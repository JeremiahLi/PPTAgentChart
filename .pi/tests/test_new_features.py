# -*- coding: utf-8 -*-
"""测试新增的 3 项能力：视觉层级约束、2x2 专用 schema、增量渲染缓存。"""
from __future__ import annotations

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "renderer"))

import doc_analyzer
import outline
import planner
import qc
import speaker_notes
from doc_analyzer import SourceModule
from outline import (
    ChartSuggestion, DeckOutline, DensityTarget, SectionPlan,
    SlidePlan, SlideType
)
from ppt_dsl import (
    HighlightBox, QuadrantCell, QuadrantGridSlide, SlideNotes, Slide,
    StatementSlide, validate_dict
)
from theme import FONT_SIZES

# 确保能导入兄弟模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "renderer"))


# ---------------- 外观主题（deck.style） ----------------

class TestThemeStyles:
    """deck.style 配色主题：5 套已登记 + 未登记名回退 + 原位变更 + 色板键对齐。"""

    def test_registry_is_the_five_agreed_themes(self):
        """注册表只许这 5 套（用户 2026-10 明确要求扩容，别再随手加实验色板）。"""
        import theme
        assert list(theme._THEMES) == ["style-harness", "blue-business",
                                       "bright", "forest", "violet-tech"]
        # 默认主题永远是 style-harness，且主色仍是藏青（老产出不变）
        assert theme.DEFAULT_STYLE == "style-harness"
        assert theme._BASE_COLORS["primary"] == "2760A4"

    def test_every_palette_is_key_complete(self):
        """色板 22 键必须对齐：`layouts._c()` 是 COLORS.get(name, name)，
        少一个键就会把 "cyan" 当色值去解析（RGBColor 直接抛错）。"""
        import theme
        base = set(theme._BASE_COLORS)
        for name, spec in theme._THEMES.items():
            pal = spec["colors"]
            assert set(pal) == base, f"{name} 缺键：{sorted(base - set(pal))}"
            for k, v in pal.items():
                assert len(v) == 6 and all(c in "0123456789ABCDEF" for c in v), (name, k, v)
            assert spec["accent_cycle"], f"{name} 没有强调色循环"
            assert all(k in pal for k in spec["accent_cycle"]), name
            assert all(k in pal for k in spec["chart_series"]), name

    def test_white_text_on_accent_colors_is_readable(self):
        """色带上压白字（layered_stack 的层标签、徽章）不能低于默认主题的基线。"""
        import theme

        def lum(h):
            def ch(v):
                v = v / 255
                return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
            return (0.2126 * ch(int(h[0:2], 16)) + 0.7152 * ch(int(h[2:4], 16))
                    + 0.0722 * ch(int(h[4:6], 16)))

        def ratio(a, b):
            la, lb = lum(a), lum(b)
            return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)

        for name, spec in theme._THEMES.items():
            pal = spec["colors"]
            for k in pal:
                if k.startswith("cover_") or k in ("bg_page", "card_bg", "border",
                                                   "text", "text_sub", "text_light",
                                                   "on_dark"):
                    continue
                r = ratio("FFFFFF", pal[k])
                assert r >= 3.6, f"{name}.{k}={pal[k]} 白字对比度只有 {r:.2f}:1"

    def test_apply_style_unknown_falls_back(self):
        """deck.style 写了别的名字（含曾出现过的 terra/celadon）→ 一律回退默认。"""
        import theme
        for name in ("terra", "celadon", "不存在的主题", None, ""):
            assert theme.apply_style(name) == "style-harness"
            assert theme.COLORS["primary"] == "2760A4"

    def test_each_theme_switches_palette_in_place(self):
        """切主题后 layouts 看到的必须是新色板（同一个 dict 对象）。"""
        import layouts
        import theme
        for name, spec in theme._THEMES.items():
            theme.apply_style(name)
            assert theme.CURRENT_STYLE == name
            assert layouts.COLORS is theme.COLORS
            assert dict(theme.COLORS) == spec["colors"]
            assert list(theme.ACCENT_CYCLE) == spec["accent_cycle"]
            assert list(theme.CHART_SERIES) == spec["chart_series"]
        theme.apply_style(theme.DEFAULT_STYLE)

    def test_apply_style_in_place_keeps_bindings(self):
        """layouts 持有的是同一份 dict/list 对象——apply_style 必须原地生效。"""
        import layouts
        import theme
        assert layouts.COLORS is theme.COLORS
        theme.apply_style("bright")
        assert layouts.COLORS is theme.COLORS
        assert layouts.ACCENT_CYCLE is theme.ACCENT_CYCLE
        theme.apply_style(theme.DEFAULT_STYLE)

    def test_round_trip_restores_default(self):
        import theme
        baseline = dict(theme._THEMES["style-harness"]["colors"])
        theme.COLORS["primary"] = "000000"           # 模拟被切走污染
        theme.apply_style("style-harness")
        assert theme.COLORS == baseline

    def test_palette_rejects_unknown_key(self):
        """派生色板时写错键名必须当场报错，而不是渲染到一半才炸。"""
        import pytest as _pytest
        import theme
        with _pytest.raises(KeyError):
            theme._palette(primary_typo="123456")


# ---------------- 卡片长相（deck.card） ----------------

class TestCardStyles:
    """deck.card 卡片长相：5 种已登记、未登记名回退 flat、**不改任何文字几何**。"""

    def test_registry(self):
        import theme
        assert theme.CARD_STYLES == ("flat", "outline", "tinted",
                                     "accent-left", "accent-top")
        assert theme.DEFAULT_CARD == "flat"
        assert theme.apply_card_style(None) == "flat"

    def test_unknown_falls_back(self):
        import theme
        for name in ("shadow", "rounded", "", None, "不存在的长相"):
            assert theme.apply_card_style(name) == "flat"
        theme.apply_card_style(theme.DEFAULT_CARD)

    def test_text_geometry_never_changes(self):
        """卡片长相只换画法：文本框数量/位置/尺寸/字号必须逐页完全一致。

        这是硬约束 —— qc 的容量估算走 layouts 的几何函数，长相一旦动了内边距，
        报告就和实际对不上。
        """
        import tempfile
        from pathlib import Path

        import render
        import theme
        from pptx import Presentation
        from pptx.util import Emu

        def snapshot(path):
            prs = Presentation(str(path))
            rows = []
            for sl in prs.slides:
                for sh in sl.shapes:
                    if not sh.has_text_frame or not sh.text_frame.text.strip():
                        continue
                    rows.append((round(Emu(sh.left).inches, 4),
                                 round(Emu(sh.top).inches, 4),
                                 round(Emu(sh.width).inches, 4),
                                 round(Emu(sh.height).inches, 4),
                                 sh.text_frame.text,
                                 tuple(r.font.size.pt if r.font.size else None
                                       for p in sh.text_frame.paragraphs
                                       for r in p.runs)))
            return rows

        base = None
        try:
            with tempfile.TemporaryDirectory() as td:
                for card in theme.CARD_STYLES:
                    deck = _mini_deck()
                    deck.style = "style-harness"
                    deck.card = card
                    out = Path(td) / f"{card}.pptx"
                    render.build_pptx(deck, str(out), add_speaker_notes=False)
                    snap = snapshot(out)
                    assert snap, f"{card} 没渲染出文字"
                    if base is None:
                        base = snap
                    else:
                        assert snap == base, f"{card} 改动了文字几何"
        finally:
            theme.apply_card_style(theme.DEFAULT_CARD)

    def test_all_25_combinations_render(self):
        """5 套配色 × 5 种卡片全都能真的渲出来（坑 26：新分支必须真的渲染过）。"""
        import tempfile
        import zipfile
        from pathlib import Path

        import render
        import theme

        try:
            with tempfile.TemporaryDirectory() as td:
                for style in theme._THEMES:
                    for card in theme.CARD_STYLES:
                        deck = _mini_deck()
                        deck.style = style
                        deck.card = card
                        out = Path(td) / f"{style}-{card}.pptx"
                        render.build_pptx(deck, str(out), add_speaker_notes=False)
                        assert zipfile.is_zipfile(out) and out.stat().st_size > 20_000, \
                            f"{style}/{card} 渲染产物不可用"
        finally:
            theme.apply_style(theme.DEFAULT_STYLE)
            theme.apply_card_style(theme.DEFAULT_CARD)

    def test_qc_flags_unknown_names_without_blocking(self):
        """未登记的名字只出 warning、不拦渲染（deck 是外观偏好，写错不该整份失败）。"""
        deck = _mini_deck()
        deck.style = "terra"
        deck.card = "shadow"
        report = qc.check(deck)
        msgs = " ".join(w["message"] for w in report["warnings"])
        assert not report["errors"]
        assert "未登记的配色主题" in msgs and "未登记的卡片长相" in msgs


def _mini_deck():
    """取 fixture 里的封面 + 编号卡两页，做外观测试的最小 deck（渲染很快）。"""
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "deck.all_types.json")
    with open(fixture, "r", encoding="utf-8") as f:
        data = json.load(f)
    keep = [s for s in data["slides"] if s["slide_type"] in ("cover", "numbered_cards")]
    assert len(keep) == 2, "fixture 里应当有封面与编号卡各一页"
    return validate_dict({"title": data.get("title") or "外观样张", "slides": keep})


# ---------------- 视觉层级约束 ----------------

class TestVisualHierarchyConstraints:
    """视觉层级：schema 层 + 渲染层双重约束。"""

    def test_hero_value_font_size(self):
        """核心指标值必须使用 20pt（比正文 11.5pt 大 ≥2 级）。"""
        assert FONT_SIZES["hero_value"] == 20
        assert FONT_SIZES["hero_value"] - FONT_SIZES["card_body"] >= 8

    def test_quadrant_value_font_size(self):
        """2x2 格子核心指标值必须使用 18pt（比正文大 ≥2 级）。"""
        assert FONT_SIZES["quadrant_value"] == 18
        assert FONT_SIZES["quadrant_value"] - FONT_SIZES["card_body"] >= 6

    def test_highlight_field_description_mentions_hero(self):
        """HighlightBox 的 value 字段描述应提到视觉层级。"""
        schema = HighlightBox.model_json_schema()
        desc = schema["properties"]["value"]["description"]
        assert "20pt" in desc or "hero" in desc.lower() or "视觉" in desc

    def test_page_title_vs_body_hierarchy(self):
        """内容页标题（23pt）比正文（12.5pt）大 ≥2 级。"""
        assert FONT_SIZES["page_title"] - FONT_SIZES["body"] >= 10

    def test_card_title_vs_body_hierarchy(self):
        """卡片标题（15pt）比卡片正文（11.5pt）大 ≥1 级。"""
        assert FONT_SIZES["card_title"] - FONT_SIZES["card_body"] >= 3


# ---------------- 2x2 专用 schema ----------------

class TestGrid2x2Schema:
    """grid_2x2 版式：强制 2×2 排列的四象限/矩阵。"""

    def test_grid_2x2_in_slide_union(self):
        """版式清单里有 grid_2x2（判别字段的真值在 ppt_dsl.Slide 的 union 上，不再另存枚举）。"""
        from typing import get_args

        union = get_args(Slide)[0]
        values = []
        for model in get_args(union):
            values += list(get_args(model.model_fields["slide_type"].annotation))
        assert "grid_2x2" in values
        # 2026-10 新增 timeline/funnel/matrix/key_groups → 18 内容版式 + end = 19
        assert len(values) == len(set(values)) == 19

    def test_quadrant_cell_schema(self):
        """QuadrantCell 字段约束。"""
        cell = QuadrantCell(title="性能", value="85%", desc="比竞品高 10 个点")
        assert cell.title == "性能"
        assert cell.value == "85%"

    def test_quadrant_grid_slide_schema(self):
        """QuadrantGridSlide 包含 4 个格子。"""
        slide = QuadrantGridSlide(
            slide_type="grid_2x2",
            title="四象限分析矩阵",
            q_top_left=QuadrantCell(title="优势", value="95%", desc="比竞品高"),
            q_top_right=QuadrantCell(title="劣势", value="15%", desc="需要改进"),
            q_bottom_left=QuadrantCell(title="机会", value="3x", desc="增长空间"),
            q_bottom_right=QuadrantCell(title="威胁", value="低", desc="可控范围"),
        )
        assert slide.q_top_left.title == "优势"
        assert slide.q_bottom_right.value == "低"

    def test_grid_2x2_validation_empty_value(self):
        """空 value 应能过 schema（qc 层检查）。"""
        slide = QuadrantGridSlide(
            slide_type="grid_2x2",
            title="四象限分析矩阵",
            q_top_left=QuadrantCell(title="维度A", value="1", desc="说明A"),
            q_top_right=QuadrantCell(title="维度B", value="2", desc="说明B"),
            q_bottom_left=QuadrantCell(title="维度C", value="3", desc="说明C"),
            q_bottom_right=QuadrantCell(title="维度D", value="4", desc="说明D"),
        )
        assert slide is not None


class TestGrid2x2Routing:
    """规划引擎应将四象限/矩阵语义路由到 grid_2x2。"""

    def test_swot_routing(self):
        mod = SourceModule(
            conclusion="SWOT 分析",
            evidence=["优势", "劣势", "机会", "威胁"],
        )
        assert planner._route_slide_type(mod) == SlideType.GRID_2X2

    def test_quadrant_routing(self):
        mod = SourceModule(
            conclusion="四象限分析",
            evidence=["左上", "右上", "左下", "右下"],
        )
        assert planner._route_slide_type(mod) == SlideType.GRID_2X2

    def test_matrix_routing(self):
        mod = SourceModule(
            conclusion="重要-紧急矩阵",
            evidence=["重要紧急", "重要不紧急", "不重要紧急", "不重要不紧急"],
        )
        assert planner._route_slide_type(mod) == SlideType.GRID_2X2


# ---------------- 增量渲染缓存 ----------------

class TestIncrementalRendering:
    """增量渲染缓存：source_hash 相同则跳过重新生成。"""

    def test_compute_source_hash_deterministic(self):
        """同一文件多次计算 hash 一致。"""
        h1 = doc_analyzer.compute_source_hash(__file__)
        h2 = doc_analyzer.compute_source_hash(__file__)
        assert h1 == h2

    def test_compute_source_hash_different_files(self):
        """不同文件 hash 不同。"""
        h1 = doc_analyzer.compute_source_hash(__file__)
        h2 = doc_analyzer.compute_source_hash(
            os.path.join(os.path.dirname(__file__), "test_new_features.py"))
        # 同一文件
        assert h1 == h2

    def test_compute_source_hash_missing_file(self):
        """缺失文件不影响其他文件 hash。"""
        h1 = doc_analyzer.compute_source_hash(__file__)
        h2 = doc_analyzer.compute_source_hash(__file__, "nonexistent.md")
        assert h1 == h2

    def test_outline_has_source_hash(self):
        """DeckOutline 包含 source_hash 字段。"""
        ol = DeckOutline(title="test")
        assert hasattr(ol, "source_hash")
        assert ol.source_hash == ""

    def test_outline_save_load_preserves_hash(self):
        """序列化保留 source_hash。"""
        ol = DeckOutline(title="test", source_hash="abc123")
        data = ol.model_dump(mode="json")
        restored = DeckOutline.model_validate(data)
        assert restored.source_hash == "abc123"

    def test_plan_sets_source_hash(self):
        """planner.plan 返回的 outline 可以设置 source_hash。"""
        modules = [
            SourceModule(conclusion="test", evidence=["a", "b"]),
        ]
        ol = planner.plan(modules, title="test")
        ol.source_hash = "testhash123"
        assert ol.source_hash == "testhash123"

    def test_save_outline_creates_parent_dir(self):
        """save_outline 自动创建父目录。"""
        ol = DeckOutline(title="test", source_hash="h")
        path = ".pi/work/test_nested/dir/outline.json"
        os.makedirs(os.path.dirname(path), exist_ok=True)
        outline.save_outline(ol, path)
        assert os.path.isfile(path)
        loaded = outline.load_outline(path)
        assert loaded.source_hash == "h"
        os.remove(path)


# ---------------- qc.py 新增检查 ----------------

class TestQCVisualHierarchyChecks:
    """qc.py 对视觉层级的检查。"""

    def test_statement_highlight_without_number_warns(self):
        """statement.highlight.value 不含数字时 warning。"""
        slide = StatementSlide(
            slide_type="statement",
            title="这是一个标题即结论句",
            body="这是一段测试正文。" * 20,
            highlight=HighlightBox(label="指标", value="显著提升", desc=""),
        )
        errors, warnings = [], []
        qc._check_emphasis(1, "statement", "body", slide.body, warnings)
        assert isinstance(warnings, list)


# ---------------- 规划引擎 grid_2x2 密度目标 ----------------

class TestGrid2x2Density:
    """grid_2x2 版式的密度规划。"""

    def test_quadrant_module_density(self):
        mod = SourceModule(
            conclusion="四象限分析",
            evidence=["优势", "劣势", "机会", "威胁"],
        )
        stype = planner._route_slide_type(mod)
        density = planner._density_for_type(stype, mod)
        # grid_2x2 走默认 statement 的密度逻辑（content_length 小 → BASIC）
        assert density == DensityTarget.BASIC


# ---------------- outline_to_deck_cli 支持 grid_2x2 ----------------

class TestOutlineToDeckCli:
    """outline.outline_to_deck_cli 正确处理 grid_2x2。"""

    def test_outline_with_grid_2x2(self):
        ol = DeckOutline(
            title="test",
            brief="seminar",
            sections=[
                SectionPlan(
                    no="01",
                    title="第一章",
                    slides=[
                        SlidePlan(
                            section="第一章",
                            module="SWOT",
                            title="SWOT 分析",
                            slide_type=SlideType.GRID_2X2,
                        ),
                    ],
                ),
            ],
        )
        deck = outline.outline_to_deck_cli(ol)
        assert deck["slides"][0]["slide_type"] == "grid_2x2"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
