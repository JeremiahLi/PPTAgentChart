# -*- coding: utf-8 -*-
"""结构化大纲：规划引擎的输出，是「内容分析 → DSL 生成」之间的可审阅中间产物。

为什么要有这一层 —— 过去的「大纲」是 LLM 一段自然语言，用户看完点头就没了；
现在 Outline 是结构化 JSON，能被校验、能被编辑、能被缓存。同一份输入资料
反复生成时，Outline 只算一次，后续复用；方向变了也不用重跑分析。

数据流：DocAnalyzer →Planner → Outline → merge_outline → Deck
"""

from __future__ import annotations

import json
import os
from enum import Enum
from typing import Annotated, Literal, Optional, Union

from pydantic import BaseModel, Field


class DensityTarget(str, Enum):
    """密度档位：每页目标信息量。"""
    BASIC = "basic"           # ~200 字/页
    STANDARD = "standard"     # ~350 字/页
    HIGH = "high"             # ~500 字/页
    EXTREME = "extreme"       # ~650 字/页


class SlideType(str, Enum):
    """与 ppt_dsl.SlideType 保持同步（规划阶段用子集）。"""
    COVER = "cover"
    TOC = "toc"
    SECTION = "section"
    STATEMENT = "statement"
    COMPARE_3COL = "compare_3col"
    NUMBERED_CARDS = "numbered_cards"
    LAYERED_STACK = "layered_stack"
    FLOW = "flow"
    HIERARCHY = "hierarchy"
    TABLE = "table"
    CHART = "chart"
    PANEL_2COL = "panel_2col"
    FIGURE = "figure"
    GRID_2X2 = "grid_2x2"
    TIMELINE = "timeline"
    FUNNEL = "funnel"
    MATRIX = "matrix"
    KEY_GROUPS = "key_groups"
    END = "end"


class ChartSuggestion(BaseModel):
    """建议使用的图表类型与数据。"""
    chart_type: Literal["column", "bar", "line", "pie", "doughnut", "radar"]
    categories: list[str] = Field(default_factory=list)
    series_name: Optional[str] = None
    values: list[float] = Field(default_factory=list)
    reason: str = ""  # 为什么建议这个图


class SlidePlan(BaseModel):
    """一页 PPT 的规划条目。"""
    # 来源追溯
    section: str = ""           # 所属章节
    module: str = ""            # 分析表里的模块名
    source_lines: list[int] = Field(default_factory=list)  # analysis.md 行号

    # 内容规划
    title: str = ""             # 判断句标题
    slide_type: SlideType = SlideType.STATEMENT
    density_target: DensityTarget = DensityTarget.STANDARD

    # 内容摘要（给 LLM 生成 DSL 时参考）
    conclusion: str = ""        # 核心结论（一句判断句）
    evidence: list[str] = Field(default_factory=list)  # 论据
    gaps: list[str] = Field(default_factory=list)      # 缺口

    # 视觉建议
    chart: Optional[ChartSuggestion] = None
    notes: list[str] = Field(default_factory=list)     # 备注建议

    # 多页关联
    merge_with: Optional[int] = None  # 与哪一页合并（内容太少时）
    split_from: Optional[int] = None  # 从哪一页拆分出来
    relation_note: str = ""     # 非空 = 本页由跨模块关系路由产生（记录合并理由）


class SectionPlan(BaseModel):
    """一个章节的规划。"""
    no: str = ""                # 章节号
    title: str = ""             # 章标题
    slides: list[SlidePlan] = Field(default_factory=list)


class DeckOutline(BaseModel):
    """完整大纲：规划引擎的输出。"""
    title: str = ""
    subtitle: Optional[str] = None
    brief: str = "seminar"
    style: str = "style-harness"
    card: str = "flat"

    sections: list[SectionPlan] = Field(default_factory=list)

    # 全局统计
    total_pages: int = 0
    content_pages: int = 0
    avg_density: float = 0.0

    # 图型覆盖度
    distinct_chart_types: list[str] = Field(default_factory=list)

    # 缓存键：输入资料变化时失效
    source_hash: str = ""

    def stats(self) -> dict:
        """统计摘要，供 planner 自检。"""
        types = [s.slide_type.value for sec in self.sections for s in sec.slides]
        return {
            "total": len(types),
            "content": sum(1 for t in types if t not in ("cover", "toc", "section", "end")),
            "by_type": {t: types.count(t) for t in set(types)},
            "chart_types": self.distinct_chart_types,
            "avg_density": self.avg_density,
        }


# ---------- 序列化 ----------

def load_outline(path: str) -> DeckOutline:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return DeckOutline.model_validate(data)


def save_outline(outline: DeckOutline, path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(outline.model_dump(mode="json"), f, ensure_ascii=False, indent=2)


# ---------- 合并：Outline → Deck（轻量版，不依赖 LLM） ----------
#
# 这一步产出的**不是成品**，是"每页都合法、但正文全是待补占位"的骨架：
# 校验能过、能渲染出一份可翻的结构稿，作者（或 LLM）在上面逐页替换正文。
#
# 之前这里只写 `{slide_type, title, _conclusion}`，而 `_conclusion` 之类带下划线
# 的字段根本不在 schema 里 —— 24 页全会缺必填字段，`merge --from-outline` 必然
# 抛 29 个 validation error。链路等于断的，plan.py 却还在结尾提示下一步跑它。

PLACEHOLDER = "待补："


def _t(text: str, default: str, min_len: int = 6, max_len: int = 32) -> str:
    """把标题补到合法长度（占位也要过 schema）。"""
    s = (text or "").strip()
    if len(s) < min_len:
        s = f"{PLACEHOLDER}{default}"
    return s[:max_len]


def _body(text: str, min_len: int, max_len: int, fallback: str) -> str:
    s = (text or "").strip()
    if not s:
        s = fallback
    if not s.startswith(PLACEHOLDER):
        s = PLACEHOLDER + s
    while len(s) < min_len:
        s += "（这里补上机制、口径与论据，不少于 %d 字。）" % min_len
    return s[:max_len]


def _shell(sp: "SlidePlan") -> dict:
    """单页规划 → 该版式的**最小合法** DSL（占位正文）。"""
    t = sp.slide_type.value
    concl = (sp.conclusion or "").strip()
    ev = [e for e in (sp.evidence or []) if e]

    if t == "cover":
        return {"slide_type": "cover", "title": _t(sp.title, "演示标题", 2, 24),
                "subtitle": PLACEHOLDER + "副标题（一句话说清这份材料回答什么问题）"}
    if t == "toc":
        return {"slide_type": "toc", "items": [
            {"no": f"0{i+1}", "text": _t(concl, f"章节 {i+1}", 2, 20)}
            for i in range(3)]}
    if t == "section":
        return {"slide_type": "section", "no": "01",
                "title": _t(concl, "章节标题", 2, 20)}
    if t == "end":
        return {"slide_type": "end", "title": "谢谢观看",
                "sections": [{"heading": "下一步", "text": PLACEHOLDER + "待补收口信息"}]}
    if t == "statement":
        return {"slide_type": "statement", "title": _t(sp.title, "判断句标题", 8, 32),
                "summary": PLACEHOLDER + "（结论段：2~3 句、40~80 字）",
                "body": _body(concl, 80, 700,
                              "待补：核心论证。先给判断，再给机制，最后给口径与论据。"),
                "points": [{"label": "待补", "text": _body(x, 6, 44, "待补论据")[:44]}
                           for x in ev[:3]]}
    if t == "compare_3col":
        mk = lambda i, x: {"label": _t(x, f"对象{i+1}", 1, 10),
                           "tag": _t("", "机制", 2, 14),
                           "desc": _body(x, 20, 220, "待补：这一栏的做法与代价。"),
                           "metaphor": "待补比喻（一句话）"}
        return {"slide_type": "compare_3col", "title": _t(sp.title, "对比标题", 6, 32),
                "columns": [mk(i, ev[i] if i < len(ev) else "") for i in range(3)]}
    if t == "numbered_cards":
        n = max(4, min(8, len(ev) or 4))
        return {"slide_type": "numbered_cards",
                "title": _t(sp.title, "卡片页标题", 6, 32),
                "cards": [{"no": f"0{i+1}", "name": _t(ev[i] if i < len(ev) else "",
                                                        f"要点{i+1}", 2, 20),
                           "desc": _body(ev[i] if i < len(ev) else "", 20, 240,
                                         "待补：这一项的机制与证据。")}
                          for i in range(n)]}
    if t == "layered_stack":
        n = max(3, min(12, len(ev) or 4))
        return {"slide_type": "layered_stack",
                "title": _t(sp.title, "层级标题", 6, 32),
                "layers": [{"label": _t("", f"层{i+1}", 2, 14),
                            "desc": _body(ev[i] if i < len(ev) else "", 15, 180,
                                          "待补：这一层负责什么、依赖谁。")}
                           for i in range(n)]}
    if t == "flow":
        n = max(3, min(12, len(ev) or 5))
        return {"slide_type": "flow", "title": _t(sp.title, "流程标题", 6, 32),
                "steps": [{"label": _t("", f"步骤{i+1}", 2, 10),
                           "desc": _body(ev[i] if i < len(ev) else "", 10, 80,
                                         "待补：这一步做什么、产出什么。")}
                          for i in range(n)]}
    if t == "hierarchy":
        n = max(2, min(8, len(ev) or 4))
        return {"slide_type": "hierarchy", "title": _t(sp.title, "层级图标题", 6, 32),
                "root": _t("", "上层整体", 2, 20),
                "nodes": [{"label": _t("", f"模块{i+1}", 2, 12),
                           "desc": _body(ev[i] if i < len(ev) else "", 6, 100,
                                         "待补：负责什么。")}
                          for i in range(n)]}
    if t == "table":
        return {"slide_type": "table", "title": _t(sp.title, "表格标题", 4, 32),
                "headers": ["待补指标", "目标", "现状"],
                "rows": [["待补：" + (ev[i][:20] if i < len(ev) else f"第{i+1}项"),
                          "待补", "待补"] for i in range(3)]}
    if t == "chart":
        c = sp.chart
        cats = [x[:12] for x in (c.categories if c and c.categories else [])][:10]
        vals = [float(v) for v in (c.values if c and c.values else [])][:10]
        while len(cats) < 2:
            cats.append(f"待补{len(cats)+1}")
        while len(vals) < len(cats):
            vals.append(0.0)
        return {"slide_type": "chart", "title": _t(sp.title, "图表标题", 6, 32),
                "summary": PLACEHOLDER + "（结论段：图表说明了什么）",
                "chart": {"type": (c.chart_type if c else "column"),
                          "categories": cats,
                          "series": [{"name": (c.series_name if c and c.series_name
                                               else "待补序列"), "values": vals}]},
                "insights": [{"text": _body(ev[i] if i < len(ev) else "", 8, 70,
                                            "待补：读图结论。")} for i in range(2)]}
    if t == "panel_2col":
        mk = lambda i, x: {"label": _t("", f"模块{i+1}", 2, 14),
                           "summary": _body(x, 16, 110, "待补：模块的总结陈述段。"),
                           "points": ["待补：要点一", "待补：要点二"]}
        return {"slide_type": "panel_2col", "title": _t(sp.title, "双模块标题", 6, 32),
                "panels": [mk(0, ev[0] if ev else ""), mk(1, ev[1] if len(ev) > 1 else "")]}
    if t == "figure":
        # 图没法占位（缺文件会直接挡渲染）—— 退化成 statement 骨架，
        # 并在标题里写明原计划，作者补好图再换回 figure。
        return {"slide_type": "statement",
                "title": _t(f"{PLACEHOLDER}【原计划图示页】{sp.title}", "图示页", 8, 32),
                "summary": PLACEHOLDER + "（结论段）",
                "body": _body(concl, 80, 700,
                              "待补：先补图（EMF/WMF 原样嵌入），再把这一页改回 figure 版式。")}
    if t == "grid_2x2":
        names = ["左上", "右上", "左下", "右下"]
        return {"slide_type": "grid_2x2",
                "title": _t(sp.title, "四象限标题", 6, 32),
                "q_top_left": {"title": f"待补{names[0][:2]}", "value": "待补", "desc": "待补：这一格的一句话说明。"},
                "q_top_right": {"title": f"待补{names[1][:2]}", "value": "待补", "desc": "待补：这一格的一句话说明。"},
                "q_bottom_left": {"title": f"待补{names[2][:2]}", "value": "待补", "desc": "待补：这一格的一句话说明。"},
                "q_bottom_right": {"title": f"待补{names[3][:2]}", "value": "待补", "desc": "待补：这一格的一句话说明。"}}
    if t == "timeline":
        n = max(3, min(12, len(ev) or 4))
        return {"slide_type": "timeline", "title": _t(sp.title, "时间轴标题", 6, 32),
                "milestones": [{"label": _t("", f"节点{i+1}", 2, 14),
                                "time": f"第{i+1}期",
                                "desc": _body(ev[i] if i < len(ev) else "", 10, 120,
                                              "待补：这一阶段做什么、交付什么。")}
                               for i in range(n)]}
    if t == "funnel":
        n = max(3, min(8, len(ev) or 4))
        return {"slide_type": "funnel", "title": _t(sp.title, "漏斗标题", 6, 32),
                "stages": [{"label": _t("", f"层{i+1}", 2, 10),
                            "desc": _body(ev[i] if i < len(ev) else "", 10, 100,
                                          "待补：这一层的量级与转化率。")}
                           for i in range(n)]}
    if t == "matrix":
        cols = [_t("", f"列{i+1}", 2, 8) for i in range(3)]
        return {"slide_type": "matrix", "title": _t(sp.title, "矩阵标题", 6, 32),
                "col_labels": cols,
                "rows": [{"label": f"待补行{i+1}",
                          "cells": [{"title": f"待补格{i+1}{j+1}", "desc": "待补：一格一句话。"}
                                    for j in range(len(cols))]} for i in range(3)]}
    if t == "key_groups":
        n = max(2, min(4, len(ev) or 3))
        return {"slide_type": "key_groups",
                "title": _t(sp.title, "分组标题", 6, 32),
                "groups": [{"title": _t("", f"组{i+1}", 2, 18),
                            "points": ["待补：要点一", "待补：要点二"]} for i in range(n)]}
    raise ValueError(f"outline 里的版式「{t}」还没有骨架生成器，请补 outline._shell")


def outline_to_deck_cli(outline: DeckOutline) -> dict:
    """把大纲转成**结构合法**的最简 deck（正文全是「待补：」占位）。

    保证：能过 `validate_dict`、能渲染出页、每页字段齐全。它不是成品 ——
    作者（或 LLM）在这份骨架上逐页替换正文，替换完再跑 render --qc。
    """
    slides = []
    for sec in outline.sections:
        for sp in sec.slides:
            slides.append(_shell(sp))
    return {
        "title": outline.title or "待补：演示标题",
        "subtitle": outline.subtitle,
        "style": outline.style,
        "card": outline.card,
        "brief": outline.brief,
        "slides": slides,
    }


def slides_plan_from_section(sec: SectionPlan) -> list[dict]:
    out = []
    for sp in sec.slides:
        d = {"slide_type": sp.slide_type.value, "title": sp.title}
        if sp.conclusion:
            d["_conclusion"] = sp.conclusion
        if sp.evidence:
            d["_evidence"] = sp.evidence
        if sp.gaps:
            d["_gaps"] = sp.gaps
        if sp.chart:
            d["_chart"] = sp.chart.model_dump()
        if sp.notes:
            d["_notes"] = sp.notes
        out.append(d)
    return out
