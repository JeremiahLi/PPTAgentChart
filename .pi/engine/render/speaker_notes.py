# -*- coding: utf-8 -*-
"""备注建议：为每页 PPT 生成「建议补充的数据/资料」，写到幻灯片备注区。

设计目标 ——
- 用户在 PowerPoint 里打开备注就能看到「这一页缺什么、要补什么」
- 备注是基于规划引擎的 SourceModule 特征自动生成的，不是 LLM 即兴发挥
- 同一份 doc 重复生成时，备注也不变（跟 outline 一起缓存）

备注内容分三类：
1. 数据缺口：本页数字太少、缺基数、缺时间序列
2. 可视化建议：当前用文字但更适合图表/流程图
3. 补充建议：需要哪些外部资料能让这页更有说服力
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class NoteItem:
    """一条备注建议。"""
    category: str  # "data" | "viz" | "source" | "gap"
    text: str


@dataclass
class SpeakerNotes:
    """一页 PPT 的完整备注。"""
    slide_title: str = ""
    slide_type: str = ""
    module: str = ""
    notes: list[NoteItem] = field(default_factory=list)

    def render(self) -> str:
        """渲染成纯文本（写入 .pptx 备注区）。"""
        if not self.notes:
            return ""

        lines = []
        # 分类标签
        cat_labels = {
            "data": "数据缺口",
            "viz": "可视化建议",
            "source": "补充资料",
            "gap": "内容缺口",
        }

        for note in self.notes:
            label = cat_labels.get(note.category, "建议")
            lines.append(f"[{label}] {note.text}")

        return "\n".join(lines)


# ---------- 规则引擎 ----------

def generate_notes(
    slide_type: str,
    module_name: str,
    conclusion: str,
    evidence: list[str],
    gaps: list[str],
    data_points: list[str],
    is_data_heavy: bool = False,
    is_text_heavy: bool = False,
    has_comparison: bool = False,
    has_process: bool = False,
    has_hierarchy: bool = False,
) -> SpeakerNotes:
    """为一页 PPT 自动生成备注建议。"""
    notes = SpeakerNotes(
        slide_title=conclusion[:40] if conclusion else module_name,
        slide_type=slide_type,
        module=module_name,
    )

    full_text = f"{conclusion} {' '.join(evidence)}"

    # ---- 数据缺口 ----
    notes.notes.extend(_check_data_gaps(slide_type, data_points, full_text))

    # ---- 可视化建议 ----
    notes.notes.extend(_check_viz_suggestions(
        slide_type, is_data_heavy, is_text_heavy,
        has_comparison, has_process, has_hierarchy,
        len(evidence), len(data_points),
    ))

    # ---- 内容缺口 ----
    notes.notes.extend(_check_content_gaps(gaps, evidence, conclusion))

    # ---- 补充资料建议 ----
    notes.notes.extend(_check_source_suggestions(
        slide_type, full_text, data_points,
    ))

    return notes


def _check_data_gaps(slide_type: str, data_points: list[str],
                     text: str) -> list[NoteItem]:
    """检查数据缺口。"""
    items = []

    if slide_type == "chart":
        if len(data_points) < 4:
            items.append(NoteItem(
                "data",
                f"当前图表仅 {len(data_points)} 个数据点，建议补充更多维度或时间切片"
            ))
        # 检查百分比 vs 绝对数
        has_pct = any("%" in dp or "％" in dp for dp in data_points)
        has_abs = any(re.search(r"\d", dp) and "%" not in dp and "％" not in dp
                      for dp in data_points)
        if has_pct and not has_abs:
            items.append(NoteItem(
                "data",
                "数据全是百分比，缺绝对基数；补充「总量/规模」让读者能判断实际大小"
            ))
        if has_abs and not has_pct:
            items.append(NoteItem(
                "data",
                "数据全是绝对数，缺相对比率；补充「占比/增长率」让跨量级对比更直观"
            ))

    if slide_type in ("statement", "numbered_cards", "compare_3col"):
        if len(data_points) == 0:
            items.append(NoteItem(
                "data",
                "该页无数字论据；从原文抽取具体数值（如准确率、耗时、规模）可增强说服力"
            ))
        elif len(data_points) < 3:
            items.append(NoteItem(
                "data",
                f"仅 {len(data_points)} 个数字，论据偏少；原文有更细的数字可补充"
            ))

    # 检查是否有对比数据但没用对比版式
    if slide_type not in ("compare_3col", "chart") and any(
            kw in text for kw in ("vs", "对比", "相比", "高于", "低于")):
        items.append(NoteItem(
            "data",
            "文本含对比语义但未用对比版式；补充两组具体数值后可改为 compare_3col"
        ))

    return items


def _check_viz_suggestions(
    slide_type: str,
    is_data_heavy: bool,
    is_text_heavy: bool,
    has_comparison: bool,
    has_process: bool,
    has_hierarchy: bool,
    n_evidence: int,
    n_data: int,
) -> list[NoteItem]:
    """检查可视化改进建议。"""
    items = []

    # 数据密集但没用图表
    if is_data_heavy and slide_type not in ("chart", "table"):
        items.append(NoteItem(
            "viz",
            f"该模块含 {n_data} 个数字，数据密度高；改用 chart 或 table 版式比纯文字更直观"
        ))

    # 文本密集但没用流程图/框架图
    if is_text_heavy and slide_type == "statement":
        if has_process:
            items.append(NoteItem(
                "viz",
                "文本含流程/步骤语义但用纯文字 statement；改用 flow 版式画流程图更清晰"
            ))
        elif has_hierarchy:
            items.append(NoteItem(
                "viz",
                "文本含层级/包含语义但用纯文字 statement；改用 hierarchy 版式画框架图"
            ))
        elif has_comparison and n_evidence >= 3:
            items.append(NoteItem(
                "viz",
                "文本含多维度对比但用纯文字 statement；改用 compare_3col 或 panel_2col"
            ))
        else:
            items.append(NoteItem(
                "viz",
                "文本密度极高但无数据，也无流程图/框架图语义；考虑精简或从原文抽取更多论据"
            ))

    # 图表类型建议
    if slide_type == "chart" and n_data >= 6:
        items.append(NoteItem(
            "viz",
            f"数据点较多（{n_data}个）；若含时间维度可换折线图，占比维度可换环形图"
        ))

    return items


def _check_content_gaps(gaps: list[str], evidence: list[str],
                        conclusion: str) -> list[NoteItem]:
    """检查内容缺口（从分析表的「缺口」列来）。"""
    items = []
    for gap in gaps:
        items.append(NoteItem("gap", f"需补充：{gap}"))

    # 结论是否足够判断化
    if conclusion:
        # 检查结论是否包含数字
        has_number = any(c.isdigit() for c in conclusion)
        # 检查结论是否包含判断性词汇
        has_judgment = any(kw in conclusion for kw in
                          ("提升", "降低", "高于", "低于", "优于", "取决于",
                           "是", "应", "需要", "应该", "必须"))
        if not has_number and not has_judgment:
            items.append(NoteItem(
                "gap",
                "结论句无数字、无判断性表述；改为「判断 + 数字」格式更有力"
            ))

    return items


def _check_source_suggestions(slide_type: str, text: str,
                              data_points: list[str]) -> list[NoteItem]:
    """建议补充哪些外部资料。"""
    items = []

    # 根据文本语义推荐资料类型
    if any(kw in text for kw in ("用户", "客户", "满意度", "NPS", "留存", "活跃")):
        items.append(NoteItem(
            "source",
            "建议补充：用户调研报告、NPS 原始数据、留存曲线、活跃用户定义口径"
        ))

    if any(kw in text for kw in ("性能", "耗时", "延迟", "QPS", "吞吐", "压测")):
        items.append(NoteItem(
            "source",
            "建议补充：性能基准测试报告、压测原始数据、对比竞品的 benchmark 数据"
        ))

    if any(kw in text for kw in ("收入", "成本", "利润", "毛利", "ROI", "增长")):
        items.append(NoteItem(
            "source",
            "建议补充：财务报表、成本拆解表、增长归因分析"
        ))

    if any(kw in text for kw in ("市场", "份额", "竞品", "行业", "趋势")):
        items.append(NoteItem(
            "source",
            "建议补充：第三方市场报告（IDC/Gartner/艾瑞）、竞品公开数据、行业白皮书"
        ))

    if any(kw in text for kw in ("风险", "合规", "安全", "隐私", "监管")):
        items.append(NoteItem(
            "source",
            "建议补充：合规检查清单、安全审计报告、风险评估矩阵、法规原文摘录"
        ))

    if any(kw in text for kw in ("实验", "A/B", "测试", "效果", "准确率")):
        items.append(NoteItem(
            "source",
            "建议补充：实验设计文档、原始数据、统计显著性检验结果"
        ))

    # 如果完全没有数据，给出通用建议
    if not data_points:
        items.append(NoteItem(
            "source",
            "该页无任何数字论据；建议从原文抽取具体数值，或补充相关数据源"
        ))

    return items
