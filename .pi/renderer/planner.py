# -*- coding: utf-8 -*-
"""规划引擎：把 SourceModule 清单 → DeckOutline（结构化大纲）。

这是整个规划层的核心。它做四件事：
1. **关系抽取**：先扫描模块之间的联系（时序/对比/包含/因果/并列/依赖），
   委托 doc_analyzer.extract_relations —— 关系里藏着布局
2. **版式路由**：两遍路由 —— 第一遍按单模块内容特征选版式；
   第二遍按跨模块关系合并/拆版（时序链→flow、对比→panel_2col、
   并列→numbered_cards、包含→hierarchy）
3. **密度规划**：算出每页该塞多少字、目标档位、是否需要拆页/并页
4. **图表建议**：识别适合图表化的数据点，给出具体的 chart_type + 数据

规则全部用代码实现（不依赖 LLM），确保同一份输入每次输出一致。
LLM 的职责变成：按大纲逐页生成 DSL 正文（而不是同时决定结构+内容）。
"""

from __future__ import annotations

import math
import re
from collections import Counter
from typing import Optional

import outline
import doc_analyzer
from doc_analyzer import SourceModule
from outline import (
    ChartSuggestion, DeckOutline, DensityTarget, SectionPlan, SlidePlan, SlideType
)


# ---------- 版式路由 ----------

def _route_slide_type(mod: SourceModule) -> SlideType:
    """根据模块内容特征自动选择最合适的版式。

    决策树（按优先级）：
    0. 四象限/矩阵语义 → GRID_2X2（SWOT、重要-紧急、四象限分析、2×2 矩阵）
    1. 数据密度高 → CHART（有数据可画）
    2. 转化/漏斗语义 → FUNNEL（进线→成交这类逐层收窄）
    3. 包含对比关系 → COMPARE_3COL（2~4 个并列对比项）
    4. 包含流程 → 时序语义 TIMELINE，否则 FLOW（有步骤/顺序/循环）
    5. 包含层级 → HIERARCHY（有包含/从属关系）
    6. 论据 >3 条并列 → NUMBERED_CARDS
    7. 论据 3 条：对比语义 → COMPARE_3COL，平行组 → KEY_GROUPS（statement 是
       最后兜底，别把能分组的内容摊成字堆）
    8. 论据 2 条对比 → PANEL_2COL
    9. 默认 → STATEMENT
    """
    # 四象限/矩阵语义 → 2x2 网格
    text = f"{mod.conclusion} {' '.join(mod.evidence)}"
    if _is_quadrant(text):
        return SlideType.GRID_2X2

    # 数据密度高且数据点多 → 图表
    if mod.is_data_heavy and len(mod.data_points) >= 4:
        return SlideType.CHART

    # 转化/漏斗语义（每层都在变小的流程）→ 漏斗
    if _is_funnel(text) and len(mod.evidence) >= 3:
        return SlideType.FUNNEL

    # 有对比关系且论据 >=3 条 → 三栏对比
    if mod.has_comparison and len(mod.evidence) >= 3:
        return SlideType.COMPARE_3COL

    # 有流程特征且论据 >=3 条 → 时序语义走时间轴，一般步骤走流程图
    if mod.has_process and len(mod.evidence) >= 3:
        return SlideType.TIMELINE if _is_timeline(text) else SlideType.FLOW

    # 有层级特征 → 层级图
    if mod.has_hierarchy:
        return SlideType.HIERARCHY

    # 论据数量判断
    if len(mod.evidence) >= 4:
        return SlideType.NUMBERED_CARDS
    elif len(mod.evidence) == 3:
        # 3 条论据：有对比语义才进三栏对比（那里还要造比喻），
        # 否则是平行主题组 → key_groups；摊给 statement 只会变成字堆
        return (SlideType.COMPARE_3COL if mod.has_comparison
                else SlideType.KEY_GROUPS)
    elif len(mod.evidence) == 2:
        return SlideType.PANEL_2COL

    return SlideType.STATEMENT


_QUADRANT_KEYWORDS = ("SWOT", "四象限", "矩阵", "重要-紧急", "优势-劣势",
                      "机会-威胁", "2x2", "2×2", "维度评估", "评估框架",
                      "四个维度", "四个角度", "四维度", "评估矩阵")

# 漏斗语义：流程逐层收窄（每层的量都在变小），而不是步骤平铺
_FUNNEL_KEYWORDS = ("漏斗", "转化率", "流失", "漏损", "成交率", "进线",
                    "签约率", "转化路径", "层层筛选")

# 时间轴语义：流程带阶段/时间刻度（只看模块名/结论，"分为四步"这类不算）
_TIMELINE_RE = re.compile(
    r"阶段|里程碑|路线图|[Rr]oadmap|季度|Q[1-4]|[上下]半年|20\d\d年|"
    r"第?[一二三四五六七八九十\d]+期")

_TL_CHAIN_RE = re.compile(r"阶段|里程碑|季度|Q[1-4]|20\d\d年|第?[一二三四五六七八九十\d]+期|[Rr]oadmap")


def _is_funnel(text: str) -> bool:
    """判断文本是否描述转化/漏斗类流程。"""
    return any(kw in text for kw in _FUNNEL_KEYWORDS)


def _is_timeline(text: str) -> bool:
    """判断流程文本是否带阶段/时间刻度（→timeline 而非 flow）。"""
    return bool(_TIMELINE_RE.search(text))


def _chain_is_timeline(members: list[SourceModule]) -> bool:
    """关系组合并时判时序链用 timeline 还是 flow：只看**模块名**。

    名字带阶段/期数/时间（第一阶段、二期、2024Q3）→ timeline；
    只在结论里提到阶段、名字是普通步骤名（需求分析→系统设计）→ flow。
    """
    hits = sum(1 for m in members if _TL_CHAIN_RE.search(m.module or ""))
    return hits * 2 > len(members)


def _is_quadrant(text: str) -> bool:
    """判断文本是否描述四象限/矩阵类内容。"""
    return any(kw in text for kw in _QUADRANT_KEYWORDS)


def _route_chart_type(mod: SourceModule) -> Optional[ChartSuggestion]:
    """根据数据特征建议图表类型。"""
    if not mod.data_points:
        return None

    n_points = len(mod.data_points)

    # 尝试解析数值
    values = _parse_values(mod.data_points)

    # 检查文本语义
    text = f"{mod.conclusion} {' '.join(mod.evidence)}"

    # 时间序列 → 折线图
    if any(kw in text for kw in ("趋势", "变化", "增长", "下降", "走势",
                                  "历史", "时间", "演进", "阶段")):
        if n_points >= 3:
            return ChartSuggestion(
                chart_type="line",
                categories=[f"P{i+1}" for i in range(n_points)],
                values=values or [0.0] * n_points,
                reason="文本含时间/趋势语义，数据点 {} 个 → 折线图展示变化".format(n_points),
            )

    # 占比/百分比 → 环形图
    if any(kw in text for kw in ("占比", "比例", "分布", "构成", "份额",
                                  "%", "百分比", "比重")):
        return ChartSuggestion(
            chart_type="doughnut",
            categories=[f"类别{i+1}" for i in range(n_points)],
            values=values or [100.0 / n_points] * n_points,
            reason="文本含占比/分布语义，数据点 {} 个 → 环形图展示构成".format(n_points),
        )

    # 对比 → 柱状图
    if mod.has_comparison or n_points <= 6:
        return ChartSuggestion(
            chart_type="column",
            categories=[f"项{i+1}" for i in range(n_points)],
            values=values or [0.0] * n_points,
            reason="对比类数据 {} 个 → 柱状图展示差异".format(n_points),
        )

    # 默认折线
    if n_points >= 3:
        return ChartSuggestion(
            chart_type="line",
            categories=[f"P{i+1}" for i in range(n_points)],
            values=values or [0.0] * n_points,
            reason="时间序列数据 {} 个 → 折线图".format(n_points),
        )

    return None


def _parse_values(data_points: list[str]) -> list[float]:
    """尝试把数字字符串转成 float。"""
    values = []
    for dp in data_points:
        # 去掉单位
        s = re.sub(r"[%％个家条张层倍次人天月年万亿分钟小时MBGB元]", "", dp)
        s = s.replace(",", "")
        try:
            values.append(float(s))
        except (ValueError, TypeError):
            continue
    return values


# ---------- 密度规划 ----------

def _density_for_type(stype: SlideType, mod: SourceModule) -> DensityTarget:
    """根据版式类型 + 内容量定密度档位。"""
    content_len = mod.content_length

    # 图表页：数据点少就基础，多就标准
    if stype == SlideType.CHART:
        if len(mod.data_points) <= 4:
            return DensityTarget.BASIC
        return DensityTarget.STANDARD

    # 流程图 / 层级图 / 时间轴 / 漏斗：文本量小
    if stype in (SlideType.FLOW, SlideType.HIERARCHY,
                 SlideType.TIMELINE, SlideType.FUNNEL):
        if content_len < 150:
            return DensityTarget.BASIC
        return DensityTarget.STANDARD

    # 二维矩阵 / 要点分组：组格多就标准档
    if stype in (SlideType.MATRIX, SlideType.KEY_GROUPS):
        return DensityTarget.STANDARD

    # 卡片类
    if stype == SlideType.NUMBERED_CARDS:
        if len(mod.evidence) >= 6:
            return DensityTarget.HIGH
        return DensityTarget.STANDARD

    if stype == SlideType.COMPARE_3COL:
        return DensityTarget.STANDARD

    if stype == SlideType.PANEL_2COL:
        return DensityTarget.STANDARD

    # statement
    if content_len >= 300:
        return DensityTarget.HIGH
    if content_len >= 150:
        return DensityTarget.STANDARD
    return DensityTarget.BASIC


# ---------- 备注建议生成 ----------

def _gen_notes(mod: SourceModule, stype: SlideType) -> list[str]:
    """为这一页生成备注建议：缺什么数据、补充什么资料会更好。"""
    notes = []

    # 基于内容特征的建议
    if stype == SlideType.CHART:
        if len(mod.data_points) < 3:
            notes.append("图表数据点较少（{}个），建议补充更多对比维度或时间点的数据".format(
                len(mod.data_points)))
        # 检查是否有绝对数值 vs 百分比
        has_pct = any("%" in dp or "％" in dp for dp in mod.data_points)
        has_abs = any(c.isdigit() for dp in mod.data_points for c in dp
                      if c not in "%％")
        if has_pct and not has_abs:
            notes.append("当前数据全是百分比，缺少绝对数值；补充基数能让读者判断实际规模")
        if has_abs and not has_pct:
            notes.append("当前数据全是绝对数，缺少占比/比率；补充相对比例能让跨量级对比更直观")

    if stype == SlideType.STATEMENT:
        if len(mod.data_points) == 0:
            notes.append("该页纯文字无数据；若原文有更细的数字/案例论据，可考虑改用 numbered_cards")
        if mod.is_text_heavy:
            notes.append("文本密度极高但数据稀疏；可考虑抽象为流程图或框架图来提升可视化")

    if stype in (SlideType.COMPARE_3COL, SlideType.NUMBERED_CARDS):
        if len(mod.evidence) < 3:
            notes.append("对比/卡片数量偏少（{}条），建议从原文补充更多对比维度或合并相关模块".format(
                len(mod.evidence)))

    if stype == SlideType.FLOW:
        if len(mod.evidence) < 4:
            notes.append("流程步骤较少（{}步）；若原文有更细的子流程可补充，或改用 hierarchy".format(
                len(mod.evidence)))

    if stype == SlideType.PANEL_2COL:
        if len(mod.evidence) < 4:
            notes.append("双模块并排时两侧内容都偏少；若相关模块内容足够，可改为 numbered_cards")

    # 基于缺口的建议
    for gap in mod.gaps:
        notes.append("需补充：{}".format(gap))

    # 数据类型建议
    if mod.is_data_heavy and stype not in (SlideType.CHART, SlideType.TABLE):
        notes.append("该模块数据密度高（{}个数字），适合用 chart 或 table 展示而非纯文字".format(
            len(mod.data_points)))

    if mod.is_text_heavy and stype == SlideType.STATEMENT:
        notes.append("该模块文本密度极高但无数据，建议抽象为流程图（flow）或层级图（hierarchy）")

    return notes


# ---------- 并页 / 拆页启发 ----------

def _should_merge_with_previous(mod: SourceModule, prev: SourceModule) -> bool:
    """判断当前模块是否应该与前一个模块并到同一页。"""
    # 内容太少（<100字）的模块考虑合并
    if mod.content_length < 100:
        return True
    # 同一章节、论据总数 <= 3 时合并
    if (mod.section == prev.section and
            len(mod.evidence) + len(prev.evidence) <= 4):
        return True
    return False


def _should_split(mod: SourceModule) -> bool:
    """判断是否需要拆成多页（内容过多）。"""
    if mod.content_length > 500:
        return True
    if len(mod.evidence) > 6:
        return True
    return False


# ---------- 第二遍路由：关系感知的合并/拆版 ----------

# 关系组的容量闸门：合计超过就放弃合并，回退到单模块路由 —— 宁可页数多，不可塞不下
REL_GROUP_MAX_CHARS = 700


def _rel_targets(mod: SourceModule, rtype: str) -> set[str]:
    """一个模块某类关系指向的所有目标模块名。"""
    return {r.target for r in mod.relations if r.type == rtype and r.target}


def _explicit_link(mod: SourceModule, rtype: str, other: SourceModule) -> bool:
    """mod 对 other 是否存在**显式**关系（点名 / 命名对仗 / 阶段序号）。

    "邻接推断"的关系（文本里没点名、只靠前后顺序猜的）不参与合并决策 ——
    推断只能做建议，不能做"把两页并成一页"这种破坏性决定。
    """
    return any(r.type == rtype and r.target == other.module
               and "（邻接推断）" not in r.evidence
               for r in mod.relations)


def _loose_link(mod: SourceModule, rtype: str, other: SourceModule) -> bool:
    """mod 对 other 是否存在关系（显式或邻接推断都算）。"""
    return any(r.type == rtype and r.target == other.module
               for r in mod.relations)


def _group_ok(mods: list[SourceModule]) -> bool:
    """关系组容量闸门：合计超限或任一模块需要拆页时不合。"""
    return (sum(m.content_length for m in mods) <= REL_GROUP_MAX_CHARS
            and not any(_should_split(m) for m in mods))


def _relation_groups(mods: list[SourceModule]) -> list[tuple[int, int, SlideType, str]]:
    """扫描一个章节内模块间的 relations，产出关系驱动的合并区段。

    返回 [(start, end, slide_type, reason)]，半开区间、互不重叠、按起点有序。
    优先级：sequence 链（→flow）> contrast 对（→panel_2col）>
    parallel 组（→numbered_cards）> contain 组（→hierarchy）。

    这就是「发现模块之间的联系 → 引入不同布局」的落点：
    过去每个模块独立路由，时序链被摊成三页卡片、对比被摊成两页 statement；
    现在链式/对仗/包含关系在成页前就把它们抓回来合成一张图。
    """
    n = len(mods)
    groups: list[tuple[int, int, SlideType, str]] = []
    taken: set[int] = set()
    i = 0
    while i < n:
        if i in taken:
            i += 1
            continue
        mod = mods[i]

        # 1) sequence 链：连续相邻、两两有时序关系 → 合成一页流程图
        #    （数据密度高的模块例外 —— 它该去图表页，不硬拽进流程图）
        chain = [i]
        while True:
            tail = chain[-1]
            nxt = tail + 1
            if (nxt >= n or nxt in taken
                    or not _loose_link(mods[tail], "sequence", mods[nxt])):
                break
            chain.append(nxt)
        if len(chain) >= 3:
            members = [mods[k] for k in chain]
            if (_group_ok(members)
                    and not any(m.is_data_heavy for m in members)):
                # 时序链分流：名字带阶段/期数/时间 → 时间轴，普通步骤 → 流程图
                gtype = (SlideType.TIMELINE if _chain_is_timeline(members)
                         else SlideType.FLOW)
                gname = "时间轴" if gtype == SlideType.TIMELINE else "流程图"
                groups.append((i, chain[-1] + 1, gtype,
                               f"{len(chain)} 个模块构成时序链"
                               f"（{mods[i].module}→{mods[chain[-1]].module}），"
                               f"合为一页{gname}"))
                taken.update(chain)
                i = chain[-1] + 1
                continue

        # 2) contrast 对：相邻且**显式**互指对比（任一方向）→ 合成一页左右对照
        if (i + 1 < n
                and (_explicit_link(mod, "contrast", mods[i + 1])
                     or _explicit_link(mods[i + 1], "contrast", mod))
                and _group_ok([mod, mods[i + 1]])):
            groups.append((i, i + 2, SlideType.PANEL_2COL,
                           f"「{mod.module}」与「{mods[i + 1].module}」构成对比，"
                           f"合为一页左右对照"))
            taken.update((i, i + 1))
            i += 2
            continue

        # 3) parallel 组：连续 ≥3 个并列模块 → 合成一页卡片组
        run = [i]
        while True:
            tail = run[-1]
            nxt = tail + 1
            if (nxt >= n or nxt in taken
                    or not _loose_link(mods[tail], "parallel", mods[nxt])):
                break
            run.append(nxt)
        if len(run) >= 3 and _group_ok([mods[k] for k in run]):
            groups.append((i, run[-1] + 1, SlideType.KEY_GROUPS,
                           f"{len(run)} 个模块并列（{mod.module} 等），"
                           f"合为一页要点分组"))
            taken.update(run)
            i = run[-1] + 1
            continue

        # 4) contain 组：i 点名包含 ≥2 个近邻后续模块 → 合成一页框架图
        contain_tgts = _rel_targets(mod, "contain")
        idxs = [k for k in range(i + 1, min(i + 5, n))
                if k not in taken and mods[k].module in contain_tgts]
        if len(idxs) >= 2 and _group_ok([mods[k] for k in [i] + idxs]):
            span_end = max(idxs) + 1
            groups.append((i, span_end, SlideType.HIERARCHY,
                           f"「{mod.module}」包含 {len(idxs)} 个子模块，"
                           f"合为一页框架图"))
            taken.update(range(i, span_end))
            i = span_end
            continue

        i += 1
    return groups


def _relation_slide(mods: list[SourceModule], stype: SlideType,
                    reason: str) -> SlidePlan:
    """把一个关系组（多个模块）合并成一页的 SlidePlan。"""
    names = [m.module for m in mods]
    evidence: list[str] = []
    for m in mods:
        evidence.extend(m.evidence)
    proxy = mods[0].model_copy(update={
        "evidence": evidence,
        "content_length": sum(m.content_length for m in mods),
        "data_points": [dp for m in mods for dp in m.data_points],
    })
    density = _density_for_type(stype, proxy)
    chart = _route_chart_type(proxy) if stype == SlideType.CHART else None
    notes = _gen_notes(proxy, stype)
    notes.append(f"关系路由：{reason}")
    return SlidePlan(
        section=mods[0].section,
        module=" → ".join(names) if stype == SlideType.FLOW else " + ".join(names),
        source_lines=[m.line_start for m in mods],
        title=_infer_title(mods[0]),
        slide_type=stype,
        density_target=density,
        conclusion=next((m.conclusion for m in mods if m.conclusion), ""),
        evidence=evidence,
        gaps=[g for m in mods for g in m.gaps],
        chart=chart,
        notes=notes,
        relation_note=reason,
    )


# ---------- 主入口 ----------

def plan(modules: list[SourceModule],
         title: str = "演示文稿",
         subtitle: str = "",
         brief: str = "seminar") -> DeckOutline:
    """把 SourceModule 清单 → DeckOutline。

    路由是两遍的：
    第一遍逐模块按内容特征选版式（_route_slide_type）；
    第 0 步先抽取跨模块关系（extract_relations），第二遍按关系合并/拆版
    （_relation_groups）—— 时序链合流程、对比合左右、并列合卡片、
    包含合框架。关系组优先于单模块路由。
    """
    # 第 0 步：跨模块关系抽取（幂等，重复调用不叠加）
    modules = doc_analyzer.extract_relations(modules)

    # 按章节分组
    section_map: dict[str, list[SourceModule]] = {}
    for m in modules:
        sec = m.section or "未分类"
        section_map.setdefault(sec, []).append(m)

    sections: list[SectionPlan] = []
    global_chart_types: list[str] = []
    total_density_chars = 0
    content_page_count = 0

    for sec_idx, (sec_name, sec_modules) in enumerate(section_map.items(), 1):
        slides: list[SlidePlan] = []
        groups = _relation_groups(sec_modules)
        group_at = {g[0]: g for g in groups}
        i = 0
        while i < len(sec_modules):
            # 第二遍路由：关系组优先 —— 多个模块合为一页（关系里藏着布局）
            if i in group_at:
                start, end, gtype, reason = group_at[i]
                sp = _relation_slide(sec_modules[start:end], gtype, reason)
                slides.append(sp)
                content_page_count += 1
                total_density_chars += _density_to_chars(sp.density_target)
                if sp.chart:
                    global_chart_types.append(sp.chart.chart_type)
                i = end
                continue

            mod = sec_modules[i]
            stype = _route_slide_type(mod)
            density = _density_for_type(stype, mod)
            chart = _route_chart_type(mod) if stype == SlideType.CHART else None
            notes = _gen_notes(mod, stype)

            # 并页判断
            merge_with = None
            if i > 0 and _should_merge_with_previous(mod, sec_modules[i - 1]):
                merge_with = len(slides) - 1

            # 拆页判断
            split_pages = []
            if _should_split(mod):
                n_splits = math.ceil(mod.content_length / 350)
                per_len = len(mod.evidence) // n_splits
                for si in range(n_splits):
                    start = si * per_len
                    end = start + per_len if si < n_splits - 1 else len(mod.evidence)
                    sub_evidence = mod.evidence[start:end]
                    sub_mod = mod.model_copy(update={
                        "evidence": sub_evidence,
                        "module": f"{mod.module}（{si+1}/{n_splits}）",
                    })
                    sub_stype = _route_slide_type(sub_mod)
                    sub_density = _density_for_type(sub_stype, sub_mod)
                    split_pages.append(SlidePlan(
                        section=mod.section,
                        module=sub_mod.module,
                        source_lines=[mod.line_start],
                        title=_infer_title(sub_mod),
                        slide_type=sub_stype,
                        density_target=sub_density,
                        conclusion=sub_mod.conclusion,
                        evidence=sub_mod.evidence,
                        gaps=sub_mod.gaps,
                        chart=_route_chart_type(sub_mod) if sub_stype == SlideType.CHART else None,
                        notes=_gen_notes(sub_mod, sub_stype),
                    ))
                slides.extend(split_pages)
                if chart and split_pages:
                    global_chart_types.append(chart.chart_type)
                i += 1
                continue

            sp = SlidePlan(
                section=mod.section,
                module=mod.module,
                source_lines=[mod.line_start],
                title=_infer_title(mod),
                slide_type=stype,
                density_target=density,
                conclusion=mod.conclusion,
                evidence=mod.evidence,
                gaps=mod.gaps,
                chart=chart,
                notes=notes,
                merge_with=merge_with,
            )
            slides.append(sp)

            if chart:
                global_chart_types.append(chart.chart_type)

            # 密度统计
            density_chars = _density_to_chars(density)
            total_density_chars += density_chars
            content_page_count += 1

            i += 1

        sections.append(SectionPlan(
            no=f"{sec_idx:02d}",
            title=sec_name,
            slides=slides,
        ))

    total_pages = sum(len(sec.slides) for sec in sections)
    avg_density = total_density_chars / content_page_count if content_page_count else 0

    return DeckOutline(
        title=title,
        subtitle=subtitle or None,
        brief=brief,
        sections=sections,
        total_pages=total_pages,
        content_pages=content_page_count,
        avg_density=avg_density,
        distinct_chart_types=list(dict.fromkeys(global_chart_types)),
    )


def _infer_title(mod: SourceModule) -> str:
    """从模块的结论提取标题。"""
    if mod.conclusion:
        # 去掉句末标点，截断
        t = mod.conclusion.rstrip("。！？；").strip()
        if len(t) <= 32:
            return t
        return t[:30] + "…"
    return mod.module


def _density_to_chars(density: DensityTarget) -> int:
    return {
        DensityTarget.BASIC: 200,
        DensityTarget.STANDARD: 350,
        DensityTarget.HIGH: 500,
        DensityTarget.EXTREME: 650,
    }.get(density, 350)


# ---------- CLI 入口 ----------

def main():
    """CLI: python planner.py <analysis.md> [-o outline.json] [--title T] [--brief B]"""
    import argparse
    parser = argparse.ArgumentParser(description="规划引擎：分析 → 大纲")
    parser.add_argument("input", help="输入分析表（analysis.md）或原始文档")
    parser.add_argument("-o", "--out", default=".pi/work/outline.json",
                        help="大纲输出路径")
    parser.add_argument("--title", default="演示文稿", help="deck 标题")
    parser.add_argument("--subtitle", default="", help="deck 副标题")
    parser.add_argument("--brief", default="seminar",
                        choices=["seminar", "briefing"], help="写作方向")
    parser.add_argument("--from-raw", action="store_true",
                        help="输入是原始文档（无 analysis.md），按标题自动切分")
    args = parser.parse_args()

    # 分析
    if args.from_raw:
        modules = doc_analyzer.analyze_raw_text(args.input)
    else:
        modules = doc_analyzer.analyze_analysis_md(args.input)

    # 计算缓存键
    source_hash = doc_analyzer.compute_source_hash(args.input)

    # 规划
    outline_result = plan(
        modules,
        title=args.title,
        subtitle=args.subtitle,
        brief=args.brief,
    )
    outline_result.source_hash = source_hash

    # 保存
    outline.save_outline(outline_result, args.out)

    # 输出摘要
    stats = outline_result.stats()
    rel_pages = sum(1 for sec in outline_result.sections
                    for s in sec.slides if s.relation_note)
    print(f"OK: {args.out}")
    print(f"  总页数: {stats['total']}（内容页 {stats['content']}）")
    print(f"  版式分布: {stats['by_type']}")
    print(f"  关系合并: {rel_pages} 页（跨模块关系路由）")
    print(f"  图表类型: {stats['chart_types'] or '无'}")
    print(f"  平均密度: {outline_result.avg_density:.0f} 字/页")
    print(f"  缓存键: {source_hash}")


if __name__ == "__main__":
    main()
