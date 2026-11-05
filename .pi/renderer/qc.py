# -*- coding: utf-8 -*-
"""自动质检：schema 之上的几何/密度检查。

输出结构化报告 {errors: [...], warnings: [...], stats: {...}}，
error 会阻止渲染，warning 只提示。报告直接反馈给 LLM 修正。
"""
from __future__ import annotations

import json
import os
import re

import briefs
import layouts
from layouts import (CLAIM_MAX, chart_bottom, col_widths, compare_desc_box,
                     flow_linear_geom, hierarchy_geom, layered_geom, numbered_geom,
                     panel_slack, split_body_groups, split_claim, split_sections,
                     statement_body_box, table_slack)
import theme
from theme import FONT_SIZES
from ppt_dsl import (ChartSlide, CompareSlide, EndSlide, FigureSlide,
                     FlowSlide, FunnelSlide, HierarchySlide, KeyGroupsSlide,
                     LayeredStackSlide, MatrixSlide, NumberedCardsSlide,
                     Panel2ColSlide, QuadrantGridSlide, StatementSlide,
                     TableSlide, TimelineSlide)


# 速度档（只对用户暴露两档，默认慢档）：
#   slow 慢 = 报告全量列出 + 要求 warning 归零（build.py 退出码 3）+ 大纲给用户过目；
#   fast 快 = 同类 warning 折成一条、不再要求归零，直接出成品。
# **两档跑的是同一套检查**：快档省的是回炉轮次与上下文长度，不是检查项 ——
# 少查一项就会漏掉密度与几何问题，那才是"排版异常"的真正来源。
SPEEDS = ("slow", "fast")
DEFAULT_SPEED = "slow"


def _item(slide_no, slide_type, field, message):
    return {"slide": slide_no, "type": slide_type, "field": field, "message": message}


# 单个加粗跨度的字数上限：只标关键词，不标整句（渲染器只认 **...**）。
# 与 `ppt-common`/方向技能里写的「每处 ≤16 字」对齐 —— 以前 qc 是 24，
# 文档是 16：模型按 20 字写能过 qc，却与规范不符（同一件事两套阈值）。
_BOLD_MAX = 16
_BOLD_RATIO = 0.40


def _bold_spans(text: str) -> list[str]:
    """取出 `**...**` 之间的加粗内容（渲染器就是按这个切的）。"""
    parts = (text or "").split("**")
    return [parts[i].strip() for i in range(1, len(parts), 2) if parts[i].strip()]


def _section_heads(text: str) -> list[str]:
    """文本里的小节标题（`**小节名**：`）—— 它们是**结构**，不是"重点"。"""
    return [h for h, _ in split_sections(text) if h]


def _check_emphasis(slide_no, slide_type, field, text, warnings, require_span=False):
    """加粗要精炼：渲染器不再自动给首句加粗，重点全靠 `**...**` 标出来。

    标得太长、标得太多（整页都粗）都会让重点失效 —— 这正是 v14 P4 的问题。
    卡内的 `**关键短语**：` 与普通加粗同一口径计数（它们同样是显眼的强调色文字），
    短语本身只有几个字，正常写法不会超；超了就说明整段都被标成了重点。
    """
    if not text:
        return
    spans = _bold_spans(text)
    emph = spans
    plain = len(text.replace("**", ""))
    over = [sp for sp in emph if len(sp) > _BOLD_MAX]
    if over:
        warnings.append(_item(slide_no, slide_type, field,
                              f"加粗跨度过长（{len(over[0])} 字：「{over[0][:18]}…」）："
                              f"只加粗关键词（≤{_BOLD_MAX} 字），整句加粗等于没加粗"))
    bold_chars = sum(len(sp) for sp in emph)
    # 短文本（一两行的结论段）标两三处关键词很正常，占比只对长文本才有意义
    if plain >= 120 and bold_chars > plain * _BOLD_RATIO:
        warnings.append(_item(slide_no, slide_type, field,
                              f"加粗占了全文 {bold_chars / plain:.0%}：重点太多就等于没有重点，"
                              f"只留最关键的 1-2 处"))
    if require_span and not spans and plain >= 60:
        warnings.append(_item(slide_no, slide_type, field,
                              "没有用 **...** 标出核心短语：渲染器不会再自动给首句加粗，"
                              "这一页现在没有任何重点（用 **最短的判断句** 标 1-2 处）"))


# ---------------- 卡片骨架：关键句 + 关键短语 ----------------
# 反馈（v22 P11）：纯文字卡片就是一坨连贯段落，整页"没有特点也没有重点"。
# 骨架**不许靠放大字号**（关键句提到卡顶放大 + 细线会把文字顶出卡片背景），
# 最终规则是「同字号 + 加粗 + 强调色」，两条都做成可机检的：
#   1. 每张卡的第一句会被自动加粗染色成**关键句** —— 写太长整张卡都成重点；
#   2. 长文本要有 2 处以上**关键短语**（`**短语**：正文…`）—— 170+ 字一整段没人读得下去。
# 渲染侧对应 layouts.split_claim / split_sections / card_paras（同一份切法）。
_CARD_LONG = 170          # 超过这个字数就该有 2 处以上关键短语
_SECTION_MIN = 2          # 长卡至少要有几处关键短语


def _check_card_structure(slide_no, s, warnings):
    """卡片要有骨架：第一句是关键句，长文标 2 处以上关键短语。"""
    if isinstance(s, StatementSlide):
        for gi, g in enumerate(split_body_groups(s.body), 1):
            claim, _rest = split_claim(g)
            cn = len(claim.replace("**", ""))
            if cn > CLAIM_MAX:
                warnings.append(_item(
                    slide_no, "statement", "body",
                    f"正文卡 {gi} 的关键句 {cn} 字（建议 ≤{CLAIM_MAX} 字）："
                    f"渲染器会把**每张卡的第一句**自动加粗染色 —— 写成一句判断"
                    f"（带结论或数字），写长了整张卡都是重点，等于没有重点"))
            heads = len(_section_heads(g))
            if len(g) > _CARD_LONG and heads < _SECTION_MIN:
                warnings.append(_item(
                    slide_no, "statement", "body",
                    f"正文卡 {gi} 有 {len(g)} 字却只有 {heads} 处关键短语："
                    f"用 `**关键短语**：正文…` 标 2~3 处（如「**判定口径**：…」"
                    f"「**反例**：…」），渲染器把它们加粗染成强调色、冒号后正文跟排 —— "
                    f"观众扫短语就能抓住这一段的骨架"))
    cards = []
    if isinstance(s, NumberedCardsSlide):
        cards = [(c.name, c.desc) for c in s.cards]
    elif isinstance(s, CompareSlide):
        cards = [(c.label, c.desc) for c in s.columns]
    for name, txt in cards:
        heads = len(_section_heads(txt))
        if len(txt) > _CARD_LONG and heads < _SECTION_MIN:
            warnings.append(_item(
                slide_no, s.slide_type, "desc",
                f"卡片「{name}」描述 {len(txt)} 字却只有 {heads} 处关键短语："
                f"长描述用 `**关键短语**：正文…` 标 2~3 处，加粗染色 + 冒号后正文跟排；"
                f"一整段读下来观众什么都记不住"))


# ---------------- 结论段（summary）硬指标 ----------------
# 规则原本只写在 SKILL.md 里（"2-4 行的结论"），qc 只查了"有没有 / 是否 <30 字"，
# 于是 LLM 随手写废话、写"1.2 节给出…"、复述标题 —— 规则等于没生效。
# 这里把它做成可机检的硬指标，全部是 warning（不拦渲染）。
#
# 2026-09 调紧：结论段从 60~140 字压到 **40~80 字 / 2~3 句**。
# 页头是每页最贵的一块版面（2 行 ≈ 0.73in），而它和卡片描述的信息高度重叠；
# 现在结论段只留「判断 + 关键数字」，机制细节一律下放到卡片/栏描述。
_SUMMARY_MIN, _SUMMARY_MAX = 40, 80       # 字数区间（不含 ** 标记）
_SUMMARY_SENTS = (2, 3)                   # 句子数区间
# 引用坐标与自指：结论段写的是判断，不是阅读指引
_SUMMARY_COORD = (
    (r"\d+(\.\d+)*\s*节", "章节编号（如「1.2 节」）"),
    (r"第\s*[0-9一二三四五六七八九十]+\s*[章节]", "章节序号（如「第 4 章」）"),
    # 注意：不要写成「上文|下文」—— 会命中「上下文工程」这种高频技术术语
    (r"本节|本章|该节|前述|如前所述|后文|上节|下节", "自指（本节/上文/后文）"),
    (r"原文|原书|原文档|原资料", "指向原文（原文/原资料）"),
    (r"如表\s*\d*|见表|见图|如图所示|如下图|参见|详见|见表\s*\d", "版式指引（见表/如图所示）"),
)
# 纯填充语：删掉不影响信息量
_SUMMARY_FILLER = ("值得注意的是", "需要指出的是", "众所周知", "不难看出",
                   "总的来说", "综上所述", "可以说", "本页", "本课件")


def _summary_sentences(text: str) -> list[str]:
    """按中文句末标点切句；** 加粗标记与换行先剥离，免得算进字数与句数。"""
    plain = (text or "").replace("**", "").replace("\n", "")
    # 只按句末标点切：分号/破折号是句内停顿，算成句子会把长句当成合规
    return [s for s in re.split(r"[。！？]", plain) if s.strip()]


def _cjk_len(text: str) -> int:
    """中文字数（只数 CJK 汉字，标点与西文不算）—— 全文件唯一一份。"""
    return sum(1 for c in text if "\u4e00" <= c <= "\u9fff")


def _repeats_title(summary: str, title: str) -> str | None:
    """summary 里复述标题的那段子串（标题 <10 字时查整串），没有则 None。

    只把**中文表述**的重合算复述：标题里的专有名词（Claude Code、LangChain、
    GPT 之类）在结论段里出现是正常的，不能算"换个说法再说一遍"。
    """
    if not summary or not title:
        return None
    s = summary.replace("**", "")
    t = title.replace("**", "").strip()
    if len(t) < 4:
        return None
    if len(t) < 10:
        return t if _cjk_len(t) >= 3 and t in s else None
    for i in range(len(t) - 10 + 1):
        seg = t[i:i + 10]
        if _cjk_len(seg) >= 6 and seg in s:
            return seg
    return None


def _check_summary(slide_no, slide_type, summary, title, warnings):
    """标题下结论段的硬指标：40~80 字 / 2~3 句 / 禁坐标词 / 禁填充语 / 不复述标题。

    结论段要写**这一页的判断**：开口即结论、带可核对的数据或口径。
    写完读一遍 —— 如果把这段拿掉，读者少知道的必须正好是这一页的结论。
    """
    if not summary:
        return
    plain = summary.replace("**", "")
    n = len(plain)
    if n < _SUMMARY_MIN:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段只有 {n} 字（要求 {_SUMMARY_MIN}~{_SUMMARY_MAX} 字）："
                              f"补上支撑判断的数据或口径，别用一句空话占位"))
    elif n > _SUMMARY_MAX:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段 {n} 字，超出 {_SUMMARY_MAX} 字上限："
                              f"只留这一页最核心的判断与关键数字，机制细节全部下放到卡片/栏描述"
                              f"（页头是每页最贵的版面，写长了就是把正文的空间挪走）"))
    n_sent = len(_summary_sentences(summary))
    if n_sent < _SUMMARY_SENTS[0]:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段只有 {n_sent} 句（要求 {_SUMMARY_SENTS[0]}~{_SUMMARY_SENTS[1]} 句）："
                              f"判断句之后要跟一句依据（数字、口径或适用边界），别只给结论不给理由"))
    elif n_sent > _SUMMARY_SENTS[1]:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段 {n_sent} 句（要求 {_SUMMARY_SENTS[0]}~{_SUMMARY_SENTS[1]} 句）："
                              f"写成小作文了 —— 只留判断 + 1~2 条依据，展开留给正文卡片"))
    hits = [label for pat, label in _SUMMARY_COORD if re.search(pat, plain)]
    if hits:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段里有 {'、'.join(dict.fromkeys(hits))}："
                              f"结论段写判断，不写阅读指引与出处（章节坐标是讲稿的话，不是 PPT 上的字）"))
    fills = [w for w in _SUMMARY_FILLER if w in plain]
    if fills:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段有填充语「{'、'.join(fills)}」：删掉不影响信息量，"
                              f"省下的字用来写数字或口径"))
    dup = _repeats_title(summary, title or "")
    if dup:
        warnings.append(_item(slide_no, slide_type, "summary",
                              f"结论段复述了标题（「{dup}…」）：标题已经说了判断，"
                              f"结论段要接着往下讲依据与边界，别换个说法再说一遍"))


def _check_table_columns(slide_no, s, warnings):
    """表格列的信息量：整列同一个值的列（典型是「原文出处」）不该占一列宽度。

    列宽现在按内容自适应，但一列零信息量仍然在吃页面 —— 密度上不去的一半原因在这。
    """
    n_rows = len(s.rows)
    if n_rows < 4:
        return
    for c in range(len(s.headers)):
        vals = [r[c] for r in s.rows if c < len(r)]
        if len(vals) < 4:
            continue
        # 只查「整列同一个值」——典型是「原文出处」列：零信息量却占一列宽度。
        # 短列（标签列、专有名词列）不算问题：列宽已经按内容自适应了。
        if len(set(vals)) == 1:
            warnings.append(_item(slide_no, "table", f"columns[{c}]",
                                  f"第 {c + 1} 列「{s.headers[c]}」{n_rows} 行全都是"
                                  f"「{vals[0][:12]}」，零信息量却占一列宽度："
                                  f"删掉这列，内容并进 `note` 一句带过"))


def _check_page_blocks(slide_no, s, warnings):
    """一页至少两个信息块 —— 整页只有一个容器，信息密度一定上不去。"""
    t = s.slide_type
    if t == "table":
        blocks = 1 + (1 if (getattr(s, "side", None) or []) else 0) \
                   + (1 if s.note else 0)
        if blocks < 2:
            warnings.append(_item(slide_no, "table", "side",
                                  "本页只有一张表：加 `side`（0-4 条：关键数字 / 边界条件 / "
                                  "待确认点）把右侧用起来，表格会自动收窄；"
                                  "表格放不下的口径、来源、待办都放这里"))
    elif t == "chart":
        has_detail = bool(getattr(s, "detail_headers", None) and s.detail_rows)
        blocks = 1 + (1 if has_detail else 0) + len(s.insights) // 2
        if blocks < 2 or (len(s.chart.categories) <= 5 and not has_detail):
            warnings.append(_item(slide_no, "chart", "detail_rows",
                                  f"图表只有 {len(s.chart.categories)} 个数据点，"
                                  f"撑不满内容区、下半页会空：用 `detail_headers` + "
                                  f"`detail_rows` 在图表下面补一张明细表"
                                  f"（指标 / 目标值 / 口径 / 验证方式），"
                                  f"也可以把 insights 写到 4 条"))
    elif t == "figure":
        # 原图本身是"第一块信息"，图注/要点是第二块 —— 两样都没有就是一页孤图。
        if not s.caption and not s.points:
            warnings.append(_item(slide_no, "figure", "caption",
                                  "本页只有一张图：补 `caption`（一句话说清图在证明什么）"
                                  "或 `points`（0-3 条图旁要点），把图和结论绑在一起"))
    elif t in ("compare_3col", "numbered_cards", "hierarchy", "layered_stack",
               "flow", "panel_2col", "timeline", "funnel", "key_groups"):
        if not getattr(s, "callout", None) and not getattr(s, "footer_note", None):
            # 三栏对比**禁 callout**（`CompareSlide` 的 validator 会直接抛 error）——
            # 指引必须按版式给对：旧文案一刀切"补 callout"，模型照做就撞 schema，白跑一轮。
            if t == "compare_3col":
                warnings.append(_item(
                    slide_no, t, "footer_note",
                    "本页没有收尾：三栏并排本身已是三块信息、**不支持 `callout`**，"
                    "写一句 `footer_note`（≤64 字，放结论或口径）即可"))
            else:
                warnings.append(_item(
                    slide_no, t, "callout",
                    "本页只有卡片、没有底部辅区：补 `callout`（一组关键数字或"
                    "待确认点）或至少写一句 `footer_note`，把卡片放不下的信息落地"))
    elif t == "matrix":
        # matrix 的格子铺满内容区，塞不下 callout —— 只要求 footer_note 收尾
        if not s.footer_note:
            warnings.append(_item(
                slide_no, t, "footer_note",
                "本页只有一张矩阵：写一句 `footer_note`（≤64 字，放交叉结论或"
                "读法说明，如「右上角是优先投入区」），别让矩阵孤零零收尾"))


# ---------------- 正文密度下限（防"精简过头"） ----------------
# 反馈：LLM 一旦知道"要高密度、要精炼"，就容易把正文压成短语 —— 版面看着满了，
# 信息却被抽干。下面给每类版式的正文一个下限，低于就点名：
# **改写润色可以，抽掉论据/机制/口径不行**。
_COMPARE_DESC_MIN = 45      # 三栏卡每栏描述下限（schema 下限 20 太松）
_CARD_DESC_MIN = 40         # 编号卡片描述下限
_LAYER_DESC_MIN = 30        # 层级堆栈层描述下限
_STEP_DESC_MIN = 16         # 流程步骤描述下限
_NODE_DESC_MIN = 20         # 层级图子模块描述下限
_PANEL_SUM = (16, 110)      # 模块顶部总结陈述段区间
_PANEL_PT_MAX = 44          # 模块内单条要点上限
_STATEMENT_BODY_MIN = 150   # statement 正文下限
_MILESTONE_DESC_MIN = 25    # 时间轴节点描述下限
_FUNNEL_DESC_MIN = 20       # 漏斗层说明下限
_MATRIX_CELL_DESC_MIN = 15  # 矩阵格子说明下限
_KG_POINT_MIN = 8           # 要点分组的单条要点下限


def _check_density_floor(slide_no, s, warnings):
    """正文不许被精简过头：字段写到接近安全上限，而不是压成一句短语。"""
    t = s.slide_type
    if isinstance(s, StatementSlide) and len(s.body) < _STATEMENT_BODY_MIN:
        warnings.append(_item(slide_no, t, "body",
                              f"正文只有 {len(s.body)} 字（{_STATEMENT_BODY_MIN}+ 才够），"
                              f"精简过头了：结论页左侧正文区能装 400+ 字（上限已提到 520），"
                              f"请把论据、机制、口径写回连贯段落"
                              f"（内容确实少就并页或拆到 panel_2col，而不是把正文抽干）"))
    if isinstance(s, CompareSlide):
        thin = [c.label for c in s.columns if len(c.desc) < _COMPARE_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "columns",
                                  f"栏「{'、'.join(thin)}」描述不足 {_COMPARE_DESC_MIN} 字："
                                  f"补机制、数字、口径，别只留一句结论"))
    if isinstance(s, NumberedCardsSlide):
        thin = [c.name for c in s.cards if len(c.desc) < _CARD_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "cards",
                                  f"卡片「{'、'.join(thin)}」描述不足 {_CARD_DESC_MIN} 字："
                                  f"补机制、数字或边界条件（描述是卡片主内容）"))
    if isinstance(s, LayeredStackSlide):
        thin = [l.label for l in s.layers if len(l.desc) < _LAYER_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "layers",
                                  f"层「{'、'.join(thin)}」描述不足 {_LAYER_DESC_MIN} 字："
                                  f"补代表技术、作用或口径，别只留一个名词"))
    if isinstance(s, FlowSlide):
        thin = [st.label for st in s.steps if len(st.desc) < _STEP_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "steps",
                                  f"步骤「{'、'.join(thin)}」描述不足 {_STEP_DESC_MIN} 字："
                                  f"写清这一步的输入/输出或判定条件"))
    if isinstance(s, HierarchySlide):
        thin = [n.label for n in s.nodes if len(n.desc) < _NODE_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "nodes",
                                  f"子模块「{'、'.join(thin)}」描述不足 {_NODE_DESC_MIN} 字："
                                  f"补包含关系或职责说明"))
    if isinstance(s, TimelineSlide):
        thin = [ms.label for ms in s.milestones if len(ms.desc) < _MILESTONE_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "milestones",
                                  f"节点「{'、'.join(thin)}」描述不足 {_MILESTONE_DESC_MIN} 字："
                                  f"写清这一步交付了什么/发生了什么变化"))
    if isinstance(s, FunnelSlide):
        thin = [st.label for st in s.stages if len(st.desc) < _FUNNEL_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "stages",
                                  f"层「{'、'.join(thin)}」说明不足 {_FUNNEL_DESC_MIN} 字："
                                  f"写口径、转化率或流失原因（相邻两层一除就是转化率）"))
    if isinstance(s, MatrixSlide):
        thin = [f"{row.label}×{s.col_labels[c_i]}"
                for row in s.rows for c_i, cell in enumerate(row.cells)
                if len(cell.desc) < _MATRIX_CELL_DESC_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "rows",
                                  f"格子「{'、'.join(thin)}」说明不足 "
                                  f"{_MATRIX_CELL_DESC_MIN} 字：写清交叉处的结论或判据"))
    if isinstance(s, KeyGroupsSlide):
        thin = [f"{grp.title}#{k}"
                for grp in s.groups for k, pt in enumerate(grp.points)
                if len(pt) < _KG_POINT_MIN]
        if thin:
            warnings.append(_item(slide_no, t, "groups",
                                  f"要点「{'、'.join(thin)}」不足 {_KG_POINT_MIN} 字："
                                  f"写成「结论 + 数字/口径」，不是光秃秃的名词"))


def _check_figure(slide_no, s, errors, warnings):
    """图示页：图必须真的放得上去，图注/要点是"第二块信息"。

    用户要求：Word 里的 EMF/WMF **原样**进 PPT。渲染侧由
    `layouts.add_image` → `_add_vector_picture` 保证（字节原样、不转格式），
    这里只查"图在不在"和"这一页是不是只有一个容器"。
    """
    path = layouts.resolve_image_path(s.image)
    if not os.path.isfile(path):
        errors.append(_item(
            slide_no, "figure", "image",
            f"图片不存在或读不到：{s.image}（按仓库根解析为 {path}）。"
            f"Word 里的图先跑 `python .pi/renderer/docx_in.py <输入.docx> -o .pi/work/input.md`，"
            f"它会去修订并把 EMF/WMF 原样抽到 .pi/work/media/，"
            f"再把 `media/NN-xxx.emf` 写进 `image`"))
        return
    size = layouts.image_size_in(path)
    if size is None:
        warnings.append(_item(
            slide_no, "figure", "image",
            f"读不出图片尺寸：{s.image} —— 渲染会按图框铺满（可能被拉伸）。"
            f"EMF/WMF 会从 rclFrame 读尺寸；若这是别的格式，请确认文件没坏"))
    if not s.caption and not s.points:
        warnings.append(_item(
            slide_no, "figure", "caption",
            "本页只有一张图：补 `caption`（一句话说清这张图在证明什么）或 `points`"
            "（0-3 条图旁要点，放右侧纵栏）—— 整页只有一个容器，信息密度上不去"))
        return
    # 右栏只有 1~2 条要点时，卡片刻度到最大也填不满 —— 实测会空出 2in 空气。
    if s.points:
        fr, _ = layouts.figure_area(s.summary, bool(s.caption), True)
        if fr["mode"] == "rail":
            _rh = fr["rail_bottom"] - fr["rail_top"]
            _need = sum(layouts.point_rail_need(s.points, fr["rail_w"])) \
                + 0.14 * (len(s.points) - 1)
            _idle = _rh - _need * layouts.RAIL_GROW_MAX
            if _idle > 0.90:
                warnings.append(_item(
                    slide_no, "figure", "points",
                    f"右侧要点栏高 {_rh:.1f} in，{len(s.points)} 条要点写满也就占 "
                    f"{_need * layouts.RAIL_GROW_MAX:.1f} in，还剩 {_idle:.1f} in 空着："
                    f"补到 3 条要点（每条 30 字以上，说清「图在证明什么 → 所以呢」），"
                    f"或把结论写进 caption 后改用 panel_2col 与相邻内容并页"))


def _check_panels(slide_no, s, warnings):
    """双模块页：每个模块的顶部总结段与内容都要够 —— 两个空框并排还是空。"""
    for j, p in enumerate(s.panels):
        n = len(p.summary or "")
        if n < _PANEL_SUM[0] or n > _PANEL_SUM[1]:
            warnings.append(_item(slide_no, "panel_2col", f"panels[{j}].summary",
                                  f"模块「{p.label}」顶部总结陈述段 {n} 字"
                                  f"（要求 {_PANEL_SUM[0]}~{_PANEL_SUM[1]} 字）："
                                  f"一两句说清这个模块的结论与依据"))
        over = [x for x in p.points if len(x) > _PANEL_PT_MAX]
        if over:
            warnings.append(_item(slide_no, "panel_2col", f"panels[{j}].points",
                                  f"模块「{p.label}」有 {len(over)} 条要点超过 "
                                  f"{_PANEL_PT_MAX} 字：要点要短，细节挪进迷你表或总结段"))
        if p.rows:
            # 列宽与单元格容量都走渲染同款口径（layouts 的 PANEL_GAP/PANEL_PAD/PANEL_TBL_CELL）
            _tw = (12.23 - layouts.PANEL_GAP) / 2 - 2 * layouts.PANEL_PAD
            _w = col_widths(p.headers, p.rows, _tw, min_w=0.62)
            for r, row in enumerate(p.rows):
                for c, val in enumerate(row):
                    cap = max(6, layouts.cell_cap_chars(_w[c], 0.22,
                                                        layouts.PANEL_TBL_CELL))
                    if len(val) > 2 * cap:
                        warnings.append(_item(
                            slide_no, "panel_2col", f"panels[{j}].rows[{r}][{c}]",
                            f"模块「{p.label}」单元格 {len(val)} 字，是本列容量"
                            f"（约 {cap} 字/行）的 2 倍以上：压缩到 {2 * cap} 字以内，"
                            f"或改成要点列表"))


def _iter_icons(s):
    """一页里所有 icon 字段：[(字段路径, 图标名), ...]（渲染器按这些位置画块头图标）。"""
    t = s.slide_type
    out: list[tuple[str, str | None]] = []
    if t == "toc":
        out += [(f"items[{i}].icon", it.icon) for i, it in enumerate(s.items)]
    elif t == "section":
        out.append(("icon", s.icon))
    elif t == "statement":
        if s.highlight:
            out.append(("highlight.icon", s.highlight.icon))
        out += [(f"points[{i}].icon", p.icon) for i, p in enumerate(s.points or [])]
    elif t == "compare_3col":
        out += [(f"columns[{i}].icon", c.icon) for i, c in enumerate(s.columns)]
    elif t == "numbered_cards":
        out += [(f"cards[{i}].icon", c.icon) for i, c in enumerate(s.cards)]
    elif t == "layered_stack":
        out += [(f"layers[{i}].icon", l.icon) for i, l in enumerate(s.layers)]
    elif t == "flow":
        out += [(f"steps[{i}].icon", st.icon) for i, st in enumerate(s.steps)]
    elif t == "hierarchy":
        out += [(f"nodes[{i}].icon", n.icon) for i, n in enumerate(s.nodes)]
    elif t == "table":
        out += [(f"side[{i}].icon", n.icon) for i, n in enumerate(s.side or [])]
    elif t == "chart":
        out += [(f"insights[{i}].icon", n.icon) for i, n in enumerate(s.insights)]
    elif t == "panel_2col":
        out += [(f"panels[{i}].icon", p.icon) for i, p in enumerate(s.panels)]
    elif t == "timeline":
        out += [(f"milestones[{i}].icon", ms.icon) for i, ms in enumerate(s.milestones)]
    elif t == "funnel":
        out += [(f"stages[{i}].icon", st.icon) for i, st in enumerate(s.stages)]
    elif t == "matrix":
        out += [(f"rows[{r}].cells[{c}].icon", cell.icon)
                for r, row in enumerate(s.rows) for c, cell in enumerate(row.cells)]
    elif t == "key_groups":
        out += [(f"groups[{i}].icon", grp.icon) for i, grp in enumerate(s.groups)]
    return [(f, n) for f, n in out if n]


def _cap_note(slide_no, s, field, label, n_chars: int, width: float, height: float,
              warnings, down: list, steps: tuple | None = None) -> int:
    """按全局降档序列核对一段文字的容量（**只出 warning，不拦渲染**）。

    适配是渲染器的事：它会把字号降到放得下（最低 10pt）。所以这里只在
    "降到最小档仍放不下" 时点名；没超容量但落了档的页面记进 `down`（供人工抽检，
    不刷 warning）。返回最小档容量（字数）。
    """
    cap = layouts.cap_chars(width, height, steps)
    size, _sp = layouts.step_for_len(n_chars, width, height, steps)
    top = (steps or layouts.FIT_STEPS)[0][0]
    if n_chars > cap:
        warnings.append(_item(
            slide_no, s.slide_type, field,
            f"{label} {n_chars} 字，降到最小字号档（{size:g}pt）仍放不下（约 {cap} 字）："
            f"删到 {cap} 字以内，或挪进 `side` / 明细表 / 相邻页"))
    elif size < top:
        down.append({"slide": slide_no, "type": s.slide_type, "field": field,
                     "size": size, "chars": n_chars})
    return cap


# ---------------- 卡片标题即重点 ----------------
# 反馈：「无法图表化的内容就用卡片 + 加粗重点 Title」。标题要写成**判断句或带数字的结论**
# （「统一任务集管理」「回归通过率 62%→88%」），不能是「任务集管理」这种名词短语 ——
# 名词标题读不出任何结论，还额外占掉一层高度（再跟一句 sub 就是三层开销）。
# 这里是启发式：既没数字、也没有判断性谓语，而且很短 or 落在名词黑名单里的，点名。
_CLAIM_SUFFIX = ("管理", "介绍", "概述", "简介", "说明", "清单", "列表", "情况", "能力",
                 "机制", "体系", "平台", "模块", "功能", "流程", "规范", "方案", "建设",
                 "工作", "内容", "配置", "架构", "设计", "实践", "方法", "工具", "指标",
                 "数据", "系统", "服务", "组件", "节点", "策略", "模式", "原则", "分类")
_CLAIM_VERB = ("提升", "降低", "减少", "增加", "下降", "高于", "低于", "优于", "统一",
               "收敛", "隔离", "复用", "回滚", "灰度", "分级", "分层", "先", "再", "才",
               "越", "靠", "把", "让", "必须", "禁止", "取决", "是否", "快", "慢", "用")


def _is_claim(name: str) -> bool:
    """粗判这条标题是不是「判断句 / 带数字的结论」。"""
    s = (name or "").strip()
    if any(ch.isdigit() for ch in s):
        return True
    return any(v in s for v in _CLAIM_VERB)


def _check_claim_titles(slide_no, s, warnings):
    """卡片标题要写成判断句 / 带数字的结论 —— 标题就是这张卡的重点。"""
    if not isinstance(s, NumberedCardsSlide):
        return
    bad = []
    for c in s.cards:
        nm = (c.name or "").strip()
        if _is_claim(nm):
            continue
        if _cjk_len(nm) <= 8 or any(nm.endswith(t) for t in _CLAIM_SUFFIX):
            bad.append(nm)
    if bad:
        warnings.append(_item(
            slide_no, "numbered_cards", "cards[].name",
            "卡片标题「" + "」「".join(bad) + "」读不出结论：标题就是这张卡的**重点**，"
            "写成判断句或带数字的结论（如「统一任务集管理」「回归通过率 62%→88%」），"
            "别只写名词短语 —— 名词标题 + 一句话定位两层占高，信息量反而最少"))


# ---------------- 块头标题 vs 图标：不许叠在一起 ----------------
# 用户实测（v22 P9 第 3/5 张卡）：3 列编号卡的块头右上角有图标，标题一路写过去
# 直接压在图标上。渲染侧已经给图标让位（`block_title_w`），但**标题仍然要短** ——
# 长标题被缩字号甚至折成两行同样难看，所以这里按同一套宽度口径点名。
def _check_block_titles(slide_no, s, warnings):
    """块头标题要给右上角图标让位，否则压图标/折行。"""
    if isinstance(s, NumberedCardsSlide):
        _g = layouts.numbered_geom(s)
        cw, has_icon = _g["cw"], _g["has_icon"]
        # 4 列窄卡不画图标，但**标题照样有宽度上限**（卡宽 - 内边距）——
        # 旧实现在这里直接 return，7~8 张卡的 4 列排版正好落在盲区里。
        avail = (layouts.block_title_w(cw, layouts.NO_GUTTER, True) if has_icon
                 else layouts.block_title_w(cw, layouts.NO_GUTTER, False))
        base = FONT_SIZES["card_title"]
        for c in s.cards:
            if layouts.text_w_in(c.name, base) <= avail:
                continue
            cap = layouts.chars_per_line(avail, base)
            if has_icon:
                _why = (f"右侧要给图标让出 "
                        f"{layouts.CARD_ICON_RIGHT + layouts.ICON_TITLE_GAP:.2f}in，"
                        f"可用 {avail:.2f}in（约 {cap} 字/行）—— 现在会压到图标或折成两行")
            else:
                _why = (f"{len(s.cards)} 张卡排成 4 列、每卡只有 {cw:.2f}in 宽，"
                        f"标题可用 {avail:.2f}in（约 {cap} 字/行）—— 现在会折成两行")
            warnings.append(_item(
                slide_no, "numbered_cards", "cards[].name",
                f"卡片标题「{c.name}」放不下：{_why}。砍到 {cap} 字以内，机制细节挪到 desc"))
    elif isinstance(s, CompareSlide):
        cw = layouts.compare_desc_box(s.summary, 0, False,
                                      n_cols=len(s.columns))[0] + 0.56
        avail = layouts.compare_label_w(cw, True)
        base = FONT_SIZES["card_title"]
        for c in s.columns:
            if layouts.text_w_in(c.label, base) <= avail:
                continue
            cap = layouts.chars_per_line(avail, base)
            warnings.append(_item(
                slide_no, "compare_3col", "columns[].label",
                f"栏名「{c.label}」居中排、两侧要给块头图标让位，可用只有 {avail:.2f}in"
                f"（约 {cap} 字）—— 会压到图标。砍到 {cap} 字以内，或把结论放进 tag"))


def _check_icons(slide_no, s, warnings):
    """图标名必须真的能画出来。

    旧实现里写错的名字（拼错的中文别名）会**静默丢图标**：调用方拿到非空返回值就不再
    兜底，块头空着而且没有任何提示。现在渲染侧会兜底并退化，qc 侧负责把它点名出来，
    免得同一个错法在整份 deck 里反复出现（P6「Git 提交记录」= 别名「版本」没登记）。
    """
    for field, name in _iter_icons(s):
        if layouts.icon_resolvable(name):
            continue
        warnings.append(_item(
            slide_no, s.slide_type, field,
            f"图标「{name}」在 .pi/assets/icons 里找不到（本次渲染会按块内文字自动补一个，"
            f"但别再这么写）：填 Tabler 英文名（versions / upload / git-branch / database…）"
            f"或 .pi/assets/icons/aliases.json 里已登记的中文别名"))


def _brief_checks(deck, errors, warnings, source_path: str | None = None):
    """按 deck.brief 指定的写作方向查红线。

    只出 warning —— 这些是方向的取舍约定，不是几何错误，不该拦渲染。
    规约写在 .pi/renderer/briefs/<name>.yaml 的 check 段，加方向不用改这里。
    """
    name = getattr(deck, "brief", None)
    if not name:
        return
    try:
        brief = briefs.load(name)
    except FileNotFoundError as e:                     # 名字写错值得拦
        errors.append(_item("-", "deck", "brief", str(e)))
        return

    c = brief.get("check") or {}
    label = brief.get("label", name)
    n = len(deck.slides)

    # 页数：研讨型不设上限（minPages 只兜底），汇报型是硬区间
    lo, hi = c.get("minPages"), c.get("maxPages")
    if lo and n < lo:
        warnings.append(_item("-", "deck", "brief",
                              f"「{label}」要求至少 {lo} 页，当前 {n} 页偏少：这个方向"
                              f"要求逐条覆盖、缺口明写，页数少通常意味着替读者做了取舍"))
    if hi and n > hi:
        warnings.append(_item("-", "deck", "brief",
                              f"「{label}」要求 {lo}~{hi} 页，当前 {n} 页偏多："
                              f"先看是不是有内容重复。**不要靠删内容达标** —— "
                              f"宁可拆成多场/多份，页面上已覆盖的论据一条都不能少"))
    ask = c.get("askAbove")
    if ask and n > ask:
        warnings.append(_item("-", "deck", "brief",
                              f"「{label}」页数不该超过 {ask} 页（当前 {n} 页）："
                              f"这类取舍**先问用户**（只讲某几章？拆成多份？还是转汇报型？），"
                              f"但要并进**大纲评审那一次**问完 —— 慢档只在那里打断用户一次，"
                              f"出图后不再另开一轮问答。**无论怎么选，都不许删内容**"))

    # 输入覆盖度：研讨型要求"交稿前核对一遍输入有没有被丢"，那就不能让它变成可选项
    if c.get("requireSource") and not source_path:
        warnings.append(_item(
            "-", "deck", "source",
            f"「{label}」要求交稿前核对**输入覆盖度**：请用 "
            f"`--qc --source .pi/work/analysis.md` 重跑一次 —— "
            f"丢内容是最难被抓到的偷懒方式（版式装不下 → 删两页 → 页内检查全绿）"))

    # 图表密度：研讨型的「尽量每页都有图表表格」不能只写在规范里
    if c.get("requireLayouts"):
        missing = [t for t in c["requireLayouts"]
                   if not any(s.slide_type == t for s in deck.slides)]
        if missing:
            warnings.append(_item("-", "deck", "brief",
                                  f"「{label}」要求至少有 {'、'.join(c['requireLayouts'])} 版式，"
                                  f"本 deck 缺：{'、'.join(missing)}"))
    ratio_min = c.get("chartTableRatio")
    if ratio_min:
        content = [s for s in deck.slides
                   if s.slide_type not in ("cover", "toc", "section", "end")]
        if content:
            hits = sum(1 for s in content if s.slide_type in ("table", "chart"))
            ratio = hits / len(content)
            if ratio < ratio_min:
                warnings.append(_item(
                    "-", "deck", "brief",
                    f"「{label}」要求内容页里「表格 + 图表」至少占 {ratio_min:.0%}，"
                    f"当前 {hits}/{len(content)}（{ratio:.0%}）偏低：优先把数值、占比、"
                    f"明细对照改成 chart/table，实在没有可图表化的数据就用两列表格兜底"))
    distinct = c.get("minDistinctChartTypes")
    if distinct:
        types = {s.chart.type for s in deck.slides if s.slide_type == "chart"}
        if len(types) < distinct:
            warnings.append(_item("-", "deck", "brief",
                                  f"「{label}」要求图表「各种」——不同图型至少 {distinct} 种，"
                                  f"当前只有 {len(types)} 种（{'、'.join(sorted(types)) or '无'}）："
                                  f"柱/条/折线/饼/环/雷达换着用"))

    cap = c.get("maxSupportsPerPage")
    if cap:
        for i, s in enumerate(deck.slides, 1):
            # **只管 `points`**（statement 的支撑论据）。不能顺手管 `chart.insights`：
            # schema 要求 insights 2~4 条（必填），briefing 的 3 条上限会把合法的 4 条
            # 判成超标 —— 局部规则压过全局规则的典型。图表那边由 schema 自己约束。
            items = getattr(s, "points", None) or []
            if len(items) > cap:
                warnings.append(_item(i, s.slide_type, "supports",
                                      f"「{label}」每页支撑论据不超过 {cap} 条，"
                                      f"本页有 {len(items)} 条，请把多出的写进正文"
                                      f"或另起一页（别删论据）"))

    black = c.get("titleBlacklist") or []
    if black:
        for i, s in enumerate(deck.slides, 1):
            title = getattr(s, "title", None)
            if not title:
                continue
            hit = [w for w in black if w in title]
            if hit:
                warnings.append(_item(i, s.slide_type, "title",
                                      f"标题「{title}」带了无信息量的词「{'/'.join(hit)}」："
                                      f"「{label}」要求标题写成一句主张"
                                      f"（如「资源应继续向县域倾斜」），不是话题"))


# ---------------- 输入覆盖度：不许悄悄丢内容 ----------------
# 研讨型的硬要求是"输入里的内容一条都不能少"，而**丢内容恰恰是最难被抓到的偷懒方式**：
# 版面装不下 → 删掉两页 → --qc 全绿（其它检查都是"页内"的，看不出整块内容没进来）。
# 这里拿**输入分析**（`.pi/work/analysis.md` 那种表）交叉核对两个最可量化的口径：
#   1. **数字口径**：带 %/单位、或 3 位以上的指标数字 —— 输入里有、deck 里找不到的，点名；
#   2. **模块名**：表格里的模块名，它的 4-gram 在 deck 文本里一个都不出现的，点名。
# 两者都是 warning：改写措辞、写成中文数字、模块合并表述都可能造成假阳性，
# 但**"页数少了、数字却不见了"正是丢内容的信号**，值得人工核一遍。
_NUM_RE = re.compile(r"\d+(?:\.\d+)?(?:%|％|[个家条张层步倍次人天月年万亿])?")
_MODULE_MIN = 4                # 模块名按 4-gram 判"有没有出现过"


def _norm_text(s: str) -> str:
    """比对前的归一化：去掉加粗标记/空白/千分位，让"同一句话的不同写法"能对上。"""
    s = ((s or "").replace("**", "").replace("`", "")
         .replace(" ", "").replace("　", "").replace("\u3000", "")
         .replace("％", "%"))
    return re.sub(r"(?<=\d),(?=\d)", "", s)      # 1,204 → 1204


def _source_rows(text: str) -> list[list[str]]:
    """从输入分析（markdown 表格）里取数据行：[章节, 模块, 核心结论, 论据, 缺口]。"""
    rows: list[list[str]] = []
    for line in (text or "").splitlines():
        s = line.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if len(cells) < 3:
            continue
        joined = "".join(cells).replace("-", "").replace(":", "").strip()
        if not joined or cells[0].startswith("章节"):
            continue
        rows.append(cells)
    return rows


def _source_numbers(rows: list[list[str]]) -> list[tuple[str, str]]:
    """[(指标数字, 出处模块)]：只留"像论据"的（带 %/单位，或 3 位以上，且不是年份）。"""
    out: list[tuple[str, str]] = []
    for cells in rows:
        field = " ".join(cells[2:4])
        for m in _NUM_RE.finditer(field):
            tok = m.group(0)
            digits = re.sub(r"\D", "", tok)
            unit = tok[len(digits):]
            if not unit and len(digits) < 3:
                continue
            if not unit and 1900 <= int(digits) <= 2100:
                continue          # 年份不是论据
            out.append((tok, cells[1] if len(cells) > 1 else ""))
    return out


def _module_covered(name: str, hay: str) -> bool:
    """模块名在 deck 全文里"露过面"就算覆盖：任一 4-gram 命中即可（抗改写）。"""
    s = re.sub(r"[（(].*?[)）]", "", _norm_text(name))
    if len(s) < _MODULE_MIN:
        return s in hay
    return any(s[i:i + _MODULE_MIN] in hay for i in range(len(s) - _MODULE_MIN + 1))


def check_source_coverage(deck, source_text: str) -> dict:
    """输入覆盖度：输入分析里的模块与数字口径，deck 覆盖了没有。

    渲染/`--qc` 只跑"页内"检查，看不出"整块内容没进 deck"；这里是唯一能反向
    抓住"为了压页数丢内容"的机制。返回 {"rows":…, "numbers_missing":[…], …}。
    """
    rows = _source_rows(source_text)
    hay = _norm_text(json.dumps(deck.model_dump(mode="json"), ensure_ascii=False))
    nums = _source_numbers(rows)
    miss_nums: list[str] = []
    for tok, _mod in nums:
        if _norm_text(tok) not in hay and tok not in miss_nums:
            miss_nums.append(tok)
    miss_mods: list[str] = []
    for cells in rows:
        name = cells[1] if len(cells) > 1 else ""
        if not name or name in ("无", "-", "—"):
            continue
        if not _module_covered(name, hay) and name not in miss_mods:
            miss_mods.append(name)
    return {"rows": len(rows), "numbers": len(nums),
            "numbers_missing": miss_nums, "modules_missing": miss_mods}


def _check_source_coverage(deck, source_path: str | None, errors, warnings) -> dict | None:
    if not source_path:
        return None
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = source_path if os.path.isfile(source_path) else os.path.join(root, source_path)
    if not os.path.isfile(path):
        errors.append(_item("-", "deck", "source",
                            f"--source 指定的输入分析读不到：{source_path}"))
        return None
    with open(path, encoding="utf-8") as f:
        text = f.read()
    cov = check_source_coverage(deck, text)
    if cov["numbers_missing"]:
        head = "、".join(cov["numbers_missing"][:6])
        warnings.append(_item(
            "-", "deck", "source",
            f"输入里有 {len(cov['numbers_missing'])} 处数字口径在 deck 里找不到"
            f"（共 {cov['numbers']} 处，如：{head}）：数字是论据的核心 —— "
            f"**每处要么落到页面上，要么确认已被别处覆盖**"
            f"（可能写成了中文数字或改了表述）。**不许为了控页数把它们删掉**："
            f"版面不够就加页/拆页。"))
    if cov["modules_missing"]:
        head = "、".join(cov["modules_missing"][:6])
        warnings.append(_item(
            "-", "deck", "source",
            f"输入分析里有 {len(cov['modules_missing'])} 个模块在 deck 里没有露面"
            f"（如：{head}）：**这就是「丢内容」** —— 补一页 / 并进相邻页 / 换版式重排，"
            f"都不要删。回到 `.pi/work/analysis.md` 逐条核对后再交稿。"))
    # 页数 vs 内容量：一轮就能看出"是不是靠少讲几页换来的紧凑"
    _content = {"statement", "compare_3col", "numbered_cards", "layered_stack",
                "flow", "hierarchy", "table", "chart", "panel_2col"}
    pages = sum(1 for s in deck.slides if s.slide_type in _content) or 1
    if cov["rows"] >= 6 and cov["rows"] / pages > 6:
        warnings.append(_item(
            "-", "deck", "source",
            f"输入分析有 {cov['rows']} 条论据行，本 deck 只有 {pages} 个内容页"
            f"（平均每页要扛 {cov['rows'] / pages:.1f} 条）：研讨型的密度是"
            f"「加页 + 每页铺满」抬起来的，不是靠少讲几页 —— 先确认没有整块内容被丢掉，"
            f"缺的内容**加页**补回来"))
    return cov


# ---------------- 报告瘦身：同一规则的重复说明只讲一次 ----------------
# 质检报告是**整份进模型上下文**的（实测 30 页 deck 16.6KB，其中 13KB 是 message 文本），
# 而同一规则在多页重复时，后半段"怎么改"是一字不差的 —— `table.rows` 报了 12 条、
# 每条 146 字，其中 95 字是同一句「把下一章节/关联章节的明细并进来…」。逐页复述 12 遍
# 既是白烧 token，也让模型更难看出"这 12 条是同一个改法"。
# 这里做**纯机械**的公共后缀折叠：同一 (type, field) 里共享 ≥ `_TAIL_MIN` 字后缀的条目，
# 第一条留完整说明，其余只留前半段（页内数字/尺寸）+「（规则见第 N 页）」。
# 不猜语义、不改条目结构（照样有 slide/type/field/message），只是不再复述。
_TAIL_MIN = 24
# 成对括号：折叠时不许把消息切在半个括号上（见 _trim_head）
_BRACKETS = {"（": "）", "「": "」", "【": "】", "《": "》", "(": ")"}


def _trim_head(head: str) -> str:
    """把折出来的前半段收在干净的边界上。

    公共后缀可能正好从「（约能再放 3 行）」的右括号起算，切完会留下
    「表格 5 行、还空 1.6 in（约能再放 3 行」这种半个括号的残句 —— 比不复述更难读。
    所以这里裁掉结尾**未闭合**的括号段，再把结尾标点收干净。
    """
    for _ in range(4):
        head = head.rstrip("：:；;，,、—-… ")
        cut = -1
        for i, ch in enumerate(head):
            if ch in _BRACKETS and _BRACKETS[ch] not in head[i:]:
                cut = i
                break
        if cut < 0:
            break
        head = head[:cut]
    return head.rstrip("：:；;，,、—-… ")


def _common_suffix_len(a: str, b: str) -> int:
    n = 0
    for x, y in zip(reversed(a), reversed(b)):
        if x != y:
            break
        n += 1
    return n


def _fold_rule_tails(items: list[dict]) -> int:
    """原地折叠重复的规则说明，返回省掉的字符数。"""
    anchors: dict[tuple, list[tuple[str, int]]] = {}
    saved = 0
    for it in items:
        msg = it["message"]
        pool = anchors.setdefault((it["type"], it["field"]), [])
        best, src = 0, None
        for other, page in pool:
            n = _common_suffix_len(msg, other)
            if n > best:
                best, src = n, page
        if src is not None and best >= _TAIL_MIN:
            head, tail = msg[:len(msg) - best], msg[len(msg) - best:]
            # 别把数值和它的单位切开：切点常落在「还空 1.6 | in（约能再放 3 行）」之间，
            # 把尾部开头的数字/单位拉回前半段（含空格，避免留下 "1.6 in" 的半截）
            if head and head[-1].isdigit():
                m = re.match(r"[A-Za-z0-9./% ]+", tail)
                if m and m.group(0).strip():
                    head += m.group(0)
            head = _trim_head(head)
            label = f"第 {src} 页" if isinstance(src, int) else "同类上一条"
            new = (head + f"（规则见{label}）") if head else f"（规则见{label}）"
            saved += len(msg) - len(new)
            it["message"] = new
        else:
            pool.append((msg, it["slide"]))
    return saved


def _slim_warnings(warnings: list[dict]) -> int:
    """快档报告瘦身：同一 (type, field) 的重复项只留第一条，其余折成「同类共 N 条」。

    **两类不折**：`source`（丢内容）与 `deck` 级（页数、方向红线）—— 它们一条就是
    一件事、而且是最严重的那种，折掉等于把"内容被删了"藏起来。

    返回瘦身后的条数。顺序按原报告保持（每类第一条所处的位置）。
    """
    protected = {"source", "brief"}
    order = {id(w): i for i, w in enumerate(warnings)}
    head: list[dict] = []
    groups: dict[tuple, list[dict]] = {}
    for w in warnings:
        if w["field"] in protected or w["type"] == "deck":
            head.append(w)
        else:
            groups.setdefault((w["type"], w["field"]), []).append(w)
    slimmed = list(head)
    for items in groups.values():
        first = items[0]
        if len(items) > 1:
            pages = "、".join(str(i["slide"]) for i in items[1:12])
            tail = "" if len(items) <= 12 else f" 等 {len(items) - 12} 页"
            first["message"] += (f"（快档每类只列 1 条，本类共 {len(items)} 条，"
                                 f"其余在第 {pages} 页{tail}；"
                                 f"要逐条清零请不带 --speed fast 重跑）")
        slimmed.append(first)
    warnings[:] = sorted(slimmed, key=lambda w: order[id(w)])
    return len(warnings)


def check(deck, source_path: str | None = None, speed: str = "slow") -> dict:
    """质检。`speed` 只改**报告详略**，不改检查项：

      - `slow`（默认）：全量列出，配合 `build.py` 退出码 3 要求 warning 归零；
      - `fast`        ：同类 warning 折成一条，报告短、可直接交付。

    **快档不许少查任何一项** —— 少查一项就漏掉密度/几何问题，那才是"排版异常"
    的真正来源；快档省的是**回炉轮次与上下文**，不是检查。
    """
    errors, warnings = [], []

    # 全局结构检查：一份完整 deck 应有封面和结尾
    types = [s.slide_type for s in deck.slides]
    if types and types[0] != "cover":
        warnings.append(_item(1, types[0], "deck", "第一页不是封面（cover），建议补上"))
    if "toc" not in types and len(types) > 6:
        warnings.append(_item(1, "deck", "deck", "超过 6 页但没有目录页（toc）"))
    if types and types[-1] != "end":
        warnings.append(_item(len(types), types[-1], "deck", "最后一页不是结尾页（end），建议补上"))

    # 外观：未登记的主题/卡片名会被确定性回退（不拦渲染），但必须点名 ——
    # 否则作者以为换了一张脸、实际出的还是默认长相，还查不出原因。
    if (deck.style or "") not in theme._THEMES:
        warnings.append(_item(
            1, "deck", "style",
            f"未登记的配色主题「{deck.style}」：已回退 {theme.DEFAULT_STYLE}，"
            f"可选 {'/'.join(theme._THEMES)}"))
    if (getattr(deck, "card", None) or "") not in theme.CARD_STYLES:
        warnings.append(_item(
            1, "deck", "card",
            f"未登记的卡片长相「{getattr(deck, 'card', None)}」：已回退 "
            f"{theme.DEFAULT_CARD}，可选 {'/'.join(theme.CARD_STYLES)}"))

    total_chars = 0
    chart_pages = 0
    down: list[dict] = []          # 落了字号档的页面（渲染正常，供抽检）
    for i, s in enumerate(deck.slides, 1):
        # 图标名写错会静默丢图标 —— 逐页点名，别让同一个错法在整份 deck 里反复出现
        _check_icons(i, s, warnings)
        # 卡片标题要写成判断句（标题即重点），不是名词短语
        _check_claim_titles(i, s, warnings)
        # 卡片骨架：关键句（第一句）+ 长文关键短语
        _check_card_structure(i, s, warnings)
        # 块头标题不许压到右上角图标（长标题）
        _check_block_titles(i, s, warnings)
        if isinstance(s, StatementSlide):
            total_chars += len(s.body)
            # 正文从 12.5pt 往下逐档降到 10pt；超了**不拦渲染**，
            # 只有降到最小档仍放不下才点名 —— 写满必须比写短更安全。
            _bw, _bh = statement_body_box(s)
            _cap_note(i, s, "body", "正文", len(s.body), _bw, _bh, warnings, down,
                      steps=layouts.STATEMENT_STEPS)
            # 正文区留白：容量上限提到 700 字后，写 200 字就会在左下空出一条
            _slack, _avail, _size = layouts.statement_slack(s)
            if _slack > 0.80:
                _more = layouts.slack_chars(_bw, _slack, _size)
                warnings.append(_item(
                    i, "statement", "body",
                    f"正文区下方还空 {_slack:.1f} in（可用 {_avail:.1f} in，本页 {_size:g}pt）："
                    f"大约还能写 {_more} 字 —— 把机制、论据、口径写满，或把相邻页同主题的"
                    f"内容并进来；内容确实不多就改用 `panel_2col` 与相邻页并成一页"))
            if "，" not in s.body and "。" not in s.body and len(s.body) < 100:
                errors.append(_item(i, "statement", "body",
                                    "正文疑似碎片化（无完整句读），请写成连贯段落"))
            # 视觉层级检查：HighlightBox 核心指标必须是大字突出
            if s.highlight:
                # 检查 value 是否有足够的"视觉冲击力"（数字/百分比/对比）
                import re as _re
                has_visual_impact = bool(_re.search(r'\d', s.highlight.value))
                if not has_visual_impact:
                    warnings.append(_item(
                        i, "statement", "highlight.value",
                        f"高亮指标「{s.highlight.value}」不含数字；视觉层级要求核心指标必须是"
                        f"可量化的大字（如 85%、99.9%、5x），不含数字的文案无法发挥"
                        f"18pt 大字的视觉冲击力"))
                # 高亮卡的 desc 以前没有任何容量检查：卡高写死 1.10in、desc 框固定
                # 0.30in，而 schema 允许 80 字（实测需 0.63in）—— 文字压出卡片下沿。
                # 现在卡高随行数长、口径统一在 highlight_h，仍放不下就得点名。
                if s.highlight.desc:
                    _hg = layouts.statement_geom(s)
                    _hl_w = _hg["rw"] - 0.66          # 卡内文本宽（按有图标，与 highlight_h 同口径）
                    _hl_h = max(0.24, _hg["hl_h"] - layouts.HL_DESC_TOP - 0.06)
                    _cap_note(i, s, "highlight.desc", "高亮卡说明",
                              len(s.highlight.desc), _hl_w, _hl_h, warnings, down)
                chart_pages += 1
        elif isinstance(s, CompareSlide):
            _co = getattr(s, "callout", None) or []
            _aux = layouts.callout_height(_co)
            for j, col in enumerate(s.columns):
                # 描述的可写高度随 summary 行数、要点条数与辅区落位变化，不能按常数算
                _dw, _dh, _ = compare_desc_box(s.summary, len(col.points), bool(_co),
                                               _aux, len(s.columns))
                _cap_note(i, s, f"columns[{j}].desc", f"第 {j+1} 栏描述",
                          len(col.desc), _dw, _dh, warnings, down)
                total_chars += len(col.desc)
            chart_pages += 1
        elif isinstance(s, NumberedCardsSlide):
            g = numbered_geom(s)
            for j, c in enumerate(s.cards):
                _cap_note(i, s, f"cards[{j}].desc", f"卡片 {c.no} 描述",
                          len(c.desc), g["desc_w"], g["desc_h"], warnings, down)
                total_chars += len(c.desc)
            chart_pages += 1
        elif isinstance(s, LayeredStackSlide):
            g = layered_geom(s)
            for j, layer in enumerate(s.layers):
                _cap_note(i, s, f"layers[{j}].desc", f"层「{layer.label}」描述",
                          len(layer.desc), g["desc_w"], g["desc_h"], warnings, down)
                total_chars += len(layer.desc)
            chart_pages += 1
        elif isinstance(s, FlowSlide):
            # 环形只在「有 loop_note 且 3≤步数≤FLOW_WRAP_AT」时成立（见 render_flow）。
            # 旧判据只看 loop_note and n>=3，于是 8~12 步带 loop_note 的页面
            # **实际走两行线性流程**却一律跳过容量检查 —— 溢出 0.62in 也零警告。
            if s.loop_note and 3 <= len(s.steps) <= layouts.FLOW_WRAP_AT:
                # 环形流程的节点描述框是固定的（窄卡片），不参与全局降档
                total_chars += sum(len(st.desc) for st in s.steps)
            else:
                g = flow_linear_geom(s)
                for j, st in enumerate(s.steps):
                    _cap_note(i, s, f"steps[{j}].desc", f"步骤「{st.label}」描述",
                              len(st.desc), g["desc_w"], g["desc_h"], warnings, down)
                    total_chars += len(st.desc)
            chart_pages += 1
        elif isinstance(s, QuadrantGridSlide):
            # 视觉层级检查：核心指标值必须是大字突出（≥18pt 在 theme.py FONT_SIZES 中定义）
            cells = [s.q_top_left, s.q_top_right, s.q_bottom_left, s.q_bottom_right]
            names = ["q_top_left", "q_top_right", "q_bottom_left", "q_bottom_right"]
            g = layouts.grid_geom(s)
            for j, cell in enumerate(cells):
                if not cell.value or len(cell.value.strip()) == 0:
                    warnings.append(_item(
                        i, "grid_2x2", names[j],
                        f"格子 {cell.title} 的核心指标值不能为空（视觉层级要求 18pt 大字突出）"))
                if not cell.desc or len(cell.desc.strip()) < 4:
                    warnings.append(_item(
                        i, "grid_2x2", names[j],
                        f"格子 {cell.title} 缺少说明文字（建议 4-60 字说明指标含义、口径、对比基线）"))
                elif cell.desc:
                    _cap_note(i, s, f"{names[j]}.desc", f"格子「{cell.title}」说明",
                              len(cell.desc), g["desc_w"], g["desc_h"], warnings, down)
                    total_chars += len(cell.desc)
            # 四格说明都只有一句话时，格子下半截会是空的 —— 不是几何问题，
            # 是信息量太少：四象限的价值在"每格给出判断依据"，不是四个标签。
            if all(len((c.desc or "").strip()) < 25 for c in cells):
                warnings.append(_item(
                    i, "grid_2x2", "q_top_left.desc",
                    "四个格子的说明都不到 25 字：格子会空出下半截。每格补上"
                    "「口径 / 对比基线 / 为什么落在这一格」（40~60 字最稳），"
                    "或把格子数量从 4 个减到 2×2 真正需要的两格"))
            chart_pages += 1
        elif isinstance(s, TimelineSlide):
            g = layouts.timeline_geom(s)
            for j, ms in enumerate(s.milestones):
                _cap_note(i, s, f"milestones[{j}].desc",
                          f"节点「{ms.label}」描述", len(ms.desc),
                          g["desc_w"], g["desc_h"], warnings, down)
                total_chars += len(ms.desc)
            chart_pages += 1
        elif isinstance(s, FunnelSlide):
            g = layouts.funnel_geom(s)
            for j, st in enumerate(s.stages):
                _cap_note(i, s, f"stages[{j}].desc", f"层「{st.label}」说明",
                          len(st.desc), g["desc_w"], g["desc_h"], warnings, down)
                total_chars += len(st.desc)
            chart_pages += 1
        elif isinstance(s, MatrixSlide):
            g = layouts.matrix_geom(s)
            for r_i, row in enumerate(s.rows):
                for c_i, cell in enumerate(row.cells):
                    _cap_note(i, s, f"rows[{r_i}].cells[{c_i}].desc",
                              f"「{row.label}×{s.col_labels[c_i]}」说明",
                              len(cell.desc), g["desc_w"], g["desc_h"],
                              warnings, down)
                    total_chars += len(cell.desc)
            chart_pages += 1
        elif isinstance(s, KeyGroupsSlide):
            g = layouts.key_groups_geom(s)
            # 卡宽/卡高逐张取（末行单卡跨整行时与常规卡不同），不能只看第一张
            for cg in g["cards"]:
                grp = s.groups[cg["i"]]
                n_pts = max(1, len(grp.points))
                for k, pt in enumerate(grp.points):
                    _cap_note(i, s, f"groups[{cg['i']}].points[{k}]",
                              f"组「{grp.title}」要点", len(pt),
                              cg["text_w"], max(0.18, cg["pts_h"] / n_pts - 0.05),
                              warnings, down)
                    total_chars += len(pt)
            chart_pages += 1
        elif isinstance(s, HierarchySlide):
            g = hierarchy_geom(s)
            for j, node in enumerate(s.nodes):
                _cap_note(i, s, f"nodes[{j}].desc", f"子模块「{node.label}」描述",
                          len(node.desc), g["desc_w"], g["desc_h"], warnings, down,
                          steps=layouts.HIER_STEPS)
                total_chars += len(node.desc)
            # 包含型层级图：卡片高度贴着文字走，留白不能靠拉长框解决
            if s.mode == "contain" and g["slack"] > 0.80:
                warnings.append(_item(
                    i, "hierarchy", "nodes",
                    f"子模块卡片只有 {g['ch']:.1f} in 高，下方还空 {g['slack']:.1f} in："
                    f"卡片高度贴着文字走（拉长就是空框）。请把每个 node.desc 写满"
                    f"（补机制、代价、度量口径），或补一个占位子模块"
                    f"（label 写「待补：XX」、desc 写明缺什么、找谁要）；"
                    f"内容本就不多就并到上/下一页"))
            chart_pages += 1
        elif isinstance(s, TableSlide):
            # 单元格过长按**列宽容量**判断；字号越小每行装字越多，
            # 所以按最小字号档（10pt）算容量 —— 渲染器会自己降档去塞。
            _w = col_widths(s.headers, s.rows, 12.23)
            for r, row in enumerate(s.rows):
                for c, val in enumerate(row):
                    cap = max(14, layouts.cell_cap_chars(_w[c], 0.30,
                                                         layouts.CAP_MIN_SIZE))
                    # 行高能容 2 行，只有超过 2 倍容量（会挤成 3 行以上）才算问题
                    if len(val) > 2 * cap:
                        warnings.append(_item(
                            i, "table", f"rows[{r}][{c}]",
                            f"单元格 {len(val)} 字，是本列容量（约 {cap} 字/行）的 2 倍以上，"
                            f"会挤成 3 行把行高顶开：删减到 {2 * cap} 字以内，"
                            f"或把这列内容挪进 `side`/`note`"))
            _check_table_columns(i, s, warnings)
            # 表格行高是紧凑的，行少就会在下方留白。留白**不靠拉高行高**解决，
            # 而是把下一章节/关联章节的明细并进来；原文没有就补一行显式占位。
            _side = getattr(s, "side", None) or []
            # 行高必须容得下单元格文字：python-pptx 里的行高只是"最小值"，
            # PowerPoint 打开时会按文字把行撑高，撑开的表就跑到页外/压住 note。
            _tg = layouts.table_geom(len(s.rows), s.summary, bool(s.note),
                                     s.headers, s.rows, bool(_side))
            if _tg["cell_overflow"] > 0.02:
                warnings.append(_item(
                    i, "table", "rows",
                    f"表格 {len(s.rows)} 行压紧后行高只有 {_tg['body_h']:.2f} in，"
                    f"单元格文字在最小字号档（{layouts.TABLE_CELL_STEPS[-1][0]:g}pt）下仍高出 "
                    f"{_tg['cell_overflow']:.2f} in —— PowerPoint 打开时会把行撑开，"
                    f"整张表会溢出内容区并压住底部 `note`：删行/并页，"
                    f"或把长单元格挪进 `side`、`note`"))
            slack, extra, ratio = table_slack(len(s.rows), s.summary, bool(s.note),
                                              s.headers, s.rows, bool(_side))
            if slack > 0.80:
                if _side:
                    warnings.append(_item(
                        i, "table", "rows",
                        f"表格 {len(s.rows)} 行、只占内容区 {ratio:.0%}，主表下方还空 "
                        f"{slack:.1f} in（约能再放 {extra} 行）：把下一章节/关联章节的明细"
                        f"并进来（补 {extra} 行左右）；若相邻页也是这种小表，"
                        f"**两页并成一页、左右并排两个模块**（panel_2col）比两页各留半页"
                        f"空白强得多。原文没有的就补一行显式占位。不要靠拉高行高硬撑。"))
                else:
                    warnings.append(_item(
                        i, "table", "rows",
                        f"表格 {len(s.rows)} 行、只占内容区 {ratio:.0%}，下方还空 {slack:.1f} in"
                        f"（约能再放 {extra} 行）：**整页只有一个信息块**是密度上不去的"
                        f"主因 —— 优先加 `side`（0-4 条关键数字/边界/待确认点），"
                        f"表格会自动收窄、右侧多一块；再把下一章节/关联章节的明细并进来"
                        f"（约 {extra} 行）；相邻页也是小表就**并成一页左右并排**"
                        f"（panel_2col）。原文没有的就补一行「待补」显式占位。"
                        f"不要靠拉高行高硬撑。"))
            chart_pages += 1
        elif isinstance(s, ChartSlide):
            total_chars += len(s.summary) + sum(len(ins.text) for ins in s.insights)
            chart_pages += 1
            # 图表紧凑是按数据点算的：数据点少，下半页就会空 —— 补明细表或并页
            _cb = chart_bottom(s.summary, s.chart.type, len(s.chart.categories),
                               len(s.detail_rows or []), bool(s.footer_note))
            if layouts.BODY_BOTTOM - _cb > 0.80:
                warnings.append(_item(
                    i, "chart", "detail_rows",
                    f"图表页下半页还空 {layouts.BODY_BOTTOM - _cb:.1f} in：数据点只有 "
                    f"{len(s.chart.categories)} 个 —— 用 `detail_headers` + `detail_rows` "
                    f"配一张明细表（指标/目标值/口径/验证方式），或和其它图并成一页"
                    f"（panel_2col 左右并排），别把半页留白交出去"))
        elif isinstance(s, Panel2ColSlide):
            for p in s.panels:
                total_chars += len(p.summary) + sum(len(x) for x in p.points)
            chart_pages += 1
            _check_panels(i, s, warnings)
            _slack, _ph = panel_slack(s)
            # 模块内容超框：渲染器已把要点压到可读下限以下、或把迷你表行高压到
            # 保底线以下 —— 这正是用户复现的「P3 右侧框里的表格被挤出」的成因，
            # 只是现在不会真的画出框了，而是变成"字越来越小 + qc 点名"。
            _over, _ptsz, _rowh, _ = layouts.panel_pressure(s)
            # 行高判据改成**相对自然行高**：旧判据是绝对下限 0.16+0.02，
            # 而自然行高 0.32 —— 被压到 0.183 照样"合格"，可那已经是行高腰斩、
            # 表格挤成一条线的观感了。低于自然行高的 3/4 就该删行。
            _row_squeezed = _rowh < layouts.PANEL_TBL_ROW * 0.75
            if _over > 0.05 and (_ptsz < layouts.PANEL_PT_FLOOR or _row_squeezed):
                warnings.append(_item(
                    i, "panel_2col", "panels",
                    f"模块内容超出模块框约 {_over:.1f} in（模块高 {_ph:.1f} in）："
                    f"渲染器只能把要点压到 {_ptsz:g}pt（可读下限 "
                    f"{layouts.PANEL_PT_FLOOR:g}pt）、迷你表行高压到 {_rowh:.2f} in 才塞进去。"
                    f"删掉一条要点或几行明细；明细超过 8 行 / 要点超过 6 条就该"
                    f"**另起一页**用 `table` 或卡片页展开，不要挤在一个模块里"))
            elif _slack > 0.80:
                warnings.append(_item(
                    i, "panel_2col", "panels",
                    f"两个模块里还剩 {_slack:.1f} in 空白（模块高 {_ph:.1f} in）："
                    f"把相邻页/关联章节的明细并进来（加要点或加迷你表行），"
                    f"或补一条显式占位；不要靠拉高行高硬撑"))
        elif isinstance(s, FigureSlide):
            total_chars += len(s.summary or "") + len(s.caption or "") \
                + sum(len(x) for x in s.points)
            chart_pages += 1
            _check_figure(i, s, errors, warnings)
        elif isinstance(s, EndSlide):
            # 结尾页以前完全在 qc 视野之外：4 段各 200 字会把第 4 张卡推到 8.75in
            # （页高只有 7.5in）。现在几何统一由 end_geom 算，放不下就点名。
            _eg = layouts.end_geom(list(s.sections or []), bool(s.contact))
            if not _eg["ok"]:
                _last = layouts.END_TEXT_STEPS[-1][0]
                warnings.append(_item(
                    i, "end", "sections",
                    f"{len(s.sections or [])} 段结尾正文降到 {_last:g}pt 仍需要 "
                    f"{_eg['total']:.1f} in，可用只有 {_eg['avail']:.1f} in —— "
                    f"删并段落（结尾页 2~3 段足够）或每段压到 80 字以内；"
                    f"硬塞会被裁到画布外"))
            for j, sec in enumerate(s.sections or []):
                _cap_note(i, s, f"sections[{j}].text", f"结尾段「{sec.heading}」",
                          len(sec.text), _eg["card_w"] - 0.62,
                          max(0.3, _eg["heights"][j] - 0.14), warnings, down,
                          steps=layouts.END_TEXT_STEPS)
        # 密度下限 + 结论段必备：内容页标题下必须有 summary（核心结论段）
        summary = getattr(s, "summary", None)
        content_types = ("statement", "compare_3col", "numbered_cards",
                         "layered_stack", "flow", "hierarchy", "table", "chart",
                         "panel_2col", "figure", "timeline", "funnel", "matrix",
                         "key_groups", "grid_2x2")
        if s.slide_type in content_types:
            if not summary:
                warnings.append(_item(i, s.slide_type, "summary",
                                      "标题下缺少「核心结论段」（summary），建议补一段 "
                                      "40~80 字 / 2~3 句的结论"))
            body = getattr(s, "body", None) or summary
            if body and len(body) < 30:
                errors.append(_item(i, s.slide_type, "summary/body",
                                    "导语/正文不足 30 字，信息密度不达标"))
            # 加粗要精炼：summary 里的重点只标关键词
            _check_emphasis(i, s.slide_type, "summary", summary or "", warnings)
            # 结论段硬指标：40~80 字 / 2~3 句 / 禁坐标词 / 禁填充语 / 不复述标题
            _check_summary(i, s.slide_type, summary, getattr(s, "title", None), warnings)
            # 一页至少两个信息块（整页只有一个容器，密度上不去）
            _check_page_blocks(i, s, warnings)
            # 正文密度下限：不许把正文精简过头（卡片描述、栏描述、层描述…）
            _check_density_floor(i, s, warnings)
        if isinstance(s, StatementSlide):
            # statement 的正文不再自动加粗，重点必须自己用 **...** 标出来
            _check_emphasis(i, "statement", "body", s.body, warnings, require_span=True)

    _brief_checks(deck, errors, warnings, source_path)
    # 输入覆盖度：唯一能反向抓"整块内容没进 deck"的检查（--source 时启用）
    coverage = _check_source_coverage(deck, source_path, errors, warnings)

    stats = {
        "brief": getattr(deck, "brief", None),
        "slides": len(deck.slides),
        "content_pages": chart_pages,
        "total_content_chars": total_chars,
        "avg_chars_per_content_page": round(total_chars / chart_pages, 1) if chart_pages else 0,
        # 落到比默认档更小字号的页面（渲染正常，列出来便于人工抽检可读性）
        "downshifted": down,
        "downshifted_pages": len({d["slide"] for d in down}),
    }
    if coverage is not None:
        stats["source_coverage"] = coverage
    # ---- 回炉成本控制：把 warning 归类计数 + 给一句优先提示 ----
    # 实测一份 30 页研讨 deck 能报出 40+ 条 warning，其中 table.rows / table.summary
    # 各 12 条是**同一改法**。逐条列出来会诱导模型逐页来回改（每轮都是一次完整生成），
    # 这里给出"共 N 条、集中在哪几类"，让它一轮把同类改完。
    _g: dict[str, int] = {}
    for w in warnings:
        key = f"{w['type']}.{w['field']}"
        _g[key] = _g.get(key, 0) + 1
    _g = dict(sorted(_g.items(), key=lambda kv: -kv[1]))
    stats["warning_groups"] = _g
    if len(warnings) >= 8:
        top = "、".join(f"{k}×{v}" for k, v in list(_g.items())[:4])
        stats["digest"] = (
            f"共 {len(warnings)} 条 warning，集中在：{top} —— 同类是同一种改法，"
            f"**一轮把这几个类型一起改完**再重跑，不要一页一页来回")
    if chart_pages and stats["avg_chars_per_content_page"] < 120:
        warnings.append(_item("-", "deck", "density",
                              f"内容页平均 {stats['avg_chars_per_content_page']} 字/页，"
                              f"低于高密度标准（120+），建议充实内容"))

    # ---- 页面充实度检查：内容太少要警告（每页都要铺满）----
    # statement 的正文下限在 _check_density_floor 里逐页点名（body < 150 字），
    # 这里只补一条"正文 + 论据总量"的兜底，避免正文写满、论据却空着。
    for i, s in enumerate(deck.slides, 1):
        if isinstance(s, StatementSlide):
            total_content = len(s.body) + sum(len(p.text) for p in (s.points or []))
            if total_content < 200:
                warnings.append(_item(i, "statement", "density",
                                      f"本页内容偏少（正文+论据共 {total_content} 字）："
                                      f"补论据/数字，或把相邻页同主题的内容并进来；"
                                      f"内容确实不多就改用 panel_2col 与相邻页并成一页"))
    # 报告瘦身：同一规则的重复说明只讲一次（报告整份进上下文，见 _fold_rule_tails）
    stats["report_chars_saved"] = (_fold_rule_tails(warnings)
                                   + _fold_rule_tails(errors))
    # ---- 速度档：只改报告详略，不改检查项（见 _slim_warnings）----
    stats["speed"] = speed if speed in SPEEDS else "slow"
    if stats["speed"] == "fast":
        stats["warnings_before_slim"] = len(warnings)
        _slim_warnings(warnings)
    stats["warnings_pending"] = len(warnings)
    return {"errors": errors, "warnings": warnings, "stats": stats}
