# -*- coding: utf-8 -*-
"""PPT-DSL：结构化中间表示（pydantic 模型）。

LLM 每次生成一页（或整份 deck）的 JSON，必须通过本模块校验后才进入渲染。
字段长度约束 = 版式容量约束，校验错误信息会直接反馈给 LLM 回炉重生成。
"""
from __future__ import annotations

import json
import sys
from typing import Annotated, Literal, Optional, Union

from pydantic import BaseModel, Field, ValidationError, model_validator


# ---------- 通用字段 ----------
Str = str

# 底部页脚条的字数上限：页脚条压扁成 0.36in（单行）且右端避开右下角单位水印后，
# 一行只放得下这么多字（渲染器还会从 11pt 降到 9pt 来兜）。写不下就并进 summary。
FOOTER_NOTE_MAX = 64


# ---------- 动态容量约束：几何是唯一口径 ----------
#
# 静态 `max_length` 只能按**最宽松**情形给：同样是 `desc`，4 张卡时每卡 4.1in 宽
# （能写 240 字），8 张卡时每卡 2.0in 宽（只能写 90 字）—— 差 2.7 倍。
# 于是"字段上限"和"实际能落地"长期脱节：写满上限的内容，渲染器降到最小字号仍
# 溢出画布，只能靠 --qc 事后点名。
#
# 这里改成**按条目数动态校验**：直接调用渲染器自己的几何函数（numbered_geom /
# timeline_geom / …）反算这一页每格能写多少字，超了就在 schema 阶段报错。
# 报告里的条数、列数、可写字数全部来自同一套几何，不会再出现"文档说 240、
# 实际 90"的口径漂移。

# 动态容量的容差：字宽是保守估算，差几个字不该拒稿。
# 只吸收估算误差 —— 真正的"翻倍超限"照样会被拦下。
CAP_TOL = 1.10        # 允许超出几何容量 10%
CAP_SLACK = 8         # 再额外给 8 字余量（短文本的百分比会被放大，需要兜底）


def _layouts():
    """延迟导入渲染器（schema 层不该在 import 时就拽进 python-pptx）。"""
    try:
        import layouts
        return layouts
    except Exception:              # 纯 schema 环境（无 pptx）：跳过几何校验
        return None


def _cap_hint(items, attr: str, width: float, height: float,
              label: str, hint: str = "", steps=None) -> None:
    """逐条核对容量：超出就把"第几条 / 写了多少 / 最多能写多少"一次报全。

    **带容差**（`CAP_TOL` / `CAP_SLACK`）：字宽是保守估算，和真实排版差几个字太正常，
    差 4 个字就把一份合法 deck 拒掉是负收益。容差只吃掉估算误差 ——
    "写满 240 但只能装 90"这类严重超限照样拦得住。
    """
    L = _layouts()
    if L is None or width <= 0 or height <= 0:
        return
    cap = L.cap_chars(width, height, steps)
    if cap <= 0:
        return
    limit = int(cap * CAP_TOL) + CAP_SLACK
    bad = [(j + 1, len(getattr(it, attr) or ""))
           for j, it in enumerate(items)
           if len(getattr(it, attr) or "") > limit]
    if not bad:
        return
    detail = "、".join(f"第 {j} 条 {n} 字" for j, n in bad[:5])
    more = "" if len(bad) <= 5 else f" 等 {len(bad)} 条"
    raise ValueError(
        f"{label} 超出本页容量：这套几何下每条最多约 {cap} 字，现在 {detail}{more}。"
        + (hint or "条目数越多每格越窄 —— 要么压字数，要么把这一页拆成两页。"))



class CoverSlide(BaseModel):
    slide_type: Literal["cover"]
    title: Str = Field(..., min_length=2, max_length=24, description="主标题（中文）")
    en_title: Optional[Str] = Field(None, max_length=60, description="英文副题")
    subtitle: Optional[Str] = Field(None, max_length=60)
    org: Optional[Str] = Field(None, max_length=60, description="落款/单位")


class TocItem(BaseModel):
    no: Str = Field(..., max_length=2, description="序号，如 01")
    text: Str = Field(..., min_length=2, max_length=20)
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标，可用 Tabler 图标名或中文语义名（如 齿轮/数据/安全）")


class TocSlide(BaseModel):
    slide_type: Literal["toc"]
    items: list[TocItem] = Field(..., min_length=3, max_length=6,
                                 description="3-6 章；章节更多就**再写一页目录**"
                                             "（或按上半场/下半场分），不要合并或删章节")


class SectionSlide(BaseModel):
    slide_type: Literal["section"]
    no: Str = Field(..., max_length=2)
    title: Str = Field(..., min_length=2, max_length=20)
    subtitle: Optional[Str] = Field(None, max_length=60)
    icon: Optional[Str] = Field(None, max_length=32, description="章节页大图标")


class HighlightBox(BaseModel):
    label: Str = Field(..., max_length=12)
    value: Str = Field(..., max_length=30, description="关键数据/结论，如 52.8% → 66.5%"
                                  "（视觉层级：使用 20pt hero_value，比正文大 ≥2 级）")
    desc: Optional[Str] = Field(None, max_length=80)
    icon: Optional[Str] = Field(None, max_length=32, description="高亮数据卡图标")


class SupportPoint(BaseModel):
    label: Str = Field(..., min_length=2, max_length=8)
    text: Str = Field(..., min_length=6, max_length=44)
    icon: Optional[Str] = Field(None, max_length=32, description="论据左侧图标")


class StatementSlide(BaseModel):
    """结论页：标题必须是完整判断句，正文高密度展开。

    兜底定位：**这是最后的落点** —— 内容能分组就走 key_groups，能对比就走
    compare/panel，有序就走 flow/timeline；只有"确实是一段连贯论述"才到这里。
    """
    slide_type: Literal["statement"]
    title: Str = Field(..., min_length=8, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="标题下的总结陈述段")
    body: Str = Field(..., min_length=80, max_length=700,
                      description="正文，150-700 字（**上限按几何容量给足**：8.46in 宽 × "
                                  "4.77in 高的正文区，降到 10pt 能装 1000+ 字，"
                                  "700 是留出可读余量的值）。高亮卡与支撑论据在右栏。"
                                  "**用换行指定卡片边界（最多 3 张卡）**：每张卡的"
                                  "**首句（≤16 字）会自动提为卡头标题行**（加粗染色的组标题，"
                                  "扫读的骨架）；首句更长的卡仍把首句留在正文流里加粗染色。"
                                  "卡内长文标 2 处以上 `**关键短语**：正文` —— 短语与正文"
                                  "**同字号**、只加粗 + 强调色，冒号后正文接着排"
                                  "（不放大字号：放大 + 另起行会把文字顶出卡片背景）")
    highlight: Optional[HighlightBox] = None
    points: list[SupportPoint] = Field(default_factory=list, min_length=0, max_length=4,
                                       description="支撑论据，0-4 条，展示在右栏")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is None:
            return self
        _bw, _bh = L.statement_body_box(self)
        _cap_hint([self], "body", _bw, _bh,
                  "statement 的 body（正文卡）",
                  hint="正文从 12.5pt 逐档降到 10pt 仍放不下：把机制细节拆到相邻页，"
                       "或用 `**短语**：正文` 分小节（条目式比整段更省高度）。",
                  steps=L.STATEMENT_STEPS)
        if self.highlight and self.highlight.desc:
            g = L.statement_geom(self)
            rail_h = g["bottom"] - g["top"]
            lim = max(0.90, rail_h * 0.45)
            if g["hl_h"] > lim:
                _cap = L.cap_chars(g["rw"] - 0.66, lim - L.HL_DESC_TOP - 0.06,
                                   ((L.HL_DESC_SIZE, L.HL_DESC_SPACING),))
                raise ValueError(
                    f"statement 的 highlight.desc 太长：高亮卡会随说明长高到 "
                    f"{g['hl_h']:.1f} in，而右栏一共只有 {rail_h:.1f} in —— "
                    f"再长就把右栏的支撑论据挤掉了。"
                    f"压到约 {_cap} 字以内（一两句话说清口径或对比基线），"
                    f"细节挪到 body 或 points。")
        return self


class Callout(BaseModel):
    """卡片型版式底部的横排辅区：2-4 条关键数字 / 边界条件 / 待确认点。

    卡片讲完"是什么"，callout 补"数字是多少、边界在哪、还缺什么" ——
    一页两个信息块，密度才上得去。
    """
    label: Str = Field(..., min_length=2, max_length=12)
    value: Optional[Str] = Field(None, max_length=18, description="关键数字/结论，大字")
    text: Optional[Str] = Field(None, min_length=4, max_length=60, description="说明，小字")
    icon: Optional[Str] = Field(None, max_length=32,
                                description="本条图标；辅区落在右栏时画在卡片左上")


class CompareColumn(BaseModel):
    label: Str = Field(..., min_length=1, max_length=10)
    tag: Str = Field(..., min_length=2, max_length=14, description="核心机制短语")
    desc: Str = Field(..., min_length=20, max_length=220,
                      description="每栏描述；3 栏时每栏可写约 200 字，4 栏时约 120 字"
                                  "（渲染器自动降档到 10pt，qc 报实际容量）")
    points: list[Str] = Field(default_factory=list, min_length=0, max_length=4,
                              description="特征要点，0-4 条，每条 ≤14 字")
    metaphor: Str = Field(..., min_length=4, max_length=30, description="生活化比喻")
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标名，见 icons.py：bolt/gear/cloud/cube/hexagon/"
                                            "donut/sun/drop/decision/document/process/database/"
                                            "refresh/star/smiley/chat/chart/arrow_up/moon/pie/can/pentagon")


class CompareSlide(BaseModel):
    """三栏对比：三栏并排已经是"三块信息"，**不要再加底部辅区**。

    实测（`layouts.compare_desc_box`）：一条 `callout` 会吃掉 0.96in 高度，
    把每栏描述容量从约 100 字砍到约 20 字 —— 这正是过去三栏页正文只剩一句话的原因。
    所以要收尾就写一句 `footer_note`，要放数字就写进各栏 `points`。
    """
    slide_type: Literal["compare_3col"]
    title: Str = Field(..., min_length=6, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240)
    columns: list[CompareColumn] = Field(..., min_length=3, max_length=4,
                                         description="**3-4 栏**（4 栏时每栏自动收窄、字号降档）；"
                                                     "对比项就是文档里的项，不许删；"
                                                     "超过 4 项才另起一页继续对比")
    # 字段保留只为给出**明确报错**（pydantic 默认会静默忽略未知字段，那就变成丢内容）
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="三栏对比不支持：会砍掉每栏描述容量，"
                                               "请改用 footer_note")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def no_callout(self):
        if self.callout:
            raise ValueError(
                "compare_3col 不支持 `callout`：它会在底部吃掉一整条高度，"
                "把每栏描述容量从约 100 字砍到约 20 字（三栏并排本身已是三块信息）。"
                "请把结论写进 `footer_note`（≤64 字），数字写进各栏的 `points`。")
        return self

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is None:
            return self
        _dw = _dh = 0.0
        for col in self.columns:
            dw, dh, _ = L.compare_desc_box(self.summary, len(col.points), False,
                                           0.0, len(self.columns))
            _dw, _dh = max(_dw, dw), max(_dh, dh)
        _cap_hint(self.columns, "desc", _dw, _dh,
                  f"compare_3col 的 columns[].desc（{len(self.columns)} 栏并排，"
                  f"每栏描述 {_dw:.2f}×{_dh:.2f}in）",
                  hint="栏越宽字数越多；三栏时每栏描述写 60~100 字最稳，"
                       "四栏会明显变窄。要点条数也会吃掉描述高度 —— 减要点或压描述。")
        return self


class NumberedCard(BaseModel):
    no: Str = Field(..., max_length=2)
    name: Str = Field(..., min_length=2, max_length=20,
                      description="卡片的**加粗重点标题**：写成判断句/带数字的结论，"
                                  "别写成「任务集管理」这种名词短语")
    sub: Optional[Str] = Field(None, max_length=32,
                               description="一句话定位（可选：标题已是判断句时可以不写）")
    desc: Str = Field(..., min_length=20, max_length=240,
                      description="卡片主内容。几何容量随张数变化：4 张卡（2 列）每卡"
                                  "可写到约 240 字，5-6 张卡（3 列）约 85 字，"
                                  "7-8 张卡（4 列）约 90 字 —— 以 --qc 报的「约 X 字」"
                                  "为准（渲染器会自动降档到 10pt）")
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标：Tabler 图标名或中文语义名（齿轮/数据/安全/循环…），见 iconlib.ALIASES")


class NumberedCardsSlide(BaseModel):
    slide_type: Literal["numbered_cards"]
    title: Str = Field(..., min_length=6, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240)
    cards: list[NumberedCard] = Field(..., min_length=4, max_length=8,
                                      description="**4-8 张卡**：4 张排 2 列（每卡最宽），"
                                                  "5-6 张排 3 列，7-8 张排 4 列（窄卡不再画"
                                                  "块头图标，把宽度还给标题与描述）。"
                                                  "卡数就是文档里的项数，不许删；"
                                                  "确实超过 8 张才另起一页继续")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX,
                                       description="底部一句话结论（与 callout 二选一，"
                                                   "卡片描述需要空间时用它顶替 callout）")

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.numbered_geom(self)
            _cap_hint(self.cards, "desc", g["desc_w"], g["desc_h"],
                      f"numbered_cards 的 cards[].desc（{len(self.cards)} 张卡排 "
                      f"{g['cols']} 列，每卡 {g['cw']:.2f}in 宽）",
                      hint=f"{len(self.cards)} 张卡时每卡只有 {g['cw']:.2f}in 宽，"
                           f"字数写满会压出卡片下沿。要么压到上述字数内，"
                           f"要么减到 4 张卡（2 列，每卡 4.1in 宽、能写 240 字），"
                           f"要么把这一页拆成两页。")
        return self


class StackLayer(BaseModel):
    label: Str = Field(..., min_length=2, max_length=14)
    desc: Str = Field(..., min_length=15, max_length=180,
                      description="层说明；层数少时每层可写约 170 字，12 层时约 50 字"
                                  "（条形随层数压紧，字数以 --qc 报的为准）")
    tags: list[Str] = Field(default_factory=list, min_length=0, max_length=3,
                            description="代表技术/关键词，0-3 个，每个 ≤8 字")
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标：Tabler 图标名或中文语义名（齿轮/数据/安全/循环…），见 iconlib.ALIASES")


class LayeredStackSlide(BaseModel):
    """层级堆栈：layers[0] 为最底层（被依赖），依次向上。"""
    slide_type: Literal["layered_stack"]
    title: Str = Field(..., min_length=6, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240)
    layers: list[StackLayer] = Field(..., min_length=3, max_length=12,
                                     description="**3-12 层**：条形高度与间距按层数自适应"
                                                 "（层多就压紧），层数就是文档里的层数，"
                                                 "不许为了凑容量删层；确实超过 12 层才另起一页")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.layered_geom(self)
            _cap_hint(self.layers, "desc", g["desc_w"], g["desc_h"],
                      f"layered_stack 的 layers[].desc（{len(self.layers)} 层，"
                      f"每条只有 {g['bar_h']:.2f}in 高）")
        return self


class FlowStep(BaseModel):
    label: Str = Field(..., min_length=2, max_length=10)
    desc: Str = Field(..., min_length=10, max_length=80)
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标：Tabler 图标名或中文语义名（齿轮/数据/安全/循环…），见 iconlib.ALIASES")


class FlowSlide(BaseModel):
    """流程/循环页。

    **步数 3~12 都能一页放下**：超过 6 步渲染器自动折成两行（第一行左→右，
    折返后第二行继续），带 loop_note 的 ≤6 步画成环形、>6 步写成折返通道上的注释。
    所以**不要因为"步数多"就删步骤或拆页** —— 一页装得下就一页装下，
    研讨型的硬要求是输入里的步骤一条都不能少。
    """
    slide_type: Literal["flow"]
    title: Str = Field(..., min_length=6, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240)
    steps: list[FlowStep] = Field(..., min_length=3, max_length=12,
                                  description="3-12 步；>6 步自动排成两行折返，别删步骤")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    loop_note: Optional[Str] = Field(None, max_length=60, description="循环/回退说明")

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is None:
            return self
        n = len(self.steps)
        if self.loop_note and 3 <= n <= L.FLOW_WRAP_AT:
            _w, _h, _steps = L.flow_cycle_desc_box()
            _cap_hint(self.steps, "desc", _w, _h,
                      f"flow 的 steps[].desc（{n} 步画成环形，节点固定 "
                      f"{L.FLOW_NODE_W:.2f}in 宽、描述区 {_w:.2f}×{_h:.2f}in）",
                      hint="环形节点是固定尺寸的小卡，描述只能一两句话。"
                           "要写长说明就删掉 loop_note（改走两行线性流程），"
                           "那样每条能写 49 字以上。",
                      steps=_steps)
        else:
            g = L.flow_linear_geom(self)
            _cap_hint(self.steps, "desc", g["desc_w"], g["desc_h"],
                      f"flow 的 steps[].desc（{n} 步排成"
                      f"{'两行折返' if n > L.FLOW_WRAP_AT else '一行'}，"
                      f"描述区 {g['desc_w']:.2f}×{g['desc_h']:.2f}in）")
        return self


class HierarchyNode(BaseModel):
    label: Str = Field(..., min_length=2, max_length=12)
    desc: Str = Field(..., min_length=6, max_length=100,
                      description="子模块说明；几何容量约 85 字"
                                  "（卡片会吃掉一部分留白，写 35~85 字最合适）")
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标：Tabler 图标名或中文语义名，见 iconlib.ALIASES")


class HierarchySlide(BaseModel):
    """层级/关系图：根节点 + 子模块，表达「包含」（树形）或「递进」（阶梯）。"""
    slide_type: Literal["hierarchy"]
    title: Str = Field(..., min_length=6, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240)
    root: Str = Field(..., min_length=2, max_length=20, description="根节点/上层整体")
    nodes: list[HierarchyNode] = Field(..., min_length=2, max_length=8,
                                       description="子模块/下层，**2-8 个**（卡片会随个数变窄，"
                                                   "描述随之变短）；子模块就是文档里的子模块，"
                                                   "不许删；超过 8 个才按子域拆两页")
    mode: Literal["contain", "progressive"] = Field("contain",
                                                    description="contain=包含树形，progressive=递进阶梯")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.hierarchy_geom(self)
            _cap_hint(self.nodes, "desc", g["desc_w"], g["desc_h"],
                      f"hierarchy 的 nodes[].desc（{len(self.nodes)} 个子模块，"
                      f"卡片 {g['desc_w']:.2f}×{g['desc_h']:.2f}in）",
                      hint="子模块卡很窄：desc 只写「负责什么」，"
                           "横向/纵向关系交给 mode 与 callout。",
                      steps=L.HIER_STEPS)
        return self


class SideNote(BaseModel):
    """右侧辅区的一条：关键数字 / 边界条件 / 待确认点。

    主内容（表格、图表）只装得下一种形态，页面剩下的 1/3 靠它承载 ——
    「每页信息量尽可能多」不是把表拉宽，而是让一页有两个信息块。
    """
    label: Str = Field(..., min_length=2, max_length=12)
    value: Optional[Str] = Field(None, max_length=18, description="关键数字/结论，大字")
    text: Optional[Str] = Field(None, min_length=4, max_length=70, description="说明，小字")
    icon: Optional[Str] = Field(None, max_length=32)


class TableSlide(BaseModel):
    slide_type: Literal["table"]
    title: Str = Field(..., min_length=4, max_length=32)
    summary: Optional[Str] = Field(None, max_length=240, description="标题下的总结陈述段")
    headers: list[Str] = Field(..., min_length=2, max_length=6)
    rows: list[list[Str]] = Field(..., min_length=1, max_length=12,
                                  description="1-12 行（行高会按行数自动压紧，"
                                              "不按页面高度摊开）；行就是数据行，不许删，"
                                              "超过 12 行才另起一页")
    side: list[SideNote] = Field(default_factory=list, min_length=0, max_length=4,
                                 description="右侧辅区 0-4 条：关键数字/边界/待确认点。"
                                             "填了它表格会自动收窄，把右侧让给它；"
                                             "整页只有一个信息块正是密度上不去的主因")
    note: Optional[Str] = Field(None, max_length=120)
    hero_col: Optional[Str] = Field(
        None, max_length=12,
        description="要点亮的主角列：填**表头文字**（与 headers 里某一项完全一致），"
                    "该列整列会加底色强调。不填就不高亮 —— 别指望渲染器猜")

    @model_validator(mode="after")
    def validate_row_width(self):
        """行宽必须与表头一致。

        这个函数以前**忘了加 `@model_validator`**，所以只是普通方法、从来没跑过：
        列数写错的行会一路渲染，单元格静默错位（最后一格被丢掉）。
        """
        for i, row in enumerate(self.rows):
            if len(row) != len(self.headers):
                raise ValueError(
                    f"第 {i+1} 行有 {len(row)} 列，表头是 {len(self.headers)} 列，必须一致")
        if self.hero_col and self.hero_col.strip() not in [h.strip() for h in self.headers]:
            raise ValueError(
                f"hero_col「{self.hero_col}」不在表头里：可选 "
                + "、".join(f"「{h}」" for h in self.headers))
        return self


class ChartSeries(BaseModel):
    name: Str = Field(..., min_length=1, max_length=20)
    values: list[float] = Field(..., min_length=2, max_length=10)


class ChartSpec(BaseModel):
    type: Literal["column", "bar", "line", "pie", "doughnut", "radar"]
    categories: list[Str] = Field(..., min_length=2, max_length=10)
    series: list[ChartSeries] = Field(..., min_length=1, max_length=3)
    unit: Optional[Str] = Field(None, max_length=6, description="数值单位，如 % / 次")

    @model_validator(mode="after")
    def check_shape(self):
        for s in self.series:
            if len(s.values) != len(self.categories):
                raise ValueError(
                    f"系列「{s.name}」有 {len(s.values)} 个值，但 categories 有 "
                    f"{len(self.categories)} 个，必须一一对应")
        if self.type in ("pie", "doughnut") and len(self.series) != 1:
            raise ValueError("饼图/环形图只允许 1 个系列")
        return self


class InsightItem(BaseModel):
    text: Str = Field(..., min_length=8, max_length=70)
    icon: Optional[Str] = Field(None, max_length=32,
                                description="图标：Tabler 图标名或中文语义名（齿轮/数据/安全/循环…），见 iconlib.ALIASES")


class ChartSlide(BaseModel):
    """图表页：左侧原生图表 + 右侧关键洞察（+ 可选的下方明细表）。"""
    slide_type: Literal["chart"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Str = Field(..., min_length=15, max_length=240)
    chart: ChartSpec
    insights: list[InsightItem] = Field(..., min_length=2, max_length=4)
    detail_headers: Optional[list[Str]] = Field(None, max_length=4,
                                                description="下方明细表表头（2-4 列）")
    detail_rows: list[list[Str]] = Field(default_factory=list, max_length=10,
                                         description="下方明细表数据，列数须与 detail_headers 一致；"
                                                     "数据点少、图很矮时用它把下半页填满"
                                                     "（1-10 行，行高自动压紧）")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)


class PanelModule(BaseModel):
    """双模块页里的一个模块：**自带背景框**，顶部是这一模块的**总结陈述段**。

    一页只放两块内容时用它 —— 两块左右并排（不是上下堆叠）。每块可以装
    「要点列表」或「迷你明细表」（二选一，也可以都有：先要点、后明细）。
    两块内容原本是两页（各自内容太少）时，优先并成一页用本版式。
    """
    label: Str = Field(..., min_length=2, max_length=14, description="模块标题")
    summary: Str = Field(..., min_length=16, max_length=110,
                         description="模块顶部的总结陈述段：1-3 句，开口即结论，带数字/口径")
    icon: Optional[Str] = Field(None, max_length=32, description="本模块唯一图标，画在块头")
    points: list[Str] = Field(default_factory=list, max_length=6,
                              description="模块内要点，0-6 条，每条 ≤44 字")
    headers: list[Str] = Field(default_factory=list, max_length=4,
                               description="模块内迷你明细表表头，0-4 列（与 rows 配套）")
    rows: list[list[Str]] = Field(default_factory=list, max_length=6,
                                  description="模块内迷你明细表数据，0-6 行")

    @model_validator(mode="after")
    def check_content(self):
        if not self.points and not self.rows:
            raise ValueError(f"模块「{self.label}」既没有 points 也没有 rows，框里会空着")
        if self.rows:
            if not self.headers:
                raise ValueError(f"模块「{self.label}」有 rows 但没有 headers")
            for i, row in enumerate(self.rows):
                if len(row) != len(self.headers):
                    raise ValueError(
                        f"模块「{self.label}」rows 第 {i + 1} 行有 {len(row)} 列，"
                        f"headers 是 {len(self.headers)} 列，必须一致")
        return self


class Panel2ColSlide(BaseModel):
    """双模块并排页：页面横向一分为二，两个模块**各自一个背景框、各自顶部一段总结**。

    这是「一页两块内容 → 左右分置」的落点，也是「上/下两页内容都太少 → 并成一页」
    的落点：把原来两页的同类明细分别塞进左右两个模块，密度立刻上来，
    而且不会出现上下堆叠留下的半页空白。
    """
    slide_type: Literal["panel_2col"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    panels: list[PanelModule] = Field(..., min_length=2, max_length=2,
                                      description="恰好 2 个模块，左、右各一")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)


class FigureSlide(BaseModel):
    """图示页：把输入里的**原图**放到一页上（Word / Visio / draw.io 导出的 EMF 也支持）。

    为什么它非有不可：文档里最贵的往往是那张手绘的架构图/流程图，它没法用文字重述，
    也不该被转成 PNG —— EMF/WMF 是 Windows 的矢量图格式，PowerPoint 原生支持无级缩放，
    转栅格再嵌进去就会糊。本版式的 `image` 走「**逐字节原封不动嵌入**」的通道
    （`layouts.add_image`：EMF/WMF 手工挂 ImagePart，不重绘、不转格式）。

    Word 里的图先用 `python .pi/renderer/docx_in.py <docx> -o .pi/work/input.md --media .pi/work/media`
    抽到 `.pi/work/media/`，再把 `media/NN-xxx.emf` 写进 `image`（路径相对仓库根目录解析）。
    """
    slide_type: Literal["figure"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="标题下的结论段")
    image: Str = Field(..., min_length=3, max_length=260,
                       description="图片路径（相对仓库根目录或绝对路径），"
                                   "支持 png/jpg/gif/bmp/tiff/**emf/wmf**（矢量图原样嵌入）；"
                                   "Word 里的图先用 .pi/renderer/docx_in.py 抽到 .pi/work/media/")
    caption: Optional[Str] = Field(None, max_length=140,
                                   description="图注：一句话说清这张图在证明什么，放在图下方")
    points: list[Str] = Field(default_factory=list, max_length=3,
                              description="图旁要点 0-3 条，每条 ≤44 字（落在右侧纵栏）；"
                                          "要点是图的补充说明，别把图里的标签再抄一遍")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)


class KeyGroup(BaseModel):
    """要点分组的一个组：组标题（判断句）+ 一组要点。

    这是「论点 + 论据成组」的通用结构 —— 组与组**平行**（不是对比、不是流程、
    也不是编号顺序），每组自带 2-6 条要点。
    """
    title: Str = Field(..., min_length=2, max_length=18,
                       description="组标题，写成判断句或带重点的短语，别用纯名词")
    points: list[Str] = Field(..., min_length=2, max_length=6,
                              description="组内要点 2-6 条，每条 4-60 字；"
                                          "写成「结论 + 数字/口径」，不是光秃秃的名词")
    icon: Optional[Str] = Field(None, max_length=32, description="组图标")


class KeyGroupsSlide(BaseModel):
    """要点分组页：**2-4 个平行主题组**，每组 = 组标题 + 要点清单。

    补的是路由缺口：3 条论据以前会被硬塞进 compare_3col（对比语义、还要造比喻），
    「分组但不对比、成组但没顺序」的内容只能摊回 statement 兜底变成字堆。
    本版式就是这类内容的正身 —— statement 仍然保留给「真·连贯论述」做最后兜底。
    """
    slide_type: Literal["key_groups"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    groups: list[KeyGroup] = Field(..., min_length=2, max_length=4,
                                   description="2-4 组；组就是输入里的平行主题，不许删")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is None:
            return self
        g = L.key_groups_geom(self)
        # 卡宽**逐张不同**：3 组时第三组跨整行（宽一倍、每条能写更多字），
        # 所以容量必须按每张卡自己的几何算，不能拿第一张当全页标准。
        for cg in g["cards"]:
            grp = self.groups[cg["i"]]
            n_pts = max(1, len(grp.points))
            cap = L.cap_chars(cg["text_w"], max(0.18, cg["pts_h"] / n_pts - 0.05))
            if cap <= 0:
                continue
            limit = int(cap * CAP_TOL) + CAP_SLACK
            bad = [(k + 1, len(pt)) for k, pt in enumerate(grp.points)
                   if len(pt) > limit]
            if bad:
                detail = "、".join(f"第 {k} 条 {n} 字" for k, n in bad[:5])
                raise ValueError(
                    f"key_groups 的 groups[{cg['i']}]「{grp.title}」要点太长："
                    f"本页 {len(self.groups)} 组时这张卡有 {n_pts} 条要点、"
                    f"每条最多约 {cap} 字，现在 {detail}。"
                    f"要点是清单不是段落 —— 每条压到 {cap} 字内，"
                    f"或把这一组拆到相邻页。")
        return self


class EndSection(BaseModel):
    heading: Str = Field(..., max_length=20)
    text: Str = Field(..., max_length=200)


class EndSlide(BaseModel):
    slide_type: Literal["end"]
    title: Str = Field("谢谢观看", max_length=20)
    sections: list[EndSection] = Field(default_factory=list, max_length=4)
    contact: Optional[Str] = Field(None, max_length=80)

    @model_validator(mode="after")
    def fit_capacity(self):
        """结尾页的纵向预算：段数 × 段长必须落在 2.95→6.80in 之间。

        旧实现按 `_est_lines` 累加，4 段各 200 字会把第 4 张卡推到 8.75in
        （页高只有 7.5in）—— 直接画到画布外。现在几何统一由 end_geom 算。
        """
        L = _layouts()
        if L is None or not self.sections:
            return self
        g = L.end_geom(list(self.sections), bool(self.contact))
        if not g["ok"]:
            _last = L.END_TEXT_STEPS[-1][0]
            _cap = L.cap_chars(g["card_w"] - 0.62,
                               max(0.3, g["avail"] / len(self.sections) - 0.42),
                               L.END_TEXT_STEPS)
            raise ValueError(
                f"结尾页正文太多：{len(self.sections)} 段降到 {_last:g}pt 仍需要 "
                f"{g['total']:.1f} in，可用只有 {g['avail']:.1f} in。"
                f"结尾页 2~3 段就够，每段压到约 {_cap} 字以内（写「下一步」"
                f"「联系方式」这类收口信息，不要复述正文）。")
        return self


class QuadrantCell(BaseModel):
    """2x2 网格的一个格子：核心指标大数字突出 + 一句说明。"""
    title: Str = Field(..., min_length=2, max_length=12,
                       description="格子标题")
    value: Str = Field(..., min_length=1, max_length=14,
                       description="核心指标值，如 85%、99.9%、5x（视觉层级：18pt 大字突出）")
    desc: Optional[Str] = Field(None, max_length=60,
                                description="一句话说明，≤60 字")
    icon: Optional[Str] = Field(None, max_length=32, description="格子图标")


class QuadrantGridSlide(BaseModel):
    """2x2 网格：强制 2×2 排列的四象限/矩阵布局。

    与 numbered_cards 的区别：
    - numbered_cards 4 张卡可以是 1×4、2×2、4×1 —— 由卡数决定
    - grid_2x2 **强制** 2×2 排列 —— 适用场景：四象限分析、2×2 矩阵、
      四个独立评估维度（如 SWOT、重要-紧急矩阵、优势-劣势-机会-威胁）

    视觉层级：
    - 格子标题 13pt（quadrant_title）
    - 核心指标值 18pt（quadrant_value，大字突出，比正文大 ≥2 级）
    - 格子说明 11pt
    """
    slide_type: Literal["grid_2x2"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    # 4 个格子，左上、右上、左下、右下
    q_top_left: QuadrantCell = Field(..., description="左上格子")
    q_top_right: QuadrantCell = Field(..., description="右上格子")
    q_bottom_left: QuadrantCell = Field(..., description="左下格子")
    q_bottom_right: QuadrantCell = Field(..., description="右下格子")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.grid_geom(self)
            cells = [self.q_top_left, self.q_top_right,
                     self.q_bottom_left, self.q_bottom_right]
            _cap_hint(cells, "desc", g["desc_w"], g["desc_h"],
                      "grid_2x2 的格子说明（q_*.desc）",
                      hint="一格一句话：说清指标含义、口径或对比基线；"
                           "结论写进 footer_note。")
            _cap_hint(cells, "value", g["cell_w"] - 0.5, 0.50,
                      "grid_2x2 的格子指标（q_*.value）",
                      hint="核心指标要能一眼读完（18pt 大字），别写整句。")
        return self


class Milestone(BaseModel):
    """时间轴上的一个节点：何时（time/label）发生了什么（desc）。"""
    label: Str = Field(..., min_length=2, max_length=14,
                       description="节点名/阶段名，如「二期上线」")
    time: Optional[Str] = Field(None, max_length=12,
                                description="时间标注，如 2024Q2 / 第 3 个月（可选）")
    desc: Str = Field(..., min_length=10, max_length=120,
                      description="这一步发生了什么/交付了什么；节点少时每个可写约 100 字，"
                                  "12 个节点时约 30 字（以 --qc 报的为准）")
    icon: Optional[Str] = Field(None, max_length=32, description="节点图标")


class TimelineSlide(BaseModel):
    """时间轴/里程碑页：**时序关系**的落点 —— 事件沿中轴上下交替排布。

    与 flow 的分工：flow 讲「步骤怎么走」（判定/分支/循环），timeline 讲
    「何时发生什么」（时间点 + 事件 + 交付物）。输入是阶段序号链（一期→二期）
    或带时间标注的里程碑时，优先用本版式而不是 flow。
    """
    slide_type: Literal["timeline"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    milestones: list[Milestone] = Field(..., min_length=3, max_length=12,
                                        description="3-12 个节点；节点就是输入里的"
                                                    "阶段/里程碑，不许删；超过 12 个才拆页")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.timeline_geom(self)
            _mode = "竖排" if g["mode"] == "vertical" else "横排"
            _cap_hint(self.milestones, "desc", g["desc_w"], g["desc_h"],
                      f"timeline 的 milestones[].desc（{len(self.milestones)} 个节点走"
                      f"{_mode}，每节点描述区 {g['desc_w']:.2f}×{g['desc_h']:.2f}in）",
                      hint=f"{_mode}时每节点只有这么点地方：写一句「谁做了什么、结果如何」"
                           f"就够，细节挪到 callout；节点少一点每格会宽一些。")
        return self


class FunnelStage(BaseModel):
    """漏斗的一层：这一层的量（value）和口径说明（desc）。"""
    label: Str = Field(..., min_length=2, max_length=10, description="层名，如「进线」")
    value: Optional[Str] = Field(None, max_length=10,
                                 description="该层关键数字，如 1,200 / 45%（大字突出）")
    desc: Str = Field(..., min_length=10, max_length=100,
                      description="这一层的口径/转化率/流失原因说明")
    icon: Optional[Str] = Field(None, max_length=32, description="层图标")


class FunnelSlide(BaseModel):
    """转化漏斗页：**转化/衰减关系**的落点 —— 层层收窄，右侧配口径说明。

    适用：进线→成交、报名→完课、线索→签约这类「每层都在变小」的流程；
    各层数值只写绝对量，转化率写进 desc（相邻两层一除就是）。
    """
    slide_type: Literal["funnel"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    stages: list[FunnelStage] = Field(..., min_length=3, max_length=8,
                                      description="3-8 层；层就是输入里的转化环节，不许删")
    callout: list[Callout] = Field(default_factory=list, max_length=4,
                                   description="底部横排辅区 0-4 条：关键数字/边界/待确认点")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.funnel_geom(self)
            _cap_hint(self.stages, "desc", g["desc_w"], g["desc_h"],
                      f"funnel 的 stages[].desc（{len(self.stages)} 层，右侧说明栏 "
                      f"{g['desc_w']:.2f}×{g['desc_h']:.2f}in）",
                      hint=f"层数越多每层越矮（现在 {g['bar_h']:.2f}in）：说明只写"
                           f"「转化率 + 关键原因」，多的挪到 callout。")
        return self


class MatrixCell(BaseModel):
    """矩阵的一个格子：行 × 列交叉处的结论。"""
    title: Str = Field(..., min_length=2, max_length=16, description="格子标题")
    desc: Str = Field(..., min_length=4, max_length=90,
                      description="交叉处的说明；格子多时每个可写约 40 字，"
                                  "以 --qc 报的为准")
    icon: Optional[Str] = Field(None, max_length=32, description="格子图标")


class MatrixRow(BaseModel):
    """矩阵的一行：行标 + 与列数等长的格子。"""
    label: Str = Field(..., min_length=1, max_length=8, description="行标")
    cells: list[MatrixCell] = Field(..., min_length=2, max_length=4,
                                    description="格子数必须等于列数")


class MatrixSlide(BaseModel):
    """二维矩阵页：**两维交叉分类**的落点 —— grid_2x2 的 m×n 泛化。

    适用：两个维度各自取值、交叉格有独立结论的分类（如「能力 × 阶段」评估矩阵）。
    行数/列数由输入决定；只有 2×2 且每格有核心指标大字时用 grid_2x2 更合适。
    """
    slide_type: Literal["matrix"]
    title: Str = Field(..., min_length=6, max_length=32, description="判断句标题")
    summary: Optional[Str] = Field(None, max_length=240, description="页面级结论段")
    col_labels: list[Str] = Field(..., min_length=2, max_length=4,
                                  description="列头 2-4 个，每个 ≤8 字")
    rows: list[MatrixRow] = Field(..., min_length=2, max_length=4,
                                  description="2-4 行；每行 cells 数必须等于列数")
    footer_note: Optional[Str] = Field(None, max_length=FOOTER_NOTE_MAX)

    @model_validator(mode="after")
    def check_row_width(self):
        for i, row in enumerate(self.rows):
            if len(row.cells) != len(self.col_labels):
                raise ValueError(
                    f"第 {i+1} 行「{row.label}」有 {len(row.cells)} 个格子，"
                    f"col_labels 是 {len(self.col_labels)} 列，必须一致")
        return self

    @model_validator(mode="after")
    def fit_capacity(self):
        L = _layouts()
        if L is not None:
            g = L.matrix_geom(self)
            cells = [c for row in self.rows for c in row.cells]
            _cap_hint(cells, "desc", g["desc_w"], g["desc_h"],
                      f"matrix 的 rows[].cells[].desc（{len(self.rows)}×"
                      f"{len(self.col_labels)} 矩阵，每格 "
                      f"{g['desc_w']:.2f}×{g['desc_h']:.2f}in）",
                      hint="格子小、行列多：一格一句话（结论或做法），"
                           "交叉读法写进 footer_note。")
        return self


Slide = Annotated[
    Union[
        CoverSlide, TocSlide, SectionSlide, StatementSlide, CompareSlide,
        NumberedCardsSlide, LayeredStackSlide, FlowSlide, HierarchySlide,
        TableSlide, ChartSlide, Panel2ColSlide, FigureSlide, QuadrantGridSlide,
        TimelineSlide, FunnelSlide, MatrixSlide, KeyGroupsSlide, EndSlide,
    ],
    Field(discriminator="slide_type"),
]


class SlideNotes(BaseModel):
    """单页备注：写入 .pptx 备注区，供演讲者参考。"""
    data_gaps: list[Str] = Field(default_factory=list,
                                  description="缺少哪些数字/数据")
    viz_suggestions: list[Str] = Field(default_factory=list,
                                       description="可视化改进建议")
    source_suggestions: list[Str] = Field(default_factory=list,
                                         description="建议补充哪些外部资料")
    content_gaps: list[Str] = Field(default_factory=list,
                                    description="内容缺口（来自分析表）")

    def render(self) -> str:
        """渲染为纯文本，写入备注区。"""
        lines = []
        if self.data_gaps:
            lines.append("【数据缺口】")
            for g in self.data_gaps:
                lines.append(f"  • {g}")
        if self.viz_suggestions:
            lines.append("【可视化建议】")
            for v in self.viz_suggestions:
                lines.append(f"  • {v}")
        if self.source_suggestions:
            lines.append("【补充资料】")
            for s in self.source_suggestions:
                lines.append(f"  • {s}")
        if self.content_gaps:
            lines.append("【内容缺口】")
            for c in self.content_gaps:
                lines.append(f"  • {c}")
        return "\n".join(lines)


class Deck(BaseModel):
    title: Str = Field(..., min_length=2, max_length=30)
    subtitle: Optional[Str] = Field(None, max_length=60)
    style: Str = Field("style-harness", max_length=40,
                       description="配色主题，见 theme.py：style-harness（藏青，默认）/"
                                   "blue-business（深蓝商务）/ bright（明亮多色）/"
                                   "forest（墨绿稳重）/ violet-tech（紫灰科技）；"
                                   "写别的名字一律回退 style-harness。别在 deck.json 里发明主题名")
    card: Str = Field("flat", max_length=20,
                      description="卡片长相，见 theme.CARD_STYLES：flat（白卡细描边，默认）/"
                                  "outline（透明底描边）/ tinted（浅色底无边）/"
                                  "accent-left（左侧色条）/ accent-top（顶部色条）；"
                                  "写别的名字一律回退 flat")
    brief: Optional[Str] = Field(None, max_length=40,
                                 description="写作方向：seminar=研讨型 / briefing=汇报型，"
                                             "见 .pi/renderer/briefs/。决定取舍规则与页数预算，"
                                             "不影响外观")
    slides: list[Slide] = Field(..., min_length=1)
    # 备注：每页对应的建议（可选，与 slides 一一对应）
    speaker_notes: list[SlideNotes] = Field(default_factory=list,
                                            description="每页备注，与 slides 一一对应")


def _check_row_widths(slide):
    """表格类版式的每行长度必须等于表头数量（明细表同理）。"""
    if isinstance(slide, TableSlide):
        slide.validate_row_width()
    elif isinstance(slide, ChartSlide) and slide.detail_headers:
        for i, row in enumerate(slide.detail_rows):
            if len(row) != len(slide.detail_headers):
                raise ValueError(
                    f"chart.detail_rows 第 {i + 1} 行有 {len(row)} 列，"
                    f"detail_headers 是 {len(slide.detail_headers)} 列，必须一致")


def load_deck(path: str) -> Deck:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    deck = Deck.model_validate(data)
    for s in deck.slides:
        _check_row_widths(s)
    return deck


def validate_dict(data: dict) -> Deck:
    deck = Deck.model_validate(data)
    for s in deck.slides:
        _check_row_widths(s)
    return deck


if __name__ == "__main__":
    # 用法: python ppt_dsl.py deck.json   —— 校验通过打印 OK，失败打印全部错误
    if len(sys.argv) < 2:
        print("用法: python ppt_dsl.py <deck.json>")
        sys.exit(2)
    try:
        deck = load_deck(sys.argv[1])
        print(f"OK: {len(deck.slides)} 页，主题 {deck.style}，卡片 {deck.card}")
        for i, s in enumerate(deck.slides, 1):
            print(f"  {i:>2}. {s.slide_type:<15} {getattr(s, 'title', getattr(s, 'text', ''))[:24]}")
    except (ValidationError, ValueError) as e:
        print("DSL 校验失败：")
        print(e)
        sys.exit(1)
