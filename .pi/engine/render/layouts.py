# -*- coding: utf-8 -*-
"""版式渲染函数：每种 slide_type 一个函数，只依赖 python-pptx 原生形状。

设计原则：
- 浅灰页面底 + 白卡片 + 彩色点缀（取自参考 PDF 实测色板）
- 高信息密度：正文 11.5-12.5pt，卡片网格，图示优先
- 所有几何常量集中定义，qc.py 从这里取容量做溢出估算
"""
from __future__ import annotations

import io
import math
import os
import re

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from ..paths import PI_ROOT
from . import theme
from .icons import draw_icon
from .theme import (ACCENT_CYCLE, CHART_SERIES, COLORS, CONTENT_W, FONT_LATIN,
                    FONT_MAIN, FONT_SIZES, LANG_ALT, LANG_MAIN, MARGIN_L, tint)

# 测量与容量全部来自 geom（含 `_est_lines` 这类内部工具，故 geom 里显式写了 __all__）。
# 绘制代码沿用原来的裸名字，不必改成 geom.xxx()。
from .geom import *  # noqa: F401,F403

In = Inches

# 图表渲染开关：native=可编辑原生图表（默认），mpl=matplotlib 高质感 PNG。
# 归 layouts 而不归 geom：它是"画法"开关，几何计算与它无关。
CHART_MODE = "native"

# 页面尺寸与背景图路径（封面用 bg-cover，其余页用 bg-content；缺图回退纯色）。
SLIDE_W, SLIDE_H = 13.333, 7.5


                                      # （实测水印 x 10.95–12.90 / y 6.97–7.15）



# ---------------- 基础工具 ----------------

def _c(name: str) -> RGBColor:
    return RGBColor.from_string(COLORS.get(name, name))






def _set_font_on_rPr(rPr, text: str | None = None) -> None:
    """把 `a:latin` / `a:ea` / `a:cs` 三处 typeface 成对写上（顺序按 ECMA-376），
    并给这个 rPr 写语言标记（`_set_lang`）。

    PowerPoint 按**字符**挑字体：西文/数字走 `a:latin`、中文走 `a:ea`。所以每个 run
    都要成对写死，不能只写一个 —— 只写 latin，中文会退回主题字体（模板自带 Calibri，
    中文就变成 PowerPoint 的替身字体）；只写 ea，英文/数字会变 Calibri。

    元素顺序必须是 `latin → ea → cs`（CT_TextCharacterProperties 的顺序约束），
    所以统一用 `latin.addnext(...)` 插进去，而不是 append 到 rPr 末尾
    （rPr 末尾可能是 `a:sym` / `a:hlinkClick`）。
    """
    if text is not None:
        _set_lang(rPr, text)
    elif not rPr.get("lang"):
        _set_lang(rPr)                    # 图表等无文本上下文：按中文默认值补
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = etree.Element(qn("a:latin"))
        tail_tags = ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")
        anchor = next((c for c in rPr if c.tag in {qn(t) for t in tail_tags}), None)
        if anchor is not None:                 # latin 必须排在 sym/hlink/… 之前
            anchor.addprevious(latin)
        else:
            rPr.append(latin)
    latin.set("typeface", FONT_LATIN)
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.Element(qn("a:ea"))
        latin.addnext(ea)
    ea.set("typeface", FONT_MAIN)
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = etree.Element(qn("a:cs"))
        ea.addnext(cs)
    cs.set("typeface", FONT_MAIN)


def apply_run_font(run) -> None:
    """给一个 run 写上「中文微软雅黑 + 西文 Times New Roman」+ 语言标记。

    全 deck 唯一的字体入口：`text_box` / `_fill_cell` / 页码（render.py）都走它。
    自己新建 run 时**必须**调一次，否则那一段会退回模板字体（Calibri）；
    语言标记按 run 的实际文本判定（纯西文 → en-US，含中文 → zh-CN）。
    """
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:                          # python-pptx 的 run.font.name 会建 a:latin
        run.font.name = FONT_LATIN
    _set_font_on_rPr(rPr, run.text)


def _mark_language_on(root) -> int:
    """给一棵 XML 树里**所有**段落补语言标记，返回改动处数。

    为什么要扫整页、而不是只在 `apply_run_font` 里写：`python-pptx` 只为带文字的
    run 建 `a:rPr`，段落末尾的 `a:endParaRPr` 它从不写。PowerPoint 在文本框末尾
    （光标停在最后一段之后）或选中整个外框时，Ribbon 字体框读的正是
    `a:endParaRPr` —— 缺了它，整框文字仍会显示乱码，哪怕每个 run 都标了语言。

    段落里残留的 `a:rPr`（`lang="en-US"` 的模板默认值）也按实际文本纠正成中文。
    """
    changed = 0
    for p in root.iter(qn("a:p")):
        para_text = "".join(t.text or "" for t in p.iter(qn("a:t")))
        lang, alt = (LANG_ALT, LANG_MAIN) if (para_text and not _has_cjk(para_text)) \
            else (LANG_MAIN, LANG_ALT)
        end = p.find(qn("a:endParaRPr"))
        if end is None:
            end = etree.SubElement(p, qn("a:endParaRPr"))   # 必须是 a:p 最后一个子元素
        if end.get("lang") != lang:
            end.set("lang", lang)
            end.set("altLang", alt)
            changed += 1
        for r in p.iter(qn("a:r")):
            rPr = r.find(qn("a:rPr"))
            if rPr is None:
                continue
            run_text = "".join(t.text or "" for t in r.iter(qn("a:t")))
            want = LANG_ALT if (run_text and not _has_cjk(run_text)) else LANG_MAIN
            if rPr.get("lang") != want:
                rPr.set("lang", want)
                rPr.set("altLang", LANG_MAIN if want == LANG_ALT else LANG_ALT)
                changed += 1
        for d in p.iter(qn("a:defRPr")):
            if d.get("lang") != lang:
                d.set("lang", lang)
                d.set("altLang", alt)
                changed += 1
    return changed


def mark_text_language(container) -> int:
    """给一页（或备注页）补语言标记：`slide` / `notes_slide` 都行。"""
    return _mark_language_on(container._element)








# ---------------- 图标自动匹配 ----------------
#
# ---- 图标配额：一个内容块最多一个图标 ----
# 每个「块」（卡片 / 栏 / 层 / 步骤 / 节点 / 正文分组）只画一个图标，位置固定在块头：
# 块自带 icon 字段就用它，没写就按块的文字自动挑一个补上（保证每块都有图标）。
# 正文 / 描述文字前面**不再额外补图标** —— 块头已经有图标时，第二个图标会和它
# 打架，整页看着发飘（v14 的 P9 / P17 / P19 就是这样）。












def _draw_icon_checked(slide, icon_name: str, x: float, y: float, size: float,
                       color: str = "primary", stroke_w: float | None = None) -> bool:
    """绘制图标，成功返回 True。失败时返回 False（不抛错）。"""
    if not icon_name:
        return False
    try:
        return draw_icon(slide, icon_name, x, y, size, color=color, stroke_w=stroke_w)
    except Exception:
        return False


def _draw_block_icon(slide, icon_name: str | None, x: float, y: float, size: float,
                     color: str = "primary", stroke_w: float | None = None) -> bool:
    """画「块头唯一图标」：**画不出来就退化成同色圆角色块，块头绝不空着**。

    这是「一块一个图标」的最后一道保障：图标名解析失败、SVG 解析报错、字体缺字……
    任何一种情况都只影响图标画法，不该让这一块看着像漏了东西。
    占用尺寸与真图标一致，所以退化不会改变任何版式几何。
    """
    if _draw_icon_checked(slide, icon_name, x, y, size, color=color, stroke_w=stroke_w):
        return True
    rect(slide, x, y, size, size, fill=color, radius=0.28,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    return True


















def draw_card_text(slide, x, y, w, h, text, size_pt, spacing, *, key="primary",
                   claim=False, color="text") -> float:
    """画卡内文本：同字号，关键句/关键短语只**加粗 + 强调色**（不置顶、不放大）。"""
    paras = card_paras(text, size_pt, claim=claim, key=key, color=color)
    if not paras:
        return y
    text_box(slide, x, y, w, h, paras, size=size_pt, spacing=spacing,
             strong_color=key)
    return y














def text_box(slide, x, y, w, h, content, size=FONT_SIZES["body"], color="text", bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.3,
             italic=False, space_after=0.0, strong_color=None):
    """content: str 或 [(text, {覆盖参数}), ...] 列表（多段落）。

    **字体不由调用方指定**：每个 run 都写成「中文微软雅黑 + 西文 Times New Roman」
    （`_apply_run_font`），由 PowerPoint 按字符自动挑 —— 这是全 deck 唯一的字体口径，
    别再引入 `font=` 之类的覆盖参数（历史上"号码用雅黑、正文用思源"就是这么漂移的）。

    `strong_color`：`**...**` 里的加粗片段用的**强调色**（关键句/关键短语靠它区分，
    字号与正文一致 —— 放大字号会把文字顶出卡片背景）。
    """
    box = slide.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = content if isinstance(content, list) else [(content, {})]
    for i, (t, kw) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = kw.get("align", align)
        p.line_spacing = kw.get("spacing", spacing)
        p.space_after = Pt(kw.get("space_after", space_after))
        size_pt = kw.get("size", size)
        is_bold = kw.get("bold", bold)
        is_italic = kw.get("italic", italic)
        run_color = kw.get("color", color)
        # 支持 **加粗** 内联标记：按 ** 切分，奇数段加粗（并可用强调色）
        parts = t.split("**") if "**" in t else [t]
        sc = kw.get("strong_color", strong_color)
        for j, part in enumerate(parts):
            if not part:
                continue
            run = p.add_run()
            run.text = part
            f = run.font
            f.size = Pt(size_pt)
            is_strong = j % 2 == 1
            f.bold = is_bold or is_strong
            f.italic = is_italic
            f.color.rgb = _c(sc) if (is_strong and sc) else _c(run_color)
            apply_run_font(run)
    return box


def rect(slide, x, y, w, h, fill: str | None = "card_bg", line: str | None = None,
         line_w=0.75, radius=None, shape=MSO_SHAPE.RECTANGLE):
    """基础形状。`fill=None` = 无填充（outline 卡片就靠它透出页面底）。"""
    sp = slide.shapes.add_shape(shape, In(x), In(y), In(w), In(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _c(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = _c(line)
        sp.line.width = Pt(line_w)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    sp.text_frame.word_wrap = True
    return sp


                               # 否则等于没换长相；卡内首行文字上内边距最小 0.10，压不到字）


def card(slide, x, y, w, h, accent="primary", radius=None, bar=False,
         rail_w=CARD_RAIL_W):
    """标准卡片。长相由 `deck.card`（`theme.CURRENT_CARD`）决定。

    accent : 色条用的颜色名（`accent-left` / `accent-top` / `bar` 才用）。调用点知道本卡的
             强调色就传进来，让色条跟着编号/层级一起循环。
    radius : 仅 `flat` 生效，取调用点原有的圆角（保证老产出逐像素不变）；
             其他长相一律用长相自带的圆角，跨版式观感才一致。
    bar    : True = "这张卡天生带一条左轨"（清单感/时间轴行/编号卡的版式语义），
             与卡片长相无关：flat 下仍是历史上的方角轨道，
             accent-left 下换成胶囊轨。面板内嵌的小卡不要传。
    rail_w : `bar=True` 时的历史轨道宽度（各类卡原本就不一样，别统一，否则老产出会变）。
    """
    style = theme.CURRENT_CARD
    if style == "outline":
        # 无填充而不是白填充：内容页底是近白底图，透出来才有"轻"的质感
        sp = rect(slide, x, y, w, h, fill=None,
                  line=tint(COLORS["primary"], 0.55), line_w=1.0,
                  radius=CARD_RADIUS_OUTLINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    elif style == "tinted":
        # 浅色成块：.90 是实测的下限 —— 再浅（.94）在 F2F2F2 页面底上就看不见卡边了。
        # 与辅区色块（.90~.93）同档，但辅区是"块内含卡片"，卡片仍是实色，不会糊在一起。
        sp = rect(slide, x, y, w, h, fill=tint(COLORS["primary"], 0.90),
                  radius=CARD_RADIUS_TINTED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    elif style in ("accent-left", "accent-top"):
        # 两种 "accent-*" 都是"白卡 + 方角 + 一条色带"，只是色带贴左边还是贴顶边
        sp = rect(slide, x, y, w, h, fill="card_bg",
                  line=None if style == "accent-top" else "border",
                  radius=CARD_RADIUS_OUTLINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    else:
        sp = rect(slide, x, y, w, h, fill="card_bg", line="border",
                  radius=CARD_RADIUS_FLAT if radius is None else radius,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE)

    if style == "accent-left" or bar:
        if style == "accent-left":
            # 胶囊竖轨：上下各留一点，避免顶到圆角
            rail_h = max(0.20, h - 0.26)
            rect(slide, x + CARD_ACCENT_PAD, y + (h - rail_h) / 2,
                 CARD_ACCENT_W, rail_h, fill=accent, radius=0.5,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        else:
            # 历史方角左轨：与改动前完全一致（通高、方角、原本的宽度）
            rect(slide, x, y, rail_w, h, fill=accent)
    if style == "accent-top" and not bar:
        # 通宽顶带：多数版式自己会在卡顶压一条 0.05~0.075 的色条，
        # 顶带做成 0.09 且同色（调用点把 accent 传成本卡的 key），两者叠成一条更厚的带子
        rect(slide, x, y, w, CARD_ACCENT_TOP_H, fill=accent)
    return sp


def page_bg(slide, kind="content"):
    """铺整页背景：封面 bg-cover.png，其余页 bg-content.png；缺图回退纯色。

    图片作为最先加入的形状，天然压在底层，后续元素绘制在其上。
    """
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _c("bg_page")
    path = _BG_COVER if kind == "cover" else _BG_CONTENT
    if os.path.exists(path):
        slide.shapes.add_picture(path, In(0), In(0), In(SLIDE_W), In(SLIDE_H))








def content_header(slide, title: str, summary: str | None = None) -> float:
    """内容页统一页头：左侧竖色条 + 判断句标题 + 核心结论段（浅色块+浅色边框打底）。

    结论段是 2~3 句 / 40~80 字的段落（非一句话），支持 **重点句** 内联加粗；块高随文字量自适应。
    返回内容区起始 y。
    """
    page_bg(slide)
    rect(slide, 0.55, 0.46, 0.075, 0.34, fill="primary")
    text_box(slide, 0.78, 0.38, 11.9, 0.55, title,
             size=FONT_SIZES["page_title"], color="text", bold=True, spacing=1.1)
    if summary:
        block_h = summary_block_h(summary)
        rect(slide, 0.78, 1.02, 11.9, block_h, fill=tint(COLORS["primary"], 0.90),
             line="border", line_w=1.0,
             radius=0.12, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, 1.02, 1.02, 11.42, block_h, summary,
                 size=FONT_SIZES["summary"], color="text", spacing=1.32,
                 anchor=MSO_ANCHOR.MIDDLE)
    return header_y(summary)


def footer_note(slide, note: str):
    """页脚条：压扁到 FOOTER_H、下移到 FOOTER_Y，右端收在 FOOTER_RIGHT。

    左端给页码留位（页码由 render.py 画在 0.55,FOOTER_Y），右端避开底图右下角的
    单位水印。高度只够一行 —— 字号从 11pt 往下缩到放得下为止（最低 9pt），
    所以 `footer_note` 字段的上限收到 64 字（见 ppt_dsl）。
    """
    w = FOOTER_RIGHT - 1.25
    rect(slide, 1.25, FOOTER_Y, w, FOOTER_H, fill=tint(COLORS["primary"], 0.93),
         radius=0.16, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tw = w - 0.25 - 0.25
    size = 11.0
    while size > 9.0 and _lines_in(note, tw, size) > 1:
        size -= 0.5
    text_box(slide, 1.5, FOOTER_Y, tw, FOOTER_H, note, size=size,
             color="dark_blue", anchor=MSO_ANCHOR.MIDDLE, spacing=1.1)


# ---------------- 各版式 ----------------

_cover_dark_cache: bool | None = None


def _cover_ink() -> dict:
    """按封面底图的明暗挑字色，返回 {title,en,sub,org} 四个色名。

    底图是发布在 .pi/assets/template/ 里的固定素材，深浅由素材决定而不是写死：
    深底用浅色字、浅底用深色字，换模板时不会出现「黑字压深蓝」这种看不清的情况。
    """
    global _cover_dark_cache
    if _cover_dark_cache is None:
        dark = True  # 缺图时 page_bg 回退浅灰底，这里同步回退到深色字
        if os.path.exists(_BG_COVER):
            try:
                from PIL import Image
                with Image.open(_BG_COVER) as im:
                    small = im.convert("L").resize((32, 18))
                    # 只在文字所在的中间横带取样，避开顶部单位 logo
                    mid = small.crop((0, 6, 32, 13))
                    px = list(mid.getdata())
                    dark = (sum(px) / len(px)) < 128
            except Exception:
                dark = True
        _cover_dark_cache = dark
    suffix = "" if _cover_dark_cache else "_on_light"
    return {k: f"cover_{k}{suffix}" for k in ("title", "en", "sub", "org")}


def render_cover(slide, m):
    page_bg(slide, "cover")
    ink = _cover_ink()
    # 顶部三色装饰条：居中压在英文题上方（原来靠左，与居中的标题块不成组）
    bar_w, bar_gap = 0.36, 0.14
    total = 3 * bar_w + 2 * bar_gap
    for i, key in enumerate(["primary", "green", "orange"]):
        rect(slide, 6.667 - total / 2 + i * (bar_w + bar_gap), 2.08, bar_w, 0.075,
             fill=tint(COLORS[key], 0.42))
    if m.en_title:
        text_box(slide, 0.55, 2.32, 12.23, 0.42, m.en_title.upper(),
                 size=FONT_SIZES["cover_en"], color=ink["en"], bold=True,
                 align=PP_ALIGN.CENTER)
    text_box(slide, 0.55, 2.86, 12.23, 1.32, m.title,
             size=FONT_SIZES["cover_title"], color=ink["title"], bold=True,
             align=PP_ALIGN.CENTER, spacing=1.18)
    if m.subtitle:
        text_box(slide, 1.5, 4.34, 10.33, 0.6, m.subtitle,
                 size=FONT_SIZES["cover_sub"], color=ink["sub"],
                 align=PP_ALIGN.CENTER, spacing=1.3)
    if m.org:
        # 落款：居中细线 + 落款文字（原来铺一条浅色带，在深蓝封面上很跳）
        rect(slide, 6.667 - 0.45, 6.72, 0.9, 0.018,
             fill=tint(COLORS["primary"], 0.45))
        text_box(slide, 0.55, 6.92, 12.23, 0.42, m.org, size=FONT_SIZES["cover_org"],
                 color=ink["org"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)




def _render_toc_item(slide, item, x, y, w, h, key):
    """一条目录项：浅色底 + 序号 + 图标 + 标题。"""
    rect(slide, x, y, w, h, fill=tint(COLORS[key], 0.90),
         line=key, line_w=1.0, radius=0.1, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text_box(slide, x + 0.16, y, 0.52, h, item.no, size=FONT_SIZES["toc_no"],
             color=key, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    tx0 = x + 0.86
    if item.icon:
        _draw_block_icon(slide, item.icon, x + 0.72, y + (h - 0.44) / 2, 0.44, color=key)
        tx0 = x + 1.30
    text_box(slide, tx0, y, x + w - tx0 - 0.16, h, item.text,
             size=FONT_SIZES["toc_text"], color="text",
             anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)


def render_toc(slide, m):
    """目录页：两列铺开，行高按条目数**撑满到内容带**。

    三处老毛病一起修：
    1. 第二列右沿原来算到 13.10in，越过内容右沿 0.32in（贴边）；
    2. 4 项时两行只占 1.68in，下半页全空；
    3. 奇数条目最后一条孤零零留在左边，右侧空半行，底部还拖一条 5in 的装饰横线。
    现在奇数时末条**跨整行**，整块在内容带里垂直居中。
    """
    page_bg(slide)
    rect(slide, 0.55, 0.55, 0.075, 0.4, fill="primary")
    text_box(slide, 0.78, 0.48, 6, 0.6, "目录", size=FONT_SIZES["toc_title"],
             color="text", bold=True, spacing=1.1)
    text_box(slide, 0.8, 1.12, 6, 0.4, "CONTENTS", size=FONT_SIZES["toc_en"],
             color="text_light")
    from .theme import ACCENT_CYCLE
    n = len(m.items)
    rows = max(1, math.ceil(n / 2))
    x0 = TOC_X0
    right = MARGIN_L + CONTENT_W
    col_w = (right - x0 - TOC_GAP_X) / 2
    full_w = right - x0
    avail = TOC_BOTTOM - TOC_TOP
    row_h = (avail - (rows - 1) * TOC_GAP_Y) / rows
    block_h = max(0.62, min(TOC_ROW_MAX, row_h))
    total = block_h * rows + TOC_GAP_Y * (rows - 1)
    y0 = TOC_TOP + max(0.0, (avail - total) / 2)
    for i, item in enumerate(m.items):
        col, row = i % 2, i // 2
        single_last = n % 2 == 1 and i == n - 1 and rows > 1
        w = full_w if single_last else col_w
        x = x0 if (col == 0 or single_last) else x0 + col_w + TOC_GAP_X
        y = y0 + row * (block_h + TOC_GAP_Y)
        _render_toc_item(slide, item, x, y, w, block_h,
                         ACCENT_CYCLE[i % len(ACCENT_CYCLE)])


def render_section(slide, m):
    page_bg(slide)
    text_box(slide, 7.6, 1.7, 5.2, 2.4, m.no, size=FONT_SIZES["section_no"],
             color="border", bold=True, align=PP_ALIGN.RIGHT)
    rect(slide, 0.9, 2.75, 0.55, 0.1, fill="primary")
    # 图标放到文字前面（标题左侧）
    if m.icon:
        _draw_block_icon(slide, m.icon, 0.95, 2.95, 0.78, color="primary", stroke_w=1.0)
        tx = 2.0
    else:
        tx = 0.9
    text_box(slide, tx, 3.08, 12.23 - tx, 0.9, m.title, size=FONT_SIZES["section_title"],
             color="text", bold=True, spacing=1.15)
    if m.subtitle:
        text_box(slide, tx, 4.25, 8.2, 0.6, m.subtitle, size=FONT_SIZES["section_sub"],
                 color="text_sub", spacing=1.3)




# ---------------- 辅区（一页的第二块信息）落位：**左右优先** ----------------
#
# 一页要装两块内容时，辅区**默认放右侧纵栏**，主内容占左边。不是审美偏好，
# 是因为底部横条会**吃掉主内容的高度**：流程图（尤其环形）被压扁、三栏对比的描述
# 被挤成一句 —— v20 之前 P5 的环形图就是这么变难看的。
# 只有主内容本身**必须靠宽度才成立**时（线性流程 ≥6 步、层级卡 ≥6 个、
# 三栏对比、双模块并排），才退回底部横条；那时主内容高度仍然自适应，不被压。
#
# 落位由每个版式自己声明「主内容至少要多宽」（main_min_w），渲染与 qc 共用同一套几何。
# AUX_RAIL_W / AUX_GAP / AUX_MIN_MAIN_W 定义在文件顶部「页面骨架」区
# （statement_geom / highlight_h 在更靠前的位置就要用到）。






def draw_aux(slide, notes, fr):
    """按 card_frame 定好的落位画辅区（右栏 / 底部横条）。"""
    if not notes or fr["mode"] == "none":
        return
    if fr["mode"] == "rail":
        _render_rail(slide, notes, fr["rail_x"], fr["rail_top"], fr["rail_w"],
                     fr["rail_bottom"])
    else:
        _render_callout(slide, notes, fr["bottom"] + fr["aux_h"] + 0.10, fr["aux_h"])






























def _draw_statement_body(slide, items, list_mode, x, y, w, h, size_pt,
                         spacing, *, claim):
    """statement 卡内正文：条目式画色点清单，否则整段流（关键句加粗）。"""
    if not items:
        return
    if not list_mode:
        draw_card_text(slide, x, y, w, h, items[0], size_pt, spacing,
                       claim=claim)
        return
    ty = y
    for it in items:
        rect(slide, x, ty + size_pt / 72 * spacing / 2 - 0.028,
             0.056, 0.056, fill="primary", shape=MSO_SHAPE.OVAL)
        text_box(slide, x + 0.13, ty, w - 0.13, max(0.2, h - (ty - y)), it,
                 size=size_pt, spacing=spacing, strong_color="primary")
        ty += _para_h(it, w - 0.13, size_pt, spacing) + 0.05


def render_statement(slide, m):
    """结论页：左侧正文卡（内容优先挑档）+ 右栏支撑论据 + 右栏底部高亮数据卡。

    排版原则：**内容优先**（先算文字需要多少高度，再决定框多大、字多大），
    而不是"先分框、再把字塞进去" —— 后者必然出现三种毛病：
    估行数偏小→文字溢出边框；框被拉长→空框；靠 MIDDLE 居中遮丑→图标在顶、
    文字在中间，二者对不齐（v15 的 P4/P11）。
    """
    content_header(slide, m.title, m.summary)
    g = statement_geom(m)
    groups = split_body_groups(m.body)
    n_groups = max(1, len(groups))
    n_cols = 2 if n_groups >= 3 else 1
    gap_x = 0.12
    card_w = (g["w"] - (n_cols - 1) * gap_x) / n_cols
    avail = g["bottom"] - g["top"]
    icon_size, icon_gap = 0.26, 0.09

    # ---- 内容优先：先按文字算每行需要多高，再挑一档放得下的字号 ----
    full_last = n_cols > 1 and n_groups % n_cols == 1
    row_hs, size_pt, spacing, pad = _fit_rows(groups, n_cols, card_w, avail,
                                              full_last=full_last)
    row_hs, row_gap = _stack_heights(row_hs, avail)
    pad_t = pad_b = pad

    tops, acc = [], g["top"]
    for rh in row_hs:
        tops.append(acc)
        acc += rh + row_gap

    # ---- 正文卡片：文字 TOP 锚定，卡头行 = 图标 + 组标题（可扫读的骨架）----
    for g_idx, group in enumerate(groups):
        col, row = (g_idx % n_cols, g_idx // n_cols) if n_cols > 1 else (0, g_idx)
        full_row = n_cols > 1 and g_idx == n_groups - 1 and n_groups % n_cols == 1
        cx = g["x"] + col * (card_w + gap_x)
        c_w, card_h = (g["w"] if full_row else card_w), row_hs[row]
        cy = tops[row]
        card(slide, cx, cy, c_w, card_h, radius=0.05)
        head, items, list_mode = statement_card_items(group)
        # 一块一个图标：按本卡文本自动挑，画不出就退化成色块（块头绝不空）。
        icon_name = _resolve_icon(None, group) or "file-text"
        if head:
            # 卡头行：图标 + 组标题（首句，加粗染 primary 色）
            tx0 = cx + 0.10
            line_h = 11 / 72 * 1.2
            if _draw_block_icon(slide, icon_name, tx0,
                                cy + pad_t + (line_h - icon_size) / 2,
                                icon_size, color="primary"):
                tx0 += icon_size + icon_gap
            text_box(slide, tx0, cy + pad_t, c_w - (tx0 - cx) - 0.14, 0.30,
                     head, size=FONT_SIZES["stmt_head"], color="primary", bold=True)
            # 卡内正文：条目式（色点清单）或整段流，整宽排
            _draw_statement_body(slide, items, list_mode,
                                 cx + 0.14, cy + pad_t + STMT_HEAD_ROW,
                                 c_w - 0.28, card_h - pad_t - pad_b - STMT_HEAD_ROW,
                                 size_pt, spacing, claim=False)
        else:
            # 无卡头：图标与首行同高（旧口径），首句加粗染色（claim）
            tx0 = cx + 0.10
            line_h = size_pt / 72 * spacing
            if _draw_block_icon(slide, icon_name, tx0,
                                cy + pad_t + (line_h - icon_size) / 2,
                                icon_size, color="primary"):
                tx0 += icon_size + icon_gap
            tw = c_w - (tx0 - cx) - 0.14
            _draw_statement_body(slide, items, list_mode,
                                 tx0, cy + pad_t, tw, card_h - pad_t - pad_b,
                                 size_pt, spacing, claim=True)

    # ---- 高亮数据卡：右栏底部（栏窄，大数字按栏宽收敛字号）----
    if m.highlight is not None and g["hl_y"] is not None:
        hl, hy, rw = m.highlight, g["hl_y"], g["rw"]
        rect(slide, g["rx"], hy, rw, g["hl_h"], fill=tint(COLORS["primary"], 0.93),
             radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        hl_icon = _resolve_icon(hl.icon, hl.label, hl.value or "", hl.desc or "")
        label_x = g["rx"] + 0.14
        if hl_icon and _draw_icon_checked(slide, hl_icon, g["rx"] + 0.14, hy + 0.12,
                                          0.26, color="primary"):
            label_x = g["rx"] + 0.14 + 0.32
        vl_w = rw - (label_x - g["rx"]) - 0.20
        text_box(slide, label_x, hy + 0.10, vl_w, HL_LABEL_H, hl.label,
                 size=FONT_SIZES["rail_label"], color="text_sub", bold=True)
        # 20pt 大数字在 3.55in 栏里可能一行放不下 —— 逐档缩到放得下（最低 13pt）
        vsize = 20.0
        while vsize > 13.0 and _lines_in(hl.value or "", vl_w, vsize) > 1:
            vsize -= 1.0
        text_box(slide, label_x, hy + 0.30, vl_w, HL_VALUE_H, hl.value,
                 size=vsize, color="primary", bold=True)
        if hl.desc:
            text_box(slide, label_x, hy + HL_DESC_TOP, vl_w,
                     max(0.24, g["hl_h"] - HL_DESC_TOP - 0.06), hl.desc,
                     size=HL_DESC_SIZE, color="text_sub", spacing=HL_DESC_SPACING)

    # ---- 右栏支撑论据：加小标题成组，卡片同样紧凑、与高亮卡同栏 ----
    if m.points:
        rx, rw = g["rx"], g["rw"]
        text_box(slide, rx, g["top"], rw, 0.26, "支撑论据", size=FONT_SIZES["block_subtitle"],
                 color="primary", bold=True)
        rect(slide, rx, g["top"] + 0.29, 0.52, 0.045, fill="primary")
        pts_top, pts_bottom = g["pts_top"], g["pts_bottom"]
        # 每条论据需要的高度：标签 + 正文 + 上下内边距
        need_h = []
        for pt in m.points:
            iw = rw - (0.62 if pt.icon or _auto_icon(pt.label + pt.text) else 0.36)
            need_h.append(0.16 + _para_h(pt.label, iw, 12, 1.2)
                          + _para_h(pt.text, iw, 10.5, 1.25) + 0.16)
        phs, pgap = _stack_heights(need_h, pts_bottom - pts_top, gap_min=0.12)
        py = pts_top
        for i, (pt, ph) in enumerate(zip(m.points, phs)):
            card(slide, rx, py, rw, ph,
                 accent=ACCENT_CYCLE[(i + 1) % len(ACCENT_CYCLE)],
                 radius=0.09, bar=True)
            # 一块一个图标：画在块头（左上），与标签同一行
            pt_icon = _resolve_icon(pt.icon, pt.label, pt.text)
            tx0, tw_ = rx + 0.24, rw - 0.45
            if pt_icon and _draw_icon_checked(slide, pt_icon, rx + 0.18, py + 0.17,
                                              0.36, color="primary"):
                tx0, tw_ = rx + 0.64, rw - 0.86
            text_box(slide, tx0, py + 0.12, tw_, 0.28, pt.label,
                     size=FONT_SIZES["block_subtitle"], color="primary", bold=True)
            text_box(slide, tx0, py + 0.42, tw_, ph - 0.54, pt.text,
                     size=FONT_SIZES["rail_text"], color="text_sub", spacing=1.2)
            py += ph + pgap
    if m.footer_note:
        footer_note(slide, m.footer_note)










def render_compare_3col(slide, m):
    from .theme import ACCENT_CYCLE
    content_header(slide, m.title, m.summary)
    g = COMPARE_GEOM
    gap = COMPARE_GAP
    # 三栏对比**不放辅区**（schema 已禁 `callout`）：三栏并排本身就是三块信息，
    # 再来一条底部横条会把每栏描述容量砍到 1/5。收尾用 footer_note。
    fr = card_frame(m.summary, False, COMPARE_MIN_W, BODY_BOTTOM,
                    top_pad=g["top_gap"])
    n_cols = len(m.columns)                    # 3~4 栏：栏数由文档里的对比项决定，不裁不删
    cw = (fr["w"] - (n_cols - 1) * gap) / n_cols
    ch = fr["bottom"] - fr["top"]
    mb_h, mb_gap = g["mb_h"], g["mb_gap"]      # 底部比喻条
    desc_w, desc_avail, _used = compare_desc_box(
        m.summary, max((len(c.points[:4]) for c in m.columns), default=0), False,
        n_cols=n_cols)
    # 从大到小挑字号档：三栏都放得下就用大档，放不下自己降（最低 10pt）；
    # 口径走 `card_text_h`（描述里的小节标题会另起一行）
    size_pt, spacing = pick_step(
        lambda sz, sp: max(card_text_h(c.desc, desc_w, sz, sp) for c in m.columns),
        desc_avail)
    line_h = size_pt / 72 * spacing
    pt_size = max(9.5, min(10.5, size_pt - 1.0))
    pt_h = g["pt_h"]                              # 每条要点占高
    for i, col in enumerate(m.columns):
        x = fr["x"] + i * (cw + gap)
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, fr["top"], cw, ch, accent=key)
        rect(slide, x, fr["top"], cw, 0.075, fill=key)
        cy = fr["top"]
        # 这一栏的唯一图标：画在块头（右上角）。没填 icon 就按本栏文字自动补一个
        col_icon = _resolve_icon(col.icon, col.label, col.tag, col.desc)
        if col_icon:
            _draw_block_icon(slide, col_icon, x + cw - COMPARE_ICON_RIGHT,
                             cy + 0.2, COMPARE_ICON_W, color=key)
        # 栏名是**居中**的，所以两侧都要给右上角图标留位（不然长栏名会压上去）
        text_box(slide, x + 0.25, cy + 0.3, cw - 0.5, 0.45, col.label,
                 size=fit_one_line(col.label,
                                   compare_label_w(cw, bool(col_icon))),
                 color=key, bold=True, align=PP_ALIGN.CENTER)
        # 机制标签（胶囊）
        pw = min(cw - 0.7, max(1.5, len(col.tag) * 0.19 + 0.5))
        rect(slide, x + (cw - pw) / 2, cy + 0.92, pw, 0.36,
             fill=tint(COLORS[key], 0.86), radius=0.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + (cw - pw) / 2, cy + 0.92, pw, 0.36, col.tag,
                 size=FONT_SIZES["card_tag"], color=key, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 正文区：描述框高度按实际字数算。原来是写死的 1.45in —— 短描述会在
        # 「描述」和「要点」之间留一块白，长描述又会顶到要点上。算完把
        # 「描述 + 要点」整块在「胶囊下沿 ~ 比喻条上沿」之间居中，余量上下平分。
        pts = col.points[:4]
        desc_h = card_text_h(col.desc, desc_w, size_pt, spacing)
        body_h = desc_h + (g["pt_gap"] + len(pts) * pt_h if pts else 0.0)
        band_top = cy + g["band_top"]
        band_bottom = cy + ch - mb_h - mb_gap - g["band_pad"]
        ty = band_top + max(0.0, (band_bottom - band_top - body_h) / 2)
        # 描述文字：不再前置图标 —— 这一栏的图标已经在块头，两个图标会打架。
        # `**关键短语**：` 加粗染色、同行跟排，观众扫短语就能抓住每栏重点
        text_box(slide, x + 0.28, ty, desc_w, desc_h,
                 card_paras(col.desc, size_pt, key=key, color="text"),
                 size=size_pt, spacing=spacing, strong_color=key)
        # 特征要点（高密度补充）
        for pi_, pt in enumerate(pts):
            py = ty + desc_h + g["pt_gap"] + pi_ * pt_h
            rect(slide, x + 0.32, py + 0.075, 0.07, 0.07, fill=key)
            text_box(slide, x + 0.5, py, cw - 0.82, pt_h, pt,
                     size=pt_size, color="text_sub")
        # 底部比喻条
        rect(slide, x + 0.2, cy + ch - mb_h - mb_gap, cw - 0.4, mb_h,
             fill=tint(COLORS[key], 0.9), radius=0.12,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + 0.38, cy + ch - mb_h - mb_gap, cw - 0.76, mb_h,
                 [("“" + col.metaphor + "”", {"italic": True})],
                 size=FONT_SIZES["card_body"], color=key, anchor=MSO_ANCHOR.MIDDLE,
                 align=PP_ALIGN.CENTER, spacing=1.2)
    if m.footer_note:
        footer_note(slide, m.footer_note)




def render_numbered_cards(slide, m):
    from .theme import ACCENT_CYCLE
    content_header(slide, m.title, m.summary)
    g = numbered_geom(m)
    draw_aux(slide, getattr(m, "callout", None) or [], g["fr"])
    cw, ch, gap, cols = g["cw"], g["ch"], g["gap"], g["cols"]
    # 描述框 = 卡片余下高度；从大到小挑字号档（放不下自己降，最低 10pt）。
    # 口径走 `card_text_h`：描述里的小节标题（`**小节**：`）会另起一行，必须算进去。
    size_pt, spacing = fit_size(
        [c.desc for c in m.cards], g["desc_w"], g["desc_h"],
        measure=lambda t, sz, sp: card_text_h(t, g["desc_w"], sz, sp))
    for i, c in enumerate(m.cards):
        col, row = i % cols, i // cols
        x = g["x"] + col * (cw + gap)
        cy = g["top"] + row * (ch + gap)
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, cy, cw, ch, accent=key, bar=True, rail_w=0.065)
        # 这张卡片的唯一图标：块头（右上角）。没填就按卡片文字自动补一个。
        # 4 列窄卡不画图标（`has_icon=False`）—— 见 numbered_geom 的说明
        card_icon = _resolve_icon(c.icon, c.name, c.sub, c.desc) if g["has_icon"] else None
        if card_icon:
            _draw_block_icon(slide, card_icon, x + cw - CARD_ICON_RIGHT,
                             cy + 0.11, CARD_ICON_W, color=key)
        text_box(slide, x + NO_X, cy + 0.16, NO_GUTTER - 0.10, 0.5, c.no,
                 size=FONT_SIZES["hero_value"], color=key, bold=True)
        # 标题即重点：标题是**判断句/带数字的结论**，下面压一道色标把它标出来
        # （与卡内要点的视觉语言一致；名词短语标题读不出结论，qc 会点名）。
        # 宽度要**扣掉右上角图标**：3 列卡里长标题会直接压在图标上（v22 P9 第 3/5 张卡），
        # 所以给图标让位，并在放不下时自动缩半档字号（最低 11.5pt）保住一行。
        title_w = block_title_w(cw, NO_GUTTER, bool(card_icon))
        text_box(slide, x + NO_GUTTER, cy + 0.13, title_w, 0.4, c.name,
                 size=fit_one_line(c.name, title_w), color="text", bold=True)
        rect(slide, x + NO_GUTTER, cy + 0.555,
             min(0.52, title_w - 0.06), 0.045, fill=key)
        if getattr(c, "sub", None):
            text_box(slide, x + NO_GUTTER, cy + 0.66, cw - NO_GUTTER - 0.25, 0.30,
                     c.sub, size=FONT_SIZES["card_sub"], color=key)
        # 描述文字：不再前置图标（卡片图标已经在块头）。
        # `**关键短语**：` 只加粗 + 染色、同行跟排（不放大、不另起行，省高度也省溢出）
        text_box(slide, x + 0.3, cy + 0.98, g["desc_w"], g["desc_h"],
                 card_paras(c.desc, size_pt, key=key, color="text_sub"),
                 size=size_pt, spacing=spacing, strong_color=key)
    if getattr(m, "footer_note", None):
        footer_note(slide, m.footer_note)








def render_layered_stack(slide, m):
    content_header(slide, m.title, m.summary)
    n = len(m.layers)
    g = layered_geom(m)
    draw_aux(slide, getattr(m, "callout", None) or [], g["fr"])
    gap, top, bar_w = g["gap"], g["top"], g["bar_w"]
    bar_h, label_w, x0 = g["bar_h"], g["label_w"], g["x0"]
    size_pt, spacing = fit_size([l.desc for l in m.layers], g["desc_w"], g["desc_h"])
    for i, layer in enumerate(m.layers):
        # layers[0] 为最底层：画在最下面
        row_from_bottom = i
        by = top + (n - 1 - row_from_bottom) * (bar_h + gap)
        key = ["primary", "green", "orange", "purple", "cyan", "pink"][i % 6]
        card(slide, x0, by, bar_w, bar_h, accent=key, radius=0.08)
        rect(slide, x0, by, label_w, bar_h, fill=key, radius=0.08,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, x0 + label_w - 0.25, by, 0.25, bar_h, fill=key)
        # 这一层的唯一图标：画在左侧色带上。没填就按本层文字自动补一个
        layer_icon = _resolve_icon(layer.icon, layer.label, layer.desc)
        if layer_icon:
            _draw_block_icon(slide, layer_icon, x0 + 0.16, by + (bar_h - 0.5) / 2,
                             0.5, color="on_dark", stroke_w=1.0)
            text_box(slide, x0 + 0.74, by, label_w - 1.0, bar_h, layer.label,
                     size=FONT_SIZES["layer_label"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        else:
            text_box(slide, x0, by, label_w - 0.1, bar_h, layer.label,
                     size=FONT_SIZES["layer_label"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        # 右端代表技术标签（按中英实际字宽估算，英文按窄字计，避免挤压正文）
        tags_w = _layer_tags_w(layer)
        if tags_w:
            tx = x0 + bar_w - 0.22 - tags_w
            for t, pw in zip(layer.tags[:3], _layer_pill_ws(layer)):
                rect(slide, tx, by + (bar_h - 0.32) / 2, pw, 0.32,
                     fill=tint(COLORS[key], 0.85), radius=0.5,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
                text_box(slide, tx, by + (bar_h - 0.32) / 2, pw, 0.32, t,
                         size=FONT_SIZES["layer_tag"], color=key, bold=True, align=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)
                tx += pw + 0.1
        desc_w = bar_w - label_w - 0.3 - (tags_w + 0.3 if tags_w else 0.15)
        # 描述文字：不再前置图标（本层图标已经在左侧色带上）
        text_box(slide, x0 + label_w + 0.3, by, desc_w, bar_h,
                 layer.desc, size=size_pt, color="text",
                 anchor=MSO_ANCHOR.MIDDLE, spacing=spacing)
    # 左侧向上箭头：上层依赖下层
    arrow_h = n * bar_h + (n - 1) * gap
    arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, In(0.62), In(top),
                                   In(0.42), In(arrow_h))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = _c(tint(COLORS["primary"], 0.55))
    arrow.line.fill.background()
    arrow.shadow.inherit = False
    if m.footer_note:
        footer_note(slide, m.footer_note)




def render_flow(slide, m):
    y = content_header(slide, m.title, m.summary)
    n = len(m.steps)
    # 环形图只在 ≤FLOW_WRAP_AT 步时用：8 个节点绕椭圆会互相压住（环是 4 步的形态）。
    # 步数多就退回两行线性流程，loop_note 写在折返通道上 —— 循环语义还在，只是不画成环。
    if m.loop_note and 3 <= n <= FLOW_WRAP_AT:
        _render_flow_cycle(slide, m, y)
    else:
        _render_flow_linear(slide, m, y)


def _render_flow_cycle(slide, m, y):
    """环形循环图：节点绕椭圆一周，中心枢纽展示 loop_note，节点间顺时针箭头闭环。"""
    from .theme import ACCENT_CYCLE
    n = len(m.steps)
    # 环形流程**高度吃紧**：辅区一律放右栏，环拿到完整高度。
    # （原来是底部横条，环被压成扁椭圆 —— 这就是「原本不错的流程图变难看」的原因。）
    co = getattr(m, "callout", None) or []
    fr = card_frame(m.summary, bool(co), FLOW_CYCLE_MIN_W, BODY_BOTTOM, top_pad=0.0,
                    aux_h=callout_height(co))
    draw_aux(slide, co, fr)
    node_w, node_h = FLOW_NODE_W, FLOW_NODE_H
    cx = fr["x"] + fr["w"] / 2
    Rx = min(3.55, fr["w"] / 2 - node_w / 2 - 0.06)
    Ry = min(1.80, max(0.85, (fr["bottom"] - fr["top"]) / 2 - node_h / 2 - 0.06))
    cy = (fr["top"] + fr["bottom"]) / 2
    # 中心枢纽：loop_note 拆成「标题：正文」
    note = m.loop_note
    if "：" in note:
        head, body = note.split("：", 1)
    elif ":" in note:
        head, body = note.split(":", 1)
    else:
        head, body = "循环", note
    oval_size = 1.7
    rect(slide, cx - oval_size/2, cy - oval_size/2, oval_size, oval_size,
         fill="primary", shape=MSO_SHAPE.OVAL)
    # 根据实际文本内容动态计算两个文本框的高度和垂直位置，避免重叠或过大空白
    box_w = 1.58
    # 枢纽标题尽量压成一行：一行放不下就降半档字号，别断成「…终止条/件」
    head_size = 12.0
    while head_size > 10.0 and _lines_in(head, box_w, head_size) > 1:
        head_size -= 0.5
    head_lines = _lines_in(head, box_w, head_size)
    body_lines = _lines_in(body, box_w, 7.5)
    head_lh = head_size / 72 * 1.3   # 标题行高（inch）
    body_lh = 7.5 / 72 * 1.12  # 正文行高（inch）
    head_h = head_lines * head_lh
    body_h = body_lines * body_lh
    gap = 0.08
    total_h = head_h + gap + body_h
    # 垂直居中于椭圆
    head_top = cy - total_h / 2
    body_top = head_top + head_h + gap
    text_box(slide, cx - box_w/2, head_top, box_w, head_h, head,
             size=head_size, color="on_dark", bold=True, align=PP_ALIGN.CENTER)
    text_box(slide, cx - box_w/2, body_top, box_w, body_h, body,
             size=FONT_SIZES["step_desc_micro"], color="on_dark", align=PP_ALIGN.CENTER,
             spacing=1.12, anchor=MSO_ANCHOR.TOP)
    # 节点中心坐标（顺时针，i=0 在正上方）
    centers = []
    for i in range(n):
        phi = math.radians(i * 360 / n)
        centers.append((cx + Rx * math.sin(phi), cy - Ry * math.cos(phi)))
    # 节点间箭头（含回到起点的闭环箭头）
    for i in range(n):
        x0, y0 = centers[i]
        x1, y1 = centers[(i + 1) % n]
        mx, my = (x0 + x1) / 2, (y0 + y1) / 2
        rot = math.degrees(math.atan2(y1 - y0, x1 - x0))
        ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                    In(mx - 0.11), In(my - 0.08), In(0.22), In(0.16))
        ar.fill.solid()
        ar.fill.fore_color.rgb = _c(tint(COLORS["primary"], 0.2))
        ar.line.fill.background()
        ar.shadow.inherit = False
        ar.rotation = rot
    # 节点卡片（最后画，盖在箭头上方）
    for i, st in enumerate(m.steps):
        x = centers[i][0] - node_w / 2
        yy = centers[i][1] - node_h / 2
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, yy, node_w, node_h, accent=key)
        rect(slide, x + node_w - 0.40, yy + 0.06, 0.28, 0.28, fill=key,
             radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + node_w - 0.40, yy + 0.06, 0.28, 0.28, str(i + 1),
                 size=FONT_SIZES["badge"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 这个节点的唯一图标：块头。没填就按标签+描述自动补一个
        step_icon = _resolve_icon(st.icon, st.label, st.desc)
        if step_icon:
            _draw_block_icon(slide, step_icon, x + node_w / 2 - 0.15, yy + 0.07, 0.30, color=key)
        text_box(slide, x + 0.06, yy + 0.40, node_w - 0.12, 0.24,
                 _strip_step_no(st.label), size=FONT_SIZES["step_label_micro"], color="text", bold=True,
                 align=PP_ALIGN.CENTER, spacing=1.05)
        # 描述文字：不再前置图标（节点图标已经在块头）
        text_box(slide, x + 0.14, yy + 0.64, FLOW_CYCLE_DESC_W, FLOW_CYCLE_DESC_H,
                 st.desc,
                 size=FLOW_CYCLE_DESC_SIZE, color="text_sub", align=PP_ALIGN.CENTER,
                 spacing=FLOW_CYCLE_DESC_SPACING)


# 环形流程的节点描述框：**固定**（节点绕椭圆排，卡宽卡高不随步数变）。
# 7.5pt 是环形的专用字号 —— 它比其他版式的 10pt 下限还小，因为一圈最多 6 个节点、
# 每卡只有 1.72in 宽。schema 的 desc 静态上限 80 字在这里放不下，靠动态容量校验拦。








def _flow_arrow(slide, shape, x, y, w, h):
    ar = slide.shapes.add_shape(shape, In(x), In(y), In(w), In(h))
    ar.fill.solid()
    ar.fill.fore_color.rgb = _c(tint(COLORS["primary"], 0.2))
    ar.line.fill.background()
    ar.shadow.inherit = False
    return ar


def _draw_flow_wrap(slide, g) -> None:
    """两行流程的折返通道：一条细线 + 右端向下箭头（"回到左边继续"）。

    步数多（7~12）时靠它把两行串成一条流水线：第一行走到最右 → 掉头 → 第二行继续。
    """
    fr, per, bw, gap, n = g["fr"], g["per"], g["bw"], g["gap"], g["n"]
    last_row1 = min(per, n) - 1
    x_r = fr["x"] + last_row1 * (bw + gap) + bw - 0.22
    x_l = fr["x"] + 0.12
    y0 = g["by"] + g["bh"]                      # 第一行卡片下沿（折返通道从这里开始）
    # 有 loop_note 时把它写在折返通道的上半段：折返本身就是"回到起点"的循环语义。
    # 注意**文字与细线要分层**（同一层会穿字），通道高 FLOW_ROW_GAP 装得下一行小字 + 细线。
    note_h = 0.17 if g.get("loop_note") else 0.0
    if note_h:
        text_box(slide, x_l, y0 + 0.01, max(1.0, x_r - x_l - 0.10), note_h,
                 g["loop_note"], size=FONT_SIZES["micro_label"], color="text_sub")
    line_y = y0 + (note_h + 0.06 if note_h else 0.13)
    rect(slide, x_l, line_y, max(0.5, x_r + 0.10 - x_l), 0.022,
         fill=tint(COLORS["primary"], 0.30))
    _flow_arrow(slide, MSO_SHAPE.DOWN_ARROW, x_r, line_y - 0.20, 0.24, 0.22)


def _render_flow_linear(slide, m, y):
    """线性流程：编号徽章 + 图标 + 强连接箭头；步数 > FLOW_WRAP_AT 时折成两行。"""
    from .theme import ACCENT_CYCLE
    g = flow_linear_geom(m)
    g["loop_note"] = getattr(m, "loop_note", None) if g["rows"] == 2 else None
    draw_aux(slide, getattr(m, "callout", None) or [], g["fr"])
    n, gap, arrow_w = g["n"], g["gap"], g["arrow_w"]
    bw, bh, by, per = g["bw"], g["bh"], g["by"], g["per"]
    # 步骤描述放在「编号 + 图标 + 标签」之下，从大到小挑档（最低 10pt）
    _desc_w = g["desc_w"]
    size_pt, spacing = fit_size([st.desc for st in m.steps], _desc_w, g["desc_h"])
    for i, st in enumerate(m.steps):
        row, col = divmod(i, per)
        x = g["fr"]["x"] + col * (bw + gap)
        yy = by + row * (bh + g["row_gap"])
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, yy, bw, bh, accent=key)
        rect(slide, x, yy, bw, 0.07, fill=key)
        rect(slide, x + 0.12, yy + 0.14, 0.32, 0.32, fill=key, radius=0.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + 0.12, yy + 0.14, 0.32, 0.32, str(i + 1),
                 size=FONT_SIZES["badge"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 这个步骤的唯一图标：块头。没填就按标签+描述自动补一个
        step_icon = _resolve_icon(st.icon, st.label, st.desc)
        if step_icon:
            _draw_block_icon(slide, step_icon, x + bw / 2 - 0.24, yy + 0.20, 0.48, color=key)
        text_box(slide, x + 0.15, yy + 0.78, bw - 0.3, 0.4, _strip_step_no(st.label),
                 size=FONT_SIZES["step_label"], color=key, bold=True,
                 align=PP_ALIGN.CENTER, spacing=1.15)
        # 描述文字：不再前置图标（步骤图标已经在块头）；
        # 框内垂直居中，卡片被拉高时不会把文字全挤在顶部、下面留一条空
        text_box(slide, x + 0.2, yy + 1.24, _desc_w, bh - 1.36, st.desc,
                 size=size_pt, color="text_sub",
                 align=PP_ALIGN.CENTER, spacing=spacing, anchor=MSO_ANCHOR.MIDDLE)
        # 行内箭头：本行还有下一步才画（行末交给折返通道）
        if col < per - 1 and i < n - 1:
            _flow_arrow(slide, MSO_SHAPE.RIGHT_ARROW, x + bw + 0.08,
                        yy + bh / 2 - 0.16, arrow_w, 0.32)
    if g["rows"] == 2:
        _draw_flow_wrap(slide, g)


def render_hierarchy(slide, m):
    y = content_header(slide, m.title, m.summary)
    if m.mode == "progressive":
        _render_hierarchy_progressive(slide, m, y)
    else:
        _render_hierarchy_contain(slide, m, y)






def _render_hierarchy_contain(slide, m, y):
    """包含（树形）：顶部根节点实心条，向下连线分流到 N 个子模块卡片。"""
    from .theme import ACCENT_CYCLE
    g = hierarchy_geom(m)
    fr, n, cw, ch = g["fr"], g["n"], g["cw"], g["ch"]
    draw_aux(slide, getattr(m, "callout", None) or [], fr)
    root_w = min(6.0, max(2.6, len(m.root) * 0.24 + 1.2), fr["w"])
    rx = fr["x"] + (fr["w"] - root_w) / 2
    ry = fr["top"]
    rect(slide, rx, ry, root_w, 0.5, fill="primary",
         radius=0.16, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text_box(slide, rx, ry, root_w, 0.5, m.root, size=FONT_SIZES["hier_root"], color="on_dark",
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    gap = g["gap"]
    # 卡片高度贴着内容（图标 + 标签 + 描述 + 内边距），**不拉满页**：
    # 拉满就会出现 P31 那种"框拉长了、下半截没内容"的空框。
    # 真的空得多，交给 --qc 提示补内容或补占位，而不是靠拉长遮丑。
    cy = g["cy"]
    # 内容确实撑不满时，把整组卡片往下挪一点，空白上下平分（比全堆在底部自然），
    # 但**不把卡片拉长** —— 拉长就是 P31 那种"框在、内容不在"。
    cy += min(0.55, g["slack"] / 2)
    bus_y = ry + 0.5 + 0.28
    line_c = tint(COLORS["primary"], 0.45)
    rect(slide, fr["x"] + fr["w"] / 2 - 0.015, ry + 0.5, 0.03, 0.28, fill=line_c)
    x0 = fr["x"] + cw / 2
    x1 = fr["x"] + (n - 1) * (cw + gap) + cw / 2
    rect(slide, x0, bus_y, x1 - x0, 0.03, fill=line_c)
    for i, node in enumerate(m.nodes):
        cx_ = fr["x"] + i * (cw + gap) + cw / 2
        x = fr["x"] + i * (cw + gap)
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        rect(slide, cx_ - 0.015, bus_y, 0.03, cy - bus_y, fill=line_c)
        card(slide, x, cy, cw, ch, accent=key)
        rect(slide, x, cy, cw, 0.06, fill=key)
        # 这个子模块的唯一图标：块头。没填就按标签+描述自动补一个
        node_icon = _resolve_icon(node.icon, node.label, node.desc)
        if node_icon:
            _draw_block_icon(slide, node_icon, x + cw / 2 - 0.22, cy + 0.14, 0.44, color=key)
        text_box(slide, x + 0.1, cy + 0.62, cw - 0.2, 0.28, node.label,
                 size=FONT_SIZES["badge"], color="text", bold=True, align=PP_ALIGN.CENTER)
        # 描述文字：不再前置图标（子模块图标已经在块头）；字号档由 hierarchy_metrics 挑好。
        # 卡片吃了留白会变高，描述**垂直居中**，不会全堆在顶部留出下半截空白。
        text_box(slide, x + 0.14, cy + 0.92, g["desc_w"], g["desc_h"], node.desc,
                 size=g["size"], color="text_sub", align=PP_ALIGN.CENTER,
                 spacing=g["spacing"], anchor=MSO_ANCHOR.MIDDLE)
    if m.footer_note:
        footer_note(slide, m.footer_note)


def _render_hierarchy_progressive(slide, m, y):
    """递进（阶梯）：子模块左低右高逐级上升，编号徽章 + 上升箭头。"""
    from .theme import ACCENT_CYCLE
    n = len(m.nodes)
    gap = 0.42
    co = getattr(m, "callout", None) or []
    # 递进阶梯同「含型」：纵向靠台阶本身占位，辅区走底部横条
    fr = card_frame(m.summary, bool(co), HIER_MIN_W, BODY_BOTTOM, top_pad=0.0,
                    aux_h=callout_height(co))
    draw_aux(slide, co, fr)
    bw = (fr["w"] - (n - 1) * gap) / n
    step = 0.24
    # 卡片高度撑满可用空间：最低那级（i=0）的底边压到内容区底线，别在下方留白
    bh = min(3.6, fr["bottom"] - fr["top"] - 0.5 - (n - 1) * step)
    _size, _sp = fit_size([nd.desc for nd in m.nodes], bw - 0.28, bh - 1.08, HIER_STEPS)
    for i, node in enumerate(m.nodes):
        x = fr["x"] + i * (bw + gap)
        top = fr["top"] + 0.5 + (n - 1 - i) * step
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, top, bw, bh, accent=key)
        rect(slide, x, top, bw, 0.06, fill=key)
        rect(slide, x + 0.12, top + 0.12, 0.32, 0.32, fill=key, radius=0.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + 0.12, top + 0.12, 0.32, 0.32, str(i + 1),
                 size=FONT_SIZES["badge"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 这个子模块的唯一图标：块头。没填就按标签+描述自动补一个
        node_icon = _resolve_icon(node.icon, node.label, node.desc)
        if node_icon:
            _draw_block_icon(slide, node_icon, x + bw / 2 - 0.2, top + 0.16, 0.4, color=key)
        text_box(slide, x + 0.1, top + 0.62, bw - 0.2, 0.28, node.label,
                 size=FONT_SIZES["card_sub"], color="text", bold=True, align=PP_ALIGN.CENTER)
        # 描述文字：不再前置图标（子模块图标已经在块头）；
        # 描述框占满卡片余下高度并垂直居中，短描述不会把空白全堆在卡片底部
        text_box(slide, x + 0.14, top + 0.94, bw - 0.28, bh - 1.08, node.desc,
                 size=_size, color="text_sub", align=PP_ALIGN.CENTER, spacing=_sp,
                 anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                        In(x + bw + 0.05), In(top + 0.55),
                                        In(0.34), In(0.22))
            ar.fill.solid()
            ar.fill.fore_color.rgb = _c(tint(COLORS["primary"], 0.2))
            ar.line.fill.background()
            ar.shadow.inherit = False
            ar.rotation = -18
    if m.footer_note:
        footer_note(slide, m.footer_note)


PANEL_PT_SIZE = 11.0     # 模块内要点正文字号
















def _panel_points(slide, x, pw, g: dict, pts, key) -> float:
    """模块内要点列表：排版参数全部取自 `panel_content_geom`（**绝不会越过 pt_bottom**）。"""
    if not pts:
        return g["cy"]
    size, spacing = g["pt_size"], g["pt_spacing"]
    hs, gap = g["pt_hs"], g["pt_gap"]
    cy = g["cy"]
    for t, h in zip(pts, hs):
        rect(slide, x + PANEL_PAD + 0.04, cy + h / 2 - 0.038, 0.075, 0.075, fill=key)
        text_box(slide, x + PANEL_PAD + 0.22, cy, g["iw"], h, t, size=size,
                 color="text", spacing=spacing)
        cy += h + gap
    return cy - gap


def _panel_table(slide, x, pw, g: dict, headers, rows, key):
    """模块内迷你明细表：位置 / 行高 / 字号全部取自 `panel_content_geom`。

    关键约定：`head_h + row_h×n ≡ g["tbl_h"]`，表**底部锚定**在内容区下沿 ——
    这样表格在数学上不可能越出模块框（旧的 `max(0.18, …)` + 居中写法会）。
    """
    from pptx.util import Inches as _In
    x0 = x + PANEL_PAD
    tw = pw - 2 * PANEL_PAD
    n = len(rows)
    head_h = g["head_h"]
    row_h = g["row_h"]
    cell_size = g["cell_size"]
    total = head_h + row_h * n
    ty = g["tbl_y"]
    gf = slide.shapes.add_table(n + 1, len(headers), _In(x0), _In(ty),
                                _In(tw), _In(total))
    tbl = gf.table
    tbl.horz_banding = False
    tbl.first_row = False
    tbl.rows[0].height = _In(head_h)
    for r in range(1, n + 1):
        tbl.rows[r].height = _In(row_h)
    for c, cw_ in enumerate(col_widths(headers, rows, tw, min_w=0.62)):
        tbl.columns[c].width = _In(cw_)
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _c(key)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = _In(0.06)
        cell.margin_top = cell.margin_bottom = _In(0.01)
        _fill_cell(cell, h, cell_size, "FFFFFF", bold=True, align=PP_ALIGN.CENTER)
    for r, row in enumerate(rows, start=1):
        band = r % 2 == 1
        for cidx, val in enumerate(row):
            cell = tbl.cell(r, cidx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _c("FFFFFF" if band else "F7F8FA")
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = _In(0.06)
            cell.margin_top = cell.margin_bottom = _In(0.01)
            _fill_cell(cell, val, cell_size,
                       "text" if cidx == 0 else "text_sub",
                       bold=cidx == 0, align=PP_ALIGN.LEFT if cidx == 0 else PP_ALIGN.CENTER)
    return ty + total


def render_panel_2col(slide, m):
    """双模块并排页：页面横向一分为二，每块一个背景框、顶部一段总结陈述段。

    两块内容上下堆叠会在下半页留出半页空白 —— 这里一律左右分置：
    左模块 + 右模块；每个模块 = 色条 + 图标/标题 + 浅色总结块 + 要点/迷你明细表。
    """
    from .theme import ACCENT_CYCLE
    y = content_header(slide, m.title, m.summary)
    co = getattr(m, "callout", None) or []
    # 双模块本身就是左右分置；再让出右栏会把两个迷你表挤瘪，所以辅区走底部横条
    fr = card_frame(m.summary, bool(co), PANEL_MIN_W, BODY_BOTTOM,
                    top_pad=PANEL_TOP_PAD, aux_h=callout_height(co))
    draw_aux(slide, co, fr)
    top, bottom = fr["top"], fr["bottom"] - 0.06
    ph = bottom - top
    pw = (fr["w"] - PANEL_GAP) / 2
    for i, p in enumerate(m.panels):
        x = fr["x"] + i * (pw + PANEL_GAP)
        key = ["primary", "green"][i % 2] if i < 2 else ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        card(slide, x, top, pw, ph, accent=key)
        rect(slide, x, top, pw, 0.075, fill=key)
        # ---- 块头：一个模块只画一个图标 ----
        icon = _resolve_icon(p.icon, p.label, p.summary)
        tx = x + PANEL_PAD
        if icon and _draw_icon_checked(slide, icon, x + PANEL_PAD, top + PANEL_HEAD_TOP,
                                       0.40, color=key):
            tx = x + PANEL_PAD + 0.52
        text_box(slide, tx, top + PANEL_HEAD_TOP, x + pw - PANEL_PAD - tx,
                 PANEL_HEAD_H, p.label, size=FONT_SIZES["card_title"], color=key,
                 bold=True, anchor=MSO_ANCHOR.MIDDLE)
        # ---- 模块顶部总结陈述段（浅色底块，与页面级结论段同一语言）----
        g = panel_content_geom(p.summary, top, ph, pw, p)
        sy, sh = g["sy"], g["sum_h"]
        sw = pw - 2 * PANEL_PAD
        rect(slide, x + PANEL_PAD, sy, sw, sh, fill=tint(COLORS[key], 0.92),
             radius=0.10, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + PANEL_PAD + 0.14, sy, sw - 0.28, sh, p.summary,
                 size=PANEL_SUM_SIZE, color="text", spacing=1.32,
                 anchor=MSO_ANCHOR.MIDDLE)
        # ---- 模块内容：要点在上、迷你明细表在下 ----
        # 空间切分只认 panel_content_geom（渲染与 qc 同一口径）：
        # 表先按内容要高度并**底部锚定**，要点只在上方剩余区间里排 —— 两者都不可能越框。
        if p.points:
            _panel_points(slide, x, pw, g, p.points[:6], key)
        if p.rows:
            _panel_table(slide, x, pw, g, p.headers, p.rows, key)
    if m.footer_note:
        footer_note(slide, m.footer_note)


# ---------------- 2x2 网格（四象限/矩阵） ----------------
#
# 与 numbered_cards 4 张卡的区别：numbered_cards 由卡数决定排列（1×4/2×2/4×1），
# grid_2x2 强制 2×2 —— 适用四象限分析、SWOT、重要-紧急矩阵等。
# 视觉层级：格子标题 13pt → 核心指标 18pt（大字突出）→ 说明 11pt。





def render_grid_2x2(slide, m):
    """2x2 网格：强制 2×2 排列的四象限/矩阵布局。"""
    from .theme import ACCENT_CYCLE

    y = content_header(slide, m.title, m.summary)
    g = grid_geom(m)
    cell_w, cell_h = g["cell_w"], g["cell_h"]
    avail_x, avail_top = MARGIN_L, g["top"]

    # 4 个格子
    cells = [
        (m.q_top_left, avail_x, avail_top),
        (m.q_top_right, avail_x + cell_w + GRID_GAP, avail_top),
        (m.q_bottom_left, avail_x, avail_top + cell_h + GRID_GAP),
        (m.q_bottom_right, avail_x + cell_w + GRID_GAP, avail_top + cell_h + GRID_GAP),
    ]

    for idx, (cell, cx, cy) in enumerate(cells):
        key = ACCENT_CYCLE[idx % len(ACCENT_CYCLE)]

        # 卡片背景
        card(slide, cx, cy, cell_w, cell_h, accent=key)

        # 顶部色条
        rect(slide, cx, cy, cell_w, 0.06, fill=key)

        # 图标
        icon = _resolve_icon(cell.icon, cell.title, cell.desc)
        icon_x = cx + cell_w - CARD_ICON_RIGHT
        has_icon = icon and _draw_block_icon(slide, icon, icon_x, cy + 0.10, CARD_ICON_W, color=key)

        # 标题（13pt，比正文 11.5pt 大一级）
        title_w = cell_w - 0.4
        if has_icon:
            title_w -= 0.5
        text_box(slide, cx + 0.25, cy + 0.10, title_w, 0.35,
                 cell.title, size=FONT_SIZES["quadrant_title"], color=key, bold=True)

        # 核心指标值（18pt，大字突出，比正文大 ≥2 级：11.5→18）
        text_box(slide, cx + 0.25, cy + 0.55, cell_w - 0.5, 0.5,
                 cell.value, size=FONT_SIZES["quadrant_value"], color="text", bold=True)

        # 说明（整页统一一档字号，放不下才降）
        if cell.desc:
            text_box(slide, cx + 0.25, cy + GRID_DESC_TOP, g["desc_w"], g["desc_h"],
                     cell.desc, size=g["desc_size"], color="text_sub",
                     spacing=g["desc_spacing"])

    if m.footer_note:
        footer_note(slide, m.footer_note)


# ---------------- 要点分组（key_groups） ----------------
#
# 关系语义落点：「平行分组」—— 2-4 个主题组，每组 = 组标题 + 要点清单。
# 补的是路由缺口：3 条论据过去被硬塞进 compare_3col（对比语义 + 造比喻），
# 成组但不对比的内容只能摊回 statement 兜底变成字堆。组与组两列并排，
# 组数 3 时第三组占整行，4 组排 2×2。







def render_key_groups(slide, m):
    content_header(slide, m.title, m.summary)
    g = key_groups_geom(m)
    for cg in g["cards"]:
        grp = m.groups[cg["i"]]
        cx, cy, c_w, c_h = cg["x"], cg["y"], cg["w"], cg["h"]
        key = ACCENT_CYCLE[cg["i"] % len(ACCENT_CYCLE)]
        card(slide, cx, cy, c_w, c_h, accent=key)
        rect(slide, cx, cy, c_w, 0.05, fill=key)
        # 组头：图标（右上）+ 组标题（左，加粗染色）
        icon = _resolve_icon(grp.icon, grp.title, " ".join(grp.points))
        title_w = c_w - 0.32
        if icon and _draw_block_icon(slide, icon,
                                     cx + c_w - CARD_ICON_RIGHT,
                                     cy + 0.12, CARD_ICON_W, color=key):
            title_w -= CARD_ICON_W + 0.14
        text_box(slide, cx + 0.16, cy + 0.12, title_w, 0.30, grp.title,
                 size=FONT_SIZES["quadrant_title"], color=key, bold=True)
        # 要点清单：小色点 + 文本，整组一起挑一档字号（高度贴着内容走）
        py = cy + _KG_TITLE_ROW
        pw = cg["text_w"]
        size, spacing = fit_size(grp.points, pw, cg["pts_h"] - 0.05 * max(0, len(grp.points) - 1))
        # 卡片会被撑高到接近可用高度（见 key_groups_geom），余量分给**条间距**，
        # 让要点在卡里均匀铺开 —— 全挤在卡顶、下半截空着，卡片撑多高都没用。
        need = sum(_para_h(pt, pw, size, spacing) for pt in grp.points)
        gap_pt = 0.05 if len(grp.points) < 2 else min(
            0.20, max(0.05, (cg["pts_h"] - need) / (len(grp.points) - 1)))
        line_h = size / 72 * spacing
        for pt in grp.points:
            rect(slide, cx + 0.18, py + line_h / 2 - 0.028, 0.056, 0.056,
                 fill=key, shape=MSO_SHAPE.OVAL)
            text_box(slide, cx + 0.30, py, pw, line_h + 0.04, pt,
                     size=size, color="text", spacing=spacing)
            py += _para_h(pt, pw, size, spacing) + gap_pt
    if m.callout:
        _render_callout(slide, m.callout, BODY_BOTTOM)
    if m.footer_note:
        footer_note(slide, m.footer_note)


# ---------------- 时间轴/里程碑（timeline） ----------------
#
# 关系语义落点：跨模块抽取出的时序链（sequence），且模块名带阶段/期数/时间语义。
# 与 flow 的分工：flow 讲「步骤怎么走」（判定/分支/循环），timeline 讲
# 「何时发生什么」（时间点 + 事件）。节点沿中轴上下交替排布，卡宽随节点数均分。





def _render_timeline_vertical(slide, m, g):
    """竖排时间轴：左侧时间列 + 竖直中轴 + 右侧「节点名 + 说明」卡片。

    行高由几何函数统一算（12 行约 0.36in/行），文字在行内**底部对齐**，
    时间与卡片同一行对齐，读起来是一条真正的时序链，而不是被压扁的一排小卡。
    """
    axis_x, card_x, card_w = g["axis_x"], g["card_x"], g["card_w"]
    rect(slide, axis_x - 0.015, g["top"], 0.03, g["bottom"] - g["top"], fill="primary")
    for i, ms in enumerate(m.milestones):
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        cy = g["top"] + i * (g["row_h"] + g["gap"])
        mid = cy + g["row_h"] / 2
        # 左侧时间（没有 time 就空着，不拿 label 顶）
        if getattr(ms, "time", None):
            text_box(slide, MARGIN_L, cy, g["time_w"] - 0.10, g["row_h"],
                     ms.time, size=FONT_SIZES["micro_label"], color=key, bold=True,
                     align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, axis_x - 0.055, mid - 0.055, 0.11, 0.11, fill=key,
             shape=MSO_SHAPE.OVAL)
        card(slide, card_x, cy, card_w, g["row_h"], accent=key,
             radius=0.06, bar=True)
        text_box(slide, card_x + 0.16, cy + 0.04, card_w - 0.30,
                 g["row_h"] - 0.08,
                 [(ms.label, {"bold": True, "color": key, "size": 11}),
                  (ms.desc, {"color": "text_sub", "size": g["size"]})],
                 size=g["size"], spacing=g["spacing"])


def render_timeline(slide, m):
    content_header(slide, m.title, m.summary)
    g = timeline_geom(m)
    if g["mode"] == "vertical":
        _render_timeline_vertical(slide, m, g)
        if m.callout:
            _render_callout(slide, m.callout, BODY_BOTTOM)
        if m.footer_note:
            footer_note(slide, m.footer_note)
        return
    axis_y = g["axis_y"]
    rect(slide, MARGIN_L, axis_y - 0.018, CONTENT_W, 0.036, fill="primary")
    for i, ms in enumerate(m.milestones):
        above = (i % 2 == 0)
        cx = MARGIN_L + i * (g["card_w"] + g["gap"])
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        cy = (axis_y - TL_AXIS_HALF - g["half_h"]) if above \
            else (axis_y + TL_AXIS_HALF)
        card(slide, cx, cy, g["card_w"], g["half_h"], accent=key)
        rect(slide, cx, cy, g["card_w"], 0.05, fill=key)
        mid = cx + g["card_w"] / 2
        # 连接线（卡片边 → 中轴）+ 节点圆点
        conn_y = (cy + g["half_h"]) if above else cy
        rect(slide, mid - 0.011, conn_y, 0.022, TL_AXIS_HALF, fill=key)
        rect(slide, mid - 0.055, axis_y - 0.055, 0.11, 0.11, fill=key,
             shape=MSO_SHAPE.OVAL)
        # 卡内：time（小字强调）→ label（加粗）→ desc（逐档降字号塞满）
        ty = cy + 0.12
        if getattr(ms, "time", None):
            text_box(slide, cx + 0.18, ty, g["card_w"] - 0.36, 0.24,
                     ms.time, size=FONT_SIZES["micro_label"], color=key, bold=True)
            ty += 0.26
        text_box(slide, cx + 0.18, ty, g["card_w"] - 0.36, 0.32,
                 ms.label, size=FONT_SIZES["quadrant_title"], color="text",
                 bold=True)
        dy = ty + 0.40
        dh = cy + g["half_h"] - dy - 0.12
        size, spacing = fit_size([ms.desc], g["desc_w"], dh)
        text_box(slide, cx + 0.18, dy, g["desc_w"], dh, ms.desc,
                 size=size, color="text_sub", spacing=spacing)
    if m.callout:
        _render_callout(slide, m.callout, BODY_BOTTOM)
    if m.footer_note:
        footer_note(slide, m.footer_note)


# ---------------- 转化漏斗（funnel） ----------------
#
# 关系语义落点：转化/衰减链 —— 进线→成交这类「每层都在变小」的流程。
# 层条居中逐层收窄，label+value 在条内，口径/转化率放右侧说明栏（一页两个信息块）。





def render_funnel(slide, m):
    content_header(slide, m.title, m.summary)
    g = funnel_geom(m)
    for i, st in enumerate(m.stages):
        key = ACCENT_CYCLE[i % len(ACCENT_CYCLE)]
        w = g["widths"][i]
        bx = MARGIN_L + (g["zone_w"] - w) / 2
        by = g["top"] + i * (g["bar_h"] + g["gap"])
        rect(slide, bx, by, w, g["bar_h"], fill=tint(COLORS[key], 0.86),
             radius=0.10, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, bx, by, 0.055, g["bar_h"], fill=key)
        # 条内：label + value（value 是这层的量，层窄也要放得下）
        text_box(slide, bx + 0.20, by + 0.06, w - 0.40, 0.30,
                 st.label, size=FONT_SIZES["quadrant_title"], color=key,
                 bold=True)
        if st.value:
            v_size = 18 if g["bar_h"] >= 0.58 else 14
            text_box(slide, bx + 0.20, by + 0.36, w - 0.40, g["bar_h"] - 0.42,
                     st.value, size=v_size, color="text", bold=True)
        # 右侧说明栏：口径 / 转化率 / 流失原因，随条行对齐
        size, spacing = fit_size([st.desc], g["desc_w"], g["desc_h"])
        text_box(slide, g["desc_x"], by, g["desc_w"], g["bar_h"], st.desc,
                 size=size, color="text_sub", spacing=spacing,
                 anchor=MSO_ANCHOR.MIDDLE)
    if m.callout:
        _render_callout(slide, m.callout, BODY_BOTTOM)
    if m.footer_note:
        footer_note(slide, m.footer_note)


# ---------------- 二维矩阵（matrix） ----------------
#
# grid_2x2 的 m×n 泛化：两个维度各自取值、交叉格有独立结论的分类。
# 列头一行 + 左侧行标一列 + 卡片网格；grid_2x2 适合"每格一个核心大数字"，
# matrix 适合"每格一段结论"。





def render_matrix(slide, m):
    content_header(slide, m.title, m.summary)
    g = matrix_geom(m)
    x0 = MARGIN_L + MATRIX_ROWLABEL_W + MATRIX_GAP
    for j, cl in enumerate(m.col_labels):
        cx = x0 + j * (g["cell_w"] + MATRIX_GAP)
        text_box(slide, cx, g["top"], g["cell_w"], MATRIX_COLHEAD_H, cl,
                 size=FONT_SIZES["quadrant_title"], color="primary", bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for r_i, row in enumerate(m.rows):
        ry = (g["top"] + MATRIX_COLHEAD_H + 0.08
              + r_i * (g["cell_h"] + MATRIX_GAP))
        text_box(slide, MARGIN_L, ry, MATRIX_ROWLABEL_W - 0.14, g["cell_h"],
                 row.label, size=FONT_SIZES["quadrant_title"], color="text",
                 bold=True, anchor=MSO_ANCHOR.MIDDLE)
        key = ACCENT_CYCLE[r_i % len(ACCENT_CYCLE)]
        for c_i, cell in enumerate(row.cells):
            cx = x0 + c_i * (g["cell_w"] + MATRIX_GAP)
            card(slide, cx, ry, g["cell_w"], g["cell_h"])
            rect(slide, cx, ry, g["cell_w"], 0.05, fill=key)
            icon = _resolve_icon(cell.icon, cell.title, cell.desc)
            icon_x = cx + g["cell_w"] - CARD_ICON_RIGHT
            has_icon = icon and _draw_block_icon(slide, icon, icon_x,
                                                 ry + 0.09, CARD_ICON_W,
                                                 color=key)
            tw = g["cell_w"] - 0.40
            if has_icon:
                tw -= 0.45
            text_box(slide, cx + 0.20, ry + 0.10, tw, 0.30, cell.title,
                     size=FONT_SIZES["block_subtitle"], color=key, bold=True)
            dy = ry + MATRIX_DESC_TOP
            dh = ry + g["cell_h"] - dy - MATRIX_BOTTOM_PAD
            size, spacing = fit_size([cell.desc], g["desc_w"], dh, MATRIX_STEPS)
            text_box(slide, cx + 0.20, dy, g["desc_w"], dh, cell.desc,
                     size=size, color="text_sub", spacing=spacing)
    if m.footer_note:
        footer_note(slide, m.footer_note)










def _add_vector_picture(slide, path: str, x, y, w, h) -> bool:
    """EMF/WMF **逐字节原样**嵌入：手工建 ImagePart + 关系 + `p:pic`。

    两个前提在 python-pptx（>=1.0）里成立：① `ImagePart` 构造只存字节、不解析；
    ② `_add_pic_from_image_part` 在**宽高都给全**时不会去问 Pillow 要原生尺寸
    （只要有一个为 None 就会走 PIL，EMF 会直接报 UnrecognizedImageError）。
    """
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.parts.image import ImagePart

    ext = os.path.splitext(path)[1].lower().lstrip(".")
    ctype = _CTYPE_VECTOR.get(ext)
    if not ctype:
        return False
    try:
        with open(path, "rb") as f:
            blob = f.read()
        pkg = slide.part.package
        part = ImagePart(pkg.next_image_partname(ext), ctype, pkg, blob,
                         os.path.basename(path))
        rId = slide.part.relate_to(part, RT.IMAGE)
        slide.shapes._add_pic_from_image_part(part, rId, In(x), In(y), In(w), In(h))
        return True
    except Exception:                                   # noqa: BLE001 —— 交给 qc 报缺图
        return False


def add_image(slide, path: str, x: float, y: float, w: float, h: float) -> bool:
    """在 (x, y, w, h) 里**等比 contain** 放图（矢量图原样嵌入），返回是否成功。

    等比 contain + 居中：宁可四周留白，也不把图拉变形、不裁剪（图画错了比留白更糟）。
    """
    ext = os.path.splitext(path)[1].lower()
    size = image_size_in(path)
    if size:
        iw, ih = size
        k = min(w / iw, h / ih)
        dw, dh = iw * k, ih * k
    else:
        dw, dh = w, h
    px, py = x + (w - dw) / 2.0, y + (h - dh) / 2.0
    if ext in FIG_EXT_VECTOR:
        return _add_vector_picture(slide, path, px, py, dw, dh)
    try:
        slide.shapes.add_picture(path, In(px), In(py), In(dw), In(dh))
        return True
    except Exception:                                   # noqa: BLE001 —— 交给 qc 报缺图
        return False












def _render_point_rail(slide, pts, x: float, y: float, w: float, bottom: float) -> float:
    """figure 的右侧要点栏（points 是纯文本，不是 SideNote，所以不复用 `_render_rail`）。"""
    if not pts:
        return y
    need = point_rail_need(pts, w)
    hs, gap, off = _fill_rail(need, bottom - y, gap_min=0.14)
    cy = y + off
    for i, (t, h) in enumerate(zip(pts, hs)):
        key = ACCENT_CYCLE[(i + 1) % len(ACCENT_CYCLE)]
        rect(slide, x, cy, w, h, fill="card_bg", line="border", radius=0.09,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, x, cy, 0.055, h, fill=key)
        rect(slide, x + 0.18, cy + 0.16, 0.30, 0.30, fill=key, radius=0.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, x + 0.18, cy + 0.16, 0.30, 0.30, str(i + 1), size=FONT_SIZES["badge"],
                 color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        text_box(slide, x + 0.60, cy + 0.08, w - 0.76, h - 0.16, t, size=FONT_SIZES["rail_text"],
                 color="text", spacing=1.25, anchor=MSO_ANCHOR.MIDDLE)
        cy += h + gap
    return cy - gap


def render_figure(slide, m):
    """图示页：原图 + 图注（+ 可选要点右栏）。

    图走 `add_image` —— EMF/WMF 原样嵌入、等比 contain 居中。图放不出来时画
    **显式占位**（写明路径与补救办法），绝不交一页白框出去。
    """
    content_header(slide, m.title, m.summary)
    pts = list(getattr(m, "points", None) or [])
    fr, box = figure_area(m.summary, bool(m.caption), bool(pts))
    card(slide, fr["x"], fr["top"], fr["w"], fr["bottom"] - fr["top"])
    path = resolve_image_path(m.image)
    if not add_image(slide, path, box[0], box[1], box[2], box[3]):
        rect(slide, box[0], box[1], box[2], box[3], fill=tint(COLORS["orange"], 0.92),
             radius=0.10, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text_box(slide, box[0], box[1], box[2], box[3],
                 f"缺图：{m.image}\n（路径不存在或格式不受支持；Word 里的图先跑 "
                 f"engine/ingest/docx_in.py 抽到 .pi/work/media/，EMF/WMF 会原样嵌入）",
                 size=FONT_SIZES["caption"], color="orange", align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
    if m.caption:
        text_box(slide, fr["x"] + FIGURE_PAD, fr["bottom"] - FIGURE_CAPTION_H,
                 fr["w"] - 2 * FIGURE_PAD, FIGURE_CAPTION_H, m.caption,
                 size=FONT_SIZES["rail_text"], color="text_sub", align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
    if pts:
        _render_point_rail(slide, pts, fr["rail_x"], fr["rail_top"], fr["rail_w"],
                           fr["rail_bottom"])
    if m.footer_note:
        footer_note(slide, m.footer_note)












def _render_callout(slide, notes, bottom: float, height: float | None = None) -> float:
    """底部横排辅区：2-4 条关键数字 / 边界 / 待确认点。返回新的内容底线。

    卡片型版式原来整页只有"一组卡片"，信息量被容器锁死；callout 让它们多装一块信息。
    **高度按内容自适应**（`callout_height`），说明文字放不下时逐档缩到 7.5pt。
    """
    from .theme import ACCENT_CYCLE
    n = len(notes)
    if not n:
        return bottom
    h = height if height else callout_height(notes)
    gap = 0.20
    cw = (12.23 - (n - 1) * gap) / n
    y = bottom - h
    for i, cn in enumerate(notes):
        x = 0.55 + i * (cw + gap)
        key = ACCENT_CYCLE[(i + 1) % len(ACCENT_CYCLE)]
        rect(slide, x, y, cw, h, fill=tint(COLORS[key], 0.92),
             radius=0.12, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, x, y, 0.05, h, fill=key)
        text_box(slide, x + 0.16, y + 0.06, cw - 0.32, _CALLOUT_LABEL_H, cn.label,
                 size=FONT_SIZES["callout_label"], color=key, bold=True)
        if cn.value:
            text_box(slide, x + 0.16, y + 0.24, cw - 0.32, _CALLOUT_VALUE_H, cn.value,
                     size=FONT_SIZES["callout_value"], color="text", bold=True)
        if cn.text:
            off = 0.55 if cn.value else 0.28
            tw, th = cw - 0.32, h - off - 0.06
            size = _CALLOUT_SIZE
            while size > 7.5 and _lines_in(cn.text, tw, size) * (size / 72 * 1.15) > th:
                size -= 0.5
            text_box(slide, x + 0.16, y + off, tw, th, cn.text,
                     size=size, color="text_sub", spacing=1.15)
    return y - 0.10


def _render_rail(slide, notes, x: float, y: float, w: float, bottom: float) -> float:
    """右侧纵栏辅区：关键数字 / 边界条件 / 待确认点。返回辅区底边。

    一页只有一块主内容 = 信息量上不去；纵栏就是"第二块信息"——
    表格页的 `side`、图表页的洞察、以及卡片型版式的 `callout` 都走这里。
    **左右分置**（不是上下堆叠）：主内容的高度完整留给主内容，
    卡片高度按内容走，不拉长成空框。
    """
    from .theme import ACCENT_CYCLE
    n = len(notes)
    if not n:
        return y
    # 装不下就逐档降字号（正文 10 → 8），宁可字小一点，不许卡片被页脚压住
    n_ = len(notes)
    gap = 0.14
    avail_side = bottom - y
    size_t = 10.0
    need_h: list[float] = []
    for size_t in (10.0, 9.5, 9.0, 8.5, 8.0):
        need_h = []
        for sn in notes:
            h = 0.12 + 0.22                              # 上边距 + label 行
            h += 0.34 if sn.value else 0.0               # 大字
            if sn.text:
                h += (_lines_in(sn.text, w - 0.66, size_t)
                      * (size_t / 72 * 1.25) + 0.10)
            need_h.append(max(0.74, h))
        if sum(need_h) + gap * (n_ - 1) <= avail_side:
            break
    # 高度分配：内容给足后，余量按「内边距 → 间距 → 长高 → 居中」四步消化
    hs, gap, off = _fill_rail(need_h, avail_side, gap_min=gap)
    cy = y + off
    for i, (sn, ch) in enumerate(zip(notes, hs)):
        key = ACCENT_CYCLE[(i + 1) % len(ACCENT_CYCLE)]
        rect(slide, x, cy, w, ch, fill="card_bg", line="border",
             radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, x, cy, 0.055, ch, fill=key)
        icon = _resolve_icon(getattr(sn, "icon", None), sn.label,
                             sn.value or "", sn.text or "")
        tx0 = x + 0.24
        if icon and _draw_icon_checked(slide, icon, x + 0.15, cy + 0.14,
                                       0.28, color=key):
            tx0 = x + 0.55
        tw = w - (tx0 - x) - 0.16
        text_box(slide, tx0, cy + 0.11, tw, 0.22, sn.label,
                 size=FONT_SIZES["rail_label"], color="text_sub", bold=True)
        off = 0.34
        if sn.value:
            text_box(slide, tx0, cy + 0.33, tw, 0.34, sn.value,
                     size=FONT_SIZES["rail_value"], color=key, bold=True)
            off = 0.68
        if sn.text:
            text_box(slide, tx0, cy + off, tw, max(0.24, ch - off - 0.10),
                     sn.text, size=size_t, color="text", spacing=1.22)
        cy += ch + gap
    return cy - gap










def render_table(slide, m):
    y = content_header(slide, m.title, m.summary)
    from pptx.util import Inches as _In
    rows, cols = len(m.rows) + 1, len(m.headers)
    # 有辅区就收窄主表：整页只放一张表是"信息量上不去"的直接原因
    side = list(getattr(m, "side", None) or [])
    g = table_geom(len(m.rows), m.summary, bool(m.note), m.headers, m.rows, bool(side))
    tx, tw = 0.55, g["tw"]
    head_h, body_h, total_h = g["head_h"], g["body_h"], g["total_h"]
    cell_size = g["cell_size"]
    ty = g["top"]
    head_size = FONT_SIZES["table_header"]
    gf = slide.shapes.add_table(rows, cols, _In(tx), _In(ty), _In(tw), _In(total_h))
    tbl = gf.table
    tbl.horz_banding = False
    tbl.first_row = False
    tbl.rows[0].height = _In(head_h)
    for r in range(1, rows):
        tbl.rows[r].height = _In(body_h)
    # 列宽按内容自适应（不再"第一列 2.7in + 其余均分"）
    for c, cw_ in enumerate(g["widths"]):
        tbl.columns[c].width = _In(cw_)
    # 主角列：由 `hero_col` 显式指定（与表头文字完全匹配的那一列整列高亮成色带）。
    # 旧实现自动匹配表头里的 "Pi" / "pi-ai" / "Pi " 前缀 —— 那是某一份 deck 的
    # 产品名，换个行业就永远匹配不上，等于把一次性的审美决定焊进了渲染器。
    hero = None
    _want = (getattr(m, "hero_col", None) or "").strip()
    if _want:
        for _i, _h in enumerate(m.headers):
            if _h.strip() == _want:
                hero = _i
                break
    # 表头
    for i, h in enumerate(m.headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _c("dark_blue" if i == hero else "primary")
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = _In(0.1)
        cell.margin_top = cell.margin_bottom = _In(0.02)
        _fill_cell(cell, h, head_size, "FFFFFF", bold=True, align=PP_ALIGN.CENTER)
    # 表体
    for r, row in enumerate(m.rows, start=1):
        band = r % 2 == 1
        for cidx, val in enumerate(row):
            cell = tbl.cell(r, cidx)
            cell.fill.solid()
            if cidx == hero:
                cell.fill.fore_color.rgb = _c(tint(COLORS["primary"], 0.84))
            elif cidx == 0:
                cell.fill.fore_color.rgb = _c("F2F3F5")
            else:
                cell.fill.fore_color.rgb = _c("FFFFFF" if band else "F7F8FA")
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = _In(0.1)
            cell.margin_top = cell.margin_bottom = _In(0.02)
            align = PP_ALIGN.LEFT if cidx == 0 else PP_ALIGN.CENTER
            bold = cidx in (0, hero)
            color = "dark_blue" if cidx == hero else ("text" if cidx == 0 else "text_sub")
            _fill_cell(cell, val, cell_size, color, bold=bold, align=align)
    if side:
        # note 已收窄到主表宽度，辅区可以用满到内容区底线
        _render_rail(slide, side, tx + tw + SIDE_GAP, g["top"], SIDE_W, TABLE_BOTTOM)
    if m.note:
        # 有辅区时 note 收窄到主表宽度（右边让给辅区），否则通栏。
        # 位置只由**表格底边**决定：`total_h ≤ avail` 是几何保证，
        # 所以 note 一定落在表格下方、且不会越过 BODY_BOTTOM（旧代码取
        # max(表格底边, 辅区底边) 再 clamp，正好把 note 压到表格上）。
        nw = tw if side else 12.23
        ny = min(g["note_y"], BODY_BOTTOM - g["note_h"])
        note_size = 11.5
        while note_size > 9.0 and _lines_in(m.note, nw - 0.62, note_size) > 1:
            note_size -= 0.5
        rect(slide, tx, ny, nw, g["note_h"], fill=tint(COLORS["primary"], 0.88),
             radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, tx + 0.17, ny + 0.09, 0.07, g["note_h"] - 0.18, fill="primary")
        text_box(slide, tx + 0.4, ny, nw - 0.62, g["note_h"], m.note, size=note_size,
                 color="dark_blue", bold=True, anchor=MSO_ANCHOR.MIDDLE)


def _fill_cell(cell, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _c(color)
    apply_run_font(run)






def render_end(slide, m):
    page_bg(slide)
    # 与封面同一组装饰条：整页内容居中，条也居中（原来靠左，落单了）
    bar_w, bar_gap = 0.36, 0.14
    total = 3 * bar_w + 2 * bar_gap
    for i, key in enumerate(["primary", "green", "orange"]):
        rect(slide, 6.667 - total / 2 + i * (bar_w + bar_gap), END_BAR_Y,
             bar_w, 0.075, fill=key)
    text_box(slide, 0.55, END_TITLE_Y, 12.23, 1.0, m.title, size=FONT_SIZES["end_title"],
             color="text", bold=True, align=PP_ALIGN.CENTER, spacing=1.15)
    g = end_geom(list(m.sections or []), bool(m.contact))
    y = g["top"]
    for s, bh in zip(m.sections or [], g["heights"]):
        card(slide, END_CARD_X, y, g["card_w"], bh)
        rect(slide, END_CARD_X + 0.15, y + 0.10, 0.07, max(0.2, bh - 0.20),
             fill="primary")
        text_box(slide, END_CARD_X + 0.45, y + 0.07, g["card_w"] - 0.62,
                 bh - 0.14,
                 [(s.heading, {"bold": True, "color": "text", "size": 12.5}),
                  (s.text, {"color": "text_sub", "size": g["size"]})],
                 size=g["size"], spacing=g["spacing"])
        y += bh + g["gap"]
    if m.contact:
        text_box(slide, 0.55, END_CONTACT_Y, 12.23, 0.4, m.contact, size=FONT_SIZES["cover_org"],
                 color="text_light", align=PP_ALIGN.CENTER)


# 图表页横向几何 + 竖向几何（渲染与 qc 共用）
CHART_X, CHART_W = 0.55, 7.5






_CHART_TYPE_MAP = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "radar": XL_CHART_TYPE.RADAR_MARKERS,
}



def _apply_chart_fonts(root) -> None:
    """图表里**每一级**文字属性都成对写上 latin/ea/cs。

    python-pptx 不会写 `a:ea`（甚至只写 `a:latin` 或不写），中文标签就会退回主题字体
    （模板自带 Calibri）。而图表的 `c:txPr` 是**逐元素**生效的：图表级 / 图例级 /
    坐标轴标签级 / 数据标签级各有一份，只改图表级会漏掉后面几级 —— 所以整棵树扫一遍。
    """
    for tag in ("a:defRPr", "a:rPr"):
        for rPr in root.iter(qn(tag)):
            _set_font_on_rPr(rPr)
    # 图表 txPr 的 a:p 也缺 endParaRPr：选中图表标题/图例文字时字体框读它
    _mark_language_on(root)


def _normalize_chart_xml(chart):
    """把 python-pptx 生成的图表 XML 修成 PowerPoint 一定认的形状。

    python-pptx 会写出两处 OOXML 违规，PowerPoint 打开时直接判定「文件损坏 /
    无法打开」，点「修复」后图表被丢弃 —— 整页看起来是空的：

    1. `c:doughnutChart` 里出现 `c:dLblPos`。ECMA-376 中环形图不允许数据标签定位
       （`CT_DLbls` 的 `dLblPos` 只对饼/柱/条/折线/雷达有效）。只要写了，
       整个 .pptx 就打不开。
    2. `c:legendPos` 缺 `val`。python-pptx 的 `OptionalAttribute` 在值等于默认值
       （RIGHT）时会把属性删掉，留下 `<c:legendPos/>`，而 `val` 是必填属性。

    这里做一次兜底：不管上层怎么设置，出去的文件都必须能被 PowerPoint 打开。

    顺带把图表字体补齐（`_apply_chart_fonts`）—— 图表不走 `text_box`，
    它有自己的 `c:txPr` 体系，字体会漏。
    """
    root = chart._chartSpace
    _apply_chart_fonts(root)
    # 1) 环形图禁写 dLblPos
    for dough in root.iter(f"{{{_C_NS}}}doughnutChart"):
        for dLblPos in list(dough.iter(f"{{{_C_NS}}}dLblPos")):
            dLblPos.getparent().remove(dLblPos)
    # 2) legendPos 必须带 val，默认 right（且必须是 legend 的第一个子元素）
    for legend in root.iter(f"{{{_C_NS}}}legend"):
        pos = legend.find(f"{{{_C_NS}}}legendPos")
        if pos is None:
            pos = etree.Element(f"{{{_C_NS}}}legendPos")
            legend.insert(0, pos)
        if pos.get("val") is None:
            pos.set("val", "r")
    return chart


def render_chart_native(slide, m):
    """原生 PowerPoint 图表（可在 PPT 内编辑，样式较朴素）。--chart-mode native 时使用。"""
    y = content_header(slide, m.title, m.summary)
    spec = m.chart
    cx, cw = CHART_X, CHART_W
    # 图表高度**按数据点紧凑自适应**，不再无脑撑满：
    # 4 个数据点占 4.7in 只会得到"4 根细条 + 大片空白"。
    # 省下来的高度给下方明细表 / 右侧洞察卡，让一页装两块信息。
    n_pts = len(spec.categories)
    has_detail = bool(m.detail_headers and m.detail_rows)
    _g = chart_metrics(m.summary, spec.type, n_pts,
                       len(m.detail_rows or []), bool(m.footer_note))
    cy, ch_ = _g["cy"], _g["ch"]

    cd = CategoryChartData()
    cd.categories = spec.categories
    unit_suffix = f"（{spec.unit}）" if spec.unit else ""
    for s in spec.series:
        cd.add_series(s.name + unit_suffix, s.values)
    gf = slide.shapes.add_chart(_CHART_TYPE_MAP[spec.type],
                                In(cx), In(cy), In(cw), In(ch_), cd)
    chart = gf.chart
    chart.has_title = False
    chart.font.size = Pt(10.5)
    chart.font.color.rgb = _c("text")
    chart.font.name = FONT_LATIN          # 西文/数字；中文由 _apply_chart_fonts 补 a:ea

    multi = len(spec.series) > 1
    single = not multi
    if multi:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
    else:
        chart.has_legend = False

    plot = chart.plots[0]
    if single and spec.type in ("column", "bar"):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(10)
        dl.font.bold = True
        dl.font.color.rgb = _c("text")
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    elif single and spec.type in ("pie", "doughnut"):
        # 饼/环图：图例显式标注分类，数据标签只显示百分比。
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_category_name = False
        dl.show_percentage = True
        dl.show_value = False
        dl.font.size = Pt(10)
        dl.font.bold = True
        if spec.type == "pie":
            # 饼图可以把标签摆到扇区外，用深色字（底是浅色页面）。
            dl.font.color.rgb = _c("text")
            try:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass
        else:
            # 环图不能设 dLblPos（写了 PowerPoint 直接打不开文件），标签只能落在
            # 环形色带上；色带是 primary/green/orange 等深色，所以标签用白色粗体。
            dl.font.color.rgb = _c("FFFFFF")

    for i, ser in enumerate(chart.series):
        color = CHART_SERIES[i % len(CHART_SERIES)]
        if spec.type in ("pie", "doughnut"):
            for j, point in enumerate(ser.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = _c(
                    ACCENT_CYCLE[j % len(ACCENT_CYCLE)])
        elif single and spec.type in ("column", "bar"):
            for j, point in enumerate(ser.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = _c(
                    ACCENT_CYCLE[j % len(ACCENT_CYCLE)])
        elif spec.type == "line":
            ser.format.line.color.rgb = _c(color)
            ser.format.line.width = Pt(2.25)
            ser.smooth = False
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = _c(color)

    if spec.type in ("column", "bar", "line", "radar"):
        try:
            va = chart.value_axis
            va.has_major_gridlines = True
            va.major_gridlines.format.line.color.rgb = _c("border")
            va.major_gridlines.format.line.width = Pt(0.5)
            va.tick_labels.font.size = Pt(9.5)
            va.tick_labels.font.color.rgb = _c("text_sub")
            ca = chart.category_axis
            ca.tick_labels.font.size = Pt(10.5)
            ca.tick_labels.font.color.rgb = _c("text")
            ca.format.line.color.rgb = _c("border")
        except Exception:
            pass
    try:
        plot.gap_width = 200 if spec.type in ("column", "bar") else 80
    except Exception:
        pass

    # 出口把关：图表 XML 必须能被 PowerPoint 打开（详见 _normalize_chart_xml）
    _normalize_chart_xml(chart)

    # 右侧洞察面板：卡片高度按文字行数走（原来固定 0.92in，两行字也占那么高）
    px, pw = 8.4, 4.38
    text_box(slide, px, cy + 0.02, pw, 0.35, "关键洞察", size=FONT_SIZES["insight_title"],
             color="primary", bold=True)
    rect(slide, px, cy + 0.4, 0.55, 0.045, fill="primary")
    ins_top = cy + 0.60
    ins_bottom = BODY_BOTTOM
    ins_need = [max(0.60, 0.20 + _lines_in(ins.text, pw - 0.94, 10.5)
                    * (10.5 / 72 * 1.25) + 0.16) for ins in m.insights[:4]]
    ins_hs, ins_gap = _stack_heights(ins_need, ins_bottom - ins_top,
                                     gap_min=0.12, gap_max=0.32, pad_max=0.20)
    iy = ins_top
    for i, (ins, ih) in enumerate(zip(m.insights[:4], ins_hs)):
        rect(slide, px, iy, pw, ih, fill="card_bg", line="border",
             radius=0.09, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        if ins.icon:
            _draw_block_icon(slide, ins.icon, px + 0.14, iy + max(0.10, (ih - 0.44) / 2),
                             0.44, color=ACCENT_CYCLE[i % len(ACCENT_CYCLE)])
        else:
            rect(slide, px + 0.18, iy + (ih - 0.36) / 2, 0.36, 0.36,
                 fill=ACCENT_CYCLE[i % len(ACCENT_CYCLE)],
                 radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            text_box(slide, px + 0.18, iy + (ih - 0.36) / 2, 0.36, 0.36,
                     str(i + 1), size=FONT_SIZES["badge"], color="on_dark", bold=True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text_box(slide, px + 0.72, iy + 0.06, pw - 0.94, ih - 0.12, ins.text,
                 size=FONT_SIZES["rail_text"], color="text", spacing=1.25, anchor=MSO_ANCHOR.MIDDLE)
        iy += ih + ins_gap

    # 下方明细表：数据点少时，把图表省下来的空间用来装第二块信息
    if has_detail and _g["detail_y"] is not None:
        dh, drows = list(m.detail_headers), m.detail_rows
        dty = _g["detail_y"]
        d_head, d_row = _g["detail_head"], _g["detail_row"]
        d_total = _g["detail_total"]
        gfd = slide.shapes.add_table(len(drows) + 1, len(dh), In(cx), In(dty),
                                     In(cw), In(d_total))
        t2 = gfd.table
        t2.horz_banding = False
        t2.first_row = False
        t2.rows[0].height = In(d_head)
        for r in range(1, len(drows) + 1):
            t2.rows[r].height = In(d_row)
        for c, ww in enumerate(col_widths(dh, drows, cw, min_w=0.9)):
            t2.columns[c].width = In(ww)
        for i2, h2 in enumerate(dh):
            cell = t2.cell(0, i2)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _c("primary")
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = In(0.08)
            _fill_cell(cell, h2, 10.5, "FFFFFF", bold=True,
                       align=PP_ALIGN.CENTER)
        for r, row in enumerate(drows, start=1):
            band = r % 2 == 1
            for cidx, val in enumerate(row):
                cell = t2.cell(r, cidx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _c("FFFFFF" if band else "F7F8FA")
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = In(0.08)
                _fill_cell(cell, val, 10, "text",
                           align=PP_ALIGN.LEFT if cidx == 0 else PP_ALIGN.CENTER)
    if m.footer_note:
        footer_note(slide, m.footer_note)


def render_chart(slide, m):
    """图表页派发：默认 native 原生可编辑图表，--chart-mode mpl 走 matplotlib 高质感 PNG。"""
    if CHART_MODE == "native":
        render_chart_native(slide, m)
    else:
        render_chart_mpl(slide, m)


def render_chart_mpl(slide, m):
    """图表页：左 matplotlib 高质感图（PNG 嵌图）+ 右关键洞察卡。

    `charts` 在这里才 import：native 是默认图表路径，全程不画 PNG，而 matplotlib
    的导入成本约 1.2s（实测 3.12 / 冷启动），占一次 build 的导入开销近一半。
    放在模块顶层会让「算字数预算」「只质检」这类命令白付这笔钱。
    """
    from . import charts
    y = content_header(slide, m.title, m.summary)
    spec = m.chart
    # 白卡片作图表底，图表 PNG 居中放入；高度撑满可用空间
    cx, cy, cw = 0.43, y + 0.08, 7.74
    ch = BODY_BOTTOM - cy
    card(slide, cx, cy, cw, ch)
    buf = io.BytesIO()
    charts.render_chart_png(spec, buf, 7.34, 4.2)
    buf.seek(0)
    slide.shapes.add_picture(buf, In(0.63), In(cy + 0.21), In(7.34), In(4.2))

    # 右侧洞察面板
    px, pw = 8.4, 4.38
    text_box(slide, px, cy + 0.02, pw, 0.35, "关键洞察", size=FONT_SIZES["insight_title"],
             color="primary", bold=True)
    rect(slide, px, cy + 0.4, 0.55, 0.045, fill="primary")
    for i, ins in enumerate(m.insights[:4]):
        iy = cy + 0.62 + i * 1.02
        rect(slide, px, iy, pw, 0.92, fill="card_bg", line="border",
             radius=0.09, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        if ins.icon:
            _draw_block_icon(slide, ins.icon, px + 0.14, iy + 0.2, 0.5,
                             color=ACCENT_CYCLE[i % len(ACCENT_CYCLE)])
        else:
            rect(slide, px + 0.18, iy + 0.29, 0.36, 0.36,
                 fill=ACCENT_CYCLE[i % len(ACCENT_CYCLE)],
                 radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            text_box(slide, px + 0.18, iy + 0.29, 0.36, 0.36, str(i + 1),
                     size=FONT_SIZES["badge"], color="on_dark", bold=True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        text_box(slide, px + 0.72, iy + 0.1, pw - 0.94, 0.75, ins.text,
                 size=FONT_SIZES["rail_text"], color="text", spacing=1.25,
                 anchor=MSO_ANCHOR.MIDDLE)
    if m.footer_note:
        footer_note(slide, m.footer_note)


RENDERERS = {
    "cover": render_cover,
    "toc": render_toc,
    "section": render_section,
    "statement": render_statement,
    "compare_3col": render_compare_3col,
    "numbered_cards": render_numbered_cards,
    "layered_stack": render_layered_stack,
    "flow": render_flow,
    "hierarchy": render_hierarchy,
    "table": render_table,
    "chart": render_chart,
    "panel_2col": render_panel_2col,
    "grid_2x2": render_grid_2x2,
    "timeline": render_timeline,
    "funnel": render_funnel,
    "matrix": render_matrix,
    "key_groups": render_key_groups,
    "figure": render_figure,
    "end": render_end,
}


# ---------------- 容量核对（qc.py 使用） ----------------
#
# 旧的 `CAPACITY` 静态容量表 + `estimate_overflow` 已经删掉：它们是"渲染算一套几何、
# qc 再背一套数字"的残留 —— 版式几何一改，表里的数字就悄悄过期（历史上反复踩）。
# 现在容量一律由**渲染同款几何函数**算：
#   文本框尺寸  → numbered_geom / layered_geom / flow_linear_geom / hierarchy_geom /
#                 statement_body_box / compare_desc_box / panel_area…
#   这个尺寸能装多少字 → cap_chars（最小字号档）/ chars_fit（指定字号档）
#   留白还能写多少字     → slack_chars / <版式>_slack
