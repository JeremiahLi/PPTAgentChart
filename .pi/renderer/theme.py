# -*- coding: utf-8 -*-
"""外观主题：色板 / 卡片长相 / 字体 / 字号阶梯 / 版面网格。

两个**正交**的外观维度，都定义在本文件里：

  - `style`：配色主题（5 套，见 `_THEMES`），决定色板 / 强调色循环 / 图表序列；
  - `card` ：卡片长相（5 种，见 `CARD_STYLES`），只决定标准卡片怎么画。

两者都由 deck.json 的同名字段选，渲染入口在画第一页之前各调一次
（`apply_style` / `apply_card_style`）。默认值 = 历史产出，改前改后像素级一致。

16:9（13.333 x 7.5 英寸）。所有版式共用本文件，改样式只改这里。
"""
from __future__ import annotations

import os

from pptx.util import Inches, Pt

# ---------- 画布 ----------
# 注意：这里全部是**浮点英寸**，不是 Inches() 的 EMU 对象 ——
# 版式函数里算完几何还要再喂给 In(x)，传 EMU 进去会被再放大 914400 倍
# （2026-10 发现 grid_2x2 就是这么飞出页面的，坑见 pitfalls 第 26 条）。
SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN_L = 0.55
MARGIN_R = 0.55
MARGIN_T = 0.42
CONTENT_W = 13.333 - 1.1


# ---------- 色板（v3：深色沉稳 + 提亮一档） ----------
COLORS = {
    "bg_page":    "F6F5F2",  # 页面暖灰底
    "card_bg":    "FFFFFF",  # 白卡片
    "border":     "E2E0DA",
    "text":       "262B33",  # 正文近黑
    "text_sub":   "59626D",  # 次要灰
    "text_light": "9AA0A8",
    "primary":    "2760A4",  # 藏青主色（提亮）
    "dark_blue":  "1A4269",
    "green":      "4A8260",  # 墨绿（提亮）
    "orange":     "B07B35",  # 赭金（提亮）
    "purple":     "6B60A0",  # 深藤紫（提亮）
    "pink":       "AD6483",  # 豆沙红（提亮）
    "cyan":       "4389BE",  # 黛蓝（提亮）
    "on_dark":    "FFFFFF",  # 深色色块上的文字（色带/徽章/封面图上的白字）

    # ---------- 封面专用 ----------
    # 封面底图 .pi/assets/template/bg-cover.png 是深蓝渐变（实测亮度 0.01–0.16），
    # 深色字压上去对比度只有 1.2–2.9:1 —— 等于看不见。封面文字一律走浅色系：
    # 白标题 6.9:1、浅蓝英文 3.5:1、浅灰蓝副题 7.1:1、落款 9.5:1。
    # 若底图被换成浅色图，layouts._cover_ink() 会自动回退到下面的深色版。
    "cover_title":     "FFFFFF",
    "cover_en":        "C5DCF0",
    "cover_sub":       "CFDCE8",
    "cover_org":       "A9C4DC",
    "cover_title_on_light": "1A4269",
    "cover_en_on_light":    "2760A4",
    "cover_sub_on_light":   "59626D",
    "cover_org_on_light":   "59626D",
}

# 编号卡片 / 对比栏循环取色的顺序
ACCENT_CYCLE = ["primary", "green", "orange", "purple", "cyan", "pink"]

# 图表系列取色顺序（多系列柱状/折线）
CHART_SERIES = ["primary", "cyan", "orange", "green", "purple"]

# 基础色板 = 默认主题（style-harness）的冻结拷贝。所有派生主题从这里合并出来，
# 不能从 COLORS 派生 —— COLORS 会被 apply_style 原地改写。
_BASE_COLORS = dict(COLORS)


# ---------- 配色主题（deck.json 的 `style` 字段按名字选） ----------
#
# 2026-10 起放宽为 5 套（用户明确要求："很多人用，外观千篇一律"）。
# 仍然只有一个默认：`style-harness`，不写 style 的老 deck 产出完全不变。
# 曾经的硬约束"非蓝主题不进注册表"已作废 —— 但**注册表仍然只许这 5 套**，
# 随手加实验色板（历史上出现过棕色 demo 混进交付目录）还是要拦，见 pitfalls 第 27 条。
#
# 为什么是「原位变更」而不是换字典：layouts / charts 都是 `from theme import
# COLORS` —— 它们持有的是**同一个 dict 对象**的引用。重新赋值 theme.COLORS
# 不会传过去，必须 clear+update 原地改（列表同理，用切片赋值）。
# 主题只能是**浅色页面**：bg-content.png / bg-cover.png 是固定素材
# （浅灰内容底 + 深蓝封面底），深色主题会跟底图打架；`cover_*` 四组封面字色
# 因此所有主题共用基础值，不随主题改。
def _palette(**overrides: str) -> dict:
    """从基础色板派生一套主题色板：只写差异项，其余键原样继承。

    必须**键对齐**：`layouts._c()` 是 `COLORS.get(name, name)`，主题少一个键就会
    走到 `RGBColor.from_string("cyan")` 直接抛错。这里合并 + 未知键当场报错，
    把这类漏键拦在 import 期而不是渲染期（tests/test_new_features.py 另有锁）。
    """
    unknown = sorted(set(overrides) - set(_BASE_COLORS))
    if unknown:
        raise KeyError(f"主题色板里没有这些键：{unknown}（键集合见 theme.COLORS）")
    return {**_BASE_COLORS, **overrides}


_THEMES: dict[str, dict] = {
    # 默认：藏青主色，色板取自《Harness工程》参考 PDF 实测值
    "style-harness": {
        "colors": dict(_BASE_COLORS),
        "accent_cycle": list(ACCENT_CYCLE),
        "chart_series": list(CHART_SERIES),
    },
    # 深蓝商务沉稳：墨蓝主色 + 石板灰中性色 + 铜金点缀，强调色收敛到 4 色
    "blue-business": {
        "colors": _palette(
            bg_page="F5F6F7", border="DDE2E8",
            text="22282F", text_sub="55606C", text_light="949BA4",
            primary="1B3A5C", dark_blue="12283D",
            green="3F6B5C", orange="A97C3F", purple="5A5F85",
            pink="8E6470", cyan="4E7A96",
        ),
        "accent_cycle": ["primary", "orange", "cyan", "green"],
        "chart_series": ["primary", "cyan", "orange", "green", "purple"],
    },
    # 明亮多色：亮蓝主色 + 珊瑚橙 / 薄荷绿 / 玫红，适合宣讲、培训、对外交流
    # （辅助色都比"纯亮色"压暗了一档：白字压在色带上要不低于默认主题的 3.66:1）
    "bright": {
        "colors": _palette(
            bg_page="F7F9FC", border="E1E8F0",
            text="1E2630", text_sub="55606E", text_light="949CA6",
            primary="1A68C8", dark_blue="0F4C8A",
            green="0B8F79", orange="E05A20", purple="7C5CE0",
            pink="D9457A", cyan="0087A0",
        ),
        "accent_cycle": ["primary", "green", "orange", "purple", "cyan", "pink"],
        "chart_series": ["primary", "cyan", "orange", "green", "purple"],
    },
    # 墨绿稳重：深墨绿主色 + 米金 / 陶土，战略、能源、可持续发展类内容
    "forest": {
        "colors": _palette(
            bg_page="F5F6F4", border="DCE3DE",
            text="222A26", text_sub="55605A", text_light="949C97",
            primary="1F5C4A", dark_blue="143F33",
            green="4F7F63", orange="B07A2F", purple="5B6E8C",
            pink="A0705F", cyan="4E8C8A",
        ),
        "accent_cycle": ["primary", "orange", "cyan", "green", "purple", "pink"],
        "chart_series": ["primary", "green", "orange", "cyan", "purple"],
    },
    # 紫灰科技：靛紫主色 + 青蓝 / 品红，产品、技术、平台方案汇报
    "violet-tech": {
        "colors": _palette(
            bg_page="F6F6FA", border="E2E2EE",
            text="24242E", text_sub="5A5A68", text_light="9797A4",
            primary="4A4E8C", dark_blue="2E3163",
            green="4C8C7A", orange="7A82C0", purple="7A6FD0",
            pink="B06090", cyan="3C8AC2",
        ),
        "accent_cycle": ["primary", "cyan", "pink", "green", "purple", "orange"],
        "chart_series": ["primary", "cyan", "pink", "green", "purple"],
    },
}

DEFAULT_STYLE = "style-harness"
CURRENT_STYLE = DEFAULT_STYLE    # 当前生效的主题名（apply_style 更新）


def apply_style(name: str | None) -> str:
    """按名字切外观主题：**原位**更新 COLORS / ACCENT_CYCLE / CHART_SERIES。

    返回实际生效的主题名：名字未登记时回退 `style-harness`（确定性，不抛错
    —— deck.style 是外观偏好，写错了不该让整份 deck 渲染失败）。
    渲染入口（render.build_pptx）每份 deck 都无条件调用一次（不能只"有值才调"：
    一个进程里会连续渲染多份 deck，否则上一份的主题会漏到下一份）。
    """
    global CURRENT_STYLE
    effective = name if (name or "") in _THEMES else DEFAULT_STYLE
    spec = _THEMES[effective]
    COLORS.clear()
    COLORS.update(spec["colors"])
    ACCENT_CYCLE[:] = spec["accent_cycle"]
    CHART_SERIES[:] = spec["chart_series"]
    CURRENT_STYLE = effective
    return CURRENT_STYLE


# ---------- 卡片长相（deck.json 的 `card` 字段按名字选） ----------
#
# 与配色主题**正交**：同一套主题可以配 5 种卡片，组合出 25 种整体观感。
# 卡片长相只改"底怎么画"（填充/描边/圆角/色条），**不许改内边距、文字框位置
# 或任何几何** —— qc 的容量估算走的是 layouts 的几何函数，改了就会与现实脱节。
#
#   flat        白底 + 细描边 + 小圆角（历史长相，默认）
#   outline     无填充（透出页面底）+ 主色浅描边 + 方角，轻量通透
#   tinted      极浅主色底 + 无描边 + 大圆角，柔和成块
#   accent-left 白卡 + 左侧细色条，条目感强
#   accent-top  白卡 + 顶部细色条，卡片头更醒目
CARD_STYLES = ("flat", "outline", "tinted", "accent-left", "accent-top")
DEFAULT_CARD = "flat"
CURRENT_CARD = DEFAULT_CARD      # 当前生效的卡片长相（apply_card_style 更新）


def apply_card_style(name: str | None) -> str:
    """按名字切卡片长相；未登记的名字确定性回退 `flat`（不抛错）。

    与 `apply_style` 一样是**进程级状态**，渲染入口每份 deck 都要调一次。
    """
    global CURRENT_CARD
    CURRENT_CARD = name if (name or "") in CARD_STYLES else DEFAULT_CARD
    return CURRENT_CARD


# ---------- 风格预设（对外只暴露 5 套，卡片长相跟着配色走） ----------
#
# 配色 5 套 × 卡片 5 种 = 25 格，让用户挑等于没给选项（他会随便选、选完还记不住）。
# 所以"PPT 风格"这个用户选项 = **一套预设**：一个名字同时定配色与卡片长相。
# 想单独换卡片长相就显式传 `card`（`--card`），显式值优先于预设 —— 这样既
# 收窄了选择面，又没把组合能力砍掉。
#
# 预设名**就是配色名**（`style` 的值），所以老的 deck.json 与命令行用法全部不变：
# 不写 card 的 deck 以前出 flat，现在出该配色的绑定卡片（视觉更成套）。
STYLE_LABELS: dict[str, str] = {
    "style-harness": "经典藏青",
    "blue-business": "深蓝商务",
    "bright":        "明亮宣讲",
    "forest":        "墨绿战略",
    "violet-tech":   "紫灰科技",
}

CARD_LABELS: dict[str, str] = {
    "flat":        "简约白卡",
    "outline":     "描边通透",
    "tinted":      "浅色柔块",
    "accent-left": "左侧色条",
    "accent-top":  "顶部色条",
}

# 每套配色绑定的卡片长相：与配色**错开搭配**（SKILL.md 里"别每次落在同一格"的
# 那条建议固化成默认值），而不是一律 flat。
STYLE_CARD_PRESET: dict[str, str] = {
    "style-harness": "flat",
    "blue-business": "accent-top",
    "bright":        "outline",
    "forest":        "tinted",
    "violet-tech":   "accent-left",
}


def resolve_card(style: str | None, card: str | None = None) -> str:
    """决定卡片长相：显式 `card` 优先，否则取该配色的预设，都落空回退 `flat`。"""
    if (card or "") in CARD_STYLES:
        return card
    preset = STYLE_CARD_PRESET.get(style if (style or "") in _THEMES else DEFAULT_STYLE)
    return preset or DEFAULT_CARD


def style_label(name: str | None) -> str:
    """配色名的中文标签（报告/交付话术用）；未登记的名字标注为回退值。"""
    return STYLE_LABELS.get(name or "", f"{name}（未登记，回退{DEFAULT_STYLE}）")


def appearance_label(style: str | None, card: str | None = None) -> str:
    """一句话说清这一份长什么样：「深蓝商务 + 顶部色条」，交付时原话给用户。"""
    eff_style = style if (style or "") in _THEMES else DEFAULT_STYLE
    eff_card = resolve_card(eff_style, card)
    return f"{STYLE_LABELS.get(eff_style, eff_style)} + {CARD_LABELS.get(eff_card, eff_card)}"


def tint(hex_color: str, ratio: float = 0.88) -> str:
    """把颜色向白色混合，ratio=0.88 表示 88% 白。用于浅色底、色块衬底。"""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    r = round(r + (255 - r) * ratio)
    g = round(g + (255 - g) * ratio)
    b = round(b + (255 - b) * ratio)
    return f"{r:02X}{g:02X}{b:02X}"


# ---------- 字体（只用系统自带字体，不随包分发字体） ----------
# 中文 = 微软雅黑，西文/数字 = Times New Roman。
#
# PowerPoint 是按**字符**在 run 的 `a:latin` / `a:ea` 之间挑字体的：同一段 run 里
# 英文/数字走 `a:latin`、中文走 `a:ea`。所以两个 typeface 必须**成对**写进每个 run
# （见 `layouts._apply_run_font`）——只写一个，另一个就会退回主题字体（Calibri）。
#
# 为什么不再随包带字体：微软雅黑 / Times New Roman 是 Windows 自带字体，内网机器上
# 一定存在（PowerPoint 属于 Office，Office 装了就有），随包分发字体既增加部署体积、
# 又带来字体授权问题。**不要再往仓库里放 fonts/ 目录或 *.ttf**。
FONT_LATIN = os.environ.get("PPT_FONT_LATIN", "Times New Roman").strip() or "Times New Roman"
                                   # 西文 / 数字（a:latin）
FONT_MAIN = "微软雅黑"              # 中文（a:ea，East Asian）
#
# 想要**全文（含英文与数字）都是微软雅黑**时，不用改代码：
#   Windows PowerShell : $env:PPT_FONT_LATIN="微软雅黑"
#   cmd                : set PPT_FONT_LATIN=微软雅黑
# 再跑 render.py。默认仍是「中文微软雅黑 + 西文 Times New Roman」。


def apply_theme_fonts(prs) -> None:
    """把母版主题的默认字体也改成「中文微软雅黑 + 西文 Times New Roman」。

    幻灯片里每段文字都显式带了 `a:latin`/`a:ea`，改主题是为了**默认值也对齐**：
    用户在 PowerPoint 里新加文本框、或编辑图表默认文本时拿到的是这两个字体，
    而不是模板自带的 Calibri。

    要同时改三处，少一处都会漏：
      - `majorFont/minorFont/a:latin` —— 西文与数字；
      - `majorFont/minorFont/a:ea` —— 中文（run 没写 a:ea 时用它）；
      - `a:font script="Hans"` —— **简体中文的按脚本覆盖项优先于 a:ea**，
        默认值是「宋体」，不改它中文就会变宋体。
    """
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for master in prs.slide_masters:
        try:
            theme_part = master.part.part_related_by(RT.THEME)
        except KeyError:
            continue
        root = etree.fromstring(theme_part.blob)
        for scheme in root.iter(f"{{{a_ns}}}fontScheme"):
            for tag in ("majorFont", "minorFont"):
                font = scheme.find(f"{{{a_ns}}}{tag}")
                if font is None:
                    continue
                for child_tag, face in (("latin", FONT_LATIN), ("ea", FONT_MAIN)):
                    node = font.find(f"{{{a_ns}}}{child_tag}")
                    if node is None:
                        node = etree.SubElement(font, f"{{{a_ns}}}{child_tag}")
                    node.set("typeface", face)
                for script in font.findall(f"{{{a_ns}}}font"):
                    if script.get("script") in ("Hans", "Hant"):
                        script.set("typeface", FONT_MAIN)
        theme_part._blob = etree.tostring(  # noqa: SLF001 —— python-pptx 没有公开的写入口
            root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------- 字号阶梯（pt，实测参考 PDF：正文 13-16，标题 20-24） ----------
FONT_SIZES = {
    "cover_title": 40,
    "cover_en": 18,
    "cover_sub": 16,
    "section_no": 96,
    "section_title": 34,
    "page_title": 23,     # 内容页大标题（判断句）
    "summary": 12.5,      # 标题下导语
    "card_title": 15,     # 卡片标题
    "card_tag": 12,
    "body": 12.5,         # 正文
    "card_body": 11.5,    # 卡片内正文（起始档；放不下由 layouts 逐档降到 10.0）
    "card_body_mid": 10.5,   # 中间档（供布局按内容挑档）
    "card_body_min": 10.0,   # 最小档：研讨型"看图说话"材料的正文下限
    "footer": 11,
    "toc_no": 20,
    "toc_text": 15,
    "toc_title": 26,      # 目录页大标题（比内容页标题 23 大一档，作为隔页）
    "toc_en": 12,         # 标题下 CONTENTS 小字
    "table_header": 12.5,
    "table_cell": 11.5,      # 起始档（放不下逐档降到 10.0）
    "table_cell_mid": 10.5,
    "table_cell_min": 10.0,
    "step_label": 13,
    "step_desc": 11,
    # 视觉层级约束（schema 层强制）
    "hero_value": 20,     # 大数字突出（比正文 card_body 大 ≥2 级：11.5→20）
    "hero_label": 11,     # 大数字下方说明
    "quadrant_title": 13, # 2x2 网格格子标题
    "quadrant_value": 18, # 2x2 网格格子大数字
    # --- 辅区 / 局部元素（原先散落在 layouts.py 里写死，收口到这里） ---
    "rail_label": 10,        # 右栏辅区卡片小标题（关键数字/边界条件）
    "rail_text": 10.5,       # 右栏辅区卡片正文（含 figure 要点栏）
    "rail_value": 17,        # 右栏辅区大字数值
    "callout_label": 9.5,    # 底部横条辅区的 label
    "callout_value": 14,     # 底部横条辅区的大字数值
    "callout_text": 8.5,     # 底部横条辅区的说明（放不下再降到 7.5）
    "block_subtitle": 12,    # 页内小标题（「支撑论据」「图注」这类）
    "caption": 12.5,         # 图注 / 图片说明
    "step_label_micro": 9,   # 环形流程节点 label（节点只有 1.72in 宽）
    "step_desc_micro": 7.5,  # 环形流程节点描述（比全局下限 10pt 更小，环是特例）
    "hero_value_big": 20,    # statement 高亮卡大数字（起始档，放不下降到 13）
    "end_title": 34,         # 结尾页大标题
    "end_heading": 12.5,     # 结尾页每段的段头
    "cover_mark": 12,        # 封面装饰性小字
    "cover_org": 11,         # 封面落款/单位
    "section_sub": 14,       # 章节页副标题
    "card_sub": 11.5,        # 卡片的一句话定位（副标题位）
    "stmt_head": 11,         # statement 卡头（组标题行）
    "layer_label": 13,       # 层级条的层标签（有无色带图标都用这一档）
    "layer_tag": 9.5,        # 层级条右端的关键词胶囊
    "badge": 11,             # 序号徽章里的数字
    "micro_label": 9.5,      # 卡片内的小标签（时间点/角标/循环说明）
    "note": 10,              # 辅助说明小字
    "insight_title": 13,     # 图表页「关键洞察」小标题
    "hier_root": 13,         # 层级图根节点（宽度按字数算，不能用 15pt 的 card_title）
}
