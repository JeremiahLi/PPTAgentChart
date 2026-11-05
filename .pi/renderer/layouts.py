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

import charts
import theme
from theme import (ACCENT_CYCLE, CHART_SERIES, COLORS, CONTENT_W, FONT_LATIN,
                   FONT_MAIN, FONT_SIZES, MARGIN_L, tint)
from icons import draw_icon

In = Inches

# 页面尺寸与背景图路径（封面用 bg-cover，其余页用 bg-content；缺图回退纯色）。
SLIDE_W, SLIDE_H = 13.333, 7.5
_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_BG_COVER = os.path.join(_BASE, "assets", "template", "bg-cover.png")
_BG_CONTENT = os.path.join(_BASE, "assets", "template", "bg-content.png")

# 图表渲染开关：native=可编辑原生图表（默认），mpl=matplotlib 高质感 PNG。
CHART_MODE = "native"

# ---------------- 页面纵向骨架（所有版式共用的唯一来源） ----------------
#
# 为什么要拎出来：密度上不去的一半原因在"固定开销"——页脚条原来占 0.56in 而且
# 每个版式还各自再扣一次（`6.55 - 0.56`），同一块页高被扣两遍。
# 现在统一成：内容带到 BODY_BOTTOM=6.80 为止，页脚条压扁成 0.36in 挪到 6.90。
# 这样"有没有写 footer_note"**不再影响内容带高度**，各版式也不用各算一套。
BODY_BOTTOM = 6.80                    # 内容带下沿（所有版式的默认底线）
FOOTER_Y = 6.90                       # 页脚条上沿（与页码同一行高）
FOOTER_H = 0.36                       # 页脚条高度（压扁：单行 9~11pt）
FOOTER_RIGHT = 10.85                  # 页脚条右端：避开底图右下角单位水印
                                      # （实测水印 x 10.95–12.90 / y 6.97–7.15）

# 右纵栏（辅区）骨架。放在这里而不是版式段内：statement_geom / highlight_h
# 在文件更靠前的位置就要用到，常量必须先生效。
AUX_RAIL_W = 3.55                     # 右侧纵栏宽度（与表格页 side 同宽）
AUX_GAP = 0.22                        # 主内容与右纵栏的间距
AUX_MIN_MAIN_W = 12.23 - (AUX_RAIL_W + AUX_GAP)   # 让出右栏后主内容还剩多宽


# ---------------- 基础工具 ----------------

def _c(name: str) -> RGBColor:
    return RGBColor.from_string(COLORS.get(name, name))


def _set_font_on_rPr(rPr) -> None:
    """把 `a:latin` / `a:ea` / `a:cs` 三处 typeface 成对写上（顺序按 ECMA-376）。

    PowerPoint 按**字符**挑字体：西文/数字走 `a:latin`、中文走 `a:ea`。所以每个 run
    都要成对写死，不能只写一个 —— 只写 latin，中文会退回主题字体（模板自带 Calibri，
    中文就变成 PowerPoint 的替身字体）；只写 ea，英文/数字会变 Calibri。

    元素顺序必须是 `latin → ea → cs`（CT_TextCharacterProperties 的顺序约束），
    所以统一用 `latin.addnext(...)` 插进去，而不是 append 到 rPr 末尾
    （rPr 末尾可能是 `a:sym` / `a:hlinkClick`）。
    """
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
    """给一个 run 写上「中文微软雅黑 + 西文 Times New Roman」。

    全 deck 唯一的字体入口：`text_box` / `_fill_cell` / 页码（render.py）都走它。
    自己新建 run 时**必须**调一次，否则那一段会退回模板字体（Calibri）。
    """
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:                          # python-pptx 的 run.font.name 会建 a:latin
        run.font.name = FONT_LATIN
    _set_font_on_rPr(rPr)


def _est_lines(s: str, width: float) -> int:
    """估算文本在 width 英寸内的换行行数（11.5pt：中文≈0.16in、ASCII≈0.085in）。"""
    w = sum(0.16 if ord(c) > 127 else 0.085 for c in s)
    return max(1, math.ceil(w / width))


def _lines_in(text: str, width: float, size_pt: float) -> int:
    """文字在 width 英寸里占几行。

    字宽取**保守**值（CJK 1.06em、ASCII 0.55em）：PowerPoint 实际排版比
    「CJK=1em」更宽（标点、括号、中英混排都要占额外空间），估乐观的后果是
    **文字溢出卡片边框**（v15 的 P4/P11 就是这样）。宁可框高一点，不许溢出。
    """
    plain = (text or "").replace("**", "")
    em = size_pt / 72
    w = sum(em * 1.06 if ord(c) > 127 else em * 0.55 for c in plain)
    return max(1, math.ceil(w / max(0.5, width)))


def _para_h(text: str, width: float, size_pt: float, spacing: float = 1.3) -> float:
    """一段文字在 width 英寸内占用的高度（英寸），供卡片按内容比例分配高度。"""
    return _lines_in(text, width, size_pt) * size_pt / 72 * spacing


# ---------------- 图标自动匹配 ----------------
#
# ---- 图标配额：一个内容块最多一个图标 ----
# 每个「块」（卡片 / 栏 / 层 / 步骤 / 节点 / 正文分组）只画一个图标，位置固定在块头：
# 块自带 icon 字段就用它，没写就按块的文字自动挑一个补上（保证每块都有图标）。
# 正文 / 描述文字前面**不再额外补图标** —— 块头已经有图标时，第二个图标会和它
# 打架，整页看着发飘（v14 的 P9 / P17 / P19 就是这样）。

# 关键词 -> 图标名 映射表（用于根据文本内容自动选图标）
_ICON_KEYWORDS: list[tuple[list[str], str]] = [
    # 效率 / 性能
    (["效率", "性能", "速度", "快速", "提升", "优化", "加速", "高效"], "gauge"),
    # 安全 / 保护
    (["安全", "保护", "防护", "风险", "合规", "权限", "边界", "约束"], "shield-check"),
    # 数据 / 分析
    (["数据", "分析", "指标", "统计", "度量", "监控", "观测"], "chart-bar"),
    # 协作 / 团队
    (["协作", "团队", "协同", "合作", "沟通", "交流"], "users"),
    # 流程 / 过程
    (["流程", "过程", "步骤", "环节", "阶段", "工序"], "process"),
    # 质量 / 品质
    (["质量", "品质", "标准", "规范", "准确", "精确", "正确"], "check-circle"),
    # 成本 / 资源
    (["成本", "资源", "投入", "产出", "收益", "预算"], "coin"),
    # 创新 / 改进
    (["创新", "改进", "改革", "升级", "迭代", "演进"], "lightbulb"),
    # 文档 / 报告
    (["文档", "报告", "记录", "文档化", "归档", "撰写"], "file-text"),
    # 部署 / 发布
    (["部署", "发布", "上线", "交付", "推出"], "rocket"),
    # 测试 / 验证
    (["测试", "验证", "检验", "评测", "验收"], "test-pipe"),
    # 配置 / 设置
    (["配置", "设置", "参数", "选项", "偏好"], "settings"),
    # 网络 / 连接
    (["网络", "连接", "通信", "传输", "协议"], "network"),
    # 用户 / 客户
    (["用户", "客户", "使用者", "体验", "界面"], "user"),
    # 搜索 / 查询
    (["搜索", "查询", "检索", "查找", "索引"], "search"),
    # 通知 / 消息
    (["通知", "消息", "提醒", "告警", "预警"], "bell"),
    # 时间 / 计划
    (["时间", "计划", "排期", "日程", "周期"], "clock"),
    # 成功 / 完成
    (["成功", "完成", "达成", "实现", "获得"], "circle-check"),
    # 失败 / 错误
    (["失败", "错误", "异常", "故障", "缺陷"], "alert-circle"),
    # 增长 / 上升
    (["增长", "上升", "增加", "提高", "扩大"], "trending-up"),
    # 下降 / 减少
    (["下降", "减少", "降低", "缩小", "衰减"], "trending-down"),
    # 稳定 / 平衡
    (["稳定", "平衡", "均衡", "持续", "稳态"], "scale"),
    # 自动化
    (["自动", "自动化", "智能", "AI", "机器"], "cpu"),
    # 数据库 / 存储
    (["数据库", "存储", "持久化", "备份", "仓库"], "database"),
    # API / 接口
    (["API", "接口", "对接", "集成", "调用"], "plug"),
    # 代码 / 开发
    (["代码", "开发", "编程", "编码", "实现"], "code"),
    # 架构 / 设计
    (["架构", "设计", "模式", "结构", "框架"], "layers"),
    # 目标 / 方向
    (["目标", "方向", "愿景", "使命", "战略"], "target"),
    # 问题 / 挑战
    (["问题", "挑战", "难点", "痛点", "障碍"], "help-circle"),
    # 解决方案
    (["解决", "方案", "对策", "策略", "方法"], "puzzle"),
    # 学习 / 培训
    (["学习", "培训", "教育", "知识", "技能"], "book-open"),
    # 共享 / 分享
    (["共享", "分享", "分发", "传播"], "share"),
    # 环境 / 生态
    (["环境", "生态", "平台", "系统", "基础设施"], "stack-3"),
    # 反馈 / 评价
    (["反馈", "评价", "评估", "评审", "回顾"], "message-circle"),
    # 循环 / 重复
    (["循环", "重复", "迭代", "轮次", "回环"], "refresh"),
    # 决策 / 选择
    (["决策", "选择", "判断", "判定", "决定"], "git-branch"),
    # 保护 / 保险
    (["保护", "保险", "保障", "保证", "确保"], "shield"),
    # 工具 / 器具
    (["工具", "器具", "设备", "装置", "仪器"], "tool"),
    # 构建 / 建造
    (["构建", "建造", "搭建", "建立", "创建"], "hammer"),
]


def _auto_icon(text: str) -> str | None:
    """根据文本内容自动匹配合适的图标名。

    遍历关键词映射表，命中即返回对应图标；都不命中时返回 None。
    """
    if not text:
        return None
    for keywords, icon_name in _ICON_KEYWORDS:
        for kw in keywords:
            if kw in text:
                return icon_name
    return None


# 兜底图标：前面都挑不出来时用它，保证「一块一个图标」永不落空。
# 这几个名字在 .pi/assets/icons 里都确认存在（.pi/tests/smoke.py 会回归验证）。
_ICON_FALLBACKS = ("file-text", "layers", "target")

_ICON_CACHE: dict[str, bool] = {}


def icon_resolvable(name: str | None) -> bool:
    """这个名字渲染器能不能真的画出图形（别名 / Tabler 名 / 库规格都算）。

    真值来自 iconlib 的实际解析，不是"填了就算" —— 拼错的中文（「版本」）在旧实现里
    会被当成有效名字返回，调用方于是不再兜底，**整块没有图标且没有任何提示**。
    P6「Git 提交记录」掉图标就是这么来的。
    """
    if not name:
        return False
    hit = _ICON_CACHE.get(name)
    if hit is None:
        try:
            from iconlib import resolve_icon as _spec
            hit = bool(_spec(name))
        except Exception:
            hit = False
        _ICON_CACHE[name] = hit
    return hit


def _resolve_icon(item_icon: str | None, *texts: str) -> str | None:
    """解析这一块的唯一图标，**返回值保证画得出来**；实在挑不出才返回 None。

    优先级：块自带的 icon（仅在能解析时采纳）→ 按块文字自动匹配 → 通用兜底图标。
    一个块只调一次，画在块头。别在正文前再补第二个图标（见上面的图标配额）。
    """
    if item_icon and icon_resolvable(item_icon):
        return item_icon
    auto = _auto_icon(" ".join(t for t in texts if t))
    if auto and icon_resolvable(auto):
        return auto
    for fb in _ICON_FALLBACKS:
        if icon_resolvable(fb):
            return fb
    return None


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


# 结论页正文：短正文**单卡连续排**（不拆卡），长正文最多 3 组
BODY_SINGLE_MAX = 160       # ≤ 这个字数就排成一张卡：拆卡会给每张卡再扣一遍内边距


# ---------------- 卡内文本结构：置顶关键句 + 小节 ----------------
#
# 反馈（v22 P11）：纯文字卡片就是一坨连贯段落，观众扫不到重点，整页"没有特点"。
# 卡片因此要有骨架 —— 但**骨架不许靠放大字号**：早先把关键句提到卡顶、大一号加粗
# 外加一条细线，多出来的那行会把文字顶出卡片背景（用户实测到了）。最终规则：
#   1. **同字号**：关键句/关键短语与正文一个字号，只**加粗 + 强调色**（strong_color）；
#   2. **关键句**：每张卡的第一句自动加粗染色（`claim=True`），仍排在正文流里；
#   3. **关键短语**：作者写 `**短语**：正文…`，冒号后直接跟正文、同一段里往下排。
# 于是卡片高度 = 普通文字的高度，不再有"关键句那行放不下"的溢出。
# 短语标记两种写法都认（冒号写在加粗外/内）：
#   `**关键短语**：正文…`  与  `**关键短语：**正文…`
# 冒号**必须存在** —— 否则 `**回归通过率**提升到 88%：…` 这种普通加粗会被误判。
_SECTION_RE = re.compile(r"\*\*(?P<h1>[^*\n]{2,14}?)\*\*[：:]\s*"
                         r"|\*\*(?P<h2>[^*\n]{2,14}?)[：:]\*\*\s*")


def _sec_head(mk) -> str:
    return (mk.group("h1") or mk.group("h2") or "").strip().rstrip("：:")


CLAIM_MAX = 48          # 关键句（第一句，会被加粗染色）建议不超过这个字数


def split_claim(text: str) -> tuple[str, str]:
    """卡内文本 → (关键句, 其余)。关键句 = 第一句（不足 8 字就连上第二句）。"""
    t = (text or "").strip()
    if not t:
        return "", ""
    cut = 0
    for _ in range(2):
        m = re.search(r"[。！？]", t[cut:])
        if m is None:
            break
        cut += m.end()
        if len(t[:cut].replace("**", "").strip()) >= 8:
            break
    return t[:cut].strip() or t, t[cut:].strip()


def split_sections(text: str) -> list[tuple[str, str]]:
    """卡内文本 → [(关键短语, 其后的正文)]；没写 `**短语**：` 就整体当作一节。

    qc 用它数"这张卡有几处关键短语"；渲染时它们是**同行跟排**的（不单独起一行）。
    """
    t = (text or "").strip()
    if not t:
        return []
    marks = list(_SECTION_RE.finditer(t))
    if not marks:
        return [("", t)]
    out: list[tuple[str, str]] = []
    if marks[0].start() > 0:
        lead = t[:marks[0].start()].strip()
        if lead:
            out.append(("", lead))
    for i, mk in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(t)
        out.append((_sec_head(mk), t[mk.end():end].strip()))
    return out


def card_paras(text: str, size_pt: float, *, claim: bool = False,
               key: str = "primary", color: str = "text") -> list[tuple[str, dict]]:
    """卡内文本 → `text_box` 的段落列表（**一段到底**）。

    关键句（第一句，`claim=True` 时）自动包进 `**…**`；作者写的 `**短语**：`
    原样保留 —— 两者都由 `text_box` 渲染成「**同字号** + 加粗 + 强调色」，
    冒号后的正文继续在同一段里往下排（不单独起行，所以不会多占高度）。
    """
    t = (text or "").strip()
    if not t:
        return []
    if claim:
        head, rest = split_claim(t)
        if head:
            t = f"**{head.replace('**', '')}**{rest}"
    if t.count("**") % 2:      # 标记不成对：宁可不加粗，也别让半张卡被误加粗
        t = t.replace("**", "")
    return [(t, {"size": size_pt, "color": color, "strong_color": key})]


def card_text_h(text: str, width: float, size_pt: float, spacing: float = 1.30,
                *, claim: bool = False, key: str = "primary") -> float:
    """卡内文本占的高度 —— **就是普通文字的高度**（关键句/短语不放大、不另起行）。"""
    return sum(_para_h(t, width, kw["size"], spacing)
               for t, kw in card_paras(text, size_pt, claim=claim, key=key))


def draw_card_text(slide, x, y, w, h, text, size_pt, spacing, *, key="primary",
                   claim=False, color="text") -> float:
    """画卡内文本：同字号，关键句/关键短语只**加粗 + 强调色**（不置顶、不放大）。"""
    paras = card_paras(text, size_pt, claim=claim, key=key, color=color)
    if not paras:
        return y
    text_box(slide, x, y, w, h, paras, size=size_pt, spacing=spacing,
             strong_color=key)
    return y


def split_body_groups(text: str) -> list[str]:
    """把正文拆成「卡片组」：每组是一张卡，返回该卡的文本。

    **作者可以用换行指定卡片边界**（每张卡的第一句会被加粗染色成关键句，
    卡内用 `**关键短语**：` 标重点）—— 版式最多 3 张卡（2 列 × 2 行），多的自动并进前一张。
    没写换行时按老规矩自动分组：短正文（≤ BODY_SINGLE_MAX 字）**单卡连续排**
    （拆卡会给每张卡再扣一遍内边距），长正文按**字数**分 2~3 组、每组 150~180 字。

    qc 也用它（"每张卡的关键句"必须和渲染出来的卡一一对应），所以是公开名。
    """
    raw = (text or "").strip()
    if not raw:
        return [raw]
    paras = [p.strip() for p in raw.split("\n") if p.strip()]
    if not paras:
        return [raw]
    if 1 < len(paras) <= 3:
        return paras
    while len(paras) > 3:
        paras[-2] += paras.pop()
    flat = "".join(paras)
    sentences = [s.strip() for s in re.split(r'(?<=[。！？])', flat) if s.strip()]
    if not sentences:
        return [flat]
    if len(flat) <= BODY_SINGLE_MAX or len(sentences) <= 2:
        return ["".join(sentences)]
    n = len(sentences)
    n_groups = min(3, max(2, math.ceil(len(flat) / 170)))
    base = n // n_groups
    extra = n % n_groups
    groups, idx = [], 0
    for g in range(n_groups):
        size = base + (1 if g < extra else 0)
        if size > 0:
            groups.append("".join(sentences[idx:idx + size]))
            idx += size
    while len(groups) > 3:
        groups[-2] += groups[-1]
        groups.pop()
    return groups


HL_LABEL_H = 0.20        # 高亮卡 label 行
HL_VALUE_H = 0.46        # 高亮卡大数字行（30 字降到 13pt 时是两行，必须给够）
HL_DESC_TOP = 0.78       # 高亮卡 desc 起点（相对卡片上沿）
HL_DESC_SIZE = 9.5
HL_DESC_SPACING = 1.2


def highlight_h(hl, rail_w: float = AUX_RAIL_W) -> float:
    """高亮数据卡的高度：**按 desc 实际行数算**。

    旧实现写死 1.10in、desc 框固定 0.30in —— desc 一过 40 字就压出卡片下沿
    （schema 允许 80 字，实测需要 0.63in）。这里让卡片随 desc 长高，
    渲染与 qc 共用，别再各写一个数。
    """
    if hl is None:
        return 0.0
    if not hl.desc:
        return HL_DESC_TOP + 0.02
    # 图标占 0.32in + 左右内边距；按有图标这个更窄的情形算（保守）
    vl_w = max(1.2, rail_w - 0.46 - 0.20)
    lines = _lines_in(hl.desc, vl_w, HL_DESC_SIZE)
    desc_h = lines * HL_DESC_SIZE / 72 * HL_DESC_SPACING
    return HL_DESC_TOP + desc_h + 0.08


def statement_geom(m) -> dict:
    """结论页几何（渲染与 qc 共用同一套数字）。

    正文吃左侧（有右栏时宽 **8.46in**），支撑论据与高亮数据卡都放**右栏** ——
    高亮卡原来钉在左下角，白压掉正文 1.26in 高度 + 1.11in 宽度。
    """
    top = header_y(getattr(m, "summary", None)) + 0.08
    bottom = BODY_BOTTOM
    pts = list(getattr(m, "points", None) or [])
    hl = getattr(m, "highlight", None)
    hl_h = highlight_h(hl) if hl is not None else 0.0
    if pts or hl is not None:
        x, w = 0.55, AUX_MIN_MAIN_W
        rx = 0.55 + AUX_MIN_MAIN_W + AUX_GAP
        rw = AUX_RAIL_W
    else:
        x, w, rx, rw = 0.55, 12.23, 0.0, 0.0
    return {"top": top, "bottom": bottom, "x": x, "w": w, "rx": rx, "rw": rw,
            "hl_h": hl_h, "hl_y": (bottom - hl_h) if hl is not None else None,
            "pts_top": top + (0.44 if pts else 0.0),
            "pts_bottom": bottom - ((hl_h + 0.16) if hl is not None else 0.0)}


def statement_body_box(m) -> tuple[float, float]:
    """结论页正文卡的文本区 (宽, 高)，qc 用它核对容量。

    按**两列**这个最窄情形算（正文被拆成 3 组时会变两列），保证是保守估计。
    """
    g = statement_geom(m)
    card_w = (g["w"] - 0.12) / 2
    return card_w - 0.60, g["bottom"] - g["top"] - 0.26


def statement_slack(m) -> tuple[float, float, int]:
    """结论页正文区 (下方留白, 可用高度, 用的字号)。

    「留白靠补内容」对结论页同样成立：正文写太短，左侧卡片下方就会空一条 ——
    qc 用它报出"还空多少、还能写几字"，让 LLM 把论据写满或并页。
    """
    g = statement_geom(m)
    groups = split_body_groups(m.body)
    n_groups = max(1, len(groups))
    n_cols = 2 if n_groups >= 3 else 1
    card_w = (g["w"] - (n_cols - 1) * 0.12) / n_cols
    avail = g["bottom"] - g["top"]
    full_last = n_cols == 2 and n_groups % n_cols == 1
    row_hs, size_pt, _sp, _pad = _fit_rows(groups, n_cols, card_w, avail,
                                           full_last=full_last)
    need = sum(row_hs) + 0.12 * (len(row_hs) - 1)
    return max(0.0, avail - need), avail, size_pt


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


# ---------------- 卡片长相（theme.CARD_STYLES 的唯一画法） ----------------
#
# 只有这里能决定"卡片长什么样"。配色主题给颜色，卡片长相给填充/描边/圆角/色条。
# 铁律：**长相只改底怎么画，不改内边距、不新建文字框、不动任何几何** ——
# qc 的容量估算走的是本文件的几何函数，动了内边距，报告就和现实对不上了。
CARD_RADIUS_FLAT = 0.055       # flat 的默认圆角（调用点可覆盖：历史版式有 0.05~0.09）
CARD_RADIUS_OUTLINE = 0.02     # outline / accent-left / accent-top：近方角
CARD_RADIUS_TINTED = 0.10      # tinted：大圆角
CARD_RAIL_W = 0.055            # 历史左轨宽（flat 下与改动前逐像素一致，不许改）
CARD_ACCENT_W = 0.045          # accent-left 胶囊轨宽（受最小内边距 0.10 约束）
CARD_ACCENT_PAD = 0.02         # 胶囊轨距卡片左边缘
CARD_ACCENT_TOP_H = 0.09       # accent-top 顶带高（必须 > 版式自带的 0.05~0.075 色条，
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


SUMMARY_W = 11.4   # 页头结论段文本区宽（英寸）


def summary_block_h(summary: str | None) -> float:
    """页头结论段打底块的高度；没写结论段就是 0。"""
    if not summary:
        return 0.0
    size_pt = FONT_SIZES["summary"]
    # 行数用 _lines_in（保守字宽），块高随之自适应——结论段文字不许溢出浅色底块
    lines = _lines_in(summary, SUMMARY_W - 0.24, size_pt)
    return min(1.9, lines * size_pt / 72 * 1.4 + 0.24)


def header_y(summary: str | None) -> float:
    """内容页在页头之下还剩多少高度 —— 即 content_header 的返回值。

    版式几何要靠它算卡片高；qc 也要靠它算容量，所以单独拎出来，
    保证「渲染用的高度」和「质检用的高度」永远是一个数。
    """
    if not summary:
        return 1.35
    return 1.02 + summary_block_h(summary) + 0.20


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


TOC_TOP = 1.90            # 目录首行上沿（大标题 + CONTENTS 之下）
TOC_BOTTOM = 6.20         # 目录末行下沿（留出页脚带）
TOC_GAP_X = 0.50
TOC_GAP_Y = 0.22
TOC_ROW_MAX = 1.45        # 目录条最大高度：再多就居中留白，不把条目拉成胖块
TOC_X0 = 0.90


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
    from theme import ACCENT_CYCLE
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


def _stack_heights(heights: list[float], avail: float,
                   gap_min: float = 0.10, gap_max: float = 0.32,
                   pad_max: float = 0.20) -> tuple[list[float], float]:
    """把一组「内容真正需要的高度」塞进 avail。

    余量**先补卡片内边距（每块 ≤pad_max）、再补间距（≤gap_max）**，
    绝不超过 avail —— 各块总高 + 间距×(n-1) 必须 ≤ avail，
    否则就是"框比文字矮"，文字必溢出（v15/v16 的 P4）。
    返回 (各块高度, 间距)；第一块从调用方给的顶边开始，不额外留顶部间距。
    """
    n = len(heights)
    if not n:
        return [], 0.0
    slack = avail - sum(heights) - gap_min * (n - 1)
    if slack <= 0:
        return list(heights), gap_min
    per = min(pad_max, slack / max(1.0, n * 1.5))
    rest = max(0.0, slack - per * n)
    gap = min(gap_max, gap_min + rest / max(1, n - 1))
    return [h + per for h in heights], gap


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


def aux_rail_fits(main_min_w: float) -> bool:
    """主内容至少需要 main_min_w 英寸才能成立时，还放得下右纵栏吗？"""
    return main_min_w <= AUX_MIN_MAIN_W


def card_frame(summary: str | None, has_aux: bool, main_min_w: float,
               bottom: float, top_pad: float = 0.12,
               aux_h: float | None = None) -> dict:
    """卡片型版式的几何：主内容框 + 辅区落位（渲染与 qc 共用，别在别处另算）。

    `bottom` 传「没有辅区时主内容的下沿」，返回值里已经按落位调整好：
      mode="none"   无辅区：主内容满宽满高
      mode="rail"   辅区在**右侧纵栏**：主内容收窄到 w，**高度不受影响**（左右分置）
      mode="bottom" 辅区在底部横条：主内容满宽，高度让出 aux_h + 0.10

    `aux_h` 传 `callout_height(notes)`；不传就按最大高度保守估算（qc 的默认口径）。
    """
    top = header_y(summary) + top_pad
    h = CALLOUT_H if aux_h is None else aux_h
    fr = {"mode": "none", "top": top, "bottom": bottom, "x": 0.55, "w": 12.23,
          "aux_h": h, "rail_x": None, "rail_w": AUX_RAIL_W,
          "rail_top": top, "rail_bottom": bottom}
    if not has_aux:
        return fr
    if aux_rail_fits(main_min_w):
        fr["mode"] = "rail"
        fr["w"] = AUX_MIN_MAIN_W
        fr["rail_x"] = 0.55 + fr["w"] + AUX_GAP
    else:
        fr["mode"] = "bottom"
        fr["bottom"] = bottom - (h + 0.10)
    return fr


def draw_aux(slide, notes, fr):
    """按 card_frame 定好的落位画辅区（右栏 / 底部横条）。"""
    if not notes or fr["mode"] == "none":
        return
    if fr["mode"] == "rail":
        _render_rail(slide, notes, fr["rail_x"], fr["rail_top"], fr["rail_w"],
                     fr["rail_bottom"])
    else:
        _render_callout(slide, notes, fr["bottom"] + fr["aux_h"] + 0.10, fr["aux_h"])


# 各版式「主内容至少要多宽才能成立」——窄于这个宽度就会被挤瘪，于是辅区退回底部横条。
# 数字不是拍脑袋：三栏/卡片栏宽 3.6in 是 11.5pt 正文能写两行以上的下限，
# 层级子模块卡 1.35in 是 9.5pt 描述能排三行的下限。
_MIN_COL_W = 3.60                  # 三栏 / 多列网格：一栏最窄宽度
_MIN_CARD_W = 3.60                 # 编号卡片：一张卡最窄宽度
NO_X = 0.24                        # 编号卡：序号左起位置
NO_GUTTER = 0.72                   # 编号卡：序号占用的横向槽宽（原来 1.05，收窄 0.33in 还宽度给正文）
# 块头右上角图标占的横向空间：**标题必须给它让位**。
# 漏掉这一步的后果：3 列卡（P9 那种 5 张卡）里长标题直接压到图标上（用户实测第 3/5 张卡）。
CARD_ICON_W = 0.48                 # 块头图标尺寸
CARD_ICON_RIGHT = 0.66             # 图标左沿距卡片右沿
ICON_TITLE_GAP = 0.14              # 标题与图标之间的最小间距


def block_title_w(card_w: float, left_gutter: float, has_icon: bool,
                  right_pad: float = 0.25) -> float:
    """块头标题可用的宽度：有图标时把图标那一段让出来，避免标题压图标。"""
    right = (CARD_ICON_RIGHT + ICON_TITLE_GAP) if has_icon else right_pad
    return max(0.8, card_w - left_gutter - right)


def fit_one_line(text: str, max_w: float, start: float | None = None,
                 min_size: float = 11.5) -> float:
    """把一行标题缩进 max_w 内：放不下就逐档缩字号（最低 min_size）。

    渲染与 qc 共用 —— qc 判断"标题会不会压图标/折成两行"用的就是这个口径。
    """
    size = start or FONT_SIZES["card_title"]
    while size > min_size and _cell_text_w(text, size) > max_w:
        size -= 0.5
    return size
_MIN_FLOW_W = 1.30                 # 线性流程：一步最窄宽度
COMPARE_MIN_W = 3 * _MIN_COL_W + 2 * 0.28        # 三栏对比 → 永远用底部横条
LAYER_MIN_W = 7.80                 # 层级堆栈：箭头槽 + 色带 2.5 + 描述 4.0
FLOW_CYCLE_MIN_W = 6.00            # 环形流程：环可以缩，太小就不像个环了
PANEL_MIN_W = 2 * 4.30 + 0.30      # 双模块：每个模块要 4.3in 才装得下迷你表
# 层级图的子模块卡**高度贴内容走**（拉长就是空框），纵向本来就有富余：
# 让出右栏只会多出纵向留白，所以辅区固定走底部横条。
HIER_MIN_W = 12.50


# ---------------- 正文自适应档位（全局降档序列） ----------------
#
# 这里是"把适配责任从 LLM 手里拿回渲染器"的落点：
# 内容放不下时**渲染器自己降档**（最低 10pt），而不是让 LLM 删字改短 ——
# 过去 6 类版式的描述溢出都是 error（阻止渲染），LLM 的理性策略就是写保守，
# 于是"过度精简"成了机制性结果，不是模型偷懒。
FIT_STEPS = ((11.5, 1.30), (11.0, 1.24), (10.5, 1.18), (10.0, 1.12))
# 结论页正文区更宽更高，允许往上走到 12.5pt（读起来更舒服），往下仍能降到 10pt
STATEMENT_STEPS = ((12.5, 1.35), (12.0, 1.30), (11.5, 1.30), (11.0, 1.24),
                   (10.5, 1.18), (10.0, 1.12))
_STATEMENT_STEPS = STATEMENT_STEPS      # 兼容旧引用（statement 一直用它挑档）
# 层级图的子模块卡很窄，字号天然更小（9.5 是它的起始档，不是"降到的最小档"）
HIER_STEPS = ((9.5, 1.20), (9.0, 1.15), (8.5, 1.10))
# 子模块卡最多吃掉的留白高度（描述居中，长了也不像空框；再多就该补内容/占位了）
HIER_GROW_MAX = 0.90


def pick_step_ex(need_fn, avail: float,
                 steps: tuple | None = None) -> tuple[float, float, bool]:
    """同 `pick_step`，但把「放不放得下」一并返回。

    很多越界不是"挑档挑错了"，而是**所有档都放不下**时仍返回最低档、渲染照画
    （文字直接压出卡片）。调用方拿到 `fits=False` 后可以：继续降到硬下限、
    或交给 `--qc` 按"必然溢出"处理 —— 总之别再假装排好了。
    """
    ladder = steps or FIT_STEPS
    for size, spacing in ladder:
        if need_fn(size, spacing) <= avail:
            return size, spacing, True
    return ladder[-1][0], ladder[-1][1], False


def pick_step(need_fn, avail: float, steps: tuple | None = None) -> tuple[float, float]:
    """按「这一档需要多少」挑档：`need_fn(size, spacing) -> 数值`，取第一个 ≤ avail 的档。

    放不下就继续降档；降到最低档仍不满足也照样返回最低档（由 `--qc` 点名
    "已降到最低档仍放不下"）—— 渲染**绝不**因此失败。
    """
    return pick_step_ex(need_fn, avail, steps)[:2]


def fit_size_ex(texts, width: float, height: float,
                steps: tuple | None = None, measure=None) -> tuple[float, float, bool]:
    """`fit_size` 的三元版：多返回一个 `fits`（最小档也放不下时为 False）。"""
    if measure is None:
        def measure(t, sz, sp):
            return _lines_in(t, width, sz) * sz / 72 * sp
    return pick_step_ex(lambda sz, sp: max((measure(t, sz, sp) for t in texts if t),
                                           default=0.0), height, steps)


def fit_size(texts, width: float, height: float,
             steps: tuple | None = None, measure=None) -> tuple[float, float]:
    """按内容挑字号档：返回第一个「所有文本都放得进 (width, height)」的 (字号, 行距)。

    `measure(text, size, spacing) -> 高度` 可换口径（卡内文本带小节标题时走
    `card_text_h`）—— **挑档与绘制必须同一口径**，否则会出现"挑了大档、画出来超高"。
    """
    return fit_size_ex(texts, width, height, steps, measure)[:2]


def step_for_len(n_chars: int, width: float, height: float,
                 steps: tuple | None = None) -> tuple[float, float]:
    """`fit_size` 的反向查询：这一段写 n_chars 字会落在哪个档位。

    qc 用它报告"本页会降到 X pt"，并据此算容量；渲染与 qc 同一套档位。
    """
    ladder = steps or FIT_STEPS
    for size, spacing in ladder:
        if n_chars <= chars_fit(width, height, size, spacing):
            return size, spacing
    return ladder[-1]


def _fit_rows(groups: list[list[str]], n_cols: int, card_w: float,
              avail: float, gap_min: float = 0.10,
              full_last: bool = False, gap_x: float = 0.12
              ) -> tuple[list[float], float, float, float]:
    """给正文分组挑一档字号/行距，使所有卡片都装得下。

    返回 (各行高度, 字号, 行距, 上下内边距)。装不下就逐档降级（字小、行距紧），
    降到最低档还放不下才交给 --qc 拦（那种情况该拆页，不该硬塞）。

    `full_last`：3 组时会排成「2 列 + 最后一张跨整行」，算高必须按**跨行后的宽度**
    算，否则末卡会按半宽估出一个偏高的框（页面上就是卡片下半截空一条）。
    """
    n_rows = max(1, math.ceil(len(groups) / n_cols))
    full_w = card_w * n_cols + (n_cols - 1) * gap_x
    row_hs, size_pt, spacing, pad = [0.42], 10.0, 1.12, 0.12
    for size_pt, spacing in _STATEMENT_STEPS:
        pad = 0.13 if size_pt >= 11.0 else 0.11      # 单侧内边距（上下各一份）
        # 每张卡 = 「卡头标题行（首句 ≤16 字时）+ 其余正文」—— 算高必须把卡头
        # 行也算进去，否则文字直接压出边框（v16 P4 的老坑，口径必须与绘制一致）
        row_hs = []
        for r in range(n_rows):
            row = groups[r * n_cols:(r + 1) * n_cols]
            needs = []
            for i, g in enumerate(row):
                is_full = full_last and r == n_rows - 1 and len(row) == 1
                needs.append(statement_card_h(g, full_w if is_full else card_w,
                                              size_pt, spacing))
            row_hs.append(max(needs or [0.42]) + 2 * pad)
        if sum(row_hs) + gap_min * (n_rows - 1) <= avail:
            break
    return row_hs, size_pt, spacing, pad


STMT_HEAD_MAX = 16       # 卡头长度上限：首句 ≤16 字才提成卡头，更长就留在正文流里
STMT_HEAD_ROW = 0.33     # 卡头行高（11pt 加粗行 + 图标 + 间隙）——渲染与算高同口径


def statement_group_head(group: str) -> tuple[str, str]:
    """卡组 → (卡头, 其余正文)。

    首句（≤16 字）提为**卡头标题行**——每张卡有一个可扫读的组标题，段落流才有
    章法（2026-10 用户反馈 statement"大量文字堆叠没章法"后的骨架升级）。
    首句过长/无法切分时返回 ("", 原文)，整组仍走旧的关键句流。渲染与
    `_fit_rows` 必须走同一个函数，保证算高与绘制同口径。
    """
    head, rest = split_claim(group or "")
    plain = head.replace("**", "").strip()
    if rest and 8 <= len(plain) <= STMT_HEAD_MAX:
        return plain, rest
    return "", (group or "").strip()


def statement_card_items(group: str) -> tuple[str, list[str], bool]:
    """卡组 → (卡头, 条目列表, 是否条目式)。

    卡内 `**短语**：` 小节 ≥2 时进入**条目式**：每个小节独立成条（色点 + 一块），
    卡内从"一段到底"变成可扫读的清单 —— 这是 statement 卡内的章法来源。
    条目文本保留 `**短语**：` 标记（text_box 渲染成加粗 + 强调色）。
    """
    head, rest = statement_group_head(group)
    items = []
    for phrase, body in split_sections(rest):
        text = (f"**{phrase}：**{body}" if phrase else body).strip()
        if text:
            items.append(text)
    return head, items, len(items) >= 2


def statement_card_h(group: str, card_w: float, size_pt: float,
                     spacing: float) -> float:
    """statement 单卡的高度（渲染与 `_fit_rows` 同一口径，宽度传**卡片宽**）。

    宽度必须按渲染实际用的文本区算 —— 旧实现传 `card_w-0.60`、绘制却用
    `c_w-0.41`，每行少算约 2 字，卡片被估高一整行，页面上就是卡片下方空一条。
    """
    head, items, list_mode = statement_card_items(group)
    if list_mode:
        inner_w = card_w - 0.41          # 色点 0.13 + 右侧内边距 0.14 + 余量
    elif head:
        inner_w = card_w - 0.28          # 卡头行缩进 0.14，两侧内边距
    else:
        inner_w = card_w - 0.59          # 图标 0.26 + 间距 0.09 + 左边距 0.10 + 右边距 0.14
    h = STMT_HEAD_ROW if head else 0.0
    if list_mode:
        h += sum(_para_h(it, inner_w, size_pt, spacing) for it in items)
        h += 0.05 * (len(items) - 1)
    else:
        h += card_text_h(items[0] if items else "", inner_w, size_pt, spacing,
                         claim=not head)
    return max(0.42, h)


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


# 三栏卡的竖向几何（英寸）。render_compare_3col 和 compare_desc_budget 共用。
COMPARE_GAP = 0.28           # 栏间距
COMPARE_GEOM = {
    "top_gap": 0.18,     # 页头之下、卡片上沿之间
    "band_top": 1.34,    # 卡片上沿 → 正文带（胶囊之下）
    "mb_h": 0.78,        # 底部比喻条高
    "mb_gap": 0.14,      # 卡片下沿 → 比喻条
    "band_pad": 0.12,    # 正文带 → 比喻条之间的余量
    "pt_h": 0.26,        # 每条要点占高
    "pt_gap": 0.16,      # 描述 → 要点之间
}


COMPARE_ICON_W = 0.52              # 三栏卡块头图标尺寸
COMPARE_ICON_RIGHT = 0.74          # 图标左沿距栏右沿


def compare_label_w(col_w: float, has_icon: bool) -> float:
    """三栏卡「栏名」可用的宽度：居中排，所以两侧都要给右上角图标让位。"""
    if not has_icon:
        return col_w - 0.5
    return max(0.8, col_w - 2 * (COMPARE_ICON_RIGHT + ICON_TITLE_GAP))


def compare_desc_box(summary: str | None, n_points: int,
                     has_callout: bool = False, aux_h: float | None = None,
                     n_cols: int = 3) -> tuple[float, float, float]:
    """对比卡「描述」的可用空间：(文本宽, 描述可用高度, 要点占高)。

    描述能写多少字取决于这几件事 —— 栏数、页头结论段长短、要点条数、辅区占位。
    渲染（挑档）与 qc（算容量）共用本函数，别在别处另写一套。
    """
    g = COMPARE_GEOM
    fr = card_frame(summary, has_callout, COMPARE_MIN_W, BODY_BOTTOM,
                    top_pad=g["top_gap"], aux_h=aux_h)
    n_cols = max(2, n_cols)
    cw = (fr["w"] - (n_cols - 1) * COMPARE_GAP) / n_cols
    ch = fr["bottom"] - fr["top"]
    band = ch - g["band_top"] - g["mb_h"] - g["mb_gap"] - g["band_pad"]
    used = g["pt_gap"] + min(n_points, 4) * g["pt_h"] if n_points else 0.0
    return cw - 0.56, max(0.0, band - used), used


def render_compare_3col(slide, m):
    from theme import ACCENT_CYCLE
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


def numbered_geom(m) -> dict:
    """编号卡网格几何（渲染与 qc 共用同一套数字）。

    4 张卡（2 列）让出右栏后每卡还有 4.1in → 左右分置；
    5-6 张卡（3 列）宽度敏感 → 辅区退回底部横条。
    """
    n = len(m.cards)
    # 4 张 → 2 列（每卡最宽）；5-6 张 → 3 列；7-8 张 → 4 列（窄卡，见下面 has_icon）
    cols = 2 if n == 4 else (4 if n >= 7 else 3)
    rows = math.ceil(n / cols)
    gap = 0.24
    co = getattr(m, "callout", None) or []
    # 用 `_MIN_CARD_W`（3.60）而不是"卡片刚好塞得下"的宽度来决定辅区落位：
    # 4 张（2 列）7.44in ≤ 右栏可用 8.46in → 右栏；5-6 张（3 列）11.28in、8 张（4 列）15.12in
    # 都超过 → 退回底部横条，卡片保持 3.9in 宽。
    # （曾经这里被写成 cols*2.60，5-6 张卡悄悄变成右栏、卡片缩到 2.66in —— 隐性漂移。）
    fr = card_frame(m.summary, bool(co), cols * _MIN_CARD_W + (cols - 1) * gap,
                    BODY_BOTTOM, aux_h=callout_height(co))
    cw = (fr["w"] - (cols - 1) * gap) / cols
    ch = (fr["bottom"] - fr["top"] - (rows - 1) * gap) / rows
    # 4 列窄卡（≈2.9in）**不画块头图标**：图标要占 0.80in，让出去标题只剩 6 字，
    # 还不如把宽度还给标题与描述（一块一个图标的规则在窄卡上让位于可读性）。
    has_icon = cols <= 3
    return {"fr": fr, "cols": cols, "rows": rows, "gap": gap, "cw": cw, "ch": ch,
            "x": fr["x"], "top": fr["top"], "has_icon": has_icon,
            "desc_w": cw - 0.60, "desc_h": ch - 1.15}


def render_numbered_cards(slide, m):
    from theme import ACCENT_CYCLE
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


def _layer_pill_ws(layer) -> list[float]:
    """本层右端标签胶囊的宽度（按中英实际字宽估算，英文按窄字计）。"""
    return [sum(0.13 if ord(c) > 127 else 0.085 for c in t) + 0.3
            for t in layer.tags[:3]]


def _layer_tags_w(layer) -> float:
    """本层右端标签组占多宽。挑档与绘制共用同一口径，别各算一套。"""
    ws = _layer_pill_ws(layer)
    return (sum(ws) + 0.1 * (len(ws) - 1)) if ws else 0.0


def layered_geom(m) -> dict:
    """层级堆栈几何（渲染与 qc 共用同一套数字）。

    辅区放**右栏**（横向条形宽度富余、高度吃紧）；描述可用宽按最窄的一层算，
    保证全页统一一档字号。
    """
    n = len(m.layers)
    # 层数多就把间距压紧（3~12 层都放得下）——**容量不够是压排版，不是删层**
    gap = 0.14 if n <= 6 else (0.10 if n <= 9 else 0.07)
    co = getattr(m, "callout", None) or []
    fr = card_frame(m.summary, bool(co), LAYER_MIN_W, BODY_BOTTOM, top_pad=0.18,
                    aux_h=callout_height(co))
    top = fr["top"]
    bar_h = (fr["bottom"] - top - (n - 1) * gap) / n
    label_w, x0 = 2.5, 1.35
    bar_w = fr["x"] + fr["w"] - x0
    _dw = bar_w - label_w - 0.3
    desc_w = min((_dw - (_layer_tags_w(l) + 0.3 if _layer_tags_w(l) else 0.15)
                  for l in m.layers), default=_dw - 0.15)
    return {"fr": fr, "gap": gap, "top": top, "bar_h": bar_h, "bar_w": bar_w,
            "x0": x0, "label_w": label_w, "desc_w": desc_w, "desc_h": bar_h - 0.08}


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


def _strip_step_no(label: str) -> str:
    """去掉标签前导序号（①/1./1、），渲染时统一改用数字徽章。"""
    s = label.strip()
    if s and 0x2460 <= ord(s[0]) <= 0x2473:
        return s[1:].lstrip(" .、:：")
    i = 0
    while i < len(s) and s[i].isdigit():
        i += 1
    if 0 < i < len(s) and s[i] in ".、:：":
        return s[i + 1:].lstrip()
    return s


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
    from theme import ACCENT_CYCLE
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
FLOW_NODE_W, FLOW_NODE_H = 1.72, 1.15
FLOW_CYCLE_DESC_W, FLOW_CYCLE_DESC_H = 1.44, 0.48
FLOW_CYCLE_DESC_SIZE = 7.5
FLOW_CYCLE_DESC_SPACING = 1.15
FLOW_CYCLE_STEPS = ((FLOW_CYCLE_DESC_SIZE, FLOW_CYCLE_DESC_SPACING),)


def flow_cycle_desc_box() -> tuple[float, float, tuple]:
    """环形流程节点描述框 (宽, 高, 字号档) —— schema 动态容量校验与 qc 共用。"""
    return FLOW_CYCLE_DESC_W, FLOW_CYCLE_DESC_H, FLOW_CYCLE_STEPS


FLOW_ARROW_W = 0.46
FLOW_ARROW_GAP = 0.16
FLOW_WRAP_AT = 6          # 超过这个步数折成两行（7~12 步一页放下，不用拆页/删步骤）
FLOW_ROW_GAP = 0.34       # 两行流程之间的折返通道（细线 + 向下箭头）


def flow_linear_geom(m) -> dict:
    """线性流程几何（渲染与 qc 共用）。

    步数少（≤4）宽度富余 → 辅区放右栏；步数多（≥5）每步至少 1.3in，退回底部横条。
    **步数超过 FLOW_WRAP_AT 就折成两行**（第一行左→右，折返后第二行继续左→右）。

    为什么必须支持折行：以前 "flow 只有 3~6 步" 是 schema 硬上限，于是出现
    「8 步的流程一页放不下 → 拆两页就丢整体感 / 删掉两步就丢内容」的死结，
    LLM 最后选了**压缩总页数、丢掉不重要的页** —— 与"研讨型不许丢输入内容"直接冲突。
    现在 7~12 步一页就装下：两行 × 每行 4 步，折返箭头标出阅读顺序。
    """
    n = len(m.steps)
    gap = FLOW_ARROW_W + FLOW_ARROW_GAP
    rows = 2 if n > FLOW_WRAP_AT else 1
    per = math.ceil(n / rows)
    co = getattr(m, "callout", None) or []
    fr = card_frame(m.summary, bool(co), per * _MIN_FLOW_W + (per - 1) * gap,
                    BODY_BOTTOM, top_pad=0.08, aux_h=callout_height(co))
    bw = (fr["w"] - (per - 1) * gap) / per
    total_h = fr["bottom"] - fr["top"]
    row_gap = 0.0 if rows == 1 else FLOW_ROW_GAP
    bh = total_h if rows == 1 else (total_h - row_gap) / rows
    return {"fr": fr, "n": n, "rows": rows, "per": per, "gap": gap,
            "arrow_w": FLOW_ARROW_W, "bw": bw, "bh": bh, "by": fr["top"],
            "row_gap": row_gap,
            "desc_w": bw - 0.40, "desc_h": bh - 1.42}


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
    from theme import ACCENT_CYCLE
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


def hierarchy_metrics(n_nodes: int, descs: list[str], summary: str | None,
                      has_footer: bool = False, has_callout: bool = False,
                      aux_h: float | None = None) -> dict:
    """包含型层级图的全部几何（渲染与 qc 共用同一套数字）。

    卡片高度贴着描述文字走、**不拉满页面**（拉满就是空框）；字号先按可用高度挑档
    （描述短就往上走到 10pt，长就往下降到 8.5pt），再用该档算卡片高。
    """
    gap = 0.32
    fr = card_frame(summary, has_callout, HIER_MIN_W, BODY_BOTTOM, top_pad=0.02,
                    aux_h=aux_h)
    cw = (fr["w"] - (n_nodes - 1) * gap) / n_nodes if n_nodes else fr["w"]
    tw = cw - 0.28
    cy = fr["top"] + 0.5 + 0.62
    avail = max(1.0, fr["bottom"] - cy)
    size, spacing = pick_step(
        lambda sz, sp: max((_lines_in(d, tw, sz) * sz / 72 * sp for d in descs),
                           default=0.0), max(0.4, avail - 1.12), HIER_STEPS)
    lines = max((_lines_in(d, tw, size) for d in descs), default=1)
    ch = max(1.30, 0.92 + lines * size / 72 * spacing + 0.20)
    # 卡片吃留白：子模块卡窄，贴内容算出来的高常常远小于可用高度 ——
    # 允许把卡片**长高最多 HIER_GROW_MAX**，描述在卡内垂直居中，就不会出现
    # "框在、内容全堆在顶部、下半截空着"（P31 那种空框的病根是顶部锚定 + 无限拉长）。
    raw_slack = max(0.0, avail - ch)
    ch = min(avail, ch + min(HIER_GROW_MAX, raw_slack * 0.75))
    return {"fr": fr, "n": n_nodes, "gap": gap, "cw": cw, "ch": ch, "cy": cy,
            "size": size, "spacing": spacing,
            "slack": max(0.0, fr["bottom"] - (cy + ch)),
            "desc_w": tw, "desc_h": max(0.3, ch - 1.04)}


def hierarchy_geom(m) -> dict:
    """包含型层级图：按页面模型算几何（qc 与渲染共用）。"""
    co = getattr(m, "callout", None) or []
    return hierarchy_metrics(len(m.nodes), [nd.desc for nd in m.nodes], m.summary,
                             bool(getattr(m, "footer_note", None)), bool(co),
                             callout_height(co))


def _render_hierarchy_contain(slide, m, y):
    """包含（树形）：顶部根节点实心条，向下连线分流到 N 个子模块卡片。"""
    from theme import ACCENT_CYCLE
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
    from theme import ACCENT_CYCLE
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


# ---------------- 双模块并排页（panel_2col） ----------------
#
# 用户的三条硬要求直接落在这里：
#   1. 一页两块内容 → **左右分置**（不是上下堆叠，堆叠会在下半页留出大片空白）；
#   2. 每个模块**独立背景框**（白卡片 + 顶部色条 + 自己的边框）；
#   3. 每个模块**顶部一段总结陈述段**（浅色块，与页面级结论段同一语言）。
PANEL_GAP = 0.30         # 两个模块之间的横向间距
PANEL_TOP_PAD = 0.06     # 页头之下、模块上沿之间
PANEL_PAD = 0.22         # 模块卡片内的左右内边距
PANEL_HEAD_TOP = 0.16    # 模块标题行距框顶
PANEL_HEAD_H = 0.40      # 模块标题行高
PANEL_SUM_SIZE = 11.5    # 模块顶部总结陈述段字号
PANEL_PT_SIZE = 11.0     # 模块内要点正文字号
PANEL_TBL_HEAD = 0.34    # 模块内迷你表表头行高
PANEL_TBL_ROW = 0.32     # 模块内迷你表表体行高（紧凑）
PANEL_TBL_CELL = 9.5     # 模块内迷你表单元格字号
# 模块内要点：起点与全局一致；空间真的不够时继续降（模块框里空间最紧）。
# **低于 PANEL_PT_SIZE 说明这一模块的内容已经超框**（渲染器只能压缩才画得进去），
# `panel_slack` 会把这件事报给 --qc，让作者去删要点/减行，而不是让框被撑破。
PANEL_PT_STEPS = FIT_STEPS + ((9.5, 1.24), (9.0, 1.18), (8.5, 1.12), (8.0, 1.08))
PANEL_PT_SIZE = 11.0        # 要点"正常"字号；低于它 = 被压缩过
PANEL_PT_FLOOR = 9.5        # 要点可读下限：渲染可以再往下压，但 qc 会点名（该删内容了）
PANEL_TBL_ROW_MIN = 0.16    # 迷你表最矮行高（再矮字会压线，此时由 qc 点名删行）
PANEL_TBL_CELL_STEPS = ((9.5, 1.16), (9.0, 1.12), (8.5, 1.08), (8.0, 1.04))


def panel_area(summary: str | None, has_callout: bool, has_footer: bool = False,
               aux_h: float | None = None) -> tuple[float, float]:
    """双模块页内容区的 (上沿, 下沿)。渲染与 qc 共用同一个函数（同一套 card_frame）。

    `has_footer` 已不再影响内容带（页脚条在 BODY_BOTTOM 之下），仅为兼容保留。
    """
    fr = card_frame(summary, has_callout, PANEL_MIN_W, BODY_BOTTOM,
                    top_pad=PANEL_TOP_PAD, aux_h=aux_h)
    return fr["top"], fr["bottom"] - 0.06


def _panel_points_need(pts, width: float, size_pt: float, spacing: float,
                       gap: float = 0.14) -> float:
    """要点列表在给定字号下需要多高（渲染与 qc 同一口径）。"""
    if not pts:
        return 0.0
    return (sum(max(0.24, _lines_in(t, width, size_pt) * size_pt / 72 * spacing)
                for t in pts) + gap * (len(pts) - 1))


def panel_points_layout(pts, width: float, budget: float, gap: float = 0.14) -> dict:
    """要点列表在 budget 高度内的排版（字号 / 行距 / 各行高 / 间距）—— 渲染与 qc 共用。

    连最小档都放不下时继续往下缩（下限 7.5pt，与底部辅区同口径）。**绝不越框**：
    越框的部分会直接压在下面的迷你明细表上。
    """
    pts = list(pts or [])
    if not pts:
        return {"size": PANEL_PT_SIZE, "spacing": 1.30, "hs": [], "gap": gap}
    budget = max(0.0, budget)

    def _need(sz, sp):
        return _panel_points_need(pts, width, sz, sp, gap)

    size, spacing = pick_step(_need, budget, PANEL_PT_STEPS)
    while size > 7.5 and _need(size, spacing) > budget:
        size -= 0.5
        spacing = max(1.05, spacing - 0.04)
    hs = [max(0.24, _lines_in(t, width, size) * size / 72 * spacing) for t in pts]
    slack = budget - sum(hs) - gap * (len(pts) - 1)
    if slack > 0 and len(pts) > 1:
        gap = min(0.34, gap + slack / (len(pts) - 1))
    return {"size": size, "spacing": spacing, "hs": hs, "gap": gap}


def panel_content_geom(summary: str, top: float, ph: float, pw: float, panel) -> dict:
    """模块内「要点 + 迷你表」的高度切分 —— **唯一判据**（渲染与 qc 共用）。

    历史的坑（用户复现的 P3「右侧框里的表格被挤出期望区域」）：
    旧实现让要点先画、表再按"剩余空间"落位，而 `pick_step` 在最小档也放不下时
    仍然返回最小档（绝不失败），于是要点越界把表一路往下顶；`_panel_table` 又用
    `max(0.18, …)` 给行高兜底、且 `ty = y + max(0, …)` 不会回收，表就画到框外去了。

    现在改成**从下往上分配**：
      1. 迷你表按内容要一个自然高度，但最多只拿 `avail - 要点最小高度 - 间距`；
      2. 表**底部锚定**在内容区下沿，行高 / 表头高 / 字号都按拿到的 `tbl_h` 反算，
         `head_h + row_h×n ≡ tbl_h`，**数学上不可能越出模块框**；
      3. 要点只在上方剩余区间里排（真的排不下就继续降字号，由 `panel_pressure` 报给 qc）。
    """
    gap = 0.14
    iw = pw - 2 * PANEL_PAD - 0.34
    sy = top + 0.62
    sh = _panel_sum_h(summary)
    cy = sy + sh + 0.16
    cy_bottom = top + ph - 0.18
    avail = max(0.0, cy_bottom - cy)

    pts = list(getattr(panel, "points", None) or [])
    rows = list(getattr(panel, "rows", None) or [])
    tbl_rows = len(rows)
    tbl_natural = (PANEL_TBL_HEAD + PANEL_TBL_ROW * tbl_rows) if tbl_rows else 0.0
    # 要点在"正常字号"下的需要量（qc 拿它算"还剩多少"）
    pt_need_natural = _panel_points_need(pts, iw, PANEL_PT_SIZE, 1.30, gap)
    pt_need_floor = _panel_points_need(pts, iw, PANEL_PT_STEPS[-1][0],
                                       PANEL_PT_STEPS[-1][1], gap)

    if tbl_rows and pts:
        tbl_h = min(tbl_natural, max(0.0, avail - pt_need_floor - gap))
    elif tbl_rows:
        tbl_h = min(tbl_natural, avail)
    else:
        tbl_h = 0.0

    if tbl_rows:
        head_h = PANEL_TBL_HEAD
        if tbl_h < head_h + PANEL_TBL_ROW_MIN * tbl_rows:
            head_h = max(0.18, tbl_h * 0.28)          # 空间极紧时表头一起压
        row_h = max(0.0, (tbl_h - head_h) / tbl_rows)
    else:
        head_h = row_h = 0.0

    tbl_y = cy_bottom - tbl_h
    pt_bottom = (tbl_y - gap) if tbl_rows else cy_bottom
    pt_budget = max(0.0, pt_bottom - cy)
    pt = panel_points_layout(pts, iw, pt_budget, gap)

    # 迷你表字号：行高必须容得下单元格文字的自然高度，否则 PowerPoint 打开时
    # 会把行撑高（表就跑到框外、压住别的东西）—— 这是「表被挤出」的第二个成因。
    cell_size, _csp = PANEL_TBL_CELL_STEPS[-1]
    if tbl_rows:
        _cw = col_widths(list(getattr(panel, "headers", None) or []), rows,
                         pw - 2 * PANEL_PAD, min_w=0.62)
        for _sz, _sp in PANEL_TBL_CELL_STEPS:
            if all(_lines_in(v, max(0.40, _cw[c] - 0.14), _sz) * _sz / 72 * _sp
                   <= row_h - 0.02
                   for row in rows for c, v in enumerate(row) if c < len(_cw)):
                cell_size, _csp = _sz, _sp
                break
    return {
        "cy": cy, "cy_bottom": cy_bottom, "avail": avail, "iw": iw, "gap": gap,
        "pt_need_natural": pt_need_natural, "pt_need_floor": pt_need_floor,
        "pt_budget": pt_budget, "pt_bottom": pt_bottom,
        "pt_size": pt["size"], "pt_spacing": pt["spacing"], "pt_hs": pt["hs"],
        "pt_gap": pt["gap"],
        "tbl_y": tbl_y, "tbl_h": tbl_h, "tbl_natural": tbl_natural,
        "head_h": head_h, "row_h": row_h, "cell_size": cell_size,
        "sum_h": sh, "sy": sy,
    }


def panel_slack(m) -> tuple[float, float]:
    """双模块页：返回 (两个模块里最空的那个还剩多少 in, 模块高度)。

    **负值 = 这一页的模块内容按正常字号/行高已经装不下**（渲染器只能压字号、
    压行高才塞得进去）。两块内容都少时是正值，此时该把相邻页内容并进来。
    """
    notes = getattr(m, "callout", None) or []
    top, bottom = panel_area(m.summary, bool(notes), bool(getattr(m, "footer_note", None)),
                             callout_height(notes))
    ph = bottom - top
    pw = (12.23 - PANEL_GAP) / 2
    slack_min = ph
    for p in m.panels:
        g = panel_content_geom(p.summary, top, ph, pw, p)
        need = g["pt_need_natural"] + (g["gap"] if p.rows and p.points else 0.0) \
            + g["tbl_natural"]
        slack_min = min(slack_min, g["avail"] - need)
    return slack_min, ph


def panel_pressure(m) -> tuple[float, float, float, float]:
    """双模块页：返回 (最紧模块的超框量, 该模块里要点实际字号, 迷你表实际行高, 模块高)。

    「超框」的正负看 `panel_pressure()[0]`；但**渲染器已经靠压缩兜住了**，
    观感是否还能接受要看后两个数：要点被压到 `PANEL_PT_FLOOR` 以下、
    或迷你表行高低于 `PANEL_TBL_ROW_MIN + 0.02`，就该让作者删要点/减行（qc 点名）。
    """
    notes = getattr(m, "callout", None) or []
    top, bottom = panel_area(m.summary, bool(notes), bool(getattr(m, "footer_note", None)),
                             callout_height(notes))
    ph = bottom - top
    pw = (12.23 - PANEL_GAP) / 2
    worst = (-9.9, PANEL_PT_SIZE, PANEL_TBL_ROW, ph)
    for p in m.panels:
        g = panel_content_geom(p.summary, top, ph, pw, p)
        need = g["pt_need_natural"] + (g["gap"] if p.rows and p.points else 0.0) \
            + g["tbl_natural"]
        over = need - g["avail"]
        if over > worst[0]:
            worst = (over, g["pt_size"],
                     g["row_h"] if p.rows else PANEL_TBL_ROW, ph)
    return worst


def _panel_sum_h(summary: str) -> float:
    """模块顶部总结陈述段占的高度（浅色底块自适应，不许文字溢出）。"""
    sw = (12.23 - PANEL_GAP) / 2 - 2 * PANEL_PAD
    lines = _lines_in(summary, sw - 0.28, PANEL_SUM_SIZE)
    return max(0.46, lines * PANEL_SUM_SIZE / 72 * 1.32 + 0.22)


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
    from theme import ACCENT_CYCLE
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

GRID_GAP = 0.28              # 格子间距
GRID_TOP_PAD = 0.12          # 标题下方的额外间距
GRID_DESC_TOP = 1.15         # 说明段相对格子上沿的起点（标题 0.35 + 大数字 0.5）


def grid_geom(m) -> dict:
    """2x2 网格几何（渲染与 qc 共用）：格宽/格高 + 说明段文本框与字号。

    原来渲染里这些数字全是散装常量（`cell_h - 1.2`、`cell_w - 0.5`、写死 11pt），
    qc 只能查"值非空"，格子写满后被撑破也没人管。现在收进几何函数，
    qc 按 desc_w/desc_h 核对容量。
    """
    top = header_y(getattr(m, "summary", None)) + GRID_TOP_PAD
    cell_w = (CONTENT_W - GRID_GAP) / 2
    cell_h = (BODY_BOTTOM - top - GRID_GAP) / 2
    desc_h = max(0.30, cell_h - GRID_DESC_TOP - 0.10)
    descs = [c.desc for c in (m.q_top_left, m.q_top_right,
                              m.q_bottom_left, m.q_bottom_right) if c.desc]
    size, spacing = fit_size(descs, cell_w - 0.5, desc_h, FIT_STEPS)
    return {"top": top, "cell_w": cell_w, "cell_h": cell_h,
            "desc_w": cell_w - 0.5, "desc_h": desc_h,
            "desc_size": size, "desc_spacing": spacing}


def render_grid_2x2(slide, m):
    """2x2 网格：强制 2×2 排列的四象限/矩阵布局。"""
    from theme import ACCENT_CYCLE

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

KG_GAP = 0.14
KG_TOP_PAD = 0.15
KG_GROW_MAX = 1.35        # 卡片相对内容需要的高度最多长到 1.35 倍（再多就是空框）
_KG_TITLE_ROW = 0.36      # 组标题行（13pt 加粗 + 图标）


def _kg_card_h(grp, card_w: float, size_pt: float, spacing: float) -> float:
    """key_groups 单卡高度：组标题行 + 要点清单（渲染与挑档同一口径）。"""
    text_w = card_w - 0.46
    h = _KG_TITLE_ROW + 0.12
    for pt in grp.points:
        h += _para_h(pt, text_w, size_pt, spacing) + 0.05
    return h


def key_groups_geom(m):
    """key_groups 几何（渲染与 qc 共用）：返回逐行布局与要点容量。

    两组一行的等网格在**奇数条目**时右下角空一整个卡位（3 组 = 空 1/4 页）。
    末行只剩 1 组时让那张卡**跨整行**：宽了一倍，一行装得下更多字，行高也随之
    变矮，多出来的高度回到上面几行去。行高按各自宽度算，不再一刀切均分。
    """
    aux = getattr(m, "callout", None)
    bottom = BODY_BOTTOM - (callout_height(aux) + 0.12 if aux else 0.0)
    top = header_y(m.summary) + KG_TOP_PAD
    groups = list(m.groups) or []
    n = max(1, len(groups))
    n_cols = 2 if n > 1 else 1
    n_rows = max(1, math.ceil(n / n_cols))
    card_w = (CONTENT_W - (n_cols - 1) * KG_GAP) / n_cols
    rows = []
    for r in range(n_rows):
        idxs = list(range(r * n_cols, min(n, (r + 1) * n_cols)))
        single = n_cols > 1 and r == n_rows - 1 and len(idxs) == 1
        rows.append({"idxs": idxs, "w": (CONTENT_W if single else card_w),
                     "single": single})
    avail = bottom - top - (n_rows - 1) * KG_GAP
    size, spacing, row_hs = FIT_STEPS[-1][0], FIT_STEPS[-1][1], []
    for size, spacing in FIT_STEPS:
        row_hs = [max(_kg_card_h(groups[i], row["w"], size, spacing)
                      for i in row["idxs"]) for row in rows]
        if sum(row_hs) + KG_GAP * (n_rows - 1) <= avail:
            break
    else:
        # 最低档仍放不下：按可用高度均分（文字会压框，由 --qc 点名）
        row_hs = [max(0.6, avail / n_rows)] * n_rows
    # 行高贴合内容后会沉到底部留白（3 组时下半页空 1.3in）—— 把余量**按比例
    # 撑高卡片（上限 KG_GROW_MAX），还剩的就上下均分、整块居中。
    used = sum(row_hs) + KG_GAP * (n_rows - 1)
    top_off = 0.0
    if used < avail - 0.02:
        room = [h * (KG_GROW_MAX - 1.0) for h in row_hs]
        total_room = sum(room)
        if total_room > 0:
            add = min(avail - used, total_room)
            row_hs = [h + r * (add / total_room) for h, r in zip(row_hs, room)]
            used += add
        top_off = max(0.0, (avail - used) / 2)
    cards, y = [], top + top_off
    for row, rh in zip(rows, row_hs):
        xs = [MARGIN_L]
        if len(row["idxs"]) > 1:
            xs = [MARGIN_L + k * (row["w"] + KG_GAP)
                  for k in range(len(row["idxs"]))]
        for k, gi in enumerate(row["idxs"]):
            cards.append({"i": gi, "x": xs[k], "y": y, "w": row["w"], "h": rh,
                          "text_w": row["w"] - 0.46,
                          "pts_h": max(0.20, rh - _KG_TITLE_ROW - 0.12)})
        y += rh + KG_GAP
    first = cards[0] if cards else {"w": card_w, "h": 1.0}
    return {"top": top, "bottom": bottom, "n_cols": n_cols, "card_w": card_w,
            "card_h": first["h"], "text_w": first["text_w"],
            "pts_h": first["pts_h"], "size": size, "spacing": spacing,
            "rows": rows, "cards": cards, "full_last": any(r["single"] for r in rows)}


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

TL_AXIS_HALF = 0.28        # 横排：中轴带到卡片的间隙（连接线画在这里）
TL_TOP_PAD = 0.22
TL_VERTICAL_AT = 7         # 节点 ≥7 改竖排：横排每卡只剩 <1.4in，文字放不下
TL_VGAP = 0.08             # 竖排：行间距
TL_TIME_W = 1.05           # 竖排：左侧时间列宽
TL_AXIS_X = 1.82           # 竖排：中轴 x
TL_CARD_X = 2.02           # 竖排：卡片左沿


def timeline_geom(m):
    """timeline 几何（渲染与 qc 共用一张图）：返回模式 / 卡宽 / 描述容量。

    **节点多于 6 个就改竖排**：横排是「卡片沿中轴上下交替、宽度按节点数均分」，
    12 个节点时每卡只有 0.85in、描述区 0.49in —— 120 字的 schema 上限在那里
    需要 5.6in 高，必然横穿整页。竖排一行一个节点（时间在左、卡片在右），
    12 行仍能每行写下一整句，容量反而接近横排的 3 倍。
    """
    aux = getattr(m, "callout", None)
    bottom = BODY_BOTTOM - (callout_height(aux) + 0.12 if aux else 0.0)
    top = header_y(m.summary) + TL_TOP_PAD
    n = max(1, len(m.milestones))
    if n >= TL_VERTICAL_AT:
        row_h = (bottom - top - (n - 1) * TL_VGAP) / n
        card_w = (MARGIN_L + CONTENT_W) - TL_CARD_X
        desc_w = card_w - 0.36
        desc_h = max(0.24, row_h - 0.20 - 0.12)      # 扣掉 label 行与上下内边距
        size, spacing = fit_size([ms.desc for ms in m.milestones],
                                 desc_w, desc_h, FIT_STEPS)
        return {"mode": "vertical", "top": top, "bottom": bottom, "n": n,
                "gap": TL_VGAP, "row_h": row_h, "card_x": TL_CARD_X,
                "card_w": card_w, "axis_x": TL_AXIS_X, "time_w": TL_TIME_W,
                "desc_w": desc_w, "desc_h": desc_h,
                "size": size, "spacing": spacing}
    gap = 0.18
    card_w = (CONTENT_W - (n - 1) * gap) / n
    half_h = (bottom - top - 2 * TL_AXIS_HALF) / 2
    return {"mode": "horizontal", "top": top, "bottom": bottom, "n": n,
            "gap": gap, "card_w": card_w,
            "half_h": half_h, "desc_w": card_w - 0.36,
            "desc_h": max(0.35, half_h - 0.86), "axis_y": (top + bottom) / 2}


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

FUNNEL_DESC_W = 4.30      # 右侧说明栏宽（英寸）


def funnel_geom(m):
    """funnel 几何（渲染与 qc 共用）：返回条高/逐层宽度/说明栏容量。"""
    aux = getattr(m, "callout", None)
    bottom = BODY_BOTTOM - (callout_height(aux) + 0.12 if aux else 0.0)
    top = header_y(m.summary) + 0.18
    n = max(1, len(m.stages))
    gap = 0.14
    bar_h = (bottom - top - (n - 1) * gap) / n
    zone_w = CONTENT_W - FUNNEL_DESC_W - 0.35
    w_top, w_bot = zone_w, zone_w * 0.40
    widths = [w_top + (w_bot - w_top) * (i / (n - 1)) for i in range(n)] \
        if n > 1 else [w_top]
    return {"top": top, "bottom": bottom, "bar_h": bar_h, "gap": gap,
            "widths": widths, "zone_w": zone_w,
            "desc_x": MARGIN_L + zone_w + 0.35, "desc_w": FUNNEL_DESC_W,
            "desc_h": max(0.35, bar_h - 0.12)}


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

MATRIX_ROWLABEL_W = 1.45   # 左侧行标列宽
MATRIX_GAP = 0.12
MATRIX_COLHEAD_H = 0.36
MATRIX_TOP_PAD = 0.10
MATRIX_DESC_TOP = 0.58     # 格子说明起点：块头图标占 0.09–0.57，说明必须整段让开
MATRIX_BOTTOM_PAD = 0.10
# 矩阵格子最窄（4 行 × 4 列时约 2.2in），正文比其它版式多给两档
MATRIX_STEPS = FIT_STEPS + ((9.5, 1.18), (9.0, 1.12))


def matrix_geom(m):
    """matrix 几何（渲染与 qc 共用）：返回格子尺寸与描述容量。

    旧实现描述框从 +0.46 起、扣 0.72 的 chrome，4 行时只剩 0.37in（10pt 约 28 字），
    与 schema 的 90 字差了 3 倍。这里把列头/间距/内边距都收了一档，并允许降到 9pt。

    说明起点**必须在图标下沿之下**（0.58）：从 0.42 起时首行文字会和右上角图标
    撞在一起（3 行格子特别明显，图标压住了第一行末尾几个字）。
    """
    n_rows = max(1, len(m.rows))
    n_cols = max(1, len(m.col_labels))
    top = header_y(m.summary) + MATRIX_TOP_PAD
    cell_w = (CONTENT_W - MATRIX_ROWLABEL_W - (n_cols - 1) * MATRIX_GAP) / n_cols
    cell_h = (BODY_BOTTOM - top - MATRIX_COLHEAD_H - 0.08
              - (n_rows - 1) * MATRIX_GAP) / n_rows
    return {"top": top, "cell_w": cell_w, "cell_h": cell_h,
            "desc_w": cell_w - 0.40,
            "desc_h": max(0.24, cell_h - MATRIX_DESC_TOP - MATRIX_BOTTOM_PAD)}


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


# ---------------- 图示页（figure） ----------------
#
# 为什么不走 `slide.shapes.add_picture`：python-pptx 用 Pillow 读图片头判断格式与尺寸，
# 而 Pillow **不认识 EMF/WMF** —— Word / Visio / draw.io 导出的图恰恰多是 EMF。
# 把 EMF 转成 PNG 再嵌进去会糊；用户的硬要求是"原封不动放到 PPT 相应位置"。
# 所以这里手工挂 ImagePart：**字节原样**进 `ppt/media/`，PowerPoint 自己按矢量渲染。
FIG_EXT_VECTOR = (".emf", ".wmf")
_CTYPE_VECTOR = {"emf": "image/x-emf", "wmf": "image/x-wmf"}


def _blob_size_in(blob: bytes, ext: str) -> tuple[float, float] | None:
    """图片的物理尺寸（英寸）：EMF 读 `rclFrame`（0.01mm）、WMF 读 placeable 头，其余用 Pillow。"""
    import struct

    ext = ext.lower()
    try:
        if ext == ".emf" and len(blob) >= 40:
            _l, _t, _r, _b = struct.unpack_from("<4i", blob, 24)
            w, h = (_r - _l) / 100.0 / 25.4, (_b - _t) / 100.0 / 25.4
            if 0.05 < w < 80 and 0.05 < h < 80:
                return w, h
        elif ext == ".wmf" and len(blob) >= 22 and blob[:4] == b"\xd7\xcd\xc6\x9a":
            _l, _t, _r, _b = struct.unpack_from("<4h", blob, 6)
            w, h = (_r - _l) / 1440.0, (_b - _t) / 1440.0
            if 0.05 < w < 80 and 0.05 < h < 80:
                return w, h
        else:
            from PIL import Image as PILImage
            with PILImage.open(io.BytesIO(blob)) as im:
                dpi = im.info.get("dpi") or (96, 96)
                return im.width / float(dpi[0]), im.height / float(dpi[1])
    except Exception:                                   # noqa: BLE001 —— 尺寸拿不到不影响嵌入
        pass
    return None


def image_size_in(path: str) -> tuple[float, float] | None:
    """图片物理尺寸（英寸）；读不到返回 None（此时按给定框铺满）。"""
    try:
        with open(path, "rb") as f:
            blob = f.read()
    except OSError:
        return None
    return _blob_size_in(blob, os.path.splitext(path)[1]) if blob else None


def resolve_image_path(path: str) -> str:
    """DSL 里的 image 路径 → 磁盘真实路径。

    依次试「仓库根 → .pi → .pi/work → 当前目录」。为什么这么绕：`docx_in.py`
    把图抽到 `.pi/work/media/`，并在 Markdown 里写成 `media/NN-x.emf`（相对 md
    文件本身）—— 于是 DSL 里出现的就是 `media/xx.emf`，而它相对的是 `.pi/work`，
    不是仓库根、更不是 cwd。旧实现只试 `_BASE`(.pi) 和 cwd，从别的目录运行时会
    拼出 `.pi/.pi/work/media/...` 这种永远不存在的路径，图页直接变成"缺图"。
    """
    if os.path.isabs(path):
        return path
    bases = (os.path.dirname(_BASE), _BASE, os.path.join(_BASE, "work"), os.getcwd())
    cands = [os.path.join(b, path) for b in bases]
    if path.replace("\\", "/").startswith("media/"):
        cands.append(os.path.join(_BASE, "work", path))
    for c in cands:
        if os.path.isfile(c):
            return c
    return cands[0]


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


# ---------------- 图示页（figure）：一张原图 + 图注（+ 可选要点右栏） ----------------
FIGURE_MIN_W = 5.20       # 图区最窄宽度：比这更窄就该把要点挪到图下方
FIGURE_PAD = 0.20         # 图/图注与卡片边框之间的留白
FIGURE_CAPTION_H = 0.42   # 图注一行高


def figure_area(summary: str | None, has_caption: bool, has_points: bool
                ) -> tuple[dict, tuple[float, float, float, float]]:
    """图示页几何：返回 (card_frame, 图片矩形 x/y/w/h)。

    图的**放法**（等比 contain + 居中）只在 `render_figure` → `add_image` 里，
    qc 不参与几何核对（它只查"图在不在"和"这一页是不是只有一个容器"）。
    """
    fr = card_frame(summary, has_points, FIGURE_MIN_W, BODY_BOTTOM)
    cap_h = FIGURE_CAPTION_H if has_caption else 0.0
    card_h = fr["bottom"] - fr["top"]
    return fr, (fr["x"] + FIGURE_PAD, fr["top"] + FIGURE_PAD,
                max(0.6, fr["w"] - 2 * FIGURE_PAD),
                max(0.6, card_h - 2 * FIGURE_PAD - cap_h))


# ---------------- 右栏卡片栈：把可用高度**消化掉** ----------------
RAIL_PAD_MAX = 0.26      # 单卡上下内边距补量上限
RAIL_GAP_MAX = 0.40      # 卡片间距上限
RAIL_GROW_MAX = 1.45     # 卡片高度相对自然高度的最大放大倍数（再高就是空框）


def _fill_rail(need: list[float], avail: float, gap_min: float = 0.14
               ) -> tuple[list[float], float, float]:
    """把右栏卡片铺进 avail：先补内边距，再拉间距，再适度长高，最后**居中**。

    返回 (各卡高, 间距, 顶部偏移)。

    旧实现 pad 封在 0.20、gap 封在 0.30 就收手，余下的高度全沉在栏底 —— 4.5in
    的右栏只堆了 2in 卡片，下面 2in 是纯空气（figure / layered / panel 页都有）。
    现在分四步消化：消化不掉的上下均分，整栈垂直居中，而不是一条贴顶一条空底。
    """
    n = len(need)
    if not n:
        return [], gap_min, 0.0
    # **等高**：右栏是同一块信息的堆叠，卡片高低不齐比留白更难看。
    # 统一到最长的那个自然高度，再一起进入四步消化。
    hs = [max(need)] * n
    slack = avail - sum(hs) - gap_min * (n - 1)
    if slack <= 0:
        return hs, gap_min, 0.0
    per = min(RAIL_PAD_MAX, slack / n)              # 1) 卡片内边距
    hs = [h + per for h in hs]
    slack -= per * n
    gap = gap_min
    if n > 1 and slack > 0:                          # 2) 卡片间距
        gap = min(RAIL_GAP_MAX, gap_min + slack / (n - 1))
        slack -= (gap - gap_min) * (n - 1)
    if slack > 0.05:                                 # 3) 适度长高
        room = [h * (RAIL_GROW_MAX - 1.0) for h in hs]
        total_room = sum(room)
        if total_room > 0:
            add = min(slack, total_room)
            hs = [h + r * (add / total_room) for h, r in zip(hs, room)]
            slack -= add
    return hs, gap, max(0.0, slack / 2)              # 4) 余量上下均分


def point_rail_need(pts, w: float) -> list[float]:
    """figure 右栏每张要点卡的自然高度（渲染与 qc 同一口径）。"""
    return [max(0.72, _para_h(t, w - 0.78, 10.5, 1.25) + 0.34) for t in (pts or [])]


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
                 f".pi/renderer/docx_in.py 抽到 .pi/work/media/，EMF/WMF 会原样嵌入）",
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


# 表格页竖向几何（英寸）。render_table 和 table_slack 共用 —— qc 报的"还能放几行"
# 必须和渲染出来的行高一模一样，别再在别处另写一套。
TABLE_TOP_PAD = 0.15     # 页头之下、表格上沿之间
TABLE_BOTTOM = BODY_BOTTOM   # 内容区底线（页脚下移后与其它版式统一）
TABLE_HEAD_H = 0.50      # 表头行高
TABLE_ROW_H = 0.46       # 表体行高（紧凑，不随页面高度摊开）
TABLE_NOTE_H = 0.44      # 底部 note 条
SIDE_W = 3.55            # 右侧辅区宽度（有 side 时主表收窄，让一页装两个信息块）
SIDE_GAP = 0.22


def _cell_text_w(text: str, size_pt: float = 11.5) -> float:
    """单元格文字需要的宽度（英寸），与 _lines_in 同一套保守字宽。"""
    em = size_pt / 72
    return sum(em * 1.06 if ord(c) > 127 else em * 0.55 for c in (text or ""))


def col_widths(headers: list[str], rows: list[list[str]], total_w: float,
               min_w: float = 1.15) -> list[float]:
    """按各列内容的实际宽度分配列宽 —— 不再等分。

    固定等分是"表格大面积留白"的元凶：单元格只有 4 个字也占 3.18in。
    这里按每列最长内容 + 内边距分配，短列窄、长列宽，省下的横向空间还给内容。
    """
    n = len(headers)
    if not n:
        return []
    need = []
    for c in range(n):
        vals = [headers[c]] + [r[c] for r in rows if c < len(r)]
        widest = max((_cell_text_w(v) for v in vals), default=1.0)
        need.append(max(min_w, widest + 0.30))
    scale = total_w / sum(need)
    widths = [max(min_w, x * scale) for x in need]
    k = total_w / sum(widths)              # 抬到下限后可能超宽，再归一一次
    return [w * k for w in widths]


CALLOUT_H_MAX = 0.86      # 底部横条的最大高度（两条说明文字）
CALLOUT_H_MIN = 0.44      # 只有一行 label 时的最小高度
CALLOUT_H = CALLOUT_H_MAX  # 兼容旧引用（qc 等按最大高度做保守估算）
_CALLOUT_LABEL_H = 0.20   # label 行
_CALLOUT_VALUE_H = 0.27   # 大字数值行
_CALLOUT_SIZE = 8.5       # 说明文字字号（放不下会往下缩到 7.5）


def callout_height(notes, width: float = 12.23) -> float:
    """底部横条按**内容**自适应高度（CALLOUT_H_MIN ~ CALLOUT_H_MAX）。

    以前固定 0.86in：只写「标签 + 数值」的条目也照样占满一整条，白扔掉 0.2~0.4in
    页高 —— 这些高度本该给主内容。渲染与 qc 共用本函数（`card_frame(aux_h=...)`）。
    """
    if not notes:
        return 0.0
    n = len(notes)
    gap = 0.20
    tw = (width - (n - 1) * gap) / n - 0.32
    need = CALLOUT_H_MIN
    for cn in notes:
        off = 0.55 if cn.value else 0.28
        h = off + 0.08
        if cn.text:
            h += _lines_in(cn.text, tw, _CALLOUT_SIZE) * (_CALLOUT_SIZE / 72 * 1.15) + 0.04
        need = max(need, h)
    return min(CALLOUT_H_MAX, max(CALLOUT_H_MIN, need))


def _render_callout(slide, notes, bottom: float, height: float | None = None) -> float:
    """底部横排辅区：2-4 条关键数字 / 边界 / 待确认点。返回新的内容底线。

    卡片型版式原来整页只有"一组卡片"，信息量被容器锁死；callout 让它们多装一块信息。
    **高度按内容自适应**（`callout_height`），说明文字放不下时逐档缩到 7.5pt。
    """
    from theme import ACCENT_CYCLE
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
    from theme import ACCENT_CYCLE
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


# 表格单元格的降档序列：比正文多给两档（9.5 / 9.0 / 8.5）——
# 行数多的表只有字号继续往下走，行高才装得下文字；否则 PowerPoint 会把行撑高。
TABLE_CELL_STEPS = FIT_STEPS + ((9.5, 1.18), (9.0, 1.14), (8.5, 1.10))
CELL_V_PAD = 0.06        # 单元格上下内边距合计（渲染里显式设 0.02+0.02，再留 0.02 余量）


def table_row_need(text: str, width: float, size_pt: float, spacing: float) -> float:
    """单元格文字在给定宽度/字号下的**自然高度**（含上下内边距）。

    **行高必须 ≥ 它**。python-pptx 写进去的行高只是"最小高度"，PowerPoint 打开时
    会按文字实际高度把行撑开 —— 撑开的行没人再回收，整张表就跑到内容区之外、
    底部的 note 也会被压住（用户复现的 P4 就是这个现象）。
    """
    return _lines_in(text, width, size_pt) * size_pt / 72 * spacing + CELL_V_PAD


def table_geom(n_rows: int, summary: str | None, has_note: bool,
               headers: list[str] | None = None, rows: list[list[str]] | None = None,
               has_side: bool = False) -> dict:
    """表格页竖向几何 —— **唯一口径**（render_table / table_slack / qc 共用）。

    三件事在这里一次性算清、彼此不矛盾：
      1. 行高：紧凑 0.46；行数多到放不下时按可用高度均分（自动压紧）；
      2. 单元格字号：往上从大到小挑，直到**每个格子的文字自然高度 ≤ 行高**
         （这是防"PowerPoint 撑行"的关键，不是防它自己不换行）；
      3. note 位置：只跟在表格下面，**不再跟右侧辅区的底边取 max** ——
         那会把 note 一路压到 BODY_BOTTOM，正好压在表格上。
    """
    top = header_y(summary) + TABLE_TOP_PAD
    note_h = TABLE_NOTE_H if has_note else 0.0
    avail = max(0.0, TABLE_BOTTOM - top - (note_h + 0.10 if note_h else 0.0))
    tw = 12.23 - (SIDE_W + SIDE_GAP) if has_side else 12.23
    widths = col_widths(headers or [], rows or [], tw) if headers else [tw / 2] * 2
    # 表头也要随可用高度收，否则 total_h 会在极紧的空间里大于 avail（表就出界了）
    head_h = min(TABLE_HEAD_H, avail)
    body_h = min(TABLE_ROW_H, max(0.0, (avail - head_h) / max(1, n_rows)))
    total_h = head_h + body_h * n_rows
    cell_size, _csp = TABLE_CELL_STEPS[-1]
    cell_overflow = 0.0
    if rows:
        for _sz, _sp in TABLE_CELL_STEPS:
            need = max((table_row_need(v, widths[c] - 0.30, _sz, _sp)
                        for r in rows for c, v in enumerate(r) if c < len(widths)),
                       default=0.0)
            if need <= body_h + 1e-6:
                cell_size, _csp = _sz, _sp
                break
            cell_overflow = max(cell_overflow, need - body_h)
    return {"top": top, "avail": avail, "head_h": head_h, "body_h": body_h,
            "total_h": total_h, "cell_size": cell_size, "cell_overflow": cell_overflow,
            "note_h": note_h, "note_y": top + total_h + 0.08, "tw": tw, "widths": widths}


def table_slack(n_rows: int, summary: str | None, has_note: bool,
                headers: list[str] | None = None, rows: list[list[str]] | None = None,
                has_side: bool = False) -> tuple[float, int, float]:
    """表格页下方还剩多少高度、还能再放几行、表格占了内容区的多少。

    「表格没铺满」不该靠拉高行高解决，而是把**下一章节或关联章节的内容并进来**；
    这个函数供 qc 把"还空多少、大约还能放几行"算给 LLM 看。
    """
    g = table_geom(n_rows, summary, has_note, headers, rows, has_side)
    avail, used = g["avail"], g["total_h"]
    slack = avail - used
    ratio = min(1.0, used / avail) if avail > 0 else 1.0
    return slack, int(max(0.0, slack) // TABLE_ROW_H), ratio


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


END_BAR_Y = 1.52          # 三色装饰条
END_TITLE_Y = 1.95        # 标题上沿
END_SECTIONS_TOP = 2.95   # 第一张信息卡上沿
END_CARD_X, END_CARD_W = 1.9, 9.5
END_SECTION_GAP = 0.22
END_CONTACT_Y = 6.75
# 结尾页正文没有卡片描述那么宽，档位单独给一套（12.5 → 9.0）
END_TEXT_STEPS = ((12.5, 1.28), (12.0, 1.26), (11.5, 1.25), (11.0, 1.22),
                  (10.5, 1.18), (10.0, 1.15), (9.5, 1.12), (9.0, 1.10))


def end_geom(sections, has_contact: bool) -> dict:
    """结尾页几何：**唯一口径**（render_end / qc 共用）。

    旧实现用 11.5pt 的旧字宽口径自己累加 y：4 段 200 字时第 4 张卡上沿到 8.75in
    （页高 7.5），contact 直接飞出画布。现在先算「所有段在某一档字号下要多高」，
    从大到小挑第一档装得下的；到 9pt 仍装不下时把间距压到 0.12 并等分可用高度，
    由 `--qc` 点名（schema 上限已按 120 字收紧，正常不会再走到这一步）。
    """
    bottom = (END_CONTACT_Y - 0.22) if has_contact else BODY_BOTTOM
    avail = max(0.6, bottom - END_SECTIONS_TOP)
    n = len(sections)
    if not n:
        return {"top": END_SECTIONS_TOP, "bottom": bottom, "avail": avail,
                "heights": [], "gap": 0.0, "size": 11.5, "spacing": 1.25,
                "total": 0.0, "card_w": END_CARD_W, "ok": True}
    text_w = END_CARD_W - 0.65          # 卡片内左右内边距 + 左侧色条
    gap = END_SECTION_GAP
    size, spacing = END_TEXT_STEPS[-1]
    heights: list[float] = []
    for size, spacing in END_TEXT_STEPS:
        heights = [0.26 + _lines_in(s.text, text_w, size) * size / 72 * spacing
                   + 0.16 for s in sections]
        if sum(heights) + gap * (n - 1) <= avail:
            break
    else:
        size, spacing = END_TEXT_STEPS[-1]
        gap = 0.12
        heights = [0.26 + _lines_in(s.text, text_w, size) * size / 72 * spacing
                   + 0.14 for s in sections]
    total = sum(heights) + gap * (n - 1)
    ok = total <= avail + 1e-6
    if not ok:                          # 硬下限仍放不下：等分，交给 qc 点名
        share = max(0.3, (avail - gap * (n - 1)) / n)
        heights = [share] * n
        total = avail
    return {"top": END_SECTIONS_TOP, "bottom": bottom, "avail": avail,
            "heights": heights, "gap": gap, "size": size, "spacing": spacing,
            "total": total, "card_w": END_CARD_W, "ok": ok}


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
CHART_TOP_PAD = 0.12
CHART_BOTTOM = BODY_BOTTOM   # 内容区底线（页脚下移后与其它版式统一）


def chart_metrics(summary: str | None, ctype: str, n_pts: int,
                  n_detail: int, has_footer: bool) -> dict:
    """图表页几何：图表与明细表的高度/位置（单一来源，渲染与 qc 共用）。

    图表**按数据点紧凑自适应**（不无脑撑满）：数据点少时省下的高度给下方明细表
    或让右侧洞察卡铺满 —— 一页始终装两块信息，别在图表下方留半页空白。
    """
    cy = header_y(summary) + CHART_TOP_PAD
    avail = CHART_BOTTOM - cy          # 页脚条已下移到内容带之外，不再预留
    ideal = 3.8 if ctype in ("pie", "doughnut", "radar") else max(2.3, 0.40 * n_pts + 1.15)
    has_detail = n_detail > 0
    ch = min(avail * (0.50 if has_detail else 1.0), ideal)
    out = {"cy": cy, "ch": ch, "detail_y": None,
           "detail_head": 0.36, "detail_row": 0.0, "detail_total": 0.0}
    if has_detail:
        dty = cy + ch + 0.20
        d_avail = CHART_BOTTOM - dty
        if d_avail >= 1.0:
            d_head = 0.36
            # 明细表把图表下方剩下的空间吃掉（上限 0.46/行）—— 半页留白不如多放几行
            d_row = min(0.46, max(0.24, (d_avail - d_head) / max(1, n_detail)))
            out.update(detail_y=dty, detail_head=d_head, detail_row=d_row,
                       detail_total=d_head + d_row * n_detail)
    return out


def chart_bottom(summary, ctype, n_pts, n_detail, has_footer) -> float:
    """图表页内容的最低底边（图表 + 可选明细表）；qc 用它算"下半页还空多少"。"""
    g = chart_metrics(summary, ctype, n_pts, n_detail, has_footer)
    if g["detail_y"] is not None:
        return g["detail_y"] + g["detail_total"]
    return g["cy"] + g["ch"]


_CHART_TYPE_MAP = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "radar": XL_CHART_TYPE.RADAR_MARKERS,
}

_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _apply_chart_fonts(root) -> None:
    """图表里**每一级**文字属性都成对写上 latin/ea/cs。

    python-pptx 不会写 `a:ea`（甚至只写 `a:latin` 或不写），中文标签就会退回主题字体
    （模板自带 Calibri）。而图表的 `c:txPr` 是**逐元素**生效的：图表级 / 图例级 /
    坐标轴标签级 / 数据标签级各有一份，只改图表级会漏掉后面几级 —— 所以整棵树扫一遍。
    """
    for tag in ("a:defRPr", "a:rPr"):
        for rPr in root.iter(qn(tag)):
            _set_font_on_rPr(rPr)


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
    """图表页派发：默认 matplotlib 高质感图，--chart-mode native 走原生可编辑图表。"""
    if CHART_MODE == "native":
        render_chart_native(slide, m)
    else:
        render_chart_mpl(slide, m)


def render_chart_mpl(slide, m):
    """图表页：左 matplotlib 高质感图（PNG 嵌图）+ 右关键洞察卡。"""
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



def chars_per_line(width_in: float, size_pt: float) -> int:
    """width 英寸里一行放得下几个**中文**字（与 `_lines_in` 同一套保守字宽）。

    qc 里几处"这个格子还能放几个字"原先各自手写 `size/72*1.06` —— 统一到这里，
    免得改了字宽口径、两边对不上。
    """
    return max(1, int(max(0.5, width_in) / (size_pt / 72 * 1.06)))


def cell_cap_chars(col_w: float, pad: float, size_pt: float) -> int:
    """表格单元格在 size_pt 下**一行**能放几个字（列宽先扣掉左右内边距）。"""
    return max(1, chars_per_line(col_w - pad, size_pt))


def text_w_in(text: str, size_pt: float = 11.5) -> float:
    """一段文字按保守字宽占多少英寸（qc 判断"标题会不会压到图标"用）。"""
    return _cell_text_w(text, size_pt)


def slack_chars(width_in: float, slack_h: float, size_pt: float,
                spacing: float = 1.30) -> int:
    """这块留白（slack_h 英寸）大约还能写几个字 —— qc 报"还能写 N 字"用。"""
    return chars_fit(width_in, slack_h, size_pt, spacing)


def chars_fit(width_in: float, height_in: float, size_pt: float,
              spacing: float = 1.35) -> int:
    """给定框能放下的最大字数（qc 报"约 X 字"用），与 `chars_per_line` 同一字宽口径。"""
    lines = int(height_in / (size_pt / 72 * spacing))
    return max(0, chars_per_line(width_in, size_pt) * lines)


CAP_MIN_SIZE = FIT_STEPS[-1][0]        # 最小正文档（10pt）——容量核对按它算


def cap_chars(width: float, height: float, steps: tuple | None = None) -> int:
    """这个文本框在**最小字号档**下能放多少字（qc 的硬容量口径）。

    渲染器会自己降档适配，所以"超容量"只可能发生在降到最小档之后 ——
    qc 按这个数点名，才不会把"渲染器本来就塞得下"的内容误判成溢出。
    """
    size, spacing = (steps or FIT_STEPS)[-1]
    return chars_fit(width, height, size, spacing)
