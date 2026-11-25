# -*- coding: utf-8 -*-
"""版式的**测量与容量**：不 import python-pptx，也不 import matplotlib。

为什么单独成模块（2026-11-25 从 `layouts.py` 拆出来）：
`layouts.py` 既是"算多大"的地方，又是"画上去"的地方，于是任何只要容量的调用方——
`plan/budgets.py`（算每页字数预算）、`dsl/ppt_dsl.py`（写页时校验放不放得下）、
`qc/qc.py`（查溢出）——都得连带拉起 python-pptx（约 0.6s）和 matplotlib（约 1.2s）。
拆开之后这条依赖链断了：`python .pi/cli.py budget` 不再为"画图"付启动成本。

**这里的函数只算数，不碰 slide 对象。** 判据很硬：一旦某个函数开始引用
`slide` / `In()` / `MSO_SHAPE` 之类的绘图符号，它就该回到 `layouts.py`。
`layouts.py` 用 `from .geom import *` 把这些名字拿回去，所以绘制代码不用改写成
`geom.xxx(...)`——改名成本为零，但两边的归属是清楚的。

改这里的常量等于改所有版式的几何，改完必须跑 `.pi/tests/smoke.py`
（19 个版式的容量与渲染回归）与 `.pi/tests/layout.py`。
"""
from __future__ import annotations

import io
import math
import os
import re

from ..paths import PI_ROOT
from .theme import CONTENT_W, FONT_SIZES, LANG_ALT, LANG_MAIN, MARGIN_L

__all__ = [
    "_BASE",
    "_BG_COVER",
    "_BG_CONTENT",
    "BODY_BOTTOM",
    "FOOTER_Y",
    "FOOTER_H",
    "FOOTER_RIGHT",
    "AUX_RAIL_W",
    "AUX_GAP",
    "AUX_MIN_MAIN_W",
    "_has_cjk",
    "_set_lang",
    "_est_lines",
    "_lines_in",
    "_para_h",
    "_ICON_KEYWORDS",
    "_auto_icon",
    "_ICON_FALLBACKS",
    "_ICON_CACHE",
    "icon_resolvable",
    "_resolve_icon",
    "BODY_SINGLE_MAX",
    "_SECTION_RE",
    "_sec_head",
    "CLAIM_MAX",
    "split_claim",
    "split_sections",
    "card_paras",
    "card_text_h",
    "split_body_groups",
    "HL_LABEL_H",
    "HL_VALUE_H",
    "HL_DESC_TOP",
    "HL_DESC_SIZE",
    "HL_DESC_SPACING",
    "highlight_h",
    "statement_geom",
    "statement_body_box",
    "statement_slack",
    "CARD_RADIUS_FLAT",
    "CARD_RADIUS_OUTLINE",
    "CARD_RADIUS_TINTED",
    "CARD_RAIL_W",
    "CARD_ACCENT_W",
    "CARD_ACCENT_PAD",
    "CARD_ACCENT_TOP_H",
    "SUMMARY_W",
    "summary_block_h",
    "header_y",
    "TOC_TOP",
    "TOC_BOTTOM",
    "TOC_GAP_X",
    "TOC_GAP_Y",
    "TOC_ROW_MAX",
    "TOC_X0",
    "_stack_heights",
    "aux_rail_fits",
    "card_frame",
    "_MIN_COL_W",
    "_MIN_CARD_W",
    "NO_X",
    "NO_GUTTER",
    "CARD_ICON_W",
    "CARD_ICON_RIGHT",
    "ICON_TITLE_GAP",
    "block_title_w",
    "fit_one_line",
    "_MIN_FLOW_W",
    "COMPARE_MIN_W",
    "LAYER_MIN_W",
    "FLOW_CYCLE_MIN_W",
    "PANEL_MIN_W",
    "HIER_MIN_W",
    "FIT_STEPS",
    "STATEMENT_STEPS",
    "_STATEMENT_STEPS",
    "HIER_STEPS",
    "HIER_GROW_MAX",
    "pick_step_ex",
    "pick_step",
    "fit_size_ex",
    "fit_size",
    "step_for_len",
    "_fit_rows",
    "STMT_HEAD_MAX",
    "STMT_HEAD_ROW",
    "statement_group_head",
    "statement_card_items",
    "statement_card_h",
    "COMPARE_GAP",
    "COMPARE_GEOM",
    "COMPARE_ICON_W",
    "COMPARE_ICON_RIGHT",
    "compare_label_w",
    "compare_desc_box",
    "numbered_geom",
    "_layer_pill_ws",
    "_layer_tags_w",
    "layered_geom",
    "_strip_step_no",
    "FLOW_CYCLE_DESC_SIZE",
    "FLOW_CYCLE_DESC_W",
    "FLOW_CYCLE_DESC_H",
    "FLOW_NODE_W",
    "FLOW_NODE_H",
    "FLOW_CYCLE_DESC_SPACING",
    "FLOW_CYCLE_STEPS",
    "flow_cycle_desc_box",
    "FLOW_ARROW_W",
    "FLOW_ARROW_GAP",
    "FLOW_WRAP_AT",
    "FLOW_ROW_GAP",
    "flow_linear_geom",
    "hierarchy_metrics",
    "hierarchy_geom",
    "PANEL_GAP",
    "PANEL_TOP_PAD",
    "PANEL_PAD",
    "PANEL_HEAD_TOP",
    "PANEL_HEAD_H",
    "PANEL_SUM_SIZE",
    "PANEL_TBL_HEAD",
    "PANEL_TBL_ROW",
    "PANEL_TBL_CELL",
    "PANEL_PT_STEPS",
    "PANEL_PT_SIZE",
    "PANEL_PT_FLOOR",
    "PANEL_TBL_ROW_MIN",
    "PANEL_TBL_CELL_STEPS",
    "panel_area",
    "_panel_points_need",
    "panel_points_layout",
    "panel_content_geom",
    "panel_slack",
    "panel_pressure",
    "_panel_sum_h",
    "GRID_GAP",
    "GRID_TOP_PAD",
    "GRID_DESC_TOP",
    "grid_geom",
    "KG_GAP",
    "KG_TOP_PAD",
    "KG_GROW_MAX",
    "_KG_TITLE_ROW",
    "_kg_card_h",
    "key_groups_geom",
    "TL_AXIS_HALF",
    "TL_TOP_PAD",
    "TL_VERTICAL_AT",
    "TL_VGAP",
    "TL_TIME_W",
    "TL_AXIS_X",
    "TL_CARD_X",
    "timeline_geom",
    "FUNNEL_DESC_W",
    "funnel_geom",
    "MATRIX_ROWLABEL_W",
    "MATRIX_GAP",
    "MATRIX_COLHEAD_H",
    "MATRIX_TOP_PAD",
    "MATRIX_DESC_TOP",
    "MATRIX_BOTTOM_PAD",
    "MATRIX_STEPS",
    "matrix_geom",
    "FIG_EXT_VECTOR",
    "_CTYPE_VECTOR",
    "_blob_size_in",
    "image_size_in",
    "resolve_image_path",
    "FIGURE_MIN_W",
    "FIGURE_PAD",
    "FIGURE_CAPTION_H",
    "figure_area",
    "RAIL_PAD_MAX",
    "RAIL_GAP_MAX",
    "RAIL_GROW_MAX",
    "_fill_rail",
    "point_rail_need",
    "TABLE_TOP_PAD",
    "TABLE_BOTTOM",
    "TABLE_HEAD_H",
    "TABLE_ROW_H",
    "TABLE_NOTE_H",
    "SIDE_W",
    "SIDE_GAP",
    "_cell_text_w",
    "col_widths",
    "CALLOUT_H_MAX",
    "CALLOUT_H_MIN",
    "CALLOUT_H",
    "_CALLOUT_LABEL_H",
    "_CALLOUT_VALUE_H",
    "_CALLOUT_SIZE",
    "callout_height",
    "TABLE_CELL_STEPS",
    "CELL_V_PAD",
    "table_row_need",
    "table_geom",
    "table_slack",
    "END_BAR_Y",
    "END_CARD_X",
    "END_CARD_W",
    "END_TITLE_Y",
    "END_SECTIONS_TOP",
    "END_SECTION_GAP",
    "END_CONTACT_Y",
    "END_TEXT_STEPS",
    "end_geom",
    "CHART_TOP_PAD",
    "CHART_BOTTOM",
    "chart_metrics",
    "chart_bottom",
    "_C_NS",
    "chars_per_line",
    "cell_cap_chars",
    "text_w_in",
    "slack_chars",
    "chars_fit",
    "CAP_MIN_SIZE",
    "cap_chars",
]


_BASE = str(PI_ROOT)                      # .pi/（资源根，见 engine/paths.py）


_BG_COVER = os.path.join(_BASE, "assets", "template", "bg-cover.png")


_BG_CONTENT = os.path.join(_BASE, "assets", "template", "bg-content.png")


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


# 右纵栏（辅区）骨架。放在这里而不是版式段内：statement_geom / highlight_h
# 在文件更靠前的位置就要用到，常量必须先生效。
AUX_RAIL_W = 3.55                     # 右侧纵栏宽度（与表格页 side 同宽）


AUX_GAP = 0.22                        # 主内容与右纵栏的间距


AUX_MIN_MAIN_W = 12.23 - (AUX_RAIL_W + AUX_GAP)   # 让出右栏后主内容还剩多宽


def _has_cjk(text: str) -> bool:
    """文本里有没有中日韩汉字（决定这段 run 的主语言）。"""
    return any("\u3400" <= ch <= "\u4dbf" or "\u4e00" <= ch <= "\u9fff" for ch in text or "")


def _set_lang(rPr, text: str | None = None) -> None:
    """写 run 级语言标记 `lang` + `altLang`。

    Office 2007 的 Ribbon 字体框按 run 的语言决定显示哪个字体名：`lang` 缺失时它按
    默认语言（`en-US`）走西文路径，中文名在那种路径下会渲染成乱码。Office 自己保存
    的 pptx 里，每个 `a:rPr` 都同时带 `lang` 与 `altLang`。

    纯西文的 run 标 `en-US`（否则英文会被当中文做校对/断词），其余标 `zh-CN`。
    图表等拿不到文本上下文的场景传 `None`，按中文交付默认值走。
    """
    if text is not None and not _has_cjk(text):
        rPr.set("lang", LANG_ALT)
        rPr.set("altLang", LANG_MAIN)
    else:
        rPr.set("lang", LANG_MAIN)
        rPr.set("altLang", LANG_ALT)


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
            from .iconlib import resolve_icon as _spec
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


TOC_TOP = 1.90            # 目录首行上沿（大标题 + CONTENTS 之下）


TOC_BOTTOM = 6.20         # 目录末行下沿（留出页脚带）


TOC_GAP_X = 0.50


TOC_GAP_Y = 0.22


TOC_ROW_MAX = 1.45        # 目录条最大高度：再多就居中留白，不把条目拉成胖块


TOC_X0 = 0.90


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


END_BAR_Y = 1.52          # 三色装饰条


END_TITLE_Y = 1.95        # 标题上沿


END_SECTIONS_TOP = 2.95   # 第一张信息卡上沿


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


_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


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


# ---- 共用尺寸常量（原为 tuple 赋值，随 end_geom / flow_cycle_desc_box 一起搬来）----
# 这些名字在 layouts 的绘制函数里也用，靠 __all__ 星号导入回去。
END_CARD_X, END_CARD_W = 1.9, 9.5
FLOW_NODE_W, FLOW_NODE_H = 1.72, 1.15
FLOW_CYCLE_DESC_W, FLOW_CYCLE_DESC_H = 1.44, 0.48
