# -*- coding: utf-8 -*-
"""matplotlib 高质感图表渲染：把 ChartSpec 画成高分辨率 PNG，再嵌入 PPT。

为什么不用 python-pptx 原生图表：原生图表样式简陋，字体/网格/配色/数值标签
可控性差。matplotlib（离线、已随环境内置）可精确控制配色、数值标签、网格线、
雷达填充、环形中心文字等，观感更接近参考 PDF。

高级手法（本模块已应用）：
- 去掉图表四边框，只留值轴一条基线；网格用极淡虚线，压到图形背后
- 刻度去掉小短线（tick length=0），只留数字，画面更干净
- 柱/条形单系列时，把最大值柱用暖色（orange）点亮，其余统一主色，一眼抓重点
- 数值标签加粗、贴近柱端，避开默认的灰小字
- 折线加标记点白描边 + 线下淡色面积填充，趋势更醒目
- 环形图减薄环宽、中心放大「单位 + 总量」，饼图楔形留白描边
- 雷达图加同心网格环 + 半透明填充 + 顶点数值

用法：charts.render_chart_png(spec, buf, w_in, h_in)
    spec  = ChartSpec（见 ppt_dsl.py）
    buf   = 文件类对象（BytesIO 即可，渲染后 seek(0) 交给 add_picture）
    w_in/h_in = 目标图幅（英寸），与嵌入尺寸一致，保证不拉伸变形。
"""
from __future__ import annotations

import logging
import math
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.ticker import MaxNLocator

from theme import ACCENT_CYCLE, CHART_SERIES, COLORS, FONT_LATIN, FONT_MAIN

# ---------- 字体（与 PPT 正文同一口径：西文 Times New Roman + 中文微软雅黑） ----------
# 只用 **Windows 自带字体**，不随包分发字体、不下载。字体候选列表按"系统里通常都在"
# 的顺序排；matplotlib 3.6+ 支持 `font.family` 传列表做**逐字符回退**，
# 所以英文/数字命中外文、中文回落雅黑，一张图里两种字体同时成立。
_FONT_FILES = [
    r"C:\Windows\Fonts\times.ttf",     # Times New Roman
    r"C:\Windows\Fonts\timesbd.ttf",   # Times New Roman 粗体
    r"C:\Windows\Fonts\msyh.ttc",      # 微软雅黑
    r"C:\Windows\Fonts\msyhbd.ttc",    # 微软雅黑粗体
]
_FONT_STACK = [FONT_LATIN, FONT_MAIN, "Microsoft YaHei", "SimHei", "DejaVu Sans"]
_setup_done = False


def _setup_font():
    global _setup_done
    if _setup_done:
        return
    _setup_done = True
    for p in _FONT_FILES:
        if os.path.exists(p):
            try:
                fm.fontManager.addfont(p)
            except Exception:
                pass
    names = {f.name for f in fm.fontManager.ttflist}
    stack = [n for n in _FONT_STACK if n in names] or ["DejaVu Sans"]
    plt.rcParams["font.family"] = stack
    plt.rcParams["font.sans-serif"] = stack
    plt.rcParams["font.serif"] = stack
    plt.rcParams["axes.unicode_minus"] = False
    # matplotlib 找不到某一档字重时会往 stderr 打 "findfont: Failed to find font weight"
    # （字号/字形仍然正常，只是那一处不加粗）。这类日志会混进 Agent 读的渲染日志里，
    # 压到 ERROR 级别；**缺字（Glyph missing）是 UserWarning，不受影响、仍会暴露**。
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)


# ---------- 配色工具 ----------
def _hex(name: str) -> str:
    return "#" + COLORS.get(name, name)


def _accent(i: int) -> str:
    return _hex(ACCENT_CYCLE[i % len(ACCENT_CYCLE)])


def _series(i: int) -> str:
    return _hex(CHART_SERIES[i % len(CHART_SERIES)])


def _fmt(v) -> str:
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:.1f}"
    except (TypeError, ValueError):
        return str(v)


def _max_val(spec) -> float:
    mx = 0.0
    for s in spec.series:
        for v in s.values:
            mx = max(mx, float(v))
    return mx if mx > 0 else 1.0


# ---------- 各图表类型 ----------
def _draw_bars(ax, spec, horizontal: bool):
    cats = spec.categories
    series = spec.series
    multi = len(series) > 1
    n = len(cats)
    mx = _max_val(spec)
    unit = spec.unit or ""

    # 单系列：点亮最大值柱（暖色），其余统一主色，一眼抓重点
    max_i = None
    if not multi:
        max_i = max(range(n), key=lambda i: float(series[0].values[i]))

    if horizontal:
        positions = list(range(n))[::-1]
        width = 0.62 / max(1, len(series))
        for i, s in enumerate(series):
            off = (i - (len(series) - 1) / 2) * width
            ys = [p + off for p in positions]
            colors = ([_series(i)] * n if multi
                      else [_hex("orange") if j == max_i else _hex("primary")
                            for j in range(n)])
            ax.barh(ys, s.values, height=width * 0.84, color=colors,
                    edgecolor="none", zorder=3)
            for j, (ypos, v) in enumerate(zip(ys, s.values)):
                c = _hex("orange") if (not multi and j == max_i) else _hex("text")
                ax.text(float(v) + mx * 0.012, ypos, _fmt(v) + unit,
                        va="center", ha="left", fontsize=9.5, fontweight="bold",
                        color=c, zorder=4)
        ax.set_yticks(positions)
        ax.set_yticklabels(cats)
        ax.set_xlim(0, mx * 1.30)
        _style_ax(ax, grid_axis="x")
    else:
        x = list(range(n))
        width = 0.62 / max(1, len(series))
        for i, s in enumerate(series):
            off = (i - (len(series) - 1) / 2) * width
            xs = [xi + off for xi in x]
            colors = ([_series(i)] * n if multi
                      else [_hex("orange") if j == max_i else _hex("primary")
                            for j in range(n)])
            ax.bar(xs, s.values, width=width * 0.84, color=colors,
                   edgecolor="none", zorder=3,
                   label=s.name if multi else None)
            for j, (xi, v) in enumerate(zip(xs, s.values)):
                c = _hex("orange") if (not multi and j == max_i) else _hex("text")
                ax.text(xi, float(v) + mx * 0.015, _fmt(v) + unit,
                        va="bottom", ha="center", fontsize=9.5, fontweight="bold",
                        color=c, zorder=4)
        ax.set_xticks(x)
        ax.set_xticklabels(cats)
        ax.set_ylim(0, mx * 1.22)
        _style_ax(ax, grid_axis="y")

    if multi:
        ax.legend(frameon=False, fontsize=10, loc="upper right",
                  labelcolor=_hex("text_sub"))
    ax.tick_params(colors=_hex("text_sub"), labelsize=10)


def _draw_line(ax, spec):
    cats = spec.categories
    x = list(range(len(cats)))
    mx = _max_val(spec)
    unit = spec.unit or ""
    for i, s in enumerate(spec.series):
        color = _series(i)
        y = [float(v) for v in s.values]
        ax.plot(x, y, marker="o", markersize=5.5, linewidth=2.6, color=color,
                solid_capstyle="round", markeredgecolor="white",
                markeredgewidth=1.2, zorder=4,
                label=s.name if len(spec.series) > 1 else None)
        # 线下淡色面积填充，趋势更醒目
        ax.fill_between(x, 0, y, color=color, alpha=0.13, zorder=1)
        for xi, v in zip(x, y):
            ax.annotate(_fmt(v) + unit, (xi, v), textcoords="offset points",
                        xytext=(0, 10), ha="center", fontsize=9.5,
                        fontweight="bold", color=color, zorder=5)
    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylim(0, mx * 1.28)
    _style_ax(ax, grid_axis="y")
    if len(spec.series) > 1:
        ax.legend(frameon=False, fontsize=10, loc="upper right",
                  labelcolor=_hex("text_sub"))
    ax.tick_params(colors=_hex("text_sub"), labelsize=10)


def _draw_pie(ax, spec, doughnut: bool):
    vals = [float(v) for v in spec.series[0].values]
    colors = [_accent(i) for i in range(len(vals))]
    wedgeprops = dict(edgecolor="white", linewidth=2.0)
    if doughnut:
        wedgeprops["width"] = 0.48
    wedges, texts, autotexts = ax.pie(
        vals, labels=spec.categories, colors=colors,
        autopct=lambda p: f"{p:.0f}%",
        startangle=90, counterclock=False, wedgeprops=wedgeprops,
        pctdistance=0.80 if doughnut else 0.62,
        labeldistance=1.10,
        textprops=dict(color=_hex("text_sub"), fontsize=10),
    )
    for at in autotexts:
        at.set_color("white")
        at.set_fontweight("bold")
        at.set_fontsize(10)
    ax.set_aspect("equal")
    ax.set_xlim(-1.4, 1.4)
    ax.set_ylim(-1.4, 1.4)
    if doughnut:
        ax.text(0, 0.10, spec.unit or "合计", ha="center", va="center",
                fontsize=12, color=_hex("text_sub"))
        ax.text(0, -0.20, _fmt(sum(vals)), ha="center", va="center",
                fontsize=20, fontweight="bold", color=_hex("text"))


def _draw_radar(ax, spec):
    cats = spec.categories
    vals = [float(v) for v in spec.series[0].values]
    n = len(cats)
    angles = [i / n * 2 * math.pi for i in range(n)]
    vals_c = vals + vals[:1]
    angles_c = angles + angles[:1]
    color = _series(0)
    unit = spec.unit or ""
    mx = _max_val(spec)
    ax.set_xticks(angles)
    ax.set_xticklabels(cats, fontsize=10, color=_hex("text"))
    ax.set_yticklabels([])
    ax.set_ylim(0, mx * 1.15)
    # 同心网格环 + 角度辐线淡化
    ax.grid(color=_hex("border"), linewidth=0.6, linestyle="--", alpha=0.6)
    ax.spines["polar"].set_color(_hex("border"))
    ax.plot(angles_c, vals_c, color=color, linewidth=2.2, marker="o",
            markersize=5, zorder=4)
    ax.fill(angles_c, vals_c, color=color, alpha=0.20, zorder=2)
    for a, v in zip(angles, vals):
        ax.annotate(_fmt(v) + unit, (a, v), textcoords="offset points",
                    xytext=(0, 7), ha="center", fontsize=9.5,
                    fontweight="bold", color=color, zorder=5)


def _style_ax(ax, grid_axis="y"):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(_hex("border"))
        ax.spines[sp].set_linewidth(0.8)
    ax.grid(True, axis=grid_axis, color=_hex("border"), linewidth=0.6,
            linestyle="--", alpha=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(colors=_hex("text_sub"), labelsize=10, length=0)


# ---------- 入口 ----------
def render_chart_png(spec, buf, w_in: float, h_in: float):
    """把 ChartSpec 渲染为 PNG 写入 buf。图幅与嵌入尺寸一致，避免拉伸。"""
    _setup_font()
    fig = plt.figure(figsize=(w_in, h_in), dpi=200)
    fig.patch.set_facecolor("white")
    t = spec.type
    if t in ("column", "bar"):
        ax = fig.add_subplot(111)
        _draw_bars(ax, spec, horizontal=(t == "bar"))
        fig.subplots_adjust(left=0.14, right=0.96, top=0.88, bottom=0.18)
    elif t == "line":
        ax = fig.add_subplot(111)
        _draw_line(ax, spec)
        fig.subplots_adjust(left=0.13, right=0.96, top=0.82, bottom=0.16)
    elif t in ("pie", "doughnut"):
        ax = fig.add_subplot(111)
        _draw_pie(ax, spec, doughnut=(t == "doughnut"))
        fig.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.02)
    elif t == "radar":
        ax = fig.add_subplot(111, projection="polar")
        _draw_radar(ax, spec)
        fig.subplots_adjust(left=0.08, right=0.92, top=0.92, bottom=0.08)
    else:
        ax = fig.add_subplot(111)
        _draw_bars(ax, spec, horizontal=False)
        fig.subplots_adjust(left=0.14, right=0.96, top=0.88, bottom=0.18)
    fig.savefig(buf, format="png", dpi=200, facecolor="white")
    plt.close(fig)
