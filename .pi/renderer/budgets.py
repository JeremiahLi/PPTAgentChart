# -*- coding: utf-8 -*-
"""字数预算表：把"动态容量"从**校验时反算**挪到**写作前下发**。

背景：ppt_dsl 的动态容量校验（`_cap_hint` + 各版式几何函数）只在生成**之后**报数，
模型写作时看不到"8 张卡每卡 90 字"这类随条目数浮动的真实容量，写满静态 max_length
就撞墙 → 回炉。本脚本读 outline.json，对每页按版式与条目数**预先**算出每字段的
写作目标字数，注入生成 prompt —— 校验与预算出自同一套几何函数，口径零漂移。

实现方式（关键决定）：**拦截式探测**，不重新实现任何几何 ——
- 接管 `ppt_dsl._cap_hint`：构造合成页实例时把"要报错"换成"记下 cap"，
  于是预算数与 schema 校验数必然一致（同一段代码算出来的）；
- key_groups / end 两处校验没走 `_cap_hint`，在下方按其校验器同款公式**直接复算**
  （各 6 行，已注明"与某校验器保持同步"，改动校验器时同步改这里）。

三个数字的口径（写进输出头，别在别处再写一遍）：
- **写作目标** = min(静态 max_length, 几何容量 cap)——照它写必过校验；
- **校验放行线** = int(cap×1.1)+8（CAP_TOL/CAP_SLACK，ppt_dsl 里的容差）；
- **写作区间** = 静态下限（含 briefs 方向红线的密度下限）到写作目标的 ~90%。

用法（全部从仓库根执行）：
  python .pi/renderer/budgets.py .pi/work/outline.json -o .pi/work/budgets.md
  python .pi/renderer/budgets.py --all -o .pi/work/budgets-all.md   # 全版式参考表
  python .pi/renderer/budgets.py --selfcheck                        # 自证"照预算写必过校验"
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pydantic import BaseModel
from typing import get_args, get_origin

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import layouts as L
import ppt_dsl

CLS = {
    "cover": ppt_dsl.CoverSlide, "toc": ppt_dsl.TocSlide,
    "section": ppt_dsl.SectionSlide, "statement": ppt_dsl.StatementSlide,
    "compare_3col": ppt_dsl.CompareSlide, "numbered_cards": ppt_dsl.NumberedCardsSlide,
    "layered_stack": ppt_dsl.LayeredStackSlide, "flow": ppt_dsl.FlowSlide,
    "hierarchy": ppt_dsl.HierarchySlide, "table": ppt_dsl.TableSlide,
    "chart": ppt_dsl.ChartSlide, "panel_2col": ppt_dsl.Panel2ColSlide,
    "figure": ppt_dsl.FigureSlide, "grid_2x2": ppt_dsl.QuadrantGridSlide,
    "timeline": ppt_dsl.TimelineSlide, "funnel": ppt_dsl.FunnelSlide,
    "matrix": ppt_dsl.MatrixSlide, "key_groups": ppt_dsl.KeyGroupsSlide,
    "end": ppt_dsl.EndSlide,
}

# 条目容器：slide_type -> (集合字段名, 条目模型, 最少, 最多)
ITEMS = {
    "compare_3col": ("columns", ppt_dsl.CompareColumn, 3, 4),
    "numbered_cards": ("cards", ppt_dsl.NumberedCard, 4, 8),
    "layered_stack": ("layers", ppt_dsl.StackLayer, 3, 12),
    "flow": ("steps", ppt_dsl.FlowStep, 3, 12),
    "hierarchy": ("nodes", ppt_dsl.HierarchyNode, 2, 8),
    "timeline": ("milestones", ppt_dsl.Milestone, 3, 12),
    "funnel": ("stages", ppt_dsl.FunnelStage, 3, 8),
    "matrix": ("rows", ppt_dsl.MatrixRow, 2, 4),
    "key_groups": ("groups", ppt_dsl.KeyGroup, 2, 4),
}

# 各版式"每条"动态字段：预算表关注的 attr（grid_2x2 多一个 value）
ITEM_ATTR = {
    "compare_3col": ["desc"], "numbered_cards": ["desc"],
    "layered_stack": ["desc"], "flow": ["desc"], "hierarchy": ["desc"],
    "timeline": ["desc"], "funnel": ["desc"], "matrix": ["desc"],
    "key_groups": ["points"], "statement": ["body"], "grid_2x2": ["desc", "value"],
}


def _z(k: int) -> str:
    """占位填充串：把静态上限撑满，让几何校验（被拦截后）报出真实容量。"""
    return "占" * k


# ---------- 合成页实例 ----------

def _synth_kwargs(st: str, n: int, aux=False, ring=False,
                  full_summary=False, right_rail=False, cols=None) -> dict:
    """构造一页合法的合成实例（静态字段全部满足，动态字段撑满）。"""
    callout = [{"label": "辅区一", "value": "12%", "text": _z(20)},
               {"label": "辅区二", "value": "3.4x", "text": _z(20)}] if aux else []
    t = "占位标题占位标题占位"
    if st == "numbered_cards":
        return dict(slide_type=st, title=t, cards=[
            {"no": str(i + 1), "name": "卡头判断句带数字", "desc": _z(240)}
            for i in range(n)], callout=callout)
    if st == "compare_3col":
        return dict(slide_type=st, title=t,
                    summary=_z(240) if full_summary else None,
                    columns=[{"label": "栏名", "tag": "核心机制短语", "desc": _z(220),
                              "points": ["要点一二三", "要点四五六"],
                              "metaphor": "生活化比喻的说法"} for _ in range(n)])
    if st == "layered_stack":
        return dict(slide_type=st, title=t, layers=[
            {"label": f"第{i + 1}层", "desc": _z(180)} for i in range(n)],
            callout=callout)
    if st == "flow":
        return dict(slide_type=st, title=t, steps=[
            {"label": f"步骤{i + 1}", "desc": _z(80)} for i in range(n)],
            callout=callout, loop_note=_z(60) if ring else None)
    if st == "hierarchy":
        return dict(slide_type=st, title=t, root="根节点整体", mode="contain",
                    nodes=[{"label": f"模块{i + 1}", "desc": _z(100)}
                           for i in range(n)], callout=callout)
    if st == "timeline":
        return dict(slide_type=st, title=t, milestones=[
            {"label": f"节点{i + 1}", "desc": _z(120)} for i in range(n)],
            callout=callout)
    if st == "funnel":
        return dict(slide_type=st, title=t, stages=[
            {"label": f"层{i + 1}", "desc": _z(100)} for i in range(n)],
            callout=callout)
    if st == "matrix":
        return dict(slide_type=st, title=t,
                    col_labels=[f"第{j + 1}列" for j in range(cols)],
                    rows=[{"label": f"行{i + 1}",
                           "cells": [{"title": f"格{i + 1}{j + 1}", "desc": _z(90)}
                                     for j in range(cols)]} for i in range(n)])
    if st == "key_groups":
        return dict(slide_type=st, title=t, groups=[
            {"title": f"组标题{i + 1}判断", "points": [_z(4), _z(4), _z(4)]}
            for i in range(n)], callout=callout)
    if st == "grid_2x2":
        cell = lambda k: {"title": f"格{k}", "value": "85%", "desc": _z(60)}
        return dict(slide_type=st, title=t, q_top_left=cell("左上"),
                    q_top_right=cell("右上"), q_bottom_left=cell("左下"),
                    q_bottom_right=cell("右下"))
    if st == "statement":
        return dict(slide_type=st, title=_z(14), body=_z(700),
                    highlight={"label": "关键数字", "value": "52.8%→66.5%",
                               "desc": _z(20)} if right_rail else None,
                    points=[{"label": "论据一", "text": _z(44)},
                            {"label": "论据二", "text": _z(44)}] if right_rail else [])
    raise ValueError(f"没有 {st} 的合成实例模板")


# ---------- 拦截式探测 ----------

class _Recorder:
    """顶替 _cap_hint：不报错，只记 (attr, cap)。"""

    def __init__(self):
        self.rows = []

    def __call__(self, items, attr, width, height, label="", hint="", steps=None):
        if width <= 0 or height <= 0:
            return
        cap = L.cap_chars(width, height, steps)
        if cap > 0:
            self.rows.append((attr, cap))


def _probe(st: str, **cfg) -> dict:
    """构造合成实例并拦截 _cap_hint，返回 {attr: cap}（同 attr 取最小 = 保守口径）。"""
    n = cfg.pop("n")
    rec = _Recorder()
    orig = ppt_dsl._cap_hint
    ppt_dsl._cap_hint = rec
    try:
        CLS[st].model_validate(_synth_kwargs(st, n, **cfg))
    finally:
        ppt_dsl._cap_hint = orig
    out = {}
    for attr, cap in rec.rows:
        out[attr] = min(out.get(attr, 10 ** 9), cap)
    return out


def _probe_configs(st: str, n: int) -> list[tuple[str, dict]]:
    """一个版式要探测哪几种配置：容量受配置影响时逐行列出，不受影响时合并。"""
    if st == "compare_3col":
        return [("页头结论段不写，每栏 2 条要点", dict(full_summary=False)),
                ("页头结论段写满 240 字，每栏 2 条要点", dict(full_summary=True))]
    if st == "flow":
        cfgs = [("线性/两行折返", dict(ring=False, aux=False)),
                ("线性 + 2 条 callout", dict(ring=False, aux=True))]
        if n <= L.FLOW_WRAP_AT:
            cfgs.append(("环形（带 loop_note）", dict(ring=True, aux=True)))
        return cfgs
    if st == "statement":
        return [("无右栏", dict(right_rail=False)),
                ("右栏放高亮卡 + 2 条论据", dict(right_rail=True))]
    if st in ("numbered_cards", "layered_stack", "timeline", "funnel",
              "hierarchy", "key_groups"):
        return [("无辅区", dict(aux=False)), ("带 2 条 callout", dict(aux=True))]
    if st == "matrix":
        return [(f"{n} 行 × {c} 列", dict(cols=c)) for c in (2, 3, 4)]
    return [("（容量不受配置影响）", {})]


# key_groups / end：校验器没走 _cap_hint，按同款公式直接复算。
# 同步约定：改 KeyGroupsSlide.fit_capacity / EndSlide.fit_capacity 时同步改这里。

def _key_groups_caps(n: int, aux: bool) -> list[int]:
    """与 KeyGroupsSlide.fit_capacity 逐卡同款：卡宽逐张不同，逐卡算。"""
    m = ppt_dsl.KeyGroupsSlide.model_validate(
        _synth_kwargs("key_groups", n, aux=aux))
    g = L.key_groups_geom(m)
    out = []
    for cg in g["cards"]:
        grp = m.groups[cg["i"]]
        k = max(1, len(grp.points))
        out.append(L.cap_chars(cg["text_w"], max(0.18, cg["pts_h"] / k - 0.05)))
    return out


def _end_caps(n: int, contact: bool) -> int:
    """与 EndSlide.fit_capacity 报错路径同款：每段正文的兜底字数。"""
    secs = [ppt_dsl.EndSection(heading=f"段{i + 1}", text=_z(4)) for i in range(n)]
    g = L.end_geom(secs, contact)
    return L.cap_chars(g["card_w"] - 0.62,
                       max(0.3, g["avail"] / n - 0.42), L.END_TEXT_STEPS)


# ---------- 静态契约（pydantic 自省，不手抄） ----------

def _len_bounds(fi):
    mn = mx = None
    for c in fi.metadata:
        v = getattr(c, "min_length", None)
        if isinstance(v, int):
            mn = v if mn is None else min(mn, v)
        v = getattr(c, "max_length", None)
        if isinstance(v, int):
            mx = v if mx is None else min(mx, v)
    return mn, mx


def _unwrap(fi):
    args = [a for a in get_args(fi.annotation) if a is not type(None)]
    return args[0] if args else fi.annotation


def _fmt_scalar(name, mn, mx):
    if mn is not None and mx is not None:
        return f"{name} {mn}~{mx}"
    if mx is not None:
        return f"{name}≤{mx}"
    if mn is not None:
        return f"{name}≥{mn}"
    return None


def _static_parts(cls, top=True):
    parts = []
    for name, fi in cls.model_fields.items():
        if name == "slide_type":
            continue
        mn, mx = _len_bounds(fi)
        base = _unwrap(fi)
        origin = get_origin(base)
        if origin is list:
            item = get_args(base)[0]
            rng = (f"{mn or ''}~{mx}" if mx else f"≥{mn}" if mn else "")
            if isinstance(item, type) and issubclass(item, BaseModel):
                inner = " · ".join(filter(None, _static_parts(item, top=False)))
                unit = "行" if name in ("rows", "detail_rows") else "项"
                body = f"{{{inner}}}" if inner else ""
                parts.append(f"{name} {rng} {unit}{body}".replace("  ", " "))
            else:
                parts.append(_fmt_scalar(f"{name}(列表)", mn, mx) or "")
        elif isinstance(base, type) and issubclass(base, BaseModel):
            inner = " · ".join(filter(None, _static_parts(base, top=False)))
            parts.append(f"{name}{{{inner}}}")
        else:
            parts.append(_fmt_scalar(name, mn, mx) or "")
    return [p for p in parts if p]


# ---------- 预算计算与渲染 ----------

def _static_cap(st: str, attr: str):
    """attr 的静态 max_length（pydantic 自省）——几何容量要比它大时以它封顶。"""
    if st == "statement":
        return _len_bounds(ppt_dsl.StatementSlide.model_fields["body"])[1]
    if st == "grid_2x2":
        fi = ppt_dsl.QuadrantCell.model_fields.get(attr)
        return _len_bounds(fi)[1] if fi else None
    if st == "matrix":  # desc 不在 MatrixRow 上，在 rows[].cells[].desc 的 MatrixCell 上
        fi = ppt_dsl.MatrixCell.model_fields.get(attr)
        return _len_bounds(fi)[1] if fi else None
    if st in ITEMS:
        fi = ITEMS[st][1].model_fields.get(attr)
        return _len_bounds(fi)[1] if fi else None
    return None


def _target(st: str, attr: str, cap: int) -> int:
    """写作目标 = min(静态上限, 几何容量)。"""
    sm = _static_cap(st, attr)
    return min(sm, cap) if sm else cap


def _merge_same(per_cfg: dict) -> list[tuple[str, dict]]:
    """配置不同但结果相同的行合并成一行（如"带不带辅区不影响"）。"""
    merged: dict[str, tuple[list[str], dict]] = {}
    for label, by_n in per_cfg.items():
        key = json.dumps(by_n, sort_keys=True, ensure_ascii=False)
        merged.setdefault(key, ([], by_n))[0].append(label)
    return [((" / ".join(ls) if len(ls) > 1 else ls[0]), by_n)
            for ls, by_n in merged.values()]


def _dynamic_lines(st: str) -> list[str]:
    """返回该版式的动态容量行（每行一个"配置 → 各条目数的写作目标"）。"""
    lines = []
    if st in ("numbered_cards", "compare_3col", "layered_stack", "flow",
              "hierarchy", "timeline", "funnel"):
        field, _, lo, hi = ITEMS[st]
        attr = ITEM_ATTR[st][0]
        per_cfg = {}
        for cnt in range(lo, hi + 1):
            for label, cfg in _probe_configs(st, cnt):
                caps = _probe(st, n=cnt, **cfg)
                if attr in caps:
                    per_cfg.setdefault(label, {})[cnt] = _target(st, attr, caps[attr])
        for label, by_n in _merge_same(per_cfg):
            body = " · ".join(f"{k} 项≈{v}" for k, v in sorted(by_n.items()))
            lines.append(f"- `{field}[].{attr}`：{body}　（{label}）")
    elif st == "matrix":
        rows_field, _, lo, hi = ITEMS["matrix"]
        per_cfg = {}
        for r in range(lo, hi + 1):
            for label, cfg in _probe_configs("matrix", r):
                caps = _probe("matrix", n=r, **cfg)
                per_cfg.setdefault(label, {})[r] = {
                    cfg["cols"]: _target("matrix", "desc", caps.get("desc", 0))}
        for label, by_r in _merge_same(per_cfg):
            seg = " · ".join(
                f"{r} 行：" + " · ".join(f"{c} 列≈{v}" for c, v in sorted(d.items()))
                for r, d in sorted(by_r.items()))
            lines.append(f"- `rows[].cells[].desc`：{seg}　（{label}）")
    elif st == "key_groups":
        per_cfg = {}
        for label, cfg in _probe_configs("key_groups", 0):
            by_n = {}
            for n in range(2, 5):
                caps = _key_groups_caps(n, cfg.get("aux", False))
                by_n[n] = "/".join(str(c) for c in caps)
            per_cfg[label] = by_n
        for label, by_n in _merge_same(per_cfg):
            body = " · ".join(f"{k} 组按 {v}" for k, v in by_n.items())
            lines.append(f"- `groups[].points` 每条（按 3 条要点算，逐卡）：{body}　"
                         f"（{label}；要点条数越多每条越矮，第 3 组跨整行能写更多）")
    elif st == "grid_2x2":
        caps = _probe("grid_2x2", n=1)
        if "desc" in caps:
            lines.append(f"- `q_*.desc`：≈{_target('grid_2x2', 'desc', caps['desc'])}"
                         "　（固定 2×2，不受配置影响）")
        if "value" in caps:
            lines.append(f"- `q_*.value`：≈{_target('grid_2x2', 'value', caps['value'])}"
                         "　（18pt 大字，写短些）")
    elif st == "statement":
        for label, cfg in _probe_configs("statement", 0):
            caps = _probe("statement", n=1, **cfg)
            if "body" in caps:
                lines.append(f"- `body`：≈{_target('statement', 'body', caps['body'])}"
                             f"　（{label}；换行分卡 ≤3 张，卡头句 ≤16 字）")
    elif st == "end":
        for contact in (False, True):
            by_n = {n: _end_caps(n, contact) for n in (2, 3, 4)}
            body = " · ".join(f"{k} 段≈{v}" for k, v in by_n.items())
            tag = "带联系方式" if contact else "无联系方式"
            lines.append(f"- `sections[].text`：{body}　（{tag}）")
    return lines


def _page_block(no: int, sl: dict) -> str:
    st = sl.get("slide_type", "")
    head = sl.get("module") or sl.get("title") or ""
    out = [f"### 第 {no} 页 · {st}　{head}", ""]
    cls = CLS.get(st)
    if cls is None:
        out.append(f"> ⚠️ 未知版式 `{st}`，查 ppt-common 版式契约表。")
        return "\n".join(out)
    dyn = _dynamic_lines(st)
    if dyn:
        out.append("**动态容量**（写作目标 = min(静态上限, 几何容量)；"
                   "写到目标的 80~90% 最稳，超目标必看放行线）：")
        out.extend(dyn)
    else:
        out.append("**动态容量**：本版式无随条目数浮动的字段（静态上限即真实容量）。")
    static = " ｜ ".join(_static_parts(cls))
    out.append("")
    out.append(f"**静态契约**：{static}")
    rng = ""
    if st in ITEMS:
        field, _, lo, hi = ITEMS[st]
        rng = f"　先定 `{field}` 条目数（{lo}~{hi}）再写内容，**定完别改**——条目数一变，上面全部预算作废。"
    out.append("")
    out.append(f"> 写作区间：每字段不低于静态下限（密度红线见 briefs 与 ppt-common，"
               f"此处不重复），上限按上方写作目标的 ~90%。{rng}")
    return "\n".join(out)


HEADER = """# 字数预算表（由 .pi/renderer/budgets.py 生成，勿手改）

来源：本表所有数字由渲染器**同一套几何函数**算出（拦截 schema 校验路径探测），
与 `build.py` 校验报的「每条最多约 X 字」同口径。

读法：
- **写作目标** = min(静态 max_length, 几何容量) —— 照它写必过校验；
- **校验放行线** = 目标 × 1.1 + 8 字（schema 容差），超放行线才被拒；
- **写作区间** = 静态下限（含方向红线密度下限）到写作目标的 ~90%；
- 条目数决定一切：先按大纲定每页条目数，再按对应行写，**写时别再改条目数**。
"""


def render_outline(outline_path: str, out_path: str) -> int:
    with open(outline_path, encoding="utf-8") as f:
        data = json.load(f)
    pages = [s for sec in data.get("sections", []) for s in sec.get("slides", [])]
    blocks = [HEADER, f"依据大纲：{outline_path}（{len(pages)} 个内容页，"
                     f"封面/目录/章节页由渲染端补齐）", ""]
    for i, sl in enumerate(pages, 1):
        blocks.append(_page_block(i, sl))
        blocks.append("")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))
    return len(pages)


def render_all(out_path: str) -> int:
    blocks = [HEADER.replace("（由 .pi/renderer/budgets.py 生成，勿手改）",
                             "（全版式参考表）"), ""]
    n = 0
    for st, cls in CLS.items():
        blocks.append(f"## {st}")
        blocks.append("")
        blocks.extend(_dynamic_lines(st))
        blocks.append("")
        blocks.append("**静态契约**：" + " ｜ ".join(_static_parts(cls)))
        blocks.append("")
        n += 1
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))
    return n


# ---------- 自证：照预算写必过校验 ----------

def _fill_to_target(st: str, n: int, cfg: dict, targets: dict) -> dict:
    """把合成实例的动态字段改成恰好等于写作目标字数。"""
    cfg = {k: v for k, v in cfg.items() if k != "n"}
    kw = _synth_kwargs(st, n, **cfg)
    if st == "numbered_cards" and "desc" in targets:
        for c in kw["cards"]:
            c["desc"] = _z(targets["desc"])
    elif st == "compare_3col" and "desc" in targets:
        for c in kw["columns"]:
            c["desc"] = _z(targets["desc"])
    elif st == "layered_stack" and "desc" in targets:
        for x in kw["layers"]:
            x["desc"] = _z(targets["desc"])
    elif st == "flow" and "desc" in targets:
        for x in kw["steps"]:
            x["desc"] = _z(targets["desc"])
    elif st == "hierarchy" and "desc" in targets:
        for x in kw["nodes"]:
            x["desc"] = _z(targets["desc"])
    elif st == "timeline" and "desc" in targets:
        for x in kw["milestones"]:
            x["desc"] = _z(targets["desc"])
    elif st == "funnel" and "desc" in targets:
        for x in kw["stages"]:
            x["desc"] = _z(targets["desc"])
    elif st == "matrix" and "desc" in targets:
        for row in kw["rows"]:
            for c in row["cells"]:
                c["desc"] = _z(targets["desc"])
    elif st == "key_groups" and "points" in targets:
        for g in kw["groups"]:
            g["points"] = [_z(targets["points"])] * 3
    elif st == "grid_2x2":
        for key in ("q_top_left", "q_top_right", "q_bottom_left", "q_bottom_right"):
            if "desc" in targets:
                kw[key]["desc"] = _z(targets["desc"])
            if "value" in targets:
                kw[key]["value"] = _z(min(14, targets["value"]))
    elif st == "statement" and "body" in targets:
        kw["body"] = _z(min(700, targets["body"]))
    return kw


def selfcheck() -> int:
    """对代表性配置验证：按写作目标填满 → 真实校验必须放行。退出码 0 = 全过。"""
    cases = [("numbered_cards", dict(n=4)), ("numbered_cards", dict(n=8)),
             ("compare_3col", dict(n=3)), ("compare_3col", dict(n=4, full_summary=True)),
             ("layered_stack", dict(n=6)), ("flow", dict(n=8)),
             ("flow", dict(n=5, ring=True)), ("hierarchy", dict(n=4)),
             ("timeline", dict(n=9)), ("funnel", dict(n=5)),
             ("matrix", dict(n=3, cols=3)), ("key_groups", dict(n=3)),
             ("grid_2x2", dict(n=1)), ("statement", dict(n=1)),
             ("statement", dict(n=1, right_rail=True))]
    fails = []
    for st, cfg in cases:
        n = cfg["n"]
        # 写作目标 = min(静态上限, cap)，与 _dynamic_lines 同一 helper
        targets = {}
        for attr, cap in _probe(st, **cfg).items():
            targets[attr] = _target(st, attr, cap)
        if st == "key_groups":
            targets["points"] = min(_key_groups_caps(n, cfg.get("aux", False)))
            caps = {"points": targets["points"]}
        kw = _fill_to_target(st, n, cfg, targets)
        try:
            CLS[st].model_validate(kw)
            ok = "PASS"
        except Exception as e:  # noqa: BLE001
            ok = "FAIL"
            fails.append((st, cfg, str(e)[:120]))
        print(f"[{ok}] {st} {cfg} 目标={targets}")
    for st, cfg, msg in fails:
        print(f"  失败详情 {st} {cfg}: {msg}", file=sys.stderr)
    if fails:
        print(f"SELFCHECK: {len(fails)} 例失败", file=sys.stderr)
        return 1
    print(f"SELFCHECK: 全部 {len(cases)} 例通过 —— 照预算表写作目标写，校验必放行")
    return 0


def main():
    ap = argparse.ArgumentParser(description="字数预算表：写作前下发动态容量")
    ap.add_argument("outline", nargs="?", help="outline.json 路径")
    ap.add_argument("-o", "--out", default=".pi/work/budgets.md",
                    help="预算表输出路径（默认 .pi/work/budgets.md）")
    ap.add_argument("--all", action="store_true", help="不读大纲，输出全部版式参考表")
    ap.add_argument("--selfcheck", action="store_true",
                    help="自证「照预算写必过校验」，退出码 0 = 通过")
    args = ap.parse_args()
    if args.selfcheck:
        sys.exit(selfcheck())
    if args.all:
        n = render_all(args.out)
        print(f"全版式参考表已写入 {args.out}（{n} 个版式）")
        return
    if not args.outline:
        ap.error("需要 outline.json 路径（或 --all / --selfcheck）")
    n = render_outline(args.outline, args.out)
    print(f"预算表已写入 {args.out}（{n} 页）。写每批 slides 前先读对应页的预算段。")


if __name__ == "__main__":
    main()
