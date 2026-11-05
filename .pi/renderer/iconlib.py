# -*- coding: utf-8 -*-
"""SVG 图标引擎：把离线 SVG 图标库转成 PowerPoint 自定义几何图形。

一个图标库 / 两个风格（共 4620 个图标），全部离线：

    .pi/assets/icons/<lib>/<style>/*.svg
      tabler    outline(3646) filled(974)

**只留 tabler 一库是有意的**（2026-10 收敛，L2 档）：原来 lucide / feather /
heroicons 四库并存，笔画宽度（2.0 vs 1.5）和图形占框比例（16.9/24 vs 18.4/24）
各不相同，混着用会让同一页的图标一半粗一半细、一半大一半小；同名图标还会让
检索结果四份重复。裁剪口径与工具见 `.pi/tests/icon_trim.py`，删除记录在
`.pi/work/icon-trim-record.json`。

两条渲染路线（由 SVG 根属性实测决定，不靠猜）：
- **描边型**（outline）：`fill="none"` + `stroke-width`
  → DrawingML `<a:noFill/>` + `<a:ln w=...>`
- **填充型**（filled）：`fill="currentColor"`
  → DrawingML `<a:solidFill/>`，不描边

`fill-rule="evenodd"` 在 DrawingML 里没有对应属性，本引擎用
**子路径绕向修正**（按嵌套深度做奇偶反转）把 even-odd 语义等价转换为 nonzero 语义。

为什么转 custGeom：
- 真矢量（`<a:custGeom>` + `<a:pathLst>`），PPT 内可改色、缩放不失真，不是图片
- 零外部资产、完全离线

命名解析（写入 DSL 的 `icon` 字段）：
    brain                 → 按 index.json 的 resolve_order 依次查找
    tabler:brain          → 在 tabler 内查找（outline 优先）
    tabler:filled/brain   → 精确定位到某个风格
    齿轮                   → 先查 .pi/assets/icons/aliases.json 语义别名

命令行：
    python .pi/renderer/iconlib.py 齿轮
    python .pi/renderer/iconlib.py --search brain --json --limit 20
    python .pi/renderer/iconlib.py --libs
    python .pi/renderer/iconlib.py --stats
    python .pi/renderer/iconlib.py --show 齿轮
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from typing import Optional

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICON_ROOT = os.path.join(ROOT, "assets", "icons")
INDEX_PATH = os.path.join(ICON_ROOT, "index.json")
ALIASES_PATH = os.path.join(ICON_ROOT, "aliases.json")

SKIP_DIRS = {"LICENSES", "__pycache__"}

# 裸名解析优先级：先描边型（视觉更轻），再到填充型
DEFAULT_ORDER = ["tabler/outline", "tabler/filled"]

# 描边宽度缩放：SVG 原始 stroke-width 按各自 viewBox 映射到目标尺寸后再削 35%，
# 视觉上更接近现代图标库的精细度。
STROKE_SCALE = 0.65

# ---------------------------------------------------------------- 别名
ALIASES: dict[str, str] = {}


def _load_aliases() -> dict[str, str]:
    global ALIASES
    if ALIASES:
        return ALIASES
    try:
        with open(ALIASES_PATH, encoding="utf-8") as f:
            ALIASES = {k: v for k, v in json.load(f).get("aliases", {}).items()}
    except Exception:
        ALIASES = {}
    return ALIASES


_INDEX_META: dict = {}


def index_meta() -> dict:
    """读取 index.json 的元信息（库 / 风格 / resolve_order / 关键词）。"""
    global _INDEX_META
    if _INDEX_META:
        return _INDEX_META
    try:
        with open(INDEX_PATH, encoding="utf-8") as f:
            _INDEX_META = json.load(f)
    except Exception:
        _INDEX_META = {}
    return _INDEX_META


def resolve_order() -> list[str]:
    order = (index_meta().get("naming") or {}).get("resolve_order")
    return list(order) if order else list(DEFAULT_ORDER)


def icon_tags() -> dict[str, dict[str, list[str]]]:
    """{'lib/style': {name: [关键词, ...]}}"""
    return index_meta().get("icons") or {}


def _style_modes() -> dict[str, str]:
    out = {}
    for lib, meta in (index_meta().get("libraries") or {}).items():
        for style, sm in (meta.get("styles") or {}).items():
            out[f"{lib}/{style}"] = sm.get("mode", "stroke")
    return out


# ---------------------------------------------------------------- 目录扫描
_SCAN: Optional[dict[str, str]] = None           # "lib/style/name" -> 绝对路径
_BY_NAME: Optional[dict[str, list[str]]] = None  # name -> ["lib/style/name", ...]


def _scan() -> dict[str, str]:
    global _SCAN, _BY_NAME
    if _SCAN is not None:
        return _SCAN
    table: dict[str, str] = {}
    by_name: dict[str, list[str]] = {}
    if os.path.isdir(ICON_ROOT):
        for lib in sorted(os.listdir(ICON_ROOT)):
            lib_dir = os.path.join(ICON_ROOT, lib)
            if not os.path.isdir(lib_dir) or lib in SKIP_DIRS or lib.startswith("_"):
                continue
            for style in sorted(os.listdir(lib_dir)):
                style_dir = os.path.join(lib_dir, style)
                if not os.path.isdir(style_dir):
                    continue
                for f in sorted(os.listdir(style_dir)):
                    if not f.endswith(".svg"):
                        continue
                    spec = f"{lib}/{style}/{f[:-4]}"
                    table[spec] = os.path.join(style_dir, f)
                    by_name.setdefault(f[:-4], []).append(spec)
    _SCAN, _BY_NAME = table, by_name
    return _SCAN


def libraries() -> dict[str, list[str]]:
    """{'tabler': ['filled', 'outline'], ...}"""
    out: dict[str, list[str]] = {}
    for spec in _scan():
        lib, style, _ = spec.split("/", 2)
        out.setdefault(lib, [])
        if style not in out[lib]:
            out[lib].append(style)
    for v in out.values():
        v.sort()
    return dict(sorted(out.items()))


def group_counts() -> dict[str, int]:
    out: dict[str, int] = {}
    for spec in _scan():
        key = spec.rsplit("/", 1)[0]
        out[key] = out.get(key, 0) + 1
    return dict(sorted(out.items()))


# ---------------------------------------------------------------- 名称解析
def _pick(style_list: list[str]) -> list[str]:
    """同一库内风格优先级：outline > lab > 其余按名。"""
    pref = {"outline": 0, "lab": 1, "filled": 2, "solid": 2}
    return sorted(style_list, key=lambda s: (pref.get(s, 5), s))


def resolve_icon(name: str) -> Optional[str]:
    """规格 -> 'lib/style/name'。找不到返回 None。"""
    if not name:
        return None
    table = _scan()
    if not table:
        return None
    raw = name.strip()
    key = raw.lower()

    # 1) 语义别名（中文原名 / 小写）
    aliases = _load_aliases()
    alias = aliases.get(raw) or aliases.get(key)
    if alias:
        hit = _resolve_spec(alias, table)
        if hit:
            return hit

    # 2) 直接按规格解析
    hit = _resolve_spec(raw, table)
    if hit:
        return hit

    # 3) 子串检索（最接近的短名优先）
    hits = [s for s in table if key in s.split("/", 2)[2]]
    if hits:
        return sorted(hits, key=lambda s: (len(s.split("/")[2]), s))[0]
    return None


def _resolve_spec(spec: str, table: dict[str, str]) -> Optional[str]:
    s = spec.strip()
    if not s:
        return None
    if ":" in s:                                     # lib:name 或 lib/style:name
        lib, rest = s.split(":", 1)
        if "/" in rest:
            style, nm = rest.split("/", 1)
            full = f"{lib}/{style}/{nm}"
            return full if full in table else None
        styles = [k.split("/")[1] for k in table if k.startswith(lib + "/")]
        for st in _pick(styles):
            full = f"{lib}/{st}/{rest}"
            if full in table:
                return full
        return None
    if "/" in s:                                     # style/name 或 lib/name 或 lib/style/name
        full = s.replace("\\", "/")
        if full in table:
            return full
        parts = full.split("/")
        if len(parts) == 3:                          # 形如 lib/style/name 但风格名不匹配
            return None
        if len(parts) == 2:
            head, nm = parts
            for k in table:                          # ① head 作风格名
                lib, st, n = k.split("/", 2)
                if st == head and n == nm:
                    return k
            styles = [k.split("/")[1] for k in table if k.startswith(head + "/")]
            for st in _pick(styles):                 # ② head 作库名
                cand = f"{head}/{st}/{nm}"
                if cand in table:
                    return cand
        return None
    low = s.lower()
    for key in resolve_order():                      # 裸名：按优先级
        cand = f"{key}/{low}"
        if cand in table:
            return cand
    for key in table:                                # 兜底：任意库
        if key.endswith("/" + low):
            return key
    return None


def search_icons(keyword: str, limit: int = 30, lib: str | None = None) -> list[dict]:
    """按名称 / 关键词 / 别名检索。返回 [{spec, icon, lib, style, mode, score}, ...]"""
    table = _scan()
    if not table:
        return []
    kw = (keyword or "").strip().lower()
    tags = icon_tags()
    modes = _style_modes()
    scored: dict[str, int] = {}

    def bump(spec: str, score: int):
        if spec not in table:
            return
        if lib and not spec.startswith(lib + "/"):
            return
        if score > scored.get(spec, 0):
            scored[spec] = score

    if kw:
        for spec in table:
            n = spec.split("/", 2)[2]
            if n == kw:
                bump(spec, 100)
            elif n.startswith(kw):
                bump(spec, 80)
            elif kw in n:
                bump(spec, 60)
            elif kw in spec.split("/", 2)[1]:
                bump(spec, 20)
        for group, name_tags in tags.items():
            for n, kws in name_tags.items():
                if any(kw == t for t in kws):
                    bump(f"{group}/{n}", 70)
                elif any(kw in t for t in kws):
                    bump(f"{group}/{n}", 45)
        for alias_key, target in _load_aliases().items():
            if kw == alias_key.lower():
                hit = _resolve_spec(target, table)
                if hit:
                    bump(hit, 90)
            elif kw in alias_key.lower():
                hit = _resolve_spec(target, table)
                if hit:
                    bump(hit, 40)
    else:
        for spec in table:
            bump(spec, 10)

    items = []
    for spec, score in scored.items():
        l, st, n = spec.split("/", 2)
        items.append({"spec": spec, "icon": n, "lib": l, "style": st,
                      "mode": modes.get(f"{l}/{st}", "stroke"), "score": score})
    order = {g: i for i, g in enumerate(resolve_order())}
    # 同分时：先按解析优先级（与 resolve_icon 一致），再按名字长度与字母序
    items.sort(key=lambda d: (-d["score"],
                              order.get(f"{d['lib']}/{d['style']}", 99),
                              len(d["icon"]), d["spec"]))
    return items[:limit]


# ---------------------------------------------------------------- SVG 解析
_NUM = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
_CMDS = set("MmLlHhVvCcSsQqTtAaZz")
_TAG = re.compile(r"<(path|circle|rect|line|polyline|polygon|ellipse)\b([^>]*?)/?>", re.I)
_ATTR = re.compile(r"([\w:-]+)\s*=\s*\"([^\"]*)\"")
_ROOT = re.compile(r"<svg\b[^>]*>", re.I | re.S)


def _attrs(s: str) -> dict:
    return {k.lower(): v for k, v in _ATTR.findall(s)}


def _elem_to_d(tag: str, a: dict) -> str:
    t = tag.lower()
    if t == "path":
        return a.get("d", "")
    if t == "circle":
        cx, cy, r = float(a.get("cx", 0)), float(a.get("cy", 0)), float(a.get("r", 0))
        return f"M{cx - r},{cy}a{r},{r} 0 1,0 {2 * r},0a{r},{r} 0 1,0 {-2 * r},0z"
    if t == "ellipse":
        cx, cy = float(a.get("cx", 0)), float(a.get("cy", 0))
        rx, ry = float(a.get("rx", 0)), float(a.get("ry", 0))
        return (f"M{cx - rx},{cy}a{rx},{ry} 0 1,0 {2 * rx},0"
                f"a{rx},{ry} 0 1,0 {-2 * rx},0z")
    if t == "line":
        return f"M{a.get('x1', 0)},{a.get('y1', 0)}L{a.get('x2', 0)},{a.get('y2', 0)}"
    if t in ("polyline", "polygon"):
        pts = [p for p in re.split(r"[ ,]+", a.get("points", "").strip()) if p]
        if len(pts) < 4:
            return ""
        d = "M" + " ".join(f"{pts[i]},{pts[i + 1]}" for i in range(0, len(pts) - 1, 2))
        return d + ("z" if t == "polygon" else "")
    if t == "rect":
        x, y = float(a.get("x", 0)), float(a.get("y", 0))
        w, h = float(a.get("width", 0)), float(a.get("height", 0))
        rx = float(a.get("rx", a.get("ry", 0)) or 0)
        if rx <= 0:
            return f"M{x},{y}h{w}v{h}h{-w}z"
        return (f"M{x + rx},{y}h{w - 2 * rx}a{rx},{rx} 0 0 1 {rx},{rx}"
                f"v{h - 2 * rx}a{rx},{rx} 0 0 1 {-rx},{rx}h{-(w - 2 * rx)}"
                f"a{rx},{rx} 0 0 1 {-rx},{-rx}v{-(h - 2 * rx)}"
                f"a{rx},{rx} 0 0 1 {rx},{-rx}z")
    return ""


def parse_svg(text: str) -> dict:
    """解析 SVG，按元素实际样式分成 filled / stroked 两组。"""
    root = _ROOT.search(text)
    ra = _attrs(root.group(0)) if root else {}
    r_fill = (ra.get("fill") or "").strip()
    r_stroke = (ra.get("stroke") or "").strip()
    r_rule = (ra.get("fill-rule") or "").strip()
    r_sw = ra.get("stroke-width")
    try:
        stroke_w = float(r_sw) if r_sw else None
    except ValueError:
        stroke_w = None

    viewbox = 24.0
    vb = re.search(r"viewBox\s*=\s*\"([^\"]+)\"", text)
    if vb:
        parts = vb.group(1).split()
        if len(parts) == 4:
            try:
                viewbox = float(parts[2])
            except ValueError:
                pass

    fill_d, stroke_d, evenodd = [], [], False
    for tag, attr_str in _TAG.findall(text):
        a = _attrs(attr_str)
        d = _elem_to_d(tag, a)
        if not d:
            continue
        fill = (a.get("fill") or r_fill).strip()
        stroke = (a.get("stroke") or r_stroke).strip()
        if (a.get("fill-rule") or r_rule).strip().lower() == "evenodd":
            evenodd = True
        has_fill = fill not in ("", "none")
        has_stroke = stroke not in ("", "none")
        if has_fill:
            fill_d.append(d)
        if has_stroke:
            stroke_d.append(d)
        if not has_fill and not has_stroke:
            stroke_d.append(d)

    # mode 只用于展示/自检，不参与出图决策（build_icon_xml 直接看 fill/stroke 分组）。
    # 一个可识别的绘制元素都没有时必须报 "none"——旧写法会算成 "stroke"，
    # 于是「文件里根本没有可画的东西」和「有元素但解析成空」被混成同一类，
    # tools/check_icons.py 的全库几何自检就分不出这两种根因了。
    if not fill_d and not stroke_d:
        mode = "none"
    else:
        mode = "mixed" if (fill_d and stroke_d) else ("fill" if fill_d else "stroke")
    return {"fill": fill_d, "stroke": stroke_d, "viewbox": viewbox,
            "stroke_width": stroke_w, "mode": mode, "evenodd": evenodd}


def _skip_ws(d: str, i: int) -> int:
    while i < len(d) and (d[i].isspace() or d[i] == ","):
        i += 1
    return i


def _read_num(d: str, i: int):
    i = _skip_ws(d, i)
    m = _NUM.match(d, i)
    if not m:
        raise ValueError(f"path 数据在位置 {i} 处不是数字：{d[i:i + 14]!r}")
    return float(m.group()), m.end()


def _read_flag(d: str, i: int):
    """圆弧的 large-arc-flag / sweep-flag。

    SVG 规范：这两个是**单字符** 0/1，因此允许其后紧跟下一个数字而不加分隔符，
    例如 `a5 5 0 0110 10` ≡ `a5 5 0 0 1 10 10`。
    绝不能用「按空格切分再定长取 7 个」的写法——那样标志位会和后面的数字
    粘连成一个数，参数整体错位，轻则把圆弧算飞，重则 int('A') 直接抛异常。
    """
    i = _skip_ws(d, i)
    if i >= len(d) or d[i] not in "01":
        raise ValueError(f"path 数据在位置 {i} 处不是圆弧标志位：{d[i:i + 14]!r}")
    return int(d[i]), i + 1


def _arc_cubic(x1, y1, rx, ry, phi_deg, fa, fs, x2, y2):
    """SVG 圆弧 -> 三次贝塞尔近似（标准算法）。"""
    if rx == 0 or ry == 0:
        return [("L", (x2, y2))]
    phi = math.radians(phi_deg)
    cos_p, sin_p = math.cos(phi), math.sin(phi)
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cos_p * dx + sin_p * dy, -sin_p * dx + cos_p * dy
    lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry)
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s
        ry *= s
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    co = math.sqrt(max(0, num / den)) if den else 0
    if fa == fs:
        co = -co
    cxp, cyp = co * rx * y1p / ry, -co * ry * x1p / rx
    cx = cos_p * cxp - sin_p * cyp + (x1 + x2) / 2
    cy = sin_p * cxp + cos_p * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        dd = math.hypot(ux, uy) * math.hypot(vx, vy)
        c = max(-1, min(1, (ux * vx + uy * vy) / dd)) if dd else 1
        a = math.acos(c)
        return -a if (ux * vy - uy * vx) < 0 else a

    th1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dth = ang((x1p - cxp) / rx, (y1p - cyp) / ry,
              (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not fs and dth > 0:
        dth -= 2 * math.pi
    elif fs and dth < 0:
        dth += 2 * math.pi
    segs = max(1, int(math.ceil(abs(dth) / (math.pi / 2))))
    delta = dth / segs
    t = 4 / 3 * math.tan(delta / 4)
    out = []
    for i in range(segs):
        th = th1 + i * delta

        def e(a):
            # W3C arc->bezier：x = cx + rx·cosθ·cosφ − ry·sinθ·sinφ
            #                  y = cy + rx·cosθ·sinφ + ry·sinθ·cosφ
            # （原实现把 y 的 rx/ry 写反了，正圆看不出来，椭圆会被画成大穹顶）
            return (cx + rx * math.cos(a) * cos_p - ry * math.sin(a) * sin_p,
                    cy + rx * math.cos(a) * sin_p + ry * math.sin(a) * cos_p)

        p0, p3 = e(th), e(th + delta)
        # 切线 dy/dθ = −rx·sinθ·sinφ + ry·cosθ·cosφ
        d1 = (-rx * math.sin(th) * cos_p - ry * math.cos(th) * sin_p,
              -rx * math.sin(th) * sin_p + ry * math.cos(th) * cos_p)
        d2 = (-rx * math.sin(th + delta) * cos_p - ry * math.cos(th + delta) * sin_p,
              -rx * math.sin(th + delta) * sin_p + ry * math.cos(th + delta) * cos_p)
        out.append(("C", (p0[0] + t * d1[0], p0[1] + t * d1[1]),
                    (p3[0] - t * d2[0], p3[1] - t * d2[1]), p3))
    return out


def _parse_d(d: str):
    """path 数据 -> 子路径列表；每个子路径是 [('M',pt), ('L',pt), ('C',c1,c2,pt), ('Z',)]。

    按字符游标逐命令解析，而不是「先切成 token 再定长取参数」。
    原因见 `_read_flag`：圆弧的两个标志位是单字符、可与后一个数字紧贴，
    定长切片会把参数整体读错。
    """
    subs, cur = [], []
    i, cmd = 0, None
    cx, cy, sx, sy = 0.0, 0.0, 0.0, 0.0
    last_c = last_q = None

    while True:
        i = _skip_ws(d, i)
        if i >= len(d):
            break
        ch = d[i]
        if ch in _CMDS:
            cmd = ch
            i += 1
            if cmd in "Zz":
                cur.append(("Z",))
                subs.append(cur)
                cur = []
                cx, cy = sx, sy
                continue
        elif cmd is None:
            raise ValueError(f"path 数据未以命令字母开头：{d[:20]!r}")

        rel = cmd.islower()
        C = cmd.upper()
        if C == "M":
            x, i = _read_num(d, i)
            y, i = _read_num(d, i)
            cx, cy = (cx + x, cy + y) if rel else (x, y)
            sx, sy = cx, cy
            cur.append(("M", (cx, cy)))
            cmd = "l" if rel else "L"      # 规范：M 之后的隐式命令是 L
            last_c = last_q = None
        elif C == "L":
            x, i = _read_num(d, i)
            y, i = _read_num(d, i)
            cx, cy = (cx + x, cy + y) if rel else (x, y)
            cur.append(("L", (cx, cy)))
            last_c = last_q = None
        elif C == "H":
            x, i = _read_num(d, i)
            cx = cx + x if rel else x
            cur.append(("L", (cx, cy)))
            last_c = last_q = None
        elif C == "V":
            y, i = _read_num(d, i)
            cy = cy + y if rel else y
            cur.append(("L", (cx, cy)))
            last_c = last_q = None
        elif C == "C":
            v = []
            for _ in range(6):
                q, i = _read_num(d, i)
                v.append(q)
            if rel:
                c1 = (cx + v[0], cy + v[1])
                c2 = (cx + v[2], cy + v[3])
                e = (cx + v[4], cy + v[5])
            else:
                c1, c2, e = (v[0], v[1]), (v[2], v[3]), (v[4], v[5])
            cur.append(("C", c1, c2, e))
            last_c, last_q, cx, cy = c2, None, e[0], e[1]
        elif C == "S":
            v = []
            for _ in range(4):
                q, i = _read_num(d, i)
                v.append(q)
            r1 = (2 * cx - last_c[0], 2 * cy - last_c[1]) if last_c else (cx, cy)
            c2 = (cx + v[0], cy + v[1]) if rel else (v[0], v[1])
            e = (cx + v[2], cy + v[3]) if rel else (v[2], v[3])
            cur.append(("C", r1, c2, e))
            last_c, last_q, cx, cy = c2, None, e[0], e[1]
        elif C == "Q":
            v = []
            for _ in range(4):
                q, i = _read_num(d, i)
                v.append(q)
            q = (cx + v[0], cy + v[1]) if rel else (v[0], v[1])
            e = (cx + v[2], cy + v[3]) if rel else (v[2], v[3])
            c1 = (cx + 2 / 3 * (q[0] - cx), cy + 2 / 3 * (q[1] - cy))
            c2 = (e[0] + 2 / 3 * (q[0] - e[0]), e[1] + 2 / 3 * (q[1] - e[1]))
            cur.append(("C", c1, c2, e))
            last_q, last_c, cx, cy = q, None, e[0], e[1]
        elif C == "T":
            x, i = _read_num(d, i)
            y, i = _read_num(d, i)
            q = (2 * cx - last_q[0], 2 * cy - last_q[1]) if last_q else (cx, cy)
            e = (cx + x, cy + y) if rel else (x, y)
            c1 = (cx + 2 / 3 * (q[0] - cx), cy + 2 / 3 * (q[1] - cy))
            c2 = (e[0] + 2 / 3 * (q[0] - e[0]), e[1] + 2 / 3 * (q[1] - e[1]))
            cur.append(("C", c1, c2, e))
            last_q, last_c, cx, cy = q, None, e[0], e[1]
        elif C == "A":
            rx, i = _read_num(d, i)
            ry, i = _read_num(d, i)
            rot, i = _read_num(d, i)
            fa, i = _read_flag(d, i)       # 单字符，可能与后一个数字紧贴
            fs, i = _read_flag(d, i)
            x, i = _read_num(d, i)
            y, i = _read_num(d, i)
            x2, y2 = (cx + x, cy + y) if rel else (x, y)
            for seg in _arc_cubic(cx, cy, rx, ry, rot, fa, fs, x2, y2):
                cur.append(seg)
            cx, cy = x2, y2
            last_c = last_q = None
        else:
            raise ValueError(f"不支持的 path 命令：{cmd!r}")

    if cur:
        subs.append(cur)
    return [s for s in subs if len(s) > 1]


# ------------------------------------------------ even-odd -> nonzero 等价转换
def _flatten(sub) -> list[tuple[float, float]]:
    pts = []
    for seg in sub:
        if seg[0] in ("M", "L"):
            pts.append(seg[1])
        elif seg[0] == "C":
            pts.extend(seg[1:4])
    return pts


def _signed_area(sub) -> float:
    pts = _flatten(sub)
    if len(pts) < 3:
        return 0.0
    s = 0.0
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return s / 2.0


def _point_in_poly(poly, pt) -> bool:
    x, y = pt
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        if (y1 > y) != (y2 > y):
            if x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
                inside = not inside
    return inside


def _reverse_subpath(sub):
    """反转子路径绕向（形状不变，只改方向）。"""
    closed = bool(sub) and sub[-1][0] == "Z"
    segs = sub[:-1] if closed else list(sub)
    if not segs or segs[0][0] != "M":
        return sub
    start = segs[0][1]
    walk = []
    for s in segs[1:]:
        if s[0] == "L":
            walk.append(("L", s[1], None))
        elif s[0] == "C":
            walk.append(("C", s[3], (s[1], s[2])))
    if not walk:
        return sub
    out = [("M", walk[-1][1])]
    for i in range(len(walk) - 1, -1, -1):
        cur = walk[i]
        prev_end = start if i == 0 else walk[i - 1][1]
        if cur[0] == "L":
            out.append(("L", prev_end))
        else:
            c1, c2 = cur[2]
            out.append(("C", c2, c1, prev_end))
    if closed:
        out.append(("Z",))
    return out


def fix_evenodd(subs: list) -> list:
    """DrawingML 没有 fill-rule，用绕向修正把 even-odd 语义转成 nonzero。

    规则：子路径按嵌套深度分层，第 d 层的绕向应为最外层祖先绕向的 (-1)^d。
    """
    if len(subs) < 2:
        return subs
    polys = [_flatten(s) for s in subs]
    areas = [_signed_area(s) for s in subs]
    n = len(subs)
    parent: list[Optional[int]] = [None] * n
    depth = [0] * n
    for i in range(n):
        if not polys[i]:
            continue
        for j in range(n):
            if i == j or not polys[j]:
                continue
            if abs(areas[j]) <= abs(areas[i]):
                continue
            if _point_in_poly(polys[j], polys[i][0]):
                depth[i] += 1
                if parent[i] is None or abs(areas[j]) < abs(areas[parent[i]]):
                    parent[i] = j
    out = []
    for i, s in enumerate(subs):
        if areas[i] == 0:
            out.append(s)
            continue
        base = parent[i]
        while base is not None and parent[base] is not None:
            base = parent[base]
        s0 = areas[base] if base is not None else areas[i]
        if s0 == 0:
            s0 = areas[i]
        target = (1 if s0 > 0 else -1) * ((-1) ** depth[i])
        cur = 1 if areas[i] > 0 else -1
        out.append(_reverse_subpath(s) if cur != target else s)
    return out


# ---------------------------------------------------------------- 加载与缓存
_CACHE: dict[str, dict] = {}


def _load(spec: str) -> dict:
    hit = _CACHE.get(spec)
    if hit is not None:
        return hit
    path = _scan().get(spec)
    missing = {"fill": [], "stroke": [], "fill_groups": [], "stroke_groups": [],
               "viewbox": 24.0, "stroke_width": None,
               "mode": "none", "evenodd": False, "path": None}
    if not path or not os.path.exists(path):
        _CACHE[spec] = missing
        return missing
    with open(path, encoding="utf-8") as f:
        info = parse_svg(f.read())
    # 每个源 <path> 元素单独成组：绕向修正只在组内做（even-odd 是元素级语义），
    # 组间保持独立 <a:path>，让 PowerPoint 取并集。详见 _path_xml 的说明。
    fill_groups, stroke_groups = [], []
    bad = []

    def _subs_of(d):
        """解析单个元素的 d；失败时**报警但不中断出图**。

        _parse_d 现在会抛异常而不是像以前那样静默截断——静默截断的后果是
        图标画不出来却没有任何提示（lucide/outline/zap 就是这样变成空白的）。
        所以这里把错误显式打到 stderr，并把出错的 spec 记进 info["parse_errors"]，
        交给 tools/check_icons.py 去兜。
        """
        try:
            return _parse_d(d)
        except ValueError as e:
            bad.append(str(e))
            print(f"[iconlib] 图标解析失败 {spec}: {e}", file=sys.stderr)
            return []

    for d in info["fill"]:
        subs = _subs_of(d)
        if info["evenodd"]:
            subs = fix_evenodd(subs)
        if subs:
            fill_groups.append(subs)
    for d in info["stroke"]:
        subs = _subs_of(d)
        if info["evenodd"]:
            subs = fix_evenodd(subs)
        if subs:
            stroke_groups.append(subs)
    info["parse_errors"] = bad
    info["fill_groups"], info["stroke_groups"] = fill_groups, stroke_groups
    info["fill"] = [s for g in fill_groups for s in g]
    info["stroke"] = [s for g in stroke_groups for s in g]
    info["path"] = path
    _CACHE[spec] = info
    return info


# ---------------------------------------------------------------- 绘制
def _next_id(spTree) -> int:
    from pptx.oxml.ns import qn
    mx = 1
    for el in spTree.iter(qn("p:cNvPr")):
        try:
            mx = max(mx, int(el.get("id")))
        except Exception:
            pass
    return mx + 1


_NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')


def _path_xml(groups, k: float, size_emu: int, fill_mode: str) -> str:
    """把子路径按「源 <path> 元素」分组，每组产出一个 <a:path>。

    分组的依据是 DrawingML 的两条特性：

    - **同一个 <a:path> 内**所有子路径共享一条 winding rule。DrawingML 没有
      fill-rule，SVG 的 even-odd 语义已由 `fix_evenodd` 转成反向绕行，
      因此**同一元素**的子路径必须合并进同一个 <a:path>，孔才挖得出来
      （反例：tabler filled/circle-check 拆成两个 <a:path> 后孔会被填死）。
    - **不同 <a:path> 之间**PowerPoint 各自独立填充后取并集，正好对应 SVG
      「每个 <path> 元素独立填充、按文档顺序叠加」的语义。若把不同元素也并进
      一个 <a:path>，两个元素绕向相反时会凭空多出一个孔
      （反例：tabler filled/arrow-big-left-line，箭头主体被挖空）。

    参数 groups：[ [子路径, ...], ... ]，每个内层列表对应一个源元素。
    """
    if groups and groups[0] and isinstance(groups[0][0], tuple):
        # 兼容直接传入「扁平子路径列表」的旧调用形式（段的第一个元素是 op 字符串）
        groups = [groups]

    paths = []
    for subs in groups:
        segs = []
        for sub in subs:
            for seg in sub:
                op = seg[0]
                if op == "M":
                    segs.append('<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>'
                                % (seg[1][0] * k, seg[1][1] * k))
                elif op == "L":
                    segs.append('<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>'
                                % (seg[1][0] * k, seg[1][1] * k))
                elif op == "C":
                    segs.append('<a:cubicBezTo><a:pt x="%d" y="%d"/><a:pt x="%d" y="%d"/>'
                                '<a:pt x="%d" y="%d"/></a:cubicBezTo>'
                                % (seg[1][0] * k, seg[1][1] * k, seg[2][0] * k, seg[2][1] * k,
                                   seg[3][0] * k, seg[3][1] * k))
                elif op == "Z":
                    segs.append("<a:close/>")
        if segs:
            paths.append('<a:path w="%d" h="%d" fill="%s">%s</a:path>'
                         % (size_emu, size_emu, fill_mode, "".join(segs)))
    return "".join(paths)


def build_icon_xml(name: str, x_emu: int, y_emu: int, size_emu: int, color_hex: str,
                   shape_id: int, stroke_w: float | None = None,
                   spec: str | None = None) -> list[str]:
    """返回若干 <p:sp> XML 字符串（填充型与描边型各一个形状）。"""
    spec = spec or resolve_icon(name)
    if not spec:
        return []
    info = _load(spec)
    if not info["fill"] and not info["stroke"]:
        return []
    k = size_emu / (info["viewbox"] or 24.0)
    out: list[str] = []

    def wrap(paths_xml: str, sp_id: int, fill: bool, w_emu: int = 0) -> str:
        if fill:
            style = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                     '<a:ln><a:noFill/></a:ln>' % color_hex)
        else:
            style = ('<a:noFill/><a:ln w="%d" cap="rnd"><a:solidFill>'
                     '<a:srgbClr val="%s"/></a:solidFill><a:round/></a:ln>'
                     % (max(1, w_emu), color_hex))
        return (
            '<p:sp %s>'
            '<p:nvSpPr><p:cNvPr id="%d" name="icon-%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="%d" b="%d"/><a:pathLst>%s</a:pathLst></a:custGeom>'
            '%s</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            % (_NS, sp_id, name.replace('"', ""), x_emu, y_emu, size_emu, size_emu,
               size_emu, size_emu, paths_xml, style))

    if info["fill"]:
        groups = info.get("fill_groups") or [info["fill"]]
        out.append(wrap(_path_xml(groups, k, size_emu, "norm"), shape_id, True))
    if info["stroke"]:
        sw = info["stroke_width"]
        w = int(stroke_w * k) if stroke_w else int((sw or 0) * k * STROKE_SCALE)
        groups = info.get("stroke_groups") or [info["stroke"]]
        out.append(wrap(_path_xml(groups, k, size_emu, "none"),
                        shape_id + len(out), False, w))
    return out


def draw_icon(slide, name: str, x: float, y: float, size: float,
              color: str = "primary", stroke_w: float | None = None,
              plate: bool = False, In=None) -> bool:
    """绘制 SVG 图标。坐标单位英寸。

    成功返回 True；图标找不到或解析失败返回 False（调用方需降级）。
    """
    if In is None:
        from pptx.util import Inches as In
    from pptx.oxml import parse_xml
    spec = resolve_icon(name)
    if not spec:
        return False
    try:
        from theme import COLORS
        hex_color = COLORS.get(color, color).lstrip("#").upper()
        if len(hex_color) != 6:
            hex_color = "2760A4"
        spTree = slide.shapes._spTree
        xmls = build_icon_xml(name, int(In(x)), int(In(y)), int(In(size)),
                              hex_color, _next_id(spTree), stroke_w=stroke_w, spec=spec)
        if not xmls:
            return False
        for xml in xmls:
            spTree.append(parse_xml(xml))
        return True
    except Exception:
        return False


def icon_info(name: str) -> dict | None:
    """调试用：某图标规格的解析详情。"""
    spec = resolve_icon(name)
    if not spec:
        return None
    info = _load(spec)
    lib, style, base = spec.split("/", 2)
    return {"query": name, "spec": spec, "icon": base, "lib": lib, "style": style,
            "mode": info["mode"], "viewbox": info["viewbox"],
            "stroke_width": info["stroke_width"],
            "subpaths_fill": len(info["fill"]), "subpaths_stroke": len(info["stroke"]),
            "groups_fill": len(info.get("fill_groups") or []),
            "groups_stroke": len(info.get("stroke_groups") or []),
            "parse_errors": list(info.get("parse_errors") or []),
            "evenodd_converted": bool(info.get("evenodd")), "path": info["path"]}


# ---------------------------------------------------------------- CLI
def _print_hits(items: list[dict], as_json: bool):
    if as_json:
        print(json.dumps({"count": len(items), "items": items}, ensure_ascii=False, indent=1))
        return
    for it in items:
        tag = {"fill": "实心", "mixed": "混合"}.get(it["mode"], "描边")
        print(f"  {it['icon']:<34} {it['lib']}/{it['style']:<9} [{tag}]  {it['spec']}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="离线图标库检索与自检")
    ap.add_argument("keyword", nargs="*", help="关键词（等价于 --search）")
    ap.add_argument("--search", "-s", help="关键词检索")
    ap.add_argument("--lib", help="限定图标库（当前只有 tabler）")
    ap.add_argument("--limit", type=int, default=30)
    ap.add_argument("--json", action="store_true", help="JSON 输出")
    ap.add_argument("--show", help="显示某个规格/别名的解析详情")
    ap.add_argument("--libs", action="store_true", help="列出图标库与风格")
    ap.add_argument("--stats", action="store_true", help="统计信息")
    args = ap.parse_args(argv)

    if args.libs:
        data = libraries()
        counts = group_counts()
        if args.json:
            print(json.dumps({"total": len(_scan()), "groups": counts},
                             ensure_ascii=False, indent=1))
        else:
            print(f"图标库（共 {len(_scan())} 个图标）：")
            for lib, styles in data.items():
                detail = ", ".join(f"{s}={counts.get(f'{lib}/{s}', 0)}" for s in styles)
                print(f"  {lib:<12} {detail}")
        return 0

    if args.stats:
        modes = _style_modes()
        counts = group_counts()
        meta = index_meta()
        rows = [{"group": g, "count": c, "mode": modes.get(g, "?")}
                for g, c in counts.items()]
        payload = {"root": ICON_ROOT, "total": len(_scan()),
                   "index_generated": meta.get("generated"),
                   "aliases": len(_load_aliases()),
                   "tagged_icons": sum(len(v) for v in icon_tags().values()),
                   "by_group": rows}
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=1))
        else:
            print(f"图标根目录 : {payload['root']}")
            print(f"图标总数   : {payload['total']}")
            print(f"索引生成于 : {payload['index_generated']}")
            print(f"语义别名   : {payload['aliases']} 条")
            print(f"带关键词   : {payload['tagged_icons']} 个")
            print("分组：")
            for r in rows:
                print(f"  {r['group']:<18} {r['count']:>5}  "
                      f"[{'实心' if r['mode'] == 'fill' else '描边'}]")
        return 0

    if args.show:
        info = icon_info(args.show)
        if not info:
            print(f"未找到：{args.show}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(info, ensure_ascii=False, indent=1))
        else:
            for k, v in info.items():
                print(f"  {k:<20} {v}")
        return 0

    kw = args.search or " ".join(args.keyword)
    items = search_icons(kw, args.limit, lib=args.lib)
    if not args.json:
        scope = f"（限定 {args.lib}）" if args.lib else ""
        print(f"关键词「{kw}」{scope}命中 {len(items)} 个 "
              f"/ 图标库共 {len(_scan())} 个（显示前 {args.limit}）：")
    _print_hits(items, args.json)
    return 0


if __name__ == "__main__":
    sys.exit(main())
