# -*- coding: utf-8 -*-
"""图标库裁剪：把 .pi/assets/icons 从「4 库 9931 个」收敛成「1 库、只留画得出、用得上」的集合。

跑法：
    python .pi/tests/icon_trim.py                 # 只出计划（默认 L2），不动文件
    python .pi/tests/icon_trim.py --level l1      # 轻档
    python .pi/tests/icon_trim.py --level l3      # 精选档
    python .pi/tests/icon_trim.py --apply         # 真正执行：移入 .pi/work/_icon_trash/ + 重建 index.json
    python .pi/tests/icon_trim.py --apply --verify-render   # 再渲染一遍范例合集确认没坏

三条删除口径（每条都写在 REASON 里，计划 JSON 可审计）：

1. **重复** —— 同一个图形换个库/换个网格再存一份。跨库同名的裸名本来就被
   resolve_order 挡住（tabler/outline 永远先命中），存着只会让 `search_icons`
   把同一个概念返回 4 遍（搜 database 的前 4 条是 4 个库的同名图标）。
2. **AI 味** —— 星芒/大脑/机器人/魔法棒/灯泡/火箭/笑脸。这类图标是"AI 生成感"的
   最大来源：任何一页放个 ✨ 都像素材站模板。
3. **粗糙** —— lucide/lab（官方实验区）、heroicons 的 16/20 网格（同图缩放版）、
   `-off` 禁用态、占画框不到一半的、以及只对 UI 工具栏有意义的类别
   （align/number/letter/math/device/brand…）。

保留集的底线是**先保护、再删除**：aliases.json 的每个目标、技能与 deck 里真实
出现过的每个名字、layouts 的兜底图标，一条都不许删 —— 删了会让页面静默掉图标
（见 .pi/docs/pitfalls.md 第 9 条）。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
ICON_ROOT = REPO / "assets" / "icons"
INDEX = ICON_ROOT / "index.json"
ALIASES = ICON_ROOT / "aliases.json"
TRASH = REPO / "work" / "_icon_trash"
PLAN_OUT = REPO / "work" / "icon-trim-plan.json"

# ---------------------------------------------------------------- 删除口径
AI_NAMES = {
    "sparkles", "sparkles-2", "star", "stars", "star-off", "star-half", "star-filled",
    "brain", "brain-cog", "robot", "robot-face", "robot-off",
    "ai", "ai-agent", "ai-agents", "ai-gateway",
    "magic-wand", "wand", "crystal-ball", "bulb", "bulb-off",
    "rocket", "rocket-off", "confetti", "zzz", "bubble-text",
}
AI_PREFIXES = ("sparkle", "stars")

# 只有 UI 工具栏才用得上的类别：deck 里不会出现"对齐方式""断开连接""阿拉伯字母"
NOISE_PREFIXES = ("mood-", "letter-", "number-", "math-",
                  "align-", "device-", "shopping-", "home-")

# 品牌 logo：整类删，只留汇报里可能真被引用的（其余是 Fortnite / Tinder /
# OnlyFans / 麦当劳这类消费与游戏品牌，放进客户 PPT 是事故）
KEEP_BRANDS = {
    "brand-wechat", "brand-qq", "brand-dingtalk", "brand-alipay", "brand-weibo",
    "brand-zhihu", "brand-miniprogram", "brand-github", "brand-docker",
    "brand-python", "brand-aws", "brand-azure", "brand-google", "brand-windows",
}

# 货币：留主流几种，其余 60 多种（bahraini / naira / …）是噪声
KEEP_CURRENCY = {
    "currency-yuan", "currency-dollar", "currency-euro", "currency-yen",
    "currency-pound", "currency-bitcoin", "currency-ethereum",
}

# L3 精选：这些前缀/名字是"业务汇报真会用到的语义"
L3_PREFIXES = (
    "chart-", "file-", "folder", "users", "user-", "shield", "settings", "database",
    "server", "cloud", "building", "coin", "alert-", "circle-check", "circle-x",
    "check", "clock", "calendar", "target", "flag", "route", "git-", "layout-",
    "stack", "arrow-up", "arrow-down", "arrow-right", "trending-", "gauge", "key",
    "lock", "unlock", "mail", "globe", "link", "list-", "search", "filter", "book",
    "briefcase", "cpu", "network", "plug", "puzzle", "receipt", "report", "school",
    "table", "tag", "tool", "truck", "wallet", "world", "api", "binary", "braces",
    "bug", "code", "columns", "compass", "components", "copy", "credit-card",
    "crown", "dashboard", "download", "upload", "edit", "eye", "hierarchy",
    "history", "info-circle", "layers", "leaf", "map-pin", "message", "messages",
    "notebook", "package", "paperclip", "pencil", "photo", "player-", "printer",
    "progress", "refresh", "repeat", "scale", "send", "share", "sitemap",
    "speakerphone", "test-pipe", "ticket", "timeline", "tools", "trash", "trophy",
    "versions", "volume", "zone",
)

# 保留档位说明：给人看的，不参与逻辑
LEVELS = {
    "l1": "轻：只删整库重复（solid-16/20 + 非 tabler 库里与 tabler 同名的）",
    "l2": "中（推荐）：收敛到 tabler 单库 + 删 AI 俗套与 UI 噪声",
    "l3": "激进：只留业务语义白名单（约一千个）",
}


# ---------------------------------------------------------------- 磁盘现状
def scan() -> dict[str, list[str]]:
    """{'lib/style': [name, ...]}"""
    out: dict[str, list[str]] = {}
    for lib in sorted(os.listdir(ICON_ROOT)):
        lib_dir = ICON_ROOT / lib
        if not lib_dir.is_dir() or lib in ("LICENSES", "__pycache__") or lib.startswith("_"):
            continue
        for style in sorted(os.listdir(lib_dir)):
            d = lib_dir / style
            if not d.is_dir():
                continue
            out[f"{lib}/{style}"] = sorted(f[:-4] for f in os.listdir(d) if f.endswith(".svg"))
    return out


def protected_names() -> set[str]:
    """不许删的名字：别名目标 + 技能与 deck 里出现的 + 兜底图标。"""
    keep: set[str] = set()
    data = json.loads(ALIASES.read_text(encoding="utf-8")).get("aliases", {})
    for v in data.values():
        v = v.strip()
        if ":" in v:
            v = v.split(":", 1)[1]
        if "/" in v:                       # lib/style/name 或 style/name
            v = v.rsplit("/", 1)[1]
        keep.add(v)

    pat = re.compile(r'["\']icon["\']\s*:\s*["\']([^"\']+)["\']')
    for base in (REPO / ".pi" / "skills", REPO / "work"):
        if not base.exists():
            continue
        for p in base.rglob("*.json"):
            try:
                txt = p.read_text(encoding="utf-8")
            except (OSError, UnicodeDecodeError):
                continue
            for m in pat.finditer(txt):
                v = m.group(1).strip()
                if ":" in v:
                    v = v.split(":", 1)[1]
                keep.add(v.rsplit("/", 1)[-1])

    # 别名里的中文词不能当图标名，过滤掉非 ASCII（它们只是键，不是文件名）
    return {n for n in keep if n and n.isascii()}


def level_rules(level: str, styles: dict[str, list[str]], protect: set[str]):
    """-> {spec/name: reason}"""
    drop: dict[str, str] = {}
    tabler = set(styles.get("tabler/outline", [])) | set(styles.get("tabler/filled", []))

    # ① 重复：整库同名 / 同图另存一份网格
    for st, names in styles.items():
        if not st.startswith("tabler/"):
            for n in names:
                if n in tabler:
                    # tabler 里还有一份，删这份不影响解析（同名裸名本来就轮到 tabler）
                    drop[f"{st}/{n}"] = "重复：与 tabler 同名，裸名永远被 tabler/outline 挡住"
                elif n in protect:
                    # 只此一份、且被别名/deck 引用 —— 留着，删了会静默掉图标
                    continue
                elif st in ("heroicons/solid-16", "heroicons/solid-20"):
                    drop[f"{st}/{n}"] = "重复：heroicons 同图的 16/20 网格版，缩放已由渲染器负责"
                elif st == "lucide/lab":
                    drop[f"{st}/{n}"] = "粗糙：lucide 官方实验区（lab）"
                elif level != "l1":
                    drop[f"{st}/{n}"] = "粗糙：非 tabler 库，笔画 1.5/2.0 与占框比例都和 tabler 不一致"

    if level == "l1":
        return drop

    # ② / ③ tabler 内部：AI 俗套 + UI 噪声 + 粗糙
    for st in ("tabler/outline", "tabler/filled"):
        for n in styles.get(st, []):
            key = f"{st}/{n}"
            if key in drop or n in protect:
                continue
            if n in AI_NAMES or n.startswith(AI_PREFIXES):
                drop[key] = "AI 味：星芒/大脑/机器人/魔法/灯泡/火箭/笑脸"
            elif n.startswith("brand-"):
                if n not in KEEP_BRANDS:
                    drop[key] = "噪声：第三方品牌 logo（消费/游戏品牌居多）"
            elif n.startswith("currency-"):
                if n not in KEEP_CURRENCY:
                    drop[key] = "噪声：非主流货币符号"
            elif n.startswith(NOISE_PREFIXES):
                drop[key] = "噪声：只对 UI 工具栏有意义的类别"
            elif n.endswith("-off"):
                drop[key] = "噪声：禁用态变体"

    if level != "l3":
        return drop

    # L3：只留白名单
    for st in ("tabler/outline", "tabler/filled"):
        for n in styles.get(st, []):
            key = f"{st}/{n}"
            if key in drop or n in protect:
                continue
            if not n.startswith(L3_PREFIXES):
                drop[key] = "L3 精选：不在业务语义白名单内"
    return drop


def rebuild_index(styles: dict[str, list[str]]) -> dict:
    """按磁盘现状过滤 index.json（保留每个图标的关键词，只重算计数与元信息）。"""
    idx = json.loads(INDEX.read_text(encoding="utf-8"))
    old_icons = idx.get("icons") or {}
    icons, libraries, order = {}, {}, []
    for st, names in styles.items():
        if not names:
            continue
        lib, style = st.split("/", 1)
        kept = set(names)
        tags = {n: t for n, t in (old_icons.get(st) or {}).items() if n in kept}
        icons[st] = dict(sorted(tags.items()))
        libs = libraries.setdefault(lib, {"license": "", "styles": {}})
        prev = ((idx.get("libraries") or {}).get(lib) or {}).get("styles", {}).get(style, {})
        libs["license"] = ((idx.get("libraries") or {}).get(lib) or {}).get("license", "")
        libs["styles"][style] = {
            "mode": prev.get("mode", "stroke"),
            "count": len(names),
            "viewBox": prev.get("viewBox", 24.0),
            "modes": prev.get("modes", {}),
            "strokeWidth": prev.get("strokeWidth"),
        }
        if prev.get("strokeWidth") is None:
            libs["styles"][style].pop("strokeWidth")

    naming = dict(idx.get("naming") or {})
    naming["resolve_order"] = [s for s in naming.get("resolve_order", []) if s in icons]
    style_mode = {}
    for _lib, meta in libraries.items():
        for _style, sm in meta.get("styles", {}).items():
            style_mode[_style] = sm.get("mode", "stroke")
    naming["style_mode"] = style_mode
    total = sum(len(v) for v in styles.values())
    out = dict(idx)
    out.update({"generated": "trimmed", "total": total, "icons": icons,
                "libraries": dict(sorted(libraries.items())), "naming": naming})
    return out


# ---------------------------------------------------------------- 校验
def verify(styles: dict[str, list[str]]) -> list[str]:
    """删除之后必须还成立的事情。返回问题清单（空 = 通过）。"""
    sys.path.insert(0, str(REPO / "renderer"))
    import importlib
    import iconlib
    importlib.reload(iconlib)          # 丢掉进程内的旧磁盘快照

    problems = []
    al = json.loads(ALIASES.read_text(encoding="utf-8")).get("aliases", {})
    for k, v in al.items():
        if not iconlib.resolve_icon(v):
            problems.append(f"别名失效：{k} -> {v}")

    for n in ("file-text", "layers", "target"):
        if not iconlib.resolve_icon(n):
            problems.append(f"兜底图标缺失：{n}")

    pat = re.compile(r'["\']icon["\']\s*:\s*["\']([^"\']+)["\']')
    for base in (REPO / ".pi" / "skills", REPO / "work"):
        if not base.exists():
            continue
        for p in base.rglob("*.json"):
            try:
                txt = p.read_text(encoding="utf-8")
            except (OSError, UnicodeDecodeError):
                continue
            for m in pat.finditer(txt):
                if not iconlib.resolve_icon(m.group(1)):
                    problems.append(f"deck 里引用的图标解析不到：{m.group(1)}（{p.name}）")
    return sorted(set(problems))


def current_styles() -> dict[str, list[str]]:
    return {k: [n for n in v if (ICON_ROOT / k / f"{n}.svg").is_file()]
            for k, v in scan().items()}


# ---------------------------------------------------------------- 主流程
def main() -> int:
    ap = argparse.ArgumentParser(description=".pi/assets/icons 裁剪（默认只出计划）")
    ap.add_argument("--level", choices=list(LEVELS), default="l2")
    ap.add_argument("--apply", action="store_true", help="真的执行（移到 .pi/work/_icon_trash/）")
    ap.add_argument("--verify-render", action="store_true", help="执行后渲染范例合集自检")
    args = ap.parse_args()

    styles = current_styles()
    total = sum(len(v) for v in styles.values())
    protect = protected_names()
    drop = level_rules(args.level, styles, protect)

    # 计划里只记还存在的文件
    drop = {k: v for k, v in drop.items() if (ICON_ROOT / k).with_suffix(".svg").is_file()}
    kept = {st: [n for n in names if f"{st}/{n}" not in drop] for st, names in styles.items()}

    by_reason: dict[str, int] = {}
    for r in drop.values():
        head = r.split("：", 1)[0]
        by_reason[head] = by_reason.get(head, 0) + 1

    print(f"档位 {args.level}：{LEVELS[args.level]}")
    print(f"现状 {total} 个 / {len(styles)} 个风格；删除 {len(drop)}，保留 {total - len(drop)}")
    for head, n in sorted(by_reason.items(), key=lambda kv: -kv[1]):
        print(f"  {head:6s} {n:5d}")
    print("保留分布：")
    for st, names in kept.items():
        if names:
            print(f"  {st:20s} {len(names):5d}")

    PLAN_OUT.parent.mkdir(parents=True, exist_ok=True)
    PLAN_OUT.write_text(json.dumps({
        "level": args.level, "before": total, "after": total - len(drop),
        "summary": by_reason, "kept": kept,
        "dropped": [{"spec": k, "reason": v} for k, v in sorted(drop.items())],
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"计划已写：{PLAN_OUT.relative_to(REPO)}")

    if not args.apply:
        print("\n未加 --apply，什么都没动。")
        return 0

    # 执行：移入隔离区（不 rm，随时能捞回来）
    moved = 0
    for spec in drop:
        src = (ICON_ROOT / spec).with_suffix(".svg")
        dst = (TRASH / spec).with_suffix(".svg")
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))
        moved += 1
    for st in styles:
        d = ICON_ROOT / st
        if d.is_dir() and not any(d.iterdir()):
            d.rmdir()
    for lib in sorted(os.listdir(ICON_ROOT)):
        d = ICON_ROOT / lib
        if d.is_dir() and lib not in ("LICENSES", "__pycache__") and not any(d.iterdir()):
            d.rmdir()
    print(f"\n已移动 {moved} 个文件到 {TRASH.relative_to(REPO)}/")

    after = current_styles()
    INDEX.write_text(json.dumps(rebuild_index(after), ensure_ascii=False), encoding="utf-8")
    print(f"index.json 已重建：{sum(len(v) for v in after.values())} 条")

    problems = verify(after)
    if problems:
        print(f"\n[FAIL] 校验发现 {len(problems)} 个问题：")
        for p in problems[:20]:
            print("  -", p)
        return 1
    print("[ok] 别名 / 兜底图标 / deck 引用全部还能解析")

    if args.verify_render:
        import subprocess
        py = sys.executable
        tmp = REPO / "work" / "_trim_check.json"
        subprocess.run([py, str(REPO / "renderer" / "merge.py"),
                        str(REPO / ".pi" / "skills" / "ppt-common" / "examples"),
                        str(tmp), "--title", "裁剪后自检"], check=True)
        r = subprocess.run([py, str(REPO / "renderer" / "render.py"), str(tmp),
                            "-o", str(REPO / "work" / "_trim_check.pptx"), "--qc"],
                           capture_output=True, text=True, encoding="utf-8")
        tail = (r.stdout or "").strip().splitlines()[-1:] or ["(无输出)"]
        print(f"渲染自检：{tail[0][:160]}")
        return r.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
