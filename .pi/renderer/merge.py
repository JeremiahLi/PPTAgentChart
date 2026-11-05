# -*- coding: utf-8 -*-
"""合并 slides 目录为 deck.json。

两种写法都支持（可以混用，按文件名排序后依次展开）：

1. **一个文件一页**：`01-cover.json` 里就是一个 slide 对象（旧写法）；
2. **一个文件多页（推荐）**：`01-batch.json` 里是一个 slide 数组，一批 5-8 页。

   为什么推荐数组：每次 `write_to_file` 都是一次模型往返，40 页按一页一文件写
   就是 40 次往返；一个文件装 6 页只需 7 次。渲染本身只要 1 秒，往返才是成本。

用法：
    python merge.py <slides_dir> <out.json> [--title 标题] [--brief 方向]
    python merge.py --from-outline .pi/work/outline.json -o .pi/work/deck.json

批量文件的好处还有一处：校验失败会指到「文件 + 第几项」（旧写法只知道 slides[51]）。
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys

import briefs
import theme
from ppt_dsl import SlideNotes, validate_dict


class SlideError(Exception):
    """slides 目录里的内容不合法：带「文件 + 第几项」的定位信息。"""


def _force_utf8_streams() -> None:
    """Windows 控制台默认 GBK：校验报错里的中文与字段值会打成乱码，
    模型读错误信息时看不清「哪个字段、当前多少字」，修复效率暴跌，
    还会连带产生 UnicodeDecodeError。统一把 stdout/stderr 切到 UTF-8。"""
    for stream in (sys.stdout, sys.stderr):
        enc = getattr(stream, "encoding", "") or ""
        if enc.lower().replace("-", "") != "utf8" and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except Exception:
                pass


def _coerce_common_mistakes(item: dict, where: str) -> list[str]:
    """两种高频结构混淆的自动纠偏（能救则救，救不了的留给 pydantic 报准确错）：

    1. `panel_2col` 的 `panels[].points` 是**字符串数组**，模型常写成
       `[{'text': ...}, ...]`（那是 statement 的写法）→ 提取 text；
    2. `statement` 的 `points` 是 `{label, text}` **对象数组**，模型常写成
       纯字符串数组（那是 panel 的写法）→ 前缀截成 label 包一层。
    两版式的 points 结构正好相反，属于复发率最高的混淆，见 pitfalls。
    """
    warns: list[str] = []
    st = item.get("slide_type")

    if st == "panel_2col":
        for pi, panel in enumerate(item.get("panels") or []):
            pts = panel.get("points") if isinstance(panel, dict) else None
            if not isinstance(pts, list):
                continue
            fixed, changed = [], False
            for p in pts:
                if (isinstance(p, dict) and "label" not in p
                        and isinstance(p.get("text"), str)):
                    fixed.append(p["text"])
                    changed = True
                else:
                    fixed.append(p)
            if changed:
                panel["points"] = fixed
                warns.append(f"{where}: panel.points 里的 {{'text':…}} 已提取为字符串（panel 版式 points 是字符串数组）")

    if st == "statement":
        pts = item.get("points")
        if isinstance(pts, list) and any(isinstance(p, str) for p in pts):
            item["points"] = [
                ({"label": p[:8], "text": p} if isinstance(p, str) and len(p) >= 6 else p)
                for p in pts
            ]
            warns.append(f"{where}: statement.points 里的纯字符串已包成 {{label,text}}（label 取前 8 字，statement 版式 points 是对象数组）")
    return warns


def collect_slides(slides_dir: str) -> tuple[list[dict], list[dict], list[str], list[str]]:
    """读取目录里所有 *.json，扁平化成 slides 列表。

    返回 `(slides, notes, origin, warns)`：
      - slides : 每个元素是一个 slide dict；
      - notes  : 与 slides 等长、逐页占位（render 用 `speaker_notes[i-1]` 按索引取，
                 只 append 有备注的页会让后面整体前移，贴到别的页上）；
      - origin : 与 slides 等长，形如 `["02-batch.json 第 3 项", ...]`，报错定位用。
      - warns  : 自动纠偏提示（type→slide_type、points 结构混淆等），不阻止合并。
    """
    files = sorted(glob.glob(os.path.join(slides_dir, "*.json")))
    if not files:
        raise SlideError(f"在 {slides_dir} 下没有找到 *.json")

    slides: list[dict] = []
    notes: list[dict] = []
    origin: list[str] = []
    warns: list[str] = []
    for path in files:
        name = os.path.basename(path)
        try:
            # utf-8-sig：兼容带 BOM 的 UTF-8（Windows 编辑器/PowerShell 写出的文件
            # 常带 BOM，按纯 utf-8 读会直接 UnicodeDecodeError/BOM 报错）。
            with open(path, "r", encoding="utf-8-sig") as fh:
                data = json.load(fh)
        except json.JSONDecodeError as e:
            raise SlideError(f"{name} 不是合法 JSON：{e.msg}（第 {e.lineno} 行）") from None
        items = data if isinstance(data, list) else [data]
        if not items:
            raise SlideError(f"{name} 是空数组，至少要有 1 页")
        for idx, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise SlideError(
                    f"{name} 第 {idx} 项不是对象（是 {type(item).__name__}）："
                    f"一个文件多页时，每项必须是一页 slide 对象")
            # 版式判别字段是 `slide_type`，但模型常照直觉写成 `type`
            #（union_tag_not_found 的头号来源）。这里做一次无损归一化。
            if "slide_type" not in item and isinstance(item.get("type"), str):
                item["slide_type"] = item.pop("type")
                warns.append(f"{name} 第 {idx} 项：字段 `type` 已归一化为 `slide_type`")
            if "slide_type" not in item:
                hint = "（这看着像一份 deck.json —— 合并输出别写在 slides 目录里）" \
                    if "slides" in item else ""
                raise SlideError(
                    f"{name} 第 {idx} 项缺 `slide_type`：每页都要有 slide_type，"
                    f"别把备注/清单/索引也塞进 slides 目录{hint}")
            slides.append(item)
            warns.extend(_coerce_common_mistakes(item, f"{name} 第 {idx} 项"))
            n = item.get("speaker_notes") or {}
            notes.append({
                "data_gaps": n.get("data_gaps", []),
                "viz_suggestions": n.get("viz_suggestions", []),
                "source_suggestions": n.get("source_suggestions", []),
                "content_gaps": n.get("content_gaps", []),
            })
            origin.append(name if len(items) == 1 else f"{name} 第 {idx} 项")
    return slides, notes, origin, warns


def _format_errors(errors, origin: list[str]) -> str:
    """把 pydantic 的 `slides[51]` 翻译成「02-batch.json 第 3 项 / 字段」。"""
    lines = []
    for e in errors:
        loc = e.get("loc") or ()
        where = "deck"
        field = ""
        if len(loc) >= 2 and isinstance(loc[1], int):
            if 0 <= loc[1] < len(origin):
                where = origin[loc[1]]
            field = " / ".join(str(x) for x in loc[2:])
        msg = e.get("msg", "")
        if "discriminator 'slide_type'" in msg:
            msg = "没有可识别的 slide_type（写错版式名或漏字段）"
        elif "union_tag" in e.get("type", ""):
            msg = "没有可识别的 slide_type（写错版式名或漏字段）"
        # 长度类错误补上「当前多少字 vs 上限多少」——只给上限不给现值时，
        # 模型不知道要砍多少，容易逐字试错、反复重跑。
        etype = e.get("type", "")
        if etype in ("string_too_long", "string_too_short"):
            ctx = e.get("ctx") or {}
            val = e.get("input")
            cur = len(val) if isinstance(val, str) else None
            limit = ctx.get("max_length") if etype == "string_too_long" else ctx.get("min_length")
            if cur is not None and limit is not None:
                msg += f"（当前 {cur} 字，{'上限' if etype == 'string_too_long' else '下限'} {limit} 字）"
        elif etype in ("string_type", "model_type"):
            msg += "（注意：panel_2col 的 points 是字符串数组，statement 的 points 是 {label,text} 对象数组，别混用）"
        lines.append(f"  · {where}" + (f" [{field}]" if field else "") + f"：{msg}")
    return "deck 校验失败（按 slide 顺序）：\n" + "\n".join(lines)


def build_deck(slides_dir: str, *, title: str | None = None, subtitle: str | None = None,
               style: str = "style-harness", card: str = "flat",
               brief: str | None = None, with_notes: bool = False):
    """合并 + 校验，返回 Deck 模型（不落盘）。build.py / merge CLI 共用这一份逻辑。"""
    from pydantic import ValidationError

    slides, notes, origin, warns = collect_slides(slides_dir)
    for w in warns:
        print(f"⚠ 自动纠偏：{w}", file=sys.stderr)
    if title is None:
        title = next((s.get("title") for s in slides
                      if s.get("slide_type") == "cover"), "演示文稿")
    deck_dict = {"title": title, "subtitle": subtitle, "style": style,
                 "card": card, "brief": brief, "slides": slides}
    if with_notes:
        deck_dict["speaker_notes"] = [SlideNotes(**n).model_dump() for n in notes]
    try:
        return validate_dict(deck_dict)
    except ValidationError as e:
        raise SlideError(_format_errors(e.errors(), origin)) from None


def write_deck(deck, out: str) -> None:
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(deck.model_dump(), fh, ensure_ascii=False, indent=2)


def _print_result(ok: bool, errors: list[str] | None = None, **extra) -> None:
    """末行打印 RESULT JSON。校验拦截（schema 不过 / 目录里没有合法页）是
    **正常业务逻辑**——schema 在正常工作，不是运行失败：退出码恒 0，
    成败只看这份 JSON，避免把「回炉」错记成一次 bash 运行失败。"""
    payload = {"ok": ok, "stage": "merge", **extra}
    if errors:
        payload["errors"] = errors
    # ensure_ascii：GBK 控制台/管道会把 UTF-8 中文误解码，转义成 \uXXXX 任何
    # 编码环境下 json.loads 都能无损还原。
    print("RESULT " + json.dumps(payload, ensure_ascii=True))


def main():
    _force_utf8_streams()
    ap = argparse.ArgumentParser()
    ap.add_argument("slides_dir", nargs="?",
                    help="slides 目录（每个 *.json 可以是一页，也可以是一页数组）")
    ap.add_argument("out", nargs="?", default=".pi/work/deck.json",
                    help="输出 deck.json 路径")
    ap.add_argument("-o", "--out-path", dest="out_path", default=None,
                    help="输出路径（与位置参数等价；plan.py 的提示用的是 -o）")
    ap.add_argument("--title", default=None, help="deck 标题（缺省取封面页 title）")
    ap.add_argument("--subtitle", default=None)
    ap.add_argument("--style", default=theme.DEFAULT_STYLE,
                    help="PPT 风格（配色 + 绑定卡片长相，默认经典藏青）："
                         + " / ".join(f"{k}={v}" for k, v in theme.STYLE_LABELS.items()))
    ap.add_argument("--card", default=None,
                    help="卡片长相，**覆盖** --style 的预设绑定："
                         + " / ".join(f"{k}={v}" for k, v in theme.CARD_LABELS.items()))
    ap.add_argument("--brief", choices=briefs.available(), default=None,
                    help="写作方向（研讨型/汇报型）：渲染外的取舍规约，见 .pi/renderer/briefs/")
    ap.add_argument("--from-outline", default=None,
                    help="从 outline.json 生成最简 deck（结构化大纲优先）")
    ap.add_argument("--with-notes", action="store_true",
                    help="保留备注到 deck.json")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="只输出一行 OK（脚本/CI 用）")
    args = ap.parse_args()
    # `-o` 与位置参数等价：plan.py 结尾打印的下一步命令用的就是 `-o`，
    # 而 argparse 之前只认位置参数 → 照着提示敲必定 unrecognized arguments。
    if getattr(args, "out_path", None):
        args.out = args.out_path

    # 从 outline 生成
    if args.from_outline:
        _merge_from_outline(args)
        return

    if not args.slides_dir:
        ap.error("需要指定 slides_dir 或 --from-outline")
    if os.path.abspath(os.path.dirname(args.out) or ".") == os.path.abspath(args.slides_dir):
        msg = (f"deck.json 不要写进 slides 目录（{args.slides_dir}）："
               f"下次合并会把它当成一页读进去。换个输出路径，例如 .pi/work/deck.json")
        print(msg, file=sys.stderr)
        _print_result(False, errors=[msg])
        return

    try:
        deck = build_deck(args.slides_dir, title=args.title, subtitle=args.subtitle,
                          style=args.style,
                          card=args.card or theme.resolve_card(args.style),
                          brief=args.brief, with_notes=args.with_notes)
    except SlideError as e:
        # 校验拦截 ≠ 运行失败：退出码 0，错误清单进 RESULT JSON
        lines = str(e).splitlines()
        print("\n".join(lines), file=sys.stderr)
        _print_result(False, errors=lines[1:] if len(lines) > 1 else lines)
        return
    n_files = len(glob.glob(os.path.join(args.slides_dir, "*.json")))
    write_deck(deck, args.out)

    tag = f"，{briefs.load(args.brief)['label']}" if args.brief else ""
    if args.quiet:
        print(f"OK: {args.out}（{len(deck.slides)} 页{tag}）")
        _print_result(True, out=args.out, slides=len(deck.slides))
        return
    print(f"OK: {args.out}（{n_files} 个文件 → {len(deck.slides)} 页{tag}）")
    _print_result(True, out=args.out, slides=len(deck.slides), files=n_files)


def _merge_from_outline(args):
    """从 outline.json 生成最简 deck。"""
    import outline

    ol = outline.load_outline(args.from_outline)
    deck_dict = outline.outline_to_deck_cli(ol)

    if args.title:
        deck_dict["title"] = args.title
    if args.subtitle:
        deck_dict["subtitle"] = args.subtitle
    if args.brief:
        deck_dict["brief"] = args.brief
    # 卡片长相：显式 --card > outline 里已带的 > 风格的预设绑定（与 build.py 同口径）
    deck_dict["card"] = (getattr(args, "card", None) or deck_dict.get("card")
                         or theme.resolve_card(deck_dict.get("style")))

    try:
        validate_dict(deck_dict)
    except Exception as e:  # pydantic.ValidationError 等：校验拦截，不是运行失败
        from pydantic import ValidationError
        if isinstance(e, ValidationError):
            lines = _format_errors(e.errors(), ["outline"] * len(e.errors())).splitlines()
        else:
            lines = [str(e)]
        print("\n".join(lines), file=sys.stderr)
        _print_result(False, errors=lines[1:] if len(lines) > 1 else lines)
        return

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(deck_dict, fh, ensure_ascii=False, indent=2)

    stats = ol.stats()
    print(f"OK: {args.out}（从 outline 生成 {stats['total']} 页）")
    _print_result(True, out=args.out, slides=stats['total'])
    if getattr(args, "quiet", False):
        return
    print(f"  内容页: {stats['content']}")
    print(f"  版式: {stats['by_type']}")
    print(f"  图表: {stats['chart_types'] or '无'}")


if __name__ == "__main__":
    main()
