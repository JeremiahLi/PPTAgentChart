# -*- coding: utf-8 -*-
"""一条命令跑完：合并 → 质检 → 渲染 → 字体自检。

    python .pi/renderer/build.py .pi/work/slides -o .pi/out/主题.pptx \
        --brief seminar --style blue-business --source .pi/work/analysis.md

为什么要有它：以前一轮迭代要敲 3 条命令（merge、render --qc、font_audit），
每条都是一次"跑完 → 读输出 → 决定下一步"的模型往返；而且合并写出的 deck.json
只为给 render 读一次。现在一次启动、一份报告、一个退出码。

三个用户维度（都不传就是默认值，见 ppt-entry 技能）：

  --brief   PPT 类型：seminar 研讨型·全量详实（默认）/ briefing 汇报型·结论先行
  --style   PPT 风格：style-harness 经典藏青（默认）/ blue-business 深蓝商务 /
            bright 明亮宣讲 / forest 墨绿战略 / violet-tech 紫灰科技
            （每套风格已绑定一种卡片长相，一般不用再管 --card）
  --speed   PPT 速度：slow 慢（默认：报告全量、要求 warning 归零）/ fast 快
            （同类 warning 折成一条、直接出成品）

  **两档速度跑的是同一套检查** —— 快档省的是回炉轮次与上下文长度，不是检查项。
  流程侧的差别（大纲给不给用户过目、要不要骨架预览）写在 ppt-entry 技能里。

退出码：**恒 0**。校验拦截（slides 不合法 / 质检 error / 字体不过 / 慢档 warning 未归零）
是**正常业务逻辑**——检查在正常工作，不是运行失败，不该以非零退出码暴露
（否则每次回炉都被计成一次 bash 运行失败）。成败看末行 `RESULT {json}`：
`{"ok": true|false, "reason": ...}`。只有未捕获的运行时异常才非零退出。

常用开关：
  --check          只质检不渲染（改完页先看红线，别急着出文件）
  --only 3,7       只渲染第 3、7 页（回炉后只想看改动那两页时用；页码与整份不一致，
                   仅供看版式，交稿必须不带 --only 整份渲染）
  --notes          额外生成备注建议
  --keep-deck      顺手把合并结果写成 .pi/work/deck.json（需要单页重渲/对比时才要）
  --card           单独换卡片长相，覆盖 --style 的预设绑定
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

import briefs
import merge
import qc
import render
import theme


def _parse_only(raw: str | None, total: int) -> list[int] | None:
    if not raw:
        return None
    try:
        idx = sorted({int(x) for x in raw.replace("，", ",").split(",") if x.strip()})
    except ValueError:
        raise SystemExit(f"--only 只接受页码，例如 --only 3,7（收到 {raw!r}）")
    bad = [i for i in idx if not 1 <= i <= total]
    if bad:
        raise SystemExit(f"--only 里的页码超出范围（这份 deck 共 {total} 页）：{bad}")
    return idx


# 慢档的「warning 未归零」提示（RESULT JSON 里 ok=false, reason=warnings_not_zero）。
# 文件照常落盘，但交稿前必须处理到 0。
_ZERO_WARN_HINT = (
    "未归零: 慢档要求 warning 归零，当前 {n} 条待处理 —— "
    "逐条修掉（确认是误报的就在交付说明里写明理由），重跑到 0 再交稿；"
    "赶时间可加 --speed fast 直接出成品（检查项一样，只是不再要求归零）。"
)


def _print_result(ok: bool, **extra) -> None:
    """末行 RESULT JSON：Agent 按它判断成败。校验拦截退出码恒 0（见 docstring）。
    ensure_ascii：GBK 控制台/管道会把 UTF-8 中文误解码，转义后任何环境可无损解析。"""
    print("RESULT " + json.dumps({"ok": ok, "stage": "build", **extra},
                                 ensure_ascii=True))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="合并 → 质检 → 渲染 → 字体自检，一条命令")
    ap.add_argument("slides_dir", help="slides 目录（一个文件可以装 1 页，也可以装一页数组）")
    ap.add_argument("-o", "--out", default=".pi/out/deck.pptx", help="输出 pptx")
    ap.add_argument("--title", default=None, help="deck 标题（缺省取封面页 title）")
    ap.add_argument("--subtitle", default=None)
    ap.add_argument("--brief", choices=briefs.available(), default=briefs.DEFAULT_BRIEF,
                    help="PPT 类型（默认 %(default)s）："
                         + " / ".join(f"{n}={briefs.load(n)['label']}"
                                      for n in briefs.available()))
    ap.add_argument("--style", default=theme.DEFAULT_STYLE,
                    help="PPT 风格（配色 + 绑定卡片长相，默认 %(default)s）："
                         + " / ".join(f"{k}={v}" for k, v in theme.STYLE_LABELS.items()))
    ap.add_argument("--card", default=None,
                    help="卡片长相，**覆盖** --style 的预设绑定："
                         + " / ".join(f"{k}={v}" for k, v in theme.CARD_LABELS.items()))
    ap.add_argument("--speed", choices=qc.SPEEDS, default=qc.DEFAULT_SPEED,
                    help="PPT 速度（默认 %(default)s）：slow=慢（报告全量、要求 warning "
                         "归零、大纲给用户过目）/ fast=快（同类 warning 折一条、直接出成品）。"
                         "两档检查项完全一致，快档省的是回炉轮次")
    ap.add_argument("--source", default=None,
                    help="输入分析（.pi/work/analysis.md）：核对论据/模块有没有被丢掉")
    ap.add_argument("--notes", action="store_true", help="自动生成备注建议")
    ap.add_argument("--no-notes", action="store_true",
                    help="关掉慢档自动带上的备注建议（快档本来就不加）")
    ap.add_argument("--chart-mode", choices=["mpl", "native"], default="native")
    ap.add_argument("--check", action="store_true", help="只质检，不渲染")
    ap.add_argument("--only", default=None, help="只渲染指定页，如 3,7（仅供看版式）")
    ap.add_argument("--keep-deck", action="store_true",
                    help="把合并结果写成 .pi/work/deck.json")
    ap.add_argument("--no-font-check", action="store_true", help="跳过字体自检")
    ap.add_argument("-q", "--quiet", action="store_true", help="只输出一行结论")
    args = ap.parse_args(argv)

    t0 = time.perf_counter()
    render.layouts.CHART_MODE = args.chart_mode
    # 卡片长相：--card 显式指定优先，否则取 --style 的预设绑定（用户只选"风格"一项）
    card = args.card or theme.resolve_card(args.style)

    # ① 合并（校验失败带「文件 + 第几项」定位；拦截属正常业务，退出码 0）
    try:
        deck = merge.build_deck(args.slides_dir, title=args.title, subtitle=args.subtitle,
                                style=args.style, card=card, brief=args.brief,
                                with_notes=True)
    except merge.SlideError as e:
        lines = str(e).splitlines()
        print("\n".join(lines), file=sys.stderr)
        _print_result(False, reason="slides_invalid",
                      errors=lines[1:] if len(lines) > 1 else lines)
        return 0
    except FileNotFoundError as e:
        msg = f"找不到 slides 目录：{e}"
        print(msg, file=sys.stderr)
        _print_result(False, reason="slides_dir_missing", errors=[msg])
        return 0

    # 该速度档的流程约定（真值在方向 yaml 的 speed 段，--speed 只是选档）：
    # 慢档"不打折扣"——备注自动开、warning 归零；快档直接出成品。
    spec = briefs.speed_of(deck.brief, args.speed)
    zero_required = bool(spec.get("requireZeroWarnings"))

    if args.keep_deck:
        merge.write_deck(deck, ".pi/work/deck.json")

    # 备注建议：慢档自动开（`--no-notes` 关掉），快档不加（要就手动 `--notes`）
    if args.notes or (spec.get("notes") and args.source and not args.no_notes):
        render._auto_generate_notes(deck, args.source)

    direction = f"，{briefs.load(deck.brief)['label']}" if deck.brief else ""
    appearance = f"，{theme.appearance_label(args.style, card)}"
    print(f"合并: {len(deck.slides)} 页{direction}{appearance}（{args.speed}）")

    # ② 质检（error 一律拦住，不落盘）
    report = qc.check(deck, source_path=args.source, speed=args.speed)
    n_err, n_warn = len(report["errors"]), len(report["warnings"])
    print(f"质检: error={n_err} warning={n_warn}")
    if n_err or n_warn:
        print(render._dump_report(report))
    if n_err:
        print(f"质检未通过（{n_err} 个 error），已阻止渲染。修正后重跑即可。",
              file=sys.stderr)
        _print_result(False, reason="qc_error", errors=report["errors"], out=None)
        return 0
    if args.check:
        print(f"仅质检（未渲染）。耗时 {time.perf_counter() - t0:.1f}s")
        if zero_required and n_warn:
            print(_ZERO_WARN_HINT.format(n=n_warn), file=sys.stderr)
            _print_result(False, reason="warnings_not_zero", warnings=n_warn, out=None)
            return 0
        _print_result(True, out=None, checked_only=True, warnings=n_warn)
        return 0

    # ③ 渲染（--only 走"只看这几页"的临时视图，不写回 deck）
    only = _parse_only(args.only, len(deck.slides))
    target = deck
    if only:
        target = deck.model_copy(update={
            "slides": [deck.slides[i - 1] for i in only],
            "speaker_notes": None,
        })
        print(f"提示: 只渲染第 {','.join(str(i) for i in only)} 页，"
              f"页码与整份不一致，仅用于看版式；交稿请整份渲染。")
    Path(os.path.dirname(args.out) or ".").mkdir(parents=True, exist_ok=True)
    render.build_pptx(target, args.out, add_speaker_notes=bool(target.speaker_notes))
    size_kb = os.path.getsize(args.out) / 1024

    # ④ 字体自检（同一进程内，省一次启动与一条命令）
    font_bad = 0
    if not args.no_font_check:
        import fontcheck
        font_bad = len(fontcheck.audit(Path(args.out), verbose=False))

    elapsed = time.perf_counter() - t0
    if font_bad:
        print(f"字体: FAIL，{font_bad} 处不符合「微软雅黑 + Times New Roman」，"
              f"跑 python .pi/renderer/fontcheck.py {args.out} 看细节", file=sys.stderr)
        _print_result(False, reason="font_check", font_issues=font_bad, out=args.out)
        return 0
    if args.quiet:
        print(f"OK: {args.out}（{len(target.slides)} 页{appearance}）")
    else:
        print(f"字体: PASS")
        print(f"OK: {args.out}（{len(target.slides)} 页{direction}{appearance}，"
              f"{size_kb:.0f} KB，{elapsed:.1f}s）")
    # 慢档要求 warning 归零：文件照常落盘（不然没法看效果），但 RESULT 里标
    # ok=false——"质检过了 80% 就交稿"是最常见的偷懒方式，必须让它显形。
    # --only 只看版式，不算交稿，不参与归零要求。
    if zero_required and n_warn and not only:
        print(_ZERO_WARN_HINT.format(n=n_warn), file=sys.stderr)
        _print_result(False, reason="warnings_not_zero", warnings=n_warn, out=args.out)
        return 0
    _print_result(True, out=args.out, slides=len(target.slides),
                  warnings=n_warn, seconds=round(elapsed, 1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
