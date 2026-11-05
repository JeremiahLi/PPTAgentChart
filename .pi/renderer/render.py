# -*- coding: utf-8 -*-
"""渲染入口：python render.py deck.json -o out.pptx [--qc]

--qc：渲染前先跑质检，有 error 级问题则不渲染并以非零码退出，
     错误报告（JSON）打印到 stdout，供 Agent 读取后回炉重生成。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import briefs
import layouts
import speaker_notes
import doc_analyzer
import theme
from theme import COLORS, apply_theme_fonts
from ppt_dsl import load_deck, SlideNotes
import qc


def build_pptx(deck, out_path: str, add_speaker_notes: bool = True):
    # 外观两个维度，都在画第一页之前切（未登记的名字确定性回退默认值）：
    #   style → 配色主题（5 套，见 theme._THEMES）
    #   card  → 卡片长相（5 种，见 theme.CARD_STYLES）
    # 都**无条件**调用：一个进程里会连续渲染多份 deck（smoke.py 就是），
    # 只在"有值才调"会让上一份的主题漏到下一份。
    theme.apply_style(getattr(deck, "style", None))
    theme.apply_card_style(getattr(deck, "card", None))
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    # 母版主题的默认字体也改成「中文微软雅黑 + 西文 Times New Roman」：
    # 幻灯片里每段文字都显式带了字体，这一步是让"默认值"也对齐
    # （用户在 PowerPoint 里新加文本/编辑图表默认文本时才不会退回 Calibri）。
    apply_theme_fonts(prs)
    blank = prs.slide_layouts[6]
    for i, slide_model in enumerate(deck.slides, 1):
        slide = prs.slides.add_slide(blank)
        layouts.RENDERERS[slide_model.slide_type](slide, slide_model)

        # 写入备注（每页 PPT 下方的备注区）
        if add_speaker_notes and deck.speaker_notes and i <= len(deck.speaker_notes):
            notes = deck.speaker_notes[i - 1]
            has_content = (notes.data_gaps or notes.viz_suggestions or
                           notes.source_suggestions or notes.content_gaps)
            if has_content:
                text = notes.render()
                if text:
                    notes_slide = slide.notes_slide
                    tf = notes_slide.notes_text_frame
                    tf.text = text
                    # 备注页的 run 也走同一字体口径：不补就会落成模板默认（Calibri），
                    # 中文在备注区显示成别的字体。
                    for _p in tf.paragraphs:
                        for _run in _p.runs:
                            layouts.apply_run_font(_run)

        # 页码：左下角。底图 bg-content.png 右下角有单位名称水印
        # （x 10.95–12.90, y 6.97–7.15），页码放右下会压在水印上。
        if slide_model.slide_type not in ("cover", "end"):
            tb = slide.shapes.add_textbox(Inches(0.55), Inches(layouts.FOOTER_Y),
                                          Inches(0.55), Inches(layouts.FOOTER_H))
            tf = tb.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            run = p.add_run()
            run.text = f"{i:02d}"
            run.font.size = Pt(10.5)
            # 底色是 #F2F2F2，原来的 9CA3AF 只有 2.0:1，几乎看不见
            run.font.color.rgb = RGBColor.from_string(COLORS["text_sub"])
            # 自己新建的 run 必须补字体（否则退回模板字体 Calibri）
            layouts.apply_run_font(run)
    prs.save(out_path)


def _dump_report(report: dict) -> str:
    """紧凑打印质检报告：一行一条，省掉 `indent=2` 那些纯结构的缩进字节。

    报告整份要进模型上下文（30 页 deck 实测 16.6KB），`indent=2` 的结构开销
    占 3.6KB —— 换成紧凑 JSON 什么都不丢，只少烧 token。
    """
    chunks = [f'"{k}":{json.dumps(v, ensure_ascii=False, separators=(",", ":"))}'
              for k, v in report.items() if k not in ("errors", "warnings")]
    for key in ("errors", "warnings"):
        rows = ",\n  ".join(json.dumps(it, ensure_ascii=False,
                                       separators=(",", ":"))
                            for it in (report.get(key) or []))
        chunks.append(f'"{key}":[{chr(10) + "  " + rows + chr(10) if rows else ""}]')
    return "{" + ",\n".join(chunks) + "}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", help="deck.json 路径")
    ap.add_argument("-o", "--out", default="out.pptx")
    ap.add_argument("--qc", action="store_true", help="渲染前执行质检")
    ap.add_argument("--speed", choices=qc.SPEEDS, default=qc.DEFAULT_SPEED,
                    help="PPT 速度（默认 %(default)s）：slow=慢（要求 warning 归零）/ "
                         "fast=快（同类 warning 折一条）。两档检查项完全一致")
    ap.add_argument("--source", default=None,
                    help="输入分析（如 .pi/work/analysis.md）：核对输入里的论据/模块"
                         "有没有被丢掉（研讨型必用，见 ppt-seminar/SKILL.md）")
    ap.add_argument("--chart-mode", choices=["mpl", "native"], default="native",
                    help="图表渲染方式：native=可编辑原生图表(默认)，mpl=matplotlib 高质感图")
    ap.add_argument("--brief", choices=briefs.available(), default=None,
                    help="写作方向（研讨型/汇报型）：决定取舍规则与页数预算，不影响外观。"
                         "默认沿用 deck.json 里的 brief 字段")
    ap.add_argument("--style", default=None,
                    help="PPT 风格，覆盖 deck.json 的 style（未登记的名字回退经典藏青）："
                         + " / ".join(f"{k}={v}" for k, v in theme.STYLE_LABELS.items()))
    ap.add_argument("--card", default=None,
                    help="卡片长相，覆盖 deck.json 的 card 与风格预设（未登记回退简约白卡）："
                         + " / ".join(f"{k}={v}" for k, v in theme.CARD_LABELS.items()))
    ap.add_argument("--notes", action="store_true",
                    help="自动生成备注建议（基于内容分析）")
    args = ap.parse_args()
    layouts.CHART_MODE = args.chart_mode

    # schema 校验拦截 = 正常业务逻辑（检查在正常工作），不是运行失败：
    # 退出码 0 + 末行 RESULT JSON，与 merge.py / build.py 同口径。
    try:
        deck = load_deck(args.deck)
    except Exception as e:
        from pydantic import ValidationError
        if isinstance(e, ValidationError):
            msg = f"{args.deck} 未通过 schema 校验：\n" + "\n".join(
                f"  · {'/'.join(map(str, err['loc']))}：{err['msg']}"
                for err in e.errors())
        else:
            msg = str(e)
        print(msg, file=sys.stderr)
        print("RESULT " + json.dumps(
            {"ok": False, "stage": "render", "reason": "schema_invalid",
             "errors": msg.splitlines()}, ensure_ascii=True))
        return
    if args.brief:
        deck.brief = args.brief
    # 外观覆盖：命令行优先于 deck.json（语义与 --brief 一致）
    if args.style:
        deck.style = args.style
    # 卡片长相：命令行 > deck.json > 风格的预设绑定（与 build.py / merge.py 同口径）
    deck.card = args.card or deck.card or theme.resolve_card(deck.style)

    # 自动生成备注
    if args.notes:
        _auto_generate_notes(deck, args.source)

    if args.qc:
        report = qc.check(deck, source_path=args.source, speed=args.speed)
        print(_dump_report(report))
        if report["errors"]:
            print(f"质检未通过（{len(report['errors'])} 个 error），已阻止渲染。请修正后重试。", file=sys.stderr)
            # 质检拦截是正常业务逻辑，不是运行失败：退出码 0 + RESULT JSON
            print("RESULT " + json.dumps(
                {"ok": False, "stage": "render", "reason": "qc_error",
                 "errors": report["errors"]}, ensure_ascii=True))
            return

    build_pptx(deck, args.out, add_speaker_notes=True)
    tag = f"，{briefs.load(deck.brief)['label']}" if deck.brief else ""
    style_tag = f"，主题 {theme.CURRENT_STYLE}，卡片 {theme.CURRENT_CARD}"
    print(f"OK: {args.out}（{len(deck.slides)} 页{tag}{style_tag}）")
    print("RESULT " + json.dumps(
        {"ok": True, "stage": "render", "out": args.out,
         "slides": len(deck.slides)}, ensure_ascii=True))


def _auto_generate_notes(deck, source_path: str | None = None):
    """为 deck 自动生成备注。"""
    import speaker_notes
    import doc_analyzer

    # 尝试读取分析表
    analysis_text = None
    if source_path and os.path.isfile(source_path):
        with open(source_path, "r", encoding="utf-8") as f:
            analysis_text = f.read()

    notes_list = []
    for slide_model in deck.slides:
        # 提取当前页的数据特征
        text = getattr(slide_model, 'title', '') or ''
        evidence = []

        if hasattr(slide_model, 'body'):
            text += ' ' + (slide_model.body or '')
        if hasattr(slide_model, 'summary'):
            text += ' ' + (slide_model.summary or '')
        if hasattr(slide_model, 'columns'):
            for col in slide_model.columns:
                evidence.append(col.label + ' ' + col.desc)
                for p in col.points:
                    evidence.append(p)
        if hasattr(slide_model, 'cards'):
            for card in slide_model.cards:
                evidence.append(card.name + ' ' + card.desc)
        if hasattr(slide_model, 'points'):
            for p in slide_model.points:
                evidence.append(getattr(p, 'label', '') + ' ' + getattr(p, 'text', ''))
        if hasattr(slide_model, 'nodes'):
            for n in slide_model.nodes:
                evidence.append(n.label + ' ' + n.desc)
        if hasattr(slide_model, 'steps'):
            for s in slide_model.steps:
                evidence.append(s.label + ' ' + s.desc)
        if hasattr(slide_model, 'panels'):
            for p in slide_model.panels:
                evidence.append(p.label + ' ' + p.summary)
                evidence.extend(p.points)

        data_points = doc_analyzer.extract_data_points(text + ' '.join(evidence))

        sn = speaker_notes.generate_notes(
            slide_type=slide_model.slide_type,
            module_name=getattr(slide_model, 'title', ''),
            conclusion=getattr(slide_model, 'summary', '') or getattr(slide_model, 'title', ''),
            evidence=evidence,
            gaps=[],
            data_points=data_points,
            is_data_heavy=doc_analyzer._is_data_heavy(text),
            is_text_heavy=doc_analyzer._is_text_heavy(text),
            has_comparison=doc_analyzer.has_comparison(text),
            has_process=doc_analyzer.has_process(text),
            has_hierarchy=doc_analyzer.has_hierarchy(text),
        )

        from ppt_dsl import SlideNotes
        notes_list.append(SlideNotes(
            data_gaps=[n.text for n in sn.notes if n.category == "data"],
            viz_suggestions=[n.text for n in sn.notes if n.category == "viz"],
            source_suggestions=[n.text for n in sn.notes if n.category == "source"],
            content_gaps=[n.text for n in sn.notes if n.category == "gap"],
        ))

    deck.speaker_notes = notes_list


if __name__ == "__main__":
    main()
