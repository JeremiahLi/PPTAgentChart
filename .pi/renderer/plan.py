# -*- coding: utf-8 -*-
"""规划引擎 CLI：分析文档 → 生成结构化大纲。

用法：
  # 从 analysis.md 分析
  python .pi/renderer/plan.py .pi/work/analysis.md -o .pi/work/outline.json --title "Pi 编程智能体" --brief seminar

  # 从原始文档分析（无 analysis.md 时）
  python .pi/renderer/plan.py input/article.md -o .pi/work/outline.json --from_raw

  # 检查大纲是否需要重跑（对比缓存键）
  python .pi/renderer/plan.py .pi/work/analysis.md --check-cache

  # 从原文自动生成 analysis.md 骨架（结论/论据/缺口留给人和模型补）
  python .pi/renderer/plan.py .pi/work/input.md --emit-analysis
"""

from __future__ import annotations

import argparse
import json
import os
import sys

# 确保能导入兄弟模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import briefs
import doc_analyzer
import planner
import outline

ANALYSIS_TEMPLATE = ".pi/work/analysis.md"


def main():
    parser = argparse.ArgumentParser(
        description="规划引擎：文档分析 → 结构化大纲")
    parser.add_argument("input", help="输入文件（analysis.md 或原始文档）")
    parser.add_argument("-o", "--out", default=".pi/work/outline.json",
                        help="大纲输出路径（默认 .pi/work/outline.json）")
    parser.add_argument("--title", default="演示文稿", help="deck 标题")
    parser.add_argument("--subtitle", default="", help="deck 副标题")
    parser.add_argument("--brief", default="seminar",
                        choices=briefs.available() or ["seminar", "briefing"],
                        help="写作方向")
    parser.add_argument("--from-raw", action="store_true",
                        help="输入是原始文档（按标题自动切分）")
    parser.add_argument("--check-cache", action="store_true",
                        help="检查缓存：对比现有 outline 的 source_hash，未变则跳过")
    parser.add_argument("--force", action="store_true",
                        help="强制重新生成（忽略缓存）")
    parser.add_argument("--emit-analysis", nargs="?", const=ANALYSIS_TEMPLATE,
                        default=None, metavar="OUT_MD",
                        help=f"只生成分析表骨架（默认 {ANALYSIS_TEMPLATE}）然后退出："
                             f"章节/模块/论据/数字按原文自动填好，只留结论与缺口给人补。"
                             f"输入应当是原文或 docx_in 产出的 input.md。")
    args = parser.parse_args()

    if args.emit_analysis:
        if args.input.endswith("analysis.md"):
            print("这已经是分析表了：--emit-analysis 的输入应当是原文/input.md。",
                  file=sys.stderr)
            sys.exit(2)
        modules = doc_analyzer.analyze_raw_text(args.input)
        if not modules:
            print(f"没能从 {args.input} 切出模块：确认原文有 # / ## 标题、"
                  f"且每节正文不少于 30 字。", file=sys.stderr)
            sys.exit(2)
        doc_analyzer.write_analysis_skeleton(args.emit_analysis, modules,
                                             source_name=os.path.basename(args.input))
        print(f"骨架已写入: {args.emit_analysis}（{len(modules)} 个模块）")
        print(f"  已自动填: 章节 / 模块 / 论据 / 数字")
        print(f"  待你补:   核心结论（判断句）、缺口（没有就写 无）")
        print(f"下一步: 改完 {args.emit_analysis} 后重跑（不带 --emit-analysis）")
        print(f"  python .pi/renderer/plan.py {args.emit_analysis} -o {args.out} "
              f"--title \"...\" --brief {args.brief}")
        return

    # 缓存检查
    if args.check_cache and not args.force and os.path.isfile(args.out):
        existing = outline.load_outline(args.out)
        new_hash = doc_analyzer.compute_source_hash(args.input)
        if existing.source_hash == new_hash:
            print(f"SKIP: 缓存命中（hash={new_hash}），大纲未变")
            stats = existing.stats()
            print(f"  {stats['total']} 页（内容 {stats['content']}）")
            return

    # 步骤 1：内容分析
    print(f"分析: {args.input}")
    if args.from_raw:
        modules = doc_analyzer.analyze_raw_text(args.input)
    else:
        modules = doc_analyzer.analyze_analysis_md(args.input)
    print(f"  识别 {len(modules)} 个内容模块")

    # 步骤 2：规划
    print("规划大纲...")
    outline_result = planner.plan(
        modules,
        title=args.title,
        subtitle=args.subtitle,
        brief=args.brief,
    )
    outline_result.source_hash = doc_analyzer.compute_source_hash(args.input)

    # 步骤 3：保存
    outline.save_outline(outline_result, args.out)

    # 输出摘要
    stats = outline_result.stats()
    rel_pages = sum(1 for sec in outline_result.sections
                    for s in sec.slides if s.relation_note)
    print(f"\n大纲已保存: {args.out}")
    print(f"  总页数: {stats['total']}（内容页 {stats['content']}）")
    print(f"  版式分布: {json.dumps(stats['by_type'], ensure_ascii=False)}")
    print(f"  关系合并: {rel_pages} 页（跨模块关系路由）")
    print(f"  图表类型: {stats['chart_types'] or '无'}")
    print(f"  平均密度: {outline_result.avg_density:.0f} 字/页")
    print(f"  缓存键: {outline_result.source_hash}")

    # 提示下一步
    print(f"\n下一步:")
    print(f"  1. 审阅大纲: cat {args.out}（这是唯一要停下来让用户过目的环节）")
    print(f"  2. 写逐页 JSON 到 .pi/work/slides/（一个文件 5-8 页，数组），一条命令出图:")
    print(f"     python .pi/renderer/build.py .pi/work/slides -o .pi/out/result.pptx "
          f"--brief {args.brief} --source .pi/work/analysis.md")
    print(f"  3. 只想先看骨架（不写页）: python .pi/renderer/merge.py --from-outline {args.out} "
          f"-o .pi/work/deck.json && python .pi/renderer/render.py .pi/work/deck.json "
          f"-o .pi/out/skeleton.pptx --qc")
    print(f"  4. 检查缓存: python .pi/renderer/plan.py {args.input} -o {args.out} --check-cache")


if __name__ == "__main__":
    main()
