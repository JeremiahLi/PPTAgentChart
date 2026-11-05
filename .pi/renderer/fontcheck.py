# -*- coding: utf-8 -*-
"""字体自检：核对交付用的 .pptx 里每段文字都是「中文微软雅黑 + 西文 Times New Roman」。

这里是**唯一实现**（`build.py` 渲染完直接调它，省一次进程启动与一条命令往返）；
命令行入口仍保留在两个地方，行为完全一致：
    python .pi/renderer/fontcheck.py .pi/out/我的汇报.pptx
    python .pi/tests/font_audit.py .pi/out/我的汇报.pptx     # 薄壳，历史命令继续可用

查三件事，任何一件不过就 exit 1：
  1. 幻灯片 / 原生图表 / 备注页里每个 `a:rPr`、`a:defRPr` 的 latin + ea **成对**写全；
     （只写一个，另一半会退回模板字体 Calibri —— 这就是"字体不对"最常见的原因）
  2. 母版主题的 majorFont / minorFont / `script="Hans"` 三项也都是这套字体；
     （改主题是为了用户新加文本框、编辑图表默认文本时拿到的仍是雅黑，不是宋体）
  3. 字体名在 XML 里是 **UTF-8 正确**的"微软雅黑"，不是被 GBK 写坏的乱码字节。

只读不写，不依赖 renderer，也不需要装 Office。
"""
from __future__ import annotations

import os
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EA_EXPECT = "微软雅黑"
# 西文口径与 theme.py 同源：设了 PPT_FONT_LATIN 就按那个核对（全文雅黑模式）
LATIN_EXPECT = os.environ.get("PPT_FONT_LATIN", "Times New Roman").strip() or "Times New Roman"


def _group(name: str) -> str:
    """按部件归类：只有真正会被看见的文字才做硬检查。"""
    if name.startswith("ppt/slides/slide"):
        return "幻灯片"
    if name.startswith("ppt/charts/"):
        return "原生图表"
    if name.startswith("ppt/notesSlides/"):
        return "备注页"
    if name.startswith(("ppt/slideLayouts/", "ppt/slideMasters/",
                        "ppt/notesMasters/")):
        return "母版/版式"
    return "其它"


def _face(el, tag: str):
    node = el.find(f"{{{A}}}{tag}")
    return None if node is None else node.get("typeface")


def audit(path: Path, *, verbose: bool = True) -> list[str]:
    """返回问题列表（空 = 通过）。`verbose=False` 时不打印细节，供 build.py 汇总。"""
    path = Path(path)
    problems: list[str] = []
    per_part: dict[str, Counter] = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.endswith(".xml"):
                continue
            raw = z.read(name)
            try:
                root = ET.fromstring(raw)
            except ET.ParseError:
                problems.append(f"{name}: XML 解析失败")
                continue
            g = _group(name)
            cnt = per_part.setdefault(g, Counter())
            for tag in ("rPr", "defRPr", "endParaRPr"):
                for el in root.iter(f"{{{A}}}{tag}"):
                    lv, ev = _face(el, "latin"), _face(el, "ea")
                    cnt[(lv, ev)] += 1
                    # endParaRPr 只是段落末尾标记，没有可见文字，不判错（否则每个图
                    # 表部件都会误报一条）
                    if tag == "endParaRPr":
                        continue
                    if g in ("幻灯片", "原生图表", "备注页") and ev != EA_EXPECT:
                        problems.append(f"{name}: <{tag}> ea={ev!r}（应为 {EA_EXPECT!r}）")
                    elif g in ("幻灯片",) and lv != LATIN_EXPECT:
                        problems.append(f"{name}: <{tag}> latin={lv!r}（应为 {LATIN_EXPECT!r}）")
            # 字体名不能是 GBK 写坏的乱码：UTF-8 有、GBK 字节没有
            if EA_EXPECT.encode("gbk") in raw and EA_EXPECT.encode("utf-8") not in raw:
                problems.append(f"{name}: 字体名被写成 GBK 乱码字节")

        if verbose:
            print(f"\n=== {path.name} ===")
            for g in ("幻灯片", "原生图表", "备注页", "母版/版式"):
                if g in per_part:
                    total = sum(per_part[g].values())
                    print(f"  {g}（{total} 段）：{per_part[g].most_common(3)}")

        try:
            theme_root = ET.fromstring(z.read("ppt/theme/theme1.xml"))
            scheme = theme_root.find(f".//{{{A}}}fontScheme")
            for tag in ("majorFont", "minorFont"):
                font = scheme.find(f"{{{A}}}{tag}")
                if font is None:
                    continue
                lv, ev = _face(font, "latin"), _face(font, "ea")
                hans = [x.get("typeface") for x in font.iter(f"{{{A}}}font")
                        if x.get("script") in ("Hans", "Hant")]
                if verbose:
                    print(f"  主题 {tag}: latin={lv!r} ea={ev!r} Hans={hans}")
                if ev != EA_EXPECT or hans != [EA_EXPECT] * len(hans) or not hans:
                    problems.append(f"主题 {tag}: ea/Hans 不是 {EA_EXPECT!r}（ea={ev!r} Hans={hans}）")
                if lv != LATIN_EXPECT:
                    problems.append(f"主题 {tag}: latin={lv!r}（应为 {LATIN_EXPECT!r}）")
        except KeyError:
            problems.append("ppt/theme/theme1.xml 缺失 —— 母版默认字体无从核对")

    if verbose:
        if problems:
            print(f"  [FAIL] {len(problems)} 处字体不对：")
            for p in problems[:10]:
                print(f"    - {p}")
        else:
            print("  [ ok ] 全部文字都是「中文微软雅黑 + 西文 Times New Roman」")
    return problems


def main(argv: list[str] | None = None, *, default_dir: Path | None = None) -> int:
    args = [Path(a) for a in (argv if argv is not None else sys.argv[1:])]
    if not args and default_dir is not None:
        args = sorted(p for p in Path(default_dir).glob("*.pptx")
                      if not p.name.startswith("~$"))
    if not args:
        print("用法: python .pi/renderer/fontcheck.py <deck.pptx> [...]")
        return 2
    total = 0
    for p in args:
        if not p.exists():
            print(f"[skip] {p} 不存在")
            continue
        total += len(audit(p))
    if total:
        print(f"\n[FAIL] 共 {total} 处字体问题")
        return 1
    print("\n[PASS] 字体口径无回归")
    return 0


if __name__ == "__main__":
    sys.exit(main())
