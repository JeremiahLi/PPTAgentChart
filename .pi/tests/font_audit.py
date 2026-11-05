# -*- coding: utf-8 -*-
"""字体自检 CLI（薄壳）。

真正的实现在 `.pi/renderer/fontcheck.py`：`build.py` 渲染完会**在同一个进程里**
顺手跑一遍，不必再敲一条命令；这个文件只负责保留历史命令行的行为与文案：

    python .pi/tests/font_audit.py .pi/out/我的汇报.pptx
    python .pi/tests/font_audit.py .pi/out/*.pptx          # 多个也行
    python .pi/tests/font_audit.py                     # 不带参数 = 扫 .pi/out/ 下所有 pptx

核对口径（中文微软雅黑 + 西文 Times New Roman，逐段 latin/ea 成对、母版主题、
编码不是 GBK 乱码）见 fontcheck.py 的模块文档。
"""
from __future__ import annotations

import sys
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "out"
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "renderer"))

import fontcheck  # noqa: E402

audit = fontcheck.audit


def main() -> int:
    return fontcheck.main(None, default_dir=OUT)


if __name__ == "__main__":
    sys.exit(main())
