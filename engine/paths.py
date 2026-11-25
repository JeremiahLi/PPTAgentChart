# -*- coding: utf-8 -*-
"""engine 里所有"找 .pi 下的资源"的唯一入口。

在这之前，每个模块各写一遍 `dirname(dirname(__file__))` 来反推 `.pi/` ——
只要文件被移进子目录（`engine/render/layouts.py`），这个推算就少一层，
路径会静默指到 `engine/` 上去，而错误要等到渲染时才以"找不到图片"的形式炸出来。
所以搬迁的第一步是把根路径收敛到这一个文件：**本文件固定放在 `<repo>/engine/paths.py`**，
`PACKAGE_ROOT` 永远是仓库根，别的模块只 import 常量，不再自己算。

为什么 `engine` 在仓库根、不在 `.pi/` 下：
要用 `from engine.dsl import ppt_dsl` 这种绝对导入，就得把仓库根加进 `sys.path`；
反过来若把 `.pi` 加进 path，`.pi/pptx.py`（唯一入口脚本）会把第三方 `pptx` 库
顶掉，报 `'pptx' is not a package`。`.pi/` 下任何一层再有同名 `.py` 都会重演，
所以代码包不能住在那儿。
"""
from __future__ import annotations

from pathlib import Path

PACKAGE_ROOT = Path(__file__).resolve().parents[1]   # 仓库根
PI_ROOT = PACKAGE_ROOT / ".pi"

BRIEF_DIR = PI_ROOT / "briefs"                       # 方向规约 yaml
ASSET_DIR = PI_ROOT / "assets"                       # 图标、模板、示例
WORK_DIR = PI_ROOT / "work"                          # 中间产物
OUT_DIR = PI_ROOT / "out"                            # 成品 pptx
