# -*- coding: utf-8 -*-
"""engine 里所有"找 .pi 下的资源"的唯一入口。

在这之前，每个模块各写一遍 `dirname(dirname(__file__))` 来反推 `.pi/` ——
只要文件被移进子目录（`engine/render/layouts.py`），这个推算就少一层，
路径会静默指到 `engine/` 上去，而错误要等到渲染时才以"找不到图片"的形式炸出来。
所以根路径收敛到这一个文件：**本文件固定放在 `<repo>/.pi/engine/paths.py`**，
往上两层就是 `.pi`，别的模块只 import 常量，不再自己算。

为什么代码整包住 `.pi/engine`（而不是仓库根）：
2026-11-25 起仓库根只留 `AGENTS.md` + `requirements.txt`，可执行代码全部收进 `.pi/`。
这要求 `python .pi/cli.py` 把 `.pi` 塞进 `sys.path`（脚本目录天然是 `sys.path[0]`），
于是 `.pi` 顶层**不能有任何与第三方库同名的 `.py`** —— 原先的入口脚本叫 `pptx.py`，
一旦 `.pi` 进 path 就顶掉 `pptx` 库，报 `'pptx' is not a package`。
入口已改名 `cli.py`，约束随之解除。**新增的顶层脚本别叫 pptx / lxml / yaml /
numpy / PIL 这类名字**，否则同样的坑会重演，且只在特定 import 顺序下才炸。
"""
from __future__ import annotations

from pathlib import Path

PACKAGE_ROOT = Path(__file__).resolve().parents[1]   # 代码包根 = .pi
PI_ROOT = PACKAGE_ROOT                               # 资源根 = .pi（两者已合一）

BRIEF_DIR = PI_ROOT / "briefs"                       # 方向规约 yaml
ASSET_DIR = PI_ROOT / "assets"                       # 图标、模板、示例
WORK_DIR = PI_ROOT / "work"                          # 中间产物
OUT_DIR = PI_ROOT / "out"                            # 成品 pptx
REPO_ROOT = PI_ROOT.parent                           # 仓库根（只剩 AGENTS.md / requirements.txt）
