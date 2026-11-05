# -*- coding: utf-8 -*-
"""图标门面：渲染器只从这里拿图标，实现全在 iconlib。

历史 composite 图标已弃用（线条粗糙、不够精致）。所有 draw_icon 都走 iconlib，
由它把 SVG 路径转成 PowerPoint 自定义几何图形（custGeom），真矢量、可改色、可缩放。

图标库（tabler 单库 / outline + filled）、命名解析与检索命令见 .pi/assets/icons/README.md。
这个文件只保留渲染器真正调用的那一个函数 —— 检索走 `python .pi/renderer/iconlib.py --search`，
别再在门面上加转发层。
"""
from __future__ import annotations


def draw_icon(slide, name: str, x: float, y: float, size: float,
              color: str = "primary", stroke_w: float | None = None,
              plate: bool = False) -> bool:
    """绘制 SVG 图标。坐标单位英寸。

    支持两种写法：`brain`（按 resolve_order 解析）、`tabler/filled/star`
    （指定库/风格），以及中文语义别名（如 `齿轮`）。

    成功返回 True；图标找不到或解析失败返回 False（调用方应降级处理）。
    """
    from iconlib import draw_icon as _draw
    return _draw(slide, name, x, y, size, color=color, stroke_w=stroke_w)
