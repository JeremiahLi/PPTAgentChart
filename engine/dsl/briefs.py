# -*- coding: utf-8 -*-
"""写作方向（brief）：研讨型 / 汇报型。

为什么单独一层 —— 两个方向共用同一套配色和版式骨架，差别**不在长相**，
而在「怎么取舍原文」：研讨型逐条摆全、缺口明写；汇报型一页一个主张、
结论先行。所以这层是策略数据，不是渲染配置。

新增一个方向 = 往 `.pi/renderer/briefs/` 里放一个 yaml，零代码改动：
  - `writing` 段给 Agent 读（开工前 Read 一次，按它选内容、定页数）
  - `check`  段给 qc 读（页数区间、每页论据上限、标题禁用词）

渲染器不消费 brief 的版式，只消费它的红线 —— 改这里不会改外观。
"""
from __future__ import annotations

import os
from functools import lru_cache

import yaml

BRIEF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "briefs")
DEFAULT_BRIEF = "seminar"


def available() -> list[str]:
    if not os.path.isdir(BRIEF_DIR):
        return []
    return sorted(f[:-5] for f in os.listdir(BRIEF_DIR) if f.endswith(".yaml"))


@lru_cache(maxsize=None)
def load(name: str) -> dict:
    """读一个方向；名字不存在时抛 FileNotFoundError，报错里带上可选项。"""
    path = os.path.join(BRIEF_DIR, f"{name}.yaml")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"没有这个写作方向：{name}（可选：{'、'.join(available()) or '无'}）")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    data.setdefault("name", name)
    data.setdefault("label", name)
    data.setdefault("check", {})
    return data


def writing_of(name: str) -> str:
    """方向里的写作规范正文（给 Agent 读）。"""
    return (load(name).get("writing") or "").strip()


# 速度档（与 qc.SPEEDS 同源，两档）。流程约定的**默认值**：方向 yaml 没写
# `speed:` 段、或只覆盖了其中几个键时，缺的键从这里补。
DEFAULT_SPEED = "slow"
_SPEED_FALLBACK = {
    "slow": {"reviewOutline": True, "skeletonPreview": True,
             "requireZeroWarnings": True, "maxRefineRounds": 3, "notes": True},
    "fast": {"reviewOutline": False, "skeletonPreview": False,
             "requireZeroWarnings": False, "maxRefineRounds": 1, "notes": False},
}


def speed_of(name: str | None, speed: str = DEFAULT_SPEED) -> dict:
    """方向在该速度档下的流程约定。

    `requireZeroWarnings` 由 `build.py` 消费（决定退出码 3），其余三项给 Agent 读
    （大纲给不给用户过目、要不要骨架预览、回炉几轮）。方向名不存在时回退默认值
    —— 速度档不该因为 yaml 少写一段就让整份 deck 出不来。
    """
    base = dict(_SPEED_FALLBACK.get(speed, _SPEED_FALLBACK[DEFAULT_SPEED]))
    if not name:
        return base
    try:
        spec = load(name).get("speed") or {}
    except FileNotFoundError:
        return base
    base.update(spec.get(speed) or {})
    return base
