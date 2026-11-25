# -*- coding: utf-8 -*-
"""转发桩完整性回归：锁定"搬迁期旧路径还能用"这件事。

搬迁（`.pi/renderer/x.py` → `engine/…/x.py`）会给旧路径留一个转发桩。
桩有两种坏法，都是**静默**的，正常回归发现不了：

  1. 被 import 时没换成真模块 —— 于是 `isinstance` 失效、模块级可变状态被复制成
     两份（theme.COLORS 会被 apply_style 就地改，两份就互相看不见）。
  2. 被直接执行时没转发 `__main__` —— `python .pi/renderer/docx_in.py in.docx`
     什么都不做、退出码 0；而 `.pi/pptx.py analyze` 正是用子进程调它，
     表现就是"命令成功但产物没生成"。这个坑真踩过一次。

所以这里两条都验，且不依赖任何第三方库（开发机也能跑）。

用法：`python .pi/tests/shims.py`
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / ".pi" / "renderer"
sys.path.insert(0, str(RENDERER))
sys.path.insert(0, str(REPO))

_mark = "sys.modules[__name__] = _real"
SHIMS = sorted(p for p in RENDERER.glob("*.py") if _mark in p.read_text(encoding="utf-8"))

_fail: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'ok' if ok else 'FAIL'}] {name}" + ("" if ok else f" —— {detail}"))
    if not ok:
        _fail.append(name)


def real_module(stub: Path):
    """从桩里读出真身模块路径，再 import 它。"""
    import importlib
    import re

    # 带 CLI 的桩把 import 放在 else: 里（避免 runpy 重复执行同一个模块），
    # 所以这里要允许行首缩进。
    m = re.search(r'^\s*from ([\w.]+) import (\w+) as _real', stub.read_text(encoding="utf-8"),
                  re.M)
    if not m:
        return None, None
    pkg, leaf = m.group(1), m.group(2)
    return f"{pkg}.{leaf}", importlib.import_module(f"{pkg}.{leaf}")


def main() -> int:
    print(f"== 1. 转发桩对象同一性（共 {len(SHIMS)} 个）==")
    for stub in SHIMS:
        full, real = real_module(stub)
        if real is None:
            check(f"{stub.name} 能认出真身", False, "桩里找不到 `from ... import ... as _real`")
            continue
        import importlib

        old = importlib.import_module(stub.stem)
        check(f"{stub.stem}: import 旧名拿到的就是 {full}",
              old is real, f"{old.__name__} vs {real.__name__}")

    print("== 2. 带 CLI 的真身：旧路径直接跑必须真的执行 ==")
    for stub in SHIMS:
        full, real = real_module(stub)
        if real is None:
            continue
        if f'runpy.run_module("{full}"' not in stub.read_text(encoding="utf-8"):
            # 真身没有 __main__ 就不该有转发段（两边都不跑主逻辑才是对的）
            has_main = (Path(real.__file__).read_text(encoding="utf-8").find("__main__") >= 0)
            check(f"{stub.stem}: 无 __main__ 的模块不该写转发段",
                  not has_main, f"{full} 有 __main__ 却没转发")
            continue
        r = subprocess.run([sys.executable, str(stub), "--help"],
                           capture_output=True, text=True, timeout=120)
        out = (r.stdout or "") + (r.stderr or "")
        # 允许"不认识 --help 而报错"，不允许"退出 0 且一句话都不说"
        check(f"{stub.stem}: `python .pi/renderer/{stub.stem}.py --help` 真的跑了",
              out.strip() != "", "输出为空、退出码 %d —— 转发段没生效" % r.returncode)

    print("== 3. 路径引导不能顶掉第三方库（`.pi` 不许进 sys.path）==")
    import importlib

    importlib.import_module(SHIMS[0].stem)          # 触发桩里的 sys.path 引导
    pptx = importlib.import_module("pptx")
    check("import pptx 拿到的是真库（不是 .pi/pptx.py 入口脚本）",
          hasattr(pptx, "__path__"), f"解析到 {getattr(pptx, '__file__', '?')}")

    print()
    if _fail:
        print(f"[FAIL] {len(_fail)} 项不通过")
        return 1
    print("[PASS] 转发桩全部正常")
    return 0


if __name__ == "__main__":
    sys.exit(main())
