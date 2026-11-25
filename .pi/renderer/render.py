"""迁移期转发桩：真身在 `engine/render/render.py`。搬迁结束后删掉本文件。

两种调用都要通，缺一种就是静默失效：
  - 被 import → 整体换成真模块（不能 `import *`：导不出下划线名，也不保证对象同一）。
  - 被 `python .pi/renderer/render.py ...` 直接跑 → 转给真身的 `__main__`。
"""
import sys
from pathlib import Path

_ROOT = str(Path(__file__).resolve().parents[2])        # 仓库根
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

if __name__ == "__main__":
    import runpy

    runpy.run_module("engine.render.render", run_name="__main__", alter_sys=True)
else:
    from engine.render import render as _real

    sys.modules[__name__] = _real
