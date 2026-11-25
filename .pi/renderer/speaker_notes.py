"""迁移期转发桩：真身在 `engine/render/speaker_notes.py`。搬迁结束后删掉本文件。

两种调用都要通，缺一种就是静默失效：
  - 被 import → 整体换成真模块（不能 `import *`：导不出下划线名，也不保证对象同一）。
  - 本模块没有 CLI，所以不写 `__main__` 转发段。
"""
import sys
from pathlib import Path

_ROOT = str(Path(__file__).resolve().parents[2])        # 仓库根
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from engine.render import speaker_notes as _real  # noqa: E402

sys.modules[__name__] = _real
