"""迁移期转发桩：真身在 `engine/dsl/briefs.py`（数据在 `.pi/briefs/`）。

搬迁结束后删掉本文件。
"""
import sys
from pathlib import Path

_ROOT = str(Path(__file__).resolve().parents[2])
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from engine.dsl import briefs as _real                  # noqa: E402

sys.modules[__name__] = _real
