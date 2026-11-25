"""迁移期转发桩：真身在 `engine/plan/planner.py`。搬迁结束后删掉本文件。"""
import sys
from pathlib import Path

_PI = str(Path(__file__).resolve().parents[1])
if _PI not in sys.path:
    sys.path.insert(0, _PI)

from engine.plan import planner as _real                # noqa: E402

sys.modules[__name__] = _real
