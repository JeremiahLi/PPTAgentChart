"""迁移期转发桩：真身在 `engine/ingest/docx_in.py`。搬迁结束后删掉本文件。"""
import sys
from pathlib import Path

_ROOT = str(Path(__file__).resolve().parents[2])
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from engine.ingest import docx_in as _real              # noqa: E402

sys.modules[__name__] = _real
