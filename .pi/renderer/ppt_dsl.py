"""迁移期转发桩：真身已在 `engine/dsl/ppt_dsl.py`。

保留这个文件，是为了让 `.pi/renderer/*.py` 的 CLI、`.pi/tests/*` 和第三方
`import ppt_dsl` 全部不用改；搬迁结束后删掉本文件即可。

做法是**把本模块对象整体换成真模块**（而不是 `import *` 转发）：`import *`
不导出下划线开头的名字，也无法保证 `ppt_dsl is engine.dsl.ppt_dsl`，
而 isinstance / 模块级可变状态都依赖"同一个对象"。
"""
import sys
from pathlib import Path

_PI = str(Path(__file__).resolve().parents[1])          # .pi/
if _PI not in sys.path:
    sys.path.insert(0, _PI)

from engine.dsl import ppt_dsl as _real                 # noqa: E402

sys.modules[__name__] = _real
