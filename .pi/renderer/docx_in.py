"""迁移期转发桩：真身在 `engine/ingest/docx_in.py`。搬迁结束后删掉本文件。

两种调用都要通，缺一种就是静默失效：
  - 被 import → 整体换成真模块。不能写成 `from ... import *`：那样导不出下划线
    开头的名字，也保证不了 `该模块 is 真模块`，而 isinstance 和模块级可变状态
    （theme.COLORS 会被 apply_style 就地改）都依赖"同一个对象"。
  - 被 `python .pi/renderer/docx_in.py ...` 直接跑 → 转给真身的 `__main__`。
    少这一段会变成"什么都不做、退出码 0"，而 `pptx.py` 正是用子进程调这些脚本，
    于是表现为"命令成功但产物没生成"。
"""
import sys
from pathlib import Path

_ROOT = str(Path(__file__).resolve().parents[2])        # 仓库根
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

if __name__ == "__main__":
    import runpy

    runpy.run_module("engine.ingest.docx_in", run_name="__main__", alter_sys=True)
else:
    from engine.ingest import docx_in as _real

    sys.modules[__name__] = _real
