# -*- coding: utf-8 -*-
"""分层结构回归：包边界没被旧路径、旧习惯悄悄破掉。

    python .pi/tests/layout.py

拆掉 `.pi/renderer/` 转发桩之后（2026-11），这套结构才有了"能被破坏"的可能，
所以必须有东西盯着。检查四件事：

1. **旧路径不再存在**。`.pi/renderer/` 一旦回来，就说明有人又把代码/桩放进了
   `.pi` —— 而 `.pi` 进 `sys.path` 会把第三方 `pptx` 顶掉（`'pptx' is not a package`）。
2. **engine 内没有裸兄弟导入**。`import theme` / `from ppt_dsl import …` 这类写法
   以前能跑，纯粹是因为转发桩把 `.pi/renderer` 塞进了 `sys.path`；桩一拆就
   ModuleNotFoundError。更坏的是 `ppt_dsl._layouts()` 那种 `try/except: return None`，
   失败会被吞掉，表现为"几何容量校验静默失效"。用 AST 扫，import 时就报，不靠运气。
3. **入口的阶段表指向真实模块**。`.pi/pptx.py` 的 `STAGE_MODULES` 写错一个名字，
   只有跑到那个阶段才知道；这里提前对一遍文件是否存在。
4. **文档里不再教旧命令**。技能/Prompt 是模型照着敲的东西，写 `python .pi/renderer/x.py`
   等于让它必失败。只允许在"已拆除/会报错"这类说明句里出现。

纯标准库，开发机（没装 python-pptx）也能跑。
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ENGINE = REPO / "engine"

# engine 里的模块名（裸导入如果命中这些，就是没走相对导入）
ENGINE_MODULES = {
    "paths", "build",
    "ppt_dsl", "briefs",
    "doc_analyzer", "outline", "planner", "plan", "budgets",
    "docx_in",
    "theme", "charts", "icons", "iconlib", "layouts",
    "merge", "render", "speaker_notes",
    "qc", "fontcheck",
}

_failures: list[str] = []


def _check(name: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  [ ok ] {name}")
    else:
        _failures.append(f"{name}: {detail}")
        print(f"  [FAIL] {name}  {detail}")


def check_no_renderer_dir() -> None:
    print("== 1. 旧路径已拆除 ==")
    _check(".pi/renderer/ 不存在", not (REPO / ".pi" / "renderer").exists(),
           "有人又把代码或桩放回 .pi/ 了")
    _check("engine/ 在仓库根（不在 .pi 下）", ENGINE.is_dir() and not (REPO / ".pi" / "engine").exists())


def check_no_bare_imports() -> None:
    print("== 2. engine 内无裸兄弟导入 ==")
    bad: list[str] = []
    for py in sorted(ENGINE.rglob("*.py")):
        try:
            tree = ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
        except SyntaxError as exc:                     # 语法错由编译/回归暴露，这里不重复报
            bad.append(f"{py.relative_to(REPO)}: SyntaxError {exc.lineno}")
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for a in node.names:
                    if a.name.split(".")[0] in ENGINE_MODULES:
                        bad.append(f"{py.relative_to(REPO)}:{node.lineno} import {a.name}")
            elif isinstance(node, ast.ImportFrom):
                # level>0 就是相对导入，正确；level==0 才需要看名字
                if node.level == 0 and node.module and node.module.split(".")[0] in ENGINE_MODULES:
                    bad.append(f"{py.relative_to(REPO)}:{node.lineno} from {node.module} import …")
    _check("没有 import 裸模块名（一律相对导入）", not bad, "\n        ".join(bad))

    # 反向确认：锚点模块确实存在，避免集合写错导致上面永远为空
    _check("模块名集合与 engine 实际文件对得上",
           {(p.stem) for p in ENGINE.rglob("*.py") if p.name != "__init__.py"} == ENGINE_MODULES,
           f"多: {(set(p.stem for p in ENGINE.rglob('*.py') if p.name != '__init__.py') - ENGINE_MODULES)} "
           f"少: {ENGINE_MODULES - set(p.stem for p in ENGINE.rglob('*.py') if p.name != '__init__.py')}")


def check_stage_modules() -> None:
    print("== 3. 入口阶段表指向真实模块 ==")
    entry = REPO / ".pi" / "pptx.py"
    src = entry.read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(entry))
    table = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", "") == "STAGE_MODULES" for t in node.targets):
            table = ast.literal_eval(node.value)
    _check("找得到 STAGE_MODULES", isinstance(table, dict) and bool(table), str(table))
    if not isinstance(table, dict):
        return
    for stage, mod in table.items():
        target = REPO / mod.replace(".", "/")
        ok = target.with_suffix(".py").exists() or (target / "__main__.py").exists()
        _check(f"{stage} → {mod}", ok, f"文件不存在：{target.relative_to(REPO)}")


def check_docs() -> None:
    print("== 4. 文档不教旧命令 ==")
    # 只扫"会给模型看"的文档；历史审计/根因记录另有"路径时效"注记，不在此列。
    docs = [REPO / "AGENTS.md", REPO / "setup.bat"]
    docs += sorted((REPO / ".pi" / "skills").rglob("*.md"))
    docs += sorted((REPO / ".pi" / "prompts").glob("*.md"))
    offenders: list[str] = []
    for d in docs:
        if not d.exists():
            continue
        for i, line in enumerate(d.read_text(encoding="utf-8").splitlines(), 1):
            if ".pi/renderer" not in line and ".pi\\renderer" not in line:
                continue
            # 允许：明确说明"已拆除 / 会报错 / 以前是桩"的句子
            if any(k in line for k in ("不存在", "已拆除", "以前", "不再", "旧路径", "不要再照抄")):
                continue
            offenders.append(f"{d.relative_to(REPO)}:{i} {line.strip()[:70]}")
    _check("没有可照抄的旧命令", not offenders, "\n        ".join(offenders))


def main() -> int:
    check_no_renderer_dir()
    check_no_bare_imports()
    check_stage_modules()
    check_docs()
    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项不符合")
        return 1
    print("[PASS] 分层结构无回归 —— 旧路径已清、导入全是相对的、阶段表与文档都指对了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
