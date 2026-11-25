# -*- coding: utf-8 -*-
"""分层结构回归：包边界没被旧路径、旧习惯悄悄破掉。

    python .pi/tests/layout.py

拆掉 `.pi/renderer/` 转发桩之后（2026-11），这套结构才有了"能被破坏"的可能，
所以必须有东西盯着。2026-11-25 代码整包迁进 `.pi/engine`，检查项跟着改，现在盯五件事：

1. **仓库根只剩 `AGENTS.md` + `requirements.txt`**。代码/资源一旦溜回根目录，
   `.pi` 就不再是唯一的家，文档里两套路径又会长回来。
2. **`.pi` 顶层没有与依赖同名的 `.py`**。`python .pi/cli.py` 会把 `.pi` 放进
   `sys.path[0]`；只要那儿出现 `pptx.py`，`import pptx` 就会命中它，报
   `'pptx' is not a package`——2026-11 之前入口就叫这个名字，靠"代码不在 .pi 下"
   绕开的。这条把绕法换成正面约束。
3. **`.pi/renderer/` 不再出现**。转发桩已拆，回来就说明有人又把代码塞进了 `.pi` 顶层。
4. **engine 内没有裸兄弟导入**。`import theme` / `from ppt_dsl import …` 这类写法
   以前能跑，纯粹是因为转发桩把 `.pi/renderer` 塞进了 `sys.path`；桩一拆就
   ModuleNotFoundError。更坏的是 `ppt_dsl._layouts()` 那种 `try/except: return None`，
   失败会被吞掉，表现为"几何容量校验静默失效"。用 AST 扫，import 时就报，不靠运气。
5. **入口的阶段表指向真实模块 / 文档里不再教旧命令**。`.pi/cli.py` 的 `STAGE_MODULES`
   写错一个名字，只有跑到那个阶段才知道；技能与 Prompt 是模型照着敲的东西，
   写 `python .pi/renderer/x.py` 等于让它必失败。

纯标准库，开发机（没装 python-pptx）也能跑。
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]      # 仓库根
PI_ROOT = REPO / ".pi"                          # 代码与资源的家
ENGINE = PI_ROOT / "engine"

# 仓库根允许出现的条目（此外一律算"溜回来了"）
ROOT_ALLOWED = {"AGENTS.md", "requirements.txt", ".pi", ".git", ".gitignore", ".workbuddy",
                ".codebuddy", ".vscode", ".pytest_cache", "__pycache__"}

# engine 里的模块名（裸导入如果命中这些，就是没走相对导入）
ENGINE_MODULES = {
    "paths", "build",
    "ppt_dsl", "briefs",
    "doc_analyzer", "outline", "planner", "plan", "budgets",
    "docx_in",
    "theme", "charts", "icons", "iconlib", "layouts", "geom",
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


def check_root_is_clean() -> None:
    print("== 1. 仓库根只有 AGENTS.md + requirements.txt ==")
    extra = sorted(p.name for p in REPO.iterdir() if p.name not in ROOT_ALLOWED)
    _check("根目录没有多余条目", not extra, f"多了：{extra}")
    _check(".pi/engine/ 存在", ENGINE.is_dir(), "代码包不在 .pi 下")
    _check("根目录没有 engine/", not (REPO / "engine").exists(), "代码又溜回根目录了")


def check_no_shadow_modules() -> None:
    print("== 2. .pi 顶层没有与依赖同名的 .py ==")
    # 依赖名从 requirements.txt 读，不手写清单：换依赖时这条自动跟着变。
    req = REPO / "requirements.txt"
    deps: set[str] = set()
    if req.exists():
        for line in req.read_text(encoding="utf-8").splitlines():
            line = line.split("#", 1)[0].strip()
            if line and not line.startswith("-"):
                deps.add(line.split("==")[0].split(">=")[0].split("[")[0].strip().lower())
    shadow = sorted(p.name for p in PI_ROOT.glob("*.py") if p.stem.lower() in deps)
    _check(".pi/*.py 与 requirements.txt 无重名", not shadow,
           f"重名：{shadow}（.pi 在 sys.path[0]，会顶掉第三方库）")
    _check("入口叫 cli.py（不是 pptx.py）", (PI_ROOT / "cli.py").exists())


def check_no_renderer_dir() -> None:
    print("== 3. 旧路径已拆除 ==")
    _check(".pi/renderer/ 不存在", not (PI_ROOT / "renderer").exists(),
           "有人又把代码或桩放进 .pi 顶层了")


def check_no_bare_imports() -> None:
    print("== 4. engine 内无裸兄弟导入 ==")
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
    actual = {p.stem for p in ENGINE.rglob("*.py") if p.name != "__init__.py"}
    _check("模块名集合与 engine 实际文件对得上", actual == ENGINE_MODULES,
           f"多: {sorted(actual - ENGINE_MODULES)} 少: {sorted(ENGINE_MODULES - actual)}")


def check_stage_modules() -> None:
    print("== 5. 入口阶段表指向真实模块 ==")
    entry = PI_ROOT / "cli.py"
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
        target = PI_ROOT / mod.replace(".", "/")
        ok = target.with_suffix(".py").exists() or (target / "__main__.py").exists()
        _check(f"{stage} → {mod}", ok, f"文件不存在：{target.relative_to(REPO)}")


def check_docs() -> None:
    print("== 6. 文档不教旧命令 ==")
    # 只扫"会给模型看"的文档；历史审计/根因记录另有"路径时效"注记，不在此列。
    docs = [REPO / "AGENTS.md", PI_ROOT / "setup.bat", PI_ROOT / "cli.py"]
    docs += sorted((PI_ROOT / "skills").rglob("*.md"))
    docs += sorted((PI_ROOT / "prompts").glob("*.md"))
    offenders: list[str] = []
    stale_entry: list[str] = []
    for d in docs:
        if not d.exists():
            continue
        for i, line in enumerate(d.read_text(encoding="utf-8").splitlines(), 1):
            if ".pi/renderer" in line or ".pi\\renderer" in line:
                # 允许：明确说明"已拆除 / 会报错 / 以前是桩"的句子
                if not any(k in line for k in ("不存在", "已拆除", "以前", "不再", "旧路径", "不要再照抄")):
                    offenders.append(f"{d.relative_to(REPO)}:{i} {line.strip()[:70]}")
            # 入口改名后，文档里不能再出现旧文件名（它会让 `.pi` 顶掉第三方 pptx 库）
            if ".pi/pptx.py" in line or ".pi\\pptx.py" in line:
                stale_entry.append(f"{d.relative_to(REPO)}:{i} {line.strip()[:70]}")
    _check("没有可照抄的旧命令", not offenders, "\n        ".join(offenders))
    _check("没有旧入口名 .pi/pptx.py", not stale_entry, "\n        ".join(stale_entry))


def main() -> int:
    check_root_is_clean()
    check_no_shadow_modules()
    check_no_renderer_dir()
    check_no_bare_imports()
    check_stage_modules()
    check_docs()
    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项不符合")
        return 1
    print("[PASS] 分层结构无回归 —— 根目录干净、无遮蔽模块、导入全是相对的、阶段表与文档都指对了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
