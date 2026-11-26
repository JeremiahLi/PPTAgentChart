# -*- coding: utf-8 -*-
"""`.pi/cli.py`（唯一入口 + 状态机 + 门禁）的回归：**纯标准库，不需要第三方依赖**。

    python .pi/tests/entry.py

为什么单独一个文件：`.pi/tests/` 里其他回归（smoke / chart_xml / font_audit）
全都要 python-pptx、matplotlib，只能在装好依赖的离线 PC 上跑；而入口层的逻辑
（跳步拦截、过期判定、失败标记、信封格式）恰恰是最容易改坏、又最该在开发机就挡住的部分。
这里用沙箱 + 桩模块，把 engine 换成 20 行的假实现，于是整套门禁可以在任何机器上验证。

沙箱做法：把 cli.py 拷到临时目录的 .pi/ 下——它按 `__file__` 上溯两级认仓库根，
所以拷过去就自动把临时目录当成仓库，不会碰真实工程。
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve()
ENTRY = HERE.parent.parent / "cli.py"           # <repo>/.pi/cli.py

# 桩模块：只做"写出产物 + 打印一行"，把 engine 的依赖全部绕开
STUB_PLAN = '''
import sys, pathlib
a = sys.argv[1:]
if "--emit-analysis" in a:
    i = a.index("--emit-analysis")
    out = a[i + 1] if i + 1 < len(a) and not a[i + 1].startswith("-") else ".pi/work/analysis.md"
    pathlib.Path(out).parent.mkdir(parents=True, exist_ok=True)
    pathlib.Path(out).write_text("| 章节 | 模块 | 核心结论 | 论据 | 缺口 |\\n", encoding="utf-8")
    print("骨架已写入:", out)
else:
    out = a[a.index("-o") + 1]
    pathlib.Path(out).parent.mkdir(parents=True, exist_ok=True)
    pathlib.Path(out).write_text('{"sections": []}', encoding="utf-8")
    print("大纲已保存:", out)
'''
STUB_BUDGETS = '''
import sys, pathlib
out = sys.argv[sys.argv.index("-o") + 1]
pathlib.Path(out).parent.mkdir(parents=True, exist_ok=True)
pathlib.Path(out).write_text("# 预算表\\n", encoding="utf-8")
'''
STUB_BUILD = '''
import sys, json, pathlib
a = sys.argv[1:]
out = a[a.index("-o") + 1]
pathlib.Path(out).parent.mkdir(parents=True, exist_ok=True)
pathlib.Path(out).write_text("fake pptx", encoding="utf-8")
print("RESULT " + json.dumps({"ok": True, "stage": "build", "out": out}))
'''
STUB_BUILD_FAIL = '''
import sys, json
print("RESULT " + json.dumps({"ok": False, "reason": "qc_error", "errors": [
    {"slide": 3, "type": "panel_2col", "field": "bullets", "message": "超上限"}]}))
'''

_ok = _bad = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global _ok, _bad
    if cond:
        _ok += 1
        print(f"  PASS  {name}")
    else:
        _bad += 1
        print(f"  FAIL  {name}" + (f"\n        {detail}" if detail else ""))


def run(root: Path, *args: str) -> tuple[int, str, dict | None]:
    """跑一次入口，返回 (退出码, 全文, 末行 RESULT)。"""
    env = {"PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    import os
    env = dict(os.environ, **env)
    p = subprocess.run([sys.executable, str(root / ".pi" / "cli.py"), *args],
                       cwd=str(root), capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env)
    out = (p.stdout or "") + (p.stderr or "")
    res = None
    for line in (p.stdout or "").splitlines():
        if line.startswith("RESULT "):
            try:
                res = json.loads(line[7:])
            except Exception:
                res = None
    return p.returncode, out, res


def codes(res: dict | None) -> list[str]:
    return [e.get("code", "") for e in (res or {}).get("errors", [])]


def main() -> int:
    if not ENTRY.exists():
        print(f"找不到 {ENTRY}")
        return 1
    root = Path(tempfile.mkdtemp(prefix="pptx-entry-"))
    try:
        (root / ".pi" / "engine" / "plan").mkdir(parents=True)
        (root / ".pi" / "work").mkdir(parents=True)
        shutil.copy2(ENTRY, root / ".pi" / "cli.py")
        # 入口用 `python -m engine.<阶段>` 起子进程（带 PYTHONPATH=.pi），
        # 所以桩必须摆成包结构，住在 .pi/engine 下，而不是以前的 .pi/renderer/*.py。
        eng = root / ".pi" / "engine"
        for pkg in (eng, eng / "plan"):
            (pkg / "__init__.py").write_text("", encoding="utf-8")
        for name, body in ((eng / "plan" / "plan.py", STUB_PLAN),
                           (eng / "plan" / "budgets.py", STUB_BUDGETS),
                           (eng / "build.py", STUB_BUILD)):
            name.write_text(body, encoding="utf-8")
        (root / "src.md").write_text("# 源文档\n正文内容，不少于三十个字，用于测试切分。\n",
                                     encoding="utf-8")

        print("空仓库")
        code, out, res = run(root)
        check("status 退出码 0", code == 0, out)
        check("空仓库指向 analyze", "analyze" in out and res and res.get("next") == "analyze", out)

        print("跳步拦截")
        code, out, res = run(root, "outline")
        check("跳 outline 被拦", res is not None and res.get("ok") is False, out)
        check("错误码 E_STAGE_MISSING", codes(res) == ["E_STAGE_MISSING"], str(codes(res)))
        check("业务失败退出码仍为 0", code == 0, f"exit={code}")
        code, out, res = run(root, "build")
        check("跳 build 被拦", codes(res) == ["E_STAGE_MISSING"], str(codes(res)))
        check("拦在根因（先补输入分析）", "输入分析" in out, out)

        print("逐阶段跑通")
        for stage, args in (("analyze", ("src.md",)), ("outline", ("--title", "T")),
                            ("budget", ())):
            code, out, res = run(root, stage, *args)
            check(f"{stage} ok", res is not None and res.get("ok") is True, out)
        check("analyze 产物存在", (root / ".pi/work/analysis.md").exists())
        check("outline 产物存在", (root / ".pi/work/outline.json").exists())
        check("budget 产物存在", (root / ".pi/work/budgets.md").exists())

        print("再跳 build")
        code, out, res = run(root, "build")
        check("没有 slides 时 build 被拦", codes(res)[0] in ("E_MISSING_INPUT", "E_STAGE_MISSING"),
              str(codes(res)))
        check("指向上一步（写 slides）", "逐页" in out or "slides" in out, out)

        print("写页后出图")
        (root / ".pi/work" / "slides").mkdir(exist_ok=True)
        (root / ".pi/work" / "slides" / "01.json").write_text(
            '[{"slide_type": "cover", "title": "测试"}]', encoding="utf-8")
        code, out, res = run(root, "build", "-o", ".pi/out/x.pptx")
        check("build ok", res is not None and res.get("ok") is True, out)
        check("产物落盘", (root / ".pi/out/x.pptx").exists())
        code, out, res = run(root)
        check("status 报全通", "整条流水线已跑通" in out, out)

        print("失败标记")
        (root / ".pi" / "engine" / "build.py").write_text(STUB_BUILD_FAIL, encoding="utf-8")
        code, out, res = run(root, "build", "-o", ".pi/out/x.pptx")
        check("失败码映射 E_QC_ERROR", codes(res)[0] == "E_QC_ERROR", str(codes(res)))
        check("子进程 errors 被带出来", len((res or {}).get("errors", [])) >= 2,
              str((res or {}).get("errors")))
        code, out, res = run(root)
        check("status 显示上次失败", "上次失败" in out, out)
        check("失败后不再自称全部跑通", "整条流水线已跑通" not in out, out)

        print("过期级联")
        (root / ".pi" / "engine" / "build.py").write_text(STUB_BUILD, encoding="utf-8")
        run(root, "build", "-o", ".pi/out/x.pptx")
        (root / "src.md").write_text("# 源文档\n内容变了。\n", encoding="utf-8")
        code, out, res = run(root)
        check("analyze 判为过期", "已过期" in out, out)
        check("下游被阻塞", "被阻塞" in out, out)
        check("下一步回到 analyze", res and res.get("next") == "analyze", str(res and res.get("next")))
        code, out, res = run(root, "build", "-o", ".pi/out/x.pptx")
        check("过期后 build 被拦", "E_STAGE_STALE" in codes(res), str(codes(res)))
        code, out, res = run(root, "build", "-o", ".pi/out/x.pptx", "--force")
        check("--force 绕过过期拦截", (res or {}).get("ok") is True, out)

        def states_of(result: dict | None) -> dict[str, str]:
            return {r["stage"]: r["state"] for r in (result or {}).get("stages", [])}

        print("人工补写分析表不算过期")
        code, out, res = run(root, "analyze", "src.md")  # 先把上面造的过期收回来
        check("重跑 analyze 复位", (res or {}).get("ok") is True, out)
        art = root / ".pi" / "work" / "analysis.md"
        art.write_text(art.read_text(encoding="utf-8") +
                       "| 四、自建扩展 | 手工补的结论 | 手工补的论据 | 无 |\n",
                       encoding="utf-8")
        code, out, res = run(root)
        check("产物被人工补写后 analyze 仍 done", states_of(res).get("analyze") == "done",
              str(states_of(res)))
        code, out, res = run(root, "outline", "--title", "T")
        check("补写后 outline 不被 stale 拦", (res or {}).get("ok") is True
              and "E_STAGE_STALE" not in codes(res), out + " / " + str(codes(res)))
        code, out, res = run(root)
        check("outline 收录了补写后的指纹", states_of(res).get("outline") == "done",
              str(states_of(res)))

        print("源文件变化仍判过期")
        (root / "src.md").write_text("# 源文档\n源文件这次真的变了，超过三十个字，用于验证输入指纹仍在校验。\n",
                                     encoding="utf-8")
        code, out, res = run(root)
        check("源文件变了 analyze 判过期", states_of(res).get("analyze") == "stale",
              str(states_of(res)))
        code, out, res = run(root, "outline", "--title", "T")
        check("源文件变了 outline 被拦", "E_STAGE_STALE" in codes(res), str(codes(res)))
    finally:
        shutil.rmtree(root, ignore_errors=True)

    print(f"\n{_ok} passed, {_bad} failed")
    return 0 if not _bad else 1


if __name__ == "__main__":
    sys.exit(main())
