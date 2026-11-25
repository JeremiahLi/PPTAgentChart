# -*- coding: utf-8 -*-
"""PPT Agent 的唯一入口：一条命令跑完整条流水线，跳步会当场失败。

    python .pi/pptx.py                          # 看状态（等价 status）
    python .pi/pptx.py analyze 材料.docx         # Word 预处理 + 分析表骨架
    python .pi/pptx.py outline --title "主题"    # 分析表 → 结构化大纲
    python .pi/pptx.py budget                   # 大纲 → 逐页字数预算表
    python .pi/pptx.py check                    # 只质检，不出文件
    python .pi/pptx.py build -o .pi/out/x.pptx  # 合并 + 质检 + 渲染 + 字体自检

为什么有这个文件（一句话）：**流程以前写在 markdown 里，靠模型读、靠模型记**。
`ppt-entry.md` 描述的"分析 → 大纲 → 预算 → 写页 → 出图"没有任何东西保证被按序执行，
模型跳过 `budgets.py` 直接写 slides，`build.py` 照样出图，直到长度校验报错才返工。
本文件把这条顺序搬进代码：每个阶段跑完写 `.pi/work/state.json`（输入/产物内容指纹），
下一阶段开跑前先核对上游产物是否在、是否已过期，不在就报错并给出该跑哪条命令。

四条硬约定（与 build.py / merge.py 对齐，不另立一套）：

1. **退出码恒 0**。校验拦截、产物缺失、上游过期都是**正常业务结果**，不是运行失败；
   成败只看末行 `RESULT {json}` 的 `ok`。只有未捕获异常才非零退出。
   以前 plan.py 退 2、build.py 恒 0、render.py 文档写"非零退出"——三种语汇，模型得逐个记。
2. **统一信封**：`{"ok","stage","reason","errors":[{code,path,hint}],"artifacts":[...],"next":{...}}`。
   `code` 是稳定枚举（`E_STAGE_MISSING` 这类），中文只放在 `message` / `hint` 里，
   这样回炉时能按 code 统计、能 diff，不必靠模型"读懂一句话"。
3. **指纹判新旧**：状态里存内容的 sha256 前 12 位，不比 mtime——
   `git checkout` 会改 mtime 但内容没变，比 mtime 会误报过期。
4. **只做适配层**：内部仍调 `.pi/renderer/*.py`，不复制任何判定逻辑
   （长度上限、几何容量、字体口径全部留在原处）。以后把 renderer 拆成 engine 库时，
   只需要改本文件的 `STAGE_SCRIPTS`，对外契约不变。

门禁分两级，别混：
  **requires**（缺了直接报错）—— 上游产物是当前阶段的前提，缺了跑下去必然白跑；
  **advises**（缺了只告警）—— 能跑但会返工，例如没跑 budget 就出图。
  确认要硬闯：`--force`（会写进 RESULT 的 `forced: true`，可追溯）。

不用第三方依赖就能跑 `status` / `--dry-run`：阶段脚本按需启动子进程，
所以开发机（没装 python-pptx）也能看流程状态，不会一上来就 ImportError。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path

# 仓库根：本文件在 <repo>/.pi/pptx.py，上溯两级
REPO = Path(__file__).resolve().parent.parent
STATE = REPO / ".pi" / "work" / "state.json"
PY = sys.executable

# ---- 产物路径（相对仓库根；状态文件里存的也是相对路径，换机器可读）----
ART_INPUT = ".pi/work/input.md"          # docx 预处理产物
ART_ANALYSIS = ".pi/work/analysis.md"    # 分析表（骨架由代码生成，结论/缺口由模型补）
ART_OUTLINE = ".pi/work/outline.json"    # 结构化大纲
ART_BUDGETS = ".pi/work/budgets.md"      # 逐页字数预算
ART_SLIDES = ".pi/work/slides"           # 逐页 DSL（模型写，非脚本产物）
ART_OUT = ".pi/out/deck.pptx"
ART_MEDIA = ".pi/work/media"
_CWD0 = REPO  # 用户调用时的 cwd（main 里覆盖），用于解析他给的相对路径

STATUS_LABEL = {
    "done": "✓ 完成",
    "untracked": "? 有产物(无记录)",
    "stale": "! 已过期",
    "failed": "× 上次失败",
    "pending": "· 未做",
    "blocked": "× 被阻塞",
    "skipped": "- 不适用",
}

# 子进程脚本位置：内部拆分时只改这里
STAGE_SCRIPTS = {
    "docx_in": ".pi/renderer/docx_in.py",
    "plan": ".pi/renderer/plan.py",
    "budgets": ".pi/renderer/budgets.py",
    "build": ".pi/renderer/build.py",
}


class Stage:
    """一个阶段 = 一段确定性的动作 + 它吃/吐哪些文件。

    requires / advises 填的是**阶段 key**（不是文件路径）：
    门禁问的是"上游那一步做完了没、过期没"，而不是"某个文件在不在"——
    后者会把"文件在但内容是旧的"当成通过，正是返工的来源。
    """

    def __init__(self, key, label, requires=(), advises=(), produces=(), manual=False):
        self.key = key
        self.label = label
        self.requires = tuple(requires)
        self.advises = tuple(advises)
        self.produces = tuple(produces)
        self.manual = manual

    @property
    def short(self):
        return self.key


# 顺序即流水线顺序（status 按它渲染，门禁按 requires/advises 递归）
STAGES = [
    # analyze 顺带产出 input.md（docx 预处理），但状态只锚定 analysis.md：
    # input.md 是派生物，有人手工修过它不该被判成"过期"再被覆盖掉。
    Stage("analyze", "输入分析", requires=(), produces=(ART_ANALYSIS,)),
    Stage("outline", "大纲规划", requires=("analyze",), produces=(ART_OUTLINE,)),
    Stage("budget", "字数预算", requires=("outline",), produces=(ART_BUDGETS,)),
    Stage("slides", "逐页 JSON", requires=("outline",), produces=(ART_SLIDES,), manual=True),
    Stage("build", "出图", requires=("slides",), advises=("budget", "analyze"),
          produces=(ART_OUT,)),
]
BY_KEY = {s.key: s for s in STAGES}


# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------
def _utf8_streams():
    """GBK 控制台上打印中文会 UnicodeEncodeError：降到 replace，不因排版崩掉流程。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _abs(rel_or_abs) -> Path:
    p = Path(rel_or_abs)
    return p if p.is_absolute() else (REPO / p)


def _rel(p: Path) -> str:
    """状态文件里统一存相对路径；仓库外的输入（用户给的 docx）存绝对路径。"""
    try:
        return p.resolve().relative_to(REPO).as_posix()
    except ValueError:
        return p.resolve().as_posix()


def sha_path(target) -> str | None:
    """内容指纹：文件取字节 hash，目录取「相对路径 + 内容」排序后的 hash。

    不用 mtime：git checkout / 解压 / 拷贝都会改 mtime 而内容没变，会误报过期。
    目录里跳过 ~$ 开头的 Office 锁文件（它们每次打开都变，纯噪音）。
    """
    p = _abs(target)
    if not p.exists():
        return None
    h = hashlib.sha256()
    if p.is_dir():
        files = sorted(f for f in p.rglob("*") if f.is_file() and not f.name.startswith("~$"))
        if not files:
            return None
        for f in files:
            h.update(f.relative_to(p).as_posix().encode("utf-8"))
            h.update(b"\0")
            h.update(f.read_bytes())
    else:
        h.update(p.read_bytes())
    return h.hexdigest()[:12]


def now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


def load_state() -> dict:
    if not STATE.exists():
        return {"version": 1, "stages": {}}
    try:
        data = json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        # 状态坏了不该拦住干活：当作空状态，下一次写入自愈
        return {"version": 1, "stages": {}}
    data.setdefault("version", 1)
    data.setdefault("stages", {})
    return data


def save_state(state: dict) -> None:
    STATE.parent.mkdir(parents=True, exist_ok=True)
    state["updated"] = now()
    tmp = STATE.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE)


def mark_failed(key: str, code: str) -> None:
    """记一次失败。不删成功记录（产物还在），只打一个"更晚的失败"标记。"""
    state = load_state()
    state.setdefault("last_error", {})[key] = {"code": code, "at": now()}
    save_state(state)


def record(key: str, inputs: dict, outputs, elapsed: float, args: dict | None = None) -> None:
    """记一次成功：inputs 是 {名字: 路径}，输出是路径列表。"""
    state = load_state()
    state.setdefault("last_error", {}).pop(key, None)  # 成功即清除失败标记
    state["stages"][key] = {
        "at": now(),
        "elapsed": round(elapsed, 2),
        "inputs": {n: {"path": _rel(_abs(p)), "sha": sha_path(p)} for n, p in inputs.items()},
        "outputs": {_rel(_abs(o)): sha_path(o) for o in outputs},
        "args": args or {},
    }
    save_state(state)


# --------------------------------------------------------------------------
# 状态评估：done / stale / pending / blocked
# --------------------------------------------------------------------------
def stage_state(key: str) -> str:
    """单个阶段的实际状态（不含上游传导）。

    - 没记录也没产物      → pending
    - 有记录但产物没了     → pending（被手工删了，当作没做过）
    - 产物在但内容变了     → stale（有人手改过产物）
    - 任一输入内容变了     → stale（上游重跑过，这一步的结论已过期）
    """
    st = BY_KEY[key]
    if st.manual:
        if key == "slides":
            p = REPO / ART_SLIDES
            has = p.is_dir() and any(f.suffix == ".json" for f in p.glob("*.json")
                                     if not f.name.startswith("~$"))
            return "done" if has else "pending"

    rec = load_state()["stages"].get(key)
    # 失败标记由成功清除（见 record），所以这里不用比时间戳——
    # 时间戳只有秒级精度，「同一秒内失败后重试成功」会被误判成还在失败。
    if (load_state().get("last_error") or {}).get(key):
        return "failed"
    if not rec:
        # 迁移期兼容：产物在、但没有指纹记录（新入口上线前跑的老流程）。
        # 不能当成 done——无法判定是否过期；也不能当成 pending——会误拦已经能跑的工作。
        # 标 untracked，门禁放行但告警，重跑一次就把指纹收进来。
        outs = st.produces
        if outs and all(sha_path(p) is not None for p in outs):
            return "untracked"
        return "pending"
    for rel, sha in rec["outputs"].items():
        cur = sha_path(rel)
        if cur is None:
            return "pending"
        if cur != sha:
            return "stale"
    for item in rec["inputs"].values():
        cur = sha_path(item["path"])
        if cur is None:
            return "pending"
        if cur != item["sha"]:
            return "stale"
    return "done"


def upstream_issues(key: str) -> tuple[list[dict], list[dict]]:
    """递归收集上游的硬/软问题。返回 (errors, warnings)。"""
    st = BY_KEY[key]
    errors, warns = [], []

    def walk(name, severity, seen):
        if name in seen or name not in BY_KEY:
            return
        seen.add(name)
        # 先追根再报本级：slides 和 outline 都没做时，该让他先去写大纲，
        # 而不是"去写 slides"（后者在缺大纲时根本写不对）。
        if severity == "error":
            for up in BY_KEY[name].requires:
                walk(up, severity, seen)
        s = stage_state(name)
        target = errors if severity == "error" else warns
        if s == "pending":
            target.append({
                "code": "E_STAGE_MISSING" if severity == "error" else "W_STAGE_MISSING",
                "stage": name,
                "label": BY_KEY[name].label,
                "message": f"上游「{BY_KEY[name].label}」还没做",
                "hint": next_cmd(name),
            })
        elif s == "untracked":
            warns.append({
                "code": "W_STAGE_UNTRACKED",
                "stage": name,
                "label": BY_KEY[name].label,
                "message": f"上游「{BY_KEY[name].label}」有产物但没有指纹记录，无法判定是否过期",
                "hint": f"跑一次 `{next_cmd(name)}` 收录指纹（或 --force 跳过告警）",
            })
        elif s == "stale":
            target.append({
                "code": "E_STAGE_STALE" if severity == "error" else "W_STAGE_STALE",
                "stage": name,
                "label": BY_KEY[name].label,
                "message": f"上游「{BY_KEY[name].label}」的输入已变，结论过期",
                "hint": next_cmd(name),
            })
        elif s == "failed":
            target.append({
                "code": "E_STAGE_FAILED" if severity == "error" else "W_STAGE_FAILED",
                "stage": name,
                "label": BY_KEY[name].label,
                "message": f"上游「{BY_KEY[name].label}」上次跑失败了，产物可能是旧的",
                "hint": next_cmd(name),
            })
        # 软依赖不传染：一条 warning 不该顺着链路滚成一串
    seen: set[str] = set()
    for up in st.requires:
        walk(up, "error", seen)
    for up in st.advises:
        walk(up, "warn", seen)
    return errors, warns


def next_cmd(key: str) -> str:
    """说人话的"下一步"：把阶段名翻成可以直接粘的命令。"""
    return {
        "analyze": "python .pi/pptx.py analyze <源文件.docx|.md>",
        "outline": "python .pi/pptx.py outline --title \"主题\"",
        "budget": "python .pi/pptx.py budget",
        "slides": "按 ppt-common 第 2 步写 .pi/work/slides/*.json（一个文件 5-8 页）",
        "build": "python .pi/pptx.py build -o .pi/out/<主题>.pptx",
    }.get(key, f"python .pi/pptx.py {key}")


def last_out() -> str:
    """最后一次实际写到哪就报哪（--out 可以是任意路径，不能写死默认值）。"""
    built = load_state()["stages"].get("build") or {}
    return next(iter(built.get("outputs", {})), ART_OUT)


def pipeline_rows() -> list[dict]:
    """按流水线顺序算一遍状态；上游坏了，下游一律标"被阻塞"。

    顺序很重要：STAGES 已经是拓扑序，所以"第一个非完成阶段"必然就是根因，
    status 直接把它当"下一步"，不用再爬依赖图。
    """
    states = {st.key: stage_state(st.key) for st in STAGES}
    for st in STAGES:  # 拓扑序，上游先算，blocked 能继续往下传
        if any(states.get(u) not in ("done", "untracked") for u in st.requires):
            states[st.key] = "blocked"
    rows = []
    out_rel = last_out()
    for st in STAGES:
        art = ART_SLIDES if st.key == "slides" else (out_rel if st.key == "build" else
                                                     (st.produces[0] if st.produces else ""))
        rows.append({"stage": st.key, "label": st.label, "state": states[st.key],
                     "artifact": art, "manual": st.manual})
    return rows


# --------------------------------------------------------------------------
# 执行：子进程 + 信封
# --------------------------------------------------------------------------
def _display_width(s: str) -> int:
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)


def _pad(s: str, width: int) -> str:
    return s + " " * max(0, width - _display_width(s))


def emit(ok: bool, stage: str, **extra) -> None:
    """末行 RESULT：ensure_ascii 让 GBK 管道也能无损解析（沿用 build.py 的做法）。"""
    payload = {"ok": ok, "stage": stage, "at": now(), **extra}
    print("RESULT " + json.dumps(payload, ensure_ascii=True))


def _strip_result(text: str) -> tuple[str, dict | None]:
    """把子进程末行 RESULT 摘出来，正文原样保留（不然会出现两个 RESULT）。"""
    body, found = [], None
    for line in text.splitlines():
        if line.startswith("RESULT "):
            try:
                found = json.loads(line[len("RESULT "):])
            except Exception:
                found = None
            continue
        body.append(line)
    return "\n".join(body), found


def run_child(script_key: str, argv: list[str], *, dry: bool = False, quiet: bool = False):
    """跑一个阶段脚本。返回 (退出码, 子进程 RESULT, 耗时)。

    子进程的 stdout 原样转发（报告是模型要读的东西），只摘掉它自己的 RESULT 行。
    """
    script = REPO / STAGE_SCRIPTS[script_key]
    if not script.exists():
        return 127, {"ok": False, "reason": "script_missing",
                     "errors": [f"找不到 {STAGE_SCRIPTS[script_key]}"]}, 0.0
    cmd = [PY, str(script)] + argv
    if dry:
        print(f"  [dry-run] {' '.join(cmd)}")
        return 0, {"ok": True, "dry_run": True}, 0.0

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    t0 = time.perf_counter()
    proc = subprocess.run(cmd, cwd=str(REPO), capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
    elapsed = time.perf_counter() - t0

    body, child = _strip_result(proc.stdout or "")
    if body.strip() and not quiet:
        for line in body.rstrip().splitlines():
            print("  " + line)
    if proc.returncode != 0 and (proc.stderr or "").strip() and not quiet:
        for line in (proc.stderr or "").rstrip().splitlines():
            print("  ! " + line, file=sys.stderr)
    return proc.returncode, child, elapsed


def _map_reason(reason: str | None) -> str:
    """把 build/merge 的 reason 翻成稳定错误码。"""
    return {
        "slides_invalid": "E_SLIDES_INVALID",
        "qc_error": "E_QC_ERROR",
        "warnings_not_zero": "E_WARNINGS_NOT_ZERO",
        "font_check": "E_FONT_CHECK",
        "slides_dir_missing": "E_MISSING_INPUT",
    }.get(reason or "", "E_STAGE_FAILED")


def fail(stage: str, code: str, message: str, *, hint: str = "", details=None,
         forced: bool = False, ran: bool = False, **kw) -> int:
    """统一的失败出口。

    注意：子进程返回的 RESULT 里可能带 `ok` / `stage` 键，直接 **child 展开会
    和 emit 自己的同名参数撞车（TypeError: got multiple values）。所以这里只收
    白名单字段（details / forced / kw），调用方不许直接展开 child。
    """
    print(f"✗ {message}", file=sys.stderr)
    if hint:
        print(f"  → {hint}", file=sys.stderr)
    if ran:  # 只有"真的跑了但没过"才算失败；门禁拦下来是 blocked，不是 failed
        mark_failed("build" if stage == "check" else stage, code)
    entry = {"code": code, "message": message}
    if hint:
        entry["hint"] = hint
    errs = [entry]
    for e in details or []:
        if isinstance(e, dict) and e.get("code") != code:
            errs.append(e)
        elif isinstance(e, str):
            errs.append({"code": code, "message": e})
    emit(False, stage, reason=code.lower().replace("e_", "", 1).replace("w_", "", 1),
         forced=forced, errors=errs, **kw)
    return 0


def _gate(key: str, force: bool) -> tuple[bool, list[dict], list[dict]]:
    """跑之前的门禁：硬问题不让跑，软问题只记下来放进信封。"""
    errors, warns = upstream_issues(key)
    if force and errors:
        for e in errors:
            e["forced"] = True
        print(f"  ! --force：忽略 {len(errors)} 个上游缺口，后果自负", file=sys.stderr)
        errors = []
    return (not errors), errors, warns


def _child_details(child) -> list | None:
    """只取子进程 RESULT 里的 errors 列表。

    子进程的 RESULT 还带 ok/stage/at 等键，`**child` 展开会和 emit 的参数撞车，
    所以这里做白名单——只放行 errors。
    """
    if not isinstance(child, dict):
        return None
    errs = child.get("errors")
    return errs if isinstance(errs, list) else None


def _check_artifact(stage: str, paths, force: bool) -> tuple[bool, list[dict]]:
    """脚本说成功了，产物却不在——这是"脚本口径"和"磁盘事实"打架，必须以磁盘为准。"""
    missing = [p for p in paths if sha_path(p) is None]
    if not missing:
        return True, []
    if force:
        return True, [{"code": "W_NO_ARTIFACT", "paths": missing}]
    return False, [{"code": "E_NO_ARTIFACT", "paths": missing,
                    "message": f"脚本报告成功，但产物不存在：{', '.join(missing)}",
                    "hint": "看子进程输出里的报错；或加 --force 忽略"}]


# --------------------------------------------------------------------------
# 子命令
# --------------------------------------------------------------------------
def cmd_status(a) -> int:
    print(f"仓库  {REPO}")
    print(f"状态  {_rel(STATE)}" + ("" if STATE.exists() else "（还没跑过任何阶段）"))
    print()
    rows = pipeline_rows()
    state = load_state()
    w_label = max(_display_width(r["label"]) for r in rows) + 2
    w_state = max(_display_width(v) for v in STATUS_LABEL.values()) + 2
    w_art = max(_display_width(r["artifact"]) for r in rows if r["artifact"]) + 2
    for r in rows:
        line = ("  " + _pad(r["label"], w_label) + _pad(STATUS_LABEL[r["state"]], w_state)
                + _pad(r["artifact"], w_art))
        if r["state"] == "failed":
            err = (state.get("last_error") or {}).get(r["stage"], {})
            line += f"上次失败 {err.get('at', '')[11:16]}  {err.get('code', '')}"
        else:
            rec = state["stages"].get(r["stage"])
            if rec:
                line += f"{rec['at'].replace('T', ' ')}  {rec['elapsed']}s"
        print(line.rstrip())

    print()
    nxt = next((r for r in rows if r["state"] in ("pending", "stale", "blocked", "failed")), None)
    if nxt:
        print(f"下一步  {next_cmd(nxt['stage'])}")
    else:
        print(f"整条流水线已跑通：{last_out()}")

    # 每条命令都以 RESULT 收尾（含 status）——"成败看末行"这条契约不能有例外，
    # 否则模型/脚本还得记住"哪几个命令没有信封"。
    emit(True, "status", stages=rows, next=(nxt["stage"] if nxt else None))
    return 0


def cmd_analyze(a) -> int:
    src = Path(a.source)
    if not src.is_absolute():
        # 相对路径按"用户敲命令时的目录"解，其次才按仓库根——
        # 因为 main() 为了内部脚本的 `.pi/work/...` 会把 cwd 切到仓库根，
        # 直接用 cwd 会把 `cd 别的目录 && python <repo>/.pi/pptx.py analyze ./x.md` 解错。
        cand = (_CWD0 / src)
        src = cand.resolve() if cand.exists() else _abs(src)
    if not src.exists():
        return fail("analyze", "E_MISSING_INPUT", f"源文件不存在：{src}",
                    hint="确认路径；.docx 要先能被 Word 打开（加密/损坏读不了）")
    if src.name.lower() == "analysis.md":
        return fail("analyze", "E_USAGE", "--emit-analysis 的输入不能是 analysis.md 本身",
                    hint="给它原文或 docx；分析表已存在就直接跑 outline")

    ok, errors, warns = _gate("analyze", a.force)
    if not ok:
        return fail("analyze", errors[0]["code"], errors[0]["message"],
                    hint=errors[0].get("hint", ""), details=errors)

    suffix = src.suffix.lower()
    inputs = {"source": src}

    # ① docx 预处理（去修订 + 抽图），只有 Word 才需要；.md 直接当分析输入
    if suffix in (".docx", ".doc"):
        print(f"① 预处理 {src.name}")
        code, child, _ = run_child("docx_in", [str(src), "-o", ART_INPUT, "--media", ART_MEDIA],
                                   dry=a.dry_run, quiet=a.quiet)
        if code != 0 or (child and not child.get("ok", True)):
            return fail("analyze", "E_DOCX_IN", "Word 预处理失败",
                        hint="看上面 docx_in 的报错：加密/损坏的 docx 读不了",
                        details=_child_details(child), ran=True)
        body_in = ART_INPUT
    else:
        body_in = str(src)

    # ② 分析表骨架：章节/模块/论据/数字自动填，结论与缺口留给人/模型补
    print("② 生成分析表骨架（结论 / 缺口两列待补）")
    code, child, elapsed = run_child("plan", [body_in, "--emit-analysis", ART_ANALYSIS],
                                     dry=a.dry_run, quiet=a.quiet)
    if a.dry_run:
        emit(True, "analyze", dry_run=True, artifacts=[], warnings=warns)
        return 0
    if code != 0:
        return fail("analyze", "E_PLAN", "分析表骨架生成失败",
                    hint="确认原文有 # / ## 标题、每节正文不少于 30 字",
                    details=_child_details(child), ran=True)

    good, art_err = _check_artifact("analyze", [ART_ANALYSIS], a.force)
    if not good:
        return fail("analyze", "E_NO_ARTIFACT", art_err[0]["message"], hint=art_err[0]["hint"],
                    ran=True)

    record("analyze", inputs, [ART_ANALYSIS],
           elapsed, {"source": str(src), "title": a.title})
    print(f"✓ {ART_ANALYSIS} 已就绪——下一步：补「核心结论 / 缺口」两列，再跑 outline")
    emit(True, "analyze", artifacts=[ART_ANALYSIS], warnings=warns,
         next={"stage": "outline", "cmd": next_cmd("outline")}, elapsed=round(elapsed, 2))
    return 0


def cmd_outline(a) -> int:
    ok, errors, warns = _gate("outline", a.force)
    if not ok:
        return fail("outline", errors[0]["code"], errors[0]["message"],
                    hint=errors[0].get("hint", ""), details=errors)

    argv = [ART_ANALYSIS, "-o", ART_OUTLINE, "--brief", a.brief]
    if a.title:
        argv += ["--title", a.title]
    if a.subtitle:
        argv += ["--subtitle", a.subtitle]
    code, child, elapsed = run_child("plan", argv, dry=a.dry_run, quiet=a.quiet)
    if a.dry_run:
        emit(True, "outline", dry_run=True, warnings=warns)
        return 0
    if code != 0:
        return fail("outline", "E_PLAN", "大纲生成失败",
                    hint="分析表格式对不对？表头应为「章节|模块|核心结论|论据|缺口」",
                    details=_child_details(child), ran=True)

    good, art_err = _check_artifact("outline", [ART_OUTLINE], a.force)
    if not good:
        return fail("outline", "E_NO_ARTIFACT", art_err[0]["message"], hint=art_err[0]["hint"],
                    ran=True)

    record("outline", {"analysis": ART_ANALYSIS}, [ART_OUTLINE], elapsed,
           {"title": a.title, "brief": a.brief})
    print(f"✓ {ART_OUTLINE} 已就绪——慢档在这里停下来让用户过目一次，然后跑 budget")
    emit(True, "outline", artifacts=[ART_OUTLINE], warnings=warns,
         next={"stage": "budget", "cmd": next_cmd("budget")}, elapsed=round(elapsed, 2))
    return 0


def cmd_budget(a) -> int:
    ok, errors, warns = _gate("budget", a.force)
    if not ok:
        return fail("budget", errors[0]["code"], errors[0]["message"],
                    hint=errors[0].get("hint", ""), details=errors)

    code, child, elapsed = run_child("budgets", [ART_OUTLINE, "-o", ART_BUDGETS],
                                     dry=a.dry_run, quiet=a.quiet)
    if a.dry_run:
        emit(True, "budget", dry_run=True, warnings=warns)
        return 0
    if code != 0:
        return fail("budget", "E_BUDGETS", "预算表生成失败",
                    details=_child_details(child), ran=True)

    good, art_err = _check_artifact("budget", [ART_BUDGETS], a.force)
    if not good:
        return fail("budget", "E_NO_ARTIFACT", art_err[0]["message"], hint=art_err[0]["hint"],
                    ran=True)

    record("budget", {"outline": ART_OUTLINE}, [ART_BUDGETS], elapsed)
    print(f"✓ {ART_BUDGETS} 已就绪——照预算写批次（一个文件 5-8 页），再跑 build")
    emit(True, "budget", artifacts=[ART_BUDGETS], warnings=warns,
         next={"stage": "slides", "cmd": next_cmd("slides")}, elapsed=round(elapsed, 2))
    return 0


def _slides_present(directory: str | None = None) -> bool:
    p = _abs(directory or ART_SLIDES)
    return p.is_dir() and any(f.suffix == ".json" for f in p.glob("*.json")
                              if not f.name.startswith("~$"))


def _build_argv(a, *, out: str, check: bool) -> list[str]:
    """只传用户显式给的维度；没给就让 build.py 用它自己的默认值。

    这里**不复制**默认值（研讨型 / 藏青 / 慢）——默认值只该有一处，
    抄一份到壳里，两处迟早漂移。信封里的 args 记录的就是"实际传了什么"。
    """
    argv = [a.slides, "-o", out]
    for flag, val in (("--brief", a.brief), ("--style", a.style), ("--speed", a.speed),
                      ("--card", a.card), ("--source", a.source), ("--chart-mode", a.chart_mode),
                      ("--title", a.title), ("--subtitle", a.subtitle), ("--only", a.only)):
        if val:
            argv += [flag, val]
    if a.notes:
        argv.append("--notes")
    if a.no_notes:
        argv.append("--no-notes")
    if a.no_font_check:
        argv.append("--no-font-check")
    if a.keep_deck:
        argv.append("--keep-deck")
    if check:
        argv.append("--check")
    return argv


def _run_build(a, *, check: bool) -> int:
    stage = "check" if check else "build"
    # 先过门禁再查 slides：空仓库里跑 build，该说的是"你连分析都还没做"，
    # 而不是"去写 slides"——后者在缺大纲时照着做也是白做。
    ok, errors, warns = _gate("build", a.force)
    if not ok:
        return fail(stage, errors[0]["code"], errors[0]["message"],
                    hint=errors[0].get("hint", ""), details=errors)
    if not _slides_present(a.slides) and not a.dry_run:
        return fail(stage, "E_MISSING_INPUT", f"{a.slides}/ 里没有 .json",
                    hint=next_cmd("slides"))
    for w in warns:
        print(f"  ! {w['message']} —— {w.get('hint', '')}", file=sys.stderr)

    out = a.out or (".pi/out/_check.pptx" if check else ART_OUT)
    code, child, elapsed = run_child("build", _build_argv(a, out=out, check=check),
                                     dry=a.dry_run, quiet=a.quiet)
    if a.dry_run:
        emit(True, stage, dry_run=True, warnings=warns)
        return 0

    child = child or {}
    child_ok = child.get("ok", code == 0)
    if code != 0 or not child_ok:
        reason = child.get("reason") or ("runtime_error" if code != 0 else "unknown")
        hint = "按上面报告里的 slide/field 逐条改，改完重跑"
        if reason == "warnings_not_zero":
            # 这里的 warnings 字段是入口层的上游告警；deck 自己的 warning 清单只在
            # 上面那行紧凑报告里。不点明的话，模型会以为"只差几条 untracked 告警"
            # 直接 --force 放行，把真正的内容问题漏掉。
            n = child.get("warnings")
            hint = (f"deck 还有 {n} 条 warning 未归零（清单在上面的报告里，按 slide/field 改）；"
                    "本 RESULT 的 warnings 只列入口层告警，不是 deck 的")
        return fail(stage, _map_reason(reason),
                    f"未通过：{reason}", hint=hint,
                    warnings=warns, details=child.get("errors"), child_reason=reason, ran=True)

    if not check:
        good, art_err = _check_artifact("build", [out], a.force)
        if not good:
            return fail("build", "E_NO_ARTIFACT", art_err[0]["message"], hint=art_err[0]["hint"],
                        ran=True)
        inputs = {"slides": ART_SLIDES}
        if a.source:
            inputs["source"] = a.source
        record("build", inputs, [out], elapsed,
               {k: v for k, v in (("brief", a.brief), ("style", a.style),
                                  ("speed", a.speed), ("out", out)) if v})
        print(f"✓ {out}")
    else:
        print("✓ 仅质检（未渲染）")

    emit(True, stage, artifacts=[out] if not check else [],
         warnings=warns, child=child, elapsed=round(elapsed, 2),
         next=({"stage": "deliver", "cmd": "交付：报路径 + 三个维度 + 每页摘要"}
               if not check else {"stage": "build", "cmd": next_cmd("build")}))
    return 0


def cmd_check(a) -> int:
    return _run_build(a, check=True)


def cmd_build(a) -> int:
    return _run_build(a, check=False)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    # 公共开关：每个子命令都认，避免"这个命令有没有那个参数"的记忆负担
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--json", action="store_true", help="额外打印结构化信封")
    common.add_argument("--quiet", "-q", action="store_true", help="不转发子进程输出")
    common.add_argument("--force", action="store_true", help="忽略上游缺口（写进 RESULT 可追溯）")
    common.add_argument("--dry-run", action="store_true", help="只打印将执行的命令，不跑不写状态")

    ap = argparse.ArgumentParser(
        prog="pptx.py",
        description="PPT Agent 唯一入口：analyze → outline → budget → 写页 → build",
        epilog="不带子命令 = 看流水线状态。成败看末行 RESULT，退出码恒 0。",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("status", parents=[common], help="看流水线状态（默认）")
    p.set_defaults(func=cmd_status)

    p = sub.add_parser("analyze", parents=[common], help="Word 预处理 + 分析表骨架")
    p.add_argument("source", help="源文件（.docx / .doc / .md）")
    p.add_argument("--title", default=None, help="deck 标题（只记进状态，供 outline 用）")
    p.set_defaults(func=cmd_analyze)

    p = sub.add_parser("outline", parents=[common], help="分析表 → 结构化大纲")
    p.add_argument("--title", default=None)
    p.add_argument("--subtitle", default=None)
    p.add_argument("--brief", default="seminar", help="写作方向（默认 seminar）")
    p.set_defaults(func=cmd_outline)

    p = sub.add_parser("budget", parents=[common], help="大纲 → 逐页字数预算表")
    p.set_defaults(func=cmd_budget)

    def add_build_args(p):
        p.add_argument("--slides", default=ART_SLIDES, help=f"slides 目录（默认 {ART_SLIDES}）")
        p.add_argument("-o", "--out", default=None, help=f"输出 pptx（默认 {ART_OUT}）")
        p.add_argument("--brief", default=None)
        p.add_argument("--title", default=None, help="deck 标题（缺省取封面页 title）")
        p.add_argument("--subtitle", default=None)
        p.add_argument("--style", default=None)
        p.add_argument("--speed", default=None)
        p.add_argument("--card", default=None)
        p.add_argument("--source", default=None,
                       help=f"输入分析（默认有 {ART_ANALYSIS} 就自动带上）")
        p.add_argument("--chart-mode", default=None, choices=["mpl", "native"])
        p.add_argument("--only", default=None, help="只渲染指定页，如 3,7")
        p.add_argument("--notes", action="store_true")
        p.add_argument("--no-notes", action="store_true")
        p.add_argument("--no-font-check", action="store_true")
        p.add_argument("--keep-deck", action="store_true")

    p = sub.add_parser("check", parents=[common], help="只质检，不出文件")
    add_build_args(p)
    p.set_defaults(func=cmd_check)

    p = sub.add_parser("build", parents=[common], help="合并 + 质检 + 渲染 + 字体自检")
    add_build_args(p)
    p.set_defaults(func=cmd_build)

    return ap


def main(argv=None) -> int:
    global _CWD0
    _utf8_streams()
    _CWD0 = Path.cwd()
    # 统一工作目录：所有内部脚本都用 `.pi/work/...` 相对路径，cwd 必须是仓库根。
    # 从别处调用（`python /path/to/.pi/pptx.py status`）也能工作；
    # 用户给的相对路径在 cmd_analyze 里按 _CWD0 兜住。
    try:
        os.chdir(REPO)
    except Exception:
        pass

    parser = build_parser()
    # 不带子命令 = status；`python .pi/pptx.py --json` 这种"只给全局开关"也当 status
    raw = list(sys.argv[1:] if argv is None else argv)
    if not raw:
        raw = ["status"]
    elif raw[0].startswith("-") and raw[0] not in ("-h", "--help"):
        raw = ["status"] + raw
    args = parser.parse_args(raw)
    # 自动带 --source：分析表在就带上，省掉"忘了写 --source 导致覆盖度不查"这个坑
    if getattr(args, "cmd", "") in ("build", "check") and not args.source:
        if (REPO / ART_ANALYSIS).exists():
            args.source = ART_ANALYSIS
    try:
        return args.func(args)
    except KeyboardInterrupt:
        print("\n中断", file=sys.stderr)
        emit(False, getattr(args, "cmd", "?"), reason="interrupted")
        return 0
    except Exception as e:  # 未捕获异常才算真失败——非零退出
        import traceback
        traceback.print_exc()
        emit(False, getattr(args, "cmd", "?"), reason="exception", error=str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
