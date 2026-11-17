// @ts-nocheck
/**
 * 依赖自举（deps bootstrap）
 *
 * 时机：pi 加载项目扩展时执行本文件，**在 session_start / 资源发现 / provider 注册之前**。
 *       pi 文档：工厂函数返回 Promise 时，pi 会 await 它再继续启动
 *       （docs/extensions.md「Writing an Extension」）。所以这里做成 async 工厂。
 *
 * 做什么：
 *   1. 定位 CPython 3.12（离线部署的目标解释器；wheels/py312 里全是 cp312 二进制）
 *   2. 探针：在这个解释器里真跑一次 import，确认 requirements.txt 的 7 个包可用
 *   3. 缺失则 `pip install` 从内网镜像 172.30.39.34:8765/simple 拉取
 *   4. 结果写缓存，24 小时内命中就秒过（否则每次启动都要付 1~3 秒的 import 成本）
 *
 * 刻意不做的事：
 *   - 安装失败**不让 pi 崩溃**。起不来的 pi 没法排查问题，改成打醒目警告后放行。
 *   - 不 block 非 3.12 解释器。本机日常跑的 3.13/3.14 不归本脚本管（wheels 也装不上），
 *     找不到 3.12 就安静跳过，同一条提示 24 小时只打一次。
 *
 * 环境变量：
 *   PI_PIP_INDEX_URL   覆盖镜像地址（默认 http://172.30.39.34:8765/simple）
 *   PI_DEPS_PY_VERSION 覆盖目标解释器版本（默认 3.12；换交付机版本时才需要动）
 *   PI_DEPS_SKIP=1     完全跳过（应急）
 *   PI_DEPS_VERBOSE=1  打印判定过程
 *   PI_DEPS_CLI=1      以命令行模式运行（执行检查/安装后退出，不挂到 pi 上）
 */

import { spawnSync } from "node:child_process"
import fs from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"

const PIP_INDEX = process.env.PI_PIP_INDEX_URL || "http://172.30.39.34:8765/simple"
// 离线交付机的解释器版本；wheels/py312 里全是 cp312 二进制
const TARGET_VERSION = process.env.PI_DEPS_PY_VERSION || "3.12"
const MODULES = ["pptx", "lxml", "pydantic", "matplotlib", "numpy", "PIL", "yaml"]
const STATE_FILE = ".pi/work/.deps-state.json"
const CACHE_TTL_MS = 24 * 60 * 60 * 1000
const PIP_TIMEOUT_MS = 15 * 60 * 1000

// 真跑 import 而不是 find_spec：装了但装坏（缺传递依赖）只有真 import 才暴露。
// matplotlib 先切 Agg，避免无头环境探测 backend 拖时间。
const PROBE_CODE = [
  "import matplotlib",
  "matplotlib.use('Agg')",
  "import pptx, lxml, pydantic, numpy, PIL, yaml",
  "print('DEPS_OK')",
].join("\n")

function verbose(msg) {
  if (process.env.PI_DEPS_VERBOSE) process.stderr.write(`[deps] ${msg}\n`)
}

function warn(msg) {
  process.stderr.write(`[deps] ${msg}\n`)
}

/** 从起始目录向上找仓库根（同时有 requirements.txt 和 .pi/ 的那层）。 */
function findRepoRoot(start) {
  let dir = start
  for (let i = 0; i < 6; i += 1) {
    if (
      fs.existsSync(path.join(dir, "requirements.txt")) &&
      fs.existsSync(path.join(dir, ".pi"))
    ) {
      return dir
    }
    const parent = path.dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return null
}

/** 扩展自身位置兜底：<root>/.pi/extensions/xxx.ts → <root> */
function rootFromExtensionFile() {
  try {
    const here = path.dirname(fileURLToPath(import.meta.url))
    return path.resolve(here, "..", "..")
  } catch {
    return null
  }
}

/**
 * 找一个目标版本的解释器。
 * 顺序：py -X.Y（Windows launcher）→ pythonX.Y → python（但版本必须精确匹配）。
 * 返回 { cmd, args, executable } 或 null。
 */
function resolveTargetPython() {
  const candidates = [
    { cmd: "py", args: [`-${TARGET_VERSION}`] },
    { cmd: `python${TARGET_VERSION}`, args: [] },
    { cmd: "python", args: [] },
  ]
  for (const candidate of candidates) {
    const probe = spawnSync(
      candidate.cmd,
      [
        ...candidate.args,
        "-c",
        "import sys;print('%d.%d'%sys.version_info[:2]);print(sys.executable)",
      ],
      { encoding: "utf8", timeout: 20000, windowsHide: true },
    )
    if (probe.status !== 0 || !probe.stdout) continue
    const [version, executable] = probe.stdout.trim().split(/\r?\n/)
    if (version?.trim() !== TARGET_VERSION) continue
    verbose(`解释器: ${candidate.cmd} ${candidate.args.join(" ")} → ${executable}`)
    return { cmd: candidate.cmd, args: candidate.args, executable: executable.trim() }
  }
  return null
}

function readState(repoRoot) {
  try {
    return JSON.parse(fs.readFileSync(path.join(repoRoot, STATE_FILE), "utf8"))
  } catch {
    return null
  }
}

function writeState(repoRoot, state) {
  try {
    const file = path.join(repoRoot, STATE_FILE)
    fs.mkdirSync(path.dirname(file), { recursive: true })
    fs.writeFileSync(file, `${JSON.stringify(state, null, 2)}\n`, "utf8")
  } catch (err) {
    verbose(`缓存写入失败（不影响启动）: ${err.message}`)
  }
}

function requirementsMtime(repoRoot) {
  try {
    return fs.statSync(path.join(repoRoot, "requirements.txt")).mtimeMs
  } catch {
    return 0
  }
}

/** 在指定解释器里跑 import 探针。 */
function probeDeps(python) {
  const result = spawnSync(
    python.cmd,
    [...python.args, "-c", PROBE_CODE],
    { encoding: "utf8", timeout: 120000, windowsHide: true },
  )
  if (result.status === 0 && (result.stdout || "").includes("DEPS_OK")) {
    return { ok: true, detail: "" }
  }
  // 只保留最后几行，pip 与 Python 的报错价值在尾部
  const detail = `${result.stderr || result.stdout || ""}`.trim().split(/\r?\n/).slice(-4).join("\n")
  return { ok: false, detail }
}

function installDeps(repoRoot, python) {
  const host = (() => {
    try {
      return new URL(PIP_INDEX).host
    } catch {
      return ""
    }
  })()

  const args = [
    ...python.args,
    "-m",
    "pip",
    "install",
    "--disable-pip-version-check",
    "--no-input",
    "--timeout",
    "30",
    "--retries",
    "2",
    "--index-url",
    PIP_INDEX,
  ]
  if (host) args.push("--trusted-host", host.split(":")[0])
  args.push("-r", path.join(repoRoot, "requirements.txt"))

  warn(`依赖缺失，从 ${PIP_INDEX} 安装（首次可能要几分钟）...`)
  const result = spawnSync(python.cmd, args, {
    cwd: repoRoot,
    stdio: ["ignore", "inherit", "inherit"], // 让 pip 的进度条直接可见
    timeout: PIP_TIMEOUT_MS,
    windowsHide: true,
  })
  if (result.error?.code === "ETIMEDOUT") {
    return { ok: false, detail: `pip install 超过 ${PIP_TIMEOUT_MS / 60000} 分钟未完成` }
  }
  if (result.status !== 0) {
    return { ok: false, detail: `pip install 退出码 ${result.status}` }
  }
  return { ok: true, detail: "" }
}

/**
 * 主流程。返回 { status, message }：
 *   ok            依赖已在位
 *   installed     本次装上了
 *   no-python     没有 3.12 解释器，跳过
 *   failed        安装或验证失败
 */
function ensureDeps(repoRoot) {
  const reqMtime = requirementsMtime(repoRoot)
  const cached = readState(repoRoot)
  const now = Date.now()

  const python = resolveTargetPython()
  if (!python) {
    // 本机日常是 3.13/3.14，这个分支会命中；同一提示 24h 只说一次。
    if (cached?.status !== "no-python" || now - (cached.checkedAt || 0) > CACHE_TTL_MS) {
      warn(`未找到 CPython ${TARGET_VERSION} 解释器，跳过依赖检查（wheels 只服务 ${TARGET_VERSION} 离线机）`)
    }
    writeState(repoRoot, { status: "no-python", checkedAt: now, requirementsMtimeMs: reqMtime })
    return { status: "no-python", message: "" }
  }

  // 缓存命中：同一解释器 + requirements.txt 没改过 + 24h 内验证过 → 不再付 import 成本
  const cacheHit =
    cached?.status === "ok" &&
    cached.python === python.executable &&
    cached.requirementsMtimeMs === reqMtime &&
    now - (cached.checkedAt || 0) < CACHE_TTL_MS
  if (cacheHit) {
    verbose("缓存命中，跳过探针")
    return { status: "ok", message: "" }
  }

  verbose("探针：检查 7 个直接依赖")
  const first = probeDeps(python)
  if (first.ok) {
    writeState(repoRoot, {
      status: "ok",
      python: python.executable,
      version: TARGET_VERSION,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
    })
    return { status: "ok", message: "" }
  }

  verbose(`探针失败: ${first.detail}`)
  const install = installDeps(repoRoot, python)
  if (!install.ok) {
    warn(`依赖安装失败：${install.detail}`)
    warn(`手动修复：${python.executable} -m pip install -r requirements.txt`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: TARGET_VERSION,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
      detail: install.detail,
    })
    return { status: "failed", message: install.detail }
  }

  const second = probeDeps(python)
  if (!second.ok) {
    warn(`安装完成但验证仍失败：${second.detail}`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: TARGET_VERSION,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
      detail: second.detail,
    })
    return { status: "failed", message: second.detail }
  }

  writeState(repoRoot, {
    status: "ok",
    python: python.executable,
    version: TARGET_VERSION,
    requirementsMtimeMs: reqMtime,
    checkedAt: now,
  })
  warn(`依赖安装完成（${MODULES.join(", ")}）`)
  return { status: "installed", message: "" }
}

function bootstrap() {
  if (process.env.PI_DEPS_SKIP === "1") {
    verbose("PI_DEPS_SKIP=1，跳过")
    return { status: "skipped", message: "" }
  }
  const repoRoot = findRepoRoot(process.cwd()) || rootFromExtensionFile()
  if (!repoRoot || !fs.existsSync(path.join(repoRoot, "requirements.txt"))) {
    verbose("不在 PPT-Agent 仓库内，跳过")
    return { status: "skipped", message: "" }
  }
  return ensureDeps(repoRoot)
}

// ---------------------------------------------------------------------------
// pi 扩展入口：async 工厂 → pi 会 await 它再继续启动
// ---------------------------------------------------------------------------
export default async function depsBootstrap(_pi) {
  try {
    bootstrap()
  } catch (err) {
    // 自举本身出错绝不能拦住 pi 启动
    warn(`自举异常（已忽略）：${err?.message || err}`)
  }
}

// ---------------------------------------------------------------------------
// 命令行模式：PI_DEPS_CLI=1 npx jiti .pi/extensions/deps-bootstrap.ts
// 或 node --experimental-strip-types 直接跑，用于人工排查与 setup.bat 调用
// ---------------------------------------------------------------------------
if (process.env.PI_DEPS_CLI === "1") {
  const result = bootstrap()
  if (result.status === "no-python") {
    process.stderr.write(`[deps] 没有 ${TARGET_VERSION} 解释器，未做任何检查\n`)
    process.exit(2)
  }
  if (result.status === "failed") process.exit(1)
  process.stderr.write(`[deps] ${result.status}\n`)
  process.exit(0)
}
