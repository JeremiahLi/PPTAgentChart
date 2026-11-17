// @ts-nocheck
/**
 * 依赖自举（deps bootstrap）—— 通用版
 *
 * 时机：pi 加载扩展时执行本文件，**在 session_start / 资源发现 / provider 注册之前**。
 *       pi 文档：工厂函数返回 Promise 时，pi 会 await 它再继续启动
 *       （docs/extensions.md「Writing an Extension」）。所以这里做成 async 工厂。
 *
 * 判据：不猜 import 名、不自己比版本，直接让 pip 判断——
 *       `pip install --dry-run --no-index -r requirements.txt`
 *       退出码 0 = 全部满足（含版本约束）；非 0 = 缺包或版本不符。
 *       这样 `python-pptx→pptx`、`Pillow→PIL`、`PyYAML→yaml` 这类映射一概不用维护，
 *       `numpy==2.4.6` 的版本约束也由 pip 权威实现负责。
 *
 * 复用：把本文件放进任意 pi 项目的 `.pi/extensions/`（或 `~/.pi/agent/extensions/`
 *       对所有项目生效），该目录只要有 requirements.txt 即可工作，零配置默认走公网 PyPI。
 *
 * 配置：`<仓库根>/.pi/deps.json`（可选，全部字段可选）
 *   {
 *     "python": "3.12",                              // 限定解释器版本；不填=用 PATH 上的 python
 *     "indexUrl": "http://172.30.39.34:8765/simple",  // 私有源；不填=官方 PyPI
 *     "timeoutMinutes": 15                            // 安装超时，默认 15
 *   }
 *
 * 环境变量（优先级高于 deps.json）：
 *   PI_DEPS_PY_VERSION   同 python
 *   PI_PIP_INDEX_URL     同 indexUrl
 *   PI_DEPS_TIMEOUT_MIN  同 timeoutMinutes
 *   PI_DEPS_SKIP=1       完全跳过（应急）
 *   PI_DEPS_VERBOSE=1    打印判定过程
 *   PI_DEPS_CLI=1        命令行模式（检查/安装后退出，不挂到 pi 上）
 *
 * 刻意不做的事：
 *   - 安装失败**不让 pi 崩溃**。起不来的 pi 没法排查问题，改成打醒目警告后放行。
 *   - 不 block 版本不符的解释器。配了 `python` 却找不到该版本时安静跳过。
 */

import { spawnSync } from "node:child_process"
import fs from "node:fs"
import path from "node:path"

const CONFIG_FILE = ".pi/deps.json"
const STATE_FILE = ".pi/work/.deps-state.json"
const CACHE_TTL_MS = 24 * 60 * 60 * 1000
const CHECK_TIMEOUT_MS = 180 * 1000

function verbose(msg) {
  if (process.env.PI_DEPS_VERBOSE) process.stderr.write(`[deps] ${msg}\n`)
}

function warn(msg) {
  process.stderr.write(`[deps] ${msg}\n`)
}

/**
 * 从起始目录向上找仓库根：有 requirements.txt 的那一层。
 * 只上溯 3 层——再往上容易跨进父目录的另一个无关项目，把依赖装错地方。
 * 刻意不用「扩展文件自身位置」兜底：那会让任何无关项目都去检查本仓库的
 * requirements.txt，并在那里执行 pip install。
 */
function findRepoRoot(start) {
  let dir = start
  for (let i = 0; i < 3; i += 1) {
    if (fs.existsSync(path.join(dir, "requirements.txt"))) return dir
    const parent = path.dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return null
}

/** 读配置：内置默认 ← .pi/deps.json ← 环境变量（后者覆盖前者）。 */
function readConfig(repoRoot) {
  const config = { python: null, indexUrl: null, timeoutMinutes: 15 }

  try {
    const raw = JSON.parse(fs.readFileSync(path.join(repoRoot, CONFIG_FILE), "utf8"))
    if (typeof raw.python === "string" && raw.python.trim()) config.python = raw.python.trim()
    if (typeof raw.indexUrl === "string" && raw.indexUrl.trim()) config.indexUrl = raw.indexUrl.trim()
    if (Number.isFinite(raw.timeoutMinutes) && raw.timeoutMinutes > 0) {
      config.timeoutMinutes = raw.timeoutMinutes
    }
    verbose(`配置: ${JSON.stringify(config)}（来自 ${CONFIG_FILE}）`)
  } catch {
    /* 没有配置文件是正常情况 */
  }

  if (process.env.PI_DEPS_PY_VERSION) config.python = process.env.PI_DEPS_PY_VERSION.trim()
  if (process.env.PI_PIP_INDEX_URL) config.indexUrl = process.env.PI_PIP_INDEX_URL.trim()
  const envTimeout = Number(process.env.PI_DEPS_TIMEOUT_MIN)
  if (Number.isFinite(envTimeout) && envTimeout > 0) config.timeoutMinutes = envTimeout

  return config
}

function run(cmd, args, opts = {}) {
  return spawnSync(cmd, args, {
    encoding: "utf8",
    timeout: opts.timeout || 30000,
    windowsHide: true,
    ...opts,
  })
}

/**
 * 定位解释器。
 * 配了版本 → 只接受精确匹配；没配 → 第一个能跑起来的。
 * 返回 { cmd, args, executable, version } 或 null。
 */
function resolvePython(config) {
  const probes = config.python
    ? [
        { cmd: "py", args: [`-${config.python}`] },
        { cmd: `python${config.python}`, args: [] },
        { cmd: "python", args: [] },
      ]
    : [
        { cmd: "python", args: [] },
        { cmd: "python3", args: [] },
        { cmd: "py", args: ["-3"] },
      ]

  for (const probe of probes) {
    const result = run(
      probe.cmd,
      [...probe.args, "-c", "import sys;print('%d.%d'%sys.version_info[:2]);print(sys.executable)"],
      { timeout: 20000 },
    )
    if (result.status !== 0 || !result.stdout) continue
    const [version, executable] = result.stdout.trim().split(/\r?\n/)
    const actual = version?.trim()
    if (!actual || !executable) continue
    if (config.python && actual !== config.python) {
      verbose(`${probe.cmd} ${probe.args.join(" ")} 是 ${actual}，跳过`)
      continue
    }
    verbose(`解释器: ${probe.cmd} ${probe.args.join(" ")} → ${executable} (${actual})`)
    return { cmd: probe.cmd, args: probe.args, executable: executable.trim(), version: actual }
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

/** pip 是否支持 --dry-run（>= 22.2）。老版本无法做无副作用检查。 */
function pipSupportsDryRun(python) {
  const result = run(python.cmd, [...python.args, "-m", "pip", "install", "--help"], {
    timeout: 60000,
  })
  return `${result.stdout || ""}${result.stderr || ""}`.includes("--dry-run")
}

/**
 * 用 pip 自己做满足性检查：--no-index 保证完全不联网，只比对已装环境。
 * 返回 { ok, missing: string[], detail }。
 */
function checkDeps(repoRoot, python) {
  const result = run(
    python.cmd,
    [
      ...python.args,
      "-m",
      "pip",
      "install",
      "--dry-run",
      "--no-index",
      "--disable-pip-version-check",
      "--no-input",
      "-r",
      path.join(repoRoot, "requirements.txt"),
    ],
    { timeout: CHECK_TIMEOUT_MS, cwd: repoRoot },
  )

  if (result.status === 0) return { ok: true, missing: [], detail: "" }

  const output = `${result.stderr || ""}\n${result.stdout || ""}`
  const missing = [...output.matchAll(/No matching distribution found for (\S+)/g)].map((m) =>
    m[1].replace(/[;,.]+$/, ""),
  )
  // 只留最后几行：pip 的报错价值在尾部
  const detail = output.trim().split(/\r?\n/).slice(-4).join("\n")
  return { ok: false, missing: [...new Set(missing)], detail }
}

function installDeps(repoRoot, python, config) {
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
  ]

  if (config.indexUrl) {
    args.push("--index-url", config.indexUrl)
    try {
      const host = new URL(config.indexUrl).hostname
      if (host) args.push("--trusted-host", host)
    } catch {
      /* 非法 URL 交给 pip 自己报错 */
    }
  }
  args.push("-r", path.join(repoRoot, "requirements.txt"))

  const source = config.indexUrl || "默认索引 (PyPI)"
  warn(`依赖未就位，从 ${source} 安装（首次可能要几分钟）...`)
  const result = spawnSync(python.cmd, args, {
    cwd: repoRoot,
    stdio: ["ignore", "inherit", "inherit"],
    timeout: config.timeoutMinutes * 60 * 1000,
    windowsHide: true,
  })

  if (result.error?.code === "ETIMEDOUT") {
    return { ok: false, detail: `pip install 超过 ${config.timeoutMinutes} 分钟未完成` }
  }
  if (result.status !== 0) {
    return { ok: false, detail: `pip install 退出码 ${result.status}` }
  }
  return { ok: true, detail: "" }
}

/**
 * 主流程。返回 { status, missing }：
 *   ok        依赖已在位
 *   installed 本次装上了
 *   no-python 没有匹配的解释器，跳过
 *   no-pip    pip 太老，无法检查，跳过
 *   failed    安装或验证失败
 */
function ensureDeps(repoRoot) {
  const config = readConfig(repoRoot)
  const reqMtime = requirementsMtime(repoRoot)
  const cached = readState(repoRoot)
  const now = Date.now()

  const python = resolvePython(config)
  if (!python) {
    const hint = config.python ? ` CPython ${config.python}` : ""
    if (cached?.status !== "no-python" || now - (cached.checkedAt || 0) > CACHE_TTL_MS) {
      warn(`未找到${hint} 解释器，跳过依赖检查`)
    }
    writeState(repoRoot, { status: "no-python", checkedAt: now, requirementsMtimeMs: reqMtime })
    return { status: "no-python", missing: [] }
  }

  // 缓存命中：同一解释器 + requirements.txt 没改过 + 24h 内验证过
  const cacheHit =
    cached?.status === "ok" &&
    cached.python === python.executable &&
    cached.requirementsMtimeMs === reqMtime &&
    now - (cached.checkedAt || 0) < CACHE_TTL_MS
  if (cacheHit) {
    verbose("缓存命中，跳过检查")
    return { status: "ok", missing: [] }
  }

  if (!pipSupportsDryRun(python)) {
    warn("目标环境的 pip 不支持 --dry-run（需 >= 22.2），无法检查依赖，建议先升级 pip")
    writeState(repoRoot, { status: "no-pip", checkedAt: now, requirementsMtimeMs: reqMtime })
    return { status: "no-pip", missing: [] }
  }

  verbose(`检查 ${path.join(repoRoot, "requirements.txt")}`)
  const first = checkDeps(repoRoot, python)
  if (first.ok) {
    writeState(repoRoot, {
      status: "ok",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
    })
    return { status: "ok", missing: [] }
  }

  verbose(`未满足: ${first.missing.join(", ") || first.detail}`)
  const install = installDeps(repoRoot, python, config)
  if (!install.ok) {
    warn(`依赖安装失败：${install.detail}`)
    warn(`手动修复：${python.executable} -m pip install -r requirements.txt`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
      detail: install.detail,
    })
    return { status: "failed", missing: first.missing }
  }

  const second = checkDeps(repoRoot, python)
  if (!second.ok) {
    warn(`安装完成但复检仍不满足：${second.missing.join(", ") || second.detail}`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
      detail: second.detail,
    })
    return { status: "failed", missing: second.missing }
  }

  writeState(repoRoot, {
    status: "ok",
    python: python.executable,
    version: python.version,
    requirementsMtimeMs: reqMtime,
    checkedAt: now,
  })
  warn(`依赖安装完成（${python.executable}）`)
  return { status: "installed", missing: [] }
}

function bootstrap() {
  if (process.env.PI_DEPS_SKIP === "1") {
    verbose("PI_DEPS_SKIP=1，跳过")
    return { status: "skipped", missing: [] }
  }
  const repoRoot = findRepoRoot(process.cwd())
  if (!repoRoot || !fs.existsSync(path.join(repoRoot, "requirements.txt"))) {
    verbose("当前目录及上 3 层没有 requirements.txt，跳过")
    return { status: "skipped", missing: [] }
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
// 用于人工排查与 setup.bat 调用
// ---------------------------------------------------------------------------
if (process.env.PI_DEPS_CLI === "1") {
  const result = bootstrap()
  if (result.status === "failed") {
    if (result.missing.length) warn(`缺失/不满足: ${result.missing.join(", ")}`)
    process.exit(1)
  }
  if (result.status === "no-python" || result.status === "no-pip") process.exit(2)
  process.stderr.write(`[deps] ${result.status}\n`)
  process.exit(0)
}
