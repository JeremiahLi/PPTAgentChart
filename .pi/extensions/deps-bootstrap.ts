// @ts-nocheck
/**
 * 依赖自举（deps bootstrap）—— 内网镜像单源版
 *
 * 时机：pi 加载扩展时执行本文件，**在 session_start / 资源发现 / provider 注册之前**。
 *       pi 文档：工厂函数返回 Promise 时，pi 会 await 它再继续启动
 *       （docs/extensions.md「Writing an Extension」）。所以这里做成 async 工厂。
 *
 * 唯一安装来源：**内网 server**（`.pi/deps.json` 的 `indexUrl`）。
 *   - 不回退公网 PyPI：离线机连不上，回退只会白等；配错了也不该被兜底掩盖。
 *   - 不使用 `.pi/wheels` 本地目录：依赖全部由内网 server 供应，本地包是另一条要维护的路径。
 *   - 没配 `indexUrl` 就**直接报错**，不猜地址、不偷偷走公网。
 *   - 开发机到不了内网 IP，所以本脚本不做任何"先探测一下能不能连"的试探，直接按配置下。
 *
 * 判据：不猜 import 名、不自己比版本，直接让 pip 判断——
 *       `pip install --dry-run --no-index -r requirements.txt`（纯本地比对，不联网）
 *       退出码 0 = 全部满足（含版本约束）；非 0 = 缺包或版本不符。
 *       这样 `python-pptx→pptx`、`Pillow→PIL`、`PyYAML→yaml` 这类映射一概不用维护。
 *
 * 终端反馈（每次启动都有，不让用户猜）：
 *   检查        `检查依赖是否就位（本地比对，不联网，约 10 秒）...`
 *   已就位      `✓ 依赖已在位（7 个直接依赖全部满足 · CPython 3.12.10，无需下载）`
 *   缓存命中    `✓ 依赖已在位（7 个直接依赖，3.2 小时前已验证；本次跳过重复检查）`
 *   下载中      `依赖未就位（缺 3/7：...），开始从内网 server 下载：<url>` + pip 逐块实时输出
 *   完成        `✓ 依赖安装完成（耗时 42s，复检通过 · CPython 3.12.10）` + `本次装入：a-1.0 b-2.0 …`
 *   失败        `✗ 依赖安装失败：<人话原因>` + 手动重试命令
 *
 * 配置：`<仓库根>/.pi/deps.json`（全部字段可选）
 *   {
 *     "python": "3.12",                              // 限定解释器版本；不填=用 PATH 上的 python
 *     "indexUrl": "http://172.30.39.34:8765/simple",  // **内网 server，装包唯一来源**
 *     "timeoutMinutes": 15                            // 安装超时，默认 15
 *   }
 *
 * 环境变量（优先级高于 deps.json）：
 *   PI_DEPS_PY_VERSION   同 python
 *   PI_PIP_INDEX_URL     同 indexUrl
 *   PI_DEPS_TIMEOUT_MIN  同 timeoutMinutes
 *   PI_DEPS_SKIP=1       完全跳过（应急）
 *   PI_DEPS_FORCE=1      忽略 24h 缓存，强制真跑一次检查
 *   PI_DEPS_VERBOSE=1    打印判定过程
 *   PI_DEPS_CLI=1        命令行模式（检查/安装后退出，不挂到 pi 上）
 *
 * 缓存为什么留：`pip install --dry-run --no-index` 实测 11 秒，每次启动都跑等于每次多等 11 秒；
 * 但缓存命中时提示里必须写清"X 前已验证、本次跳过"，不能让人误以为刚查过。
 * 要立刻真查一次：`PI_DEPS_FORCE=1 pi`。
 *
 * 刻意不做的事：
 *   - 安装失败**不让 pi 崩溃**。起不来的 pi 没法排查问题，改成打醒目警告 + 人话原因后放行。
 *   - 不 block 版本不符的解释器。配了 `python` 却找不到该版本时跳过并提示。
 */

import { spawn, spawnSync } from "node:child_process"
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
 * 流式执行：边跑边把子进程输出转发到 stderr，同时累积全文供事后解析。
 * 为什么不用 spawnSync + stdio:inherit——那样拿不到任何文本，"失败原因"就只能报退出码。
 * 用 pipe 手动转发，实时性和可解析性同时拿到。
 */
function runStreaming(cmd, args, { cwd, timeoutMs }) {
  return new Promise((resolve) => {
    let child
    try {
      child = spawn(cmd, args, { cwd, windowsHide: true })
    } catch (err) {
      resolve({ code: -1, output: "", error: err, timedOut: false })
      return
    }

    let output = ""
    let timedOut = false
    let settled = false
    const timer = setTimeout(() => {
      timedOut = true
      try {
        child.kill("SIGKILL")
      } catch {
        /* 已经退出 */
      }
    }, timeoutMs)
    const finish = (payload) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(payload)
    }

    const forward = (chunk) => {
      const text = chunk.toString("utf8")
      output += text
      process.stderr.write(text)
    }
    child.stdout?.on("data", forward)
    child.stderr?.on("data", forward)
    child.on("error", (err) => finish({ code: -1, output, error: err, timedOut }))
    child.on("close", (code) => finish({ code, output, error: null, timedOut }))
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

/** 直接依赖条数（去掉空行与注释），只用于提示语里报个数。 */
function countRequirements(repoRoot) {
  try {
    return fs
      .readFileSync(path.join(repoRoot, "requirements.txt"), "utf8")
      .split(/\r?\n/)
      .filter((line) => line.trim() && !line.trim().startsWith("#")).length
  } catch {
    return 0
  }
}

/** 把毫秒差写成"刚刚 / 23 分钟 / 3.2 小时 / 2 天"，给缓存提示用。 */
function formatAge(ms) {
  const minutes = Math.max(0, Math.round(ms / 60000))
  if (minutes < 1) return "刚刚"
  if (minutes < 60) return `${minutes} 分钟`
  const hours = Math.round((minutes / 60) * 10) / 10
  if (hours < 48) return `${hours} 小时`
  return `${Math.round(hours / 24)} 天`
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

/**
 * 把 pip 的原始报错翻成一句能照着修的人话。
 * 顺序有意义：先判"根本没连上"，再判"连上了但没这个包"——两者 pip 都可能吐
 * "No matching distribution found"，先判网络才不会把断网误报成缺包。
 */
function explainFailure(output, indexUrl) {
  const text = String(output || "")
  let host = indexUrl || "未配置"
  try {
    host = new URL(indexUrl).host || indexUrl
  } catch {
    /* 保留原串 */
  }

  if (/WinError 10061|Connection refused|NewConnectionError|Failed to establish a new connection/i.test(text)) {
    return `内网 server 拒绝连接（${host}）：这台机器到该 IP/端口不通，或 server 服务未启动`
  }
  if (/WinError 10060|WinError 10065|timed out|TimeoutError|ReadTimeoutError/i.test(text)) {
    return `连接内网 server 超时（${host}）：网络不通、防火墙拦截，或服务响应过慢`
  }
  if (/WinError 11001|WinError 11002|Name or service not known|getaddrinfo failed|nodename nor servname/i.test(text)) {
    return `内网 server 主机名解析失败（${host}）：indexUrl 里应写 IP，别写主机名`
  }
  if (/No matching distribution found for ([^\s(]+)/i.test(text)) {
    const m = text.match(/No matching distribution found for ([^\s(]+)/i)
    return `内网 server 上没有 ${m ? m[1] : "所需版本"}：该包未同步到镜像，或镜像里的版本与 requirements.txt 不一致`
  }
  if (/SSLError|certificate verify failed/i.test(text)) {
    return "证书校验失败：http 源已自动加 --trusted-host；https 自签名证书需要把 CA 装进系统信任区"
  }
  if (/WinError 5|Access is denied|Permission denied|Errno 13/i.test(text)) {
    return "没有写入权限：Python 装在受保护目录，用管理员身份重开终端再启动 pi"
  }
  if (/No space left|WinError 112/i.test(text)) {
    return "磁盘空间不足，清理后重试"
  }

  const tail = text
    .trim()
    .split(/\r?\n/)
    .filter((line) => line.trim())
    .slice(-3)
    .join(" | ")
  return tail ? `pip 报错（未归类）：${tail}` : "pip 没有输出可读原因（可能被强杀或超时）"
}

/** 从 pip 输出里捞 "Successfully installed a-1.0 b-2.0" 那一行，供完成后呈现。 */
function summarizeInstalled(output) {
  const m = String(output || "").match(/Successfully installed ([^\r\n]+)/i)
  if (!m) return null
  const pkgs = m[1].trim().split(/\s+/).filter(Boolean)
  if (!pkgs.length) return null
  if (pkgs.length <= 8) return pkgs.join("  ")
  return `${pkgs.slice(0, 8).join("  ")}  …等 ${pkgs.length} 个`
}

/**
 * 从内网 server 安装。强制单源：没配 indexUrl 直接失败，不偷偷走公网。
 * --prefer-binary：离线机通常没有编译链，优先拿 wheel，别现场编译 sdist。
 * --progress-bar on：输出不是 tty 时 pip 默认关进度条，强制打开才能看到刷新。
 */
async function installDeps(repoRoot, python, config, missing, reqCount) {
  const startedAt = Date.now()
  const elapsed = () => Math.round((Date.now() - startedAt) / 1000)
  const lacks = missing.length
    ? `缺 ${missing.length}/${reqCount || "?"}：${missing.slice(0, 6).join(", ")}`
    : "版本不满足"

  if (!config.indexUrl) {
    return {
      ok: false,
      elapsed: 0,
      output: "",
      installed: null,
      detail: "未配置内网 server 地址，拒绝走公网 PyPI",
      hint: `在 ${CONFIG_FILE} 里写 {"indexUrl": "http://<内网IP>:<端口>/simple"}`,
    }
  }

  const args = [
    ...python.args,
    "-m",
    "pip",
    "install",
    "--disable-pip-version-check",
    "--no-input",
    "--prefer-binary",
    "--progress-bar",
    "on",
    "--timeout",
    "30",
    "--retries",
    "2",
    "--index-url",
    config.indexUrl,
  ]
  try {
    const host = new URL(config.indexUrl).hostname
    if (host) args.push("--trusted-host", host)
  } catch {
    /* 非法 URL 交给 pip 自己报错 */
  }
  args.push("-r", path.join(repoRoot, "requirements.txt"))

  warn(`依赖未就位（${lacks}），开始从内网 server 下载：${config.indexUrl}`)
  warn(`pip 实时输出如下（下载进度即时刷新）...`)
  const result = await runStreaming(python.cmd, args, {
    cwd: repoRoot,
    timeoutMs: config.timeoutMinutes * 60 * 1000,
  })

  if (result.timedOut) {
    return {
      ok: false,
      elapsed: elapsed(),
      output: result.output,
      installed: null,
      detail: `下载超过 ${config.timeoutMinutes} 分钟未完成（内网 server 太慢或卡住）`,
      hint: "加大超时：设 PI_DEPS_TIMEOUT_MIN=30 后重启 pi",
    }
  }
  if (result.error) {
    return {
      ok: false,
      elapsed: elapsed(),
      output: result.output,
      installed: null,
      detail: `pip 无法启动：${result.error.message}`,
      hint: `手动重试：${python.executable} -m pip install -r requirements.txt`,
    }
  }
  if (result.code !== 0) {
    return {
      ok: false,
      elapsed: elapsed(),
      output: result.output,
      installed: null,
      detail: explainFailure(result.output, config.indexUrl),
      hint: `手动重试：${python.executable} -m pip install --index-url ${config.indexUrl} -r requirements.txt`,
    }
  }

  return {
    ok: true,
    elapsed: elapsed(),
    output: result.output,
    installed: summarizeInstalled(result.output),
    detail: "",
    hint: "",
  }
}

/**
 * 主流程。返回 { status, missing }：
 *   ok        依赖已在位
 *   installed 本次装上了
 *   no-python 没有匹配的解释器，跳过
 *   no-pip    pip 太老，无法检查，跳过
 *   failed    安装或验证失败
 */
async function ensureDeps(repoRoot) {
  const config = readConfig(repoRoot)
  const reqMtime = requirementsMtime(repoRoot)
  const reqCount = countRequirements(repoRoot)
  const cached = readState(repoRoot)
  const now = Date.now()
  const force = process.env.PI_DEPS_FORCE === "1"

  const python = resolvePython(config)
  if (!python) {
    // 解释器缺失是硬阻塞（依赖一个都不会装），不吃缓存、每次启动都说
    const hint = config.python ? ` CPython ${config.python}` : ""
    warn(`! 未找到${hint} 解释器，本次跳过依赖检查，依赖不会被安装`)
    warn(`  py -0 查看本机已装版本；装好后重开 pi 即生效（PI_DEPS_SKIP=1 可静音）`)
    writeState(repoRoot, { status: "no-python", checkedAt: now, requirementsMtimeMs: reqMtime })
    return { status: "no-python", missing: [] }
  }

  // 缓存命中：同一解释器 + requirements.txt 没改过 + 24h 内验证过。
  // 提示里必须写清"多久前验的、本次没重跑"，否则会被误读成"刚检查过"。
  const cacheHit =
    !force &&
    cached?.status === "ok" &&
    cached.python === python.executable &&
    cached.requirementsMtimeMs === reqMtime &&
    now - (cached.checkedAt || 0) < CACHE_TTL_MS
  if (cacheHit) {
    verbose("缓存命中，跳过检查")
    const age = formatAge(now - (cached.checkedAt || 0))
    const when = age === "刚刚" ? "刚刚已" : `${age}前已`
    warn(`✓ 依赖已在位（${reqCount} 个直接依赖，${when}验证；本次跳过重复检查）`)
    return { status: "ok", missing: [] }
  }
  if (force) verbose("PI_DEPS_FORCE=1，忽略缓存重查")

  if (!pipSupportsDryRun(python)) {
    warn("! 目标环境的 pip 不支持 --dry-run（需 >= 22.2），无法检查依赖，建议先升级 pip")
    writeState(repoRoot, { status: "no-pip", checkedAt: now, requirementsMtimeMs: reqMtime })
    return { status: "no-pip", missing: [] }
  }

  warn("检查依赖是否就位（本地比对，不联网，约 10 秒）...")
  const first = checkDeps(repoRoot, python)
  if (first.ok) {
    writeState(repoRoot, {
      status: "ok",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
    })
    warn(`✓ 依赖已在位（${reqCount} 个直接依赖全部满足 · CPython ${python.version}，无需下载）`)
    return { status: "ok", missing: [] }
  }

  verbose(`未满足: ${first.missing.join(", ") || first.detail}`)
  const install = await installDeps(repoRoot, python, config, first.missing, reqCount)
  if (!install.ok) {
    warn(`✗ 依赖安装失败：${install.detail}`)
    if (install.hint) warn(`  ${install.hint}`)
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
    warn(`✗ 下载完成但复检仍不满足：${second.missing.join(", ") || second.detail}`)
    warn("  说明内网 server 上的包版本与 requirements.txt 不一致，或传递依赖被裁掉了")
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
  warn(`✓ 依赖安装完成（耗时 ${install.elapsed}s，复检通过 · CPython ${python.version}）`)
  if (install.installed) warn(`  本次装入：${install.installed}`)
  return { status: "installed", missing: [] }
}

async function bootstrap() {
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
    await bootstrap()
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
  bootstrap().then((result) => {
    if (result.status === "failed") {
      if (result.missing.length) warn(`缺失/不满足: ${result.missing.join(", ")}`)
      process.exit(1)
    }
    if (result.status === "no-python" || result.status === "no-pip") process.exit(2)
    process.stderr.write(`[deps] ${result.status}\n`)
    process.exit(0)
  })
}
