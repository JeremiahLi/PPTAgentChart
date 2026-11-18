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
 *   PI_DEPS_VERIFY=1     预检模式：只向 server 查询能否取全，不安装、不写状态，退出码 0/1
 *   PI_DEPS_VERBOSE=1    打印判定过程
 *   PI_DEPS_CLI=1        命令行模式（检查/安装后退出，不挂到 pi 上）
 *
 * 两套输出通道，两条都要留着：
 *   1) TUI widget（首页可见）—— `session_start` + `ctx.ui.setWidget("deps-status", ..., belowEditor)`。
 *      **这个是给用户看的**。pi 进 interactive 模式会切 alternate screen 重绘，stderr 全被盖掉；
 *      2026-11-18 就是因此首页一片干净，用户压根不知道依赖检没检查。
 *   2) stderr `[deps] ...` —— CLI 模式、print 模式、日志重定向、人工排查用。
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

// ---------------------------------------------------------------------------
// TUI 状态通道
//
// 为什么必须走 ctx.ui.setWidget 而不是 process.stderr：pi 进入 interactive 模式后
// 会切到 alternate screen 并重绘，启动期间写到 stderr 的 `[deps] ...` 全部被盖掉——
// 用户首页看到的就是一片干净，压根不知道依赖检没检查（2026-11-18 实地确认）。
// 所以把同一份状态同时维护成 snapshot，由 session_start 里的 widget 渲染出来。
// stderr 那条照样保留：CLI 模式、print 模式、日志重定向都还在用。
// ---------------------------------------------------------------------------
const snapshot = {
  phase: "init", // init | checking | installing | ok | failed | no-python | no-pip | skipped
  reqCount: 0,
  missing: [],
  version: null,
  wantVersion: null,
  detail: "",
  hint: "",
  note: "",
  progress: "",
}

let redraw = null

function update(patch) {
  Object.assign(snapshot, patch)
  if (!redraw) return
  try {
    redraw()
  } catch {
    /* TUI 已经关掉，忽略 */
  }
}

/** 按显示宽度截断，避免把 widget 撑成不可读的长行。 */
function clip(text, width) {
  const s = String(text || "")
  let w = 0
  let out = ""
  for (const ch of s) {
    const cp = ch.codePointAt(0)
    const wide =
      (cp >= 0x1100 && cp <= 0x115f) ||
      (cp >= 0x2e80 && cp <= 0x303e) ||
      (cp >= 0x3041 && cp <= 0x33ff) ||
      (cp >= 0x3400 && cp <= 0x4dbf) ||
      (cp >= 0x4e00 && cp <= 0x9fff) ||
      (cp >= 0xa000 && cp <= 0xa4cf) ||
      (cp >= 0xac00 && cp <= 0xd7a3) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xfe30 && cp <= 0xfe4f) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      (cp >= 0xffe0 && cp <= 0xffe6) ||
      (cp >= 0x20000 && cp <= 0x3fffd)
    w += wide ? 2 : 1
    if (w > width) return `${out}…`
    out += ch
  }
  return out
}

/** 首页 widget 的内容：1-2 行，任何阶段都有一句能看懂的话。 */
function renderStatus() {
  const s = snapshot
  switch (s.phase) {
    case "init":
    case "checking":
      return [`⏳ 正在检索开源依赖（共 ${s.reqCount || "?"} 个直接依赖 + 传递依赖）...`]
    case "installing":
      return [
        `⏳ 依赖未就位（缺 ${s.missing.length || "?"}/${s.reqCount || "?"}），正在从内网 server 下载安装...`,
        `   ${clip(s.progress || "等待 pip 输出首个包...", 60)}`,
      ]
    case "ok":
      return [
        `✓ 依赖已就位（${s.reqCount}/${s.reqCount} 直接依赖 · CPython ${s.version || "?"}）${s.note || ""}`,
      ]
    case "failed":
      return [
        `✗ 依赖未就位：${clip(s.detail, 56)}`,
        `   ${clip(s.hint || "跑 PI_DEPS_CLI=1 看详细原因，或 PI_DEPS_VERBOSE=1 看判定过程", 60)}`,
      ]
    case "no-python":
      return [
        `! 未找到 CPython ${s.wantVersion || ""} 解释器，依赖未检查也未安装`,
        `   py -0 查看本机已装版本；装好后重开 pi`,
      ]
    case "no-pip":
      return ["! pip 太老（需 >= 22.2），无法检查依赖，建议先升级 pip"]
    default:
      return []
  }
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
 * 为什么不用 spawnSync + stdio:inherit——那样拿不到任何文本，"失败原因"就只能报退出码；
 * 更要紧的是 spawnSync 会**阻塞事件循环**，TUI 的 widget 在检查/下载期间一次都刷不出来，
 * 用户看到的"正在检索"就是一句假话。
 *
 * silent：只累积不转发（检查阶段用，pip 的 dry-run 日志对用户没价值）
 * onLine：每读到一段输出回调一次，供安装阶段把最新进度顶到 widget 上
 */
function runStreaming(cmd, args, { cwd, timeoutMs = 120000, silent = false, onLine = null }) {
  // 兜底：timeoutMs 传错/漏传时 setTimeout(fn, undefined) 会当 0 处理，
  // 子进程刚起来就被杀掉、输出为空——2026-11-18 就是因此在检查阶段永远判失败。
  const limit = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 120000
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
    }, limit)
    const finish = (payload) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(payload)
    }

    const forward = (chunk) => {
      const text = chunk.toString("utf8")
      output += text
      if (!silent) process.stderr.write(text)
      if (onLine) {
        // pip 的进度条用 \r 原地刷新，取最后一段才有意义
        const piece = text
          .split(/[\r\n]+/)
          .filter((s) => s.trim())
          .pop()
        if (piece) onLine(piece.trim())
      }
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

/** pip 是否支持 --dry-run（>= 22.2）。老版本无法做无副作用检查。异步跑，别阻塞 TUI 重绘。 */
async function pipSupportsDryRun(python) {
  const result = await runStreaming(
    python.cmd,
    [...python.args, "-m", "pip", "install", "--help"],
    { timeoutMs: 60000, silent: true },
  )
  return String(result.output || "").includes("--dry-run")
}

/**
 * 这段输出是不是"根本没连上"。
 * 必须和"连上了但没这个包"分开——pip 两种情形都会吐 No matching distribution found，
 * 连接失败时去报"server 缺包"会把人引到镜像同步上去，白折腾。
 */
function isNetworkFailure(text) {
  return /WinError 10061|WinError 10060|WinError 10065|WinError 11001|WinError 11002|Connection refused|NewConnectionError|Failed to establish a new connection|timed out|TimeoutError|ReadTimeoutError|Name or service not known|getaddrinfo failed|nodename nor servname/i.test(
    String(text || ""),
  )
}

/** 从 pip 输出里提取缺包名（checkDeps 与 verify 共用，口径必须一致）。 */
function extractMissing(output) {
  const names = [...String(output || "").matchAll(/No matching distribution found for (\S+)/g)].map(
    (m) => m[1].replace(/[;,.]+$/, ""),
  )
  return [...new Set(names)]
}

/**
 * 用 pip 自己做满足性检查：--no-index 保证完全不联网，只比对已装环境。
 * 异步跑（不阻塞事件循环），检查期间 widget 才能刷新出"正在检索"。
 * 返回 { ok, missing: string[], detail }。
 */
async function checkDeps(repoRoot, python) {
  const result = await runStreaming(
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
    { timeoutMs: CHECK_TIMEOUT_MS, cwd: repoRoot, silent: true },
  )

  if (result.code === 0) return { ok: true, missing: [], detail: "" }

  const output = result.output || ""
  // 只留最后几行：pip 的报错价值在尾部
  const detail = output.trim().split(/\r?\n/).slice(-4).join("\n")
  return { ok: false, missing: extractMissing(output), detail }
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
 * 内网 server 的 pip 参数：--index-url + 自动补 --trusted-host。
 * http 源不带 trusted-host 会被 pip 直接拒绝，所以两处（安装 / 预检）必须同口径。
 */
function indexArgs(indexUrl) {
  const args = ["--index-url", indexUrl]
  try {
    const host = new URL(indexUrl).hostname
    if (host) args.push("--trusted-host", host)
  } catch {
    /* 非法 URL 交给 pip 自己报错 */
  }
  return args
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
    ...indexArgs(config.indexUrl),
    "-r",
    path.join(repoRoot, "requirements.txt"),
  ]

  warn(`依赖未就位（${lacks}），开始从内网 server 下载：${config.indexUrl}`)
  warn(`pip 实时输出如下（下载进度即时刷新）...`)
  update({ phase: "installing", missing, progress: "" })

  // 把 pip 的最新一行顶到 widget 上。必须节流：进度条用 \r 高频刷新，
  // 每次 chunk 都重绘会让 TUI 抖得看不清。
  let lastPaint = 0
  const onLine = (line) => {
    // 只留真正体现进度的行：pip 还会吐 Looking in indexes / Requirement already satisfied
    // 这类噪音，顶到 widget 上等于没信息
    if (!/Downloading|Collecting|Installing collected|Processing|Using cached|Building wheel/i.test(line)) return
    const now = Date.now()
    if (now - lastPaint < 400) return
    lastPaint = now
    update({ progress: line })
  }

  const result = await runStreaming(python.cmd, args, {
    cwd: repoRoot,
    timeoutMs: config.timeoutMinutes * 60 * 1000,
    onLine,
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
 * 预检：只问内网 server「这些包能不能拿全」，不安装、不写缓存状态。
 * 意义是能把「网络不通」和「server 上缺包」分开——两者在 pip 输出里都会出现
 * "No matching distribution found"，运营时最容易被搞混的就是这两件事。
 *
 * --ignore-installed 是关键：不加的话 pip 会把"本机已装"当成已满足、根本不去 server 查，
 * 预检就退化成自查，白跑一遍。
 */
async function verifyAgainstServer(repoRoot) {
  const config = readConfig(repoRoot)
  const reqCount = countRequirements(repoRoot)
  const python = resolvePython(config)
  if (!python) {
    warn("! 未找到匹配的解释器，无法预检（py -0 查看本机已装版本）")
    return { status: "no-python", missing: [] }
  }
  if (!config.indexUrl) {
    warn("✗ 未配置内网 server 地址（.pi/deps.json 的 indexUrl），无法预检")
    return { status: "no-index", missing: [] }
  }
  if (!(await pipSupportsDryRun(python))) {
    warn("! pip 不支持 --dry-run（需 >= 22.2），无法预检")
    return { status: "no-pip", missing: [] }
  }

  const args = [
    ...python.args,
    "-m",
    "pip",
    "install",
    "--dry-run",
    "--ignore-installed",
    "--disable-pip-version-check",
    "--no-input",
    "--prefer-binary",
    "--timeout",
    "30",
    "--retries",
    "2",
    ...indexArgs(config.indexUrl),
    "-r",
    path.join(repoRoot, "requirements.txt"),
  ]

  warn(`预检：向 ${config.indexUrl} 查询 ${reqCount} 个直接依赖及其传递依赖（只查询，不安装）...`)
  const result = await runStreaming(python.cmd, args, {
    cwd: repoRoot,
    timeoutMs: config.timeoutMinutes * 60 * 1000,
  })

  if (result.timedOut) {
    warn(`✗ 预检超时（超过 ${config.timeoutMinutes} 分钟），server 响应过慢`)
    return { status: "failed", missing: [] }
  }
  if (result.code !== 0) {
    warn(`✗ 预检失败：${explainFailure(result.output, config.indexUrl)}`)
    const missing = extractMissing(result.output)
    // 网络不通时 pip 也会吐 No matching distribution，这时报缺包就是误导
    if (missing.length && !isNetworkFailure(result.output)) {
      warn(`  内网 server 上缺这些包：${missing.join(", ")}`)
    }
    return { status: "failed", missing }
  }

  const would = (result.output.match(/Would install ([^\r\n]+)/i) || [])[1]
  warn("✓ 预检通过：requirements.txt 的直接依赖与传递依赖都能从内网 server 取到")
  if (would) {
    const pkgs = would.trim().split(/\s+/).filter(Boolean)
    const shown = pkgs.length > 8 ? `${pkgs.slice(0, 8).join("  ")}  …等 ${pkgs.length} 个` : pkgs.join("  ")
    warn(`  解析出 ${pkgs.length} 个包：${shown}`)
  }
  return { status: "ok", missing: [] }
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
  update({ phase: "checking", reqCount, wantVersion: config.python, missing: [] })

  const python = resolvePython(config)
  if (!python) {
    // 解释器缺失是硬阻塞（依赖一个都不会装），不吃缓存、每次启动都说
    const hint = config.python ? ` CPython ${config.python}` : ""
    warn(`! 未找到${hint} 解释器，本次跳过依赖检查，依赖不会被安装`)
    warn(`  py -0 查看本机已装版本；装好后重开 pi 即生效（PI_DEPS_SKIP=1 可静音）`)
    writeState(repoRoot, { status: "no-python", checkedAt: now, requirementsMtimeMs: reqMtime })
    update({ phase: "no-python", wantVersion: config.python })
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
    update({ phase: "ok", version: python.version, note: `（${when}验证）` })
    return { status: "ok", missing: [] }
  }
  if (force) verbose("PI_DEPS_FORCE=1，忽略缓存重查")

  if (!(await pipSupportsDryRun(python))) {
    warn("! 目标环境的 pip 不支持 --dry-run（需 >= 22.2），无法检查依赖，建议先升级 pip")
    writeState(repoRoot, { status: "no-pip", checkedAt: now, requirementsMtimeMs: reqMtime })
    update({ phase: "no-pip" })
    return { status: "no-pip", missing: [] }
  }

  warn("检查依赖是否就位（本地比对，不联网，约 10 秒）...")
  update({ phase: "checking", detail: "比对本地已装环境" })
  const first = await checkDeps(repoRoot, python)
  if (first.ok) {
    writeState(repoRoot, {
      status: "ok",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
    })
    warn(`✓ 依赖已在位（${reqCount} 个直接依赖全部满足 · CPython ${python.version}，无需下载）`)
    update({ phase: "ok", version: python.version, note: "" })
    return { status: "ok", missing: [] }
  }

  verbose(`未满足: ${first.missing.join(", ") || first.detail}`)
  update({ phase: "installing", missing: first.missing })
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
    update({
      phase: "failed",
      detail: install.detail,
      // 完整重试命令留在 stderr（终端/日志里能复制），widget 只放一句怎么查
      hint: "详细原因与重试命令见终端 [deps] 输出，或跑 PI_DEPS_CLI=1 排查",
    })
    return { status: "failed", missing: first.missing }
  }

  const second = await checkDeps(repoRoot, python)
  if (!second.ok) {
    const detail = `下载完成但复检仍不满足：${second.missing.join(", ") || second.detail}`
    warn(`✗ ${detail}`)
    warn("  说明内网 server 上的包版本与 requirements.txt 不一致，或传递依赖被裁掉了")
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      requirementsMtimeMs: reqMtime,
      checkedAt: now,
      detail: second.detail,
    })
    update({
      phase: "failed",
      detail,
      hint: "核对内网 server 上的包版本与 requirements.txt 是否一致",
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
  update({ phase: "ok", version: python.version, note: `（本次装入 ${install.elapsed}s）` })
  return { status: "installed", missing: [] }
}

async function bootstrap() {
  if (process.env.PI_DEPS_SKIP === "1") {
    verbose("PI_DEPS_SKIP=1，跳过")
    update({ phase: "skipped" })
    return { status: "skipped", missing: [] }
  }
  const repoRoot = findRepoRoot(process.cwd())
  if (!repoRoot || !fs.existsSync(path.join(repoRoot, "requirements.txt"))) {
    verbose("当前目录及上 3 层没有 requirements.txt，跳过")
    update({ phase: "skipped" })
    return { status: "skipped", missing: [] }
  }
  return ensureDeps(repoRoot)
}

// ---------------------------------------------------------------------------
// pi 扩展入口
//
// 工厂**同步**返回：pi 才不会在启动阶段被卡住（启动阶段 TUI 还没渲染，卡在那里
// 用户只看到一片空白，比慢更难忍）。检查/安装在后台跑，进度通过 session_start
// 里的 widget 显示——这是唯一能在首页看见的通道。
// ---------------------------------------------------------------------------
export default function depsBootstrap(pi) {
  let finish
  const done = new Promise((resolve) => {
    finish = resolve
  })

  bootstrap()
    .catch((err) => {
      // 自举本身出错绝不能拦住 pi 启动
      warn(`自举异常（已忽略）：${err?.message || err}`)
      update({ phase: "failed", detail: `自举异常：${err?.message || err}` })
    })
    .finally(() => finish())

  if (!pi || typeof pi.on !== "function") return

  pi.on("session_start", async (_event, ctx) => {
    if (!ctx || !ctx.hasUI) {
      await done
      return
    }
    const paint = () => {
      const lines = renderStatus()
      try {
        // 放输入框下方：aboveEditor 已被 pptagent-welcome 占了，别抢位置
        ctx.ui.setWidget("deps-status", lines.length ? lines : undefined, { placement: "belowEditor" })
      } catch {
        /* 会话已被替换，ctx 失效 */
      }
    }

    redraw = paint
    paint()
    // 最多等 90 秒：装依赖可能几分钟，不能把 pi 卡死；超时先放行，装完 widget 自己会刷新
    await Promise.race([done, new Promise((resolve) => setTimeout(resolve, 90_000))])
    paint()

    if (snapshot.phase !== "ok" && snapshot.phase !== "skipped") {
      try {
        ctx.ui.setStatus("deps", "依赖未就位")
      } catch {
        /* 同上 */
      }
    }
  })

  pi.on("session_shutdown", () => {
    redraw = null
  })
}

// ---------------------------------------------------------------------------
// 命令行模式：
//   预检（只问 server 能不能拿全，不装）  PI_DEPS_VERIFY=1 npx jiti <本文件>
//   检查 + 需要就装                      PI_DEPS_CLI=1    npx jiti <本文件>
// setup.bat 与人工排查都走这两条
// ---------------------------------------------------------------------------
if (process.env.PI_DEPS_VERIFY === "1") {
  const repoRoot = findRepoRoot(process.cwd())
  if (!repoRoot) {
    warn("! 当前目录及上 3 层没有 requirements.txt，无需预检")
    process.exit(0)
  }
  verifyAgainstServer(repoRoot).then((result) => process.exit(result.status === "ok" ? 0 : 1))
} else if (process.env.PI_DEPS_CLI === "1") {
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
