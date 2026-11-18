// @ts-nocheck
/**
 * 依赖自举（deps bootstrap）—— 内网镜像单源版
 *
 * 时机：pi 加载扩展时执行本文件，**在 session_start / 资源发现 / provider 注册之前**。
 *       pi 文档：工厂函数返回 Promise 时，pi 会 await 它再继续启动
 *       （docs/extensions.md「Writing an Extension」）。所以这里做成 async 工厂。
 *
 * 唯一安装来源：**内网 server**（本文件内 `CONFIG.indexUrl`，固定值）。
 *   - 不回退公网 PyPI：离线机连不上，回退只会白等；配错了也不该被兜底掩盖。
 *   - 不使用 `.pi/wheels` 本地目录：依赖全部由内网 server 供应，本地包是另一条要维护的路径。
 *   - 地址写死在代码里，**不接受文件或环境变量覆盖**——要换就改这一处。
 *
 * 依赖清单：**只认仓库根的 requirements.txt**，不在本文件里另存一份。
 *   一份真相就不会漂移，也不会有"ts 里写着 A、requirements.txt 写着 B"的歧义。
 *   传递依赖由 pip 自行解析，清单里只写直接依赖。
 *
 * 判据：不猜 import 名、不自己比版本，直接让 pip 判断——
 *       `pip install --dry-run --no-index -r requirements.txt`（纯本地比对，不联网）
 *       退出码 0 = 全部满足（含版本约束）；非 0 = 缺包或版本不符。
 *       这样 `python-pptx→pptx`、`Pillow→PIL`、`PyYAML→yaml` 这类映射一概不用维护。
 *
 * 反馈口径：**报全部依赖（直接 + 传递）**，不是只报 requirements.txt 里那几行。
 *   `requirements.txt` 只写 7 个直接依赖，实际要装 20 个（7 直接 + 13 传递）；
 *   只显示 7 会让用户以为"就位"指的是 7 个，2026-11-18 用户明确提过这一点。
 *   闭包由 `collectDepTree` 用 importlib.metadata 在本地递归算出（不联网、约 0.3 秒），
 *   而"是否满足"仍由 pip 的 --dry-run 裁定——两者分工别搞混。
 *
 * 终端反馈（每次启动都有，不让用户猜）：
 *   检查        `检查依赖是否就位（本地比对，不联网，约 10 秒）...`
 *   已就位      `✓ 依赖已在位（20 个包：7 直接 + 13 传递 全部满足 · CPython 3.12.10，无需下载）`
 *   缓存命中    `✓ 依赖已在位（20 个包：7 直接 + 13 传递，3.2 小时前已验证；本次跳过重复检查）`
 *   下载中      `依赖未就位（缺 3/7：...），开始从内网 server 下载：<url>` + pip 逐块实时输出
 *   完成        `✓ 依赖安装完成（耗时 42s，复检通过 · CPython 3.12.10）` + `本次装入：a-1.0 b-2.0 …`
 *   失败        `✗ 依赖安装失败：<人话原因>` + 手动重试命令
 *
 * 配置：**全部内联在本文件的 CONFIG 常量里**（约第 100 行），不读任何外部配置文件。
 *   - 内网地址、解释器版本、安装超时、是否拦指令都在 CONFIG 里，固定值，
 *     `.pi/deps.json` 已废弃并删除，也不接受环境变量覆盖这几项。
 *   - **依赖清单是唯一例外**：它不内联，只从仓库根的 requirements.txt 读。
 *     清单属于项目内容、会随版本变，放在项目里改起来自然；内联一份副本只会制造漂移。
 *     所以本文件不持有任何依赖版本号，也就没有"两份清单要对齐"这件事。
 *
 * 环境变量只剩**应急开关**（不是配置，按需临时设，不影响地址与清单）：
 *   PI_DEPS_SKIP=1       完全跳过（应急）
 *   PI_DEPS_FORCE=1      忽略 24h 缓存，强制真跑一次检查
 *   PI_DEPS_VERIFY=1     预检模式：只向 server 查询能否取全，不安装、不写状态，退出码 0/1
 *   PI_DEPS_VERBOSE=1    打印判定过程
 *   PI_DEPS_CLI=1        命令行模式（检查/安装后退出，不挂到 pi 上）
 *   PI_DEPS_GATE=0       依赖未就位时不拦用户指令（默认拦）
 *   PI_DEPS_NO_PROBE=1   跳过第 5 步的连通性探测
 *
 * 这个脚本是**完全独立的**：一个 .ts 文件丢进任意 pi 项目的 `.pi/extensions/`
 * （或 `~/.pi/agent/extensions/` 全局）就能工作——不需要 deps.json、不需要额外配置，
 * 内网地址与解释器版本都在文件里。唯一的运行时输入是仓库根的 requirements.txt
 * （依赖清单）与"当前目录往上 3 层内有没有 `.pi` / requirements.txt"（定位仓库根）。
 *
 * 两套输出通道，两条都要留着：
 *   1) TUI widget（首页可见）—— `session_start` + `ctx.ui.setWidget("deps-status", ..., belowEditor)`。
 *      **这个是给用户看的**。pi 进 interactive 模式会切 alternate screen 重绘，stderr 全被盖掉；
 *      2026-11-18 就是因此首页一片干净，用户压根不知道依赖检没检查。
 *   2) stderr `[deps] ...` —— CLI 模式、print 模式、日志重定向、人工排查用。
 *
 * 日志按**步骤**走，每一步都说清"在做什么、结果如何"（序号 1-7）：
 *   1) 定位解释器     2) 统计依赖闭包   3) 读来源地址     4) 探测 server 连通性
 *   5) 本地比对       6) 下载安装       7) 复检
 * 依赖缺了才知道去看日志就太晚了，所以每一步都尽量给出"通/不通/通到哪一层"。
 *
 * 缓存为什么留：`pip install --dry-run --no-index` 实测 11 秒，每次启动都跑等于每次多等 11 秒；
 * 但缓存命中时提示里必须写清"X 前已验证、本次跳过"，不能让人误以为刚查过。
 * 要立刻真查一次：`PI_DEPS_FORCE=1 pi`。
 *
 * 刻意不做的事：
 *   - 安装失败**不让 pi 崩溃**。起不来的 pi 没法排查问题，改成打醒目警告 + 人话原因后放行。
 *   - 不 block 版本不符的解释器。配了 `python` 却找不到该版本时跳过并提示。
 *   - 探测不到 server 时不假装自己能连：TCP 明确拒绝/DNS 解析失败就直接给结论，
 *     不白等 pip 的 30s 超时 × 2 次重试；只有超时这种"可能被防火墙拖住"的情况才继续试 pip。
 */

import { spawn, spawnSync } from "node:child_process"
// 全部用命名空间导入，不用 default 导入：`node:net` 这类内置模块**没有 default export**，
// 在原生 ESM 下 `import net from "node:net"` 会直接抛
// "does not provide an export named 'default'"，而且报错发生在模块加载阶段——
// factory 根本不会被调用，widget 也不会画，现象就是"这个扩展毫无反应"。
// jiti 转 CJS 时能 interop 掉，所以开发机上用 jiti 跑是好的，换台机器换了加载路径就未必。
// 命名空间导入在两种加载方式下都合法，是这里唯一安全写法。
import * as net from "node:net"
import * as http from "node:http"
import * as https from "node:https"
import * as fs from "node:fs"
import * as path from "node:path"

const STATE_FILE = ".pi/work/.deps-state.json"
const CACHE_TTL_MS = 24 * 60 * 60 * 1000
const CHECK_TIMEOUT_MS = 180 * 1000
const PROBE_TIMEOUT_MS = 5000

// ---------------------------------------------------------------------------
// 全部配置。这是**唯一**的配置来源：不读 .pi/deps.json、不接受
// PI_PIP_INDEX_URL / PI_DEPS_PY_VERSION / PI_DEPS_TIMEOUT_MIN 这类覆盖。
// 要换地址、换解释器版本、换超时——改这里，别去环境里设变量。
//
// 依赖清单不在下面：它只从仓库根的 requirements.txt 读（见 CONFIG.requirementsFile）。
// 版本号属于项目内容、随版本走，抄一份到代码里只会多出一个要同步的地方。
// ---------------------------------------------------------------------------
const CONFIG = {
  // 目标解释器：只认这个精确版本；设 null 则用 PATH 上第一个能跑起来的 python
  python: "3.12",

  // 内网 server：装包唯一来源，固定值，不回退公网 PyPI
  indexUrl: "http://172.30.39.34:18080/simple",

  // 安装总超时（分钟）
  timeoutMinutes: 15,

  // 依赖未就位时是否拦用户指令
  gate: true,

  // 依赖清单：相对仓库根的文件名。只写直接依赖，传递依赖交给 pip 解析。
  // 这是**唯一**的清单真相，本文件不另存副本。
  requirementsFile: "requirements.txt",
}

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
  reqCount: 0, // requirements.txt 里的直接依赖条数
  directCount: 0, // 闭包里实际解析到的直接依赖个数
  transitiveCount: 0, // 闭包里传递依赖个数
  total: 0, // directCount + transitiveCount（用户要看的是这个）
  pkgs: [], // 闭包里的包名-版本，用于展示
  missing: [],
  version: null,
  wantVersion: null,
  pythonPath: "", // 实际用的解释器，回答"这套依赖装在哪个 Python 里"
  source: "", // 依赖来源地址（内网 server），回答"从哪拿的"
  step: 0, // 当前步骤号（1-7）
  stepLabel: "", // 当前步骤在做什么
  stepDetail: "", // 当前步骤的结果（可达/不可达/缺几个…）
  installedNow: false, // 本次是否真的下载安装了（区分"本来就有"和"刚装好"）
  cachedAt: 0, // 缓存命中时，上次验证的时间戳
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

// ---------------------------------------------------------------------------
// 步骤日志：1-7 步，每一步都说清"在做什么、结果如何"
//
// 用户的原话是"期望通过日志一步一步告知访问哪个服务器地址、服务地址是否通"。
// 所以不是"最后给一句失败原因"，而是每一步都有落点：解释器是哪个、来源是哪个地址、
// TCP 通不通、索引页返回什么、缺几个包、从哪儿下的、复检过没过。
// stderr 给完整版（终端/日志可回溯），widget 给"第几步 + 这一步的结论"。
// ---------------------------------------------------------------------------
const TOTAL_STEPS = 7

function stepHead(n, label) {
  update({ step: n, stepLabel: label, stepDetail: "" })
  warn(`${n}/${TOTAL_STEPS} ${label}`)
}

function stepInfo(text) {
  update({ stepDetail: text })
  warn(`    ${text}`)
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

/**
 * "20 个包：7 直接 + 13 传递"——用户要的是这个数，不是 requirements.txt 的 7 行。
 * 闭包没算出来（老 pip / 读取失败）时如实退回"只报了直接依赖"，不假装统计过。
 */
function counts() {
  const s = snapshot
  if (s.total) return `${s.total} 个包：${s.directCount} 直接 + ${s.transitiveCount} 传递`
  // 闭包还没算出来（启动最初的零点几秒，或极端情况下读取失败）：只报直接依赖，
  // 不编造传递依赖个数——这个分支只会短暂出现，稳态永远是上面那行。
  if (s.reqCount) return `${s.reqCount} 个直接依赖`
  return "依赖"
}

// ---------------------------------------------------------------------------
// server 连通性探测
//
// 为什么单独探：pip 失败时只会吐一坨重试日志和 "No matching distribution found"，
// 分不清"地址不通"和"server 上没这个包"——而这两件事的处置完全不同（找网络 vs 补包）。
// 先探 TCP、再探索引页，能把失败钉到具体一层。
//
// 探测失败 ≠ 装了会失败：代理、只放行 HTTP 的防火墙都会让 TCP 探测失真。
// 所以只有**确定性失败**（拒绝连接 / 解析失败 / 路由不可达 / 索引页 404、401、403）
// 才跳过 pip 直接给结论；超时这类不确定情况仍然让 pip 试一次。
// ---------------------------------------------------------------------------
function parseIndexUrl(indexUrl) {
  try {
    const u = new URL(indexUrl)
    const port = Number(u.port) || (u.protocol === "https:" ? 443 : 80)
    return { host: u.hostname, port, path: u.pathname || "/", protocol: u.protocol }
  } catch {
    return null
  }
}

function probeTcp(host, port, timeoutMs) {
  return new Promise((resolve) => {
    const startedAt = Date.now()
    const socket = new net.Socket()
    let settled = false
    const done = (ok, code) => {
      if (settled) return
      settled = true
      const ms = Date.now() - startedAt
      try {
        socket.destroy()
      } catch {
        /* 已关 */
      }
      resolve({ ok, code, ms })
    }
    socket.setTimeout(timeoutMs)
    socket.once("connect", () => done(true, "OK"))
    socket.once("timeout", () => done(false, "ETIMEDOUT"))
    socket.once("error", (err) => done(false, err?.code || "ERROR"))
    try {
      socket.connect(port, host)
    } catch (err) {
      done(false, err?.code || "ERROR")
    }
  })
}

/** 只取状态码，不下载页面内容——我们要的是"索引入口通不通"。 */
function probeHttp(indexUrl, timeoutMs) {
  return new Promise((resolve) => {
    const target = parseIndexUrl(indexUrl)
    if (!target) return resolve({ ok: false, code: "BAD_URL", status: 0, ms: 0 })

    const startedAt = Date.now()
    let settled = false
    const done = (payload) => {
      if (settled) return
      settled = true
      resolve({ ms: Date.now() - startedAt, ...payload })
    }
    const lib = target.protocol === "https:" ? https : http
    try {
      const req = lib.get(
        {
          host: target.host,
          port: target.port,
          path: target.path,
          timeout: timeoutMs,
          headers: { "User-Agent": "pi-deps-bootstrap" },
        },
        (res) => {
          res.resume() // 丢弃正文，只要状态码
          done({ ok: res.statusCode >= 200 && res.statusCode < 400, code: "HTTP", status: res.statusCode })
        },
      )
      req.on("timeout", () => {
        req.destroy()
        done({ ok: false, code: "ETIMEDOUT", status: 0 })
      })
      req.on("error", (err) => done({ ok: false, code: err?.code || "ERROR", status: 0 }))
    } catch (err) {
      done({ ok: false, code: err?.code || "ERROR", status: 0 })
    }
  })
}

/** 把 socket 错误码翻成人话。 */
function explainProbeCode(code) {
  switch (code) {
    case "ECONNREFUSED":
      return "端口拒绝连接：该端口没有服务在听，或防火墙显式拒绝"
    case "ENOTFOUND":
    case "EAI_AGAIN":
      return "主机名解析失败：indexUrl 里应写 IP，别写主机名"
    case "EHOSTUNREACH":
    case "ENETUNREACH":
      return "路由不可达：这台机器没有到该网段的路由 / VPN 没连"
    case "ETIMEDOUT":
      return "连接超时：网络不通，或防火墙静默丢包"
    case "ECONNRESET":
      return "连接被重置：对端或中间设备中断了握手"
    case "BAD_URL":
      return "indexUrl 不是合法 URL"
    default:
      return `连接失败（错误码 ${code}）`
  }
}

/** 确定性失败：不用再让 pip 白等 30s×2 重试。 */
const FATAL_TCP_CODES = new Set([
  "ECONNREFUSED",
  "ENOTFOUND",
  "EAI_AGAIN",
  "EHOSTUNREACH",
  "ENETUNREACH",
  "ECONNRESET",
])

/**
 * 探测结果：
 *   ok     true 表示 TCP 通且索引页能返回 2xx/3xx
 *   fatal  true 表示确定拿不到包，不必再试 pip
 *   lines  每一步的结论，按顺序打到日志和 widget 上
 */
async function probeServer(indexUrl) {
  const target = parseIndexUrl(indexUrl)
  if (!target) {
    return { ok: false, fatal: true, lines: ["✗ CONFIG.indexUrl 不是合法 URL（本文件顶部的 CONFIG 常量）"] }
  }

  const tcp = await probeTcp(target.host, target.port, PROBE_TIMEOUT_MS)
  if (!tcp.ok) {
    return {
      ok: false,
      fatal: FATAL_TCP_CODES.has(tcp.code),
      lines: [`✗ TCP ${target.host}:${target.port} 不可达（${tcp.ms}ms）→ ${explainProbeCode(tcp.code)}`],
    }
  }

  const lines = [`✓ TCP ${target.host}:${target.port} 可达（${tcp.ms}ms）`]
  const res = await probeHttp(indexUrl, PROBE_TIMEOUT_MS)
  if (res.ok) {
    lines.push(`✓ 索引页 ${target.path} 返回 ${res.status}（${res.ms}ms）`)
    return { ok: true, fatal: false, lines }
  }
  if (res.code === "ETIMEDOUT") {
    lines.push(`? 索引页 ${target.path} 请求超时（${res.ms}ms），仍让 pip 试一次`)
    return { ok: false, fatal: false, lines }
  }
  const hint =
    res.status === 404
      ? `✗ 索引页 ${target.path} 返回 404：server 活着但路径不对（pip 要求 /simple），或镜像没建索引`
      : res.status === 401 || res.status === 403
        ? `✗ 索引页返回 ${res.status}：server 要求认证，pip 需要凭据才能取包`
        : `✗ 索引路径 ${target.path} 返回 ${res.status}：server 活着但取包入口异常`
  lines.push(hint)
  return { ok: false, fatal: [404, 401, 403].includes(res.status), lines }
}

/** 解释器路径只留尾部两级，够识别又不撑爆一行。 */
function shortPath(p) {
  const parts = String(p || "").split(/[\\/]/).filter(Boolean)
  return parts.length <= 2 ? String(p || "") : `…\\${parts.slice(-2).join("\\")}`
}

/** 包清单：一行放得下就全列，放不下就截断并标注还剩几个。 */
function pkgLine() {
  const s = snapshot
  if (!s.pkgs.length) return ""
  const shown = []
  let width = 0
  for (const p of s.pkgs) {
    // 中文宽度按 2 算太苛，包名全是 ASCII，直接按字符数控制
    if (width + p.length + 2 > 62) break
    shown.push(p)
    width += p.length + 2
  }
  const rest = s.pkgs.length - shown.length
  return `已装 ${shown.join("  ")}${rest > 0 ? `  …另 ${rest} 个` : ""}`
}

/**
 * 首页 widget 的内容：任何阶段都有一句能看懂的话。
 * 就位时给三件事：多少包、**从哪个地址来的**、装在哪个解释器里；
 * 未就位时给两件事：停在第几步、那一步的结论是什么。
 */
function renderStatus() {
  const s = snapshot
  const at = (n, label) => (n ? `${n}/${TOTAL_STEPS} ` : "") + label

  switch (s.phase) {
    case "init":
    case "checking":
      return [
        `⏳ ${at(s.step, s.stepLabel || "正在检索开源依赖")}（${counts()}）`,
        s.stepDetail ? `   ${clip(s.stepDetail, 64)}` : "",
      ].filter((line) => line.trim())

    case "installing":
      return [
        `⏳ ${at(s.step || 6, "从内网 server 下载安装")}（缺 ${s.missing.length || "?"} 个）`,
        `   ${clip(s.progress || s.stepDetail || "等待 pip 输出首个包...", 64)}`,
      ]

    case "ok":
      return [
        `✓ 依赖已就位（${counts()} · CPython ${s.version || "?"}）${s.note || ""}`,
        `   来源 ${clip(s.source || "未配置", 46)}${s.installedNow ? "（本次刚下载）" : "（本次未下载）"}`,
        `   解释器 ${shortPath(s.pythonPath) || "?"}`,
        s.pkgs.length ? `   ${pkgLine()}` : "",
      ].filter((line) => line.trim())

    case "failed":
      return [
        `✗ 依赖未就位（停在第 ${s.step || "?"}/${TOTAL_STEPS} 步：${s.stepLabel || "检查"}）`,
        `   ${clip(s.detail || "原因见终端 [deps] 输出", 60)}`,
        s.hint ? `   ${clip(s.hint, 60)}` : "",
      ].filter((line) => line.trim())

    case "no-python":
      return [
        `! 未找到 CPython ${s.wantVersion || ""} 解释器（第 1/${TOTAL_STEPS} 步），依赖未检查也未安装`,
        `   py -0 查看本机已装版本；装好后重开 pi`,
      ]
    case "no-pip":
      return ["! pip 太老（需 >= 22.2），无法检查依赖，建议先升级 pip"]
    case "skipped":
      // 跳过也要留痕：静默跳过等于让用户以为扩展坏了。
      // 但要克制——findRepoRoot 没找到仓库根（cwd 不对）时也走这里，
      // 而那种情况下用户并不在项目里工作，不该在首页挂一行常驻提示。
      return s.note ? [`- 依赖自举跳过：${clip(s.note, 64)}`] : []
    default:
      return []
  }
}

/**
 * 从起始目录向上找项目根：有 requirements.txt 或 `.pi` 目录的那一层。
 * 只上溯 3 层——再往上容易跨进父目录的另一个无关项目，把依赖装错地方。
 * 两个锚点都认：`.pi` 是本文件自己住的地方（本项目必然有），requirements.txt
 * 是依赖清单（脚本要读它）。任一存在都足以定位根目录。
 */
function findRepoRoot(start) {
  let dir = start
  for (let i = 0; i < 3; i += 1) {
    if (fs.existsSync(path.join(dir, CONFIG.requirementsFile)) || fs.existsSync(path.join(dir, ".pi"))) {
      return dir
    }
    const parent = path.dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return null
}

/**
 * 取配置：直接返回内联常量的一份拷贝。
 * 保留函数形态是为了调用点不用改，也留一个将来加校验的位置（比如断言 indexUrl 是合法 URL）。
 * 不做任何文件读取、不接受环境变量覆盖——"固定不变"就是这么落实的。
 * 唯一的外部输入 requirementsFile 只是文件名，内容仍然从仓库里读。
 */
function readConfig() {
  const config = {
    python: CONFIG.python,
    indexUrl: CONFIG.indexUrl,
    timeoutMinutes: CONFIG.timeoutMinutes,
    gate: CONFIG.gate,
    requirementsFile: CONFIG.requirementsFile,
    origin: "扩展内置配置",
  }
  if (process.env.PI_DEPS_GATE === "0") config.gate = false
  verbose(`配置: ${JSON.stringify(config)}（来自扩展内置 CONFIG）`)
  return config
}

/** 依赖清单的绝对路径：仓库根 + CONFIG.requirementsFile。 */
function reqFileOf(repoRoot) {
  return path.join(repoRoot, CONFIG.requirementsFile)
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
function runStreaming(cmd, args, { cwd, timeoutMs = 120000, silent = false, onLine = null, env = null }) {
  // 兜底：timeoutMs 传错/漏传时 setTimeout(fn, undefined) 会当 0 处理，
  // 子进程刚起来就被杀掉、输出为空——2026-11-18 就是因此在检查阶段永远判失败。
  const limit = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 120000
  return new Promise((resolve) => {
    let child
    try {
      child = spawn(cmd, args, {
        cwd,
        windowsHide: true,
        env: env ? { ...process.env, ...env } : process.env,
      })
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

/** 读清单原文；文件不存在时返回 null（调用方负责报错/跳过）。 */
function readReqText(reqFile) {
  try {
    return fs.readFileSync(reqFile, "utf8")
  } catch {
    return null
  }
}

/** 清单里的有效行（去掉空行与 # 注释）。 */
function reqLines(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith("#"))
}

/**
 * 清单指纹，用于缓存失效判定。取文件内容的简单哈希而不是 mtime：
 * mtime 在 git checkout / 拷贝后可能整体变化，内容没变却让缓存失效；
 * 反过来 mtime 精度在部分文件系统上只到秒，同一秒内改动会漏判。内容哈希没这两个问题。
 */
function requirementsStamp(reqFile) {
  const text = readReqText(reqFile)
  const lines = reqLines(text)
  let hash = 0
  const body = lines.join("|")
  for (let i = 0; i < body.length; i += 1) hash = (hash * 31 + body.charCodeAt(i)) | 0
  return `${lines.length}-${hash}`
}

/** 直接依赖条数，只用于提示语里报个数。 */
function countRequirements(reqFile) {
  return reqLines(readReqText(reqFile)).length
}

// ---------------------------------------------------------------------------
// 依赖闭包（含传递依赖）
//
// 为什么需要：requirements.txt 只有 7 行，但 pip 会顺着依赖树再装 13 个传递依赖，
// 合计 20 个。只报"7 个直接依赖"会让人以为就位与否只看这 7 个。
//
// 为什么不用 pip 拿这个数：`pip install --dry-run` 的输出只有直接依赖
// （"Requirement already satisfied: xxx"），拿不到闭包；`--report` 又不含已满足项，
// 除非 `--ignore-installed`，而那个必须联网。
// 所以本地读 importlib.metadata 的 Requires-Dist 递归——不联网、不装包、约 0.3 秒。
// 注意分工：**闭包只用于显示；是否满足仍由 pip 的 --dry-run 裁定**（权威口径不变）。
// ---------------------------------------------------------------------------
const DEP_SCAN_SCRIPT = String.raw`
import importlib, json, re, sys
from importlib import metadata

def norm(n):
    return re.sub(r"[-_.]+", "-", (n or "")).lower().strip()

def req_name(spec):
    m = re.match(r"\s*([A-Za-z0-9][A-Za-z0-9._-]*)", spec or "")
    return norm(m.group(1)) if m else None

def load_requirement():
    for mod in ("packaging.requirements", "pip._vendor.packaging.requirements"):
        try:
            return importlib.import_module(mod).Requirement
        except Exception:
            continue
    return None

Requirement = load_requirement()

def allowed(spec):
    """环境 marker 不满足的依赖不算进来（extra 的依赖也一律不算）。"""
    if ";" not in spec:
        return True
    if Requirement is None:
        # 没有 packaging 时保守保留：多算一个名字（它多半没装，会被跳过）比漏算安全
        return True
    try:
        r = Requirement(spec)
        return r.marker is None or r.marker.evaluate({"extra": ""})
    except Exception:
        return True

def dist_of(name):
    try:
        return metadata.distribution(name)
    except Exception:
        return None

reqfile = sys.argv[1]
roots = []
for line in open(reqfile, encoding="utf-8", errors="replace"):
    t = line.strip()
    if not t or t.startswith("#") or t.startswith("-"):
        continue
    n = req_name(t)
    if n:
        roots.append(n)

direct, transitive, missing = [], [], []
seen = set()
queue = [(n, 0) for n in roots]
while queue:
    name, depth = queue.pop(0)
    if name in seen:
        continue
    seen.add(name)
    dist = dist_of(name)
    if dist is None:
        # 直接依赖没装 = 真缺；传递依赖没装可能是 marker 不适用，不当成缺失
        if depth == 0:
            missing.append(name)
        continue
    label = "%s-%s" % (dist.metadata["Name"] or name, dist.version)
    (direct if depth == 0 else transitive).append(label)
    for spec in (dist.requires or []):
        if not allowed(spec):
            continue
        dep = req_name(spec)
        if dep and dep not in seen:
            queue.append((dep, depth + 1))

print(json.dumps({
    "direct": direct,
    "transitive": transitive,
    "missing": missing,
    "total": len(direct) + len(transitive),
}))
`

/**
 * 算出"直接 + 传递"的完整闭包。失败返回 null，调用方退回只报直接依赖条数。
 * 用临时脚本而不是 `python -c`：多行脚本在 Windows 上过 cmd 的引号转义很容易翻车。
 */
async function collectDepTree(repoRoot, python, reqFile) {
  const scriptPath = path.join(repoRoot, ".pi", "work", ".depscan.py")
  try {
    fs.mkdirSync(path.dirname(scriptPath), { recursive: true })
    fs.writeFileSync(scriptPath, DEP_SCAN_SCRIPT, "utf8")
  } catch (err) {
    verbose(`依赖闭包脚本写入失败（不影响检查）: ${err.message}`)
    return null
  }

  const result = await runStreaming(
    python.cmd,
    [...python.args, scriptPath, reqFile],
    { timeoutMs: 30000, cwd: repoRoot, silent: true, env: { PYTHONIOENCODING: "utf-8" } },
  )
  if (result.code !== 0) {
    verbose(`依赖闭包计算失败（不影响检查）: ${String(result.output || "").slice(-200)}`)
    return null
  }
  try {
    const last = String(result.output || "").trim().split(/\r?\n/).pop()
    const parsed = JSON.parse(last)
    if (!parsed || !Array.isArray(parsed.direct)) return null
    return parsed
  } catch {
    return null
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
async function checkDeps(reqFile, python, cwd) {
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
      reqFile,
    ],
    { timeoutMs: CHECK_TIMEOUT_MS, cwd, silent: true },
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
async function installDeps(repoRoot, python, config, missing, reqFile) {
  const startedAt = Date.now()
  const elapsed = () => Math.round((Date.now() - startedAt) / 1000)
  const lacks = missing.length
    ? `${missing.length} 个未满足：${missing.slice(0, 6).join(", ")}${missing.length > 6 ? " …" : ""}`
    : "版本不满足"

  if (!config.indexUrl) {
    return {
      ok: false,
      elapsed: 0,
      output: "",
      installed: null,
      detail: "CONFIG.indexUrl 为空，拒绝走公网 PyPI",
      hint: "把本文件顶部的 CONFIG.indexUrl 改回内网地址后重开 pi",
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
    reqFile,
  ]

  warn(`  ${lacks}`)
  warn(`  下载中，pip 原始输出如下（逐包进度实时刷新）...`)
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
      hint: `加大超时：把本文件 CONFIG.timeoutMinutes 改成 30 后重开 pi（当前 ${config.timeoutMinutes}）`,
    }
  }
  if (result.error) {
    return {
      ok: false,
      elapsed: elapsed(),
      output: result.output,
      installed: null,
      detail: `pip 无法启动：${result.error.message}`,
      hint: `手动重试：${python.executable} -m pip install --index-url ${config.indexUrl} -r ${reqFile}`,
    }
  }
  if (result.code !== 0) {
    return {
      ok: false,
      elapsed: elapsed(),
      output: result.output,
      installed: null,
      detail: explainFailure(result.output, config.indexUrl),
      hint: `手动重试：${python.executable} -m pip install --index-url ${config.indexUrl} -r ${reqFile}`,
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
  const config = readConfig()
  const reqFile = reqFileOf(repoRoot)
  const reqCount = countRequirements(reqFile)
  const python = resolvePython(config)
  if (!python) {
    warn("! 未找到匹配的解释器，无法预检（py -0 查看本机已装版本）")
    return { status: "no-python", missing: [] }
  }
  if (!config.indexUrl) {
    warn("✗ CONFIG.indexUrl 为空，无法预检（本文件顶部的 CONFIG 常量被改坏了）")
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
    reqFile,
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
  warn(`✓ 预检通过：${CONFIG.requirementsFile} 的直接依赖与传递依赖都能从内网 server 取到`)
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
  const config = readConfig()
  const reqFile = reqFileOf(repoRoot)
  const reqStamp = requirementsStamp(reqFile)
  const reqCount = countRequirements(reqFile)
  const cached = readState(repoRoot)
  const now = Date.now()
  const force = process.env.PI_DEPS_FORCE === "1"
  const startedAt = Date.now()

  gateEnabled = config.gate
  update({ phase: "checking", reqCount, wantVersion: config.python, missing: [], source: config.indexUrl })

  // ---- 1/7 定位解释器 ------------------------------------------------------
  stepHead(1, config.python ? `定位 CPython ${config.python} 解释器` : "定位 Python 解释器")
  const python = resolvePython(config)
  if (!python) {
    // 解释器缺失是硬阻塞（依赖一个都不会装），不吃缓存、每次启动都说
    const hint = config.python ? ` CPython ${config.python}` : ""
    stepInfo(`✗ 没有找到${hint}解释器，依赖一个都装不了`)
    warn("  py -0 查看本机已装版本；装好后重开 pi 即生效（PI_DEPS_SKIP=1 可静音）")
    writeState(repoRoot, { status: "no-python", checkedAt: now, requirementsStamp: reqStamp })
    update({ phase: "no-python", wantVersion: config.python })
    return { status: "no-python", missing: [] }
  }
  stepInfo(`✓ ${python.executable}（${python.version}）`)
  update({ pythonPath: python.executable })

  // ---- 2/7 统计依赖闭包（本地，不联网） ------------------------------------
  stepHead(2, `统计依赖闭包（${CONFIG.requirementsFile} ${countRequirements(reqFile)} 条）`)
  const tree = await collectDepTree(repoRoot, python, reqFile)
  if (tree) {
    update({
      directCount: tree.direct.length,
      transitiveCount: tree.transitive.length,
      total: tree.total,
      pkgs: [...tree.direct, ...tree.transitive],
    })
    stepInfo(`✓ 统计完成：${tree.total} 个包（${tree.direct.length} 直接 + ${tree.transitive.length} 传递）`)
    if (tree.missing.length) stepInfo(`  直接依赖里已缺：${tree.missing.join(", ")}`)
  } else {
    stepInfo(`? 闭包统计失败，退回只报 ${reqCount} 个直接依赖（不影响判定）`)
  }

  // ---- 3/7 来源地址 --------------------------------------------------------
  stepHead(3, "读取依赖来源地址")
  stepInfo(`✓ ${config.indexUrl}（来自${config.origin}）`)

  // 缓存命中：同一解释器 + 同一来源 + 清单没改过 + 24h 内验证过。
  // 提示里必须写清"多久前验的、本次没重跑"，否则会被误读成"刚检查过"。
  const cacheHit =
    !force &&
    cached?.status === "ok" &&
    cached.python === python.executable &&
    cached.requirementsStamp === reqStamp &&
    cached.indexUrl === config.indexUrl &&
    now - (cached.checkedAt || 0) < CACHE_TTL_MS
  if (cacheHit) {
    const age = formatAge(now - (cached.checkedAt || 0))
    const when = age === "刚刚" ? "刚刚已" : `${age}前已`
    stepInfo(`✓ ${when}用这个解释器、这个地址验过同一份清单，24 小时内不重复检查`)
    stepInfo("  要立刻真查一次：PI_DEPS_FORCE=1 重开 pi")
    warn(`✓ 依赖已在位（${counts()}，${when}验证；本次跳过重复检查）`)
    warn(`  来源 ${config.indexUrl}（本次未下载）`)
    warn(`  解释器 ${python.executable}`)
    update({
      phase: "ok",
      version: python.version,
      note: `（${when}验证）`,
      cachedAt: cached.checkedAt,
      installedNow: false,
    })
    return { status: "ok", missing: [] }
  }
  if (force) verbose("PI_DEPS_FORCE=1，忽略缓存重查")

  // ---- 4/7 本地比对 --------------------------------------------------------
  // 先比对本机装没装、再探 server 通不通：依赖本来就齐时根本不需要碰网络，
  // 探测就没有存在意义（也避免了"开发机不该连内网地址"这种无谓的越界）。
  stepHead(4, "本地比对已装环境（pip --dry-run --no-index，不联网，约 10 秒）")
  if (!(await pipSupportsDryRun(python))) {
    stepInfo("✗ 目标环境的 pip 不支持 --dry-run（需 >= 22.2）")
    warn("  先升级 pip：python -m pip install -U pip")
    writeState(repoRoot, { status: "no-pip", checkedAt: now, requirementsStamp: reqStamp })
    update({ phase: "no-pip" })
    return { status: "no-pip", missing: [] }
  }

  const first = await checkDeps(reqFile, python, repoRoot)
  if (first.ok) {
    writeState(repoRoot, {
      status: "ok",
      python: python.executable,
      version: python.version,
      indexUrl: config.indexUrl,
      requirementsStamp: reqStamp,
      checkedAt: now,
    })
    stepInfo(`✓ ${counts()} 全部满足，无需下载`)
    warn(`✓ 依赖已在位（${counts()}，全部满足 · CPython ${python.version}，无需下载）`)
    warn(`  来源 ${config.indexUrl}（来自${config.origin}，本次未下载）`)
    warn(`  解释器 ${python.executable}`)
    update({ phase: "ok", version: python.version, note: "", installedNow: false })
    return { status: "ok", missing: [] }
  }

  stepInfo(`✗ 未满足 ${first.missing.length || "?"} 个：${first.missing.slice(0, 6).join(", ") || "版本不符"}`)
  verbose(`未满足明细: ${first.missing.join(", ") || first.detail}`)

  // ---- 5/7 探测 server -----------------------------------------------------
  stepHead(5, `探测内网 server 连通性：${config.indexUrl}`)
  let probe = { ok: false, fatal: false, lines: [] }
  if (process.env.PI_DEPS_NO_PROBE === "1") {
    stepInfo("? PI_DEPS_NO_PROBE=1，跳过探测")
  } else {
    probe = await probeServer(config.indexUrl)
    for (const line of probe.lines) stepInfo(line)
  }

  // 第 4 步已确定拿不到包（拒绝连接 / 解析失败 / 索引页 404）：不再让 pip 白等一下。
  if (probe.fatal) {
    const detail = (probe.lines[probe.lines.length - 1] || "").replace(/^✗\s*/, "")
    warn("✗ 第 5 步已确定这台机器取不到包，跳过下载安装（不再等 pip 超时重试）")
    warn(`  原因：${detail}`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      indexUrl: config.indexUrl,
      requirementsStamp: reqStamp,
      checkedAt: now,
      detail,
    })
    update({
      phase: "failed",
      step: 5,
      stepLabel: "探测内网 server 连通性",
      detail,
      hint: `先解决到 ${config.indexUrl} 的连通性；修好后重开 pi`,
    })
    return { status: "failed", missing: first.missing }
  }

  // ---- 6/7 下载安装 --------------------------------------------------------
  stepHead(6, `从内网 server 下载安装（${first.missing.length || "?"} 个未满足）`)
  update({ phase: "installing", missing: first.missing })
  const install = await installDeps(repoRoot, python, config, first.missing, reqFile)
  if (!install.ok) {
    stepInfo(`✗ ${install.detail}`)
    if (install.hint) warn(`  ${install.hint}`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      indexUrl: config.indexUrl,
      requirementsStamp: reqStamp,
      checkedAt: now,
      detail: install.detail,
    })
    update({
      phase: "failed",
      step: 6,
      stepLabel: "从内网 server 下载安装",
      detail: install.detail,
      // 完整重试命令留在 stderr（终端/日志里能复制），widget 只放一句怎么查
      hint: "完整原因与重试命令见终端 [deps] 输出，或跑 PI_DEPS_CLI=1 排查",
    })
    return { status: "failed", missing: first.missing }
  }
  stepInfo(`✓ pip 退出码 0，耗时 ${install.elapsed}s`)
  if (install.installed) stepInfo(`  本次装入：${install.installed}`)

  // ---- 7/7 复检 ------------------------------------------------------------
  stepHead(7, "复检（再跑一次本地比对）")
  const second = await checkDeps(reqFile, python, repoRoot)
  if (!second.ok) {
    const detail = `下载完成但复检仍不满足：${second.missing.join(", ") || second.detail}`
    stepInfo(`✗ ${detail}`)
    warn(`  说明内网 server 上的包版本与 ${CONFIG.requirementsFile} 不一致，或传递依赖被裁掉了`)
    writeState(repoRoot, {
      status: "failed",
      python: python.executable,
      version: python.version,
      indexUrl: config.indexUrl,
      requirementsStamp: reqStamp,
      checkedAt: now,
      detail: second.detail,
    })
    update({
      phase: "failed",
      detail,
      hint: `核对内网 server 上的包版本与 ${CONFIG.requirementsFile} 是否一致`,
    })
    return { status: "failed", missing: second.missing }
  }

  // 装完必须重算闭包：安装前那次数出来的是"当时已装的部分"，缺的包不在里面，
  // 直接拿它显示会报出偏小的总数（比如 7 个），正是用户之前吐槽的那个错。
  const afterTree = await collectDepTree(repoRoot, python, reqFile)
  if (afterTree) {
    update({
      directCount: afterTree.direct.length,
      transitiveCount: afterTree.transitive.length,
      total: afterTree.total,
      pkgs: [...afterTree.direct, ...afterTree.transitive],
    })
  }
  stepInfo(`✓ ${counts()} 全部满足`)

  writeState(repoRoot, {
    status: "ok",
    python: python.executable,
    version: python.version,
    indexUrl: config.indexUrl,
    requirementsStamp: reqStamp,
    checkedAt: now,
  })
  const totalSeconds = Math.round((Date.now() - startedAt) / 1000)
  warn(`✓ 依赖安装完成（耗时 ${install.elapsed}s，${counts()} 全部满足）`)
  warn(`  来源 ${config.indexUrl} · 解释器 ${python.executable}`)
  update({
    phase: "ok",
    version: python.version,
    note: `（本次装入 ${install.elapsed}s）`,
    installedNow: true,
  })
  return { status: "installed", missing: [], elapsed: totalSeconds }
}

async function bootstrap() {
  if (process.env.PI_DEPS_SKIP === "1") {
    // 这一条必须能看见：应急开关静默生效时，用户唯一看到的就是"扩展毫无反应"，
    // 会误判成扩展没加载。所以 skip 也要留一行 stderr + 一行 widget。
    warn("PI_DEPS_SKIP=1，已显式跳过依赖检查（取消该环境变量后重开 pi 才会检查）")
    update({ phase: "skipped", note: "PI_DEPS_SKIP=1 显式跳过" })
    return { status: "skipped", missing: [] }
  }
  const repoRoot = findRepoRoot(process.cwd())
  if (!repoRoot) {
    // cwd 不在项目内是隐蔽的失败：什么都不做 → 首页一片干净，
    // 和"扩展没加载"完全同徵。所以把 cwd 一起报出来，让原因当场可见。
    warn(`当前目录不在项目内（向上 3 层没找到 requirements.txt / .pi 目录），跳过依赖检查`)
    warn(`  cwd = ${process.cwd()}`)
    warn("  从项目根目录启动 pi，或把 pi 的工作目录切到项目根")
    update({ phase: "skipped", note: `cwd 不在项目内：${process.cwd()}` })
    return { status: "skipped", missing: [] }
  }
  const config = readConfig()
  // 清单是依赖的唯一真相，没有它就没有可检查的东西。这算配置缺失而不是"依赖已就位"，
  // 所以显式报出来并跳过——静默跳过会被误读成"检查过了没问题"。
  const reqFile = reqFileOf(repoRoot)
  if (!fs.existsSync(reqFile)) {
    warn(`✗ 依赖清单不存在：${reqFile}`)
    warn(`  本脚本只从 ${CONFIG.requirementsFile} 读取依赖（CONFIG.requirementsFile 指定文件名）`)
    update({ phase: "skipped", note: `依赖清单不存在：${CONFIG.requirementsFile}` })
    return { status: "skipped", missing: [] }
  }
  gateEnabled = config.gate
  return ensureDeps(repoRoot)
}

// ---------------------------------------------------------------------------
// 依赖未就位时拦住用户指令
//
// 为什么必须拦：依赖缺了 Agent 照样能起来，但生成 PPT 时才会 `import matplotlib` 崩掉，
// 用户看到的是"莫名其妙失败"，还得自己回溯到依赖问题。在这里挡住，原因当场可见。
//
// 用什么拦：`input` 事件返回 { action: "handled" } 会跳过整个 agent 运行。
//   - 走 `input` 而不是 `before_agent_start`：前者在 skill/template 展开之前，
//     拦得住 `/skill:xxx` 这类输入；后者已经在跑 turn 了。
//   - `/` 开头的扩展命令不经过 input（pi 先查命令，命中就跳过 input），
//     所以 /help /exit 这类照样能用，用户不会被锁死。
//   - `!` bash 命令走 user_bash，不经过 agent，也就不拦——它不依赖 Python 包。
//   - 扩展自己发的消息（source=extension）放行：拦了会打断别的扩展的流程，
//     而且用户要拦的是"人敲进来的自然语言"。
// ---------------------------------------------------------------------------
let gateEnabled = true

function isReady() {
  return snapshot.phase === "ok" || snapshot.phase === "skipped"
}

function blockReason() {
  const s = snapshot
  switch (s.phase) {
    case "checking":
      return `依赖正在检索（${counts()}），完成后才能接收指令。进度见输入框下方的状态行`
    case "installing":
      return `依赖正在从内网 server 下载（缺 ${s.missing.length || "?"} 个），装完才能接收指令。进度见输入框下方`
    case "failed":
      return `依赖未就位：${s.detail}。修好后重启 pi；完整原因见终端 [deps] 输出`
    case "no-python":
      return `未找到 CPython ${s.wantVersion || ""} 解释器，依赖没装。装好 3.12 后重启 pi（py -0 查看现有版本）`
    case "no-pip":
      return "pip 太老（需 >= 22.2），无法检查/安装依赖。升级 pip 后重启 pi"
    default:
      return "依赖还没就位，暂时不能接收指令。进度见输入框下方的状态行"
  }
}

// ---------------------------------------------------------------------------
// pi 扩展入口
//
// 工厂**同步**返回：pi 才不会在启动阶段被卡住（启动阶段 TUI 还没渲染，卡在那里
// 用户只看到一片空白，比慢更难忍）。检查/安装在后台跑，进度通过 session_start
// 里的 widget 显示——这是唯一能在首页看见的通道。
// ---------------------------------------------------------------------------
export default function depsBootstrap(pi) {
  // 这一行是"扩展是否真的被加载"的唯一硬证据：模块加载失败时 factory 压根不会被调用，
  // 任何内部日志都不会出现。所以先证明自己活着，再证明依赖检查跑没跑。
  // （TUI 里它会被重绘盖掉，但 pi --help / CLI 模式 / 重定向下都在。）
  warn(`扩展已加载，cwd=${process.cwd()}`)

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
        // footer 常驻一条：依赖没就位时不光要在 widget 里说，也要在状态栏一直挂着，
        // 否则 widget 滚上去就看不见了
        ctx.ui.setStatus("deps", isReady() ? undefined : "依赖未就位")
      } catch {
        /* 会话已被替换，ctx 失效 */
      }
    }

    redraw = paint
    paint()
    // 最多等 90 秒：装依赖可能几分钟，不能把 pi 卡死；超时先放行，装完 widget 自己会刷新
    await Promise.race([done, new Promise((resolve) => setTimeout(resolve, 90_000))])
    paint()
  })

  pi.on("input", async (event, ctx) => {
    if (!event || event.source === "extension") return { action: "continue" }
    if (!gateEnabled || isReady()) return { action: "continue" }

    try {
      ctx?.ui?.notify?.(blockReason(), "warning")
    } catch {
      /* 没有 UI 通道（print/json 模式）时静默，靠退出码与 stderr 说话 */
    }
    return { action: "handled" }
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
    warn("! 当前目录及上 3 层没有 requirements.txt / .pi 目录，不在项目内，无需预检")
    process.exit(0)
  }
  if (!fs.existsSync(reqFileOf(repoRoot))) {
    warn(`✗ 依赖清单不存在：${reqFileOf(repoRoot)}，无法预检`)
    process.exit(1)
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
