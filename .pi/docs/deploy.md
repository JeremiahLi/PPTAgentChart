# 部署与离线依赖

> 本文是 `AGENTS.md` 拆出来的长文（2026-11-25）。原因是它每次对话都进上下文，
> 而生成 PPT 用不到这里任何一段；部署只在装机、换机、补包时读。
> **命令口径与 `AGENTS.md` 一致**：入口是 `.pi/cli.py`，从仓库根执行。
>
> 一句话现状：代码整包在 `.pi/engine/`，仓库根只剩 `AGENTS.md` + `requirements.txt`；
> 依赖只有一条安装通路（内网 server），入口 `.pi/setup.bat`。

## 依赖与离线部署

**目标解释器 = CPython 3.12 / win_amd64。** `.pi/wheels/py312/` 里的二进制全是 cp312，
3.13 / 3.14 装不上也不该装（本机开发用的 3.14 与交付无关）。

发布结构 = `requirements.txt` 的 **7 个直接依赖** + **13 个传递依赖** = 20 个 wheel / 38 MB。
传递依赖不写进 requirements.txt，pip 顺着依赖树从同一目录自己取。

三层保障，按触发顺序：

| 时机 | 入口 | 行为 |
|---|---|---|
| pi 启动前 | `.pi/extensions/deps-bootstrap.ts`（同步工厂，长任务丢后台，进度走 widget） | **按 7 步走，每步都有日志**：①定位解释器 ②统计依赖闭包 ③读来源地址 ④本地比对（`--no-index`，不联网）⑤探测 server 连通性（TCP + 索引页）⑥从内网 server 下载（pip 输出逐块实时转发）⑦复检。依赖本来就齐时只走 1-4 步，**根本不碰网络**。结果缓存 24h |
| 首次部署 / 环境坏了 | `.pi/setup.bat`（从仓库根执行） | 优先 `py -3.12`；有 `.pi/wheels/py3XY` 就离线装，没有才走内网镜像 |
| 换机 / 补包 | `.pi/wheels/README.md` | `pip download` 到对应版本子目录 |

> **依赖只有一条安装通路：内网 server（扩展文件顶部 `CONFIG.indexUrl`，固定值）。**
> 自举**不回退公网 PyPI、也不读 `.pi/wheels` 本地目录**——离线机连不上公网，回退只会白等；
> 本地包是另一条要维护的路径，已从自举里删除。地址写死在 `CONFIG` 里，**既没有配置文件也没有
> 环境变量能改它**，永不偷偷回退公网 PyPI。
>
> **开发机到不了那个内网 IP，这是正常状态**：源码在这边写、拷到离线 PC 上跑，只有那台机器能连通
> server。所以别在开发机上试连 `172.30.39.34`（只会白等超时），也别指望开发机能验证下载路径——
> 下载链路的验证只能在离线 PC 上做。开发机上能验证的只有：检查分支（`--no-index`，纯本地）、
> 以及失败原因分类（用回环地址 `127.0.0.1:9` 造连接拒绝）。

**两套输出通道，缺一不可**：

| 通道 | 内容 | 给谁看 |
|---|---|---|
| **TUI widget**（`deps-status`，输入框下方 `belowEditor`） | ⏳ `3/7 探测内网 server 连通性` + 该步结论 / ✓ 已就位时给**四行**：多少个包、**来源地址**、**解释器路径**、包清单 / ✗ 未就位时给"停在第几步 + 原因 + 怎么办" | **用户**。pi 进 interactive 模式会切 alternate screen 重绘，stderr 全被盖掉——2026-11-18 就是因此首页一片干净，用户压根不知道依赖检没检查 |
| stderr `[deps] ...` | 同上内容的完整版 + 手动重试命令 | CLI / print 模式 / 日志 / 排查 |

依赖不是 `ok` 时，footer 常驻一行 `ctx.ui.setStatus("deps", "依赖未就位")`，每次重绘同步更新。

**依赖没就位时用户指令会被挡住**（`pi.on("input")` 返回 `{ action: "handled" }`）：

- 拦的是 `checking` / `installing` / `failed` / `no-python` / `no-pip` 五种状态；
  `ok` 与 `skipped`（含 `PI_DEPS_SKIP=1`）放行。
- 拦下来时用 `ctx.ui.notify(原因, "warning")` 说明为什么，不静默丢消息。
- 为什么拦：依赖缺了 Agent 照样能起，但生成 PPT 时才 `import matplotlib` 崩，
  用户看到的是莫名失败。挡在输入处，原因当场可见。
- **不会把人锁死**：`/` 开头的扩展命令不经过 `input` 事件（pi 先查命令），
  `/help`、`/exit` 照常可用；`!` bash 命令走 `user_bash`，也不拦。
- 扩展自己发的消息（`event.source === "extension"`）放行，免得打断别的扩展流程。

**启动时终端/stderr 会看到的分步日志**（前缀 `[deps]`，序号 1-7）：

```
1/7 定位 CPython 3.12 解释器
    ✓ C://...//pythoncore-3.12-64//python.exe（3.12）
2/7 统计依赖闭包（requirements.txt 7 条）
    ✓ 统计完成：20 个包（7 直接 + 13 传递）
3/7 读取依赖来源地址
    ✓ http://172.30.39.34:18080/simple（来自扩展内置配置）
4/7 本地比对已装环境（pip --dry-run --no-index，不联网，约 10 秒）
    ✓ 20 个包：7 直接 + 13 传递 全部满足，无需下载
✓ 依赖已在位（20 个包：7 直接 + 13 传递，全部满足 · CPython 3.12，无需下载）
  来源 http://172.30.39.34:18080/simple（来自扩展内置配置，本次未下载）
  解释器 C://...//pythoncore-3.12-64//python.exe
```

依赖缺了才会走到 5-7 步，第 5 步把"地址通不通"说透：

```
4/7 本地比对已装环境（pip --dry-run --no-index，不联网，约 10 秒）
    ✗ 未满足 3 个：numpy, pillow, matplotlib
5/7 探测内网 server 连通性：http://172.30.39.34:18080/simple
    ✓ TCP 172.30.39.34:18080 可达（12ms）
    ✓ 索引页 /simple 返回 200（18ms）
6/7 从内网 server 下载安装（3 个未满足）
    下载中，pip 原始输出如下（逐包进度实时刷新）...
    ✓ pip 退出码 0，耗时 42s
      本次装入：numpy-2.4.6  pillow-12.3.0  matplotlib-3.11.2  …
7/7 复检（再跑一次本地比对）
    ✓ 20 个包：7 直接 + 13 传递 全部满足
```

第 5 步的失败会直接给出结论，**不再让 pip 白等 30s×2 重试**：

```
5/7 探测内网 server 连通性：http://172.30.39.34:18080/simple
    ✗ TCP 172.30.39.34:18080 不可达（3ms）→ 端口拒绝连接：该端口没有服务在听，或防火墙显式拒绝
✗ 第 5 步已确定这台机器取不到包，跳过下载安装（不再等 pip 超时重试）
```

探测分两层：**TCP**（能不能握手）→ **索引页 GET**（`/simple` 返回什么）。
只有确定性失败（拒绝连接 / 解析失败 / 路由不可达 / 索引页 404、401、403）才跳过 pip；
超时这种"可能被防火墙拖住"的情况仍然让 pip 试一次——探测失败 ≠ 装了会失败，
代理和只放行 HTTP 的防火墙都会让 TCP 探测失真。

**口径是全部依赖，不是 requirements.txt 那 7 行**：`requirements.txt` 只写 7 个直接依赖，
pip 会顺着依赖树再装 13 个传递依赖。闭包由 `collectDepTree` 用 `importlib.metadata`
在本地递归算出（不联网、约 0.3 秒，装完会重算），**但"是否满足"仍由 pip 的 `--dry-run`
裁定**——分工别搞混：闭包只管显示，判定权威只有 pip。

失败原因是 pip 原始报错翻成的人话，分类顺序是"先判没连上、再判连上了没这个包"——
两种情形 pip 都会吐 `No matching distribution found`，顺序反了就会把断网误报成缺包。
覆盖：拒绝连接 / 超时 / 主机名解析失败 / server 上没这个包 / 证书 / 无写权限 / 磁盘满 / 未配地址。

缓存命中那行故意写明"多久前验证过"，是为了不让人误读成"刚才查过"——真检查一次要 11 秒
（`pip --dry-run --no-index` 实测），每次启动都跑不可接受。要立刻真查：`PI_DEPS_FORCE=1 pi`。

**判定方式**：不猜 import 名、不自己比版本——直接跑
`pip install --dry-run --no-index -r requirements.txt`，退出码 0 = 全部满足（含版本约束）。
`python-pptx→pptx`、`Pillow→PIL`、`PyYAML→yaml` 这类映射一概不用维护，
`numpy==2.4.6` 是否真装对也由 pip 权威实现负责。

## 单文件脚本：配置内联，清单只读 requirements.txt

`.pi/extensions/deps-bootstrap.ts` **不绑定本项目、不读任何外部配置文件**：一个 .ts 丢进任意
pi 项目的 `.pi/extensions/`（放 `~/.pi/agent/extensions/` 则全局生效）就能工作。

分工是明确的两半：

- **配置内联**：文件顶部 `CONFIG` 常量（约第 100 行）承载 `python` / `indexUrl` /
  `timeoutMinutes` / `gate` / `requirementsFile`。要换地址、换解释器版本、改超时就改这里，
  没有第二份配置、也没有环境变量能改它们。
- **依赖清单外置**：只从仓库根的 `requirements.txt` 读（文件名由 `CONFIG.requirementsFile` 指定）。
  版本号属于项目内容、随版本走，抄进代码只会多一处要同步的副本——所以脚本里**不含任何版本号**。

- `.pi/deps.json` **已删除**：它曾经承载 `python` / `indexUrl` / `timeoutMinutes` /
  `blockOnMissing`，现在全部并入 `CONFIG`。环境变量 `PI_PIP_INDEX_URL`、
  `PI_DEPS_PY_VERSION`、`PI_DEPS_TIMEOUT_MIN` 也都移除了——地址"固定不变"就是这么落实的，
  不存在"某台机器上被环境变量改了却不知道"的情况。
- 仓库根**先按扩展自身位置上溯两级**认定（`<repo>/.pi/extensions/` → `<repo>`），
  且该目录下必须真有 `requirements.txt` 才采信——这样 cwd 在哪都不影响（从家目录启动也能找对）；
  把 .ts 拷到 `~/.pi/agent/extensions/` 全局用时该位置没有清单，自动退回兜底。
  兜底才是**从 cwd 向上 3 层只找 `requirements.txt`**。
  **不认 `.pi` 目录当锚点**：`~/.pi` 是 pi 的全局配置目录，认它会把 `C:\Users\me`
  （及任何 home 子目录）判成仓库根，报出误导性的"依赖清单不存在"（2026-11-18 踩过）。
- 清单不存在时**显式报错并跳过**（`✗ 依赖清单不存在：…`），不静默当作"已就位"。
- 状态缓存写在项目自己的 `.pi/work/.deps-state.json`，键含**清单内容指纹**（不是 mtime，
  免得 git checkout 让 mtime 变、内容没变却失效），不会跨项目串。
- 不想要"未就位时拦指令"？改 `CONFIG.gate: false` 或临时设 `PI_DEPS_GATE=0`。
- **强制内网 server 单源**：不回退公网 PyPI、不读 `.pi/wheels`。

本项目的两个私有约束就写在这几行里：

```ts
python: "3.12",                              // 只认 3.12，找不到就每次提示（不吃缓存）
indexUrl: "http://172.30.39.34:18080/simple", // 内网镜像，装包唯一来源
requirementsFile: "requirements.txt",        // 清单唯一真相，脚本里不存副本
```

- `indexUrl` 走 pip 标准索引路径 `/simple`（不是 `/simples`）；脚本会自动补 `--trusted-host`
  （http 源不带会被 pip 直接拒绝）。
- 日志里的"来自XXX"现在固定是 `扩展内置配置`——不存在"到底读了哪份配置"的猜谜空间了。
- 安装失败**不阻塞 pi 启动**，只打警告 + 人话原因 + 手动重试命令——起不来的 pi 没法排查问题。
- 应急开关（只有这些，不碰配置）：`PI_DEPS_SKIP=1` 跳过、`PI_DEPS_VERBOSE=1` 看判定、
  `PI_DEPS_FORCE=1` 忽略 24h 缓存强制重查、`PI_DEPS_VERIFY=1` 预检 server 能否取全
  （只查询不安装，退出码 0/1，只在离线 PC 上跑）、`PI_DEPS_NO_PROBE=1` 跳过第 5 步探测、
  `PI_DEPS_GATE=0` 不拦用户指令。
- **server 端要齐的不是 7 个包，是 20 个**：7 个直接依赖 + 13 个传递依赖（清单见
  `.pi/wheels/README.md`）。少一个传递依赖，pip 会转去找 sdist、离线机没编译链就失败。
  另外 server 得是合规的 PEP 503 索引（`/simple/<包名>/` 有文件链接），
  光把 whl 堆进目录不算——索引 404 时 pip 报的也是 `No matching distribution found`，极易误判。
- 仓库根只向上找 3 层就停：再往上是父目录的别的项目，会把依赖装错地方。
- 已知边界：`pip install` 装进目标解释器本身，**没有 venv 隔离**。多项目共用一个 Python
  会互相顶版本，靠「一个项目一个 venv」的约定规避。
- `requirements-full.txt` 已删除：里面的 pdf / plotly / markdown / svgwrite / requests 整条链
  当前代码一个都没 import——`docx_in.py` 读 Word 用的是 `zipfile` + `lxml` 直接解 OOXML，
  不需要 python-docx。要恢复某项能力，按 README 从镜像补回来，别再维护一个全量清单。
