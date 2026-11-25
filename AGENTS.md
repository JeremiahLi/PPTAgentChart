# AGENTS.md

本仓库是**基于 pi 的 PPT 制作 Agent**（内网部署、不可联网的私有模型）：把文档/思路做成
高信息密度、图表丰富的 .pptx。
出图唯一通道是 `engine/`（PPT-DSL 结构化 JSON → pptx），**禁止手写 python-pptx / OOXML**。

> **代码分层（2026-11-25）**：可执行代码是仓库根的 `engine/` 包，不再是 `.pi/` 下的散脚本。
>
> | 目录 | 装什么 |
> |---|---|
> | `engine/dsl/` | 契约层：`ppt_dsl.py`（19 版式 schema）、`briefs.py`（方向规约加载） |
> | `engine/plan/` | 规划层：`doc_analyzer` → `planner` → `outline` / `plan` / `budgets` |
> | `engine/render/` | 渲染层：`theme` / `charts` / `icons` / `iconlib` / `layouts` / `merge` / `render` / `speaker_notes` |
> | `engine/qc/` | 检查层：`qc.py`（几何/密度/红线/覆盖度）、`fontcheck.py` |
> | `engine/ingest/` | 摄入层：`docx_in.py`（去修订 + 抽图） |
> | `engine/build.py` | 编排层：merge → qc → render → fontcheck 一条命令 |
> | `engine/paths.py` | 资源根路径唯一来源（`.pi/` 下 briefs / assets / work / out） |
>
> **迁移期的转发桩已拆除**（2026-11）。`.pi/renderer/` 目录不再存在，调用一律走
> `python -m engine.<子包>.<模块>`，例如 `python -m engine.build …`。旧命令
> `python .pi/renderer/build.py …` **现在会直接报文件不存在**，不要再照抄旧文档。
>
> **`.pi/` 下不放可执行代码**：一旦 `.pi` 进了 `sys.path`，`.pi/pptx.py` 会把第三方
> `pptx` 库顶掉（`'pptx' is not a package`），而且只在特定 import 顺序下才炸。
> 这也是 `engine/` 必须待在仓库根、而不是 `.pi/` 下的原因。

> **目录结构（2026-10 调整）**：配置、技能、测试、数据在 **`.pi/` 下**（`assets/`、`tests/`、
> `work/`、`out/`、`docs/`、`wheels/`、`briefs/`、`skills/`、`prompts/`）；
> 仓库根放 `AGENTS.md`、`requirements.txt`、`setup.bat` 和 `engine/`。
> **所有命令一律从仓库根执行，路径必须带 `.pi/` 前缀**：
> `python .pi/pptx.py build -o .pi/out/x.pptx`
> （一条命令 = 合并 + 质检 + 渲染 + 字体自检；拆开跑用 `merge.py` + `render.py --qc`）。
> 文件位置不再靠 `dirname(dirname(__file__))` 反推——`engine/paths.py` 是唯一来源，
> 模块再往子目录搬也不会把 `.pi` 算错一层。

> **自动提交约定（2026-11-05 起，强制执行）**：每次对代码/文档/规范做了一轮优化调整，
> 任务收尾时必须当场 `git add` + `git commit`（commit 信息一律**简体中文**、概括本次改动）+
> `git push origin main`，不要等用户提醒。禁止 `--force`；`.gitignore` 已排除的产物不入库。

## 唯一入口：`.pi/pptx.py`

以前"分析 → 大纲 → 预算 → 写页 → 出图"这条顺序**只写在 markdown 里**，没有任何东西保证
被按序执行：跳过 `budgets.py` 直接写 slides，`build.py` 照样出图，直到长度校验报错才返工。
现在顺序在代码里，**跳一步必然失败**：

```bash
python .pi/pptx.py                       # 看状态（不带子命令 = status）
python .pi/pptx.py analyze 材料.docx      # Word 预处理 + 分析表骨架
python .pi/pptx.py outline --title "主题"  # 分析表 → 结构化大纲
python .pi/pptx.py budget                # 大纲 → 逐页字数预算
python .pi/pptx.py build -o .pi/out/x.pptx   # 合并 + 质检 + 渲染 + 字体自检
python .pi/pptx.py check                 # 只质检不出文件
```

三条契约（与 `build.py` 同口径，不另立一套）：

| 项 | 约定 |
|---|---|
| **成败** | 每条命令末行都是 `RESULT {json}`（`status` 也有），看 `ok` + `errors[].code`；`code` 是稳定枚举（`E_STAGE_MISSING` / `E_STAGE_STALE` / `E_QC_ERROR` …），中文只在 `message` / `hint` 里 |
| **退出码** | 业务结果（校验拦截、上游缺失、过期、warning 未归零）**恒 0**；只有未捕获异常退 1，参数拼错退 2（这类错误不该被静默） |
| **状态** | `.pi/work/state.json`：每阶段记输入的**内容指纹**（sha256 前 12 位，不是 mtime）。改过源文件，下游自动判过期并拦住 |

门禁分两级：**requires**（缺了就报错，如大纲缺分析表）、**advises**（缺了只告警，如没跑
budget 就出图）。`--force` 可硬闯，`RESULT` 里留 `forced: true` 可追溯。
`untracked` 是迁移态：产物在但无指纹记录（新入口上线前跑的老流程），门禁放行但告警，
重跑一次就收录。

`.pi/pptx.py` 只是适配层：它不复制任何判定逻辑（字数上限、几何容量、字体口径仍在
`engine/` 里），阶段与模块的对应只有 `STAGE_MODULES` 一张表。入口层回归：
`python .pi/tests/entry.py`（纯标准库，没装 python-pptx 也能跑）。

## 谁负责什么

- **三个用户维度 + 默认值**（`build.py --brief / --style / --speed`）：
  **类型 = 研讨型 · 全量详实**、**风格 = 经典藏青**、**速度 = 慢**。
  用户没提就按默认走、提了才切换，**不要逐项追问**（每问一次一轮往返）。
  风格一项已绑定卡片长相，用户不点名就不用传 `--card`。详见 `ppt-entry` 第 0 步。
- **触发入口 + 工作流**：`.pi/skills/ppt-entry.md` + `/ppt <文档>`（`.pi/prompts/ppt.md`）。
  需求澄清 → **输入分析（`.pi/work/analysis.md`：骨架由 `plan.py --emit-analysis` 生成，
  人或模型只补结论/缺口两列）** → 大纲（唯一停下来让用户过目的环节）→
  **成批生成（一个文件 5-8 页）** → `build.py` 一条命令出图 → 回炉 → 交付，
  以及"标题即结论 / 图示优先 / 信息密度 / 结论段硬指标"铁律，都在 `ppt-entry` + `ppt-common` 里。
- **公共规范 + few-shot 范例**：`.pi/skills/ppt-common/SKILL.md`（出图通道 / 19 版式契约 /
  Word 读入 / 图标 / 通用铁律）+ `examples-all.json`（**19 版式合集，开工时一次读完**；
  要抠某个版式的写法再单独读 `examples/NN-*.json`）。
  **范例是最低质量线**，密度只能更高不能更低（范例本身也要过 `--qc` 的 summary 硬指标）。
  别每生成一页就去 Read 一个范例文件 —— 那是 40 页 deck 白搭 40 次读盘与上下文。
- **两个写作方向技能**：`.pi/skills/ppt-seminar/SKILL.md`（研讨型）、
  `.pi/skills/ppt-briefing/SKILL.md`（汇报型）。**二选一**，都引用 ppt-common。
- **运行环境**：内网部署、不可联网的私有模型，全程中文（简体）。
  `~/.pi/agent/AGENTS.md` 只放对所有项目成立的通用约束，PPT 专属规则全在本仓库。
- **字体**：中文微软雅黑 + 西文 Times New Roman，只用 Windows 自带字体，
  **仓库不带字体包**（不要放 `fonts/`、`*.ttf`）；口径与实现见 `engine/render/theme.py`
  与 `.pi/docs/pitfalls.md` 第 24 条。交付前跑 `python .pi/tests/font_audit.py .pi/out/xxx.pptx`
  自检（逐段核对 latin+ea 字体对 + `lang`/`endParaRPr` 语言标记 + 母版主题 + 编码）；
  要让**英文与数字也用雅黑**，
  渲染前设 `$env:PPT_FONT_LATIN="微软雅黑"` 即可，不用改代码。

## 依赖与离线部署

**目标解释器 = CPython 3.12 / win_amd64。** `.pi/wheels/py312/` 里的二进制全是 cp312，
3.13 / 3.14 装不上也不该装（本机开发用的 3.14 与交付无关）。

发布结构 = `requirements.txt` 的 **7 个直接依赖** + **13 个传递依赖** = 20 个 wheel / 38 MB。
传递依赖不写进 requirements.txt，pip 顺着依赖树从同一目录自己取。

三层保障，按触发顺序：

| 时机 | 入口 | 行为 |
|---|---|---|
| pi 启动前 | `.pi/extensions/deps-bootstrap.ts`（同步工厂，长任务丢后台，进度走 widget） | **按 7 步走，每步都有日志**：①定位解释器 ②统计依赖闭包 ③读来源地址 ④本地比对（`--no-index`，不联网）⑤探测 server 连通性（TCP + 索引页）⑥从内网 server 下载（pip 输出逐块实时转发）⑦复检。依赖本来就齐时只走 1-4 步，**根本不碰网络**。结果缓存 24h |
| 首次部署 / 环境坏了 | 仓库根 `setup.bat` | 优先 `py -3.12`；有 `.pi/wheels/py3XY` 就离线装，没有才走内网镜像 |
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

## 引擎模块（`engine/`）

> 全部是**可导入的包模块**，从仓库根用 `python -m …` 运行。
> 分层：`engine/dsl/`（schema、briefs）· `engine/plan/`（分析/规划/预算）·
> `engine/render/`（主题/图表/图标/版式/合并/渲染/备注）· `engine/qc/`（质检/字体）·
> `engine/ingest/`（docx 预处理）· `engine/build.py`（编排）。
> **不在 `.pi/` 下**：`.pi` 进 `sys.path` 会把第三方 `pptx` 顶掉，见开头。

| 模块 | 作用 |
|---|---|
| `python -m engine.dsl.ppt_dsl <deck.json>` | 19 版式的 pydantic schema（字段名 + 长度上限），校验入口 |
| `python -m engine.ingest.docx_in <输入.docx> -o .pi/work/input.md [--media .pi/work/media]` | **Word 读入预处理**：去修订（接受插入 / 丢弃删除与批注）+ 图片原样抽盘（EMF/WMF 不转格式）+ 图位 markdown 与清单 |
| `python -m engine.build <slides_dir> -o <out.pptx> [--brief] [--style] [--speed] [--card] [--source] [--check] [--only 3,7]` | **一条命令出图**：合并 + 质检 + 渲染 + 字体自检，一份紧凑报告。**退出码恒 0**——校验拦截（slides 不合法 / 质检 error / 字体不过 / warning 未归零）属正常业务逻辑，不是运行失败；成败看末行 `RESULT {json}`（`{"ok": …, "reason": …}`）。日常和交稿都走这条 |
| `python -m engine.render.render <deck.json> -o <out.pptx> --qc [--brief 方向] [--style 配色] [--card 卡片] [--source 输入分析]` | 渲染 + 质检（error 会阻止渲染）；已有 deck.json 时用 |
| `python -m engine.render.merge <slides_dir> <out.json> [--brief 方向] [--style 配色] [--card 卡片]` | 合并 slides 为 deck：**一个文件可装 1 页，也可装一页数组（5-8 页/文件，推荐）**；校验失败定位到「文件 + 第几项 + 字段 + 当前字数 vs 上限」；自动纠偏 `type`→`slide_type`、两种 points 结构混淆（stderr 有提示）；读文件兼容 UTF-8 BOM；**校验拦截退出码恒 0**，成败看末行 `RESULT {json}`（与 build.py / render.py 同口径） |
| `python -m engine.plan.budgets <outline.json> -o .pi/work/budgets.md` | **字数预算表**：大纲定稿后、写批次前跑，逐页下发动态容量（写作目标 = min(静态上限, 几何容量)，与 schema 校验**同一套几何**，照预算写必过校验）；`--all` 出全版式参考表；`--selfcheck` 自证"照预算写必过校验"，几何改动后重跑它当回归。预算表是"校验报数"的前置版——动态数字不在 SKILL.md 里背 |
| `python -m engine.qc.fontcheck <pptx>` | 字体自检的唯一实现（`build.py` 同进程调用；`.pi/tests/font_audit.py` 是薄壳 CLI） |
| `engine/qc/qc.py` | 几何容量 / 密度 / 结构检查 + 方向红线 + 输入覆盖度 |
| `.pi/briefs/` + `briefs.py` | **写作方向**（研讨型 / 汇报型）的取舍规约与加载，见下 |
| `engine/render/theme.py` | 外观的两个白名单维度（2026-10 扩容）：**配色 `style` 5 套**（`style-harness` 默认藏青 / `blue-business` / `bright` / `forest` / `violet-tech`）+ **卡片长相 `card` 5 种**（`flat` 默认 / `outline` / `tinted` / `accent-left` / `accent-top`），`STYLE_CARD_PRESET` 把每套配色绑定一种卡片（用户只选"风格"一项，`--card` 可覆盖）；`apply_style()` / `apply_card_style()` **原位**变更、未登记名确定性回退默认；**字体口径**（中文微软雅黑 + 西文 Times New Roman，只用系统字体、不随包分发）与 `apply_theme_fonts()`。底图与字体不随主题变，卡片长相不许动文字几何（见 pitfalls 第 27 条） |
| `engine/render/iconlib.py` | SVG → custGeom 的图标引擎（解析规则、检索 CLI） |
| `.pi/assets/icons/` | 图标库：**tabler 单库 4620 个**（outline 3646 + filled 974）+ 中文别名 `aliases.json`，见其 `README.md` |

## 写作方向（研讨型 / 汇报型）

两套东西，一对一行：

| | 研讨型 `seminar` | 汇报型 `briefing` |
|---|---|---|
| 技能（给人/AI 读的规范） | `.pi/skills/ppt-seminar/SKILL.md` | `.pi/skills/ppt-briefing/SKILL.md` |
| 规约（给 `--qc` 读的红线） | `.pi/briefs/seminar.yaml` | `.pi/briefs/briefing.yaml` |
| 受众 | 会逐条讨论的同事 | 只看这一份的决策者 |
| 取舍 | 资料论点论据**全部上页不丢弃** | 每页一个主张，可抽象可砍 |
| 页数 | **不限**，由资料量决定；预计 >50 页先问用户 | **不限**（2026-09 起与研讨型一致）；预计 >30 页先问用户 |
| 每页论据 | 不限 | ≤3 |
| 图表 | 尽量每页都有，图型至少 3 种 | 够用即可 |
| 标题 | 判断句 | 判断句 + 禁"情况介绍/相关说明/概述" |
| 缺口 | 必须显式占位 | 可选 |

**默认组合 = 研讨型 + 经典藏青（`style-harness` + `flat`）+ 慢档**；
速度档只改"跑几个回合"（大纲给不给用户过目、warning 要不要归零），
**两档跑同一套检查**，快档一项不少。

两个方向**共用同一套版式和配色**，差别不在长相，在怎么取舍原文，所以规约是策略数据
（`.pi/briefs/*.yaml`），不是渲染配置。`deck.json` 写 `"brief"` 字段
（合并时 `merge.py --brief`，渲染时 `render.py --brief` 可覆盖）。`--qc` 按该方向的
`check` 段查红线，**只出 warning**。

**改方向规约时，SKILL.md 和 YAML 要一起改**——`.pi/tests/briefs.py` 会校验 YAML 里的
`skill:` 指向的文件存在，但内容对不对得上得靠人。

新增方向 = 往 `.pi/briefs/` 放一个 yaml + 在 `.pi/skills/` 放一个 SKILL.md，
**零代码改动**。回归：`python .pi/tests/briefs.py`。

注意：`src/ppt/styles/*.yaml` 是旧流水线的同名概念，与这里无关，别改串了。
`deck.style` / `deck.card` 是**外观**（配色 / 卡片长相），与写作方向（`brief`）无关，
三个字段正交、别互相牵连，也别跟着改名字。

## 已踩过的坑（改了别再踩回去）→ `.pi/docs/pitfalls.md`

**28 条坑位 + 一份「关键长度陷阱」摘要全部搬到 [`.pi/docs/pitfalls.md`](.pi/docs/pitfalls.md)**，
每条都写清"错在哪、现在怎么做的、为什么别改回去"。三条约定：

- **改渲染器 / 改规则之前**先扫一遍那份清单，尤其别把已经修好的坑改回去。
- **生成 PPT 时不需要读它** —— 生成时只读 `ppt-entry` → `ppt-common` → 方向技能 →
  `examples-all.json`（这两节以前写在 AGENTS.md 里，等于每次生成都白读 170 行）。
- 规则分三层、各自只放该放的东西：**判据（schema 上限 / 几何函数）在代码里、
  机检阈值在 `briefs/*.yaml` 的 `check` 段、写作规范在 `.pi/skills/*.md`** ——
  同一句话不要在三处都写一遍（历史上"12 版式""60~140 字"就是这样漂移的）。

## 已删除的旧路径（2026-09 清理）

`src/ppt/`、`SPEC/`、`tools/run_ppt.py`、`tools/mock_model.py` 是**旧的确定性流水线**
（L2 直连 GLM 的 HTTP 端点）。这些路径**已整体删除**（2026-09 清理）：
`src/ppt/`、`SPEC/`、`tools/`（run_ppt、mock_model、offline_demo、template_scan、
install_template_assets）、`demo/`、`run/`。别再按旧文档去找它们。
