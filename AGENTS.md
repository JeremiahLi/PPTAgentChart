# AGENTS.md

本仓库是**基于 pi 的 PPT 制作 Agent**（内网部署、不可联网的私有模型）：把文档/思路做成
高信息密度、图表丰富的 .pptx。
出图唯一通道是 `.pi/renderer/`（PPT-DSL 结构化 JSON → pptx），**禁止手写 python-pptx / OOXML**。

> **目录结构（2026-10 调整）**：工程主体 `renderer/`、`assets/`、`tests/`、`work/`、`out/`、
> `docs/`、`wheels/` 全部位于 **`.pi/` 下**；仓库根只留 `AGENTS.md`、`requirements.txt`、`setup.bat`。
> **所有命令一律从仓库根执行，路径必须带 `.pi/` 前缀**：
> `python .pi/renderer/build.py .pi/work/slides -o .pi/out/x.pptx --source .pi/work/analysis.md`
> （一条命令 = 合并 + 质检 + 渲染 + 字体自检；拆开跑用 `merge.py` + `render.py --qc`）。
> 这么做的原因只有一个——根目录清爽；代价是每条路径多一层。
> 代码侧零改动即可工作：`renderer/` 与 `tests/` 全部用 `__file__` 推父目录（`_BASE`/`REPO`），
> `layouts.resolve_image_path` 先试 `_BASE` 再试 cwd，整体平移一层后相对关系不变。

> **自动提交约定（2026-11-05 起，强制执行）**：每次对代码/文档/规范做了一轮优化调整，
> 任务收尾时必须当场 `git add` + `git commit`（commit 信息一律**简体中文**、概括本次改动）+
> `git push origin main`，不要等用户提醒。禁止 `--force`；`.gitignore` 已排除的产物不入库。

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
  **仓库不带字体包**（不要放 `fonts/`、`*.ttf`）；口径与实现见 `.pi/renderer/theme.py`
  与 `.pi/docs/pitfalls.md` 第 24 条。交付前跑 `python .pi/tests/font_audit.py .pi/out/xxx.pptx`
  自检（逐段核对 latin+ea 字体对 + 母版主题 + 编码）；要让**英文与数字也用雅黑**，
  渲染前设 `$env:PPT_FONT_LATIN="微软雅黑"` 即可，不用改代码。

## 依赖与离线部署

**目标解释器 = CPython 3.12 / win_amd64。** `.pi/wheels/py312/` 里的二进制全是 cp312，
3.13 / 3.14 装不上也不该装（本机开发用的 3.14 与交付无关）。

发布结构 = `requirements.txt` 的 **7 个直接依赖** + **13 个传递依赖** = 20 个 wheel / 38 MB。
传递依赖不写进 requirements.txt，pip 顺着依赖树从同一目录自己取。

三层保障，按触发顺序：

| 时机 | 入口 | 行为 |
|---|---|---|
| pi 启动前 | `.pi/extensions/deps-bootstrap.ts`（async 工厂，pi 会 await 再继续启动） | 定位解释器 → `pip install --dry-run --no-index -r requirements.txt` 判定（纯本地，不联网）→ 缺则**从内网 server（`indexUrl`）`--prefer-binary` 安装，pip 输出逐块实时转发**，失败给人话原因 → 结果缓存 24h；**每次启动必打状态行**，不吃静默 |
| 首次部署 / 环境坏了 | 仓库根 `setup.bat` | 优先 `py -3.12`；有 `.pi/wheels/py3XY` 就离线装，没有才走内网镜像 |
| 换机 / 补包 | `.pi/wheels/README.md` | `pip download` 到对应版本子目录 |

> **依赖只有一条安装通路：内网 server（`.pi/deps.json` 的 `indexUrl`）。**
> 自举**不回退公网 PyPI、也不读 `.pi/wheels` 本地目录**——离线机连不上公网，回退只会白等；
> 本地包是另一条要维护的路径，已从自举里删除。`indexUrl` 没配就直接报错，绝不偷偷走公网。
>
> **开发机到不了那个内网 IP，这是正常状态**：源码在这边写、拷到离线 PC 上跑，只有那台机器能连通
> server。所以别在开发机上试连 `172.30.39.34`（只会白等超时），也别指望开发机能验证下载路径——
> 下载链路的验证只能在离线 PC 上做。开发机上能验证的只有：检查分支（`--no-index`，纯本地）、
> 未配置地址的报错、以及失败原因分类（用回环地址 `127.0.0.1:9` 造连接拒绝）。

**两套输出通道，缺一不可**：

| 通道 | 内容 | 给谁看 |
|---|---|---|
| **TUI widget**（`deps-status`，输入框下方 `belowEditor`） | ⏳ 正在检索（20 个包：7 直接 + 13 传递）/ ⏳ 下载安装（带 pip 最新进度行）/ ✓ 已就位 20 个包 + 包清单 / ✗ 未就位 + 一句怎么查 | **用户**。pi 进 interactive 模式会切 alternate screen 重绘，stderr 全被盖掉——2026-11-18 就是因此首页一片干净，用户压根不知道依赖检没检查 |
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

**启动时终端/stderr 会看到的七种行**（前缀 `[deps]`）：

```
依赖闭包：20 个（7 直接 + 13 传递）                                ← 本地读 metadata 算的，PI_DEPS_VERBOSE=1 才显示
检查依赖是否就位（本地比对，不联网，约 10 秒）...                   ← 非缓存命中时先报这句
✓ 依赖已在位（20 个包：7 直接 + 13 传递，3.2 小时前已验证；本次跳过重复检查） ← 24h 缓存命中，没重跑
✓ 依赖已在位（20 个包：7 直接 + 13 传递，全部满足 · CPython 3.12.10，无需下载） ← 真跑过 pip 检查
依赖未就位（3 个未满足：numpy, pillow, ...），开始从内网 server 下载：<url>
pip 实时输出如下（下载进度即时刷新）...                           ← 后面是 pip 的逐包下载输出
✓ 依赖安装完成（耗时 42s，复检通过 · 20 个包：7 直接 + 13 传递） + 本次装入：a-1.0 b-2.0 …
✗ 依赖安装失败：内网 server 拒绝连接（172.30.39.34:8765）：这台机器到该 IP/端口不通…
! 未找到 CPython 3.12 解释器，本次跳过依赖检查，依赖不会被安装   ← 每次都打，不吃缓存
```

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
`numpy==2.4.6` 是否真装对也由 pip 权威实现负责。**这个脚本与具体项目无关**：
任何有 `requirements.txt` 的 pi 项目丢进 `.pi/extensions/` 即可用（含 `~/.pi/agent/extensions/`
全局生效）。`.pi/deps.json` 的 `indexUrl` 不配就**直接报错**，不拿公网 PyPI 兜底。

**本项目靠 `.pi/deps.json` 落地两个私有约束**：

```json
{ "python": "3.12", "indexUrl": "http://172.30.39.34:8765/simple" }
```

- `python: "3.12"` → 只认 3.12 解释器，找不到就跳过并每次提示（不吃缓存）；不配则用
  PATH 上的 python，不看版本。
- `indexUrl` → 内网镜像，**装包唯一来源**，pip 标准索引路径 `/simple`（不是 `/simples`）。
  脚本会自动补 `--trusted-host`（http 源不带会被 pip 直接拒绝）；不配则直接报错，**不走公网**。
- 安装失败**不阻塞 pi 启动**，只打警告 + 人话原因 + 手动重试命令——起不来的 pi 没法排查问题。
- 应急开关：`PI_DEPS_SKIP=1` 跳过、`PI_DEPS_VERBOSE=1` 看判定、`PI_DEPS_FORCE=1` 忽略 24h 缓存强制重查、
  `PI_DEPS_VERIFY=1` 预检内网 server 能否取全（只查询不安装，退出码 0/1，只在离线 PC 上跑）、
  `PI_PIP_INDEX_URL` / `PI_DEPS_PY_VERSION` / `PI_DEPS_TIMEOUT_MIN` 覆盖 deps.json。
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

## 渲染器（.pi/renderer/）

| 文件 | 作用 |
|---|---|
| `.pi/renderer/ppt_dsl.py` | 19 版式的 pydantic schema（字段名 + 长度上限），校验入口 |
| `.pi/renderer/docx_in.py <输入.docx> -o .pi/work/input.md [--media .pi/work/media]` | **Word 读入预处理**：去修订（接受插入 / 丢弃删除与批注）+ 图片原样抽盘（EMF/WMF 不转格式）+ 图位 markdown 与清单 |
| `.pi/renderer/build.py <slides_dir> -o <out.pptx> [--brief] [--style] [--speed] [--card] [--source] [--check] [--only 3,7]` | **一条命令出图**：合并 + 质检 + 渲染 + 字体自检，一份紧凑报告。**退出码恒 0**——校验拦截（slides 不合法 / 质检 error / 字体不过 / warning 未归零）属正常业务逻辑，不是运行失败；成败看末行 `RESULT {json}`（`{"ok": …, "reason": …}`）。日常和交稿都走这条 |
| `.pi/renderer/render.py <deck.json> -o <out.pptx> --qc [--brief 方向] [--style 配色] [--card 卡片] [--source 输入分析]` | 渲染 + 质检（error 会阻止渲染）；已有 deck.json 时用 |
| `.pi/renderer/merge.py <slides_dir> <out.json> [--brief 方向] [--style 配色] [--card 卡片]` | 合并 slides 为 deck：**一个文件可装 1 页，也可装一页数组（5-8 页/文件，推荐）**；校验失败定位到「文件 + 第几项 + 字段 + 当前字数 vs 上限」；自动纠偏 `type`→`slide_type`、两种 points 结构混淆（stderr 有提示）；读文件兼容 UTF-8 BOM；**校验拦截退出码恒 0**，成败看末行 `RESULT {json}`（与 build.py / render.py 同口径） |
| `.pi/renderer/budgets.py <outline.json> -o .pi/work/budgets.md` | **字数预算表**：大纲定稿后、写批次前跑，逐页下发动态容量（写作目标 = min(静态上限, 几何容量)，与 schema 校验**同一套几何**，照预算写必过校验）；`--all` 出全版式参考表；`--selfcheck` 自证"照预算写必过校验"，几何改动后重跑它当回归。预算表是"校验报数"的前置版——动态数字不在 SKILL.md 里背 |
| `.pi/renderer/fontcheck.py <pptx>` | 字体自检的唯一实现（`build.py` 同进程调用；`.pi/tests/font_audit.py` 是薄壳 CLI） |
| `.pi/renderer/qc.py` | 几何容量 / 密度 / 结构检查 + 方向红线 + 输入覆盖度 |
| `.pi/renderer/briefs/` + `briefs.py` | **写作方向**（研讨型 / 汇报型）的取舍规约与加载，见下 |
| `.pi/renderer/theme.py` | 外观的两个白名单维度（2026-10 扩容）：**配色 `style` 5 套**（`style-harness` 默认藏青 / `blue-business` / `bright` / `forest` / `violet-tech`）+ **卡片长相 `card` 5 种**（`flat` 默认 / `outline` / `tinted` / `accent-left` / `accent-top`），`STYLE_CARD_PRESET` 把每套配色绑定一种卡片（用户只选"风格"一项，`--card` 可覆盖）；`apply_style()` / `apply_card_style()` **原位**变更、未登记名确定性回退默认；**字体口径**（中文微软雅黑 + 西文 Times New Roman，只用系统字体、不随包分发）与 `apply_theme_fonts()`。底图与字体不随主题变，卡片长相不许动文字几何（见 pitfalls 第 27 条） |
| `.pi/renderer/iconlib.py` | SVG → custGeom 的图标引擎（解析规则、检索 CLI） |
| `.pi/assets/icons/` | 图标库：**tabler 单库 4620 个**（outline 3646 + filled 974）+ 中文别名 `aliases.json`，见其 `README.md` |

## 写作方向（研讨型 / 汇报型）

两套东西，一对一行：

| | 研讨型 `seminar` | 汇报型 `briefing` |
|---|---|---|
| 技能（给人/AI 读的规范） | `.pi/skills/ppt-seminar/SKILL.md` | `.pi/skills/ppt-briefing/SKILL.md` |
| 规约（给 `--qc` 读的红线） | `.pi/renderer/briefs/seminar.yaml` | `.pi/renderer/briefs/briefing.yaml` |
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
（`.pi/renderer/briefs/*.yaml`），不是渲染配置。`deck.json` 写 `"brief"` 字段
（合并时 `merge.py --brief`，渲染时 `render.py --brief` 可覆盖）。`--qc` 按该方向的
`check` 段查红线，**只出 warning**。

**改方向规约时，SKILL.md 和 YAML 要一起改**——`.pi/tests/briefs.py` 会校验 YAML 里的
`skill:` 指向的文件存在，但内容对不对得上得靠人。

新增方向 = 往 `.pi/renderer/briefs/` 放一个 yaml + 在 `.pi/skills/` 放一个 SKILL.md，
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
