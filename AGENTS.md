# AGENTS.md

本仓库是**基于 pi 的 PPT 制作 Agent**（内网部署、不可联网的私有模型）：把文档/思路做成
高信息密度、图表丰富的 .pptx。出图唯一通道是 `.pi/engine/`（PPT-DSL 结构化 JSON → pptx），
**禁止手写 python-pptx / OOXML**。

## 目录（2026-11-25 起）

```
<repo>/AGENTS.md            # 本文件，一屏索引
<repo>/requirements.txt     # 7 个直接依赖（传递依赖 13 个，见 docs/deploy.md）
<repo>/.pi/
    cli.py                  # 唯一入口：analyze → outline → budget → 写页 → build
    engine/                 # 全部可执行代码（见下表）
    skills/ briefs/ prompts/ extensions/   # 技能 · 方向规约 · 斜杠提示 · pi 扩展
    assets/ work/ out/ wheels/ docs/ tests/ tools/ setup.bat
```

**仓库根只有 `AGENTS.md` + `requirements.txt`**（外加 `.git` / `.gitignore` / `.workbuddy`）。
代码与资源一律放 `.pi/` 下。

> **`.pi` 顶层不许有与第三方同名的 `.py`**：`python .pi/cli.py` 会把 `.pi` 放进
> `sys.path[0]`，一旦那儿有 `pptx.py`，`import pptx` 就命中它，报 `'pptx' is not a package`。
> 入口当年就叫 `pptx.py`，改名 `cli.py` 才解开这个死结——**别改回去**。
> `.pi/tests/layout.py` 会守这条（含 requirements 依赖名的自动比对）。

| 目录 | 装什么 |
|---|---|
| `.pi/engine/dsl/` | 契约层：`ppt_dsl.py`（19 版式 schema）、`briefs.py`（方向规约加载） |
| `.pi/engine/plan/` | 规划层：`doc_analyzer` → `planner` → `outline` / `plan` / `budgets` |
| `.pi/engine/render/` | 渲染层：`geom`（测量/容量，不碰 pptx）、`theme` / `charts` / `icons` / `iconlib` / `layouts` / `merge` / `render` / `speaker_notes` |
| `.pi/engine/qc/` | 检查层：`qc.py`（几何/密度/红线/覆盖度）、`fontcheck.py` |
| `.pi/engine/ingest/` | 摄入层：`docx_in.py`（去修订 + 抽图） |
| `.pi/engine/build.py` | 编排层：merge → qc → render → fontcheck 一条命令 |
| `.pi/engine/paths.py` | 资源根路径唯一来源（`.pi` 本身 = `PI_ROOT`） |

**分层靠依赖方向守着**：`geom.py` 只 import 标准库 + `theme`，不 import python-pptx /
matplotlib / iconlib。所以 `dsl` / `plan` / `qc` 只要几何就能跑，`budget` 这类命令
不必为"画图"付启动成本（实测 3.0s → 1.3s）。往 `geom` 里加函数前先问一句：
它用到 `slide` / `In()` / `MSO_SHAPE` 了吗？用到了就该留在 `layouts.py`。

## 唯一入口：`.pi/cli.py`

顺序以前只写在 markdown 里，跳步没人拦；现在顺序在代码里，**跳一步必然失败**：

```bash
python .pi/cli.py                          # 看状态（不带子命令 = status）
python .pi/cli.py analyze 材料.docx         # Word 预处理 + 分析表骨架
python .pi/cli.py outline --title "主题"    # 分析表 → 结构化大纲
python .pi/cli.py budget                   # 大纲 → 逐页字数预算
python .pi/cli.py build -o .pi/out/x.pptx  # 合并 + 质检 + 渲染 + 字体自检
python .pi/cli.py check                    # 只质检不出文件
python .pi/cli.py run engine.qc.fontcheck .pi/out/x.pptx   # 直通任意 engine 模块
```

**只有这一套命令口径。** 不再有 `python -m engine.x`（engine 在 `.pi` 下，裸敲找不到包）；
要单独跑某个模块就 `cli.py run engine.<子包>.<模块>`。

三条契约（与 `build.py` 同口径）：

| 项 | 约定 |
|---|---|
| **成败** | 每条命令末行都是 `RESULT {json}`（`status` 也有），看 `ok` + `errors[].code`；`code` 是稳定枚举（`E_STAGE_MISSING` / `E_STAGE_STALE` / `E_QC_ERROR` …），中文只在 `message` / `hint` 里 |
| **退出码** | 业务结果（校验拦截、上游缺失、过期、warning 未归零）**恒 0**；只有未捕获异常退 1，参数拼错退 2（`run` 子命令原样透传目标模块的退出码） |
| **状态** | `.pi/work/state.json`：每阶段记输入的**内容指纹**（sha256 前 12 位，不是 mtime），改过源文件下游自动判过期 |

门禁两级：**requires**（缺了报错，如大纲缺分析表）、**advises**（缺了只告警，如没跑 budget
就出图）。`--force` 可硬闯，`RESULT` 留 `forced: true`。入口层回归 `python .pi/tests/entry.py`
（纯标准库）；结构回归 `python .pi/tests/layout.py`。

## 谁负责什么

- **三个用户维度 + 默认值**（`build --brief / --style / --speed`）：
  **类型 = 研讨型 · 全量详实**、**风格 = 经典藏青**、**速度 = 慢**。
  用户没提就按默认走、提了才切换，**不要逐项追问**（每问一次一轮往返）。
  风格已绑定卡片长相，用户不点名就不用传 `--card`。详见 `ppt-entry` 第 0 步。
- **触发入口 + 工作流**：`.pi/skills/ppt-entry.md` + `/ppt <文档>`（`.pi/prompts/ppt.md`）。
  需求澄清 → **输入分析**（`.pi/work/analysis.md`，骨架由 `plan --emit-analysis` 生成，
  只补结论/缺口两列）→ 大纲（**唯一停下来让用户过目的环节**）→
  **成批生成（一个文件 5-8 页）** → `build` 一条命令出图 → 回炉 → 交付。
  "标题即结论 / 图示优先 / 信息密度 / 结论段硬指标"这些铁律在 `ppt-entry` + `ppt-common`。
- **公共规范 + 范例**：`.pi/skills/ppt-common/SKILL.md`（出图通道 / 19 版式契约 / Word 读入 /
  图标 / 通用铁律）+ `examples-all.json`（**19 版式合集，开工读一次**）。
  **合集是生成物**：源是 `examples/NN-*.json`，改完跑 `python .pi/tools/gen_examples.py`，
  别手改合集（`--check` 只比对）。**范例是最低质量线**，密度只能更高不能更低。
- **两个写作方向技能**：`ppt-seminar`（研讨型）、`ppt-briefing`（汇报型）。**二选一**。
- **运行环境**：内网部署、不可联网的私有模型，全程中文（简体）。
- **字体**：中文微软雅黑 + 西文 Times New Roman，只用 Windows 自带字体，
  **仓库不带字体包**；口径见 `.pi/engine/render/theme.py` 与 `docs/pitfalls.md` 第 24 条。
  交付前 `python .pi/tests/font_audit.py .pi/out/xxx.pptx`。要让英文数字也用雅黑，
  渲染前设 `$env:PPT_FONT_LATIN="微软雅黑"`。

## 引擎模块速查

| 模块 | 作用 |
|---|---|
| `engine.dsl.ppt_dsl` | 19 版式的 pydantic schema（字段名 + 长度上限），校验入口 |
| `engine.ingest.docx_in` | Word 读入：去修订 + 图片原样抽盘（EMF/WMF 不转格式）+ 图位 markdown |
| `cli.py build <slides_dir> -o <out> [--brief/--style/--speed/--card/--source/--check/--only 3,7]` | 一条命令出图，日常与交稿都走这条 |
| `engine.render.render` | 渲染 + 质检（error 阻止渲染）；已有 deck.json 时用 |
| `engine.render.merge` | slides → deck；**一个文件可装 1 页，也可装一页数组（5-8 页/文件，推荐）**；校验失败定位到「文件 + 第几项 + 字段 + 当前字数 vs 上限」 |
| `cli.py budget` | 字数预算表：写作目标 = min(静态上限, 几何容量)，与 schema **同一套几何**，照预算写必过校验；`--all` 全版式参考表，`--selfcheck` 自证 |
| `engine.qc.fontcheck` | 字体自检唯一实现（`build` 同进程调用；`.pi/tests/font_audit.py` 是薄壳） |
| `engine.qc.qc` | 几何容量 / 密度 / 结构检查 + 方向红线 + 输入覆盖度 |
| `engine.render.theme` | 配色 `style` 5 套 + 卡片长相 `card` 5 种（`STYLE_CARD_PRESET` 绑定）；`apply_style` **原位**变更；字体口径也在这里 |
| `engine.render.iconlib` | SVG → custGeom 图标引擎（解析规则、检索 CLI） |
| `.pi/assets/icons/` | tabler 单库 4620 个 + 中文别名 `aliases.json`，见其 `README.md` |

## 写作方向（研讨型 / 汇报型）

| | 研讨型 `seminar`（默认） | 汇报型 `briefing` |
|---|---|---|
| 技能 | `.pi/skills/ppt-seminar/SKILL.md` | `.pi/skills/ppt-briefing/SKILL.md` |
| 规约（`--qc` 红线） | `.pi/briefs/seminar.yaml` | `.pi/briefs/briefing.yaml` |
| 受众 | 会逐条讨论的同事 | 只看这一份的决策者 |
| 取舍 | 资料论点论据**全部上页不丢弃** | 每页一个主张，可抽象可砍 |
| 页数 | **不限**，由资料量决定；预计 >50 页先问 | **不限**；预计 >30 页先问 |
| 每页论据 | 不限 | ≤3 |
| 图表 | 尽量每页都有，图型至少 3 种 | 够用即可 |
| 标题 | 判断句 | 判断句 + 禁"情况介绍/相关说明/概述" |
| 缺口 | 必须显式占位 | 可选 |

**默认组合 = 研讨型 + 经典藏青（`style-harness` + `flat`）+ 慢档**。速度档只改"跑几个回合"
（大纲给不给用户过目、warning 要不要归零），**两档跑同一套检查**。

两个方向**共用同一套版式和配色**，差别在怎么取舍原文，所以规约是策略数据（`.pi/briefs/*.yaml`）。
`deck.json` 写 `"brief"` 字段（`merge --brief` / `render --brief` 可覆盖），`--qc` 按该方向的
`check` 段查红线，**只出 warning**。**改方向时 SKILL.md 和 YAML 要一起改**（回归：
`python .pi/tests/briefs.py`）。`deck.style` / `deck.card` 是**外观**，与 `brief` 正交，别互相牵连。

新增方向 = 一个 yaml + 一个 SKILL.md，**零代码改动**。

## 已踩过的坑 → `.pi/docs/pitfalls.md`

**28 条坑位 + 关键长度陷阱摘要**全在那份文件里，每条写清"错在哪、现在怎么做、为什么别改回去"。

- **改渲染器 / 改规则之前**先扫一遍，尤其别把修好的坑改回去。
- **生成 PPT 时不需要读它** —— 生成只读 `ppt-entry` → `ppt-common` → 方向技能 → `examples-all.json`。
- 规则分三层：**判据在代码里、机检阈值在 `briefs/*.yaml`、写作规范在 `.pi/skills/*.md`**，
  同一句话不要在三处都写（历史上"12 版式""60~140 字"就是这么漂移的）。

## 部署与离线依赖 → `.pi/docs/deploy.md`

装依赖、换机、补包、deps 自举的 7 步日志、wheel 清单口径、内网索引要求，全在那份文件。
**生成 PPT 用不到，别读。** 记住三条就够：目标解释器 CPython 3.12 / win_amd64；
依赖只有一条通路（内网 server，写死在扩展的 `CONFIG.indexUrl`）；首次部署跑 `.pi/setup.bat`。

## 已删除的旧路径（别再找）

- `src/ppt/`、`SPEC/`、`tools/run_ppt.py`、`tools/mock_model.py`：旧的确定性流水线，2026-09 整体删除。
- `.pi/renderer/`：迁移期转发桩，2026-11 拆除。旧命令会直接报文件不存在。
- `python -m engine.<子包>.<模块>`：2026-11-25 起改为 `python .pi/cli.py run engine.…`。

## 提交约定（强制执行）

每轮对代码/文档/规范做完优化，任务收尾时当场 `git add` + `git commit`（信息一律**简体中文**、
概括本次改动）+ `git push origin main`，不要等用户提醒。禁止 `--force`；`.gitignore` 已排除的产物不入库。
