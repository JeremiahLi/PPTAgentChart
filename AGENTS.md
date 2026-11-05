# AGENTS.md

本仓库是**基于 pi 的 PPT 制作 Agent**（内网部署、不可联网的私有模型）：把文档/思路做成
高信息密度、图表丰富的 .pptx。
出图唯一通道是 `.pi/renderer/`（PPT-DSL 结构化 JSON → pptx），**禁止手写 python-pptx / OOXML**。

> **目录结构（2026-10 调整）**：工程主体 `renderer/`、`assets/`、`tests/`、`work/`、`out/`、
> `docs/`、`wheels/` 全部位于 **`.pi/` 下**；仓库根只留 `AGENTS.md`、`requirements*.txt`、`setup.bat`。
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
