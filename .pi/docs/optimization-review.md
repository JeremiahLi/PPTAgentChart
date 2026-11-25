# PPTAgent 优化建议

日期：2026-11-25。这是评估结论，不是改动清单。落地时只做一包，不要 P0 / P1 / P2 一起做。

产物慢在模型回合和常驻上下文，不在出图本身。代码上最值得做的是把几何从 `layouts.py` 拆出去；目录主体已经符合 pi 的排法，要修的是口径漂移和常驻文档过长。

待办：

- P0：缩短 `AGENTS.md`、修正 14/18 版式与汇报型默认口径、清掉 zip / tsbak / 无引用 echarts、范例合集改为生成物
- P1：`build` 四段计时；按结果决定 matplotlib 是否懒加载；`--only` 预览跳过字体全检
- P2：从 `layouts.py` 拆出无 pptx / matplotlib 依赖的 `geom.py`，并把后 5 个版式补进 smoke

结论先说：**出一份 deck 的时间主要耗在模型回合，不在 python-pptx。** 技能自己写的口径是 30 页渲染约 1 秒（[`.pi/skills/ppt-common/SKILL.md`](../skills/ppt-common/SKILL.md) 工作流第 3 步）。这次没有做计时，下面凡是「秒级」的说法都沿用这份未复测的记录。默认慢档会把这一秒乘上：大纲停一次、骨架预览、备注、warning 归零最多 3 轮。

已经做对、不要再动的部分：出图只走 DSL；`pptx.py` 用内容指纹拦跳步；5–8 页一批是为了压工具往返（40 页逐页写会变成 40 次调用）；图标 SVG 解析有进程内缓存；主题 `COLORS` 必须原位变更（[`engine/render/theme.py`](../../engine/render/theme.py) 有注释）；可执行代码放在仓库根的 `engine/`，不能搬回 `.pi/`（`.pi` 进 `sys.path` 会顶掉第三方 `pptx`）。

```mermaid
flowchart LR
  subgraph turns [模型回合 主成本]
    skills[开工技能约10万字节]
    slow[慢档: 大纲停顿加最多3轮回炉]
  end
  subgraph build [一次 build 次要成本]
    merge[merge 校验]
    qc[qc 再算几何]
    draw[layouts 再算并绘制]
    font[fontcheck 重开 zip]
  end
  skills --> slow --> merge --> qc --> draw --> font
```

## 1. 产物速度

**先动回合数，再动渲染。**

- 慢档是回合乘数，检查项与快档相同（[`.pi/briefs/seminar.yaml`](../briefs/seminar.yaml) 的 `reviewOutline` / `skeletonPreview` / `requireZeroWarnings` / `maxRefineRounds`）。这是产品默认，不建议改成快档。能省的是：骨架预览不要为了「给用户看结构」再整份渲一次；回炉时报告只带回有问题的页，避免整份 QC 文本再进上下文。回炉按批重写是对的，保持 5–8 页一批。
- 开工必读偏重。默认研讨路径要读 `ppt-entry` + `ppt-common`（约 34KB）+ `ppt-seminar` + `seminar.yaml` + `examples-all.json`（32KB）。`ppt-common` 里的 19 版式字段表和 [`engine/dsl/ppt_dsl.py`](../../engine/dsl/ppt_dsl.py) 是两份契约，模型写错时以 schema 报错为准，技能里的字段表可以收成「版式名 + 何时用 + 一条范例」，字段上限不要在 markdown 里再抄一遍。
- `examples-all.json` 与 `examples/*.json` 是同一套范例存了两遍。保留拆分文件当源，合集用 `merge.py` 生成，避免两处各改各的。开工仍只读合集一次，这个决定是对的。
- Python 侧先在 [`engine/build.py`](../../engine/build.py) 的 merge / qc / render / fontcheck 四段各打一个耗时，再决定优化哪段。未计时之前不要为几何重算做大重构。从调用关系能确定的常数成本：
  - [`engine/render/layouts.py`](../../engine/render/layouts.py) 第 26 行顶层 `from . import charts`，而 [`engine/render/charts.py`](../../engine/render/charts.py) 在 import 时就 `matplotlib.use("Agg")`。默认图表是 native，这条路径不画 PNG。注释把这写成「硬依赖、缺了就在 import 时报」。若计时显示 matplotlib 占掉那约 1 秒的大头，再改成只在 `--chart-mode mpl` 时导入，并把「缺 matplotlib」留在 `setup.bat` / deps 自举里报。
  - 同一套 `*_geom` 在 schema 的 `fit_capacity`、`qc.check`、绘制时各跑一遍。函数只有一份，这是故意的，避免预算和校验漂移。要省的是重复调用，不是再抄一套公式。
  - `--only` 仍先合并并质检全 deck，再切页（[`engine/build.py`](../../engine/build.py) 约 118–189 行）。预览可以跳过字体全检；交稿路径保持全检。
  - 图标：4620 个 SVG 的目录扫描和 603KB `index.json` 是每个进程一次，不是每页一次。渲染热路径只需要 `resolve_order` 和文件名表，关键词索引可以留在 `iconlib` 的检索 CLI。这是小头。

## 2. 代码架构

分层在目录上有了，在依赖上还没有。[`engine/render/layouts.py`](../../engine/render/layouts.py) 约 3870 行，schema、预算、质检、绘制都引用它。于是 [`engine/plan/budgets.py`](../../engine/plan/budgets.py) 算字数也会拉上 python-pptx 和 matplotlib。

值得做的一刀：抽出 `engine/render/geom.py`，只放测量和容量，不 import pptx / matplotlib / iconlib。`ppt_dsl.fit_capacity`、`budgets`、`qc` 改依赖 geom；`layouts.py` 只留绘制，内部再调 geom。几何结果可以在一次 build 里算完挂在页上，绘制和质检读同一份。这是唯一能让「dsl / plan 不依赖渲染器」成立的拆分。

其余按收益排后：

- [`engine/qc/qc.py`](../../engine/qc/qc.py) 的 `check` 是一条 400 行的版式分支。拆成每版式一个函数，行为不变。顺手去掉 statement 分支里每次 `import re`。
- native 与 mpl 是两套版面（mpl 写死图框，QC 跟的是 native 的 `chart_metrics`）。默认保持 native；mpl 标成可选，不要两套同时当契约。
- 转发桩可以留着。[`.pi/pptx.py`](../pptx.py) 的 `STAGE_MODULES` 已经指向 `engine.*`，相对导入也已改完。删桩要等文档和 `.pi/tests` 不再调用 `.pi/renderer/*.py`，单独一轮，不要和 geom 拆分混做。
- 测试缺口比再写一层包更要紧。[`.pi/tests/smoke.py`](../tests/smoke.py) 仍锁 14 种版式；`grid_2x2` / `timeline` / `funnel` / `matrix` / `key_groups` 只做了 schema 拒绝。[`pitfalls.md`](pitfalls.md) 已经写明这五个没有渲染回归。`test_planner.py` 和 `test_new_features.py` 依赖 pytest，而离线 wheels 的 20 个包里没有 pytest，离线机跑不全。要么改成标准库脚本（与 `entry.py` / `briefs.py` 一致），要么把 pytest 放进 wheels。

不要改的架构：主题字典原位更新；`build` 业务失败退出码恒 0、看末行 `RESULT`。

## 3. 文件与工作区

- 磁盘上有 [`PPTAgent1124.zip`](../../PPTAgent1124.zip)（约 121MB）和 [`.pi/extensions/deps-guard.tsbak`](../extensions/deps-guard.tsbak)。[`.gitignore`](../../.gitignore) 没排除它们。zip 应删除或忽略；`.tsbak` 不是扩展，pi 不会加载，删掉即可。
- [`.pi/assets/echarts.min.js`](../assets/echarts.min.js) 约 1MB，仓库内 py/md/ts/json/html/yaml 都没有引用。确认没有外部页面靠相对路径加载后可以删。
- 图标关键词有两份：[`layouts.py`](../../engine/render/layouts.py) 的 `_ICON_KEYWORDS`（约第 220 行）和 [`.pi/assets/icons/aliases.json`](../assets/icons/aliases.json)。检索口径收进 `aliases.json`，代码里只留「别名没命中时的兜底」。这是 [redundancy-audit](redundancy-audit.md) 里还没做的一项。
- 图标 README 仍写库约 8.7MB；当前 SVG 树约 3.3MB。改掉，避免以后按错误体积做裁剪。
- [`.pi/work/`](../work/) 里同时躺着 `outline.json`、`outline-pi*.json`、多套 `slides-*`。这些被 gitignore，不进仓库，但模型容易读错文件。技能只认 `analysis.md` / `outline.json` / `budgets.md` / `slides/`。若经常并行出多份，改成 `.pi/work/runs/<slug>/`，入口用参数指向当前 run。单份作业不必上这套。

## 4. 典型 Agent 目录

现状已经接近一个 pi 项目该有的样子：短斜杠提示 [`.pi/prompts/ppt.md`](../prompts/ppt.md)、技能在 `.pi/skills/`、扩展在 `.pi/extensions/`、方向规约在 `.pi/briefs/`、运行时 `work/` 与 `out/` 不入库、代码在可 import 的 `engine/`。

和惯例差的只有三处：

- [`AGENTS.md`](../../AGENTS.md) 每次对话都会进上下文，里面仍有整段「依赖与离线部署」和「deps-bootstrap 配置」（大约从「依赖与离线部署」到引擎模块表之前）。生成 PPT 用不到这段。收成一屏：默认三个维度、唯一入口 `python .pi/pptx.py`、分层表、指向技能。部署长文挪到 `.pi/docs/deploy.md`。生成时继续不读 `pitfalls.md`，这条约定保持。
- `ppt-entry` 是 [`.pi/skills/ppt-entry.md`](../skills/ppt-entry.md)，另外三个是 `<name>/SKILL.md`。它自己有 frontmatter，现在能被点名加载；若 pi 的技能索引只扫子目录里的 `SKILL.md`，入口技能会漏。收成 `.pi/skills/ppt-entry/SKILL.md`，与另外三个一致。
- 口径在技能之间是错的，会直接让模型用错版式或写错方向：
  - [`ppt-seminar/SKILL.md`](../skills/ppt-seminar/SKILL.md) 和 [`ppt-briefing/SKILL.md`](../skills/ppt-briefing/SKILL.md) 开头仍写「14 版式」。`ppt-common` 和合集是 19。
  - `ppt-briefing` 第 8 行写「本仓库目前的默认呈现效果就是汇报型」。入口、`build.py` 和 `AGENTS.md` 的默认都是研讨型。
  - [`examples/README.md`](../skills/ppt-common/examples/README.md) 写「18 个版式」。

目标树（不搬家图标、不把代码放回 `.pi`）：

```
AGENTS.md                      # 一屏索引
engine/                        # 代码；geom.py 从 layouts 拆出
.pi/prompts/ppt.md
.pi/skills/ppt-entry/SKILL.md
.pi/skills/ppt-common|ppt-seminar|ppt-briefing/
.pi/briefs/*.yaml
.pi/extensions/*.ts
.pi/assets/                    # 图标与底图；去掉无引用的 echarts
.pi/docs/                      # pitfalls、deploy，生成时不读
.pi/tests/
.pi/work/  .pi/out/            # 继续 gitignore
```

## 5. 不建议做的

- 不要把 4620 个图标改成运行时下载。目标机离线，库在仓库里是对的。
- 不要做按页缓存的 pptx 拼接。`--only` 跳过字体全检就够；JSON 重读不是热点。
- 不要为了「分层好看」把几何公式复制进 qc。
- 自动提交约定和当前操作习惯相反：`AGENTS.md` 要求每轮改完就 `git commit` 并 `git push origin main`；本会话的用户规则是只有明确要求才提交、且不主动 push。后续 agent 会二选一。把 `AGENTS.md` 改成与真实习惯一致，避免有人把未请求的提交推上 main。

## 落地顺序

若下一轮要改代码，只做一包，不要四包一起：

- **P0** 口径与上下文：缩短 `AGENTS.md`、三处版式/默认方向写错、删 zip / tsbak / 无引用 echarts、examples 合集改为生成物。不动渲染。
- **P1** 先加四段计时；若 matplotlib 占比高，再懒加载。`--only` 预览跳过字体全检。
- **P2** 拆 `geom.py`，并把后 5 个版式补进 smoke。删 `.pi/renderer` 桩另开一轮。

---

## 落地记录（2026-11-25）

**这份文档已经执行过了，别再照它动手。** 实际做的与原文有几处不同，记在这里免得下一轮重复讨论：

已落地：

- 代码整包迁进 `.pi/engine/`，仓库根只剩 `AGENTS.md` + `requirements.txt`；入口
  `.pi/pptx.py` → `.pi/cli.py`（改名是解开"`.pi` 进 `sys.path` 会顶掉第三方 `pptx`"的关键），
  `setup.bat` → `.pi/setup.bat`。命令口径统一为 `python .pi/cli.py <阶段>` 与
  `python .pi/cli.py run engine.<模块>`，不再有 `python -m engine.x`。
- `AGENTS.md` 从 26.4KB 压到 11.1KB，部署长文（11.9KB）拆到 `.pi/docs/deploy.md`。
- `engine/render/geom.py` 已拆出（215 个定义）：只 import 标准库 + `theme`，
  **不 import python-pptx / matplotlib / iconlib**。`theme.py` 顺手删掉了从未使用的
  `from pptx.util import Inches, Pt`（它才是 geom 仍会拉起 pptx 的真正原因）。
- matplotlib 改为在 `render_chart_mpl` 内局部 import（native 路径不再为它付 1.2s）。
- smoke 的 fixture 从 14 版式补到 19；`examples-all.json` 改由
  `.pi/tools/gen_examples.py` 生成（原来两份内容已经漂移：合集里两处正文比拆分文件新）。
- 三处口径修正（14→19 版式 ×2、"默认是汇报型"、18→19）、死文件清理
  （`deps-guard.tsbak`、1MB `echarts.min.js`）。

实测（8 页 deck，目标解释器 3.12）：

| 命令 | 改前 | 改后 |
|---|---|---|
| `build` | 3.58s | 2.29s |
| `budget` | 3.04s | 1.28s |
| `import engine.build` | 2.42s | 1.13s |

原文的错漏（供以后引以为戒）：

- 「SVG 树约 3.3MB」错，实际 7.7MB（含索引 8.3MB）。README 写的 8.7MB 只虚高 4.8%。
- 「磁盘上有 `PPTAgent1124.zip`（121MB）」错，从来没有过。
- 「pi 可能只扫子目录里的 `SKILL.md`，入口技能会漏」错：pi 会加载根目录直挂的 `.md`
  （`dist/core/skills.js` 的 `includeRootFiles`），`ppt-entry.md` 一直是能用的。
- 「ppt-entry + ppt-common 约 34KB」少算一半，实际 49KB。
- 「qc 可以完全改依赖 geom」不成立：它取的 44 个名字里混着 `add_image` /
  `resolve_image_path` / `image_size_in` / `icon_resolvable`——前三个已随 geom 搬走
  （`add_image` 要 slide 对象，留在 layouts），qc 现在确实只依赖 geom。

未做（留待下一轮，都不影响正确性）：图标关键词表与 `aliases.json` 的合并、
`qc.check` 那条 400 行分支的拆分、`.pi/work/runs/<slug>/` 的多作业隔离。
