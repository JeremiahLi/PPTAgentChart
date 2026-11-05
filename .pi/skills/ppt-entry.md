---
name: ppt-entry
description: 把文档/思路做成演示文稿 —— 先定写作方向（研讨型/汇报型），再走渲染工作流
---

# 做演示文稿

用户说"做 PPT / 演示文稿 / 汇报材料 / 研讨材料"时触发本技能。

**不要自己重写规范**：出图通道、版式契约、few-shot 范例在 `ppt-common` 技能里；
两个方向的取舍规则在 `ppt-seminar` / `ppt-briefing` 里。

## 第 0 步：三个维度都有默认值 —— 用户没提就别问

**默认即交付**：三个维度都定了默认值，用户**明确提了某一项才切换**。
不要为了显得专业把三个问题挨个问一遍 —— 每问一次就是一轮往返，而用户要的
往往就是"照默认先出一份看看"。

| 维度 | 默认值 | 换成 |
|---|---|---|
| **PPT 类型** `--brief` | 研讨型 · 全量详实（`seminar`） | 汇报型 · 结论先行（`briefing`） |
| **PPT 风格** `--style` | 经典藏青（`style-harness`） | 深蓝商务 / 明亮宣讲 / 墨绿战略 / 紫灰科技 |
| **PPT 速度** `--speed` | 慢（`slow`） | 快（`fast`） |

只有用户的话里出现了下表这些词，才切换对应维度；没出现就用默认值往下走：

| 用户的话 | 类型 | 风格 | 速度 |
|---|---|---|---|
| 研讨 / 一起过一遍 / 逐条讨论 / 评审 / 别漏内容 | `seminar`（默认） | | |
| 汇报 / 给领导看 / 结论先行 / 五分钟讲完 / 摘要 | `briefing` | | |
| 深蓝 / 正式 / 稳一点 / 对上汇报 | | `blue-business` | |
| 亮一点 / 宣讲 / 培训 / 对外交流 | | `bright` | |
| 绿色 / 战略 / 能源 / 可持续 | | `forest` | |
| 科技感 / 产品方案 / 技术平台 | | `violet-tech` | |
| 快点 / 直接出 / 先来一版 / 赶时间 | | | `fast` |
| 仔细点 / 要评审 / 别出错 | | | `slow`（默认） |

用户也可能直接报**欢迎面板上的编号**（1-9），按下表解析，与词语触发同权；
只报编号、其余没提的维度照旧用默认值：

| 编号 | 指向 |
|---|---|
| 1 / 2 | `seminar`（默认）/ `briefing` |
| 3 / 4 / 5 / 6 / 7 | `style-harness`（默认）/ `blue-business` / `bright` / `forest` / `violet-tech` |
| 8 / 9 | `slow`（默认）/ `fast` |

**风格一项已绑定卡片长相**：每套配色配一种最搭的卡片（25 种组合不用用户挑）。
只有用户明确说"卡片换成描边的 / 带色条的"时才额外加 `--card`。

出图命令**把三个维度显式写出来**（别依赖默认值，交付时可追溯）：

```bash
python .pi/renderer/build.py .pi/work/slides -o .pi/out/主题.pptx \
    --brief seminar --style style-harness --speed slow \
    --source .pi/work/analysis.md
```

> **慢档是默认值**：大纲给用户过目、warning 归零（末行 `RESULT` JSON 里
> `ok=false, reason=warnings_not_zero`）这些是默认动作。用户说"快点"才切 `fast`。

> **成败看 `RESULT`，不看退出码**：`build.py` / `merge.py` / `render.py --qc` 的退出码恒 0
> ——校验拦截（schema 不过、质检 error、字体不过）是检查在正常工作，不是运行失败。
> 每次输出末行都有 `RESULT {json}`，`ok=false` 就按 `reason` 和 `errors` 回炉。

## 第 1 步：加载规范

1. `Skill(ppt-common)` —— 公共规范（必加载）。
2. 按方向 `Skill(ppt-seminar)` **或** `Skill(ppt-briefing)`（二选一，别两个都加载）。
3. Read `.pi/renderer/briefs/<方向>.yaml` —— 机检红线（页数、图表占比、标题禁用词）。

然后按 ppt-common 的工作流走：
**输入分析 → 大纲（给用户过目）→ 预算表 → 成批生成（一个文件 5-8 页）→
build.py 一条命令出图 → 回炉 → 交付**。

> **交稿前必跑一次输入覆盖度**：`python .pi/renderer/build.py .pi/work/slides -o .pi/out/x.pptx
> --source .pi/work/analysis.md`。它会把"输入里有、deck 里没露面"的数字与模块点名 ——
> **丢了内容就必须加页补回来，不许删**。研讨型尤其必须过这一关。

## 速度档：决定"跑几个回合"，不是"查几项"

**两档跑同一套检查** —— 快档一项都不少。少查一项就会漏掉密度与几何问题，
那才是"排版异常"的真正来源；快档省的是**回炉轮次与人工环节**。

| | 慢 `slow`（默认） | 快 `fast` |
|---|---|---|
| **打断用户** | **只在评审大纲时打断一次** | 一次都不打断 |
| 大纲给用户过目 | ✅（页数 / 拆分 / 换类型 **一并这次问完**） | ❌ 生成后直接出成品 |
| 骨架预览 | ✅ **内部定稿，不给用户看** | ❌ |
| 覆盖度 `--source` | ✅ 必跑（不跑会被点名 → 归不了零） | ✅ 同样跑 |
| warning | **逐条归零**（未归零 build.py 退 3） | 同类折成一条，不再要求归零 |
| 回炉轮次 | ≤3 轮，**模型自己按报告改，不问用户** | ≤1 轮（只改 error） |
| `--notes` | ✅ 自动开（`--no-notes` 可关） | 不加 |

**慢档"不打折扣"的三件事** = 覆盖度必跑 + 骨架预览必做 + 回炉 ≤3 轮：

- 覆盖度由**代码**兜底：研讨型 `requireSource: true`，不带 `--source` 就出一条 warning，
  而慢档要求归零 → 归不了零就退 3，等于强制跑。
- 骨架预览是**给自己钉页序**用的（页序、版式、每页讲什么），不是给用户的第二稿。
- 回炉是模型读 qc 报告自己改，**不再问用户**；改到 warning 归零为止。

**除评审大纲外，任何环节都不再问用户**：页数超限（`askAbove`：研讨 50 / 汇报 30）
也是并进大纲评审那一次问完 —— 出图后看到这条 warning，按已确认的方案做即可，
**不要再开一轮问答**。

**快档也不省的**：`--source` 覆盖度核对、字体自检。两者都是本地毫秒级，
却是"不丢内容、字体不出错"的唯一保障 —— 省不掉，也不值得省。

## 第 2 步：输入分析（不可跳过）

**输入是 `.docx` 时先跑一次预处理**（去修订 + 抽图，别自己写脚本解析）：

```bash
python .pi/renderer/docx_in.py <输入.docx> -o .pi/work/input.md      # 图片落到 .pi/work/media/
```

规划引擎现在支持**自动切分 + 骨架 + 结构化大纲**：

```bash
# ① 先让代码把分析表切出来（章节/模块/论据/数字自动填好，你只补"结论"和"缺口"）
python .pi/renderer/plan.py .pi/work/input.md --emit-analysis      # 默认写 .pi/work/analysis.md

# ② 补完 analysis.md 的「核心结论 / 缺口」两列后再规划（这才是产出大纲的那一步）
python .pi/renderer/plan.py .pi/work/analysis.md -o .pi/work/outline.json \
    --title "主题" --brief seminar
```

> `--emit-analysis` **只切骨架、不出大纲**：它把"读全文写整张表"换成"改三列"，
> 少几千字输出。骨架里的「章节/模块」可以合并改名，但**不要删掉要上页的模块** ——
> `--source` 覆盖度检查按最终这张表核对。没有 `input.md`（非 docx 输入）时，
> 把原文 md 路径直接喂给它即可。

> **分析表的一切修改都在骨架文件上小步 edit，禁止整段重写**（硬规则）：
> 补「结论/缺口」两列、合并、改名，一律对现文件用 edit 逐处改；
> **不许在 thinking 里整段起草新表再一次性写盘**。超长单轮输出会拉长流式回合
> （实测 LongCat 一轮 84s thinking 被服务层断流，整轮作废重来），
> 也放大字段长度类校验错误的批量出现。已有 `analysis.md` 时同理：在它上面改，不重写。

大纲是结构化 JSON，包含：每页的版式类型、密度档位、图表建议、备注建议，
以及**关系合并记录**（`relation_note`：时序链合流程图、对比合左右对照、
并列合卡片组、包含合框架图——模块之间的联系决定布局，合并理由可追溯）。
**检查缓存**：同一份文档再次生成时用 `--check-cache`，未变则跳过。

```bash
# 缓存命中时直接复用
python .pi/renderer/plan.py .pi/work/analysis.md -o .pi/work/outline.json --check-cache
```

大纲可以一键落成一份**结构稿**（每页合法、正文全是「待补：」占位）：

```bash
python .pi/renderer/merge.py --from-outline .pi/work/outline.json -o .pi/work/deck.json
python .pi/renderer/render.py .pi/work/deck.json -o .pi/out/骨架.pptx   # 给自己钉页序
```

> 这份骨架**不是成品，也不给用户看**（慢档只在评审大纲时打断用户一次）：它的作用是
> 把"页序、版式、每页要讲什么"先钉下来，再逐页替换正文。`--from-outline` 以前会抛 29 个 validation error（骨架缺必填字段），
> 现在每一步都能跑通；`figure` 会退化成 `statement` 并在标题里注明（缺图没法占位）。

大纲定稿后、写批次前，跑一次**字数预算表**（逐页下发动态容量，与 build.py 校验
同一套几何——照预算写校验必放行，长度类返工在写前就消灭）：

```bash
python .pi/renderer/budgets.py .pi/work/outline.json -o .pi/work/budgets.md
```

然后按 `ppt-common` 的工作流走：
**输入分析（--emit-analysis 出骨架）→ 大纲（给用户过目）→ 预算表 →
成批生成（一个文件 5-8 页）→ build.py 一条命令出图 → 回炉（按批重写）→ 交付**。

## 修 slides 报错的纪律（先读后改，别盲改）

历次失败的大头不是渲染器，是"凭记忆操作"：凭记忆写 JSON → 校验失败；凭记忆写补丁脚本 → 结构猜错；
凭记忆写 oldText → edit 失败。四条硬规则：

- **edit 的 oldText 必须来自刚 Read 过的原文**。改哪份 JSON 就先 Read 哪份；
  连续 2 次报"找不到原文"说明在猜，停下来重新 Read。
- **禁止用 `python -c` / heredoc 内联补丁脚本改 JSON**（多层括号嵌套必炸 SyntaxError、
  结构靠猜必 KeyError）。改 JSON 只用 edit 工具。
- slides 目录里每份文件是**单页对象或页数组**，没有 `slides`/`sections` 外层键
  （那是 deck.json 的结构），别按 deck 结构去取。
- 校验报错已带「文件 + 第几项 + 字段 + 当前字数 vs 上限」：**一次把点名的字段全部改完再重跑**，
  不要改一个跑一次（每次重跑都是一轮往返）。
  另注意两种版式的 points 结构**正好相反**：`panel_2col` 是字符串数组、
  `statement` 是 `{label,text}` 对象数组；merge 会自动纠偏最常见的两种混淆并在 stderr 提示。

> 45 行那条覆盖度检查，现在等价于 build.py 的 `--source`：`--qc --source` 与
> `build.py ... --source` 走的是同一份 `qc.check`，报告格式也一致。

> **备注建议**：加 `--notes` 后渲染器会根据每页内容特征自动在 PPT 备注区写入
> 「缺什么数据 / 补充什么资料 / 可视化改进」建议。这些建议来自规则引擎，
> 不依赖 LLM，同一份文档重复生成时保持一致。

## 本技能补的三件事

1. **需求澄清**：确认主题、受众、有无现成素材文档。素材优先读用户给的文件，
   不假设能联网。
   **三个维度默认不问**（类型=研讨型、风格=经典藏青、速度=慢），用户提了才切换 ——
   见第 0 步。（页数同样不单独问：两个方向都由资料量决定、都不设上限 ——
   研讨型超 50 页、汇报型超 30 页时，把"只讲某几章 / 拆成多份 / 换类型"
   **并进大纲评审那一次**问完，不要出图后再开一轮。）
2. **产出落点**：输入分析写 `.pi/work/analysis.md`（骨架由 `--emit-analysis` 生成）；
   逐页 DSL 写成 `.pi/work/slides/NN-batch.json`，**一个文件装 5-8 页（数组）**；
   出图一条命令：`build.py .pi/work/slides -o .pi/out/<主题>.pptx --brief seminar
   --style style-harness --speed slow --source .pi/work/analysis.md`（自带质检 + 字体自检），
   成品输出到 `.pi/out/<主题>.pptx`。**中间产物只留这几样**：`analysis.md`、
   `outline.json`、`slides/*.json`、`deck.json`（需要时 `--keep-deck` 才生成）、`.pptx`；
   不要再写 `qc_report.txt` / `render_report.json` 这类重复报告（报告就在 stdout）。
3. **交付**：给出 .pptx 路径 + **本份用的三个维度**（如"研讨型 / 深蓝商务 / 慢档"，
   用户一句话就能换其中任一项）+ 每页一句话摘要（页码 / 版式 / 标题），请用户检查。
   研讨型还要额外交一份**论点论据落点核对表**（`analysis.md` 里每一项 → 落在第几页）；
   分析表里"缺口"列有内容的，交付时一并列出**还缺什么、要找谁要**。
   **慢档交付前 warning 必须是 0**（`build.py` 结果里 `ok=false, reason=warnings_not_zero` 就是在拦这件事）；
   快档交付时如实说明"按快档放行，还有 N 条 warning 未处理"。

## 边界

- 渲染器在 `.pi/renderer/`，出图只走
  `python .pi/renderer/build.py .pi/work/slides -o <out.pptx> --source .pi/work/analysis.md`
  （或 `render.py <deck.json> -o <out.pptx> --qc --source ...` 的老路子）。
- 旧的确定性流水线（`src/ppt/`、`SPEC/`、`tools/`、`demo/`、`run/`）**已删除**，
  别再按旧文档去找；本仓库只走 `.pi/renderer/` 这一条出图通道。
- 不要手写 python-pptx / OOXML / HTML→pptx。
