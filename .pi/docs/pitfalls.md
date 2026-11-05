# 渲染器 / 规则的坑位清单

> **这份文件是给「改渲染器、改规则」的人看的**（人或 AI）。**生成 PPT 时不需要读** ——
> 生成时只读 `.pi/skills/ppt-entry.md` → `ppt-common` → 方向技能 → `examples-all.json`。
>
> 规则分三层，各自只放该放的东西：**判据（schema 上限 / 几何函数）在代码里、
> 机检阈值在 `.pi/renderer/briefs/*.yaml` 的 `check` 段、写作规范在 `.pi/skills/*.md`** ——
> 同一句话不要在三处都写一遍（历史上"12 版式""60~140 字"就是这样漂移的）。
>
> 每条都写清**错在哪、现在怎么做的、为什么别改回去**。编号会随增删变化，**引用请引用标题**。
> 共 **24 条坑位**，结尾另附一份「关键长度陷阱」摘要（容量 / 下限 / 结论段 / 留白的现行口径）。
> 回归入口：`python .pi/tests/smoke.py`（渲染/几何）、`python .pi/tests/briefs.py`（方向规约）、
> `python .pi/tests/chart_xml.py`（图表 XML + 真开一次 Office）、
> `python .pi/tests/docx_in.py`（Word 去修订 + EMF 原样抽盘）。
> 仓库总览、工作流与写作方向见 [`../AGENTS.md`](../AGENTS.md)。

## 1. 环形图不许写 `c:dLblPos`

ECMA-376 里 `doughnutChart` 不允许数据标签定位，python-pptx 却照写不误；PowerPoint 打开时
会判「文件损坏」，点「修复」后**图表被丢弃**，那一页整页空白。
`layouts._normalize_chart_xml()` 做了出口兜底，**别删**。

回归：`python .pi/tests/chart_xml.py`（6 种图型逐一出图 + XML 合规 + 装了 Office 就真开一次文件）。

## 2. 封面底图是深蓝的

`.pi/assets/template/bg-cover.png` 实测亮度 0.01–0.16。封面文字必须走 `COLORS["cover_*"]` 浅色系；
`text` / `text_sub` 这类深色字压上去对比度只有 1.2–2.9:1，等于看不见。换底图时
`layouts._cover_ink()` 会自动按明暗回退到 `cover_*_on_light`。

## 3. 一个内容块只画一个图标

块自带 `icon` 就画在块头；没写就按文字自动补一个，**正文/描述前面不再补第二个** ——
v14 的 P9/P17/P19 就是块头一个、正文前一个，两个图标打架、整页发飘。
statement 的正文分组卡不属于"已有图标"，它按组自动配一个（那就是"框里没图标就补一个"的形态）。

## 4. 加粗只认 `**...**`

渲染器**不再**自动把每组首句整句加粗（v14 的 P4 就是整句加粗，重点被废话稀释）。
重点由 LLM 用 `**` 显式标出，`--qc` 会检查跨度 >16 字、占比 >40%（长文本）、
statement.body 一个都没标这三种情况。

## 5. 标题下那段结论段是"判断"，不是"导语"

规则本来只写在 SKILL.md 里（"2-4 行的结论"），qc 只查了"有没有 / 是否 <30 字"，
于是 LLM 写出一堆废话、"1.2 节给出…"、复述标题 —— 规则等于没生效。
现在 `qc._check_summary` 逐条点名：**40~80 字 / 2~3 句 / 禁坐标词与自指 / 禁填充语 /
不复述标题**（全 warning）。改词表就在 `.pi/renderer/qc.py` 的 `_SUMMARY_COORD` /
`_SUMMARY_FILLER`，注意别把"上下文"当成自指（`上文|下文` 会误报）。
区间 2026-09 由 60~140 收到 40~80：页头是每页最贵的版面（多一行 ≈ 一张卡的高度），
而结论段与卡片描述的信息高度重叠 —— 现在只留「判断 + 关键数字」，机制下放卡片。

## 6. 铺满靠补内容，不靠拉高排版

表格行高**保持紧凑**（`TABLE_ROW_H`），不许按可用高度摊开 —— 留白要用「把下一章节/关联
章节的明细并进来」或「补一行显式占位」解决，`layouts.table_slack()` 会把「还空多少、
还能放几行」算给 `--qc` 报出来。statement 的行高按各组内容比例分配后再拉伸铺满，
文字垂直居中、行距自适应；高亮卡的**高度要先把底部扣掉**，否则它会直接压在正文卡上
（v14 P4 的排版错乱）。

## 7. 密度低不是"没写满"，是三个具体病

①整页只有一个容器；②内容少还硬占一页（半页留白，v19 的 P19/P20/P44 就是这样）；
③一块内容上下叠一块。落点是版式 **`panel_2col`**：左右两个模块，**每个模块一个背景框 +
框顶一段总结陈述段**（`panels[].summary`），下面放要点或迷你明细表。
`--qc` 报「主表下方还空 X in / 下半页还空 X in / 两个模块还剩 X in」时**去并页**，
不要拉高行高、也不要就这么交出去。留白函数单一来源：`table_slack` / `chart_bottom` /
`panel_slack`（qc 与渲染共用，别在别处另算一套）。

## 8. "精简过头"是反向的同一个病

（v19 的教训）：把正文压成短语，版面看着满、信息却被抽干。`qc._check_density_floor`
给每类版式定了正文下限 —— statement.body **150+**、compare 每栏 desc **45+**、
numbered_cards 每卡 desc **40+**、layered_stack 每层 desc **30+**、flow 每步 desc **16+**、
hierarchy node.desc **20+**。阈值集中在 `.pi/renderer/qc.py` 的 `_*_DESC_MIN` /
`_STATEMENT_BODY_MIN`，要调改那里。
**卡片型版式要 description 就要狠心少放 callout**：callout 占 0.96in，会把卡片描述空间
压掉一大截（numbered_cards 因此有了 `footer_note` 这条更省的出路）。

## 9. flow 环形图以前忽略 `callout`

带 `loop_note` 的 flow 走 `_render_flow_cycle`，现在先画辅区、再按剩余高度把椭圆收窄
（`Ry` 收紧）；改这段别又把 callout 丢掉。

## 10. 图标名写错会"静默丢图标"

（用户反馈的 P6「Git 提交记录」= 别名「版本」没登记）。旧 `_resolve_icon` 把用户填的名字
**原样返回**，值非空 → 调用方不再兜底 → 块头空着且没有任何提示。现在三层保障：
① `_resolve_icon` 先用 `iconlib.resolve_icon` 验证，认不出来按块内文字自动补、再退到
`_ICON_FALLBACKS`；② `_draw_block_icon` 连画都画不出来时退化成同色圆角色块，
**块头永远不空**；③ `qc._check_icons` 逐页点名认不出的名字。
新增常用中文别名就加到 `.pi/assets/icons/aliases.json`。

## 11. 辅区（一页的第二块信息）默认放右侧纵栏，不是底部横条

底部横条会**吃掉主内容的高度** —— 环形流程被压成扁椭圆（用户反馈的 P5），三栏对比的
描述被挤成一句。落位由每个版式声明「主内容至少要多宽」：`_MIN_CARD_W` / `_MIN_COL_W` /
`_MIN_FLOW_W` / `COMPARE_MIN_W` / `LAYER_MIN_W` / `FLOW_CYCLE_MIN_W` / `PANEL_MIN_W` /
`HIER_MIN_W`；`card_frame()` 统一算、`draw_aux()` 统一画，qc 的 `compare_desc_box` /
`hierarchy_geom` / `panel_slack` / `table_slack` / `statement_slack` / `flow_linear_geom` /
`numbered_geom` 全部走它。

- 右栏：表格 `side`、图表 `insights`、环形流程、≤4 步线性流程、层级堆栈、4 张编号卡
- 底部：≥5 步线性流程、5-6 张编号卡、三栏对比、层级图、双模块页（后两者本身已横向铺开，
  或卡片高度贴内容走，让出右栏只会多出纵向留白）

## 12. 页脚条压扁 + 内容带下沿统一到 `BODY_BOTTOM = 6.80`

（2026-09）页脚条原来 0.56in，而且每个版式还各自再扣一次（`6.55 - 0.56`）——同一块页高被
扣两遍。现在页脚条是 `FOOTER_Y=6.90 / FOOTER_H=0.36 / FOOTER_RIGHT=10.85`（右端避开右下角
单位水印 x 10.95~12.90），页码也在 6.90，**有没有写 `footer_note` 不再影响内容带高度**：
各版式统一用 `BODY_BOTTOM`，`TABLE_BOTTOM` / `CHART_BOTTOM` 全部指向它。
**别再往各版式里塞底线数字。** `footer_note` 上限因此收到 64 字（单行，字号 11→9 自适应）。

## 13. "放不下"从 error 改成"渲染器自己降档"

旧实现里 6 类版式的几何溢出都是 **error（阻止渲染）**，LLM 的理性策略就是写短 ——
这是"过度精简"的机制性根因。现在渲染器按 `FIT_STEPS` 自己降档（最低 10pt），
`qc._cap_note` 只在"降到最小档仍放不下"时出 warning，并把降到几 pt 记进
`stats.downshifted`；error 只剩「内容不足 30 字」「正文碎片化」「方向名写错」
「`figure` 的图找不到」四种（最后一种是唯一"产物会空"类，见第 23 条）。
**改渲染/质检时别把 error 加回来。**

## 14. 卡片标题即重点（判断句）

`numbered_cards[].name` 要写成判断句或带数字的结论（「统一任务集管理」
「回归通过率 62%→88%」），不是名词短语（「任务集管理」）；`sub` 已改为可选 ——
标题本身是判断句就别再补一句定位。`qc._check_claim_titles` 按「无数字 + 无判断性谓语 +
过短/名词后缀」启发式点名，词表在 `_CLAIM_SUFFIX` / `_CLAIM_VERB`。
这就是"无法图表化的内容 → 卡片 + 加粗重点标题"的落点。

## 15. 纯文字卡必须有骨架：关键句 + 关键短语

（2026-09，用户反馈 v22 P11"三个卡片式的文本段，没有任何特点和重点…有些页面文字已经
超出整个背景框了"）。最终定的形态是「**同字号 + 加粗 + 强调色**」：

- `split_body_groups`：`statement.body` 里**用换行指定卡片边界**（≤3 张，排 2 列 × 2 行）；
  没写换行仍按老规矩自动分（短正文单卡、长正文按字数 2~3 组）—— **向后兼容**。
- `card_paras`：**一段到底**。每张卡的**第一句**自动包进 `**…**`（关键句），
  作者写的 `**短语**：` 原样保留 —— 两者都由 `text_box(strong_color=…)` 渲染成
  加粗 + 强调色，**字号与正文相同**、冒号后正文在同一段里接着排。
  `card_text_h` 因此**就等于普通文字的高度** —— 这是防溢出的关键。
- 曾经的错误做法（**别再改回去**）：把关键句提到卡顶、`size+1.2` 加粗、下面压一条
  细线 —— 多出来的那行直接把文字顶出卡片背景。
- `_SECTION_RE` 认 `**短语**：正文` 与 `**短语：**正文` 两种写法。**冒号必须有**，
  否则普通加粗（`**回归通过率**提升到…`）会被误判成关键短语。
- 渲染/算高/挑档必须都走 `card_paras` / `card_text_h`（statement 正文卡、
  `numbered_cards.desc`、`compare_3col.desc` 已切过去）。
- qc：`_check_card_structure` 点名「关键句 > 48 字」与「长文（>170 字）<2 处关键短语」。

## 16. 块头标题必须给右上角图标让位

（2026-09，用户实测 v22 P9 第 3/5 张卡"标题盖住了图标"）。编号卡的块头是
「序号 + 标题（左上） … 图标（右上）」，标题可用宽度必须扣掉图标那一段，
否则 3 列窄卡（5-6 张卡）里长标题一路写到图标上：

- `block_title_w(card_w, left_gutter, has_icon)` 是**唯一口径**（扣 `CARD_ICON_RIGHT +
  ICON_TITLE_GAP`）；三栏对比的栏名是居中的，用 `compare_label_w(col_w, True)`
  （两侧都要让位）。渲染与 qc 都调这两个函数，别在别处另算宽度。
- 标题放不下时 `fit_one_line()` 自动缩半档字号（最低 11.5pt）保住一行 ——
  但**缩字号不是解法**：`qc._check_block_titles` 会按 `card_title` 原字号点名
  「标题 N 字放不下（约 X 字/行），会压到图标或折成两行」，让作者把标题砍短。
- 3 列卡（5-6 张）标题 ≈ **10 字以内**；2 列卡（4 张）约 20 字。

## 17. 版式上限 ≠ 删内容的理由

（2026-09，兄弟项目的运行日志暴露的真问题）日志里的原话是「第 5 页 flow 要承载
4+2+2=8 步，超过 flow 的 3-6 步上限」→ 于是模型**不断压缩总页数、丢掉不重要的页**
（"同一份输入，页数越来越少的 PPT"）。这是**激励设计错了**：上限拦住了"并页"，
而页数压力又把"拆页"堵死，剩下唯一出路就是删内容。本轮修了三件事：

- **flow 放开到 3~12 步**：`flow_linear_geom` 在 >`FLOW_WRAP_AT(6)` 步时折成两行
  （第一行左→右、折返通道 `_draw_flow_wrap` 回到第二行，`loop_note` 写在通道上）。
  环形图只在 ≤6 步时用（`render_flow` 判断）—— 8 个节点绕椭圆会互相压住。
  所以"8 步的流程"现在是**一页一个版式**，既不拆也不删。
- **其余上限一律对齐"真实几何容量"，而不是拍脑袋**（2026-09 第二轮放宽，用户追问
  "文档里有几层就该画几层"）：`layered_stack` 3→**12 层**（层数多时 gap 0.14→0.10→0.07
  自动压紧）、`numbered_cards` 6→**8 张**（7-8 张排 4 列，且 `has_icon=False`
  **窄卡不画块头图标** —— 0.80in 的图标会让 2.9in 的卡标题只剩 6 字，可读性优先于
  "一块一个图标"）、`hierarchy` 6→**8 个**、`compare_3col` 3→**4 栏**
  （`compare_desc_box(..., n_cols=)` 与渲染共用栏数，别在渲染里再写死 `/3`）、
  `table` 10→**12 行**、`chart.detail_rows` 8→**10 行**；
  正文上限 `statement.body` 520→**700 字**、`numbered_cards.desc` 160→**240**、
  `compare_3col.desc` 140→**220**、`layered_stack.desc` 110→**180**
  —— 这几个都是原来**远低于几何容量**的人为保守值，是"密度上不去"的直接原因。
  放宽前按各版式新上限塞满做过探针：**9 页 0 溢出、无 <9.5pt 正文**。
- **每个上限都在字段 description 里写明"超了就另起一页，不要删条目"** ——
  pydantic 的报错消息是模型唯一看得见的反馈，报错里说清楚该干什么，
  它才不会用"删内容"去满足约束。

## 18. 输入覆盖度检查：唯一能反向抓"整块内容没进 deck"的机制

（`--qc --source`）其它检查全是"页内"的 —— 删掉两页，页内检查一个都不会响。所以新增
`qc.check(deck, source_path=...)`：拿 `.pi/work/analysis.md`（章节/模块/结论/论据 表）交叉核对
两个口径：①带 %/单位或 3 位以上的**数字口径**在 deck 全文找不到；②**模块名的 4-gram**
在 deck 里一个都不出现。再把「论据行数 ÷ 内容页数」报出来（平均每页要扛 6 条以上 =
页数明显偏少）。研讨型由 `briefs/seminar.yaml` 的 `requireSource: true` 强制要求跑
（不带 `--source` 也会被点名）。

实测：拿真实的 44 页 Pi 研讨 deck 对 `.pi/work/analysis.md` 跑，报了 18 个模块没露面 ——
这正是"页数被压缩 + 密度上不去"的根因，现在能被机器抓到了。

## 19. 结论页（statement）重排 + 层级卡吃留白

正文不再强制拆 2~4 张卡（`split_body_groups` 按**字数**分组：≤160 字单卡连续排，
超过按每 170 字一组、最多 3 组），高亮数据卡从"固定占左下"移到**右栏底部**，
正文宽度 7.35→8.46in、高度完整 —— `statement.body` 上限因此从 320 提到 700 字
（520 → 700 是 2026-09 第二轮放宽），`statement_slack()` 报"还空 X in、还能写 N 字"。
`hierarchy` 另有一条：子模块卡允许吃掉最多 `HIER_GROW_MAX=0.90` 的留白
（描述垂直居中），否则窄卡片永远填不满页面、只剩"补内容/占位"一条路。

## 20. 耗时大户不是渲染，是「回合数 × 每回合重发的上下文」

（2026-09，用户反馈"PPT 生成还是有些慢"）实测把耗时拆开看，机械部分完全不是瓶颈：

| 环节 | 实测 | 说明 |
|---|---|---|
| merge + render + qc | **1.2s / 轮** | 与页数几乎无关（30 页 vs 12 页只差 0.15s） |
| 入口固定上下文 | **65KB**（研讨）/ **57KB**（汇报） | 每回合都随对话重发一遍 |
| 一页的 DSL | ~1.3KB / 页 | LLM 必须逐字生成的**内容下限**，40 页 ≈ 52KB 输出 |
| 一轮 qc 报告 | 16.6KB → **10.0KB** | 46 条 warning，其中 13KB 原本是重复的 message 文本 |

所以耗时 ≈ **回合数 × 每回合的上下文**。旧工作流把回合数放大了三倍：

- ❌ `逐页生成：一页一次` + ❌ `生成前先 Read examples/ 里对应版式的范例`；
- ❌ `回炉：只改那一页，回第 3 步` —— 每轮只改一页，却每轮都 merge+render 一次、再吞一份报告。

现在改成：**成批生成（一个回合并排写 3-5 页）**、**范例只在开工读一次**（只在开新版式
或拿不准时才单独读）、**回炉一个回合把同类问题一次改完**、**渲染一批一次**。
每页的内容要求一个字没变，只是不再把同一份上下文重发几十遍。报告也瘦了两回：

- `render._dump_report`：`--qc` 输出改紧凑 JSON，省掉 `indent=2` 的 3.6KB 结构缩进；
- `qc._fold_rule_tails`：同一 `(type, field)` 且共享 ≥24 字后缀的重复说明，只留第一条完整版，
  其余折叠成「（规则见第 N 页）」（message 4344 → 2714 字符）。

**别再写回"一页一次 / 每页读范例 / 一轮只改一页"** —— 那是纯浪费，且不影响任何一条质量指标。

## 21. 一个模块里的「要点 + 迷你表」必须从下往上分配空间

（2026-09，用户实测 v23 P3「右侧框里的表格被挤出期望区域了」）`panel_2col` 的每个模块
= 色条 + 块头 + 顶部总结段 + （要点 + 迷你明细表）。旧实现是**自上而下**画的：
先要点、表再按"剩余空间"落位。两个地方一起出错：

1. `_panel_points` 用 `pick_step`，而 `pick_step` 在最小档也放不下时**照样返回最小档**
   （设计如此：渲染绝不失败）——于是要点越界把表一路往下顶；
2. `_panel_table` 用 `row_h = max(0.18, …)` 给行高兜底、`ty = y + max(0, (avail-total)/2)`
   **从不上收**，`total` 一旦大于 `avail` 就整块画到模块框外。

现在 `render_panel_2col` 只调 **`panel_content_geom`**（渲染与 qc 共用）：
表按内容要高度、但最多拿 `avail - 要点最小高度 - 间距`，然后**底部锚定**在内容区下沿，
`head_h + row_h×n ≡ tbl_h` —— 表在数学上不可能越框；要点只在上方剩余区间里排
（放不下就继续降字号，下限 7.5pt）。
**别改回"自上而下 + 居中兜底"**：那正是这次溢出的成因。
代价是"内容真的太多时会变成字很小"，所以 `panel_pressure()` 把这件事报给 `--qc`：
要点被压到 `PANEL_PT_FLOOR` 以下、或迷你表行高压到 `PANEL_TBL_ROW_MIN` 以下就点名。

## 22. 表格行高必须容得下文字；note 只跟表格底边走

（2026-09，用户实测 v23 P4「表格整体超出 PPT 呈现区域，底部总结句盖住表格」）
实测复现：11 行 × `TABLE_ROW_H(0.46)` = 5.56in，从 `top=2.096` 一路画到 `7.656`
（页高只有 7.5），note 被压到 6.376 正好压在表上。三处要一起记住：

- **行高是"最小值"，不是"固定值"**：python-pptx 写进 `a:tr@h` 的只是最小高度，
  PowerPoint 会按单元格文字把行**撑高**；撑开的部分没人回收。所以行高必须 ≥
  `table_row_need()`（文字自然高度 + 上下内边距），字号要**按行高反推**
  （`TABLE_CELL_STEPS` 比正文多给 9.5/9/8.5 三档），不是先定字号再砍行高。
- **`body_h` 要按可用高度压紧**：`table_geom()` 统一算 `head_h` / `body_h` / `cell_size` /
  `total_h`，`render_table` 与 `table_slack` 共用 —— 别再在渲染里另写一套（历史上
  `table_slack` 写死 `TABLE_ROW_H`，报的"超出 1.8in"和实际渲染对不上）。
- **note 的位置只由表格底边决定**：旧代码 `ny = min(max(表底, 辅区底边) + 0.08,
  BODY_BOTTOM - note_h)` —— 右侧辅区比表高时会把 note 往页面底部推，再被 `clamp`
  压在表格上。`table_geom` 保证 `total_h ≤ avail`，所以 `note_y = 表底 + 0.08`
  永远落在表格下方、且不越 `BODY_BOTTOM`，**不要再去和辅区底边取 max**。

`--qc` 新增点名：行高在最小字号档仍装不下文字（`cell_overflow > 0.02`）时报
「PowerPoint 打开时会把行撑开，整张表会溢出内容区并压住底部 note」。

## 23. Word 输入：修订记录不呈现；EMF/WMF 原样嵌入不许转格式

（2026-09，用户需求）素材是 `.docx` 时走 **`.pi/renderer/docx_in.py`**（stdlib `zipfile` +
`ElementTree`，**零新依赖**，不要引入 python-docx）：

- **修订记录不呈现**：按「接受所有修订」的口径 —— `w:ins` 保留内容、`w:del` 与
  `w:moveFrom` 整棵丢弃、`w:commentRange*/w:commentReference` 丢弃、`*Change` 忽略、
  被删的整行（`w:trPr/w:del`）跳过。别把 `w:delText` 当正文读出来。
- **图片原样**：`word/media/*` 逐字节写到 `.pi/work/media/NN-<原名>`，EMF 的物理尺寸从
  头里的 `rclFrame`（偏移 24，4×int32，单位 0.01mm）读；WMF 读 placeable 头（twip）。
- **EMF/WMF 必须原样进 pptx**：`slide.shapes.add_picture` 走 Pillow 读头，
  **Pillow 不认识 EMF/WMF**，转 PNG 再嵌必糊。`layouts._add_vector_picture` 手工挂
  `ImagePart`（content-type `image/x-emf`）+ 关系 + `p:pic`；两个前提别踩坏：
  ① `ImagePart` 构造只存字节不解析；② `_add_pic_from_image_part` 在**宽高都给全**时
  才不去问 Pillow 要原生尺寸（只给一个就会走 PIL 直接报错）。
- 版式落点是 **`figure`**（第 14 个版式）：一图 + 图注 + 可选要点右栏。`qc` 在图找不到时
  直接 **error**（一页空框比渲染失败更糟），这是 error 的第四种、也是唯一的"产物会空"
  类错误。

## 24. 字体只用系统自带：中文微软雅黑 + 西文 Times New Roman，不随包分发字体

（2026-09，用户要求）**本仓库不再带字体包**（以前 `theme.py` 的注释写着"思源黑体
`fonts/` 随包分发、部署时安装"—— 那是旧约定，已废）。现在是：

- `theme.FONT_MAIN = "微软雅黑"`（中文 / `a:ea`）、`theme.FONT_LATIN = "Times New Roman"`
  （西文 / 数字 / `a:latin`）。两个都是 **Windows 自带**（装了 Office 就有），
  所以不需要下载、不需要部署字体、也没有字体授权问题。
- **两个 typeface 必须成对写进每个 run**：PowerPoint 是按**字符**在 `a:latin` / `a:ea`
  之间挑字体的。只写 latin，中文会退回主题字体（Calibri → PowerPoint 自己替换一个），
  只写 ea，英文/数字变 Calibri。统一入口 `layouts.apply_run_font(run)`，
  顺序 `latin → ea → cs`（CT_TextCharacterProperties 的顺序约束，插在 `a:sym` 之前）。
- **自己新建 run 时一定要调 `apply_run_font`**：历史上页码（`render.py` 里手写的
  `p.add_run()`）就漏了这一步，整份 deck 里只有页码是 Calibri。
- **图表有自己的 `c:txPr` 体系**，不走 `text_box`：`_normalize_chart_xml` 里会
  `_apply_chart_fonts` 把**每一级**（图表级 / 图例级 / 轴标签级 / 数据标签级）
  的 `a:defRPr` / `a:rPr` 都补齐 —— 只改图表级会漏掉后面几级。
- **主题默认字体也要改**（`theme.apply_theme_fonts`，在 `build_pptx` 里调）：
  幻灯片文字都显式带字体，改主题是为了"默认值"也对齐（用户在 PowerPoint 里新加
  文本框/编辑时才不会拿到 Calibri）。要改三处：`majorFont/minorFont` 的 `a:latin`、
  `a:ea`，以及 **`a:font script="Hans"`**（简体中文的按脚本覆盖项**优先于 `a:ea`**，
  默认值是宋体，不改它中文就变宋体）。
- matplotlib 图表同口径（`charts._setup_font`）：`font.family` 传
  `[Times New Roman, Microsoft YaHei, …]` 做**逐字符回退**，英文命中外文、中文回落雅黑。
- 回归：`.pi/tests/smoke.py` 第 17 节 —— 幻灯片/图表里每一段文字都必须带这对字体，
  且主题 `fontScheme` 的三处都要写对。**不要再往仓库里放 `fonts/` 或 `*.ttf`**，
  也**不要**为了"号码好看"再引入 `text_box(font=...)` 之类的覆盖参数
  （历史上"号码用雅黑、正文用思源"就是这么漂移的）。

> 提醒：`.pi/wheels/fonttools-*.whl` 是 matplotlib 的依赖**库**（解析字体文件用的），
> 不是字体包，别删。

## 25. 关系路由的合并是破坏性决定，只用显式关系；邻接推断只能做建议

（2026-10，规划引擎加第二遍路由时定的约定）`doc_analyzer.extract_relations` 抽
跨模块关系时，没点名目标的关系一律在 evidence 里标**「（邻接推断）」**——
文本里只出现"相比/因此/此外"这类词、没出现对方模块名的，方向纯靠前后顺序猜。
`planner._explicit_link` 对这类推断关系**一律不参与合并决策**：把两页并成一页
是破坏性操作，猜错了就丢内容。时序链（flow）和并列组（numbered_cards）可以
用推断关系（成组本来就有"连续 ≥3 个"的门槛兜底），**对比对（panel_2col）和
包含组（hierarchy）必须显式**（点名 / 命名对仗 / 阶段序号）。另外容量闸门
`REL_GROUP_MAX_CHARS = 700` 和"任一模块要拆页就不合"优先于一切关系合并——
宁可页数多，不可塞不下。别把 `_explicit_link` 改回 `_loose_link` 一刀切，
也别把闸门调大去"硬合"。

## 26. theme 的画布常量是浮点英寸，不是 EMU；新进仓的版式必须真的渲染过一次

（2026-10 加 timeline/funnel/matrix 时挖出的存量 bug，两处一起记）

- **常量口径**：`theme.py` 的 `MARGIN_L / CONTENT_W / SLIDE_W / SLIDE_H` 曾写成
  `Inches(0.55)` 这类 **EMU 对象**。版式函数算完几何还要把结果喂给 `In(x)`，
  传 EMU 进去会被**再放大 914400 倍**（`In(MARGIN_L)` ≈ 4.6 亿英寸）——
  `render_grid_2x2` 就是这么把整页画飞了。现在这四个常量全是**浮点英寸**，
  别再改回 `Inches(...)`；要 EMU 的地方在用点现场 `In()`。
- **新代码必须真的渲染过**：grid_2x2 之前带着三处必炸的伤进了仓（常量没导入
  → NameError、`layouts.CARD_ICON_RIGHT` 自引用 → NameError、EMU 双重缩放），
  而它**从没被任何测试渲染过**——smoke 的 fixture 还是旧 14 版式、
  `examples-all.json` 里没有它。约定：**每加一个 slide_type，当天就要有
  examples/NN-*.json 并过一遍 `merge.py + render.py --qc`（0 error / 0 warning）**，
  合集缺哪个版式，哪个版式就等于不存在。

## 27. 外观是「5 套配色 × 5 种卡片」的白名单，不是自由配色

（2026-10 两轮定稿）先是用户看到棕色 demo 后定下"主色必须深蓝"，注册表砍到只剩
`style-harness`；随后（2026-10）用户明确要求**解决多人共用时"外观千篇一律"**，
于是按他的口径扩成 5 套配色 + 5 种卡片，**但仍全部是白名单、全部蓝/中性色系为骨**：

- `_THEMES` **只许这 5 套**：`style-harness`（默认、藏青）/ `blue-business`（墨蓝商务）/
  `bright`（亮蓝多色）/ `forest`（墨绿）/ `violet-tech`（靛紫）。测试锁死名单与顺序
  （`TestThemeStyles.test_registry_is_the_five_agreed_themes`），**别在实现里随手加第 6 套**。
- `CARD_STYLES` **只许这 5 种**：`flat`（默认、与历史产出一致）/ `outline` / `tinted` /
  `accent-left` / `accent-top`。同样是白名单。
- 两个都是**未登记名确定性回退**（主题→`style-harness`、卡片→`flat`），
  `apply_style` / `apply_card_style` 必须**原位**改（`COLORS.clear()+update()`、列表切片），
  因为 `layouts.py` 顶部 `from theme import COLORS...` 持有的是同一对象。
  回退时 `--qc` 会出一条 warning 点名 —— 否则作者以为换了脸、实际没换。
- **新主题必须键完全对齐**（22 个颜色键）：`layouts._c()` 是 `COLORS.get(name, name)`，
  少一个键就会把 `"cyan"` 当 hex 解析直接抛错。用 `theme._palette(**overrides)` 合并生成，
  测试逐个主题比键集合。
- **白字压色带要有底线**：强调色同时是 `layered_stack` 层标签、徽章的底色（白字）。
  新增/调整辅助色时对比度不得低于默认主题的 3.66:1（`test_white_text_on_accent_colors_is_readable`）。
  `bright` 的第一版就是踩这个坑（薄荷/珊瑚白字只有 2.6~3.0），整组压暗了一档。
- **卡片长相只许改"怎么画底"**：填充/描边/圆角/色条都能换，**内边距、文字框几何一律不许动** ——
  `qc` 的容量估算走 `layouts` 的几何函数，动了内边距报告就和现实对不上
  （`test_text_geometry_never_changes` 逐页比对文本框 left/top/width/height/字号）。
  版式自己画的那条顶色带（0.05~0.075）是"版式语义"，别删；`accent-top` 的顶带做到 0.09
  并让调用点传 `accent=key`，两者同色叠成一条更厚的带子，才不会出现双色条。
- **底图与字体不随主题变**：`assets/template/bg-cover.png` / `bg-content.png`（封面深蓝渐变、
  内容页顶部蓝条 + 院所 logo + 水印）是固定素材，任何主题都不许替换或覆盖；
  字体仍是「中文微软雅黑 + 西文 Times New Roman」一套。换主题只换点缀色。
- 演示品不进交付目录：`out/` 只放正式产出，实验产物用 `_` 前缀或放临时目录
  （样张放 `.pi/work/_preview/`）。

## 附：关键长度陷阱（schema 上限 ≠ 几何容量）

一份"现行口径"摘要（不是坑，是常量与标准的当前位置；改密度/改阈值时对着它改）。

**容量不再是写死的常数**：渲染器会按内容逐档降字号（`FIT_STEPS` 11.5→10pt，
层级图 `HIER_STEPS` 9.5→8.5），所以"能写多少字"由 `layouts` 的几何函数算出来，
`--qc` 直接报「约 X 字」。改密度时**别再来回改一堆魔法数字**，改几何就行。

历史写死的安全值（供参考，已被上面的动态口径取代）：
numbered_cards.desc 80、layered_stack.desc 55、flow.desc 40、hierarchy.node.desc 60。
注意 layered_stack 的辅区在右栏 → 条形收窄，描述实际只有 **4.0in** 宽。

**下限同样是红线**（写少了 `--qc` 点名"精简过头"）：statement.body ≥150、
compare 每栏 desc ≥45、numbered_cards 每卡 desc ≥40、layered_stack 每层 desc ≥30、
flow 每步 desc ≥16、hierarchy node.desc ≥20、panel_2col 每模块 summary ≥16。
字段应写到**接近上限**，下限只是防"抽干"的最后一道。

**结论段是 40~80 字 / 2~3 句**（`qc._SUMMARY_*`）：页头是每页最贵的版面，
结论段只留「判断 + 关键数字」，机制细节下放卡片。footer_note ≤ **64 字**（一行）。

**"还剩多少留白"也由几何函数算**：`table_slack` / `chart_bottom` / `panel_slack` /
`statement_slack` / `hierarchy_metrics` 都会报「下方还空 X in，还能写 N 字」——
遇到这种 warning **去补内容或并页**，不要拉高行高、也不要缩小字号硬撑。

**容器是硬的（第 21/22 条）**：`BODY_BOTTOM=6.80` 之下不许有内容；`panel_2col` 的模块里
`head_h + row_h×n ≡ tbl_h`、表**底部锚定**；表格页 `total_h ≤ avail` 且
`body_h ≥ table_row_need(cell_size)`。可读下限：`PANEL_PT_FLOOR=9.5pt`、
`PANEL_TBL_ROW_MIN=0.16in`、`TABLE_CELL_STEPS` 最低 8.5pt —— 低于这些 `--qc` 会点名，
处理方式是**删条目/并页**，不是继续压字号。

## 附二：v23 几何大修 —— 这几处口径变过，改代码前先读

这一轮把「字段上限 / 几何 / 质检」三处口径统一到**同一个函数**，期间踩到的坑记在这里，
免得后面又写出第二套数字。

**1. `pick_step` / `fit_size` 现在会告诉你"放不下"。**
新增 `pick_step_ex` / `fit_size_ex` 返回 `(size, spacing, fits)`；`fits=False` 表示
连最低档都装不下。渲染侧不再假装成功，`--qc` 把这种情况点名成"要么删内容、要么拆页"。

**2. schema 上限改成按条目数动态校验（ppt_dsl._cap_hint）。**
静态 `max_length` 只能按"最宽松情形"给（4 张卡 240 字 vs 8 张卡 90 字）；
现在各 Slide 的 `model_validator` 直接调 `layouts.*_geom(self)` 反算容量。
**带容差**（`CAP_TOL=1.10` / `CAP_SLACK=8`）：字宽是保守估算，差 4 个字就拒稿是负收益。
新增/调整版式时，记得同步写一个 `fit_capacity`，否则这个版式又变成"文档说一套、几何做另一套"。

**3. `timeline` 节点 ≥7 个改竖排（`TL_VERTICAL_AT`）。**
横排是按节点数均分宽度：12 个节点时每卡 0.85in、描述区 0.49in，而 schema 允许 120 字
（实测要 5.6in 高）—— 必然横穿整页。竖排一行一个节点（时间在左、卡片在右），
12 行容量约是横排的 3 倍。`timeline_geom` 返回 `mode` 字段，**qc 必须按 mode 取键**。

**4. 奇数条目末位跨整行（statement / key_groups / toc 共用这套逻辑）。**
2 列网格在 3 组时右下空一整个卡位（1/4 页）。现在末行只剩 1 个时让它跨整行，
**行高按跨行后的宽度算**（宽一倍 → 一行装得下更多字 → 行高反而更矮）。
算高与绘制必须都走"这一行实际多宽"，否则末卡会按半宽估高、卡片下方空一条。

**5. 右栏卡片栈要"消化"余量，不是堆在顶部（`_fill_rail`）。**
旧实现把 pad 封在 0.20、gap 封在 0.30 就收手，4.5in 的右栏只堆了 2in 卡片、
下面 2in 纯空气。现在四步：内边距 → 间距 → 适度长高（≤1.45×） → 剩余上下均分（垂直居中）。
卡片**等高**（统一到最长那个的自然高度）—— 高低不齐比留白更难看。

**6. `matrix` 的格子说明起点必须让开块头图标（`MATRIX_DESC_TOP=0.58`）。**
从 0.42 起时，首行文字会和右上角图标（0.09–0.57）撞在一起，3 行格子特别明显。

**7. `grid_2x2` / `end` / `highlight.desc` 补了几何与检查。**
之前：grid_2x2 只查"值非空"（没有任何容量概念）、end 页完全在 qc 视野外
（4 段各 200 字会把第 4 张卡推到 8.75in，页高只有 7.5in）、高亮卡 desc 无容量检查
（卡高写死 1.10in，而 schema 允许 80 字）。现在分别走 `grid_geom` / `end_geom` / `highlight_h`。

**8. `table.hero_col` 取代了硬编码的 "Pi" 列识别。**
旧代码自动匹配表头里的 `Pi` / `pi-ai` / `Pi ` 前缀 —— 那是某一份 deck 的产品名，
换个行业永远匹配不上。现在由 `hero_col` 显式指定（填表头文字），不填就不高亮。
同时 `TableSlide.validate_row_width` **以前忘了加 `@model_validator`**，
是个从没执行过的普通方法（列数写错会静默错位），现已生效。

**9. 字号与颜色收进 `theme`。**
`layouts.py` 里 40+ 处 `size=11` / `color="FFFFFF"` 全部换成 `FONT_SIZES[...]` /
`COLORS["on_dark"]`。同类元素**同一档字号**（右栏 label 10 / 正文 10.5、徽章 11、
小标签 9.5、环形节点 7.5）。新增版式时别再写字面量。

**10. 规划链路（plan → merge --from-outline → render）现在是通的。**
`outline._shell` 为 19 个版式各生成"必填字段齐全、正文全是「待补：」"的骨架；
`merge` 接受 `-o`（以前只认位置参数，而 plan.py 打印的下一步正是 `-o`）；
`--with-notes` 改为**逐页**占位（原来只 append 有备注的页，导致备注整体前移贴错页）。
`smoke.py` 第 18/19 节锁住这两条：骨架能渲染、动态容量会拦。

## 28. 慢的不是渲染，是「每个产物都要一次模型往返」

（2026-10 提速改造）实测把口径钉死：`import render` 0.89s（matplotlib 占大头）、
4620 个图标扫描 7ms、`analyze+plan` 近 0s、30 页渲染约 1 秒 —— **代码侧没有可优化的空间**。
真正的成本是**工具调用次数**：每次 `write_to_file` / `replace_in_file` / 跑一条命令，
都是一次完整的模型往返，而且每次都要带上"已经写了什么"，40 页按一页一文件写就是
40 次往返、输入近似 O(N²) 增长。现在的口径：

- **写页按批**：`slides/NN-batch.json` 里放**数组**，一个文件 5-8 页。`merge.py` 的
  `collect_slides()` 两种都吃（一个文件一页 / 一页数组），按文件名排序后展开。
  一批太大就降到 3-4 页 —— 单批输出超长时模型会截断，回炉反而更贵。
- **出图一条命令**：`build.py`（合并 → 质检 → 渲染 → 字体自检）一次跑完、一份报告、
  一个退出码。别再把 `merge.py` / `render.py --qc` / `font_audit.py` 拆三次敲。
- **回炉按批重写**：改一个 batch 文件一次写，比逐页 `replace_in_file` 少 N-1 次往返；
  只有改一句话才用 `replace_in_file`。
- **分析表只补两列**：`plan.py --emit-analysis` 从原文自动切章节/模块/论据/数字并写成
  `analysis.md` 骨架，「核心结论 / 缺口」才需要人或模型填。骨架的「章节/模块」可以
  合并改名，但**不能删要上页的模块**（`--source` 覆盖度检查按这张表核对）。
  `analyze_raw_text` 之前用带捕获组的 `re.split` 切标题，拿到的是 `"#"`/`"##"` 标记本身，
  `section` 永远为空、标题还混进正文 —— 已改成按行解析，别再改回去。
- **中间产物只留 5 类**：`analysis.md`、`outline.json`、`slides/*.json`、`deck.json`
  （`build.py --keep-deck` 才生成）、`.pptx`。`qc_report.txt` / `qc_report2.txt` /
  `render_report.json` 这类"把 stdout 重定向一遍"的副本**不要再产生** —— 报告就在 stdout，
  `build.py` 一次全打出来。
- **报表口径不变**：`build.py` 打的是同一份 `qc.check()` 结果、同一份紧凑 JSON
  （`render._dump_report`），所以"读报告 → 回炉"的语义没变，只是少了两次进程启动与往返。

## 29. 动态容量要写前下发（budgets.py），不是写后靠报错回炉

（2026-11，"长度类校验返工"的根治）动态容量（`_cap_hint` + 各版式几何）原来只在
生成**之后**由 schema 反算报数——模型写作时看不到"8 张卡每卡 90 字"这类随条目数
浮动的真实容量，顶格写静态 `max_length` 就撞墙回炉。现在 `budgets.py` 用**拦截式
探测**（接管 `ppt_dsl._cap_hint` 构造合成页，预算数与校验数出自同一段代码）把容量
前置成逐页预算表，写批次前读表、按目标的 80~90% 写。三条别改回去：

- **别删预算环节、别把 SKILL.md 的动态数字抄回来**：表内"4 张≈240、7-8 张≈90"
  这类手抄数字曾同时漏掉"带不带辅区"两种情形（实测 8 卡无辅区可写 120、带辅区
  才是 90）——动态数字只活在预算表和 qc 报告里，文档里只写"查预算表"。
- **别去掉拦截探测改手写几何调用**：拦截保证预算与校验零漂移；`key_groups` / `end`
  两处没走 `_cap_hint`，在 budgets.py 里按校验器同款公式复算，改那两个校验器时
  要同步改 budgets.py（已注明）。
- **几何改动后跑 `python .pi/renderer/budgets.py --selfcheck`**：15 个代表性配置
  "照预算写满 → 真实校验必放行"，几何一变它就是回归线。

