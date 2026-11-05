# ppt-common 范例（few-shot）

这里是**最低质量线**——注意是"线"不是"上线"：照着范例的
"判断句标题 / 信息密度 / 比喻 / 图标 / 底部收束"来写，
内容换成你手头文档的真实信息，密度**只能更高不能更低**（不要照抄示例主题）。

> **开工时读一次 `../examples-all.json` 就够**（18 个版式全在里面）——
> 只有开一个**新**版式、或对某个版式拿不准时，才单独 Read 这里对应的一个文件。
> 「生成一页就读一个范例」是旧工作流，40 页 deck 会白搭 40 次读盘与上下文。

- `01-cover.json` … `18-key_groups.json`（+ `99-end.json`；其中 `11-grid_2x2`、
  `15~17`（timeline/funnel/matrix）、`18-key_groups` 是 2026-10 新增的关系语义版式）
  是**逐页范例**，每个文件是一页 DSL，与"逐页生成"工作流一一对应。
  `99-end` 特意排到最后，合并出来的顺序才是「封面 → … → 结尾」。
- 这些页合起来就是一份完整、连贯的 deck（主题：智能体评测平台），
  可以直接 `merge.py` + `render.py --qc` 渲染验证 —— **这套范例自己就是 0 error / 0 warning**，
  照这个标准写。

## 自检（渲染前必过）

```bash
# 合并成整份 deck 并校验 schema
python .pi/renderer/merge.py .pi/skills/ppt-common/examples /tmp/_ex_deck.json --title "智能体评测平台建设方案"

# 渲染 + 质检（有 error 会阻止渲染）
python .pi/renderer/render.py /tmp/_ex_deck.json -o /tmp/_ex.pptx --qc
```

两份都 `OK`、`errors` 为空，才说明范例本身是对的。

## 密度与留白（最常犯的三个错）

1. **一页两块内容 → 左右分置**：`13-panel_2col.json` 就是模板 —— 左右两个模块，
   各自一个背景框，框内**顶部一段总结陈述段**，下面是要点或迷你明细表。
   不要上半页一块、下半页一块，那样下半页必然空着。
2. **内容少就并页**：表格 < 5 行、图表数据点 ≤ 4 且补不出明细表，
   就把相邻页同主题的内容并进来（并成 `panel_2col` 也行）；
   `09-table.json` 是 6 行 + 右侧辅区，`10-chart.json` 是 4 个数据点 + 下方明细表。
3. **正文不许精简过头**：statement.body ≥ 150 字、compare 每栏 desc ≥ 45、
   numbered_cards 每卡 desc ≥ 40、layered_stack 每层 desc ≥ 30、flow 每步 desc ≥ 16、
   hierarchy 每个 node.desc ≥ 20、panel_2col 每模块 summary ≥ 16、
   timeline 节点 desc ≥ 25、funnel 层 desc ≥ 20、matrix 格子 desc ≥ 15。
   范例里的描述都是"接近上限"的写法，不要写成标签。
4. **卡片不许写成一坨段落**：`04-statement.json` 就是模板 —— `body` 用**换行分卡**
   （≤3 张），**每张卡的首句（≤16 字）会被自动提为卡头标题行**（加粗染色的组标题，
   所以首句要写成判断句）；卡内长文标 2~3 处 `**关键短语**：正文`；
   `06-numbered_cards.json` 的 `desc` 同样这么写。
   注意是「**同字号** + 加粗 + 强调色」—— 渲染器不会放大关键短语（放大 + 另起行会把
   文字顶出卡片背景），冒号后的正文就在同一段里接着排。
   **分组的正身是 `18-key_groups.json`**：2-4 个平行主题组、每组标题+要点清单，
   别把能分组的内容摊进 statement 兜底（那是字堆的来源）。

## 加粗与图标（两个容易做过头的地方）

- **加粗只标关键词**：用 `**…**`，每页 **1-2 处**、每处 **≤16 字**。
  渲染器**不会**自动把首句加粗，所以 `statement.body` 不标就等于没有重点；
  整句/大段加粗会被 `--qc` 报 warning —— 整句加粗等于没加粗。
- **一个内容块只有一个图标**：卡片/栏/层/步骤/节点/洞察/panel 各配一个 `icon`，画在块头。
  块内的大段正文前面**不再**放图标（渲染器也不会再自动补第二个）。
  没写 `icon` 的块会按文字自动补一个，但别依赖回退。

## 关键长度陷阱（重复强调）

- **容量是动态的**：渲染器会按内容逐档降字号（正文 11.5→10pt），`--qc` 直接报
  「约 X 字 / 已降到 X pt」。所以**别背魔法数字**：写满，超了它会告诉你还能写多少。
- 各内容页 `summary` **40~80 字 / 2~3 句**（只留判断 + 关键数字，机制下放卡片）；
  `footer_note` ≤ **64 字**（页脚条只有一行）。
- 各版式正文下限（写少了 `--qc` 点名"精简过头"）：statement.body ≥ 150、
  compare 每栏 desc ≥ 45、numbered_cards 每卡 desc ≥ 40、layered_stack 每层 desc ≥ 30、
  flow 每步 desc ≥ 16、hierarchy node.desc ≥ 20、panel_2col 每模块 summary ≥ 16。
- **卡片标题即重点**：`numbered_cards[].name` 写判断句/带数字的结论
  （「统一任务集管理」「回归通过率 62%→88%」），不是名词短语（「任务集管理」）；
  标题已是判断句就**省掉 `sub`**（`sub` 已改为可选）。
- table 每行长度必须等于 headers 长度；`panel_2col.panels[].rows` 同理。
- `panel_2col` 恰好 2 个 panel，每个 panel 至少要有点要或迷你表。
