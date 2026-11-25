---
description: 把文档做成演示文稿
argument-hint: "<文档路径> [研讨型|汇报型] [经典藏青|深蓝商务|明亮宣讲|墨绿战略|紫灰科技] [快|慢]"
---
用 ppt-entry 技能做一份演示文稿。素材文档：$ARGUMENTS

先确认文档存在；然后按 ppt-entry 的流程走。

**三个维度都有默认值，用户没提就别问**：类型=研讨型、风格=经典藏青、速度=慢。
只有 $ARGUMENTS 里（或对话里）明确出现对应说法时才切换，否则按默认值一路往下做。

定了维度再加载 `ppt-common` + 对应方向技能（`ppt-seminar` / `ppt-briefing`），
然后按顺序跑（每一步都会核对上游，跳步会被拦）：
`pptx.py analyze <文档>` → `pptx.py outline --title ...`（慢档给用户过目）→
**成批生成（一个文件 5-8 页，别一页一个文件）** → `pptx.py build -o .pi/out/主题.pptx`
（三个维度显式写进命令行）→ 回炉（慢档到 warning 归零）→ 交付。

不确定做到哪了：`python .pi/pptx.py`（不带参数）会列出进度和下一步该跑哪条。

全程用 `engine/` 出图。
