# 冗余审计（2026-10-22）

> 注：本文路径已按同日完成的目录调整（工程主体移入 `.pi/`）统一改写；
> 去掉 `.pi/` 前缀即为调整前的原位置。

针对 `.pi/renderer/`（10.5k 行、16 个模块）、`.pi/assets/icons/`（9931 个文件）与仓库根目录资产做的一次静态+动态盘点。
每条都给了定位，判据是"能不能被验证"，不是"看起来像"。

---

## 一、先说结论

| 层面 | 冗余程度 | 主要形态 |
|---|---|---|
| 布局/质检代码（`layouts.py` / `qc.py`） | **低** | 3068 行的几何计算几乎全是各版式专用，`qc` 复用 `layouts` 的几何函数而不是另背一套数字 —— 这是本项目做得最对的地方，别动 |
| 规划链路（`doc_analyzer` / `planner` / `outline` / `plan`） | **中** | 枚举与解析逻辑各存一份 |
| 图标引擎（`iconlib.py` / `icons.py` / `aliases.json`） | **高** | 资产层面 4 库重复；代码层面两套「中文→图标」映射 |
| 仓库根目录 | **高** | `packages/` 是 `.pi/wheels/` 的完整副本（81 MB） |

一句话：**代码本身冗余不多，冗余集中在"同一份知识存了两个地方"和"资产整库复制"**。

---

## 二、代码冗余（按证据强度）

### 2.1 同一份知识存两份，且已经不一致

| # | 位置 | 问题 |
|---|---|---|
| 1 | `ppt_dsl.py:43-75` 与 `ppt_dsl.py:475-507` | `QuadrantCell` / `QuadrantGridSlide` **各定义了两次**，后半段覆盖前半段，前一份是死代码。两份只差 `slide_type` 有没有默认值 |
| 2 | `ppt_dsl.py:17` 的 `SlideType` vs `outline.py:29` 的 `SlideType` | 同名枚举存了两份，15 个值逐条手抄；`ppt_dsl` 那份**全仓库无人使用**，`outline` 那份被 planner/tests 用 |
| 3 | `layouts.py:139` `_ICON_KEYWORDS`（39 组 ≈180 个中文词）vs `.pi/assets/icons/aliases.json`（230 条） | 两套「中文词→图标名」映射，取值还不一致：`安全` 别名表给 `shield-lock`、自动映射给 `shield-check`；`协作` 一边 `users-group` 一边 `users`；`数据` 一边 `database` 一边 `chart-bar`。同页里"作者写的图标"和"没写时自动补的图标"会走两条不同的规则 |
| 4 | `doc_analyzer.py:178` `_parse_analysis_table` vs `qc.py:678` `_source_rows` | 解析 `analysis.md` 那张 markdown 表的循环逐行相同（`|` 开头、`cells<3` 跳过、`章节` 表头跳过），qc 为了做输入覆盖度又写了一遍。表格式一改就得改两处 |
| 5 | `qc.py:154` `_cjk_count` vs `qc.py:478` `_cjk_len` | 函数体一模一样（都数 `\u4e00-\u9fff`），只是名字不同 |
| 6 | `doc_analyzer.py:287` `available_briefs()` | 一行包一层 `briefs.available()`，且无人调用；`doc_analyzer` 因此反向依赖 `briefs` |

### 2.2 死代码（全仓库 0 引用）

| 位置 | 符号 |
|---|---|
| `layouts.py:964` | `_block_icon_size` |
| `layouts.py:1121` | `compare_desc_h` |
| `layouts.py:1606` | `hierarchy_contain_h` |
| `doc_analyzer.py:267` | `modules_to_json` |
| `doc_analyzer.py:271` | `modules_from_json` |
| `iconlib.py:308` | `search_icon`（与 `search_icons` 同功能两套 API） |
| `theme.py:163` | `PT` |

`icons.py` 整个模块（37 行）只是把 `iconlib` 的三个函数再转一手，`icon_info` / `search_icons` 两个门面函数无人调用 —— 留 `draw_icon` 一个就够。

### 2.3 结构性重复（可合并但不算错）

- `doc_analyzer.py:83-113`：`has_comparison` / `has_process` / `has_hierarchy` / `_is_data_heavy` / `_is_text_heavy` 五个函数，骨架相同（关键词表 + `any/all`），只差词表。合并成"一张特征表 + 一个函数"能少 30 行。
- `docx_in.py:73` `image_size_in` 与 `layouts.py:2143` `_blob_size_in`：两个"从图片字节读尺寸"的实现，一个给 docx 抽图用、一个给 pptx 嵌图用。
- `charts.py`（matplotlib 路线）与 `layouts.py:2948` 之后的原生图表路线并存。**这条不算冗余**：`--chart-mode` 是有意的双通道，原生可编辑 vs 高质感渲染，别合。

---

## 三、仓库层面的死重

| 项 | 证据 |
|---|---|
| `packages/`（39 个 wheel，81 MB） | 与 `.pi/wheels/` 的 39 个文件**文件名完全相同**，是完整副本；`setup.bat` 只认 `wheels\`，全仓库无任何引用 |
| `.pi/out/~$*.pptx` × 3 | Office 打开文件时留下的锁文件（各 165 字节），不是产物 |
| `.pi/assets/icons/README.md` | 仍在教人跑 `tools/sync_icons.py`、`tools/check_icons.py`、`tools/icon_sheet.py`、`tools/verify_holes.py` —— `tools/` 已在 2026-09 清理时整体删除，四个命令全部失效 |
| `AGENTS.md:41` | 写的是 `.pi/renderer/assets/icons/`，实际路径是 `.pi/assets/icons/` |
| `.pi/assets/icons/index.json` | 1.5 MB，占图标库体积的 23%（图标本身才 16 MB）；它只提供关键词标签，解析走的是磁盘扫描 |
| `.pi/assets/icons/heroicons/solid-16` + `solid-20` | 640 个文件，是 `solid/` 同一批 324 个图形的 16/20 网格另存版；缩放在渲染器里已经做过 |

---

## 四、图标库：冗余在哪（详见 `.pi/docs/icon-audit.html` 样张）

- **9931 个文件只表达 6911 个名字**，其中 2350 个文件的裸名被 `resolve_order` 挡住，永远解析不到。
- **同概念 4 份**：搜 `database` 的前 4 条是 tabler/lucide/feather/tabler-filled 的同名图标，30 条结果里 11 条是跨库重复。
- **视觉语言不统一**（这是"粗糙"的根因，不是审美问题）：
  - 笔画：tabler / lucide / feather 统一 `stroke-width=2.0`，heroicons 是 `1.5` → 同尺寸下 heroicons 细 34%；
  - 占框：tabler 平均 16.9 / 24，lucide 18.3，feather 18.4 → 同名义尺寸下 tabler 小 8%。
  混用时同一页的图标会一半粗一半细、一半大一半小 —— 这比"图标不够漂亮"更伤观感。
- **实际使用面极小**：全仓库（含 14 页范例、历史 deck）只引用过 **86 个不同的名字**，99.1% 的图标从没被用过。
- **AI 俗套是别名表自己埋的**：`智能`/`能力`/`亮点 → sparkles`、`模型`/`推理`/`大脑 → brain`、`机器人 → robot`。模型照着别名填，页面自然全是星芒和大脑。

裁剪方案与执行工具：`.pi/tests/icon_trim.py`（默认 L2，`--apply` 才动文件，删除物移入 `.pi/work/_icon_trash/` 可回滚）。

---

## 五、本轮已执行（2026-10-22）

用户选择：图标 **L2 单库收敛**；代码 **只修确定的**。

### 图标库

| | 前 | 后 |
|---|---|---|
| 文件数 | 9931 | **4620**（−53%） |
| 库 / 风格 | 4 库 / 9 风格 | **1 库 / 2 风格**（tabler outline 3646 + filled 974） |
| 体积 | 18 MB（索引 1.5 MB） | **8.7 MB**（索引 0.59 MB） |
| 别名 | 222 条 | 222 条（不增不删：15 条的**目标**改了，见下） |

删除构成：粗糙 2078 · 重复 1669 · 噪声 1533 · AI 味 31。
逐条理由见 `.pi/work/icon-trim-record.json`，被删文件在 `.pi/work/_icon_trash/`（可回滚，确认后整目录删除）。

连带做的两件"不只是删"的事：

1. **别名表去 AI 味**：`智能/能力/亮点 → sparkles`、`模型/推理/大脑/brain → brain`、`机器人 → robot`
   全部改指 `cpu / target / cube / diamond / user-cog`。不改这里，删掉图标也会被别名重新引回来。
2. **Lucide 显式引用清零**：`甘特/算法/论文/质量/价值` 五个原来指向 lucide 的别名改指 tabler，
   这样 lucide 才能整库删干净。

### 代码

| 改动 | 文件 |
|---|---|
| 删掉重复定义两遍的 `QuadrantCell` / `QuadrantGridSlide`（保留后一份） | `.pi/renderer/ppt_dsl.py` |
| 删掉无人使用、与 `outline.SlideType` 重复的 `SlideType` 枚举 | `.pi/renderer/ppt_dsl.py` |
| `_cjk_count` / `_cjk_len` 合并为一个 `_cjk_len` | `.pi/renderer/qc.py` |
| 删死代码：`_block_icon_size`、`compare_desc_h`、`hierarchy_contain_h` | `.pi/renderer/layouts.py` |
| 删死代码：`modules_to_json`、`modules_from_json`、`available_briefs` | `.pi/renderer/doc_analyzer.py` |
| 删死代码：`search_icon`（保留 `search_icons`） | `.pi/renderer/iconlib.py` |
| 删死代码：`PT` | `.pi/renderer/theme.py` |
| `icons.py` 收敛到只剩 `draw_icon`（删 `icon_info` / `search_icons` 两个转发） | `.pi/renderer/icons.py` |
| `plan.py` 改为直接 `briefs.available()`，去掉对 `doc_analyzer` 的反向依赖与 `hasattr` 试探 | `.pi/renderer/plan.py` |
| `test_new_features.py` 的版式清单断言改为读 `ppt_dsl.Slide` 的判别字段（不再依赖已删的枚举） | `.pi/tests/test_new_features.py` |
| 删 `packages/`（39 个 wheel，81 MB，与 `.pi/wheels/` **逐字节相同**，`setup.bat` 只用 `.pi/wheels/`） | 仓库根 |
| 删 `.pi/out/~$*.pptx` × 3（Office 锁文件） | `.pi/out/` |
| 修 `AGENTS.md` 的图标路径（`.pi/renderer/assets/icons/` → `.pi/assets/icons/`）、补 `iconlib.py` 一行 | `AGENTS.md` |
| 重写 `.pi/assets/icons/README.md`（4 库 9931 → 1 库 4620；删掉 4 个指向已删除 `tools/` 的命令） | `.pi/assets/icons/README.md` |
| 新增裁剪工具（L1/L2/L3 + 保护集 + 执行后校验） | `.pi/tests/icon_trim.py` |

### 回归结果

```
.pi/tests/smoke.py      PASS（14 版式渲染 + 字体口径 + 越框检查）
.pi/tests/briefs.py     PASS
.pi/tests/chart_xml.py  PASS（含 PowerPoint 真机打开环形图）
.pi/tests/docx_in.py    PASS
.pi/tests/font_audit.py PASS（.pi/out/ 里的既有 deck）
.pi/tests/test_planner.py + test_new_features.py   59 个用例全过（见下方说明）
```

> **顺带发现（未修）**：`.pi/tests/test_planner.py` 与 `.pi/tests/test_new_features.py` 依赖 `pytest`，
> 但 `.pi/wheels/` 里没有 pytest，离线环境根本跑不了这两个回归 —— 它们引到 pytest 的地方只有
> 文件末尾的 `pytest.main`，等于只用到了入口。要么把 pytest 补进 `.pi/wheels/`，
> 要么把这两个文件改成和 `smoke.py` 一样的纯脚本自跑。本轮是用一个 pytest 桩手动跑通的。
