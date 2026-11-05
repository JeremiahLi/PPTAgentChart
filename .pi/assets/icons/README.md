# assets/icons —— 离线矢量图标库

**4620 个图标 / 1 个库（tabler）/ 2 个风格**，全部离线随包分发，无任何 CDN 或在线引用。
引擎把它们转成 PowerPoint 自定义几何图形（`<a:custGeom>`）：真矢量、可改色、可缩放，PPT 内是原生图形而非图片。

尺寸：约 8.7 MB，其中索引 `index.json` 占 0.6 MB。

> **2026-10 收敛说明**：原先是 4 库 / 9 风格 / 9931 个。四库并存的代价不是磁盘，而是
> **同一页的图标粗细与大小时不一致**（笔画 2.0 vs 1.5、占框 16.9/24 vs 18.4/24），
> 以及检索时同一概念返回 4 份。收敛到 tabler 单库后：9931 → 4620。
> 删除口径、执行工具见 `.pi/tests/icon_trim.py`，逐条理由见 `.pi/work/icon-trim-record.json`，
> 被删文件在 `.pi/work/_icon_trash/`（确认无问题后可整目录删除）。

---

## 1. 目录结构

```
assets/icons/
├── index.json          全量索引：风格 / 渲染模式 / 关键词标签
├── aliases.json        语义别名：中文词 -> 图标规格（人工维护）
├── README.md           本文件
├── LICENSES/           各库原始许可文件（交付留证用）
│   ├── tabler.txt   （lucide / feather / heroicons 的许可随库删除保留，见第 7 节）
└── tabler/
    ├── outline/        3646 个  描边型
    └── filled/          974 个  填充型
```

## 2. 命名规范

- **文件名**：小写 kebab-case，如 `academic-cap.svg`。
- **目录即分类**：`<库>/<风格>` 两级，风格名固定为 `outline` / `filled`。
- **同库内重名**：`outline` 与 `filled` 允许同名（如 `star.svg` 两处都有），靠目录区分；
  解析时 `outline` 优先。

## 3. 渲染模式（描边 / 填充）

渲染模式**不靠目录名猜**，由引擎实测 SVG 根属性决定，并记录进 `index.json`：

| SVG 特征 | 模式 | DrawingML 输出 |
|---|---|---|
| `fill="none"` + `stroke` + `stroke-width` | 描边 | `<a:noFill/>` + `<a:ln w=…>` |
| `fill="currentColor"`（无 stroke） | 填充 | `<a:solidFill/>`，不描边 |

三点实现细节（前两点都是踩过坑的）：

1. **合并的粒度是「源 `<path>` 元素」，不是整个图标**。
   - 同一个源元素内的子路径 → 合并进**一个** `<a:path>`。它们共享 winding，孔才挖得出来。
     反例：`tabler/filled/circle-check` 拆成两个 `<a:path>` 后，PowerPoint 逐个并集填充，对勾的孔被填死。
   - 不同源元素 → **各自一个** `<a:path>`，由 PowerPoint 取并集，对应 SVG「每个 `<path>`
     独立填充再叠加」的语义。反例：`tabler/filled/arrow-big-left-line` 的三个元素若并进
     一个 `<a:path>`，箭头主体与内层箭头绕向相反，会在主体上凭空多挖一个洞。
2. **`fill-rule="evenodd"` 需要转换**。DrawingML 没有 fill-rule 属性，引擎在**元素内部**
   按子路径的嵌套深度做奇偶绕向反转，把 even-odd 语义等价转换成正负 winding。
3. **已知局限**：若单个子路径自身相交（一笔画成的粗笔画），even-odd 与 nonzero 的差异
   无法靠「整条子路径翻转」表达，需要拆分自相交点。这条局限对应的两个样本
   （`heroicons/solid/at-symbol`、`heroicons/solid/cursor-arrow-ripple`）已随 heroicons
   库删除；现在 974 个填充型图标全是 tabler，未发现自相交样本。
   本轮裁剪**没有改动任何 SVG 内容**，只做筛选，custGeom 转换路径未变。

## 4. 名称解析规则

写入 DSL `icon` 字段的写法，按顺序尝试：

| 写法 | 解析方式 | 例子 |
|---|---|---|
| `库/风格/名字` | 精确路径 | `tabler/filled/star` |
| `库/风格:名字` | 精确路径（冒号写法） | `tabler/filled:star` |
| `库:名字` | 在指定库内查找（`outline` 优先） | `tabler:brain` |
| `名字` | 按 `index.json` 的 `resolve_order` 依次查找 | `brain` |
| 中文词 | 先查 `aliases.json` 语义别名 | `齿轮` |

默认解析优先级（`resolve_order`）：

```
tabler/outline → tabler/filled
```

先描边型是为了正文小图标视觉更轻；找不到时再退到填充型。
全部未命中 → 在图标名里做子串检索（取最短的）；仍未命中 → 静默忽略，页面照常渲染。

**填中文别名时优先用 `aliases.json` 里的词**，而不是自己拼英文名 —— 别名表里的目标是
逐个校验过能画出来的；拼错的名字会让那一块静默掉图标（见 `.pi/docs/pitfalls.md` 第 9 条）。

## 5. 命令行

```bash
python .pi/renderer/iconlib.py 齿轮                  # 中文别名检索
python .pi/renderer/iconlib.py --search brain        # 英文名 / 标签检索
python .pi/renderer/iconlib.py --search 安全 --lib tabler
python .pi/renderer/iconlib.py --json --search data  # 结构化输出（供程序消费）
python .pi/renderer/iconlib.py --show 齿轮            # 看某名字最终解析成哪个图标
python .pi/renderer/iconlib.py --libs                # 列出库与风格
python .pi/renderer/iconlib.py --stats               # 总数 / 索引时间 / 别名数 / 分组统计
```

## 6. 维护

```bash
# 裁剪：默认只出计划（不删文件），--apply 才执行并重建 index.json
python .pi/tests/icon_trim.py                        # L2 计划（推荐档）
python .pi/tests/icon_trim.py --level l1             # 只删跨库重复
python .pi/tests/icon_trim.py --level l3             # 只留业务语义白名单
python .pi/tests/icon_trim.py --apply --verify-render # 执行 + 渲染范例合集自检
```

`icon_trim.py` 的保护底线：`aliases.json` 的每个目标、技能与 deck 里真实出现过的每个名字、
`layouts` 的兜底图标，一条都不删；执行后会逐个校验别名与 deck 引用还能解析。

> 早期还有 `tools/sync_icons.py`（从外部图标包重建索引）、`tools/check_icons.py`（索引与磁盘一致性）、
> `tools/icon_sheet.py`（视觉回归样张）、`tools/verify_holes.py`（填充区栅格比对）四个脚本，
> 已随 `tools/` 目录在 2026-09 清理时删除。现在索引由 `.pi/tests/icon_trim.py` 维护。

## 7. 许可

| 库 | 许可 | 现状 |
|---|---|---|
| Tabler Icons | MIT | **3646 outline + 974 filled（当前唯一在库的库）** |
| Lucide | ISC | 已删除（2026-10，与 tabler 同名 605 个、笔画与占框比例不同） |
| Feather | MIT | 已删除（2026-10，同名 184 个） |
| Heroicons | MIT | 已删除（2026-10，同名 324 个 + solid-16/20 网格副本） |

四条的许可文件仍留在 `LICENSES/`：被删的文件如果日后恢复，许可凭据要一起在。
Tabler 为 MIT、可商用。
