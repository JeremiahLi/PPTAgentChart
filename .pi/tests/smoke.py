# -*- coding: utf-8 -*-
"""离线回归冒烟测试（纯 stdlib + renderer，`python .pi/tests/smoke.py` 即可跑）。

它锁定五件不该倒退的事：
  1. 14 种 slide_type 全都能过 schema + 几何质检 + 渲染；
  2. 内容放不下时**渲染器自己降档**，`--qc` 只点名"降到最小档仍放不下"，
     不再用 error 阻止渲染（写满必须比写短更安全）;
  3. 渲染产物非空、可被 python-pptx 重新打开；
  4. 留白/并页与「正文精简过头」能被 qc 点名（密度不能靠抽干正文换来）；
  5. **任何内容都不许越出容器**：panel_2col 的迷你表、表格页的表体与 note、
     figure 的 EMF 原样嵌入（用户复现的 P3/P4 两处溢出就是这条的反例）。

**不需要模型**：只走 renderer 的纯 L0 路径。
"""
from __future__ import annotations

import json
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "renderer"))

import layouts                             # noqa: E402
import qc                                  # noqa: E402
import render                              # noqa: E402
import theme                               # noqa: E402
from ppt_dsl import load_deck, validate_dict  # noqa: E402
from pydantic import ValidationError       # noqa: E402

FIXTURE = REPO / "tests" / "fixtures" / "deck.all_types.json"

_failures: list[str] = []


def _check(name: str, cond: bool, detail: str = "") -> None:
    if not cond:
        _failures.append(f"{name}: {detail}")
        print(f"  [FAIL] {name}  {detail}")
    else:
        print(f"  [ ok ] {name}")


def _validate(data: dict):
    """schema 校验：通过返回 (deck, [])，失败返回 (None, 错误字符串)。"""
    try:
        return validate_dict(data), []
    except (ValidationError, ValueError) as e:
        return None, [str(e)[:400]]


def _check_deck(data: dict):
    """schema + 几何质检：返回 (report, exit_code)，语义与旧适配器一致。"""
    deck, errs = _validate(data)
    if deck is None:
        return {"errors": [{"slide": "-", "type": "schema", "field": "-",
                            "message": errs[0]}], "warnings": []}, 1
    report = qc.check(deck)
    return report, (1 if report["errors"] else 0)


def main() -> int:
    deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    slides = deck["slides"]
    types = [s["slide_type"] for s in slides]

    print("== 1. 14 种 slide_type 全覆盖 ==")
    expect = {
        "cover", "toc", "section", "statement", "compare_3col",
        "numbered_cards", "layered_stack", "flow", "hierarchy",
        "table", "chart", "panel_2col", "figure", "end",
    }
    _check("14 种类型各占一页", set(types) == expect, f"实际 {sorted(set(types))}")

    print("== 2. schema 校验 ==")
    model, errs = _validate(deck)
    _check("schema 通过", model is not None and not errs, "; ".join(errs[:1]))

    print("== 3. 几何质检（.pi/renderer/qc.py）==")
    report, code = _check_deck(deck)
    errs = report.get("errors")
    _check("qc 0 error", code == 0 and isinstance(errs, list) and not errs,
           json.dumps(errs, ensure_ascii=False)[:200])

    print("== 4. 渲染 14 页 pptx（含一页 EMF 原样嵌入）==")
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "deck.pptx"
        render.build_pptx(load_deck(str(FIXTURE)), str(out))
        _check("render ok 且产物非空", out.is_file() and out.stat().st_size > 10_000,
               f"{out}")
        # user 要求：Word 里的 EMF/WMF 必须**原封不动**进 PPT（不转格式、不重绘）
        src_emf = (REPO / "tests" / "fixtures" / "sample-diagram.emf").read_bytes()
        with zipfile.ZipFile(out) as z:
            parts = [n for n in z.namelist() if n.endswith(".emf")]
            _check("EMF 进了 ppt/media/ 且与源文件逐字节相同",
                   bool(parts) and z.read(parts[0]) == src_emf, str(parts))
            _check("EMF 的 content-type 已声明（PowerPoint 靠它识别矢量图）",
                   "image/x-emf" in z.read("[Content_Types].xml").decode("utf-8", "replace"))

    print("== 5. 放不下时渲染器降档；qc 不再用 error 拦渲染 ==")

    def _layer_deck(desc: str) -> dict:
        """造一页 layered_stack：第 1 层用 desc，其余层写正常内容。

        层级堆栈的条形高度是均分的，最小字号档容量（约 90 字）比 schema 上限
        （110 字）还紧 —— 正好能命中"降到最小档仍放不下"这条边界。
        """
        layers = [{"label": "待测层", "desc": desc, "tags": ["标签"], "icon": "gear"}]
        for k in (2, 3, 4, 5):
            layers.append({"label": f"对照层{k}",
                           "desc": "正常长度的层描述：命中率从 62% 提到 88%，口径为满 3 个月。",
                           "tags": ["标签"], "icon": "gear"})
        return {"slide_type": "layered_stack",
                "title": "层级描述的超容量边界验证示例",
                "summary": "这一页验证：层描述超过最小字号档容量时只出 warning，"
                           "不再阻止渲染；**渲染器会先把字号降到位**，写满比写短安全。",
                "layers": layers,
                # 带 callout → 辅区落在右栏、条形收窄，描述容量变得更紧（约 90 字）
                "callout": [{"label": "对照一", "value": "62%→88%",
                             "text": "用于让辅区落位到右栏"},
                            {"label": "对照二", "value": "4 项",
                             "text": "用于对照的辅助信息"}]}

    # ① 宽裕：降档就能塞下 —— 不该有任何 error，也不该报超容量
    deck2 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in deck2["slides"]:
        if s.get("slide_type") == "numbered_cards":
            s["cards"][0]["desc"] = "这是一段刻意写长的卡片描述，用来验证渲染器会自己降档：" * 4
            break
    report2, code2 = _check_deck(deck2)
    _check("长描述不再让渲染失败（error 为空）",
           code2 == 0 and not report2.get("errors"),
           json.dumps(report2.get("errors"), ensure_ascii=False)[:160])
    _check("渲染器降档即可塞下时不报超容量",
           not any("仍放不下" in w["message"] for w in report2["warnings"]),
           json.dumps([w for w in report2["warnings"] if "仍放不下" in w["message"]],
                      ensure_ascii=False)[:160])

    # ② 机制层：容量核对本身的两条分支（超容量点名 / 未超容量只记降档）
    model4, errs4 = _validate({"title": "密度规则验证",
                               "slides": [_layer_deck("描述文字" * 27)]})
    _check("超容量边界用例本身能过 schema", model4 is not None, "; ".join(errs4[:1]))
    _slide4 = model4.slides[0]

    msgs4, down4 = [], []
    cap4 = qc._cap_note(1, _slide4, "layers[0].desc", "第 1 层描述",
                        300, 2.0, 0.30, msgs4, down4)
    _check("降到最小档仍放不下时点名（并给出容量）",
           bool(msgs4) and "仍放不下" in msgs4[0]["message"] and str(cap4) in msgs4[0]["message"],
           str(msgs4)[:220])

    msgs5, down5 = [], []
    qc._cap_note(1, _slide4, "layers[0].desc", "第 1 层描述",
                 13, 2.0, 0.30, msgs5, down5)
    _check("未超容量但落到更低档位时只记录、不点名",
           not msgs5 and len(down5) == 1, f"{msgs5} / {down5}")

    _check("stats 里能看到降档页",
           "downshifted_pages" in report2["stats"], str(report2["stats"])[:120])
    _check("正常 fixture 不被误报超容量",
           not any("仍放不下" in w["message"]
                   for w in _check_deck(json.loads(FIXTURE.read_text(encoding="utf-8")))[0]["warnings"]),
           "")

    print("== 6. 密度规则：留白/并页与「精简过头」都要被点名 ==")
    deck3 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in deck3["slides"]:
        if s.get("slide_type") == "table":
            s["rows"] = s["rows"][:2]      # 只剩 2 行 → 主表下方大片留白
            s.pop("side", None)
            break
    r3 = qc.check(validate_dict(deck3))
    msgs3 = [w["message"] for w in r3["warnings"] if w["type"] == "table"]
    _check("小表留白被点名并提示并页/panel_2col",
           any("还空" in m and "panel_2col" in m for m in msgs3), str(msgs3)[:160])

    deck4 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in deck4["slides"]:
        if s.get("slide_type") == "numbered_cards":
            for c in s["cards"]:
                c["desc"] = "这是一段明显偏短的卡片描述文字凑够二十个字"
            break
    r4 = qc.check(validate_dict(deck4))
    _check("精简过头的卡片描述被点名",
           any("不足" in w["message"] and "描述" in w["message"] for w in r4["warnings"]),
           str([w["message"] for w in r4["warnings"] if w["type"] == "numbered_cards"])[:160])
    _check("密度规则只出 warning，不拦渲染", not r3["errors"] and not r4["errors"])

    print("== 7. 图标：名字必须已登记，写错也不许缺图 ==")

    def _all_icons(obj, out: list):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == "icon" and v:
                    out.append(v)
                _all_icons(v, out)
        elif isinstance(obj, list):
            for x in obj:
                _all_icons(x, out)

    names: list[str] = []
    for p in sorted((REPO / ".pi" / "skills" / "ppt-common" / "examples").glob("*.json")):
        _all_icons(json.loads(p.read_text(encoding="utf-8")), names)
    _all_icons(json.loads(FIXTURE.read_text(encoding="utf-8")), names)
    bad = sorted({n for n in names if not layouts.icon_resolvable(n)})
    _check("范例与 fixture 里的 icon 全部能解析", not bad, f"未登记：{bad}")

    # 「版本」这种没登记的中文名曾经静默丢图标 —— 现在必须兜底成画得出来的名字
    _check("写错的 icon 名会兜底成可绘制的图标",
           layouts.icon_resolvable(layouts._resolve_icon("完全不存在-图标名", "项目规范与编码标准")),
           "兜底后仍不可解析")

    deck5 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in deck5["slides"]:
        if s.get("slide_type") == "numbered_cards":
            s["cards"][0]["icon"] = "完全不存在-图标名"
            break
    r5 = qc.check(validate_dict(deck5))
    _check("未登记的 icon 名被 qc 点名",
           any("找不到" in w["message"] for w in r5["warnings"]),
           str([w["message"] for w in r5["warnings"] if "icon" in w["field"]])[:120])

    print("== 8. 辅区落位：能放右边就放右边（左右优先）==")
    B = layouts.BODY_BOTTOM
    rail = layouts.card_frame("结" * 90, True, layouts.FLOW_CYCLE_MIN_W, B, top_pad=0.0)
    _check("环形流程 + callout → 右栏", rail["mode"] == "rail", str(rail["mode"]))
    _check("右栏时主内容高度不被压缩",
           rail["bottom"] == B and rail["w"] < 12.23,
           f"w={rail['w']:.2f} bottom={rail['bottom']:.2f}")
    stack = layouts.card_frame("结" * 90, True, layouts.LAYER_MIN_W, B, top_pad=0.18)
    _check("层级堆栈 + callout → 右栏", stack["mode"] == "rail", str(stack["mode"]))
    comp = layouts.card_frame("结" * 90, True, layouts.COMPARE_MIN_W, B, top_pad=0.18)
    _check("三栏对比 + callout → 底部横条（宽度敏感）", comp["mode"] == "bottom",
           str(comp["mode"]))
    none = layouts.card_frame("结" * 90, False, layouts.COMPARE_MIN_W, B)
    _check("没有辅区时主内容满宽满高",
           none["mode"] == "none" and none["w"] == 12.23 and none["bottom"] == B,
           str(none))
    # 页脚条压扁下移到 FOOTER_Y 之后，内容带净高（2 行结论段）必须比旧版更高
    band = B - layouts.header_y("结" * 90)
    _check(f"内容带净高 ≥ 4.85in（实测 {band:.2f}）", band >= 4.85, f"{band:.2f}")
    _check("页脚条在内容带之下、且右端避开右下角水印",
           layouts.FOOTER_Y >= B + 0.10 and layouts.FOOTER_RIGHT <= 10.95,
           f"FOOTER_Y={layouts.FOOTER_Y} FOOTER_RIGHT={layouts.FOOTER_RIGHT}")

    print("== 9. 卡片标题即重点：名词短语被点名、判断句放行 ==")
    noun_deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in noun_deck["slides"]:
        if s.get("slide_type") == "numbered_cards":
            for c, nm in zip(s["cards"], ["任务集管理", "指标标准化",
                                          "回归基线", "归因下钻"]):
                c["name"] = nm
            break
    rn = qc.check(validate_dict(noun_deck))
    _check("名词式卡片标题被点名",
           any("读不出结论" in w["message"] for w in rn["warnings"]),
           str([w["message"][:80] for w in rn["warnings"] if "name" in w["field"]]))

    claim_deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in claim_deck["slides"]:
        if s.get("slide_type") == "numbered_cards":
            for c, nm in zip(s["cards"], ["统一任务集管理", "指标口径必须统一",
                                          "劣化项 100% 告警", "失败样本下钻到能力维度"]):
                c["name"] = nm
            break
    rc = qc.check(validate_dict(claim_deck))
    _check("判断句/带数字的卡片标题放行",
           not any("读不出结论" in w["message"] for w in rc["warnings"]),
           str([w["message"][:80] for w in rc["warnings"] if "name" in w["field"]]))

    print("== 10. 纯文字卡骨架：换行分卡 + 关键句/关键短语（同字号 + 加粗 + 强调色）==")
    body = ("先造尺子，再谈迭代速度。没有统一口径，迭代只能在经验里打转。\n"
            "统一评测把三件事固化进流水线。**失败可归因：** 按能力维度下钻。")
    _check("换行 = 卡片边界（2 张卡）",
           layouts.split_body_groups(body) == body.split("\n"),
           str(layouts.split_body_groups(body)))
    _check("每卡首句被切成关键句",
           layouts.split_claim(body.split("\n")[0])[0] == "先造尺子，再谈迭代速度。",
           layouts.split_claim(body.split("\n")[0])[0])
    sec = layouts.split_sections("长文一段。**判定口径：** 重叠约 10 个。**反例：** 只会放大。")
    _check("`**小节名：**` 被切成 2 节（前置导语保留成无标题的一节）",
           [h for h, _ in sec] == ["", "判定口径", "反例"], str(sec))
    sec2 = layouts.split_sections("**口径：**冒号写在加粗里的写法也认。")
    _check("`**小节名：**` 另一种写法也认", [h for h, _ in sec2] == ["口径"], str(sec2))
    sec3 = layouts.split_sections("把**回归通过率**提升到 88%：这是普通加粗。")
    _check("普通加粗不会被误判成小节", [h for h, _ in sec3] == [""], str(sec3))
    paras = layouts.card_paras(body.split("\n")[0], 11.5, claim=True)
    _check("卡内文本是一段到底（关键句不另起一行、不放大）",
           len(paras) == 1 and paras[0][1]["size"] == 11.5
           and paras[0][1]["strong_color"] == "primary",
           str(paras))
    _check("关键句被自动包进 ** ** 走强调色",
           paras[0][0].startswith("**先造尺子，再谈迭代速度。**"),
           paras[0][0][:40])
    _check("关键句不占额外高度（同字号 = 同高度，防溢出卡片背景）",
           layouts.card_text_h("甲。" * 40, 7.8, 11.5, claim=True)
           <= layouts.card_text_h("甲。" * 40, 7.8, 11.5, claim=False))
    sec_paras = layouts.card_paras("正文。**判定口径：** 重叠约 10 个。", 11.5)
    _check("关键短语与正文同一段、同字号",
           len(sec_paras) == 1 and sec_paras[0][1]["size"] == 11.5
           and "**判定口径：**" in sec_paras[0][0],
           str(sec_paras))

    long_deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in long_deck["slides"]:
        if s.get("slide_type") == "statement":
            # 一整句 195 字、没有任何 `**小节**：`：关键句过长 + 长卡无小节，两条都该点名
            s["body"] = "关键句与机制细节全部堆在同一句里，" * 15
            break
    rl = qc.check(validate_dict(long_deck))
    msgs = " ".join(w["message"] for w in rl["warnings"])
    _check("关键句过长被点名", "关键句" in msgs and "建议 ≤48" in msgs, msgs[:140])
    _check("长卡缺关键短语被点名", "处关键短语" in msgs, msgs[:140])
    sec_deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in sec_deck["slides"]:
        if s.get("slide_type") == "statement":
            s["body"] = ("先造尺子。**收口到一处：** 任务集与指标统一维护。"
                         "**反例：** 没有统一口径，迭代只能在经验里打转。" * 2)
            break
    rs = qc.check(validate_dict(sec_deck))
    _check("关键短语不被当成「加粗滥用」",
           not any("加粗" in w["message"] for w in rs["warnings"]),
           str([w["message"][:70] for w in rs["warnings"] if "加粗" in w["message"]]))

    print("== 11. 块头标题必须给右上角图标让位（P9 第 3/5 张卡压图标）==")
    tw_icon = layouts.block_title_w(3.9, layouts.NO_GUTTER, True)
    tw_free = layouts.block_title_w(3.9, layouts.NO_GUTTER, False)
    _check("有图标时标题可用宽度更窄", tw_icon < tw_free - 0.5,
           f"icon={tw_icon:.2f} free={tw_free:.2f}")
    _check("一行放不下的标题会自动缩字号（最低 11.5pt）",
           layouts.fit_one_line("标题" * 12, tw_icon) < layouts.FONT_SIZES["card_title"]
           and layouts.fit_one_line("标题" * 40, tw_icon) >= 11.5,
           f"{layouts.fit_one_line('标题' * 12, tw_icon)}")
    long_title = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in long_title["slides"]:
        if s.get("slide_type") == "numbered_cards":
            s["cards"] = (s["cards"] + s["cards"][:1])[:5]     # 5 张 → 3 列（窄卡）
            s["cards"][0]["name"] = "隔离执行并探索互斥方案"
            break
    rl2 = qc.check(validate_dict(long_title))
    _check("三列窄卡里的长标题被点名（会压图标）",
           any("压到图标" in w["message"] for w in rl2["warnings"]),
           str([w["message"][:80] for w in rl2["warnings"] if "name" in w["field"]]))

    print("== 12. flow 8 步一页装下（两行折返）——不许因步数多而删步骤 ==")
    fl_deck = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in fl_deck["slides"]:
        if s.get("slide_type") == "flow":
            s["steps"] = (list(s["steps"]) * 2)[:8]
            break
    d8 = validate_dict(fl_deck)
    flow_slide = [s for s in d8.slides if s.slide_type == "flow"][0]
    g8 = layouts.flow_linear_geom(flow_slide)
    _check("8 步折成两行、每行 4 步", g8["rows"] == 2 and g8["per"] == 4,
           f"rows={g8['rows']} per={g8['per']}")
    _check(f"折行后每步仍 ≥1.2in 宽（实测 {g8['bw']:.2f}）", g8["bw"] >= 1.2,
           f"{g8['bw']:.2f}")
    _check("折行后有折返通道（行间距 >0）", g8["row_gap"] > 0.2, str(g8["row_gap"]))
    six = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in six["slides"]:
        if s.get("slide_type") == "flow":
            s["steps"] = list(s["steps"])[:6]
            break
    _check("6 步仍是一行（不折）",
           layouts.flow_linear_geom(
               [s for s in validate_dict(six).slides
                if s.slide_type == "flow"][0])["rows"] == 1)
    big = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in big["slides"]:
        if s.get("slide_type") == "flow":
            s["steps"] = [dict(s["steps"][0], label=f"步骤{i}") for i in range(12)]
            break
    try:
        validate_dict(big)
        _check("schema 放行 12 步", True)
    except ValidationError as e:
        _check("schema 放行 12 步", False, str(e)[:120])
    over = json.loads(json.dumps(big))
    for s in over["slides"]:
        if s.get("slide_type") == "flow":
            s["steps"] = s["steps"] + [dict(s["steps"][0], label="第13步")]
            break
    try:
        validate_dict(over)
        _check("schema 拒绝 13 步（报错里写明拆页而非删）", False, "没拦住")
    except ValidationError as e:
        _check("schema 拒绝 13 步（报错里写明拆页而非删）", True, str(e)[:60])
    with tempfile.TemporaryDirectory() as td:
        out8 = Path(td) / "flow8.pptx"
        p8 = Path(td) / "flow8.json"
        p8.write_text(json.dumps(fl_deck, ensure_ascii=False), encoding="utf-8")
        render.build_pptx(load_deck(str(p8)), str(out8))
        _check("8 步流程页渲染成功", out8.is_file() and out8.stat().st_size > 10_000)

    print("== 13. 输入覆盖度：丢了内容能被抓到 ==")
    deck6 = validate_dict(json.loads(FIXTURE.read_text(encoding="utf-8")))
    src = ("| 章节 | 模块 | 核心结论 | 论据 | 缺口 |\n|---|---|---|---|---|\n"
           "| 1 | 存活率显著高于基线 | 试点跑赢基线 | 综合存活率 94.2%，样本 1,204 家 | 无 |\n"
           "| 1 | 门店数字化改造 | 单店损益实时可见 | 处置时效缩短 60% | 无 |\n")
    cov = qc.check_source_coverage(deck6, src)
    _check("已覆盖的模块不误报（存活率显著高于基线）",
           "存活率显著高于基线" not in cov["modules_missing"],
           str(cov["modules_missing"]))
    _check("没进 deck 的模块被点名（门店数字化改造）",
           "门店数字化改造" in cov["modules_missing"], str(cov["modules_missing"]))
    _check("千分位数字不算丢（1,204 已在表里）",
           "1204" not in cov["numbers_missing"], str(cov["numbers_missing"]))
    _check("覆盖度返回统计（行数/数字数）",
           cov["rows"] == 2 and cov["numbers"] >= 2, str(cov))
    miss_src = ("| 章节 | 模块 | 核心结论 | 论据 | 缺口 |\n|---|---|---|---|---|\n"
                "| 2 | 未落地的能力 | 需求预测准确率 88.7% | 缺样本口径 | 无 |\n")
    cov2 = qc.check_source_coverage(deck6, miss_src)
    _check("缺失的数字口径被点名（88.7%）",
           "88.7%" in cov2["numbers_missing"], str(cov2["numbers_missing"]))
    sem = json.loads(FIXTURE.read_text(encoding="utf-8"))
    sem["brief"] = "seminar"
    r_sem = qc.check(validate_dict(sem))
    _check("研讨型不带 --source 会被点名（requireSource）",
           any("输入覆盖度" in w["message"] for w in r_sem["warnings"]),
           str([w["message"][:70] for w in r_sem["warnings"] if w["field"] == "source"]))
    r_src_bad = qc.check(validate_dict(sem), source_path="不存在/analysis.md")
    _check("--source 路径写错 → error",
           any("读不到" in e["message"] for e in r_src_bad["errors"]),
           str(r_src_bad["errors"])[:100])

    print("== 14. 上限按真实几何容量放宽：文档有几项就摆几项 ==")

    def _mutate(kind: str, fn) -> dict:
        d = json.loads(FIXTURE.read_text(encoding="utf-8"))
        for s in d["slides"]:
            if s.get("slide_type") == kind:
                fn(s)
                break
        return d

    lay = _mutate("layered_stack", lambda s: s.__setitem__(
        "layers", (s["layers"] * 3)[:12]))
    m12 = [x for x in validate_dict(lay).slides if x.slide_type == "layered_stack"][0]
    g12 = layouts.layered_geom(m12)
    _check("堆栈 12 层过 schema 且条形仍 ≥0.3in 高",
           len(m12.layers) == 12 and g12["bar_h"] >= 0.30, f"{g12['bar_h']:.2f}")
    _check("12 层时间距自动压紧（0.14 → 0.07）", g12["gap"] <= 0.08, str(g12["gap"]))

    num = _mutate("numbered_cards", lambda s: s.__setitem__(
        "cards", [dict(c, no=f"{i + 1:02d}") for i, c in enumerate(
            (s["cards"] * 2)[:8])]))
    m8c = [x for x in validate_dict(num).slides if x.slide_type == "numbered_cards"][0]
    g8c = layouts.numbered_geom(m8c)
    _check("编号卡 8 张排 4 列 × 2 行", g8c["cols"] == 4 and g8c["rows"] == 2,
           f"cols={g8c['cols']} rows={g8c['rows']}")
    _check("4 列窄卡不画块头图标（宽度还给标题与描述）",
           g8c["has_icon"] is False and g8c["desc_w"] >= 2.0, str(g8c["has_icon"]))

    cmp4 = _mutate("compare_3col", lambda s: s.__setitem__(
        "columns", s["columns"] + [dict(s["columns"][0], label="混合甲")]))
    m4col = [x for x in validate_dict(cmp4).slides if x.slide_type == "compare_3col"][0]
    dw4 = layouts.compare_desc_box(m4col.summary, 2, False, n_cols=4)[0]
    dw3 = layouts.compare_desc_box(m4col.summary, 2, False, n_cols=3)[0]
    _check("对比页 4 栏过 schema，且描述宽随栏数收窄", dw4 < dw3 - 0.5,
           f"4栏={dw4:.2f} 3栏={dw3:.2f}")

    tab = _mutate("table", lambda s: s.__setitem__(
        "rows", [list(s["rows"][0]) for _ in range(12)]))
    m12t = [x for x in validate_dict(tab).slides if x.slide_type == "table"][0]
    _sl = layouts.table_slack(12, m12t.summary, bool(m12t.note))[0]
    _check("表格 12 行过 schema，且行高自动压紧后不再溢出（slack≈0）",
           len(m12t.rows) == 12 and _sl > -0.1, f"rows=12 slack={_sl:.2f}")
    hie = _mutate("hierarchy", lambda s: s.__setitem__(
        "nodes", (s["nodes"] * 4)[:8]))
    m8n = [x for x in validate_dict(hie).slides if x.slide_type == "hierarchy"][0]
    _check("层级图 8 个子模块过 schema", len(m8n.nodes) == 8, str(len(m8n.nodes)))

    st7 = _mutate("statement", lambda s: s.__setitem__("body", "甲" * 700))
    _check("statement.body 700 字过 schema",
           len([x for x in validate_dict(st7).slides
                if x.slide_type == "statement"][0].body) == 700)
    st_over = _mutate("statement", lambda s: s.__setitem__("body", "甲" * 701))
    try:
        validate_dict(st_over)
        _check("statement.body 701 字被 schema 拒绝", False, "没拦住")
    except ValidationError:
        _check("statement.body 701 字被 schema 拒绝", True)
    big_lay = _mutate("layered_stack", lambda s: s.__setitem__(
        "layers", (s["layers"] * 5)[:13]))
    try:
        validate_dict(big_lay)
        _check("堆栈 13 层被拒绝（报错写明另起一页）", False, "没拦住")
    except ValidationError:
        _check("堆栈 13 层被拒绝（报错写明另起一页）", True)

    print("== 15. 质检报告：合法 JSON + 同一规则不复述（报告整份进上下文）==")
    rep7 = qc.check(validate_dict(json.loads(FIXTURE.read_text(encoding="utf-8"))))
    txt = render._dump_report(rep7)
    _check("--qc 报告是合法 JSON（模型就是按 JSON 读它）",
           json.loads(txt)["warnings"] == rep7["warnings"], txt[:80])
    _check("紧凑打印比 indent=2 省（结构缩进不再进上下文）",
           len(txt) < len(json.dumps(rep7, ensure_ascii=False, indent=2)),
           f"{len(txt)} vs {len(json.dumps(rep7, ensure_ascii=False, indent=2))}")
    _dup = [
        {"slide": 4, "type": "table", "field": "rows",
         "message": "表格 4 行、还空 1.8 in（约能再放 3 行）：把下一章节的明细并进来，"
                    "补 3 行左右；原文没有就补一行显式占位。不要靠拉高行高硬撑。"},
        {"slide": 9, "type": "table", "field": "rows",
         "message": "表格 5 行、还空 1.6 in（约能再放 3 行）：把下一章节的明细并进来，"
                    "补 3 行左右；原文没有就补一行显式占位。不要靠拉高行高硬撑。"},
    ]
    _saved = qc._fold_rule_tails(_dup)
    _check("同规则第 2 条：只留页内数字 + 指向首条，规则说明不再复述",
           _saved > 0
           and "（规则见第 4 页）" in _dup[1]["message"]
           and "1.6 in" in _dup[1]["message"]          # 数值与单位不许被切开
           and "拉高行高" not in _dup[1]["message"]     # 说明只在首条保留
           and "拉高行高" in _dup[0]["message"],
           f"saved={_saved} / {_dup[1]['message']}")
    _check("折叠不会留下半个括号的残句",
           _dup[1]["message"].count("（") == _dup[1]["message"].count("）"),
           _dup[1]["message"])
    _check("不同规则不误折叠（后缀不同的两条都保留）",
           qc._fold_rule_tails([
               {"slide": 1, "type": "table", "field": "rows", "message": "甲" * 40},
               {"slide": 2, "type": "table", "field": "rows", "message": "乙" * 40},
           ]) == 0)

    print("== 16. 内容不许越出容器（P3 迷你表 / P4 表体 + note 两处溢出回归）==")

    def _one_slide(td: str, slide_json: dict, name: str) -> Path:
        p = Path(td) / f"{name}.pptx"
        render.build_pptx(validate_dict({"title": "容器溢出回归", "slides": [slide_json]}),
                          str(p))
        return p

    def _worst_content_bottom(path: Path) -> float:
        """一页里所有**内容形状**的最低边：排除整页底图与页脚/页码条。"""
        from pptx import Presentation as _Prs
        from pptx.util import Emu as _Emu
        worst = 0.0
        for sl in _Prs(str(path)).slides:
            for sh in sl.shapes:
                top, h = _Emu(sh.top).inches, _Emu(sh.height).inches
                if h >= layouts.SLIDE_H - 0.01:          # 整页底图
                    continue
                if top >= layouts.FOOTER_Y - 0.01:       # 页脚条 / 页码
                    continue
                worst = max(worst, top + h)
        return worst

    # ---- P3：双模块页，右模块「要点 + 迷你表」——要点多到放不下时，表必须仍留在框内 ----
    p3 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in p3["slides"]:
        if s.get("slide_type") == "panel_2col":
            s["panels"][1]["points"] = [
                "短期记忆：上下文窗口内的临时状态，退出会话即失效",
                "长期记忆：持久化存储，跨会话可恢复，是 Agent 区别于 Chatbot 的关键",
                "样例一：这一条是为了把下方的迷你表顶出模块框而加的",
                "样例二：再加一条，模块里的内容就彻底超框了",
                "样例三：要点超过 6 条 / 明细超过 8 行就该另起一页",
            ]
            s["panels"][1]["headers"] = ["实现方式", "适用场景", "特点"]
            s["panels"][1]["rows"] = [
                ["SQLite/Redis", "大规模会话存储", "数据库持久化，高并发读写"],
                ["命名空间 JSON", "复杂状态管理", "结构化存储，适合嵌套状态"],
                ["MEMORY.md", "用户偏好与学习记录", "自动记忆文件，持续积累"],
                ["CLAUDE.md", "项目规范与编码标准", "Markdown 持久文件，团队共享"],
                ["样例样例", "样例样例详情", "样例样例详情"],
                ["样例样例", "样例样例详情", "样例样例详情"],
            ]
            break
    m_p3 = validate_dict(p3)
    s_p3 = [x for x in m_p3.slides if x.slide_type == "panel_2col"][0]
    _notes = s_p3.callout or []
    _top, _bottom = layouts.panel_area(s_p3.summary, bool(_notes),
                                       bool(s_p3.footer_note), layouts.callout_height(_notes))
    _ph = _bottom - _top
    _pw = (12.23 - layouts.PANEL_GAP) / 2
    _g = layouts.panel_content_geom(s_p3.panels[1].summary, _top, _ph, _pw, s_p3.panels[1])
    _check("迷你表底部不越出模块内容区（底部锚定，数学保证）",
           _g["tbl_y"] + _g["tbl_h"] <= _g["cy_bottom"] + 1e-6,
           f"表底 {_g['tbl_y'] + _g['tbl_h']:.3f} > 内容区底 {_g['cy_bottom']:.3f}")
    _pts_used = sum(_g["pt_hs"]) + _g["pt_gap"] * max(0, len(_g["pt_hs"]) - 1)
    _check("要点占位不超过留给它的区间（不会压到迷你表）",
           _g["cy"] + _pts_used <= _g["pt_bottom"] + 1e-6,
           f"要点底 {_g['cy'] + _pts_used:.3f} > {_g['pt_bottom']:.3f}")
    with tempfile.TemporaryDirectory() as td:
        wb = _worst_content_bottom(_one_slide(td, s_p3.model_dump(), "p3"))
    _check(f"渲染后最低内容边 {wb:.2f} ≤ BODY_BOTTOM {layouts.BODY_BOTTOM}",
           wb <= layouts.BODY_BOTTOM + 0.02, f"{wb:.3f}")
    _r3 = qc.check(m_p3)
    _check("要点过多被 qc 点名（内容超框）",
           any(w["type"] == "panel_2col" and "超出模块框" in w["message"]
               for w in _r3["warnings"]),
           str([w["message"][:60] for w in _r3["warnings"] if w["type"] == "panel_2col"]))

    # ---- P4：表格 11 行 + note —— 表体不许出界，note 不许压表 ----
    p4 = json.loads(FIXTURE.read_text(encoding="utf-8"))
    for s in p4["slides"]:
        if s.get("slide_type") == "table":
            s["rows"] = [list(s["rows"][0]) for _ in range(11)]
            s["rows"][0] = ["观测屏蔽", "隐藏冗余输出", "20-30%"]
            s["headers"] = ["策略", "作用", "Token 节省"]
            break
    m_p4 = validate_dict(p4)
    s_p4 = [x for x in m_p4.slides if x.slide_type == "table"][0]
    _side4 = s_p4.side or []
    _g4 = layouts.table_geom(len(s_p4.rows), s_p4.summary, bool(s_p4.note),
                             s_p4.headers, s_p4.rows, bool(_side4))
    _check("11 行表格的行高按可用高度压紧，总高不出内容区",
           _g4["total_h"] <= _g4["avail"] + 1e-6,
           f"total_h={_g4['total_h']:.3f} avail={_g4['avail']:.3f}")
    _check("底部 note 落在表格下方（不再被辅区底边推到表格上）",
           _g4["note_y"] >= _g4["top"] + _g4["total_h"] - 1e-6
           and _g4["note_y"] + _g4["note_h"] <= layouts.BODY_BOTTOM + 1e-6,
           f"note_y={_g4['note_y']:.3f} 表底={_g4['top'] + _g4['total_h']:.3f}")
    _check("行高容得下单元格文字（PowerPoint 打开不会被撑行）",
           _g4["cell_overflow"] <= 0.02,
           f"最小档仍高出 {_g4['cell_overflow']:.3f} in")
    with tempfile.TemporaryDirectory() as td:
        wb4 = _worst_content_bottom(_one_slide(td, s_p4.model_dump(), "p4"))
    _check(f"渲染后最低内容边 {wb4:.2f} ≤ BODY_BOTTOM {layouts.BODY_BOTTOM}",
           wb4 <= layouts.BODY_BOTTOM + 0.02, f"{wb4:.3f}")

    print("== 17. 字体口径：中文微软雅黑 + 西文 Times New Roman（不再随包分发字体）==")
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    _check("字体口径常量就是这两个（全 deck 唯一来源）",
           layouts.FONT_LATIN == "Times New Roman" and layouts.FONT_MAIN == "微软雅黑"
           and not hasattr(theme, "FONT_FALLBACK"),
           f"latin={layouts.FONT_LATIN} ea={layouts.FONT_MAIN}")

    def _font_of(el, tag):
        node = el.find(f"{{{A_NS}}}{tag}")
        return None if node is None else node.get("typeface")

    def _bad_font_runs(xml_bytes) -> tuple[int, list[str]]:
        """每个 a:rPr / a:defRPr 都必须成对带上 latin + ea。返回 (检查过的段数, 问题)。"""
        import xml.etree.ElementTree as ET
        root = ET.fromstring(xml_bytes)
        n, bad = 0, []
        for tag in ("rPr", "defRPr"):
            for el in root.iter(f"{{{A_NS}}}{tag}"):
                n += 1
                if _font_of(el, "latin") != "Times New Roman" or _font_of(el, "ea") != "微软雅黑":
                    bad.append(f"<{tag} latin={_font_of(el, 'latin')} ea={_font_of(el, 'ea')}>")
        return n, bad

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "fonts.pptx"
        render.build_pptx(load_deck(str(FIXTURE)), str(out))
        with zipfile.ZipFile(out) as z:
            runs, bad = 0, []
            for name in z.namelist():
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    n, b = _bad_font_runs(z.read(name))
                    runs += n
                    bad += [f"{name}: {x}" for x in b]
            _check(f"幻灯片里 {runs} 段文字都带 latin+ea 字体对", runs > 50 and not bad,
                   "; ".join(bad[:3]))
            charts = [n for n in z.namelist()
                      if n.startswith("ppt/charts/chart") and n.endswith(".xml")]
            cruns, cbad = 0, []
            for name in charts:
                n, b = _bad_font_runs(z.read(name))
                cruns += n
                cbad += [f"{name}: {x}" for x in b]
            _check(f"原生图表 {len(charts)} 个部件里 {cruns} 段文字也补齐了字体对",
                   bool(charts) and cruns > 0 and not cbad, "; ".join(cbad[:3]))

            import xml.etree.ElementTree as ET
            theme_root = ET.fromstring(z.read("ppt/theme/theme1.xml"))
            scheme = theme_root.find(f".//{{{A_NS}}}fontScheme")
            faces = {}
            for tag in ("majorFont", "minorFont"):
                font = scheme.find(f"{{{A_NS}}}{tag}")
                faces[tag] = (font.find(f"{{{A_NS}}}latin").get("typeface"),
                              font.find(f"{{{A_NS}}}ea").get("typeface"))
            hans = {f.get("script"): f.get("typeface")
                    for f in scheme.iter(f"{{{A_NS}}}font")}
            _check("母版主题默认字体：西文 Times New Roman / 中文微软雅黑",
                   all(v == ("Times New Roman", "微软雅黑") for v in faces.values())
                   and hans.get("Hans") == "微软雅黑", f"{faces} Hans={hans.get('Hans')}")

    print("== 18. 规划链路：outline → deck 骨架 → 渲染（这条链以前是断的）==")
    import outline as outline_mod
    from outline import SlidePlan, SlideType, _shell

    plans = [SlidePlan(title=f"待补测试页标题{i}", slide_type=t,
                       conclusion="这是一句待补的结论。")
             for i, t in enumerate(SlideType)]
    bad = []
    for p in plans:
        try:
            validate_dict({"title": "测试 deck 标题", "slides": [_shell(p)]})
        except (ValidationError, ValueError) as e:
            bad.append(f"{p.slide_type.value}: {str(e)[:100]}")
    _check(f"{len(plans)} 种版式的骨架都能过 schema（plan.py 提示的下一步真的能跑）",
           not bad, "; ".join(bad[:3]))

    skel = {"title": "规划链路骨架", "slides": [_shell(p) for p in plans]}
    skel_deck, skel_errs = _validate(skel)
    _check("整份骨架 deck 可校验", skel_deck is not None, (skel_errs or [""])[0][:120])
    if skel_deck is not None:
        with tempfile.TemporaryDirectory() as td:
            out = str(Path(td) / "shell.pptx")
            render.build_pptx(skel_deck, out, add_speaker_notes=False)
            n = len(skel_deck.slides)
            ok = zipfile.is_zipfile(out) and Path(out).stat().st_size > 20_000
            _check(f"骨架 deck 渲染出 {n} 页且产物可用", ok)

    print()
    print("== 19. 容量上限按**条目数**动态校验（schema 不再说一套、几何做一套）==")
    from ppt_dsl import (EndSlide, FlowSlide, KeyGroupsSlide, MatrixSlide,
                         NumberedCardsSlide, StatementSlide, TableSlide,
                         TimelineSlide)
    _long = ("这是一段用于测试容量校验的正文，字数必须够多才能触发拦截，"
             "内容本身没有意义只看长度。") * 12

    def _rejects(cls, payload) -> bool:
        try:
            cls(**payload)
            return False
        except (ValidationError, ValueError):
            return True

    def _cards(n, size):
        return [{"no": f"0{i+1}", "name": "统一任务集管理", "desc": _long[:size]}
                for i in range(n)]

    _check("4 张卡 × 240 字放行",
           not _rejects(NumberedCardsSlide, {"slide_type": "numbered_cards",
                                             "title": "测试卡片页标题", "cards": _cards(4, 240)}))
    _check("8 张卡 × 240 字被拦（同字段、不同条目数 → 不同容量）",
           _rejects(NumberedCardsSlide, {"slide_type": "numbered_cards",
                                         "title": "测试卡片页标题", "cards": _cards(8, 240)}))
    _check("12 节点 × 120 字被拦（timeline 竖排每节点只有一行）",
           _rejects(TimelineSlide, {"slide_type": "timeline", "title": "测试时间轴标题",
                                    "milestones": [{"label": f"节点{i}", "desc": _long[:120]}
                                                   for i in range(12)]}))
    _check("4 组 × 6 条 × 58 字被拦（key_groups 每格只有 ~37 字）",
           _rejects(KeyGroupsSlide, {"slide_type": "key_groups", "title": "测试分组页标题",
                                     "groups": [{"title": f"组{i}",
                                                 "points": [_long[:58]] * 6}
                                                for i in range(4)]}))
    _check("5 步环形 × 80 字被拦（节点固定 1.72in 宽），12 步线性 49 字放行",
           _rejects(FlowSlide, {"slide_type": "flow", "title": "测试流程页标题",
                                "loop_note": "循环：失败回退",
                                "steps": [{"label": f"步骤{i}", "desc": _long[:80]}
                                          for i in range(5)]})
           and not _rejects(FlowSlide, {"slide_type": "flow", "title": "测试流程页标题",
                                        "steps": [{"label": f"步骤{i}", "desc": _long[:49]}
                                                  for i in range(12)]}))
    _check("matrix 4×4 × 90 字被拦（每格 ~28 字）",
           _rejects(MatrixSlide, {"slide_type": "matrix", "title": "测试矩阵页标题",
                                  "col_labels": ["待补列1", "待补列2", "待补列3", "待补列4"],
                                  "rows": [{"label": f"行{i}",
                                            "cells": [{"title": "格子标题",
                                                       "desc": _long[:90]}
                                                      for _ in range(4)]}
                                           for i in range(4)]}))
    _check("表格行宽不匹配被拦（这个 validator 以前忘了加装饰器，从没跑过）",
           _rejects(TableSlide, {"slide_type": "table", "title": "测试表格标题",
                                 "headers": ["指标", "目标", "现状"],
                                 "rows": [["只有两列", "缺一列"]]}))
    _check("表格 hero_col 指向不存在的列被拦",
           _rejects(TableSlide, {"slide_type": "table", "title": "测试表格标题",
                                 "headers": ["指标", "目标", "现状"],
                                 "rows": [["a", "b", "c"]], "hero_col": "没有这列"}))
    _check("statement 700 字正文（schema 上限）放行 —— 动态校验带容差、不误伤",
           not _rejects(StatementSlide, {"slide_type": "statement",
                                         "title": "测试判断句标题六个字",
                                         "body": _long[:700]}))
    _check("结尾页 2 段正常放行",
           not _rejects(EndSlide, {"slide_type": "end",
                                   "sections": [{"heading": "下一步", "text": _long[:80]}
                                                for _ in range(2)]}))

    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项失败：")
        for f in _failures:
            print(f"  - {f}")
        return 1
    print("[PASS] 全部通过 —— renderer 离线链路无回归")
    return 0


if __name__ == "__main__":
    sys.exit(main())
