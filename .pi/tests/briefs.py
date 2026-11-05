# -*- coding: utf-8 -*-
"""写作方向（研讨型 / 汇报型）回归测试。

两个方向共用一套版式与配色，差别只在「怎么取舍原文」—— 所以规则必须能机检，
否则就退化成一句口号：Agent 说自己是汇报型，产出的却是逐条堆料。
本测试断言 `.pi/renderer/briefs/*.yaml` 的 check 段真的被 qc 执行了。

跑法：`python .pi/tests/briefs.py`（纯离线，不需要 Office）
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "renderer"))

import briefs          # noqa: E402
import qc              # noqa: E402
import theme           # noqa: E402
from ppt_dsl import validate_dict  # noqa: E402

_failures: list[str] = []

# 给 chart 页轮换图型，用来验证「各种图表」那条规则
CHART_TYPES = ["column", "bar", "line", "pie", "doughnut", "radar"]


def _check(name: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  [ ok ] {name}")
    else:
        _failures.append(f"{name}: {detail}")
        print(f"  [FAIL] {name}  {detail}")


def _summary(tag: str) -> str:
    """**合规**的结论段样例：40~80 字、2~3 句、无坐标词与填充语、不复述标题。

    它同时是"结论段硬指标"用例的基准线 —— 改这里要连 `== 10 ==` 一起看。
    区间 2026-09 由 60~140 收到 40~80（页头是最贵的版面，机制细节下放卡片）。
    """
    return (f"{tag}的落地效果可以直接用数据核对：本季度回归通过率从 62% 提到 88%。"
            f"剩下的差异集中在长尾任务上，要等补齐口径后才能判断。")


def _slide(kind: str, i: int, tag: str = "主题") -> dict:
    if kind == "statement":
        return {"slide_type": "statement",
                "title": f"第 {i} 页的判断句标题，说明一个明确的主张而非话题",
                "summary": _summary(tag),
                "body": "这是一段足够长的正文，用来满足 statement 至少 80 字的下限要求。" * 4,
                "points": [{"label": "论据", "text": f"第 {k} 条支撑论据的一句话说明"}
                           for k in range(1, 5)]}
    if kind == "chart":
        return {"slide_type": "chart",
                "title": f"第 {i} 页的判断句标题，用一张图佐证一个主张",
                "summary": _summary(tag),
                "chart": {"type": CHART_TYPES[i % len(CHART_TYPES)],
                          "categories": ["甲", "乙", "丙"],
                          "series": [{"name": "占比", "values": [1, 2, 3]}], "unit": "%"},
                "insights": [{"text": f"第 {k} 条洞察，说明图上的一个关键事实"}
                             for k in range(1, 5)]}
    if kind == "table":
        return {"slide_type": "table", "title": "对照表：三条路线的关键差异",
                "summary": _summary(tag),
                "headers": ["路线", "代价"], "rows": [["甲", "低"], ["乙", "高"]]}
    raise ValueError(kind)


def _deck(brief: str | None, n_content: int, kind: str = "statement",
          title: str | None = None) -> dict:
    slides = [{"slide_type": "cover", "title": "测试用封面标题"},
              {"slide_type": "toc", "items": [{"no": "01", "text": "第一章"},
                                              {"no": "02", "text": "第二章"},
                                              {"no": "03", "text": "第三章"}]}]
    for i in range(n_content):
        s = _slide(kind, i + 1)
        if title:
            s["title"] = title
        slides.append(s)
    slides.append({"slide_type": "end", "title": "谢谢观看", "sections": []})
    d: dict = {"title": "测试 deck", "slides": slides}
    if brief:
        d["brief"] = brief
    return d


def _messages(report: dict, field: str) -> list[str]:
    return [i["message"] for i in report["warnings"] if i["field"] == field]


_GOOD_TITLE = "回归通过率提升靠的是门禁而不是加班"


def _summary_warnings(summary: str, title: str = _GOOD_TITLE) -> list[str]:
    """把一页 statement 的 summary 换成待测文本，返回 summary 相关的 warning。"""
    deck = _deck(None, 1, "statement")
    deck["slides"][2]["summary"] = summary
    deck["slides"][2]["title"] = title
    return _messages(qc.check(validate_dict(deck)), "summary")


def main() -> int:
    print("== 0. 方向清单 ==")
    names = briefs.available()
    _check("研讨型/汇报型都在", {"seminar", "briefing"} <= set(names), f"实际 {names}")
    for n in names:
        b = briefs.load(n)
        _check(f"{n} 有 label/writing/check", bool(b.get("label") and b.get("writing")
                                                   and b.get("check")))

    print("== 1. 汇报型：页数不设上限，超过 30 页才先问用户 ==")
    r = qc.check(validate_dict(_deck("briefing", 20)))
    _check("20 页不再报「偏多」（2026-09 起页数不限）",
           not any("偏多" in m for m in _messages(r, "brief")),
           str(_messages(r, "brief")))
    r = qc.check(validate_dict(_deck("briefing", 35)))
    _check("35 页报「该先问用户」", any("先问用户" in m for m in _messages(r, "brief")),
           str(_messages(r, "brief")))

    print("== 2. 汇报型：每页支撑论据 > 3 条（图表 insights 不受此限）==")
    # 2026-09 修：`maxSupportsPerPage` 只管 `points`（statement 的支撑论据）。
    # 原先它顺手管了 `chart.insights`，而 schema 要求 insights **2~4 条**（必填）——
    # 局部规则压过全局规则，写 4 条合法洞察反而被点名。这条用例以前固化了那个错误行为。
    r = qc.check(validate_dict(_deck("briefing", 5, "statement")))
    _check("statement 的 4 条论据被点名", any("不超过 3 条" in i["message"]
                                             for i in r["warnings"] if i["field"] == "supports"),
           str([i["message"][:60] for i in r["warnings"] if i["field"] == "supports"]))
    r = qc.check(validate_dict(_deck("briefing", 5, "chart")))
    _check("chart 的 4 条 insights 不再误伤（schema 要求 2~4 条）",
           not any(i["field"] == "supports" for i in r["warnings"]),
           str([i["message"][:60] for i in r["warnings"] if i["field"] == "supports"]))

    print("== 3. 汇报型：标题里的无信息量词 ==")
    r = qc.check(validate_dict(_deck("briefing", 5, title="渠道建设相关情况介绍及进展")))
    _check("「情况介绍」被点名", any("情况介绍" in i["message"]
                                     for i in r["warnings"] if i["field"] == "title"))

    print("== 4. 研讨型：页数不设上限 ==")
    r = qc.check(validate_dict(_deck("seminar", 40, "table")))
    _check("40 页不报「偏多」", not any("偏多" in m for m in _messages(r, "brief")),
           str(_messages(r, "brief")))
    r = qc.check(validate_dict(_deck("seminar", 51, "table")))
    _check("51 页报「该先问用户」", any("先问用户" in m for m in _messages(r, "brief")),
           str(_messages(r, "brief")))

    print("== 5. 研讨型：每页要有图表表格 ==")
    r = qc.check(validate_dict(_deck("seminar", 8, "statement")))
    _check("缺 table/chart 被点名", any("缺：table、chart" in m
                                        for m in _messages(r, "brief")),
           str(_messages(r, "brief")))
    _check("图表占比过低被点名", any("表格 + 图表" in m for m in _messages(r, "brief")))

    r = qc.check(validate_dict(_deck("seminar", 2, "chart")))
    _check("图型不够 3 种被点名", any("不同图型至少 3 种" in m
                                      for m in _messages(r, "brief")),
           str(_messages(r, "brief")))

    good = _deck("seminar", 8, "chart")
    good["slides"][2:4] = [_slide("table", 1), _slide("table", 2)]
    r = qc.check(validate_dict(good))
    _check("图表充足 + 图型够多：无 brief 问题", not _messages(r, "brief"),
           str(_messages(r, "brief")))

    print("== 6. 方向名写错 ==")
    r = qc.check(validate_dict(_deck("seminer", 12, "table")))
    _check("未知方向是 error 且列出可选项",
           any("seminar" in i["message"] and "briefing" in i["message"]
               for i in r["errors"]),
           str(r["errors"]))

    print("== 7. 不指定方向：行为与从前一致 ==")
    r = qc.check(validate_dict(_deck(None, 3)))
    _check("不产生 brief 类 warning", not _messages(r, "brief"), str(_messages(r, "brief")))
    _check("stats.brief 为 None", r["stats"]["brief"] is None)

    print("== 8. 方向只出 warning，不拦渲染 ==")
    r = qc.check(validate_dict(_deck("briefing", 20)))
    _check("超页数不进 errors", not r["errors"], str(r["errors"]))

    print("== 9. 结论段（summary）硬指标 ==")
    good = _summary("统一评测平台")
    _check("合规结论段无 warning", not _summary_warnings(good),
           str(_summary_warnings(good)))

    cases = [
        ("章节坐标被点名", "1.2 节给出四条驱动力，其中只有第一条附了可核对的实验数字。",
         "章节编号"),
        ("自指与原文被点名", "原文给出的三条约束都写在上面，本节只做归纳说明。", "指向原文"),
        ("过短被点名", "评测平台必须一次做对四件事。", "只有"),
        ("过长被点名", good + good, "超出"),
        ("只有一句被点名", good.replace("。", "，") + "。", "只有 1 句"),
        ("填充语被点名", "值得注意的是，" + good, "填充语"),
        ("超句数被点名", good + "而且还要再补一句解释。还要再补第二句说明。再补第三句。",
         "句（要求"),
        ("复述标题被点名", "回归通过率提升靠的是门禁而不是加班，这一点在本季度数据里非常明显。",
         "复述了标题"),
    ]
    for name, summary, expect in cases:
        msgs = _summary_warnings(summary)
        _check(name, any(expect in m for m in msgs), str(msgs)[:90])

    _check("结论段问题只出 warning，不拦渲染",
           not qc.check(validate_dict(_deck(None, 1, "statement")))["errors"])

    print("== 10. 方向规约与技能文件不脱节 ==")
    for n in briefs.available():
        skill = REPO / briefs.load(n).get("skill", "")
        _check(f"{n} 指向的技能文件存在", skill.is_file(), str(skill))

    print("== 11. 方向规约覆盖新密度规则（并页 / 左右分置 / 正文下限）==")
    for n in briefs.available():
        # 2026-09 起 `writing:` 只留方向要点，完整规范在方向 SKILL.md —— 两者合起来看，
        # 这样"规则被删掉"仍然会被这条用例拦住，同时不必在两个文件里重复同一段话。
        b = briefs.load(n)
        skill = b.get("skill") or ""
        text = briefs.writing_of(n) + (
            (REPO / skill).read_text(encoding="utf-8") if skill else "")
        _check(f"{n} 规定「内容太少就并页 / 左右分置」",
               ("并页" in text) and ("panel_2col" in text), text[:80])
        _check(f"{n} 规定「正文不许精简过头」",
               "精简过头" in text, text[:80])

    print("== 12. 三个用户维度：风格预设 + 速度档 ==")
    # 风格：5 套配色各绑定一种卡片，用户只选一项；--card 可覆盖。
    # 锁的是"预设表与 theme 的白名单对齐"—— 少一个键就会静默回退，用户换了等于没换。
    _check("每套配色都有中文标签",
           set(theme.STYLE_LABELS) == set(theme._THEMES),
           f"{sorted(theme.STYLE_LABELS)} vs {sorted(theme._THEMES)}")
    _check("每套配色的预设卡片都是合法卡片名",
           all(v in theme.CARD_STYLES for v in theme.STYLE_CARD_PRESET.values()),
           str(theme.STYLE_CARD_PRESET))
    _check("预设卡片覆盖全部 5 套配色",
           set(theme.STYLE_CARD_PRESET) == set(theme._THEMES),
           str(sorted(theme.STYLE_CARD_PRESET)))
    _check("未登记的配色回退默认卡片",
           theme.resolve_card("不存在的主题") == theme.DEFAULT_CARD,
           theme.resolve_card("不存在的主题"))
    _check("显式 --card 优先于预设",
           theme.resolve_card("bright", "tinted") == "tinted",
           theme.resolve_card("bright", "tinted"))
    _check("外观一句话标签带中文", " + " in theme.appearance_label("forest"),
           theme.appearance_label("forest"))

    # 速度档：只改报告详略与回炉要求，**检查项一档都不能少**。
    for n in briefs.available():
        slow = briefs.speed_of(n, "slow")
        fast = briefs.speed_of(n, "fast")
        _check(f"{n} 慢档要求 warning 归零、快档不要求",
               slow["requireZeroWarnings"] and not fast["requireZeroWarnings"],
               f"{slow} / {fast}")
        _check(f"{n} 慢档评审大纲 + 骨架预览，快档都跳过",
               slow["reviewOutline"] and slow["skeletonPreview"]
               and not fast["reviewOutline"] and not fast["skeletonPreview"],
               f"{slow} / {fast}")
        _check(f"{n} 慢档自动带备注、快档不带",
               slow["notes"] and not fast["notes"], f"{slow} / {fast}")

    deck = validate_dict(_deck("seminar", 8, "statement"))   # 会报一堆 warning
    r_slow = qc.check(deck, speed="slow")
    r_fast = qc.check(deck, speed="fast")
    _check("快档检查项一项不少（error 数完全相同）",
           len(r_slow["errors"]) == len(r_fast["errors"]),
           f"{len(r_slow['errors'])} vs {len(r_fast['errors'])}")
    _check("快档把重复 warning 折叠了（条数变少）",
           len(r_fast["warnings"]) < len(r_slow["warnings"]),
           f"{len(r_slow['warnings'])} → {len(r_fast['warnings'])}")
    _check("快档不折叠 deck 级红线（方向红线一条都不能藏）",
           sum(1 for w in r_fast["warnings"] if w["type"] == "deck")
           == sum(1 for w in r_slow["warnings"] if w["type"] == "deck"),
           "deck 级条数不一致")
    _check("未知方向的速度约定回退默认值",
           briefs.speed_of(None, "slow")["requireZeroWarnings"] is True,
           str(briefs.speed_of(None, "slow")))

    # 慢档只在评审大纲时打断用户一次：超页数的取舍也要并进那一次，不能出图后再开一轮
    r_ask = qc.check(validate_dict(_deck("seminar", 51, "table")))
    _check("超页数：点名要并进大纲评审那一次问完",
           any("大纲评审" in m for m in _messages(r_ask, "brief")),
           str(_messages(r_ask, "brief")))
    _check("超页数：仍不许靠删内容达标",
           any("不许删内容" in m or "删掉" in m for m in _messages(r_ask, "brief")),
           str(_messages(r_ask, "brief")))

    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项失败：")
        for f in _failures:
            print(f"  - {f}")
        return 1
    print("[PASS] 研讨型/汇报型规约无回归")
    return 0


if __name__ == "__main__":
    sys.exit(main())
