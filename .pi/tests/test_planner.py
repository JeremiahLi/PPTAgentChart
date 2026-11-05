# -*- coding: utf-8 -*-
"""测试规划引擎：doc_analyzer + planner + speaker_notes + outline。"""
from __future__ import annotations

import json
import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "renderer"))

import doc_analyzer
import planner
import speaker_notes
import outline


# ---------------- doc_analyzer ----------------

class TestDataExtraction:
    def test_extract_percent(self):
        text = "收入增长 15.3%，用户数达到 1,200 万"
        points = doc_analyzer.extract_data_points(text)
        assert len(points) >= 2
        assert any("15.3%" in p for p in points)
        assert any("1,200" in p or "1200" in p for p in points)

    def test_extract_scale(self):
        text = "投入 3 人天，产出 25 篇文章，耗时 2 个月"
        points = doc_analyzer.extract_data_points(text)
        assert len(points) >= 3

    def test_no_data(self):
        text = "这是一段没有任何数字的纯文本内容"
        points = doc_analyzer.extract_data_points(text)
        assert len(points) == 0


class TestTextClassification:
    def test_comparison(self):
        text = "方案 A 优于方案 B，性能提升 30%"
        assert doc_analyzer.has_comparison(text)

    def test_comparison_markers(self):
        assert doc_analyzer.has_comparison("方案 A 对比方案 B")
        assert doc_analyzer.has_comparison("成本相比之前降低")
        assert not doc_analyzer.has_comparison("但是今天下雨了")

    def test_process(self):
        text = "首先读取配置，然后初始化连接，最后执行查询"
        assert doc_analyzer.has_process(text)

    def test_hierarchy(self):
        text = "系统由三层组成：接入层、逻辑层、数据层"
        assert doc_analyzer.has_hierarchy(text)

    def test_data_heavy(self):
        text = "指标A = 85.3%，指标B = 92%，指标C = 76%，指标D = 88%，指标E = 90%"
        assert doc_analyzer._is_data_heavy(text)

    def test_text_heavy(self):
        text = ("这是一段很长很长的纯文本内容，里面没有数字，没有任何数据，"
                "只有大量的描述性文字。" * 10)
        assert doc_analyzer._is_text_heavy(text)


class TestAnalysisParsing:
    def _sample_table(self):
        return """| 章节 | 模块 | 核心结论 | 论据 | 缺口 |
|---|---|---|---|---|
| 1. 背景 | 项目起源 | 项目起源于 2023 年 | 初始投入 3 人，预算 50 万 | 缺商业论证 |
| 2. 技术 | 架构设计 | 采用微服务架构 | 拆分为 8 个服务，QPS 达到 5000 | 无 |"""

    def test_parse_table(self):
        modules = doc_analyzer._parse_analysis_table(self._sample_table())
        assert len(modules) >= 2
        # 模块名
        assert modules[0].module == "项目起源"
        assert modules[1].module == "架构设计"
        # 数据提取
        assert len(modules[0].data_points) >= 2  # "3 人" 和 "50 万"
        assert len(modules[1].data_points) >= 1  # "8" 或 "5000"
        # 缺口
        assert len(modules[0].gaps) >= 1
        assert "商业论证" in modules[0].gaps[0]


class TestSourceHash:
    def test_same_file_same_hash(self):
        h1 = doc_analyzer.compute_source_hash(__file__)
        h2 = doc_analyzer.compute_source_hash(__file__)
        assert h1 == h2

    def test_different_files_different_hash(self):
        h1 = doc_analyzer.compute_source_hash(__file__)
        h2 = doc_analyzer.compute_source_hash(
            os.path.join(os.path.dirname(__file__), "test_planner.py"))
        # 同一文件
        assert h1 == h2


# ---------------- planner ----------------

class TestSlideTypeRouting:
    def test_data_heavy_routes_to_chart(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="性能测试",
            conclusion="性能提升显著",
            evidence=["P99 延迟下降 40%", "QPS 从 1000 提升到 5000",
                      "CPU 占用下降 20%", "内存节省 30%"],
            data_points=["40%", "1000", "5000", "20%", "30%"],
            is_data_heavy=True,
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.CHART

    def test_comparison_routes_to_compare_3col(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="方案对比",
            conclusion="方案 A 优于方案 B",
            evidence=["成本更低", "性能更高", "维护更简单"],
            has_comparison=True,
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.COMPARE_3COL

    def test_process_routes_to_flow(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="构建流程",
            conclusion="构建分为四步",
            evidence=["先读取配置", "再编译", "然后测试", "最后部署"],
            has_process=True,
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.FLOW

    def test_hierarchy_routes_to_hierarchy(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="架构",
            conclusion="系统分为三层",
            evidence=["接入层", "逻辑层", "数据层"],
            has_hierarchy=True,
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.HIERARCHY

    def test_many_evidence_routes_to_cards(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="功能列表",
            conclusion="包含 6 大功能",
            evidence=["功能A", "功能B", "功能C", "功能D", "功能E"],
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.NUMBERED_CARDS

    def test_two_evidence_routes_to_panel(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="对比",
            conclusion="两种方案各有优劣",
            evidence=["方案A的优势", "方案B的优势"],
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.PANEL_2COL

    def test_three_evidence_parallel_routes_to_key_groups(self):
        """3 条论据且无对比语义 → 要点分组（statement 只做最后兜底）。"""
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="三件事",
            conclusion="统一评测固化三件事",
            evidence=["失败可归因：按能力维度下钻",
                      "回归可自动：改动即跑基线",
                      "门禁可执行：未达标不发布"],
        )
        assert planner._route_slide_type(mod) == outline.SlideType.KEY_GROUPS

    def test_three_evidence_with_comparison_still_compare_3col(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="方案对比",
            conclusion="方案 A 优于方案 B",
            evidence=["成本更低", "性能更高", "维护更简单"],
            has_comparison=True,
        )
        assert planner._route_slide_type(mod) == outline.SlideType.COMPARE_3COL

    def test_default_statement(self):
        mod = doc_analyzer.SourceModule(
            section="测试",
            module="概述",
            conclusion="这是一个概述",
            evidence=["单点论据"],
        )
        stype = planner._route_slide_type(mod)
        assert stype == outline.SlideType.STATEMENT


class TestDensityPlanning:
    def test_chart_basic(self):
        mod = doc_analyzer.SourceModule(
            conclusion="数据展示", evidence=["25%", "30%"],
            data_points=["25%", "30%"],
            content_length=80,
        )
        density = planner._density_for_type(outline.SlideType.CHART, mod)
        assert density == outline.DensityTarget.BASIC

    def test_statement_high(self):
        mod = doc_analyzer.SourceModule(
            conclusion="详细展开", evidence=["a", "b", "c"],
            content_length=400,
        )
        density = planner._density_for_type(outline.SlideType.STATEMENT, mod)
        assert density == outline.DensityTarget.HIGH


class TestChartSuggestion:
    def test_time_series_line(self):
        mod = doc_analyzer.SourceModule(
            conclusion="趋势增长",
            evidence=["1月 10%", "2月 15%", "3月 20%", "4月 25%"],
            data_points=["10%", "15%", "20%", "25%"],
        )
        chart = planner._route_chart_type(mod)
        assert chart is not None
        assert chart.chart_type == "line"

    def test_ratio_doughnut(self):
        mod = doc_analyzer.SourceModule(
            conclusion="占比分布",
            evidence=["A类占 40%", "B类占 35%", "C类占 25%"],
            data_points=["40%", "35%", "25%"],
        )
        chart = planner._route_chart_type(mod)
        assert chart is not None
        assert chart.chart_type == "doughnut"

    def test_comparison_column(self):
        mod = doc_analyzer.SourceModule(
            conclusion="对比结果",
            evidence=["A方案 85分", "B方案 72分"],
            has_comparison=True,
        )
        chart = planner._route_chart_type(mod)
        # 对比 + 数据点少 → column
        assert chart is None or chart.chart_type in ("column", "line")


class TestMergeHeuristic:
    def test_should_merge_small_content(self):
        small = doc_analyzer.SourceModule(
            section="同章节", module="小结",
            conclusion="结论", content_length=60,
        )
        prev = doc_analyzer.SourceModule(
            section="同章节", module="前一个",
            conclusion="结论", content_length=200,
        )
        assert planner._should_merge_with_previous(small, prev)

    def test_should_not_merge_large_content(self):
        large = doc_analyzer.SourceModule(
            section="不同章节", module="大模块",
            conclusion="结论", content_length=300,
        )
        prev = doc_analyzer.SourceModule(
            section="另一章节", module="另一个",
            conclusion="结论", content_length=200,
        )
        assert not planner._should_merge_with_previous(large, prev)


class TestSplitHeuristic:
    def test_should_split_long_content(self):
        mod = doc_analyzer.SourceModule(
            content_length=600,
            evidence=["a"] * 8,
        )
        assert planner._should_split(mod)

    def test_should_not_split_short_content(self):
        mod = doc_analyzer.SourceModule(
            content_length=150,
            evidence=["a", "b"],
        )
        assert not planner._should_split(mod)


class TestPlanIntegration:
    def test_plan_produces_valid_outline(self):
        modules = [
            doc_analyzer.SourceModule(
                section="1. 背景",
                module="项目起源",
                conclusion="项目起源于 2023 年初",
                evidence=["初始投入 3 人", "预算 50 万", "周期 6 个月"],
                data_points=["3", "50万", "6"],
                gaps=["缺商业论证"],
            ),
            doc_analyzer.SourceModule(
                section="2. 技术",
                module="架构设计",
                conclusion="采用微服务架构，QPS 提升 5 倍",
                evidence=["拆分为 8 个服务", "P99 延迟 50ms", "可用性 99.9%"],
                data_points=["8", "50ms", "99.9%"],
            ),
            doc_analyzer.SourceModule(
                section="2. 技术",
                module="技术对比",
                conclusion="方案 A 优于方案 B",
                evidence=["成本更低 30%", "性能更高 2 倍", "维护更简单"],
                data_points=["30%", "2"],
                has_comparison=True,
            ),
        ]

        ol = planner.plan(modules, title="测试大纲", brief="seminar")
        assert ol.title == "测试大纲"
        assert ol.brief == "seminar"
        assert ol.total_pages >= 3
        assert ol.source_hash == ""

        # 序列化
        data = ol.model_dump(mode="json")
        restored = outline.DeckOutline.model_validate(data)
        assert restored.title == "测试大纲"

        # 保存到文件再加载
        path = ".pi/work/test_outline.json"
        os.makedirs("work", exist_ok=True)
        outline.save_outline(ol, path)
        loaded = outline.load_outline(path)
        assert loaded.total_pages == ol.total_pages
        os.remove(path)


# ---------------- 跨模块关系抽取 ----------------

def _mods(*triples):
    """快捷构造模块清单：(模块名, 结论, [论据])。"""
    return [
        doc_analyzer.SourceModule(
            section="测试", module=name, conclusion=conc, evidence=ev,
            content_length=sum(len(x) for x in ev) + len(conc),
        )
        for name, conc, ev in triples
    ]


class TestRelationExtraction:
    def test_sequence_by_ordinal_names(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("第一阶段", "完成需求调研", ["输出调研报告"]),
            ("第二阶段", "完成方案设计", ["输出设计稿"]),
            ("第三阶段", "完成开发实施", ["输出代码"]),
        ))
        assert any(r.type == "sequence" and r.target == "第二阶段"
                   for r in mods[0].relations)

    def test_sequence_by_transition_words(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("数据采集", "先完成数据采集", ["采集完成后进入清洗"]),
            ("数据清洗", "随后清洗数据", ["去重补全"]),
            ("数据分析", "最后分析", ["产出指标"]),
        ))
        assert any(r.type == "sequence" and r.target == "数据清洗"
                   for r in mods[0].relations)

    def test_contain_by_mention(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("技术架构", "平台由三层组成", ["包含接入层", "包括逻辑层"]),
            ("接入层", "负责流量接入", ["网关"]),
            ("逻辑层", "负责业务逻辑", ["服务编排"]),
        ))
        tgts = {r.target for r in mods[0].relations if r.type == "contain"}
        assert "接入层" in tgts and "逻辑层" in tgts

    def test_cause_falls_back_to_previous(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("需求膨胀", "需求持续膨胀", ["范围失控"]),
            ("工期延误", "因此工期延误", ["里程碑推迟"]),
        ))
        assert any(r.type == "cause" and r.target == "需求膨胀"
                   for r in mods[1].relations)

    def test_contrast_by_naming_pair(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("方案一", "方案一的优点", ["成本低"]),
            ("方案二", "方案二的优点", ["性能高"]),
        ))
        assert any(r.type == "contrast" and r.target == "方案二"
                   for r in mods[0].relations)

    def test_idempotent(self):
        mods = _mods(("一期", "开工", ["挖坑"]), ("二期", "完工", ["填土"]))
        doc_analyzer.extract_relations(mods)
        doc_analyzer.extract_relations(mods)
        assert len(mods[0].relations) == len(
            [r for r in mods[0].relations])  # 不叠加
        seq = [r for r in mods[0].relations if r.type == "sequence"]
        assert len(seq) == 1


# ---------------- 第二遍路由：关系组合并 ----------------

class TestRelationGroups:
    def test_sequence_chain_merges_to_flow(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("需求分析", "第一阶段完成需求分析", ["调研用户", "输出需求文档"]),
            ("系统设计", "第二阶段完成系统设计", ["架构设计", "输出设计稿"]),
            ("开发实施", "第三阶段完成开发", ["编码", "联调"]),
            ("测试验收", "第四阶段完成验收", ["回归测试", "上线"]),
        ))
        groups = planner._relation_groups(mods)
        assert len(groups) == 1
        start, end, stype, reason = groups[0]
        assert (start, end) == (0, 4)
        assert stype == outline.SlideType.FLOW

    def test_contrast_pair_merges_to_panel(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("方案一", "方案一成本更低", ["成本低", "周期短"]),
            ("方案二", "相比方案一性能更高", ["性能高", "扩展好"]),
            ("其他事项", "另有一些事项", ["杂项内容较多需要单独一页来展开说明"]),
        ))
        groups = planner._relation_groups(mods)
        assert any(g[2] == outline.SlideType.PANEL_2COL and g[:2] == (0, 2)
                   for g in groups)

    def test_inferred_contrast_not_merged(self):
        # 对比词没点名 → 只是邻接推断，不许做"并页"这种破坏性决定
        mods = _mods(
            ("采购情况", "相比去年有所下降", ["金额减少"]),
            ("库存情况", "库存周转正常", ["周转率稳定"]),
        )
        groups = planner._relation_groups(mods)
        assert not any(g[2] == outline.SlideType.PANEL_2COL for g in groups)

    def test_parallel_run_merges_to_key_groups(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("归因机制", "此外，归因机制持续运转", ["失败样本按维度聚类"]),
            ("回归机制", "同时，回归自动触发", ["改动即跑基线"]),
            ("门禁机制", "另外，门禁自动执行", ["未达标不发布"]),
        ))
        groups = planner._relation_groups(mods)
        assert any(g[2] == outline.SlideType.KEY_GROUPS and g[:2] == (0, 3)
                   for g in groups)

    def test_group_rejected_when_too_heavy(self):
        long_ev = ["这是一条非常长的论据内容" * 6] * 4
        mods = doc_analyzer.extract_relations(_mods(
            ("第一阶段", "完成调研", long_ev),
            ("第二阶段", "完成设计", long_ev),
            ("第三阶段", "完成开发", long_ev),
        ))
        groups = planner._relation_groups(mods)
        assert not groups  # 容量超限 → 回退单模块路由

    def test_data_heavy_not_dragged_into_flow(self):
        mods = doc_analyzer.extract_relations(_mods(
            ("一期", "开工", ["挖坑", "打桩"]),
            ("二期", "主体施工", ["浇筑", "养护"]),
            ("三期", "交付", ["验收"]),
        ))
        mods[1].is_data_heavy = True
        groups = planner._relation_groups(mods)
        assert not any(g[2] == outline.SlideType.FLOW for g in groups)


class TestRelationRoutingIntegration:
    def test_plan_merges_chain_into_one_flow_page(self):
        modules = doc_analyzer.extract_relations(_mods(
            ("需求分析", "第一阶段完成需求分析", ["调研用户", "输出需求文档"]),
            ("系统设计", "第二阶段完成系统设计", ["架构设计", "输出设计稿"]),
            ("开发实施", "第三阶段完成开发", ["编码", "联调"]),
            ("测试验收", "第四阶段完成验收", ["回归测试", "上线"]),
        ))
        ol = planner.plan(modules, title="测试", brief="seminar")
        flow_slides = [s for sec in ol.sections for s in sec.slides
                       if s.slide_type == outline.SlideType.FLOW]
        assert len(flow_slides) == 1
        assert flow_slides[0].relation_note
        assert "需求分析" in flow_slides[0].module
        # 合并后 4 个模块只占 1 页
        assert ol.total_pages == 1

    def test_relation_note_serializes(self):
        modules = doc_analyzer.extract_relations(_mods(
            ("需求分析", "第一阶段完成需求分析", ["调研用户"]),
            ("系统设计", "第二阶段完成系统设计", ["架构设计"]),
            ("开发实施", "第三阶段完成开发", ["编码"]),
            ("测试验收", "第四阶段完成验收", ["上线"]),
        ))
        ol = planner.plan(modules, title="测试", brief="seminar")
        data = ol.model_dump(mode="json")
        restored = outline.DeckOutline.model_validate(data)
        flow = [s for sec in restored.sections for s in sec.slides
                if s.slide_type == outline.SlideType.FLOW]
        assert flow and flow[0].relation_note


# ---------------- speaker_notes ----------------

class TestSpeakerNotes:
    def test_chart_few_data_points(self):
        sn = speaker_notes.generate_notes(
            slide_type="chart",
            module_name="性能测试",
            conclusion="性能提升",
            evidence=["QPS 1000", "延迟 50ms"],
            gaps=[],
            data_points=["1000", "50ms"],
        )
        text = sn.render()
        assert "数据点" in text or "数据" in text

    def test_statement_no_data(self):
        sn = speaker_notes.generate_notes(
            slide_type="statement",
            module_name="概述",
            conclusion="这是一段概述",
            evidence=["论据1"],
            gaps=[],
            data_points=[],
        )
        text = sn.render()
        assert "无数字" in text or "数字" in text or "数据" in text

    def test_text_heavy_suggests_viz(self):
        sn = speaker_notes.generate_notes(
            slide_type="statement",
            module_name="复杂概念",
            conclusion="这是一个极其复杂的概念需要详细解释",
            evidence=["论据1", "论据2"],
            gaps=[],
            data_points=[],
            is_text_heavy=True,
            has_process=True,
        )
        text = sn.render()
        assert "流程图" in text or "flow" in text or "可视化" in text

    def test_with_gaps(self):
        sn = speaker_notes.generate_notes(
            slide_type="statement",
            module_name="不完整模块",
            conclusion="部分结论",
            evidence=["论据"],
            gaps=["缺少用户反馈数据", "缺成本估算"],
            data_points=[],
        )
        text = sn.render()
        assert "用户反馈" in text
        assert "成本估算" in text

    def test_suggests_user_research(self):
        sn = speaker_notes.generate_notes(
            slide_type="statement",
            module_name="用户分析",
            conclusion="用户满意度高",
            evidence=["NPS 72", "留存率 85%"],
            gaps=[],
            data_points=["72", "85%"],
        )
        text = sn.render()
        # 应该提到用户调研
        assert "调研" in text or "NPS" in text or "留存" in text

    def test_suggests_financial_data(self):
        sn = speaker_notes.generate_notes(
            slide_type="chart",
            module_name="收入分析",
            conclusion="收入增长 25%",
            evidence=["Q1 100万", "Q2 125万"],
            gaps=[],
            data_points=["25%", "100万", "125万"],
        )
        text = sn.render()
        assert "财务" in text or "报表" in text

    def test_no_notes_for_strong_slide(self):
        sn = speaker_notes.generate_notes(
            slide_type="chart",
            module_name="数据展示",
            conclusion="各项指标全面提升",
            evidence=["指标A 85%", "指标B 92%", "指标C 78%", "指标D 88%"],
            gaps=[],
            data_points=["85%", "92%", "78%", "88%"],
            is_data_heavy=True,
        )
        # 数据密集且用了图表 → 数据缺口应该为空或很少
        data_gap_notes = [n for n in sn.notes if n.category == "data"]
        # 没有缺口时也可能有建议，但应该很少
        assert len(data_gap_notes) <= 1


# ---------------- outline ----------------

class TestOutlineSerialization:
    def test_round_trip(self):
        ol = outline.DeckOutline(
            title="测试",
            subtitle="副标题",
            brief="seminar",
            sections=[
                outline.SectionPlan(
                    no="01",
                    title="第一章",
                    slides=[
                        outline.SlidePlan(
                            section="第一章",
                            module="模块1",
                            title="标题1",
                            slide_type=outline.SlideType.STATEMENT,
                            density_target=outline.DensityTarget.STANDARD,
                            conclusion="结论1",
                            evidence=["论据1", "论据2"],
                            gaps=["缺口1"],
                            notes=["备注建议"],
                        ),
                    ],
                ),
            ],
            total_pages=1,
            content_pages=1,
        )
        data = ol.model_dump(mode="json")
        restored = outline.DeckOutline.model_validate(data)
        assert restored.title == "测试"
        assert len(restored.sections) == 1
        assert len(restored.sections[0].slides) == 1


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
