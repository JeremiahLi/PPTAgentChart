# -*- coding: utf-8 -*-
"""自动内容分析：把原始文档/分析表转为结构化模块清单。

这是「规划引擎」的第一步 —— 过去完全靠 LLM 读一遍文档然后自己排大纲；
现在由 analyzer 自动完成「切分 → 提取 → 标注」，输出一个结构化的
SourceModule 列表，planner 再据此生成大纲。

为什么要自动做 ——
1. 用户说「再生成一遍更好的」时，文档没变就不用重分析
2. analyzer 用规则切分（按标题、数字密度、段落长度），比 LLM 稳定
3. 文档没为 PPT 化考虑（内容揉成一团）时，analyzer 自动识别哪些段落
   有数据可以图表化、哪些只能抽象成流程图/框架图
"""

from __future__ import annotations

import hashlib
import os
import re
from typing import Literal, Optional

from pydantic import BaseModel, Field


# ---------- ModuleRelation：模块之间的一条关系（第二遍路由的输入） ----------

class ModuleRelation(BaseModel):
    """模块 A 与模块 B 之间的关系。

    为什么要有这一层 —— 过去的版式路由只看单模块内部特征（有没有对比词/流程词），
    模块与模块之间的结构信息（谁先谁后、谁是谁的一部分、谁依赖谁）在切分完就丢了，
    导致大量本该合成一张图的内容被摊平成一排各自成页的卡片。
    relations 就是把这部分信息显式化，供 planner 的第二遍路由做合并/拆版决策。

    type 取值与典型触发：
    - sequence: 时序先后（"随后/接着/完成后"、模块名带阶段序号）
    - contrast: 对比（"相比/区别/差异"、命名对仗如 方案A/方案B）
    - contain:  整体-部分（"包含/分为/其中"且点名了其他模块）
    - cause:    因果（"导致/因此/由于"）
    - parallel: 并列（"此外/同时/另外"）
    - depend:   依赖（"基于/依托/取决于"，方向 = 依赖被指向者）
    """
    type: Literal["sequence", "contrast", "contain", "cause", "parallel", "depend"]
    target: str = Field("", description="目标模块名；空串 = 邻接的下一模块")
    evidence: str = Field("", description="触发依据（命中词/命名规则），可追溯")


class SourceModule(BaseModel):
    """从文档里切出的一个内容单元。"""
    # 来源
    section: str = ""           # 所属章节
    module: str = ""            # 模块名（从分析表提取或自动识别）
    source_file: str = ""       # 来源文件路径
    line_start: int = 0         # 在 analysis.md 的行号

    # 内容
    conclusion: str = ""        # 核心结论（判断句）
    evidence: list[str] = Field(default_factory=list)  # 论据
    gaps: list[str] = Field(default_factory=list)      # 缺口

    # 跨模块关系（extract_relations 填充）
    relations: list[ModuleRelation] = Field(default_factory=list)

    # 可计算指标
    data_points: list[str] = Field(default_factory=list)  # 数字/百分比/金额
    has_comparison: bool = False   # 是否包含对比关系
    has_process: bool = False      # 是否包含流程/步骤
    has_hierarchy: bool = False    # 是否包含层级/包含关系
    is_data_heavy: bool = False    # 数据密度高（适合图表）
    is_text_heavy: bool = False    # 纯文本密度高（适合抽象为图）
    content_length: int = 0        # 正文字数（估算）


# ---------- 数据提取 ----------

_NUM_RE = re.compile(
    r"(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*(?:%|％|个|家|条|张|层|步|倍|次|人|天|月|年|万|亿|元|分钟|小时|MB|GB|TB)?)"
    r"|(?:\d+\.\d+)"
)

_COMPARISON_MARKERS = ("vs", "vs.", "对比", "相比", "优于", "低于", "高于",
                       "胜过", "不及", "区别", "差异", "区分", "对照",
                       "然而", "相对", "相较于")

_PROCESS_MARKERS = ("首先", "然后", "接着", "最后", "第一步", "第二步",
                    "流程", "步骤", "阶段", "过程", "顺序", "循环",
                    "输入", "输出", "判断", "判定", "条件", "分支")

_HIERARCHY_MARKERS = ("包含", "由...组成", "分为", "层级", "层次",
                      "上下级", "父级", "子模块", "架构", "体系",
                      "树形", "树状", "从属", "依赖")

# 更灵活的模式匹配
_HIERARCHY_PATTERNS = [
    re.compile(r"由.{1,10}组成"),
    re.compile(r"分为.{1,10}"),
    re.compile(r"层级|层次|上下级|父级|子模块|架构|体系|树形|树状|从属|依赖|包含"),
]


def extract_data_points(text: str) -> list[str]:
    """提取文本里的数字/百分比/金额。"""
    return [m.group(0).strip() for m in _NUM_RE.finditer(text)
            if len(m.group(0).strip()) >= 2]


def has_comparison(text: str) -> bool:
    """判断文本是否包含对比关系。"""
    # 移除"而"和"但"，太通用，容易误判
    return any(m in text for m in _COMPARISON_MARKERS)


def has_process(text: str) -> bool:
    """判断文本是否包含流程/步骤描述。"""
    return any(m in text for m in _PROCESS_MARKERS)


def has_hierarchy(text: str) -> bool:
    """判断文本是否包含层级/包含关系。"""
    return any(p.search(text) for p in _HIERARCHY_PATTERNS)


def _is_data_heavy(text: str) -> bool:
    """数据密度判断：每 100 字含 3+ 数字点。"""
    if len(text) < 50:
        return False
    data_count = len(extract_data_points(text))
    return data_count / (len(text) / 100) >= 3.0


def _is_text_heavy(text: str) -> bool:
    """纯文本密度判断：长段落但数据稀疏，适合抽象为图。"""
    if len(text) < 200:
        return False
    data_count = len(extract_data_points(text))
    return data_count / (len(text) / 100) < 1.0


# ---------- 跨模块关系抽取（第二遍路由的输入） ----------

# 模块间关系触发词。全部当正则用（普通子串天然兼容），
# 命中后由 extract_relations 决定指向谁（点名 > 邻接推断）。
_RELATION_MARKERS: dict[str, tuple[str, ...]] = {
    "sequence": ("随后", "接着", "之后", "然后", "接下来", "完成后",
                 "在此之前", "在此基础上", "进入下一", "下一步"),
    "contrast": ("相比", "对比", "相反", "区别", "差异", "优于", "劣于",
                 "不同于", "相对于", "而.*则"),
    "contain": ("包含", "包括", "分为", "涵盖", "其中", "子模块",
                "由.{1,10}组成"),
    "cause": ("导致", "因此", "由于", "因为", "使得", "造成",
              "引发", "带来", "归因", "源于"),
    "parallel": ("此外", "同时", "另外", "以及", "并列", "与此同时",
                 "另一方面"),
    "depend": ("依赖", "基于", "依托", "取决于", "借助",
               "建立在.{1,6}之上", "以.{1,6}为基础"),
}

# 模块名里的阶段序号：第X阶段 / 第X步 / X期 / X部分 / X批
_ORDINAL_RE = re.compile(r"(?:第)?\s*([0-9一二三四五六七八九十]+)\s*(?:期|步|阶段|部分|批|章|节)")
_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _ordinal(name: str) -> Optional[int]:
    """从模块名里提取阶段序号（第一阶段→1、二期→2、3.→3），没有则 None。"""
    m = _ORDINAL_RE.search(name)
    if not m:
        return None
    s = m.group(1)
    if s.isdigit():
        return int(s)
    return _CN_NUM.get(s)


# 命名对仗：词干 + 结尾计数（方案一→("方案",1)、Phase B→("Phase ",2)）
_TRAILING_RE = re.compile(r"^(.{2,}?)([0-9一二三四五六七八九十]+|[A-Za-z])$")


def _name_counter(name: str) -> Optional[tuple[str, int]]:
    """把「方案一 / 方案二」这类对仗名拆成（词干, 序号），不对仗则 None。"""
    m = _TRAILING_RE.match(name.strip())
    if not m:
        return None
    stem, ctr = m.group(1), m.group(2)
    if ctr.isdigit():
        return stem, int(ctr)
    if ctr in _CN_NUM:
        return stem, _CN_NUM[ctr]
    if ctr.isalpha():
        return stem, ord(ctr.upper()) - ord("A") + 1
    return None


def _mentioned_modules(text: str, modules: list["SourceModule"],
                       exclude: int) -> list[int]:
    """找出 text 里点名了的其他模块（按名字出现，≥2 字才认）。"""
    hits = []
    for j, m in enumerate(modules):
        if j == exclude or not m.module or len(m.module) < 2:
            continue
        if m.module in text:
            hits.append(j)
    return hits


def _first_hit(text: str, rtype: str) -> str:
    """返回该关系类型第一个命中的触发词，用于 evidence 可追溯。"""
    for pat in _RELATION_MARKERS[rtype]:
        m = re.search(pat, text)
        if m:
            return m.group(0)
    return ""


def _hit_any(text: str, rtype: str) -> bool:
    return any(re.search(pat, text) for pat in _RELATION_MARKERS[rtype])


def extract_relations(modules: list[SourceModule]) -> list[SourceModule]:
    """两遍扫描模块清单，抽取跨模块关系（确定性规则，不依赖 LLM）。

    决策顺序（一条 (i, j) 只落一类关系，先命中先得）：
    1. 时序链：相邻模块名带递增阶段序号（一期→二期 / 第1步→第2步）
    2. 点名关系：i 的文本里出现 j 的模块名 + 触发词 → 按触发词类型落关系
    3. 邻接推断：触发词出现但没点名 → 因果/依赖指向前一模块，
       时序/并列/对比指向后一模块（这是文档叙述的自然方向）
    4. 命名对仗：相邻模块名共享词干、仅序号/字母不同（方案A/方案B）→ 对比

    同一对模块可能同时有多类关系（如 A 包含 B 且 B 依赖 A），全部保留。
    """
    for m in modules:
        m.relations = []  # 幂等：重复调用不叠加

    n = len(modules)
    for i, mod in enumerate(modules):
        text = f"{mod.conclusion} {' '.join(mod.evidence)}"
        if not text.strip():
            continue

        # 1) 时序链：阶段序号递增的相邻模块（序号可能写在模块名或结论里）
        my_ord = _ordinal(mod.module) or _ordinal(mod.conclusion)
        if my_ord is not None and i + 1 < n:
            nxt = modules[i + 1]
            nxt_ord = _ordinal(nxt.module) or _ordinal(nxt.conclusion)
            if nxt_ord is not None and nxt_ord == my_ord + 1:
                mod.relations.append(ModuleRelation(
                    type="sequence", target=nxt.module,
                    evidence=f"阶段序号 {my_ord}→{nxt_ord}"))

        # 2) + 3) 逐类扫触发词
        for rtype in ("contrast", "contain", "cause", "depend",
                      "sequence", "parallel"):
            if not _hit_any(text, rtype):
                continue
            ev = _first_hit(text, rtype)
            mentioned = _mentioned_modules(text, modules, exclude=i)
            if mentioned:
                # 点名了：contained 可能点名多个，其余类型取第一个
                targets = mentioned if rtype == "contain" else mentioned[:1]
                for j in targets:
                    mod.relations.append(ModuleRelation(
                        type=rtype, target=modules[j].module, evidence=ev))
            elif i > 0 and rtype in ("cause", "depend"):
                # 没点名：因果/依赖回指前文（"由于上文…""基于上述…"）
                mod.relations.append(ModuleRelation(
                    type=rtype, target=modules[i - 1].module,
                    evidence=ev + "（邻接推断）"))
            elif i + 1 < n and rtype in ("sequence", "parallel", "contrast"):
                # 没点名：时序/并列/对比顺承后文
                mod.relations.append(ModuleRelation(
                    type=rtype, target=modules[i + 1].module,
                    evidence=ev + "（邻接推断）"))

        # 4) 命名对仗：方案一 vs 方案二 / 第一步 vs 第二步 / Phase A vs Phase B
        if i + 1 < n:
            mine = _name_counter(mod.module)
            theirs = _name_counter(modules[i + 1].module)
            if (mine and theirs
                    and mine[0] == theirs[0] and mine[1] != theirs[1]
                    and abs(mine[1] - theirs[1]) == 1
                    and not any(r.type == "contrast"
                                and r.target == modules[i + 1].module
                                for r in mod.relations)):
                mod.relations.append(ModuleRelation(
                    type="contrast", target=modules[i + 1].module,
                    evidence="命名对仗"))

    return modules


# ---------- 分析：analysis.md → 模块清单 ----------

def _row_to_module(cells: list[str], source_file: str = "",
                   line_no: int = 0) -> Optional[SourceModule]:
    """把 analysis.md 的一行表格转成 SourceModule。"""
    if len(cells) < 4:
        return None

    section = cells[0].strip()
    module = cells[1].strip()
    conclusion = cells[2].strip()
    evidence_text = cells[3].strip()
    gaps_text = cells[4].strip() if len(cells) > 4 else ""

    if not module or module in ("无", "-", "—"):
        return None

    # 合并所有文本用于指标计算
    full_text = f"{conclusion} {evidence_text}"

    # 解析论据（按分号/换行/数字分隔）
    evidence_items = _split_evidence(evidence_text)

    # 解析缺口
    gaps = [g.strip() for g in re.split(r"[;；\n]", gaps_text)
            if g.strip() and g.strip() not in ("无", "-", "—")]

    data_points = extract_data_points(full_text)

    return SourceModule(
        section=section,
        module=module,
        source_file=source_file,
        line_start=line_no,
        conclusion=conclusion,
        evidence=evidence_items,
        gaps=gaps,
        data_points=data_points,
        has_comparison=has_comparison(full_text),
        has_process=has_process(full_text),
        has_hierarchy=has_hierarchy(full_text),
        is_data_heavy=_is_data_heavy(full_text),
        is_text_heavy=_is_text_heavy(full_text),
        content_length=len(full_text),
    )


def _split_evidence(text: str) -> list[str]:
    """把论据字段拆成条目。"""
    # 优先按换行 / 分号 / 句号分隔
    items = re.split(r"[;；\n]", text)
    result = []
    for item in items:
        item = item.strip()
        if not item:
            continue
        # 过滤掉纯数字或太短的
        if len(item) < 4:
            continue
        result.append(item)
    return result


def _parse_analysis_table(text: str, source_file: str = "") -> list[SourceModule]:
    """从 analysis.md 的 markdown 表格解析所有行。"""
    modules = []
    for line_no, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue
        cells = [c.strip() for c in stripped.strip("|").split("|")]
        if len(cells) < 3:
            continue
        joined = "".join(cells).replace("-", "").replace(":", "").strip()
        if not joined or cells[0].startswith("章节"):
            continue
        mod = _row_to_module(cells, source_file, line_no)
        if mod:
            modules.append(mod)
    return modules


def analyze_analysis_md(path: str) -> list[SourceModule]:
    """分析 .pi/work/analysis.md，返回结构化模块清单。"""
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    return _parse_analysis_table(text, path)


def analyze_raw_text(path: str, section_prefix: str = "") -> list[SourceModule]:
    """分析原始文本/Markdown（没有 analysis.md 时直接分析原文）。

    按标题切分：`#` = 章节，`##`/`###` = 模块，标题正文 = 该模块的内容。

    （旧实现用 `re.split` 带捕获组切标题，拿到的是 "#"/"##" 这串**标记本身**，
    结果 `section` 永远是空串、标题还混进了模块正文 —— `--emit-analysis`
    生成的骨架会全错，这里改成按行读、标题与正文分开。）
    """
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()

    modules: list[SourceModule] = []
    state = {"section": "", "module": ""}
    buf: list[str] = []

    def flush():
        body = "\n".join(buf).strip()
        buf.clear()
        if len(body) < 30:
            return
        paragraphs = [p.strip() for p in re.split(r"\n\s*\n", body) if p.strip()]
        module = state["module"] or (paragraphs[0].split("\n")[0][:30] if paragraphs else "")
        sentences = re.split(r"[。！？]", body)
        conclusion = sentences[-2].strip() if len(sentences) > 1 else sentences[0].strip()
        modules.append(SourceModule(
            section=state["section"] or section_prefix,
            module=module.strip(),
            source_file=path,
            line_start=0,
            conclusion=conclusion[:200],
            evidence=_split_evidence(body),
            gaps=[],
            data_points=extract_data_points(body),
            has_comparison=has_comparison(body),
            has_process=has_process(body),
            has_hierarchy=has_hierarchy(body),
            is_data_heavy=_is_data_heavy(body),
            is_text_heavy=_is_text_heavy(body),
            content_length=len(body),
        ))

    for line in text.splitlines():
        heading = re.match(r"^(#{1,6})\s+(.+?)\s*$", line)
        if heading:
            flush()
            level, title = len(heading.group(1)), heading.group(2).strip()
            if level == 1:
                state["section"], state["module"] = title, ""
            else:
                state["module"] = title
            continue
        buf.append(line)
    flush()
    return modules


# ---------- analysis.md 骨架（--emit-analysis） ----------

_ANALYSIS_HEAD = """# 输入分析

> 这张表有两个下游：`plan.py` 拿它规划大纲，`render.py --source` / `build.py --source`
> 拿它核对"原文里的论据有没有被丢掉"。所以**别删模块**，只补内容。
>
> - 章节 / 模块：已按原文标题自动切出，可以合并、可以改名；
> - 核心结论：写一句**判断句**（"X 是 Y 的前提"），不要写"介绍 X"；
> - 论据：分号分隔 2-4 条，**数字优先**（覆盖率、时长、人数、金额）；
> - 缺口：没有就写 `无`；有就写"缺什么 + 找谁要"，研讨型会在交付时原样列给用户。
>
> 改完直接重跑（**不带** `--emit-analysis`）：
> `python .pi/renderer/plan.py .pi/work/analysis.md -o .pi/work/outline.json --title "..." --brief seminar`

| 章节 | 模块 | 核心结论 | 论据 | 缺口 |
|---|---|---|---|---|
"""


def _cell(text: str, limit: int = 120) -> str:
    """表格单元格消毒：竖线会把列切断，换行会把行切断。"""
    s = re.sub(r"\s*[\r\n]+\s*", " ", (text or "").strip())
    s = s.replace("|", "／").replace("\t", " ")
    s = re.sub(r"\s{2,}", " ", s)
    if len(s) > limit:
        s = s[:limit].rstrip() + "…"
    return s


def render_analysis_skeleton(modules: list[SourceModule], source_name: str = "") -> str:
    """把原文自动切出来的模块渲染成 analysis.md 骨架（五列表格）。

    自动填的是**原文里已经存在的东西**：章节、模块名、分号/项目符号切出的论据、数字。
    「核心结论」必须是判断句，机器判断不了，所以留 `待补：` —— 但只有一条论据时
    把那一句当结论候选，免得同一句话在「结论」「论据」两列各写一遍。
    """
    rows = []
    for m in modules:
        items = []
        for e in (m.evidence or []):
            e = re.sub(r"^[-*•·\d.、\s]+", "", (e or "").strip())
            if len(e) >= 8:
                items.append(_cell(e, 60))
        evidence = "；".join(items[:4])
        if len(items) <= 1 and m.conclusion:
            conclusion, evidence = _cell(m.conclusion), evidence or "待补：2-4 条论据"
        else:
            conclusion = "待补：一句话判断句"
        gaps = "；".join(_cell(g, 60) for g in (m.gaps or [])) or "无"
        rows.append(f"| {_cell(m.section, 30) or '（未分章）'} | {_cell(m.module, 30)} "
                    f"| {conclusion} | {evidence} | {gaps} |")
    head = _ANALYSIS_HEAD
    if source_name:
        head = head.replace("# 输入分析", f"# 输入分析（{source_name}）", 1)
    return head + "\n".join(rows) + "\n"


def write_analysis_skeleton(path: str, modules: list[SourceModule],
                            source_name: str = "") -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(render_analysis_skeleton(modules, source_name))


# ---------- SourceModule 序列化 ----------
# ---------- 缓存键 ----------

def compute_source_hash(*paths: str) -> str:
    """计算一个或多个文件的 hash，用于判断是否需要重跑分析。"""
    h = hashlib.md5()
    for p in paths:
        if os.path.isfile(p):
            with open(p, "rb") as f:
                h.update(f.read())
    return h.hexdigest()[:12]


