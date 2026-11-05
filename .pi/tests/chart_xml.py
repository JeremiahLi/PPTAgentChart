# -*- coding: utf-8 -*-
"""图表 XML 回归测试：渲染器产出的 .pptx 必须是 PowerPoint 能打开的。

为什么单独立这个测试 —— 出过一次真实事故：
`render_chart_native` 对环形图设了 `data_labels.position`，python-pptx 照写
`<c:dLblPos val="outEnd"/>`；但 ECMA-376 里 `doughnutChart` 不允许数据标签定位。
PowerPoint 打开这份 pptx 时判定「文件损坏 / 无法打开」，点「修复」后图表被丢弃，
**那一页就整页空白**（页眉和右侧洞察还在，中间的图没了）。

原有的 `.pi/tests/smoke.py` 覆盖不到：它的 fixture 里只有一张柱状图。
本测试对 6 种图型逐一渲染并断言 XML 合规，另在装有 PowerPoint 的机器上
真开一次文件做端到端确认。

跑法：`python .pi/tests/chart_xml.py`（纯离线；没有 PowerPoint 时自动跳过第 3 步）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "renderer"))

import layouts            # noqa: E402
from ppt_dsl import ChartSlide  # noqa: E402
from pptx import Presentation   # noqa: E402
from pptx.util import Inches    # noqa: E402

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

CHART_CASES = [
    # (说明, chart spec)
    ("column", {"type": "column", "categories": ["V1", "V2", "V3", "V4"],
                "series": [{"name": "通过率", "values": [62, 71, 79, 88]}], "unit": "%"}),
    ("bar", {"type": "bar", "categories": ["A", "B", "C"],
             "series": [{"name": "耗时", "values": [12, 31, 7]}], "unit": "s"}),
    ("line", {"type": "line", "categories": ["W1", "W2", "W3", "W4"],
              "series": [{"name": "分数", "values": [50, 61, 58, 70]}], "unit": ""}),
    ("pie", {"type": "pie", "categories": ["甲", "乙", "丙"],
             "series": [{"name": "占比", "values": [55, 30, 15]}], "unit": "%"}),
    ("doughnut", {"type": "doughnut", "categories": ["甲", "乙", "丙"],
                  "series": [{"name": "占比", "values": [80, 15, 5]}], "unit": "%"}),
    ("radar", {"type": "radar", "categories": ["可执行", "可控", "可靠", "可恢复"],
               "series": [{"name": "基础", "values": [95, 85, 80, 70]},
                          {"name": "生产", "values": [98, 95, 95, 90]}], "unit": "%"}),
]

_failures: list[str] = []


def _check(name: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  [ ok ] {name}")
    else:
        _failures.append(f"{name}: {detail}")
        print(f"  [FAIL] {name}  {detail}")


def _build(spec: dict, out_path: Path) -> None:
    """用真实渲染路径出一页原生图表 pptx。"""
    model = ChartSlide.model_validate({
        "slide_type": "chart",
        "title": "回归测试图表页标题即结论",
        "summary": "这一段是标题下的核心结论段，用于验证图表页能正常渲染，不参与本次断言。",
        "chart": spec,
        "insights": [{"text": "第 1 条洞察，用于填充右侧面板", "icon": "增长"},
                     {"text": "第 2 条洞察，用于填充右侧面板", "icon": "分析"}],
    })
    layouts.CHART_MODE = "native"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    layouts.RENDERERS["chart"](slide, model)
    prs.save(str(out_path))


def _chart_xml(pptx: Path) -> list[bytes]:
    with zipfile.ZipFile(pptx) as z:
        return [z.read(n) for n in z.namelist()
                if n.startswith("ppt/charts/chart") and n.endswith(".xml")]


def _powerpoint_available() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"):
            return True
    except OSError:
        return False


def _powerpoint_opens(pptx: Path) -> tuple[bool, str]:
    """真让 PowerPoint 开一次文件；打不开说明产物是坏的。"""
    ps = (
        "param([string]$P)\n"
        "$ErrorActionPreference='Stop'\n"
        "$app = New-Object -ComObject PowerPoint.Application\n"
        "try { $pres = $app.Presentations.Open($P, $true, $false, $false) }\n"
        "catch { Write-Output ('OPENFAIL ' + $_.Exception.Message); $app.Quit(); exit 0 }\n"
        "$pres.Close(); $app.Quit(); Write-Output 'OPENOK'\n"
    )
    tmp = Path(tempfile.gettempdir()) / "_chart_xml_check.ps1"
    tmp.write_text(ps, encoding="utf-8")
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(tmp),
         "-P", str(pptx)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
    out = (r.stdout or "") + (r.stderr or "")
    return "OPENOK" in out, out.strip()[:200]


def main() -> int:
    from lxml import etree

    td = Path(tempfile.mkdtemp(prefix="chartxml_"))
    try:
        print("== 1. 6 种图型渲染 + XML 合规 ==")
        pptxs = {}
        for name, spec in CHART_CASES:
            out = td / f"{name}.pptx"
            _build(spec, out)
            pptxs[name] = out
            parts = _chart_xml(out)
            _check(f"{name}: 产出 1 个图表部件", len(parts) == 1, f"实际 {len(parts)} 个")

            bad_dlbl, bad_legend = [], []
            for raw in parts:
                root = etree.fromstring(raw)
                # 环图内不得出现 dLblPos（PowerPoint 会判定文件损坏）
                for dough in root.iter(f"{{{C_NS}}}doughnutChart"):
                    for el in dough.iter(f"{{{C_NS}}}dLblPos"):
                        bad_dlbl.append(etree.tostring(el).decode())
                # legendPos 必须带必填属性 val
                for leg in root.iter(f"{{{C_NS}}}legend"):
                    pos = leg.find(f"{{{C_NS}}}legendPos")
                    if pos is None or pos.get("val") is None:
                        bad_legend.append(etree.tostring(leg).decode()[:80])

            _check(f"{name}: 无非法 dLblPos", not bad_dlbl, "; ".join(bad_dlbl)[:120])
            _check(f"{name}: legendPos 带 val", not bad_legend, "; ".join(bad_legend)[:120])

        print("== 2. 环图数据标签用浅色（落在深色环带上，深色字看不见）==")
        dough_xml = _chart_xml(pptxs["doughnut"])[0].decode("utf-8", "replace")
        _check("doughnut: dLbls 文字色为白色", "FFFFFF" in dough_xml, dough_xml[-500:][:200])

        print("== 3. PowerPoint 真机打开（没装 Office 时跳过）==")
        if _powerpoint_available():
            if shutil.which("powershell"):
                ok, detail = _powerpoint_opens(pptxs["doughnut"])
                _check("PowerPoint 能打开环形图 pptx", ok, detail)
            else:
                print("  [skip] 找不到 powershell")
        else:
            print("  [skip] 本机未注册 PowerPoint.Application")
    finally:
        shutil.rmtree(td, ignore_errors=True)

    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项失败：")
        for f in _failures:
            print(f"  - {f}")
        return 1
    print("[PASS] 图表 XML 无回归")
    return 0


if __name__ == "__main__":
    sys.exit(main())
