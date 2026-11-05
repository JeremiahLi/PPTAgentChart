# -*- coding: utf-8 -*-
"""Word 输入预处理回归：`.pi/renderer/docx_in.py`。

锁定四件用户明确要求 / 踩过的事：
  1. **修订记录不呈现**：`w:ins`（插入）按"接受所有修订"保留，`w:del`、批注一律丢弃；
  2. **EMF/WMF 原样落盘**：逐字节相同、不转栅格、不重绘；
  3. 标题 / 列表 / 表格 结构正确落到 markdown；
  4. 图片清单给出序号、格式、物理尺寸（EMF 从 `rclFrame` 读，单位 0.01mm）。

不需要 Word / 不需要 Office：fixture 是脚本现拼的最小 .docx（zip）。
跑法：`python .pi/tests/docx_in.py`
"""
from __future__ import annotations

import json
import struct
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "renderer"))

import docx_in  # noqa: E402

_failures: list[str] = []


def _check(name: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  [ ok ] {name}")
    else:
        _failures.append(f"{name}: {detail}")
        print(f"  [FAIL] {name}  {detail}")


NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

EMF_BYTES = b"".join([
    struct.pack("<i", 1),                      # iType = EMR_HEADER
    struct.pack("<i", 88),                     # nSize
    struct.pack("<4i", 0, 0, 240, 120),        # rclBounds（设备单位）
    struct.pack("<4i", 0, 0, 2540, 1270),      # rclFrame（0.01mm → 1.0in × 0.5in）
    struct.pack("<I", 0x464D4520),             # dSignature " EMF"
    b"\x00" * 24,                              # 其余头字段（本测试不解析）
])

_DOC = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NS}><w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 上下文工程</w:t></w:r></w:p>
  <w:p>
    <w:r><w:t>保留的正文。</w:t></w:r>
    <w:ins w:id="1" w:author="张三" w:date="2026-09-01T00:00:00Z">
      <w:r><w:t>【这是插入的句子，应当保留】</w:t></w:r>
    </w:ins>
    <w:del w:id="2" w:author="张三" w:date="2026-09-01T00:00:00Z">
      <w:r><w:delText>【这是被删除的句子，不应出现】</w:delText></w:r>
    </w:del>
    <w:commentRangeStart w:id="0"/><w:r><w:t>批注锚点文字</w:t></w:r>
    <w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>
  </w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
    <w:r><w:t>列表第一项</w:t></w:r></w:p>
  <w:p><w:r><w:drawing><wp:inline>
    <a:graphic><a:graphicData><pic:pic><pic:blipFill>
      <a:blip r:embed="rId9"/>
    </pic:blipFill></pic:pic></a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r></w:p>
  <w:tbl>
    <w:tr><w:tc><w:p><w:r><w:t>策略</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>节省</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>观测屏蔽</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>20-30%</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:trPr><w:del w:id="3" w:author="张三"/></w:trPr>
          <w:tc><w:p><w:r><w:t>被删的整行</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>不该出现</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
</w:body></w:document>
"""

_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId9"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/diagram1.emf"/>
</Relationships>
"""

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
"""


def _make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("word/document.xml", _DOC)
        z.writestr("word/_rels/document.xml.rels", _RELS)
        z.writestr("word/media/diagram1.emf", EMF_BYTES)


def main() -> int:
    td = Path(tempfile.mkdtemp(prefix="docxin_"))
    docx = td / "样例.docx"
    md = td / "input.md"
    media = td / "media"
    _make_docx(docx)

    print("== 1. 转换：修订记录按“接受插入 / 丢弃删除与批注”处理 ==")
    m = docx_in.convert(str(docx), str(md), str(media))
    text = md.read_text(encoding="utf-8")

    _check("插入的内容保留", "【这是插入的句子，应当保留】" in text)
    _check("删除的内容丢弃", "【这是被删除的句子，不应出现】" not in text
           and "delText" not in text)
    _check("批注引用不呈现", "commentReference" not in text and "commentRange" not in text)
    _check("批注锚点正文仍保留（批注本身不呈现）", "批注锚点文字" in text)
    _check("被删除的整行不出现", "被删的整行" not in text and "不该出现" not in text)
    _check("修订统计：插入 1 / 删除 2（一段 + 一整行）/ 批注标记 3（起·止·引用）",
           m["revisions"]["ins"] == 1 and m["revisions"]["del"] == 2
           and m["revisions"]["comments"] == 3, json.dumps(m["revisions"]))

    print("== 2. 结构：标题 / 列表 / 表格 落成 markdown ==")
    _check("一级标题", "## 第一章 上下文工程" in text, text[:200])
    _check("列表项", "- 列表第一项" in text)
    _check("表格两行（被删的整行不计）",
           "| 策略 | 节省 |" in text and "| 观测屏蔽 | 20-30% |" in text)
    _check("表格分隔行", "|---|---|" in text)

    print("== 3. 图片：EMF 原样落盘 + 清单给出物理尺寸 ==")
    out_emf = media / "01-diagram1.emf"
    _check("图片被抽到 media/ 且逐字节相同",
           out_emf.is_file() and out_emf.read_bytes() == EMF_BYTES,
           f"{out_emf}")
    _check("markdown 里有图位标记", "![图1](media/01-diagram1.emf)" in text)
    img = m["images"][0]
    _check("清单标记为矢量（原样嵌入）", img["vector"] is True and img["ext"] == "emf",
           json.dumps(img, ensure_ascii=False))
    _check("EMF 物理尺寸取自 rclFrame（1.0 × 0.5 in）",
           img["width_in"] == 1.0 and img["height_in"] == 0.5,
           f"{img['width_in']} × {img['height_in']}")
    _check("清单文件已写出", (td / "input.md.manifest.json").is_file())
    _check("矢量图在 markdown 顶部被点名“不要转格式”",
           "原样嵌入的矢量图" in text and "不要转 PNG" in text)

    print("== 4. 非 docx 输入给明确报错（不静默产出空文件）==")
    bad = td / "not_a_docx.docx"
    bad.write_bytes(b"plain text, not a zip")
    try:
        docx_in.convert(str(bad), str(td / "x.md"), str(td / "xmedia"))
        _check("非 zip 输入应当抛 BadZipFile", False, "没有抛错")
    except zipfile.BadZipFile:
        _check("非 zip 输入抛 BadZipFile（CLI 会转成人话提示）", True)

    print()
    if _failures:
        print(f"[FAIL] {len(_failures)} 项失败：")
        for f in _failures:
            print(f"  - {f}")
        return 1
    print("[PASS] Word 输入预处理无回归")
    return 0


if __name__ == "__main__":
    sys.exit(main())
