# -*- coding: utf-8 -*-
"""Word 输入预处理：去修订 + 抽图（含 EMF/WMF 原样保留）→ 给 Agent 读的 markdown。

用法：
    python .pi/renderer/docx_in.py <输入.docx> -o .pi/work/input.md [--media .pi/work/media]

产出三样东西：
  1. `-o` 指定的 markdown：**修订记录已经处理干净**的正文（标题 / 段落 / 列表 / 表格），
     图片按原文位置留下 `![图N](media/NN-file.emf)` 标记；
  2. `--media` 目录：原文里的图片**逐字节原样**落盘（EMF/WMF 不转栅格、不重绘）；
  3. 同名的 `*.manifest.json`：图片清单（序号 / 文件 / 格式 / 物理尺寸 / 所在段落），
     给 Agent 排 `figure` 版式时用。

修订记录的口径（与 Word 的「接受所有修订」一致，但**批注不呈现**）：
  - `w:ins`（插入）→ **接受**，内容保留；
  - `w:del`（删除）→ **丢弃**，`w:delText` 一律不输出；
  - `w:moveFrom` / `w:moveTo`（移动）→ 前者丢弃、后者接受；
  - 各类 `*Change`（格式修订）、`w:commentRangeStart/End`、`w:commentReference` → 丢弃；
  - `w:numPr`（列表）、`w:pStyle`（标题）→ 保留成 markdown 的标题/列表结构。

**零新依赖**：.docx 就是 zip，本模块只用 stdlib 的 zipfile + xml.etree。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import struct
import sys
import zipfile
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
VML = "urn:schemas-microsoft-com:vml"

# 图片扩展名：**矢量图必须原样嵌入**（EMF/WMF 是 PowerPoint 原生格式，转了必糊）
VECTOR_EXT = {"emf", "wmf"}
RASTER_EXT = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"}

# 丢弃整棵子树（删除的正文 / 批注 / 格式修订记录）
_DROP = {
    "del", "delText", "moveFrom", "moveFromRangeStart", "moveFromRangeEnd",
    "commentRangeStart", "commentRangeEnd", "commentReference", "annotationRef",
    "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange",
    "sectPrChange", "tblGridChange", "numberingChange", "cellDel", "cellMerge",
    "customXmlDelRangeStart", "customXmlDelRangeEnd",
}
# 拆掉外壳、保留内容（接受的修订 / 容器）
_UNWRAP = {
    "ins", "moveTo", "moveToRangeStart", "moveToRangeEnd", "smartTag", "sdt",
    "sdtContent", "hyperlink", "bdo", "dir", "fldSimple", "subDoc", "altChunk",
    "customXmlInsRangeStart", "customXmlInsRangeEnd", "cellIns",
    "proofErr", "bookmarkStart", "bookmarkEnd", "lastRenderedPageBreak",
}


def _tag(el) -> str:
    t = el.tag
    return t.rsplit("}", 1)[-1] if isinstance(t, str) and "}" in t else str(t)


def _val(el) -> str | None:
    """w:* 元素的 w:val 属性。"""
    return el.attrib.get(f"{{{W}}}val")


# ---------------- 图片：识别 + 物理尺寸 ----------------

def image_size_in(blob: bytes, ext: str) -> tuple[float | None, float | None]:
    """图片的物理尺寸（英寸），取不到返回 (None, None)。

    - EMF：头里 `rclFrame`（偏移 24，4×int32，单位 0.01mm）就是"设备无关"的物理尺寸；
    - WMF：placeable 头的 bbox（偏移 6，4×int16，单位 twip = 1/1440 in）；
    - 栅格图：交给 Pillow（python-pptx 的依赖，一定在）。
    """
    ext = ext.lower()
    try:
        if ext == "emf" and len(blob) >= 48:
            _l, _t, _r, _b = struct.unpack_from("<4i", blob, 24)
            w = (_r - _l) / 100.0 / 25.4          # 0.01mm → inch
            h = (_b - _t) / 100.0 / 25.4
            if 0.05 < w < 60 and 0.05 < h < 60:
                return round(w, 3), round(h, 3)
        elif ext == "wmf" and len(blob) >= 22 and blob[:4] == b"\xd7\xcd\xc6\x9a":
            _l, _t, _r, _b = struct.unpack_from("<4h", blob, 6)
            w, h = (_r - _l) / 1440.0, (_b - _t) / 1440.0
            if 0.05 < w < 60 and 0.05 < h < 60:
                return round(w, 3), round(h, 3)
        else:
            import io

            from PIL import Image as PILImage
            with PILImage.open(io.BytesIO(blob)) as im:
                dpi = im.info.get("dpi") or (96, 96)
                return round(im.width / float(dpi[0]), 3), round(im.height / float(dpi[1]), 3)
    except Exception:                                   # noqa: BLE001 —— 尺寸拿不到不影响主流程
        pass
    return None, None


def _find_blips(root) -> list[str]:
    """一段内容里所有图片的关系 id（`a:blip@r:embed` + VML `v:imagedata@r:id`）。"""
    out: list[str] = []
    for el in root.iter():
        if _tag(el) == "blip":
            rid = el.attrib.get(f"{{{R}}}embed") or el.attrib.get(f"{{{R}}}link")
            if rid:
                out.append(rid)
        elif _tag(el) == "imagedata":
            rid = el.attrib.get(f"{{{R}}}id")
            if rid:
                out.append(rid)
    return out


# ---------------- 正文遍历（去修订） ----------------

class DocxReader:
    def __init__(self, path: str):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.rels = self._load_rels()
        self.media: dict[str, bytes] = {
            n: self.zf.read(n) for n in self.zf.namelist()
            if n.startswith("word/media/") and not n.endswith("/")
        }
        self.stats = {"ins": 0, "del": 0, "comments": 0, "moves": 0, "changes": 0}
        self.images: list[dict] = []          # 清单（按出现顺序）
        self._img_no = 0

    # ---- 关系表：rId -> word/media/xxx ----
    def _load_rels(self) -> dict[str, str]:
        name = "word/_rels/document.xml.rels"
        if name not in self.zf.namelist():
            return {}
        root = ET.fromstring(self.zf.read(name))
        out = {}
        for rel in root:
            rid, target = rel.attrib.get("Id"), rel.attrib.get("Target")
            if not rid or not target:
                continue
            out[rid] = target.lstrip("/") if target.startswith("/") else "word/" + target
        return out

    # ---- 修订标记的统计与丢弃 ----
    def _note_drop(self, tag: str) -> bool:
        if tag == "del":
            self.stats["del"] += 1
            return True
        if tag == "moveFrom":
            self.stats["moves"] += 1
            return True
        if tag.endswith("Change"):
            self.stats["changes"] += 1
            return True
        if tag.startswith("comment") or tag in ("annotationRef",):
            self.stats["comments"] += 1
            return True
        return False

    def _emit_images(self, el, para_no: int, sink: list[str]) -> None:
        """把这一段里的图片按出现顺序登记进清单，并往正文里插标记。"""
        for rid in _find_blips(el):
            src = self.rels.get(rid, "")
            base = os.path.basename(src)
            blob = self.media.get(src)
            if blob is None:
                continue
            ext = (base.rsplit(".", 1)[-1] if "." in base else "bin").lower()
            self._img_no += 1
            no = self._img_no
            fname = f"{no:02d}-{base}"
            w_in, h_in = image_size_in(blob, ext)
            self.images.append({
                "no": no, "file": fname, "source": src, "ext": ext,
                "vector": ext in VECTOR_EXT, "bytes": len(blob),
                "width_in": w_in, "height_in": h_in, "paragraph": para_no,
            })
            sink.append(f"![图{no}](media/{fname})")

    def _text_of(self, el, para_no: int, sink: list[str]) -> None:
        """递归收集一段文字（含去修订规则）。

        默认分支就是 `_UNWRAP` 的语义（拆壳取内容）；只有 `_DROP` 里的标签整棵丢弃。
        """
        tag = _tag(el)
        if self._note_drop(tag) or tag in _DROP:
            return
        if tag == "ins":
            self.stats["ins"] += 1
        if tag in ("t", "tab", "br", "cr", "noBreakHyphen"):
            if tag == "t":
                sink.append(el.text or "")
            elif tag == "tab":
                sink.append(" ")
            else:
                sink.append(" ")
            return
        if tag in ("drawing", "pict", "object"):
            self._emit_images(el, para_no, sink)
            return
        for child in el:
            self._text_of(child, para_no, sink)

    def _para(self, p, para_no: int) -> tuple[str, int | None, bool]:
        """返回 (文本, 大纲级别, 是否列表项)。"""
        ppr = next((c for c in p if _tag(c) == "pPr"), None)
        level = None
        is_list = False
        if ppr is not None:
            style = next((c for c in ppr if _tag(c) == "pStyle"), None)
            if style is not None:
                sv = (_val(style) or "").lower()
                m = re.search(r"(?:heading|标题)\s*([1-9])", sv)
                if m:
                    level = int(m.group(1))
            outline = next((c for c in ppr if _tag(c) == "outlineLvl"), None)
            if outline is not None:
                try:
                    level = min(level or 9, int(_val(outline) or 9) + 1)
                except ValueError:
                    pass
            is_list = any(_tag(c) == "numPr" for c in ppr)
        sink: list[str] = []
        self._text_of(p, para_no, sink)
        text = "".join(sink).strip()
        if text and "![图" in text:
            is_list = False
        return text, level, is_list

    def _table(self, tbl, para_no: int) -> list[list[str]]:
        out: list[list[str]] = []
        for tr in tbl:
            if _tag(tr) != "tr":
                continue
            trpr = next((c for c in tr if _tag(c) == "trPr"), None)
            if trpr is not None and any(_tag(c) == "del" for c in trpr):
                self.stats["del"] += 1                      # 整行被删（修订）
                continue
            row: list[str] = []
            for tc in (c for c in tr if _tag(c) == "tc"):
                tcpr = next((c for c in tc if _tag(c) == "tcPr"), None)
                if tcpr is not None:
                    if any(_tag(c) == "cellDel" and (_val(c) or "").lower() in ("1", "true")
                           for c in tcpr):
                        row.append("")
                        continue
                cells: list[str] = []
                for p in (c for c in tc if _tag(c) == "p"):
                    t, _lvl, _ls = self._para(p, para_no)
                    if t:
                        cells.append(t)
                row.append(" ".join(cells).replace("|", "\\|"))
            if row:
                out.append(row)
        return out

    # ---- 主流程 ----
    def read(self) -> tuple[str, list[str]]:
        body = ET.fromstring(self.zf.read("word/document.xml")).find(f"{{{W}}}body")
        lines: list[str] = []
        para_no = 0
        for el in (body if body is not None else []):
            tag = _tag(el)
            if tag == "p":
                para_no += 1
                text, level, is_list = self._para(el, para_no)
                if not text:
                    continue
                if level:
                    lines.append(f"{'#' * min(6, level + 1)} {text}")
                elif is_list:
                    lines.append(f"- {text}")
                else:
                    lines.append(text)
                lines.append("")
            elif tag == "tbl":
                rows = self._table(el, para_no)
                if rows:
                    ncol = max(len(r) for r in rows)
                    rows = [r + [""] * (ncol - len(r)) for r in rows]
                    lines.append("| " + " | ".join(rows[0]) + " |")
                    lines.append("|" + "---|" * ncol)
                    for r in rows[1:]:
                        lines.append("| " + " | ".join(r) + " |")
                    lines.append("")
        return "\n".join(lines).strip() + "\n", [i["file"] for i in self.images]

    def close(self) -> None:
        self.zf.close()


def convert(src: str, out_md: str, media_dir: str) -> dict:
    """把 .docx 转成 markdown + 图片目录 + 清单，返回汇总信息。"""
    reader = DocxReader(src)
    try:
        md, files = reader.read()
        os.makedirs(media_dir, exist_ok=True)
        for f in files:
            src_name = next((i["source"] for i in reader.images if i["file"] == f), None)
            if src_name and src_name in reader.media:
                with open(os.path.join(media_dir, f), "wb") as fh:
                    fh.write(reader.media[src_name])

        head = ["<!-- 由 .pi/renderer/docx_in.py 生成：修订记录已按“接受插入 / 丢弃删除”处理，",
                "     批注不呈现；图片为**逐字节原样**（EMF/WMF 未转格式）。 -->", ""]
        vec = [i for i in reader.images if i["vector"]]
        if vec:
            head.append("> **原样嵌入的矢量图**（请用 `figure` 版式的 `image` 字段放上页，"
                        "不要转 PNG、不要重绘）：")
            for i in vec:
                size = f"{i['width_in']}×{i['height_in']}in" if i["width_in"] else "尺寸未知"
                head.append(f"> - 图{i['no']} → `media/{i['file']}`（{i['ext'].upper()}，"
                            f"{size}，原文第 {i['paragraph']} 段）")
            head.append("")

        with open(out_md, "w", encoding="utf-8") as fh:
            fh.write("\n".join(head) + "\n" + md)

        manifest = {
            "source": os.path.abspath(src),
            "markdown": os.path.abspath(out_md),
            "media_dir": os.path.abspath(media_dir),
            "revisions": reader.stats,
            "images": reader.images,
        }
        with open(out_md + ".manifest.json", "w", encoding="utf-8") as fh:
            json.dump(manifest, fh, ensure_ascii=False, indent=2)
        return manifest
    finally:
        reader.close()


def _report(m: dict) -> str:
    r, imgs = m["revisions"], m["images"]
    vec = sum(1 for i in imgs if i["vector"])
    ras = len(imgs) - vec
    return (
        f"OK: {m['markdown']}\n"
        f"  修订处理：接受插入 {r['ins']} 处（本轮按“接受所有修订”口径，"
        f"插入内容保留）/ 丢弃删除 {r['del']} 处 / 丢弃批注 {r['comments']} 处 / "
        f"丢弃格式修订 {r['changes']} 处 / 丢弃移动 {r['moves']} 处\n"
        f"  图片：{len(imgs)} 张（其中矢量 EMF/WMF {vec} 张 → 原样嵌入，"
        f"栅格 {ras} 张）→ {m['media_dir']}\n"
        f"  清单：{m['markdown']}.manifest.json"
    )


def main() -> int:
    ap = argparse.ArgumentParser(description="Word(.docx) → markdown + 原样图片")
    ap.add_argument("docx", help="输入 .docx 路径")
    ap.add_argument("-o", "--out", required=True, help="输出 markdown 路径，如 .pi/work/input.md")
    ap.add_argument("--media", default=None,
                    help="图片输出目录，默认与 -o 同级的 media/ 目录")
    args = ap.parse_args()
    if not os.path.isfile(args.docx):
        print(f"读不到输入文件：{args.docx}", file=sys.stderr)
        return 2
    media = args.media or os.path.join(os.path.dirname(os.path.abspath(args.out)), "media")
    try:
        m = convert(args.docx, args.out, media)
    except zipfile.BadZipFile:
        print(f"{args.docx} 不是 .docx（zip）文件：如果是 .doc 老格式，"
              f"请先用 Word 另存为 .docx", file=sys.stderr)
        return 1
    print(_report(m))
    return 0


if __name__ == "__main__":
    sys.exit(main())
