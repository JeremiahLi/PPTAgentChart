# -*- coding: utf-8 -*-
"""生成 `examples-all.json`：19 个版式范例的合集，**唯一的源是 examples/*.json**。

    python .pi/tools/gen_examples.py           # 重新生成
    python .pi/tools/gen_examples.py --check    # 只比对，不写文件（CI/回归用）

为什么要有这个脚本（2026-11-25）：
合集与拆分文件以前是同一套内容存了两份，谁改了拆分文件、忘了同步合集，
模型读到的范例就和实际契约不一致了——而且是**静默**不一致，直到某页渲染出错才发现。
现在合并动作由脚本执行，人手只维护 `examples/NN-*.json`，合集是产物。

排序按文件名数字前缀（`01-` … `18-`、`99-end`），所以 `99-end.json` 天然排最后。
合集的结构是完整 Deck（title/subtitle/style/brief/slides），与 build 的输入同构，
模型开工时读一次就能看到全部版式与字段写法。
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PI_ROOT = Path(__file__).resolve().parents[1]
EX_DIR = PI_ROOT / "skills" / "ppt-common" / "examples"
OUT = PI_ROOT / "skills" / "ppt-common" / "examples-all.json"

DECK_HEAD = {
    "title": "版式范例合集（19 版式）",
    "subtitle": None,
    "style": "style-harness",
    "brief": None,
}


def _order_key(p: Path) -> tuple[int, str]:
    m = re.match(r"(\d+)", p.name)
    return (int(m.group(1)) if m else 10**6, p.name)


def build() -> dict:
    files = sorted((p for p in EX_DIR.glob("*.json")), key=_order_key)
    if not files:
        raise SystemExit(f"找不到范例文件：{EX_DIR}")
    slides = []
    for f in files:
        try:
            slides.append(json.loads(f.read_text(encoding="utf-8")))
        except json.JSONDecodeError as e:
            raise SystemExit(f"{f.name} 不是合法 JSON：{e}") from e
    deck = dict(DECK_HEAD)
    deck["slides"] = slides
    return deck


def main(argv: list[str]) -> int:
    deck = build()
    text = json.dumps(deck, ensure_ascii=False, indent=2) + "\n"
    n = len(deck["slides"])

    if "--check" in argv:
        cur = OUT.read_text(encoding="utf-8") if OUT.exists() else ""
        # 比对语义而非字节：缩进/空行差异不算漂移，字段差异才算
        same = False
        try:
            same = json.loads(cur) == deck
        except json.JSONDecodeError:
            same = False
        if same:
            print(f"OK: examples-all.json 与 {n} 个拆分文件一致")
            return 0
        print(f"[FAIL] examples-all.json 与 examples/ 不一致（共 {n} 个拆分文件）。"
              f"跑 python .pi/tools/gen_examples.py 重新生成。", file=sys.stderr)
        return 1

    OUT.write_text(text, encoding="utf-8")
    print(f"已写入 {OUT.relative_to(PI_ROOT.parent)}：{n} 页，"
          f"{len(text.encode('utf-8'))} 字节")
    types = [s.get("slide_type") for s in deck["slides"]]
    print("版式：" + " / ".join(types))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
