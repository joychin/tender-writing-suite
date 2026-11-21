#!/usr/bin/env python3
"""make_template.py — 生成内置基准版式模板 templates/base-template.docx。

改版式 = 改 docx_toolkit.py 里的样式定义（setup_styles）后重跑本脚本。
模板只承载样式与页面设置，正文为空——create/assemble 从它起建。
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx_toolkit import enable_footer_pagenum, setup_styles  # noqa: E402

from docx import Document  # noqa: E402
from docx.shared import Cm  # noqa: E402


def main() -> None:
    doc = Document()
    setup_styles(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.18)
    enable_footer_pagenum(doc, different_first=True)
    out = Path(__file__).resolve().parent.parent / "templates" / "base-template.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"[已生成] {out}（内置基准版式：宋体小四正文/黑体标题族/封面样式/页脚页码）")


if __name__ == "__main__":
    main()
