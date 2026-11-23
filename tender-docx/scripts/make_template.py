#!/usr/bin/env python3
"""make_template.py — 生成内置基准版式模板 templates/base-template.docx。

改版式 = 改 docx_toolkit.py 里的样式定义（setup_styles）后重跑本脚本。
模板只承载样式与页面设置，正文为空——create/assemble 从它起建。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx_toolkit import enable_footer_pagenum, require_docx, setup_styles  # noqa: E402


def main() -> None:
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(errors="replace")  # 中文 Windows 管道（GBK）下 ⚠️ 会编码崩，降级输出
    # 带参数解析：--help 只看说明不落盘（曾无 argparse，冒烟 --help 也会重建模板，
    # 每次发版都给 git 造一份内容等价的 docx 元数据差异）
    argparse.ArgumentParser(
        description="生成内置基准版式模板 templates/base-template.docx（改版式=改 docx_toolkit setup_styles 后重跑）"
    ).parse_args()
    require_docx()  # --help 已在 parse_args 内退出；顶层 import 会让缺依赖环境起不来
    from docx import Document
    from docx.shared import Cm

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
