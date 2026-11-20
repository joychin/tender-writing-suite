#!/usr/bin/env python3
"""parse_document.py — 招标文件解析：docx/pdf/txt/md → Markdown + 行号大纲 + 元信息。

用法：
    python3 parse_document.py <文件路径> [-o 输出根目录]

输出（默认落 work/parse/<文件名含扩展名>/）：
    <名>.md            全文 Markdown（PDF 每页前有 <!-- p:N --> 页码锚点）
    <名>.outline.json  标题树，节点带 start_line/end_line（1 起，闭区间）
    <名>.meta.json     sha256 / 识别方式 / 规模 / 警示

幂等：三件产物已存在且 meta.sha256 与源文件一致 → [解析跳过]，不重转。
已知边界：.doc 与扫描件（图片型 PDF）不支持；PDF 通道按纯文本逐行提取，
不识别表格框线。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from pathlib import Path

MIN_TEXT_CHARS = 100
BLANK_KEEP = 2

_CN = r"[一二三四五六七八九十百零]+"
_CN_NUM = r"[一二三四五六七八九十百零\d]+"
_PAT_CHAPTER = re.compile(rf"^第\s*{_CN_NUM}\s*[章节篇部分]")
_PAT_CN_LIST = re.compile(rf"^{_CN}\s*、")
_PAT_CN_PAREN = re.compile(rf"^[（(]{_CN}[)）]")
_PAT_DECIMAL = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,3})(?=\s|$)")
_PAT_SENT_PUNCT = re.compile(r"[，。；！？…]")
_MD_HEADING = re.compile(r"^(#{1,6})\s+(\S.*)$")
_PAGE_NUM = re.compile(r"^\s*[-—–]?\s*\d{1,4}\s*[-—–]?\s*$")

CONV_CN = {
    "docx-native": "原生样式提取（结构由文档样式声明，可信）",
    "docx-numbered": "中文编号识别（标题印在原文、可验证，层级可能不完整）",
    "docx-plain": "未识别出章节结构（大纲不可用，定位请用关键词检索）",
    "pdf-numbered": "中文编号识别（标题印在原文、可验证，层级可能不完整）",
    "pdf-plain": "未识别出章节结构（大纲不可用，定位请用关键词检索）",
    "passthrough": "文本透传（无结构信息，定位请用关键词检索）",
}


def die(msg: str, code: int = 1):
    print(f"[解析失败] {msg}")
    sys.exit(code)


def need_deps(pip_names: list[str]):
    print("[解析失败] 缺少依赖：" + "、".join(pip_names))
    print("请先执行：pip install " + " ".join(pip_names))
    sys.exit(2)


class Lines:
    """正文行收集器：空行折叠 + 行号标题登记。"""

    def __init__(self):
        self.lines: list[str] = []
        self.headings: list[tuple[int, str, int]] = []  # (idx0, text, level)
        self.tables = 0
        self.images = 0

    def emit(self, text: str) -> None:
        if text == "" and len(self.lines) >= BLANK_KEEP \
                and self.lines[-1] == "" and self.lines[-2] == "":
            return
        self.lines.append(text)

    def emit_heading(self, text: str, level: int) -> None:
        self.emit(text)
        self.headings.append((len(self.lines) - 1, text, level))


# ---------------------------------------------------------------- docx 通道

def _iter_blocks(doc):
    from docx.document import Document as _Document  # noqa: F401
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _heading_by_style(p) -> int | None:
    try:
        name = (p.style.name or "").strip()
    except Exception:
        name = ""
    m = re.match(r"^(?:heading|标题)\s*(\d)$", name, re.I)
    if m:
        return min(int(m.group(1)), 6)
    try:
        lvl = p._p.xpath("./w:pPr/w:outlineLvl")
        if lvl:
            return min(int(lvl[0].get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")) + 1, 6)
    except Exception:
        pass
    return None


def _is_list_item(p) -> bool:
    try:
        return bool(p._p.xpath("./w:pPr/w:numPr"))
    except Exception:
        return False


def _image_count(p) -> int:
    try:
        return len(p._p.xpath(".//w:drawing") + p._p.xpath(".//w:pict"))
    except Exception:
        return 0


def _table_pipe(tbl) -> list[str]:
    rows = [[" ".join((c.text or "").split())[:200] or " " for c in r.cells]
            for r in tbl.rows]
    if not rows:
        return []
    ncols = max(len(r) for r in rows)

    def fmt(cells: list[str]) -> str:
        cells = (list(cells) + [" "] * ncols)[:ncols]
        return "| " + " | ".join((c or " ").replace("|", "\\|") for c in cells) + " |"

    out = [fmt(rows[0]), "| " + " | ".join(["---"] * ncols) + " |"]
    out += [fmt(r) for r in rows[1:]]
    return out


def parse_docx(src: Path, L: Lines) -> tuple[str, list[str], list[list[int]]]:
    try:
        from docx import Document
    except ImportError:
        need_deps(["python-docx"])
    doc = Document(str(src))
    warnings: list[str] = []
    elements: list[list[int]] = []
    styled = False
    for ordinal, block in enumerate(_iter_blocks(doc)):
        start = len(L.lines) + 1  # 下一个将写入的行（1 基）
        if hasattr(block, "rows"):  # Table
            L.tables += 1
            for line in _table_pipe(block):
                L.emit(line)
            L.emit("")
        else:
            text = (block.text or "").strip()
            imgs = _image_count(block)
            L.images += imgs
            lvl = _heading_by_style(block)
            if lvl and text:
                styled = True
                L.emit_heading(text, lvl)
                L.emit("")
            elif _is_list_item(block) and text:
                L.emit("- " + text)
                L.emit("")
            elif not text and imgs:
                L.emit("![](图片)")
                L.emit("")
            elif text:
                L.emit(text)
                L.emit("")
        end = len(L.lines)
        if end >= start:
            elements.append([ordinal, start, end])
    L.emit("")
    if styled:
        return "docx-native", warnings, elements
    _numbering_pass(L)
    if L.headings:
        return "docx-numbered", warnings, elements
    return ("docx-plain", ["未识别出章节结构（无样式标题、无编号），大纲不可用，定位请用关键词检索"],
            elements)


# ---------------------------------------------------------------- 编号启发式

def _clean_candidate(s: str) -> bool:
    return 0 < len(s) <= 60 and not _PAT_SENT_PUNCT.search(s)


def _numbering_pass(L: Lines) -> None:
    stripped = [l.strip() for l in L.lines]
    has_l1 = any(_PAT_CHAPTER.match(s) and _clean_candidate(s) for s in stripped)
    cands: list[tuple[int, str, int]] = []
    for i, s in enumerate(stripped):
        if not s or s.startswith("<!--"):
            continue
        lvl = _level_by_number(s, has_l1)
        if lvl:
            cands.append((i, s, lvl))
    L.headings = cands


def _level_by_number(s: str, has_l1: bool) -> int | None:
    if not _clean_candidate(s):
        return None
    if _PAT_CHAPTER.match(s):
        return 1
    if _PAT_CN_LIST.match(s):
        return 2
    if _PAT_CN_PAREN.match(s):
        return 3
    m = _PAT_DECIMAL.match(s)
    if m:
        segs = m.group(1).count(".") + 1
        return min((2 if has_l1 else 1) + segs - 1, 6)
    return None


# ---------------------------------------------------------------- pdf 通道

def parse_pdf(src: Path, L: Lines) -> tuple[str, list[str]]:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        need_deps(["pypdfium2"])
    pdf = pdfium.PdfDocument(str(src))
    pages: list[list[str]] = []
    total_chars = 0
    for i in range(len(pdf)):
        tp = pdf[i].get_textpage()
        txt = (tp.get_text_range() or "").replace("\x0c", "")
        txt = txt.replace("\r\n", "\n").replace("\r", "\n")
        total_chars += len(txt.strip())
        lines = [l.strip() for l in txt.split("\n")]
        while lines and lines[0] == "":
            lines.pop(0)
        while lines and lines[-1] == "":
            lines.pop()
        pages.append(lines)
    if total_chars < MIN_TEXT_CHARS:
        die(f"转换结果过短（{total_chars} 字符）。若输入为 PDF，疑似扫描件（图片型 PDF）："
            f"请提供文本型 PDF 或 .docx（扫描件云端识别不在本技能范围内）")
    pages, hf_warnings = _strip_headers_footers(pages)
    for i, p in enumerate(pages):
        L.emit(f"<!-- p:{i + 1} -->")
        for line in p:
            L.emit(line)
        L.emit("")
    L.emit("")
    warnings = ["PDF 通道按纯文本逐行提取（不识别表格框线），"
                "评分表/前附表等请按文本行逐行核对"] + hf_warnings
    _numbering_pass(L)
    conv = "pdf-numbered" if L.headings else "pdf-plain"
    if conv == "pdf-plain":
        warnings.append("未识别出章节结构（无书签/无编号），大纲不可用，定位请用关键词检索")
    meta_pages = len(pdf)
    return conv, warnings, meta_pages


def _strip_headers_footers(pages: list[list[str]]) -> tuple[list[list[str]], list[str]]:
    from collections import Counter

    n = len(pages)
    if n < 3:
        return pages, []
    threshold = max(3, int(n * 0.25))
    kill: set[str] = set()
    for getter in (lambda p: p[0] if p else None, lambda p: p[-1] if p else None):
        cnt = Counter(s for p in pages if (s := getter(p)) is not None)
        kill |= {s for s, c in cnt.items() if c >= threshold and 0 < len(s) <= 60}
    stripped = []
    for p in pages:
        q = list(p)
        while q and (q[0] in kill or _PAGE_NUM.match(q[0])):
            q.pop(0)
        while q and (q[-1] in kill or _PAGE_NUM.match(q[-1])):
            q.pop()
        stripped.append(q)
    warnings = [f"已剔除跨页重复的页眉/页脚 {len(kill)} 行"] if kill else []
    return stripped, warnings


# ---------------------------------------------------------------- 透传通道

def parse_passthrough(src: Path, L: Lines) -> tuple[str, list[str]]:
    raw = src.read_bytes()
    text = None
    for enc in ("utf-8", "gb18030", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    if text is None:
        die("无法识别文件编码（已尝试 UTF-8 / GB18030 / UTF-16）")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    for line in text.split("\n"):
        L.emit(line.rstrip())
    is_md = src.suffix.lower() == ".md"
    for i, line in enumerate(L.lines):
        m = _MD_HEADING.match(line.strip())
        if m:
            L.headings.append((i, m.group(2).strip(), len(m.group(1))))
    warnings = []
    if is_md and not L.headings:
        warnings.append("Markdown 内无 # 标题，大纲为空，定位请用关键词检索")
    if not is_md:
        warnings.append("纯文本无结构信息，大纲为空，定位请用关键词检索")
    return "passthrough", warnings


# ---------------------------------------------------------------- 大纲与落盘

def build_outline(headings: list[tuple[int, str, int]], total_lines: int) -> list[dict]:
    nodes: list[dict] = []
    stack: list[dict] = []
    for idx0, text, level in headings:
        while stack and stack[-1]["level"] >= level:
            stack.pop()["end_line"] = idx0  # 1 基：新标题行的前一行
        node = {"标题": text, "level": level, "start_line": idx0 + 1,
                "end_line": total_lines, "children": []}
        (stack[-1]["children"] if stack else nodes).append(node)
        stack.append(node)
    return nodes


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def artifacts_digest(paths: list[Path]) -> str | None:
    """对产物文件（名+字节）串接计算摘要——产物侧完整性登记。"""
    h = hashlib.sha256()
    got = False
    for p in paths:
        if not p.is_file():
            return None
        h.update(p.name.encode("utf-8"))
        h.update(p.read_bytes())
        got = True
    return h.hexdigest() if got else None


def _outline_max_end(nodes: list) -> int:
    mx = 0
    for n in nodes:
        mx = max(mx, int(n.get("end_line", 0) or 0))
        mx = max(mx, _outline_max_end(n.get("children", []) or []))
    return mx


def artifacts_consistent(out_root: Path, src: Path, old: dict) -> tuple[bool, str]:
    """幂等跳过前的产物对账：文件齐全 + 产物哈希与登记一致 + 行号不越界。"""
    stem = src.stem
    md = out_root / f"{stem}.md"
    ol = out_root / f"{stem}.outline.json"
    el = out_root / f"{stem}.elements.json"
    files = [md, ol] + ([el] if src.suffix.lower() == ".docx" else [])
    if not all(p.is_file() for p in files):
        return False, "解析产物缺失"
    digest = artifacts_digest(files)
    if digest is None or old.get("artifacts_sha256") != digest:
        return False, "产物内容与登记不一致（可能被外部编辑）"
    n_lines = len(md.read_text(encoding="utf-8", errors="replace").splitlines())
    try:
        if _outline_max_end(json.loads(ol.read_text(encoding="utf-8"))) > n_lines:
            return False, f"outline 行号越界（> md 实际 {n_lines} 行）"
    except Exception:
        return False, "outline.json 损坏"
    if el.is_file():
        try:
            els = json.loads(el.read_text(encoding="utf-8"))
            if els and max(int(e[2]) for e in els) > n_lines:
                return False, f"elements 行号越界（> md 实际 {n_lines} 行）"
        except Exception:
            return False, "elements.json 损坏"
    return True, ""


def print_summary(tag: str, name: str, out_root: Path, meta: dict) -> None:
    parts = [f"约 {meta['chars']:,} 字符"]
    if meta.get("pages"):
        parts.append(f"{meta['pages']} 页")
    parts += [f"一级标题 {len(meta['top_level_titles'])} 个", f"表格 {meta['tables']} 张"]
    if meta.get("images"):
        parts.append(f"图片 {meta['images']} 处")
    print(f"[{tag}] {name} → {out_root.as_posix()}/")
    print("  规模：" + " ｜ ".join(parts))
    print(f"  识别方式：{meta['conversion']} —— {CONV_CN.get(meta['conversion'], '')}")
    for w in meta["warnings"]:
        print(f"  ⚠️ {w}")


def main() -> None:
    ap = argparse.ArgumentParser(description="招标文件解析（docx/pdf/txt/md → md+outline+meta）")
    ap.add_argument("input", help="源文件路径")
    ap.add_argument("-o", "--out-dir", default="work/parse", help="输出根目录（默认 work/parse）")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.is_file():
        die(f"文件不存在：{src}")
    suffix = src.suffix.lower()
    if suffix not in {".docx", ".pdf", ".txt", ".md"}:
        die(f"暂不支持该格式：{suffix}（支持 .docx / .pdf / .txt / .md；"
            f".doc 请用 Word 另存为 .docx；扫描件请提供文本型 PDF）")

    out_root = Path(args.out_dir) / src.name
    md_path = out_root / (src.stem + ".md")
    ol_path = out_root / (src.stem + ".outline.json")
    meta_path = out_root / (src.stem + ".meta.json")

    sha = sha256_of(src)
    if meta_path.is_file():
        try:
            old = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            old = None  # meta 损坏视为无，走全量重析
        if old is not None and old.get("sha256") == sha:
            ok, reason = artifacts_consistent(out_root, src, old)
            if ok:
                print_summary("解析跳过", src.name, out_root, old)
                print("  内容未变化（sha256 一致），沿用已有产物")
                return
            print(f"[解析重建] {reason}——源文件未变，重新解析并恢复产物")

    L = Lines()
    pages = None
    elements: list[list[int]] | None = None
    try:
        if suffix == ".docx":
            conv, warnings, elements = parse_docx(src, L)
        elif suffix == ".pdf":
            conv, warnings, pages = parse_pdf(src, L)
        else:
            conv, warnings = parse_passthrough(src, L)
    except SystemExit:
        raise
    except Exception as exc:
        die(f"解析器内部错误（{type(exc).__name__}: {exc}）。"
            f"可尝试换用另一种格式（如 .docx / 文本型 PDF）后重试")

    while L.lines and L.lines[-1] == "":
        L.lines.pop()
    L.lines.append("")  # 文件以单个换行结尾
    total = len(L.lines) - 1  # splitlines 口径的实际行数（outline 端点与之一致）
    outline = build_outline(L.headings, total)
    top_titles = [n["标题"] for n in outline][:12]
    out_root.mkdir(parents=True, exist_ok=True)
    md_path.write_text("\n".join(L.lines), encoding="utf-8")
    ol_path.write_text(json.dumps(outline, ensure_ascii=False, indent=1), encoding="utf-8")
    if elements is not None:
        elements = [[o, s_, min(e_, total)] for o, s_, e_ in elements]
        (out_root / (src.stem + ".elements.json")).write_text(
            json.dumps(elements, ensure_ascii=False), encoding="utf-8")

    meta = {
        "source": src.name,
        "sha256": sha,
        "artifacts_sha256": artifacts_digest(
            [md_path, ol_path] + ([out_root / (src.stem + ".elements.json")] if elements is not None else [])
        ) or "",
        "bytes": src.stat().st_size,
        "conversion": conv,
        "chars": sum(len(l) for l in L.lines),
        "pages": pages,
        "headings": len(L.headings),
        "top_level_titles": top_titles,
        "tables": L.tables,
        "images": L.images,
        "warnings": warnings,
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    print_summary("解析成功", src.name, out_root, meta)


if __name__ == "__main__":
    main()
