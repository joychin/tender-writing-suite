#!/usr/bin/env python3
"""parse_document.py — 招标文件解析：docx/pdf/txt/md → Markdown + 行号大纲 + 元信息。

用法：
    python3 parse_document.py <文件路径> [-o 输出根目录] [--page-cap N]

输出（默认落 work/parse/<文件名含扩展名>/）：
    <名>.md            全文 Markdown（PDF 每页前有 <!-- p:N --> 页码锚点）
    <名>.outline.json  标题树，节点带 start_line/end_line（1 起，闭区间）
    <名>.meta.json     sha256 / 识别方式 / 规模 / 警示

扫描件（图片型 PDF）两段式：首跑栅格化各页到 pages/ 并生成 <名>.transcript.md
转写模板 → [解析转写待办]；按规程逐页视觉誊录后重跑同一命令 → 归一化落盘
三件套 → [解析成功]（conversion=pdf-scanned-transcript）。页数超上限先停轮，
确认成本后用 --page-cap 提升。
幂等：产物已存在且 meta.sha256 与源文件一致 → [解析跳过]，不重转。
已知边界：PDF 通道按纯文本逐行提取，不识别表格框线（关键表格须页图目视核对）。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

MIN_TEXT_CHARS = 100
BLANK_KEEP = 2
SCAN_PAGE_CAP = 40

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
    "pdf-scanned": "扫描件已栅格化为页图，待视觉转写后重跑本命令完成解析",
    "pdf-scanned-transcript": "扫描件视觉转写（模型逐页读图誊录；数字/专名须对照 pages/ 页图复核）",
    "passthrough": "文本透传（无结构信息，定位请用关键词检索）",
}


def die(msg: str, code: int = 1):
    first, *rest = str(msg).split("\n")
    print(f"[解析失败] {first}")
    for ln in rest:  # 多行消息（如 .doc 的回退指引）：续行缩进对齐，不散落成正文
        print(f"           {ln}")
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
    # 单元格全文进 md（不做截断）：截断会让评分表长单元格尾项与前附表关键值
    # （如「磋商保证金：无」）静默丢失，且派生虚假「缺失型澄清」——宁可行长勿缺。
    rows = [[" ".join((c.text or "").split()) or " " for c in r.cells]
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

def _pdf_extract_pages(pdf) -> tuple[list[list[str]], int]:
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
    return pages, total_chars


def parse_pdf(src: Path, L: Lines, out_root: Path,
              page_cap: int) -> tuple[str, list[str], int | None]:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        need_deps(["pypdfium2"])
    pdf = pdfium.PdfDocument(str(src))
    raw_pages, total_chars = _pdf_extract_pages(pdf)
    n = len(pdf)
    if total_chars < MIN_TEXT_CHARS:
        return _scanned_flow(src, pdf, L, out_root, page_cap)
    pages, hf_warnings = _strip_headers_footers(raw_pages)
    for i, p in enumerate(pages):
        L.emit(f"<!-- p:{i + 1} -->")
        for line in p:
            L.emit(line)
        L.emit("")
    L.emit("")
    warnings = ["PDF 通道按纯文本逐行提取（不识别表格框线），"
                "评分表/前附表等请按文本行逐行核对，关键表格对照页图目视核对"] + hf_warnings
    warnings += _mixed_scan_warning(raw_pages)
    _numbering_pass(L)
    conv = "pdf-numbered" if L.headings else "pdf-plain"
    if conv == "pdf-plain":
        warnings.append("未识别出章节结构（无书签/无编号），大纲不可用，定位请用关键词检索")
    return conv, warnings, n


def _mixed_scan_warning(raw_pages: list[list[str]]) -> list[str]:
    """混排 PDF：部分页文本层缺失（疑似扫描页）时告警，不阻断文本解析。"""
    n = len(raw_pages)
    if n < 5:
        return []
    weak = [i + 1 for i, p in enumerate(raw_pages) if sum(len(x) for x in p) < 30]
    if len(weak) < max(2, int(n * 0.25)):
        return []
    shown = "、".join(f"p{w}" for w in weak[:8]) + ("等" if len(weak) > 8 else "")
    return [f"第 {shown} 页文本层缺失（疑似扫描页），这些页内容不在解析产物中——"
            f"必要时对 {len(weak)} 页单独视觉补录"]


_TRANSCRIPT_HINT = ("<!-- 扫描件视觉转写模板：把每页文字誊录到对应页锚之下（规程见技能 "
                    "references/transcribe-scans.md），誊完重跑本命令完成解析 -->")


def _scanned_flow(src: Path, pdf, L: Lines, out_root: Path,
                  page_cap: int) -> tuple[str, list[str], int]:
    n = len(pdf)
    if n > page_cap:
        die(f"扫描件共 {n} 页，超过视觉转写页数上限 {page_cap}（转写耗时与页数成正比）。"
            f"确认成本后可用 --page-cap <N> 提升上限重跑")
    pages_dir = out_root / "pages"
    if not (pages_dir.is_dir() and len(list(pages_dir.glob("p*.png"))) == n):
        try:
            from PIL import Image  # noqa: F401  确认 Pillow 可用
        except ImportError:
            need_deps(["pypdfium2", "Pillow"])
        pages_dir.mkdir(parents=True, exist_ok=True)
        for i in range(n):
            pdf[i].render(scale=2).to_pil().save(pages_dir / f"p{i + 1:03d}.png")
    template = out_root / f"{src.stem}.transcript.md"
    if not template.is_file():
        parts = [_TRANSCRIPT_HINT, ""]
        for i in range(n):
            parts.append(f"<!-- p:{i + 1} -->")
            parts.append(f"<!-- 待转写：对照 pages/p{i + 1:03d}.png 逐行誊录本页文字 -->")
            parts.append("")
        template.write_text("\n".join(parts), encoding="utf-8")
    sections = _transcript_sections(template.read_text(encoding="utf-8"))
    missing = [p for p in range(1, n + 1) if not sections.get(p)]
    if missing:
        shown = "、".join(f"p{m}" for m in missing[:10]) + ("等" if len(missing) > 10 else "")
        print(f"[解析转写待办] {src.name} → {out_root.as_posix()}/")
        print(f"  扫描件 {n} 页已栅格化 → pages/pNNN.png；转写模板 {template.name}")
        print("  按 references/transcribe-scans.md 规程逐页誊录后重跑本命令完成解析")
        print(f"  缺转写页：{shown}（每页锚下至少 1 行正文，空白页写（空白页））")
        return "pdf-scanned", [], n
    _finalize_transcript(sections, n, L)
    _numbering_pass(L)
    warnings = ["扫描件视觉转写：文字/数字/专名由模型读图誊录，可能识别错误——"
                "关键值须对照 pages/ 页图复核后使用",
                "评分表/前附表等关键表格请对照页图目视核对框线、合并格与行序"]
    if not L.headings:
        warnings.append("转写文本未识别出编号标题，大纲不可用，定位请用关键词检索")
    return "pdf-scanned-transcript", warnings, n


def _transcript_sections(text: str) -> dict[int, list[str]]:
    """转写模板 → {页号: [内容行]}（注释行与空行不算内容）。"""
    sections: dict[int, list[str]] = {}
    cur: list[str] | None = None
    for raw in text.splitlines():
        line = raw.strip()
        m = re.match(r"^<!--\s*p:(\d+)\s*-->$", line)
        if m:
            cur = sections.setdefault(int(m.group(1)), [])
            continue
        if cur is None or not line or line.startswith("<!--"):
            continue
        cur.append(line)
    return sections


def _finalize_transcript(sections: dict[int, list[str]], n_pages: int, L: Lines) -> None:
    """转写内容归一化为标准 md：页锚 + 正文行，注释行剔除。"""
    for p in range(1, n_pages + 1):
        L.emit(f"<!-- p:{p} -->")
        for line in sections.get(p, []):
            L.emit(line)
        L.emit("")


# ---------------------------------------------------------------- .doc 通道

def _soffice_registered() -> list[str]:
    """Windows 注册表登记的 LibreOffice 安装目录（安装器写在 SOFTWARE\\LibreOffice\\UNO\\InstallPath，
    32/64 位各一处）；非 Windows 或读不到返回空，任何异常都只是少一个候选。"""
    try:
        import winreg
    except ImportError:
        return []
    out: list[str] = []
    views = dict.fromkeys((0, getattr(winreg, "KEY_WOW64_32KEY", 0), getattr(winreg, "KEY_WOW64_64KEY", 0)))
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in views:
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\LibreOffice\UNO\InstallPath", 0,
                                    winreg.KEY_READ | view) as key:
                    val = winreg.QueryValueEx(key, "")[0]
            except OSError:
                continue
            if val:
                out.append(str(val).rstrip("\\/") + "\\")
    return out


def _soffice_candidates() -> list[str]:
    """LibreOffice 可执行文件候选，按优先级排列（存在性由 _find_soffice 逐个探测）。

    与 docx_toolkit（交付环节）的同名函数同口径（改一处同步另一处）。
    Windows 安装器不写 PATH，只认 which 会把装好的 LibreOffice 判成「未找到」。
    soffice.com 是控制台版：转换完才返回、报错有输出可回显；soffice.exe 同目录
    并存，缺 .com 才用。当前用户级安装落在 LOCALAPPDATA，装在非默认位置靠注册表。
    """
    import shutil
    env = os.environ
    out = [w for w in (shutil.which("soffice"), shutil.which("libreoffice")) if w]
    dirs = [Path(root) / "LibreOffice" / "program"
            for root in (env.get("ProgramFiles"), env.get("ProgramW6432"), env.get("ProgramFiles(x86)"))
            if root]
    if env.get("LOCALAPPDATA"):
        dirs.append(Path(env["LOCALAPPDATA"]) / "Programs" / "LibreOffice" / "program")
    dirs += [Path(p) for p in _soffice_registered()]
    out += [str(d / exe) for d in dirs for exe in ("soffice.com", "soffice.exe")]
    out += ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice",
            "/usr/local/bin/soffice", "/opt/homebrew/bin/soffice", "/opt/homebrew/bin/libreoffice"]
    return out


def _find_soffice() -> str | None:
    for cand in _soffice_candidates():
        if cand and Path(cand).exists():
            return cand
    return None


# 脚本内置的自动转换只有 LibreOffice（本机唯一实测过的通道）。没有它时「.doc → .docx」
# 交给模型选办法：这里只把「本机装了哪些能打开 .doc 的 Office 套件」探出来写进提示，
# 不替它调用——Word/WPS 的自动化各平台各版本差异大，没在真机验证过的不内置
# （WPS for Mac 实测没有脚本接口；Windows 的自动化接口只能在 mac 上模拟）。
# 下面这一段在套件里共三份（脱敏 strip_metadata、解析 parse_document、版式读取
# docx_toolkit）——环节技能各自独立安装不共享模块，改一处同步三处，
# engine_smoke E76g 钉住三份逐字一致。
_MAC_OFFICE = (("Microsoft Word", "Microsoft Word.app"), ("WPS Office", "wpsoffice.app"),
               ("Pages", "Pages.app"))
_WIN_PROGID = (("Microsoft Word", "Word.Application"), ("WPS Office", "KWPS.Application"))


def _installed_office() -> list[str]:
    """本机已安装、能把 .doc 另存为 .docx 的 Office 套件名（去重保序）。只做存在性探测：
    macOS 看应用包，Windows 看自动化接口的注册；探测出错或别的系统返回空——只是少一条线索。"""
    found: list[str] = []
    try:
        if sys.platform == "darwin":
            for root in ("/Applications", os.path.expanduser("~/Applications")):
                for name, bundle in _MAC_OFFICE:
                    if name not in found and os.path.isdir(os.path.join(root, bundle)):
                        found.append(name)
        elif sys.platform == "win32":
            import winreg
            for name, progid in _WIN_PROGID:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                        found.append(name)
                except OSError:
                    continue
    except Exception:  # noqa: BLE001
        pass
    return found


def _hand_over_hint(why: str, next_step: str) -> str:
    """没有 LibreOffice 时给模型的回退指引：转换这一步交它选办法，后面的事仍由脚本接着做。
    why 说这份脚本为什么只认 .docx，next_step 说转好之后怎么接。"""
    found = _installed_office()
    if found:
        tools = ("转换这一步交给你选办法（宿主自带的文档转换能力也算）。本机检测到已装："
                 + "、".join(found) + "。\n"
                 "能用就自动化另存——Windows 上 Word/WPS 有自动化接口，macOS 上 Word、Pages 可脚本化，"
                 "WPS for Mac 实测没有脚本接口；不能自动化就请用户在里面对话框另存。\n")
    else:
        tools = ("转换这一步交给你选办法（宿主自带的文档转换能力也算）。"
                 "本机没检测到常见的 Office 套件（Word / WPS / Pages）。\n")
    return ("这份是 .doc（旧版二进制格式），" + why + "；本机没有找到 LibreOffice，脚本内置的自动转换只有它。\n"
            + tools + next_step)


# 超过这条线的 .doc 不自动转换（用户 2026-11-27 定：超过 5M 的提醒用户自己转换）。起因是
# 真机上一份 26.7MB 的 .doc 让 LibreOffice 转满 300 秒超时，模型随后把文字抽出来自己拼了
# 一份丢光版式的 docx 交回来——大文件既不耗那 300 秒，也不请模型另想办法。5MB 是用户定的
# 值、不是标定值：本机 9 份真实 .doc 实测 5.4–23.6MB 的都是 2–3 秒转完，只有那一份转不出来。
# 这一段同样三份逐字同口径（engine_smoke E79f 钉住）。
_DOC_AUTO_LIMIT = 5_000_000


def _oversize_doc_hint(size_bytes: int, next_step: str) -> str:
    """超过 _DOC_AUTO_LIMIT 的 .doc 的提示：不自动转换，也不请模型另想办法，请用户自己另存。
    next_step 说转好之后怎么接。"""
    return (f"这份 .doc 有 {size_bytes / 1_000_000:.1f}MB，超过 5MB：脚本不自动转换——LibreOffice 转大文件"
            "可能很久都转不完（实测一份 26.7MB 的跑满 28 分钟没转出来），另想办法转又容易丢表格和"
            "版式。\n"
            "请用户在 Word/WPS 里打开原文件另存为 .docx，这一步不要替用户自动化。\n"
            + next_step)


def _doc_too_big(path) -> bool:
    """这份 .doc 是否超过自动转换的大小线；取不到大小就当没超，照常走后面的判断。"""
    try:
        return os.stat(path).st_size > _DOC_AUTO_LIMIT
    except OSError:
        return False


# LibreOffice 转换的时限，与超时之后的处置（用户 2026-11-27 定：转换超时就 kill 掉，然后提醒用户）。
# 起因是真机上一份 .doc 让 LibreOffice 的排版陷进死循环：进程占满一个核、永不结束。两个要点：
# ①杀要杀整棵进程树——soffice 启动器在 Linux／Windows 上只是外壳，真正干活的是它拉起来的
# soffice.bin，只杀外壳会把干活的留在后台继续空转；②提醒的是用户不是模型——真机上模型被请去
# 「换一种办法」，交回来的是把文字抽出来重新拼的、丢光版式的 docx。300 秒是经验值：本机 10 份
# 能转的真实 .doc 最慢 38 秒，转不出的那一份 28 分钟也没出来。这一段同样三份逐字同口径
# （engine_smoke E80 钉住）。
_LO_TIMEOUT = 300


class _LoTimeout(Exception):
    """LibreOffice 转换超过时限没结束；抛出时进程树已经杀干净。"""


def _kill_tree(proc) -> None:
    """杀掉 proc 和它拉起来的全部子进程。POSIX 下 proc 起在独立进程组里（见 _run_lo），整组
    SIGKILL；Windows 没有进程组信号，用 taskkill /T 按进程树强杀。哪一步出错都不抛，最后都补一刀
    proc.kill()。"""
    import signal
    import subprocess
    try:
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True,
                           timeout=30, check=False)
        else:
            os.killpg(proc.pid, signal.SIGKILL)
    except Exception:  # noqa: BLE001
        pass
    try:
        proc.kill()
    except Exception:  # noqa: BLE001
        pass


def _run_lo(cmd: list, seconds: int):
    """起 LibreOffice 转换并等它结束，返回 CompletedProcess（stdout／stderr 已按本机代码页宽松
    解码，并滤掉 macOS soffice 每次先吐的系统噪音行）；超过 seconds 秒没结束就把进程树杀干净，
    抛 _LoTimeout。起不来（找不到可执行文件等）时抛 OSError，由调用方当转换失败处理。"""
    import subprocess

    def quiet(text):
        return "".join(ln for ln in (text or "").splitlines(keepends=True) if "Task policy set failed" not in ln)

    kw = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE, "text": True, "errors": "replace"}
    if sys.platform != "win32":
        kw["start_new_session"] = True    # 自成进程组：超时后才能整组杀，不牵连本脚本自己
    proc = subprocess.Popen(cmd, **kw)
    try:
        out, err = proc.communicate(timeout=seconds)
    except subprocess.TimeoutExpired:
        _kill_tree(proc)
        try:
            proc.communicate(timeout=30)  # 回收并关掉管道；万一还有活口占着管道，也不让这里卡住
        except Exception:  # noqa: BLE001
            pass
        raise _LoTimeout(seconds) from None
    return subprocess.CompletedProcess(cmd, proc.returncode, quiet(out), quiet(err))


def _stuck_doc_hint(what: str, next_step: str) -> str:
    """LibreOffice 转不出这份 .doc 时给模型的提示：转换已终止，请用户自己另存，不请模型另想办法。
    what 说这次是怎么回事（超时还是失败），next_step 说转好之后怎么接。"""
    return (f"LibreOffice 转换这份 .doc{what}，已终止——多半是个别图片／绘图对象的数据不规范，"
            "LibreOffice 排版排不出来，与文件大小无关。\n"
            "请用户在 Word/WPS 里打开原文件另存为 .docx；这一步不要替用户自动化，也不要另想办法转"
            "（把文字抽出来重新拼的 docx 会丢光表格和版式）。\n"
            + next_step)


def _doc_to_docx(src: Path, out_root: Path, handoff: Path | None = None) -> Path:
    """.doc → docx，落 out_root/_converted.docx（后续按行号拷贝格式件时从这里取）。

    模型用自选办法转好的 docx（handoff）直接采用；没给才找 LibreOffice——脚本内置的
    自动转换只有它，没有它时把「怎么转」交给模型（LibreOffice 不是前提）。
    """
    target = out_root / "_converted.docx"
    if handoff is not None:
        out_root.mkdir(parents=True, exist_ok=True)
        target.write_bytes(handoff.read_bytes())
        return target
    def give_up(what: str) -> None:
        """LibreOffice 转不出这份 .doc（超时或失败）时的唯一出口：请用户自己另存，转好带 --converted 回来。"""
        die(_stuck_doc_hint(
            what,
            f"转好后带上 --converted <那份 .docx> 重跑本命令（如 parse_document.py \"{src.name}\" "
            f"--converted \"{src.stem}.docx\"），解析仍由脚本做——来源登记、解析目录与后续按行号"
            "拷贝格式件都照原样走；原 .doc 不用动。"), 2)

    soffice = _find_soffice()
    if not soffice:
        die(_hand_over_hint(
            "解析器只认 .docx，得先把它转成 .docx",
            f"转好后带上 --converted <那份 .docx> 重跑本命令（如 parse_document.py \"{src.name}\" "
            f"--converted \"{src.stem}.docx\"），解析仍由脚本做——来源登记、解析目录与后续按行号"
            "拷贝格式件都照原样走；原 .doc 不用动。本机确实没有可用的转换方式时，"
            "再请用户在 Word/WPS 里另存为 .docx。"), 2)
    out_root.mkdir(parents=True, exist_ok=True)
    # 配置目录地址必须由 as_uri 生成：绝对路径，否则 file:// URL 失效、soffice 挂起；
    # 手工拼 file://+路径 在 Windows 上是坏地址（盘符冒号、反斜杠），含空格/# 也不转义
    profile = (out_root / ".lo-profile").resolve()
    cmd = [soffice, "--headless", f"-env:UserInstallation={profile.as_uri()}",
           "--convert-to", "docx", "--outdir", str(out_root), str(src)]
    try:
        _run_lo(cmd, _LO_TIMEOUT)
    except _LoTimeout:
        give_up(f" 超时（超过 {_LO_TIMEOUT} 秒没转完）")
    except OSError as exc:
        give_up(f" 失败（LibreOffice 起不来：{exc}）")
    produced = out_root / (src.stem + ".docx")
    if not produced.is_file():
        give_up(" 失败（没产出文件，可能版本不兼容）")
    target.write_bytes(produced.read_bytes())
    produced.unlink()
    return target


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
    text = text.lstrip("\ufeff")  # 带 BOM 的 UTF-8（Windows 记事本/PowerShell 另存）：留着会让首行标题识别不到
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


def artifacts_consistent(out_root: Path, stem: str, is_docx: bool, old: dict,
                         extra_files: list[Path] | None = None) -> tuple[bool, str]:
    """幂等跳过前的产物对账：文件齐全 + 产物哈希与登记一致 + 行号不越界。"""
    md = out_root / f"{stem}.md"
    ol = out_root / f"{stem}.outline.json"
    el = out_root / f"{stem}.elements.json"
    files = [md, ol] + ([el] if is_docx else []) + list(extra_files or [])
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


def _load_old_meta(meta_path: Path, sha: str) -> dict | None:
    """上次解析的登记：文件在、读得出、且记的就是现在这份源文件（sha256 一致）才返回，否则 None
    （meta 损坏视为无，走全量重析）。"""
    if not meta_path.is_file():
        return None
    try:
        old = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return old if isinstance(old, dict) and old.get("sha256") == sha else None


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
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(errors="replace")  # 中文 Windows 管道（GBK）下 ⚠️ 会编码崩，降级输出
    ap = argparse.ArgumentParser(description="招标文件解析（docx/doc/pdf/txt/md → md+outline+meta）")
    ap.add_argument("input", help="源文件路径")
    ap.add_argument("-o", "--out-dir", default="work/parse", help="输出根目录（默认 work/parse）")
    ap.add_argument("--page-cap", type=int, default=SCAN_PAGE_CAP,
                    help=f"扫描件视觉转写页数上限（默认 {SCAN_PAGE_CAP}）")
    ap.add_argument("--converted", default=None, metavar="转换件.docx",
                    help="这份 .doc 已经转成的 .docx（本机没有 LibreOffice，或这份 .doc 超过 5MB、脚本不自动转时用）："
                         "解析仍按 .doc 源登记，转换件落到解析目录供后续按行号拷贝格式件")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.is_file():
        die(f"文件不存在：{src}")
    suffix = src.suffix.lower()
    if suffix not in {".docx", ".doc", ".pdf", ".txt", ".md"}:
        die(f"暂不支持该格式：{suffix}（支持 .docx / .doc / .pdf / .txt / .md）")

    out_root = Path(args.out_dir) / src.name
    stem = src.stem
    md_path = out_root / (stem + ".md")
    ol_path = out_root / (stem + ".outline.json")
    meta_path = out_root / (stem + ".meta.json")
    transcript_path = out_root / (stem + ".transcript.md")

    handoff: Path | None = None   # 模型用自选办法把这份 .doc 转好的 docx（--converted）
    if args.converted:
        handoff = Path(args.converted)
        if suffix != ".doc":
            die(f"--converted 只在源文件是 .doc 时用（它是这份 .doc 转出来的 .docx），这份是 {src.name}", 2)
        if not handoff.is_file():
            die(f"--converted 指向的文件不存在：{handoff}", 2)
        if handoff.suffix.lower() != ".docx":
            die(f"--converted 须指向 .docx 文件：{handoff.name}", 2)
        with open(handoff, "rb") as fh:
            if fh.read(2) != b"PK":
                die(f"--converted 指向的不是 Word 文档（文件头不是 ZIP）：{handoff.name}", 2)

    parse_src = src
    is_docx = suffix == ".docx"
    conv_warnings: list[str] = []
    sha = sha256_of(src)

    # 已解析过且产物完好：既没有 LibreOffice 也没人交接力时照样沿用，不必再找转换件；
    # 超过自动转换大小线的 .doc 同理（下面不会再去转它）
    if suffix == ".doc" and handoff is None and (not _find_soffice() or _doc_too_big(src)):
        old = _load_old_meta(meta_path, sha)
        if old is not None and (out_root / "_converted.docx").is_file() \
                and artifacts_consistent(out_root, stem, True, old)[0]:
            print_summary("解析跳过", src.name, out_root, old)
            print("  内容未变化（sha256 一致），沿用已有产物")
            return

    if suffix == ".doc" and handoff is None and _doc_too_big(src):
        die(_oversize_doc_hint(
            os.stat(src).st_size,
            f"转好后带上 --converted <那份 .docx> 重跑本命令（如 parse_document.py \"{src.name}\" "
            f"--converted \"{src.stem}.docx\"），解析仍由脚本做——来源登记、解析目录与后续按行号"
            "拷贝格式件都照原样走；原 .doc 不用动。"), 2)

    if suffix == ".doc":
        parse_src = _doc_to_docx(src, out_root, handoff)
        is_docx = True
        conv_warnings.append("源为 .doc：" + ("已用 --converted 交来的 docx 解析" if handoff is not None
                                              else "已经 LibreOffice 转换为 docx 解析")
                             + "（版式/图片可能失真），交付级内容请核对原件")

    old = _load_old_meta(meta_path, sha)
    if old is not None:
        extra = [transcript_path] if old.get("conversion") == "pdf-scanned-transcript" else []
        ok, reason = artifacts_consistent(out_root, stem, is_docx, old, extra)
        if ok:
            print_summary("解析跳过", src.name, out_root, old)
            print("  内容未变化（sha256 一致），沿用已有产物")
            return
        print(f"[解析重建] {reason}——源文件未变，重新解析并恢复产物")

    L = Lines()
    pages = None
    elements: list[list[int]] | None = None
    try:
        if is_docx:
            conv, warnings, elements = parse_docx(parse_src, L)
        elif suffix == ".pdf":
            conv, warnings, pages = parse_pdf(src, L, out_root, args.page_cap)
        else:
            conv, warnings = parse_passthrough(src, L)
    except SystemExit:
        raise
    except Exception as exc:
        die(f"解析器内部错误（{type(exc).__name__}: {exc}）。"
            f"可尝试换用另一种格式（如 .docx / 文本型 PDF）后重试")
    warnings = conv_warnings + warnings

    if conv == "pdf-scanned":
        return  # 扫描件已栅格化待转写：待办信息已打印，产物待转写后重跑生成

    while L.lines and L.lines[-1] == "":
        L.lines.pop()
    L.lines.append("")  # 文件以单个换行结尾
    total = len(L.lines) - 1  # splitlines 口径的实际行数（outline 端点与之一致）
    outline = build_outline(L.headings, total)
    top_titles = [n["标题"] for n in outline][:12]
    out_root.mkdir(parents=True, exist_ok=True)
    md_path.write_text("\n".join(L.lines), encoding="utf-8")
    ol_path.write_text(json.dumps(outline, ensure_ascii=False, indent=1), encoding="utf-8")
    artifact_files = [md_path, ol_path]
    if elements is not None:
        elements = [[o, s_, min(e_, total)] for o, s_, e_ in elements]
        (out_root / (stem + ".elements.json")).write_text(
            json.dumps(elements, ensure_ascii=False), encoding="utf-8")
        artifact_files.append(out_root / (stem + ".elements.json"))
    if conv == "pdf-scanned-transcript":
        artifact_files.append(transcript_path)

    meta = {
        "source": src.name,
        "sha256": sha,
        "artifacts_sha256": artifacts_digest(artifact_files) or "",
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
    if conv == "pdf-scanned-transcript":
        meta["transcript_sha256"] = sha256_of(transcript_path)
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    print_summary("解析成功", src.name, out_root, meta)


if __name__ == "__main__":
    main()
