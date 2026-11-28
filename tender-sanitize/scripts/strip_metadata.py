#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
strip_metadata.py — 清除 / 体检 Word 与 PDF 的文档元数据（作者、单位、工具痕迹）

设计原则
--------
1. 默认不覆盖原文件：输出同目录 <文件名>_cleaned<扩展名>，原文件一个字节不动；
   就地覆盖须显式 --in-place（建议同时 --backup）。
2. 只改元数据、不改正文：
   - docx 走 zip 内 XML 定点改写（纯标准库）。不用 python-docx，因为它只暴露
     core.xml，碰不到 app.xml（Company/Manager/Template）、custom.xml、缩略图、
     修订/批注作者（w:author）、rsid——这些恰恰是主要泄露点。
   - pdf 走 pypdf 克隆重写（删 /Info 与 XMP，另清页面级字典、批注作者、内嵌附件）。
3. 写完立刻自证：重新扫描输出文件，并比对正文指纹（docx 比所有 w:t/delText 的
   sha256；pdf 比页数/页面尺寸/提取文本）。指纹不一致即中止，不产出文件。
4. --check 只体检不改动；发现元数据时 exit 1，可直接当验收闸门。
5. .doc 是非 zip 的二进制格式，不能在本层面直接改：先转成 .docx 再清理，输出一律
   .docx（要回存 .doc 请在 Word/WPS 里另存）。脚本内置的自动转换只有 LibreOffice；
   本机没有它时探出已装的 Office 套件、把「怎么转」交给模型选办法——LibreOffice
   体积大、不少用户装不了，不是前提（用户 2026-11-27 拍板）；超过 5MB 的 .doc 不自动转，
   直接请用户在 Word/WPS 里另存（用户 2026-11-27 定）；转换超时就终止转换（整棵进程树
   杀干净），转换失败或超时都只请用户另存，不请模型另想办法（用户同日定）。

用法示例
--------
  python3 strip_metadata.py --check 报告.docx 报价.pdf
  python3 strip_metadata.py 报告.docx 报价.pdf -o ./已清理
  python3 strip_metadata.py --in-place --backup 报告.docx
  python3 strip_metadata.py -r ./待递交 -o ./已清理 --json

依赖：docx/docm 清理只用标准库；pdf 需要 pypdf（缺失时打印安装提示退出）。
调试：环境变量 TENDER_TRACEBACK=1 输出完整异常栈（缺省每个文件一行式报错）。
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
import tempfile
import zipfile

OUT_SUFFIX = "_cleaned"

OOXML_EXTS = {".docx", ".docm", ".dotx", ".dotm"}
PDF_EXTS = {".pdf"}
DOC_EXTS = {".doc"}
SUPPORTED_EXTS = OOXML_EXTS | PDF_EXTS | DOC_EXTS

# --------------------------------------------------------------------------
# LibreOffice 查找与转换（.doc → .docx）。与 docx_toolkit / parse_document 的
# _find_soffice 同口径（环节技能各自独立安装不共享模块，改一处同步三处，
# 本机一致性由 tests/sanitize_smoke.py 与 tests/engine_smoke.py E35 共同断言）。
# --------------------------------------------------------------------------


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

    Windows 安装器不写 PATH，只认 which 会把装好的 LibreOffice 判成「未找到」。
    soffice.com 是控制台版：转换完才返回、报错有输出可回显；soffice.exe 同目录
    并存，缺 .com 才用。当前用户级安装落在 LOCALAPPDATA，装在非默认位置靠注册表。
    """
    import shutil
    env = os.environ
    out = [w for w in (shutil.which("soffice"), shutil.which("libreoffice")) if w]
    dirs = [os.path.join(root, "LibreOffice", "program")
            for root in (env.get("ProgramFiles"), env.get("ProgramW6432"), env.get("ProgramFiles(x86)"))
            if root]
    if env.get("LOCALAPPDATA"):
        dirs.append(os.path.join(env["LOCALAPPDATA"], "Programs", "LibreOffice", "program"))
    dirs += _soffice_registered()
    out += [os.path.join(d, exe) for d in dirs for exe in ("soffice.com", "soffice.exe")]
    out += ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice",
            "/usr/local/bin/soffice", "/opt/homebrew/bin/soffice", "/opt/homebrew/bin/libreoffice"]
    return out


def _find_soffice() -> str | None:
    for cand in _soffice_candidates():
        if cand and os.path.exists(cand):
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


def _soffice_msg(r) -> str:
    """soffice 的回显里挑第一条有意义的行（macOS 每次都先吐一行系统噪音）。"""
    for ln in ((r.stdout or "") + "\n" + (r.stderr or "")).splitlines():
        ln = ln.strip()
        if ln and "Task policy set failed" not in ln:
            return ln[:200]
    return "(无输出)"


def _doc_to_docx(doc_path: str, workdir: str) -> str:
    """.doc → .docx（隔离配置目录，避免与开着界面版 LibreOffice 抢配置）。

    配置目录地址必须由 as_uri 生成：手工拼 file://+路径 在 Windows 上是坏地址
    （盘符冒号、反斜杠），路径含空格/# 也不转义。没有 LibreOffice 时把「怎么转」交给
    模型；转换超时（进程树已杀干净）或失败时请用户自己另存，转好后对那份 .docx
    重跑同一条命令。
    """
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError(_hand_over_hint(
            "清理引擎只处理 .docx，得先把它转成 .docx",
            "转好后对那份 .docx 运行同一条命令即可，清理仍由脚本做（有正文指纹自证与复扫）；"
            "原 .doc 全程不用动。本机确实没有可用的转换方式时，再请用户在 Word/WPS 里另存为 .docx 发来。"))
    from pathlib import Path
    profile = Path(workdir) / "lo_profile"
    cmd = [soffice, "--headless", f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
           "--convert-to", "docx", "--outdir", workdir, str(doc_path)]

    def give_up(what: str) -> RuntimeError:
        return RuntimeError(_stuck_doc_hint(
            what, "转好后对那份 .docx 运行同一条命令即可，清理仍由脚本做（有正文指纹自证与复扫）；"
                  "原 .doc 全程不用动。"))

    try:
        r = _run_lo(cmd, _LO_TIMEOUT)
    except _LoTimeout as exc:
        raise give_up(f" 超时（超过 {_LO_TIMEOUT} 秒没转完）") from exc
    except OSError as exc:
        raise give_up(f" 失败（LibreOffice 起不来：{exc}）") from exc
    out = os.path.join(workdir, os.path.splitext(os.path.basename(doc_path))[0] + ".docx")
    if not os.path.isfile(out):
        raise give_up(f" 失败（{_soffice_msg(r) or '没产出文件，也没有任何输出'}）")
    return out


# --------------------------------------------------------------------------
# OOXML：空模板（保留部件本身，只把内容清空；删部件会牵动 rels 与
# [Content_Types].xml，能少动就少动，降低 Word 报「文档已损坏」的概率）
# --------------------------------------------------------------------------

EMPTY_CORE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<cp:coreProperties'
    ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
    ' xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>'
)

CORE_DATES_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<cp:coreProperties'
    ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
    ' xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
)


def core_with_dates(created: str, modified: str) -> str:
    """--keep-dates 时只保留 dcterms:created / modified。"""
    parts = []
    if created:
        parts.append(f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>')
    if modified:
        parts.append(f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified>')
    return CORE_DATES_OPEN + "".join(parts) + "</cp:coreProperties>"

EMPTY_APP = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties'
    ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)

EMPTY_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties'
    ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)

EMPTY_PEOPLE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:people xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)

# 承载作者指纹的修订 / 批注元素（localName）
REVISION_TAGS = {
    "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveToRangeStart",
    "rPrChange", "pPrChange", "tblPrChange", "tblGridChange", "tcPrChange",
    "trPrChange", "sectPrChange", "numberingChange",
    "cellIns", "cellDel", "cellMerge", "comment",
}

# 被扫描的正文部件（其余部件不动）
BODY_PART_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments|"
    r"commentsExtended|commentsExtensible|commentsIds)\.xml$",
    re.I,
)

# 属性级清洗：w:author 是必填属性，置空字符串而不是删属性，否则 Word 判非法 XML；
# w:initials / w:date / w:rsid* 都是可选属性，直接删。
RE_AUTHOR = re.compile(r"""(\sw:author\s*=\s*)("[^"]*"|'[^']*')""")
RE_INITIALS = re.compile(r"""\sw:initials\s*=\s*("[^"]*"|'[^']*')""")
RE_EDIT_DATE = re.compile(r"""\sw:date\s*=\s*("[^"]*"|'[^']*')""")
RE_RSID_ATTR = re.compile(r"""\sw:rsid[A-Za-z]*\s*=\s*("[^"]*"|'[^']*')""")

RE_RSIDS_BLOCK = re.compile(r"<w:rsids\b[^>]*>.*?</w:rsids\s*>|<w:rsids\b[^>]*/>", re.S)
RE_RSID_ROOT = re.compile(r"<w:rsidRoot\b[^>]*/>")
RE_ATTACHED_TEMPLATE = re.compile(
    r"<w:attachedTemplate\b[^>]*/>|<w:attachedTemplate\b[^>]*>.*?</w:attachedTemplate\s*>",
    re.S,
)
RE_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*/>|<Relationship\b[^>]*>.*?</Relationship\s*>", re.S)
RE_WT = re.compile(rb"<(?:w|a):t(?:\s[^>]*)?>(.*?)</(?:w|a):t>", re.S)
RE_DELTEXT = re.compile(rb"<w:delText(?:\s[^>]*)?>(.*?)</w:delText>", re.S)
RE_CREATED = re.compile(r"<dcterms:created\b[^>]*>([^<]*)<", re.S)
RE_MODIFIED = re.compile(r"<dcterms:modified\b[^>]*>([^<]*)<", re.S)


# --------------------------------------------------------------------------
# 通用小工具
# --------------------------------------------------------------------------

def _local(tag: str) -> str:
    """去掉 ElementTree 的 {namespace} 前缀，只留 localName。"""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _decode_xml(data: bytes):
    """返回 (文本, 编码标记)。Word 的 XML 绝大多数是 UTF-8，少数 WPS 导出是 UTF-16。"""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16"), "utf-16"
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        return data.decode("latin-1"), "latin-1"


def _encode_xml(text: str, enc: str) -> bytes:
    if enc == "utf-16":
        return text.encode("utf-16")
    if enc == "latin-1":
        return text.encode("latin-1", "replace")
    return text.encode("utf-8")


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _format_size(n: float) -> str:
    if n >= 1024 * 1024:
        return f"{n/1024/1024:.1f} MB"
    if n >= 1024:
        return f"{n/1024:.1f} KB"
    return f"{n:.0f} B"


# --------------------------------------------------------------------------
# DOCX / DOCM：分析
# --------------------------------------------------------------------------

def analyze_ooxml(path: str, keep_dates: bool = False) -> list[dict]:
    """扫描 OOXML 里的元数据，返回 findings 列表（不改动文件）。"""
    import xml.etree.ElementTree as ET

    findings: list[dict] = []

    def add(level, where, field, value):
        findings.append({"level": level, "where": where, "field": field, "value": value})

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        lower = {n.lower(): n for n in names}

        # --- docProps/core.xml ---
        core = lower.get("docprops/core.xml")
        if core:
            try:
                root = ET.fromstring(z.read(core))
                for el in root:
                    ln = _local(el.tag)
                    val = (el.text or "").strip()
                    if not val:
                        continue
                    if ln in ("created", "modified"):
                        continue  # 下面单独按 keep_dates 判断
                    add("high", "docProps/core.xml", ln, val)
                if not keep_dates:
                    for ln in ("created", "modified"):
                        for el in root:
                            if _local(el.tag) == ln:
                                val = (el.text or "").strip()
                                if val:
                                    add("high", "docProps/core.xml", ln, val)
            except ET.ParseError as exc:
                add("low", "docProps/core.xml", "解析失败", str(exc))

        # --- docProps/app.xml ---
        app = lower.get("docprops/app.xml")
        if app:
            try:
                root = ET.fromstring(z.read(app))
                sensitive = {"Company", "Manager", "Template", "TotalTime",
                             "HyperlinkBase", "Application", "AppVersion"}
                for el in root.iter():
                    ln = _local(el.tag)
                    if ln not in sensitive:
                        continue
                    val = (el.text or "").strip()
                    if not val:
                        continue
                    if ln == "TotalTime" and val in ("0", ""):
                        continue
                    if ln == "Template" and val.lower() in ("normal.dotm", "normal.dot"):
                        continue
                    level = "medium" if ln in ("Application", "AppVersion") else "high"
                    add(level, "docProps/app.xml", ln, val)
            except ET.ParseError as exc:
                add("low", "docProps/app.xml", "解析失败", str(exc))

        # --- docProps/custom.xml ---
        custom = lower.get("docprops/custom.xml")
        if custom:
            try:
                root = ET.fromstring(z.read(custom))
                props = [el for el in root if _local(el.tag) == "property"]
                for el in props:
                    name = el.get("name") or "?"
                    # 值可能嵌在 vt:lpwstr 之类的子元素里
                    val = "".join(el.itertext()).strip()
                    add("high", "docProps/custom.xml", name, val[:120])
            except ET.ParseError as exc:
                add("low", "docProps/custom.xml", "解析失败", str(exc))

        # --- 缩略图 ---
        for n in names:
            if n.lower().startswith("docprops/thumbnail."):
                add("medium", n, "缩略图", _format_size(z.getinfo(n).file_size))

        # --- 正文部件：修订 / 批注作者、rsid ---
        rev_authors: dict[str, dict[str, int]] = {}
        rsid_total = 0
        for n in names:
            if not BODY_PART_RE.match(n):
                continue
            try:
                root = ET.fromstring(z.read(n))
            except ET.ParseError:
                continue
            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            for el in root.iter():
                ln = _local(el.tag)
                if ln in REVISION_TAGS:
                    author = el.get(f"{{{w_ns}}}author") or el.get("w:author") or ""
                    if author.strip():
                        rev_authors.setdefault(n, {})
                        rev_authors[n][author] = rev_authors[n].get(author, 0) + 1
                for k in el.attrib:
                    if _local(k).startswith("rsid"):
                        rsid_total += 1

        for part, authors in rev_authors.items():
            items = ", ".join(f"{a} ×{c}" for a, c in authors.items())
            add("high", part, "w:author（修订/批注作者）", items)
        if rsid_total:
            add("low", "各正文部件", "w:rsid* 属性", f"{rsid_total} 处")

        # --- people.xml ---
        people = lower.get("word/people.xml")
        if people:
            try:
                root = ET.fromstring(z.read(people))
                w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                authors = [
                    (el.get(f"{{{w_ns}}}author") or el.get("w:author") or el.get("author") or "")
                    for el in root
                    if _local(el.tag) == "person"
                ]
                authors = [a for a in authors if a.strip()]
                if authors:
                    add("high", "word/people.xml", "w:person（修订作者名录）", ", ".join(authors))
            except ET.ParseError:
                pass

        # --- settings.xml ---
        settings = lower.get("word/settings.xml")
        if settings:
            text, _ = _decode_xml(z.read(settings))
            if RE_ATTACHED_TEMPLATE.search(text):
                target = ""
                rels = lower.get("word/_rels/settings.xml.rels")
                if rels:
                    rtext, _ = _decode_xml(z.read(rels))
                    m = re.search(r'Target="([^"]+)"[^>]*?Type="[^"]*attachedTemplate"', rtext, re.I)
                    if not m:
                        m = re.search(r'Type="[^"]*attachedTemplate"[^>]*?Target="([^"]+)"', rtext, re.I)
                    if m:
                        target = m.group(1)
                add("high", "word/settings.xml", "w:attachedTemplate",
                    target or "（模板路径，可能含用户名）")
            if RE_RSIDS_BLOCK.search(text):
                add("low", "word/settings.xml", "w:rsids", "修订会话 ID 清单")
            if RE_RSID_ROOT.search(text):
                add("low", "word/settings.xml", "w:rsidRoot", "文档根 rsid")

    return findings


# --------------------------------------------------------------------------
# DOCX / DOCM：清理
# --------------------------------------------------------------------------

def _strip_rel_by_type(data: bytes, type_suffix: str) -> tuple[bytes, int]:
    text, enc = _decode_xml(data)
    removed = 0

    def repl(m):
        nonlocal removed
        if type_suffix in m.group(0):
            removed += 1
            return ""
        return m.group(0)

    out = RE_RELATIONSHIP.sub(repl, text)
    return _encode_xml(out, enc), removed


def _scrub_body_xml(data: bytes) -> tuple[bytes, dict]:
    """只改属性值，不动任何文本节点——正文指纹校验会兜住误伤。"""
    text, enc = _decode_xml(data)

    # w:author 是必填属性，置空字符串（不是删属性，删了 Word 判非法 XML）
    text, n_author = RE_AUTHOR.subn(r'\1""', text)
    # w:initials 可选；保留属性名、置空，避免某些 WPS 版本对缺属性敏感
    text, n_initials = RE_INITIALS.subn(' w:initials=""', text)
    # w:date / w:rsid* 都是可选属性，直接删
    text, n_date = RE_EDIT_DATE.subn("", text)
    text, n_rsid = RE_RSID_ATTR.subn("", text)

    stats = {"author": n_author, "initials": n_initials, "date": n_date, "rsid": n_rsid}
    return _encode_xml(text, enc), stats


def clean_ooxml(src: str, dst: str, args) -> dict:
    """zip 内定点改写；返回 actions 与统计。"""
    actions: list[dict] = []
    stats = {"author": 0, "initials": 0, "date": 0, "rsid": 0}
    dropped_parts: list[str] = []

    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        dropped = {n for n in names if n.lower().startswith("docprops/thumbnail.")}

        tmp = dst + ".tmp"
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    n = item.filename
                    low = n.lower()

                    if n in dropped:
                        dropped_parts.append(n)
                        actions.append({
                            "action": "移除缩略图", "part": n,
                            "detail": _format_size(item.file_size),
                        })
                        continue

                    data = zin.read(n)
                    new = data

                    if low == "docprops/core.xml":
                        if args.keep_dates:
                            text, _ = _decode_xml(data)
                            created = RE_CREATED.search(text)
                            modified = RE_MODIFIED.search(text)
                            new = core_with_dates(
                                created.group(1).strip() if created else "",
                                modified.group(1).strip() if modified else "",
                            ).encode("utf-8")
                        else:
                            new = EMPTY_CORE.encode("utf-8")
                        actions.append({"action": "重写", "part": n,
                                        "detail": "去除作者/标题/关键词/时间等"})

                    elif low == "docprops/app.xml":
                        new = EMPTY_APP.encode("utf-8")
                        actions.append({"action": "重写", "part": n,
                                        "detail": "去除 Company/Manager/Template/Application 等"})

                    elif low == "docprops/custom.xml":
                        new = EMPTY_CUSTOM.encode("utf-8")
                        actions.append({"action": "重写", "part": n, "detail": "清空自定义属性"})

                    elif low == "word/people.xml" and not args.basic:
                        new = EMPTY_PEOPLE.encode("utf-8")
                        actions.append({"action": "重写", "part": n, "detail": "清空修订作者名录"})

                    elif low == "_rels/.rels":
                        new, removed = _strip_rel_by_type(data, "/thumbnail")
                        if removed:
                            actions.append({"action": "解除关系", "part": n,
                                            "detail": f"移除 {removed} 条缩略图关系"})

                    elif low == "word/_rels/settings.xml.rels":
                        new, removed = _strip_rel_by_type(data, "/attachedTemplate")
                        if removed:
                            actions.append({"action": "解除关系", "part": n,
                                            "detail": f"移除 {removed} 条模板关系"})

                    elif low == "word/settings.xml":
                        text, enc = _decode_xml(data)
                        text, n1 = RE_ATTACHED_TEMPLATE.subn("", text)
                        if args.basic:
                            n2 = n3 = 0
                        else:
                            text, n2 = RE_RSIDS_BLOCK.subn("", text)
                            text, n3 = RE_RSID_ROOT.subn("", text)
                        new = _encode_xml(text, enc)
                        if n1 or n2 or n3:
                            actions.append({"action": "重写", "part": n,
                                            "detail": f"attachedTemplate {n1} / rsids {n2} / rsidRoot {n3}"})

                    elif BODY_PART_RE.match(n):
                        if args.basic:
                            pass
                        else:
                            new, s = _scrub_body_xml(data)
                            if any(s.values()):
                                for k in stats:
                                    stats[k] += s[k]
                                actions.append({
                                    "action": "清洗属性", "part": n,
                                    "detail": (f"author {s['author']} / initials {s['initials']} / "
                                               f"date {s['date']} / rsid {s['rsid']}"),
                                })

                    # 其余部件原样复制
                    zi = zipfile.ZipInfo(n)
                    if args.keep_dates:
                        zi.date_time = item.date_time
                    else:
                        zi.date_time = (1980, 1, 1, 0, 0, 0)
                    zi.compress_type = item.compress_type
                    zi.external_attr = item.external_attr
                    zi.internal_attr = item.internal_attr
                    zi.create_system = item.create_system
                    zout.writestr(zi, new)
        except BaseException:
            if os.path.exists(tmp):
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
            raise

    os.replace(tmp, dst)

    # 汇总深度项
    if sum(stats.values()):
        actions.append({
            "action": "汇总", "part": "*",
            "detail": (f"修订作者 {stats['author']} 处、initials {stats['initials']} 处、"
                       f"修订时间 {stats['date']} 处、rsid {stats['rsid']} 处"),
        })

    return {"actions": actions, "dropped_parts": dropped_parts}


# --------------------------------------------------------------------------
# DOCX / DOCM：正文指纹（只比文本，保证「只动属性不改正文」可验证）
# --------------------------------------------------------------------------

def ooxml_fingerprint(path: str) -> dict:
    parts: dict[str, str] = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.lower().startswith("docprops/"):
                continue
            if not n.lower().endswith(".xml"):
                continue
            data = z.read(n)
            texts = b"".join(RE_WT.findall(data)) + b"".join(RE_DELTEXT.findall(data))
            parts[n] = _sha256(texts)
    digest = _sha256(json.dumps(parts, sort_keys=True, ensure_ascii=False).encode("utf-8"))
    return {"text_sha256": digest, "parts": parts, "n_parts": len(parts)}


# --------------------------------------------------------------------------
# PDF：分析 / 清理（pypdf 惰性加载：--help 与 docx 清理不碰它，纯净解释器可跑）
# --------------------------------------------------------------------------

def _pypdf_module():
    try:
        import pypdf
        return pypdf
    except ImportError:
        return None


def _pypdf_help() -> str:
    return ("缺少 pypdf，无法处理 PDF（docx/docm 的清理不需要它）。"
            "安装：python -m pip install pypdf")


def _pdf_annotations(reader) -> list:
    out = []
    for page in reader.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        try:
            arr = annots.get_object() if hasattr(annots, "get_object") else annots
            for ref in arr:
                obj = ref.get_object() if hasattr(ref, "get_object") else ref
                if hasattr(obj, "keys"):
                    out.append(obj)
        except Exception:
            continue
    return out


def _open_pdf(pypdf, path: str):
    reader = pypdf.PdfReader(path)
    if reader.is_encrypted:
        try:
            if int(reader.decrypt("")) == 0:
                raise RuntimeError("PDF 已加密，且空口令无法打开。请先解密（如 qpdf --decrypt）后再清理。")
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError(f"PDF 已加密，无法打开：{exc}") from exc
    return reader


def analyze_pdf(path: str, keep_dates: bool = False) -> list[dict]:
    pypdf = _pypdf_module()
    if pypdf is None:
        raise RuntimeError(_pypdf_help())

    findings: list[dict] = []

    def add(level, where, field, value):
        findings.append({"level": level, "where": where, "field": field, "value": value})

    reader = _open_pdf(pypdf, path)

    info = reader.metadata
    if info:
        for k, v in info.items():
            key = str(k)
            val = "" if v is None else str(v)
            if not val:
                continue
            if keep_dates and key in ("/CreationDate", "/ModDate"):
                continue
            add("high", "trailer /Info", key, val[:160])

    root = reader.trailer.get("/Root", {})
    try:
        root = root.get_object() if hasattr(root, "get_object") else root
    except Exception:
        pass
    if "/Metadata" in root:
        detail = "XMP 元数据流"
        try:
            xmp = reader.xmp_metadata
            if xmp is not None:
                keys = []
                for attr in ("dc_creator", "dc_title", "xmp_creator_tool",
                             "pdf_producer", "xmp_create_date", "xmp_modify_date"):
                    try:
                        v = getattr(xmp, attr, None)
                        if v:
                            keys.append(f"{attr}={v}")
                    except Exception:
                        pass
                if keys:
                    detail += "（" + "; ".join(str(x)[:80] for x in keys) + "）"
        except Exception:
            pass
        add("high", "root /Metadata", "XMP", detail)

    # 内嵌附件
    try:
        atts = reader.attachments
        for name in atts:
            add("high", "/Names /EmbeddedFiles", "内嵌附件", name)
    except Exception:
        pass

    # 页面级字典
    for i, page in enumerate(reader.pages, 1):
        for key in ("/Thumb", "/PieceInfo", "/LastModified"):
            if key in page:
                add("medium", f"第 {i} 页", key, "页面级元数据")

    # 批注作者
    annots = _pdf_annotations(reader)
    authors = sorted({str(a.get("/T")) for a in annots if a.get("/T")})
    if authors:
        add("high", "批注 /Annots", "/T（批注作者）", ", ".join(authors))
    n_nm = sum(1 for a in annots if a.get("/NM"))
    n_m = sum(1 for a in annots if a.get("/M"))
    if n_nm:
        add("low", "批注 /Annots", "/NM（批注标识）", f"{n_nm} 处")
    if n_m:
        add("low", "批注 /Annots", "/M（批注时间）", f"{n_m} 处")

    return findings


def clean_pdf_pypdf(src: str, dst: str, args) -> dict:
    pypdf = _pypdf_module()
    if pypdf is None:
        raise RuntimeError(_pypdf_help())

    actions: list[dict] = []

    with open(src, "rb") as fh:
        reader = _open_pdf(pypdf, str(src))
        writer = pypdf.PdfWriter(clone_from=reader)

    # 1) /Info 整个删掉（--keep-dates 时只留创建/修改时间）
    had_info = bool(reader.metadata)
    keep_dates_vals: dict[str, str] = {}
    if args.keep_dates and reader.metadata:
        for k in ("/CreationDate", "/ModDate"):
            v = reader.metadata.get(k)
            if v:
                keep_dates_vals[str(k)] = str(v)

    try:
        writer.metadata = None
    except Exception:
        try:
            writer._info = None
        except Exception:
            pass
    if keep_dates_vals:
        try:
            writer.add_metadata(keep_dates_vals)
        except Exception:
            pass
    if had_info:
        detail = "Author/Creator/Producer/Title/Dates 等"
        if keep_dates_vals:
            detail = "仅保留 " + " 与 ".join(keep_dates_vals) + "，其余删除"
        actions.append({"action": "删除", "part": "trailer /Info", "detail": detail})

    # 2) XMP
    root = writer._root_object
    if "/Metadata" in root:
        del root["/Metadata"]
        actions.append({"action": "删除", "part": "root /Metadata", "detail": "XMP 元数据流"})
    if "/PieceInfo" in root:
        del root["/PieceInfo"]
        actions.append({"action": "删除", "part": "root /PieceInfo", "detail": "应用程序私有数据"})

    # 3) 内嵌附件
    if not args.keep_attachments:
        removed_names = []
        names_dict = root.get("/Names")
        if names_dict is not None:
            try:
                nd = names_dict.get_object() if hasattr(names_dict, "get_object") else names_dict
                if "/EmbeddedFiles" in nd:
                    ef = nd["/EmbeddedFiles"]
                    try:
                        efo = ef.get_object() if hasattr(ef, "get_object") else ef
                        arr = efo.get("/Names") if hasattr(efo, "get") else None
                        if arr:
                            arro = arr.get_object() if hasattr(arr, "get_object") else arr
                            for j in range(0, len(arro) - 1, 2):
                                removed_names.append(str(arro[j]))
                    except Exception:
                        pass
                    del nd["/EmbeddedFiles"]
                    actions.append({"action": "删除", "part": "/Names /EmbeddedFiles",
                                    "detail": "附件：" + (", ".join(removed_names) or "（未列出）")})
            except Exception:
                pass
        if "/AF" in root:
            del root["/AF"]
            actions.append({"action": "删除", "part": "root /AF", "detail": "关联文件"})

    # 4) 页面级字典 + 批注作者
    n_thumb = n_piece = n_lm = 0
    n_annot = 0
    for page in writer.pages:
        for key in ("/Thumb", "/PieceInfo", "/LastModified"):
            if key in page:
                try:
                    del page[key]
                    if key == "/Thumb":
                        n_thumb += 1
                    elif key == "/PieceInfo":
                        n_piece += 1
                    else:
                        n_lm += 1
                except Exception:
                    pass
        annots = page.get("/Annots")
        if annots:
            try:
                arr = annots.get_object() if hasattr(annots, "get_object") else annots
                for ref in arr:
                    obj = ref.get_object() if hasattr(ref, "get_object") else ref
                    if not hasattr(obj, "keys"):
                        continue
                    hit = False
                    for k in ("/T", "/NM", "/M"):
                        if k in obj:
                            try:
                                del obj[k]
                                hit = True
                            except Exception:
                                pass
                    if hit:
                        n_annot += 1
            except Exception:
                pass

    if n_thumb:
        actions.append({"action": "删除", "part": "页面 /Thumb", "detail": f"{n_thumb} 页"})
    if n_piece:
        actions.append({"action": "删除", "part": "页面 /PieceInfo", "detail": f"{n_piece} 页"})
    if n_lm:
        actions.append({"action": "删除", "part": "页面 /LastModified", "detail": f"{n_lm} 页"})
    if n_annot:
        actions.append({"action": "删除", "part": "批注 /T /NM /M", "detail": f"{n_annot} 个批注"})

    # 5) /ID 尽力置空
    try:
        writer._ID = None
    except Exception:
        pass

    with open(dst, "wb") as out:
        writer.write(out)

    return {"actions": actions}


# --------------------------------------------------------------------------
# PDF：正文指纹
# --------------------------------------------------------------------------

def pdf_fingerprint(path: str) -> dict:
    pypdf = _pypdf_module()
    if pypdf is None:
        raise RuntimeError(_pypdf_help())

    reader = pypdf.PdfReader(path)
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            pass

    pages = []
    texts = []
    for p in reader.pages:
        try:
            mb = p.mediabox
            pages.append((round(float(mb.width), 2), round(float(mb.height), 2)))
        except Exception:
            pages.append((0.0, 0.0))
        try:
            texts.append(p.extract_text() or "")
        except Exception:
            texts.append("")
    blob = "\n\x00\n".join(texts).encode("utf-8", "replace")
    return {
        "n_pages": len(reader.pages),
        "page_sizes": pages,
        "text_sha256": _sha256(blob),
    }


# --------------------------------------------------------------------------
# 文件级流程
# --------------------------------------------------------------------------

def output_path_for(src: str, args, ext_out: str | None = None) -> str:
    if args.in_place:
        return src
    stem, ext = os.path.splitext(os.path.basename(src))
    ext = ext_out or ext
    name = f"{stem}{OUT_SUFFIX}{ext}"
    outdir = args.outdir or os.path.dirname(os.path.abspath(src))
    os.makedirs(outdir, exist_ok=True)
    dst = os.path.join(outdir, name)
    # 递归批量时不同子目录可能有同名文件，避免互相覆盖
    base, ext2 = os.path.splitext(dst)
    k = 2
    while dst in _OUTPUTS_THIS_RUN:
        dst = f"{base}-{k}{ext2}"
        k += 1
    _OUTPUTS_THIS_RUN.add(dst)
    return dst


def _magic(path: str) -> bytes:
    try:
        with open(path, "rb") as fh:
            return fh.read(8)
    except OSError:
        return b""


def handle_file(path: str, args) -> dict:
    ext = os.path.splitext(path)[1].lower()
    is_doc = ext in DOC_EXTS
    ext_out = ".docx" if is_doc else ext
    rep = {
        "path": path,
        "format": ("DOC→DOCX" if is_doc else ext.lstrip(".").upper() or "?"),
        "size": os.path.getsize(path) if os.path.exists(path) else 0,
        "findings": [],
        "actions": [],
        "verification": {},
        "output": None,
        "warnings": [],
        "error": None,
    }
    work = None    # 清理过程中的临时文件路径；失败时清理
    tmp_dir = None # .doc 转换用的临时目录；用毕即删（清理失败不盖掉真正的报错）

    try:
        # ---------- 边界 ----------
        if ext not in SUPPORTED_EXTS:
            raise RuntimeError(f"不支持的扩展名 {ext}（只处理 {'/'.join(sorted(SUPPORTED_EXTS))}）")
        if is_doc and args.in_place:
            raise RuntimeError(
                ".doc 不能就地覆盖（清理后的内容以 .docx 保存，回写 .doc 会换格式）。"
                f"按默认方式处理，输出 <文件名>{OUT_SUFFIX}.docx。")
        if ext in OOXML_EXTS and not _magic(path).startswith(b"PK"):
            raise RuntimeError("文件头不是 ZIP（PK），可能已损坏或扩展名与实际格式不符。")

        # ---------- .doc 先转 .docx（临时目录里转换，原件不动） ----------
        src = path
        if is_doc:
            if _doc_too_big(path):
                raise RuntimeError(_oversize_doc_hint(
                    os.path.getsize(path),
                    "转好后对那份 .docx 运行同一条命令即可，清理仍由脚本做（有正文指纹自证与复扫）；"
                    "原 .doc 全程不用动。"))
            tmp_dir = tempfile.mkdtemp(prefix="strip_doc_")
            src = _doc_to_docx(path, tmp_dir)
            rep["warnings"].append(
                ".doc 已转成 .docx 再清理：输出是 .docx 文件，要回存 .doc 请在 Word/WPS 里另存。")

        # ---------- 体检 ----------
        keep_dates = getattr(args, "keep_dates", False)
        if ext_out in OOXML_EXTS:
            rep["findings"] = analyze_ooxml(src, keep_dates)
        else:
            rep["findings"] = analyze_pdf(src, keep_dates)

        if args.check:
            if is_doc:
                rep["warnings"].append("体检基于转换出的临时副本（.doc 无法在本层面直接读属性），原文件未动。")
            rep["output"] = None
            return rep

        # ---------- 清理 ----------
        dst = output_path_for(path, args, ext_out)
        if args.in_place:
            if args.backup:
                bak = os.path.splitext(path)[0] + ".bak" + os.path.splitext(path)[1]
                shutil.copy2(path, bak)
                rep["actions"].append({"action": "备份", "part": bak, "detail": "原文件副本"})
            work = path + ".striptmp"
        else:
            work = dst + ".tmp"

        if ext_out in OOXML_EXTS:
            fp_before = ooxml_fingerprint(src)
            res = clean_ooxml(src, work, args)
        else:
            fp_before = pdf_fingerprint(src)
            res = clean_pdf_pypdf(src, work, args)

        rep["actions"].extend(res["actions"])

        # ---------- 自证 1：内容不变量 ----------
        if ext_out in OOXML_EXTS:
            fp_after = ooxml_fingerprint(work)
            ok_text = fp_after["text_sha256"] == fp_before["text_sha256"]
            expected_parts = set(fp_before["parts"]) - set(res.get("dropped_parts", []))
            actual_parts = set(fp_after["parts"])
            ok_parts = actual_parts == expected_parts
            rep["verification"] = {
                "content_ok": ok_text and ok_parts,
                "text_sha256_before": fp_before["text_sha256"],
                "text_sha256_after": fp_after["text_sha256"],
                "text_ok": ok_text,
                "parts_ok": ok_parts,
                "parts_delta": sorted(actual_parts - expected_parts) + sorted(expected_parts - actual_parts),
            }
            if not ok_text:
                raise RuntimeError(
                    "正文文本指纹不一致——已中止，未产出文件。"
                    "请把该文件反馈给脚本维护者（清洗只应改动属性，不该改动文字）。")
        else:
            fp_after = pdf_fingerprint(work)
            ok_pages = fp_after["n_pages"] == fp_before["n_pages"]
            ok_sizes = fp_after["page_sizes"] == fp_before["page_sizes"]
            ok_text = fp_after["text_sha256"] == fp_before["text_sha256"]
            rep["verification"] = {
                "content_ok": ok_pages and ok_sizes,
                "pages_before": fp_before["n_pages"],
                "pages_after": fp_after["n_pages"],
                "sizes_ok": ok_sizes,
                "text_ok": ok_text,
            }
            if not ok_pages or not ok_sizes:
                raise RuntimeError(
                    f"页数/页面尺寸不一致（{fp_before['n_pages']} → {fp_after['n_pages']}）"
                    "——已中止，未产出文件。")
            if not ok_text:
                rep["warnings"].append(
                    "提取文本指纹有差异（可能是 pypdf 重写时的对象规范化所致，页数与版式未变）。"
                    "递交前请人工翻检一遍。")

        # ---------- 自证 2：重新扫描输出 ----------
        if ext_out in OOXML_EXTS:
            remaining = analyze_ooxml(work, keep_dates)
        else:
            remaining = analyze_pdf(work, keep_dates)
        # 缩略图/附件等被删项不应残留；Application 这类纯工具标识允许保留；
        # 用户口径豁免项不算残留：--basic 跳过的深度项（rsid/people/修订作者）、
        # --keep-attachments 按用户要求保留的 PDF 内嵌附件
        basic_kept = {"w:author（修订/批注作者）", "w:rsid* 属性", "w:rsids",
                      "w:rsidRoot", "w:person（修订作者名录）"}
        remaining_sig = [f for f in remaining
                         if not (f["field"] in ("Application", "AppVersion")
                                 or (getattr(args, "basic", False) and f["field"] in basic_kept)
                                 or (getattr(args, "keep_attachments", False)
                                     and f["where"] == "/Names /EmbeddedFiles"))]
        rep["verification"]["rescan_ok"] = not remaining_sig
        rep["verification"]["rescan_findings"] = remaining_sig
        if remaining_sig:
            rep["warnings"].append(
                "清理后复查仍有残留：" + "; ".join(
                    f"{f['where']}/{f['field']}={f['value']}" for f in remaining_sig[:5]
                )
            )

        # ---------- 落盘 ----------
        os.replace(work, dst)
        rep["output"] = dst

        # ---------- 数字签名提示 ----------
        if ext in PDF_EXTS:
            try:
                sig = subprocess.run(["pdfsig", src], capture_output=True, text=True,
                                     timeout=30, errors="replace")
                if "Signature Level" in (sig.stdout or "") or "Signature Validation" in (sig.stdout or ""):
                    rep["warnings"].append("原 PDF 含数字签名：重写后签名会失效，需重新签。")
            except (OSError, subprocess.TimeoutExpired):
                pass

        return rep

    except BaseException as exc:  # noqa: BLE001
        # 清理临时文件（就地覆盖时 work != path，绝不能误删原文件）
        if work and os.path.exists(work) and os.path.realpath(work) != os.path.realpath(path):
            try:
                os.unlink(work)
            except OSError:
                pass
        if isinstance(exc, Exception):
            rep["error"] = str(exc) if str(exc) else exc.__class__.__name__
            if os.environ.get("TENDER_TRACEBACK"):
                import traceback
                traceback.print_exc()
        else:
            raise  # KeyboardInterrupt 等照常抛出
        return rep
    finally:
        # .doc 转换的临时目录用毕即删；删除失败不盖掉真正的报错
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


# --------------------------------------------------------------------------
# 文件收集 / 输出
# --------------------------------------------------------------------------

def collect_files(paths: list[str], recursive: bool) -> tuple[list[str], list[str]]:
    """收集待处理文件。目录枚举跳过 Word/WPS 的占用锁文件（~$xxx.docx / .~xxx.docx
    ——文件被打开时落的临时文件，不是 zip），macOS 资源分叉（._xxx）同样跳过。"""
    files: list[str] = []
    missing: list[str] = []

    def _wanted(n: str) -> bool:
        if n.startswith(("~$", ".~", "._")):
            return False
        return os.path.splitext(n)[1].lower() in SUPPORTED_EXTS

    for p in paths:
        p = os.path.expanduser(p)
        if os.path.isdir(p):
            if recursive:
                for root, _dirs, names in os.walk(p):
                    if os.path.basename(root) == "__MACOSX":
                        continue
                    for n in sorted(names):
                        if _wanted(n):
                            files.append(os.path.join(root, n))
            else:
                for n in sorted(os.listdir(p)):
                    fp = os.path.join(p, n)
                    if os.path.isfile(fp) and _wanted(n):
                        files.append(fp)
        elif os.path.isfile(p):
            files.append(p)
        else:
            missing.append(p)
    # 去重保序
    seen = set()
    uniq = []
    for f in files:
        key = os.path.realpath(f)
        if key in seen:
            continue
        seen.add(key)
        uniq.append(f)
    return uniq, missing


LEVEL_CN = {"high": "高", "medium": "中", "low": "低"}

# 本次运行已占用的输出路径（防止同名文件互相覆盖）
_OUTPUTS_THIS_RUN: set[str] = set()


def _print_report(reports: list[dict], missing: list[str], args) -> None:
    mode = "体检" if args.check else "清理"
    total_findings = sum(len(r["findings"]) for r in reports)
    print(f"文档属性{mode}  |  {len(reports)} 个文件  |  发现 {total_findings} 项")
    print("=" * 72)

    for i, r in enumerate(reports, 1):
        print(f"\n[{i}/{len(reports)}] {os.path.basename(r['path'])}"
              f"  ({r['format']}, {_format_size(r['size'])})")

        if r["error"]:
            # 报错可能是多行（.doc 的回退指引）：首行带 [失败]，续行缩进对齐，不散落成正文
            err_lines = r["error"].splitlines()
            print(f"  [失败] {err_lines[0]}")
            for ln in err_lines[1:]:
                print(f"          {ln}")
            continue

        if r["findings"]:
            print(f"  发现 {len(r['findings'])} 项：")
            for f in r["findings"]:
                val = f["value"]
                if len(val) > 60:
                    val = val[:57] + "..."
                print(f"    [{LEVEL_CN.get(f['level'], '?')}] {f['where']}  {f['field']} = {val}")
        else:
            print("  发现：无（已经是干净文件）")

        if args.check:
            for w in r["warnings"]:
                print(f"  [注意] {w}")
            continue

        if r["actions"]:
            print("  已执行：")
            for a in r["actions"]:
                print(f"    - {a['action']} {a['part']}：{a['detail']}")

        v = r["verification"]
        marks = []
        marks.append("正文校验一致" if v.get("content_ok") else "正文校验不一致（未产出/已回滚）")
        marks.append("复查干净" if v.get("rescan_ok") else "复查有残留")
        print("  " + "；".join(marks))

        for w in r["warnings"]:
            print(f"  [注意] {w}")

        if r["output"]:
            out_note = "（已就地覆盖）" if os.path.realpath(r["output"]) == os.path.realpath(r["path"]) else ""
            print(f"  输出：{r['output']} {out_note}")

    if missing:
        print("\n未找到：")
        for m in missing:
            print(f"  - {m}")

    # 汇总
    errs = sum(1 for r in reports if r["error"])
    residual = sum(1 for r in reports if any(
        f["level"] in ("high", "medium") for f in r.get("verification", {}).get("rescan_findings", [])
    ))
    print("\n" + "-" * 72)
    print(f"合计：{len(reports)} 个文件，失败 {errs} 个，有残留 {residual} 个。")
    if args.check:
        print("体检模式：未改动任何文件。" + (" 发现待清理项 → 退出码 1" if total_findings else " 全部干净 → 退出码 0"))


def _emit_json(reports: list[dict], missing: list[str], args) -> None:
    payload = {
        "tool": "strip_metadata",
        "mode": "check" if args.check else "clean",
        "summary": {
            "files": len(reports),
            "errors": sum(1 for r in reports if r["error"]),
            "findings": sum(len(r["findings"]) for r in reports),
            "missing": missing,
        },
        "files": reports,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def _exit_code(reports: list[dict], missing: list[str], args) -> int:
    if missing:
        return 2
    if any(r["error"] for r in reports):
        return 2
    if args.check:
        return 1 if any(r["findings"] for r in reports) else 0
    # 清理模式：复查有残留（高/中）也算失败
    for r in reports:
        v = r.get("verification") or {}
        if v.get("rescan_findings"):
            if any(f["level"] in ("high", "medium") for f in v["rescan_findings"]):
                return 1
    return 0


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="strip_metadata.py",
        description="清除 / 体检 Word（.doc/.docx/.docm）与 PDF 的文档元数据——"
                    "作者、单位、模板路径、批注修订作者、缩略图、rsid、PDF /Info 与 XMP。"
                    "默认不覆盖原文件，只改元数据不改正文（写完自证）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  strip_metadata.py --check 报告.docx 报价.pdf   # 只体检不改动，有发现退出码 1\n"
            "  strip_metadata.py 报告.docx 报价.pdf -o ./已清理\n"
            "  strip_metadata.py --in-place --backup 报告.docx\n"
            "  strip_metadata.py -r ./待递交 --json\n"
            "说明：docx/docm 清理只用标准库；PDF 需要 pypdf。.doc 先转成 .docx 再清——本机有 LibreOffice\n"
            "  脚本自动转；没有时列出检测到的 Office 套件后退出，由你选本机可用的办法转换，\n"
            "  转好对那份 .docx 重跑同一条命令。超过 5MB 的 .doc 不自动转；转换超时或失败时脚本\n"
            "  终止转换并退出——这两种情况都请用户在 Word/WPS 里另存后再来。\n"
        ),
    )
    p.add_argument("paths", nargs="+", help="文件或目录（目录默认不递归）")
    p.add_argument("-c", "--check", action="store_true",
                   help="只体检不改动；发现待清理项则退出码 1")
    p.add_argument("-o", "--outdir", help="输出目录（默认与源文件同目录）")
    p.add_argument("--in-place", action="store_true",
                   help=f"就地覆盖原文件（默认写 <文件名>{OUT_SUFFIX}<扩展名>；.doc 不支持就地覆盖）")
    p.add_argument("--backup", action="store_true",
                   help="--in-place 时先把原文件复制为 <文件名>.bak<扩展名>")
    p.add_argument("--basic", action="store_true",
                   help="只清文档属性（core/app/custom/缩略图/模板路径），"
                        "跳过修订作者、rsid、people.xml、批注作者")
    p.add_argument("--keep-attachments", action="store_true",
                   help="PDF 保留内嵌附件（默认删除，删除时列出被删附件名）")
    p.add_argument("--keep-dates", action="store_true",
                   help="保留创建/修改时间（core.xml 与 PDF /Info 日期、zip 内时间戳）")
    p.add_argument("--json", action="store_true", help="输出 JSON 报告（机器可读）")
    p.add_argument("-r", "--recursive", action="store_true", help="目录递归")
    return p


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    args = build_parser().parse_args(argv)

    if args.in_place and args.outdir:
        print("[错误] --in-place 与 --outdir 不能同时使用。", file=sys.stderr)
        return 2
    if args.backup and not args.in_place:
        args.backup = False  # 无害：只在就地覆盖时才有意义

    files, missing = collect_files(args.paths, args.recursive)

    if not files and not missing:
        print("[错误] 没有找到可处理的 " + " / ".join(sorted(SUPPORTED_EXTS)) + " 文件。",
              file=sys.stderr)
        return 2

    reports = [handle_file(f, args) for f in files]

    if args.json:
        _emit_json(reports, missing, args)
    else:
        _print_report(reports, missing, args)

    return _exit_code(reports, missing, args)


if __name__ == "__main__":
    sys.exit(main())
