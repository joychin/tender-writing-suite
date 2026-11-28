#!/usr/bin/env python3
"""sanitize_smoke.py — 文档脱敏引擎回归电池（跨平台，零第三方依赖）。

夹具运行时合成：docx 用 zipfile 拼（含 core/app/custom/缩略图/批注修订作者/people/
attachedTemplate/rsid 全泄露面），pdf 手写字节（/Info + XMP + /ID + 批注 + 页级
/LastModified）。PDF 用例需要 pypdf——缺了记失败不跳过（同 engine_smoke 对
LibreOffice 的口径：本地看不见的环境维度不能静默放过）。

S01 体检：--check 退出码 1，类目全覆盖（core/app/custom/缩略图/批注作者/people/
    attachedTemplate/rsids），原文件 sha 不变、不产任何输出文件
S02 清理：退出码 0，输出 _cleaned.docx，原文件 sha 不变，正文指纹一致，
    输出复扫 --check 退出码 0
S03 --basic 只清属性：输出里 core 作者没了、rsid 还在（复扫退出码 1 且点名 rsid）
S04 --keep-dates：输出 core.xml 保留 created/modified，作者字段为空
S05 PDF：体检出 /Info、XMP、批注作者、页级 /LastModified；清理后 /Info 与 XMP
    没了、批注 /T /NM /M 没了、页数与文本指纹不变、复扫干净
S06 --json 结构：mode/summary/files 字段齐全，findings/actions 能对上
S07 缺失文件：退出码 2
S08 --in-place --backup：原路径被替换、.bak 与原始字节一致
S09 目录枚举跳过 Word/WPS 占用锁文件（~$x.docx、.~x.docx）与资源分叉（._x.docx）
S10 递归批量 + -o：两个子目录同名文件各得一份输出（-2 去重后缀）
S11 soffice 查找与 docx_toolkit 同口径（_find_soffice 三份拷贝同步的断言）
S12 .doc：假 .doc 走转换失败路径（退出码 2，提示 LibreOffice/docx）；--in-place
    对 .doc 直接拒绝（不碰 soffice，确定性路径）
S13 纯净解释器 --help：python3 -S 启动第一步不碰 pypdf（缺依赖形态）
S14 清理机理抽查：输出 zip 内 core/app/custom 空、缩略图部件消失、settings 无
    attachedTemplate/rsids、正文 rsid 属性清零、w:author 置空（不删属性）、
    people.xml 空名录、w:t 文字原样
S15 正文指纹灵敏度：改动任一 w:t 文本后指纹必变（自证闸有效性的依据）
S16 用户豁免项不判残留：--keep-attachments 清理后退出码 0、附件真被保留（真机回放坐实）
S17 没有 LibreOffice 时把 .doc→.docx 交给模型选办法（用户 2026-11-27 拍板）：探测已装 Office
    （macOS 应用包／Windows 自动化接口注册，模拟）、回退指引点名套件并说清「转好后对 .docx
    重跑同一条命令」、不劝装 LibreOffice、原件不动、临时目录即删
S18 超过 5MB 的 .doc 不自动转换（用户 2026-11-27 定）：有 LibreOffice 也不转、不请模型另想办法，
    直接请用户在 Word/WPS 里另存；原件不动、临时目录即删；正好 5MB 不算超；命令行与 --help
S19 LibreOffice 转换卡死（用户 2026-11-27 定：超时就 kill 掉，再提醒用户）：时限一到就返回、整棵进程树
    杀干净、临时目录即删、提示请用户自己另存（另两处通道见 engine_smoke E80）
"""
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
STRIP = ROOT / "tender-sanitize" / "scripts" / "strip_metadata.py"
TOOLKIT = ROOT / "tender-docx" / "scripts" / "docx_toolkit.py"
sys.path.insert(0, str(STRIP.parent))
import strip_metadata as strip  # noqa: E402

HAS_PYPDF = importlib.util.find_spec("pypdf") is not None

PASS, FAIL = [], []


def check(name, cond, detail=""):
    if cond:
        PASS.append(name)
    else:
        FAIL.append(f"{name}{(' — ' + detail) if detail else ''}")


def run(name, fn):
    """一条用例出错只记它自己失败，后面的用例照跑。"""
    try:
        fn()
    except Exception as exc:  # noqa: BLE001
        FAIL.append(f"{name} — 用例异常 {type(exc).__name__}: {exc}")


def cli(*argv):
    """按模型使用方式走真实命令行，返回 (退出码, stdout, stdout+stderr)。

    stdout 单独返回：JSON 只认它——pypdf 等第三方库的警告走 stderr，
    合并进来会污染 json.loads。诊断信息用第三项。
    """
    r = subprocess.run([sys.executable, str(STRIP), *argv],
                       capture_output=True, text=True, errors="replace", timeout=120)
    return r.returncode, r.stdout, r.stdout + r.stderr


def sha(p: Path) -> str:
    import hashlib
    return hashlib.sha256(p.read_bytes()).hexdigest()


# --------------------------------------------------------------------------
# 合成夹具
# --------------------------------------------------------------------------

_XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _make_docx(path: Path) -> None:
    """一个把 docx 全泄露面都占齐的夹具（无需 Word 能打开，引擎按 zip+XML 读）。"""
    parts = {}
    parts["[Content_Types].xml"] = (
        _XML_HEAD
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + "</Types>")
    parts["_rels/.rels"] = (
        _XML_HEAD
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/thumbnail" Target="docProps/thumbnail.jpeg"/>'
        + '<Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        + "</Relationships>")
    parts["docProps/core.xml"] = (
        _XML_HEAD
        + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        + ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
        + ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        + "<dc:title>XX项目投标-内部讨论稿</dc:title><dc:subject>投标</dc:subject>"
        + "<dc:creator>张三</dc:creator><cp:keywords>机密 内部</cp:keywords>"
        + '<dc:description>generated by 测试</dc:description><cp:lastModifiedBy>李四</cp:lastModifiedBy>'
        + '<cp:revision>7</cp:revision><cp:category>内部归档</cp:category>'
        + '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>'
        + '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-03-04T05:06:07Z</dcterms:modified>'
        + "</cp:coreProperties>")
    parts["docProps/app.xml"] = (
        _XML_HEAD
        + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        + ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        + "<Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion>"
        + "<Company>某某科技有限公司</Company><Manager>李四</Manager>"
        + "<Template>C:\\Users\\zhangsan\\AppData\\Roaming\\Microsoft\\Templates\\公司模板.dotx</Template>"
        + "<TotalTime>42</TotalTime><Pages>30</Pages><Words>9000</Words><HyperlinkBase>\\\\fileserver\\投标\\</HyperlinkBase>"
        + "</Properties>")
    parts["docProps/custom.xml"] = (
        _XML_HEAD
        + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
        + ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        + '<property name="KSOProductBuildVer" vt:lpwstr="11.1.0"/>'
        + '<property name="ICV" vt:lpwstr="A1B2C3D4"/>'
        + '<property name="项目负责人" vt:lpwstr="王五"/>'
        + "</Properties>")
    parts["docProps/thumbnail.jpeg"] = b"\xff\xd8\xff\xe0" + b"\x00" * 64
    parts["word/document.xml"] = (
        _XML_HEAD
        + f'<w:document xmlns:w="{strip_w_ns()}"><w:body>'
        + '<w:p w:rsidR="00A1B2C3" w:rsidP="00A1B2C3" w:rsidRDefault="00A1B2C3">'
        + '<w:r w:rsidRPr="00D4E5F6"><w:t>这是正文第一段，清洗后必须原样保留。</w:t></w:r></w:p>'
        + '<w:p><w:ins w:id="1" w:author="王五" w:date="2026-02-02T10:00:00Z">'
        + "<w:r><w:t>修订插入的一句</w:t></w:r></w:ins></w:p>"
        + "</w:body></w:document>")
    parts["word/comments.xml"] = (
        _XML_HEAD
        + f'<w:comments xmlns:w="{strip_w_ns()}">'
        + '<w:comment w:id="9" w:author="赵六" w:date="2026-02-03T11:00:00Z">'
        + "<w:p><w:r><w:t>批注正文，作者要匿名但内容保留</w:t></w:r></w:p></w:comment>"
        + "</w:comments>")
    parts["word/people.xml"] = (
        _XML_HEAD
        + f'<w:people xmlns:w="{strip_w_ns()}">'
        + '<w:person w:author="王五"/><w:person w:author="赵六"/>'
        + "</w:people>")
    parts["word/settings.xml"] = (
        _XML_HEAD
        + f'<w:settings xmlns:w="{strip_w_ns()}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + '<w:attachedTemplate r:id="rIdTpl"/>'
        + '<w:rsidRoot w:val="00AABBCC"/>'
        + '<w:rsids><w:rsid w:val="00AABBCC"/><w:rsid w:val="00D4E5F6"/></w:rsids>'
        + "</w:settings>")
    parts["word/_rels/settings.xml.rels"] = (
        _XML_HEAD
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rIdTpl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"'
        + ' Target="file:///C:/Users/zhangsan/AppData/Roaming/Microsoft/Templates/公司模板.dotm"/>'
        + "</Relationships>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, text in parts.items():
            if isinstance(text, bytes):
                z.writestr(n, text)
            else:
                z.writestr(n, text.encode("utf-8"))


def strip_w_ns() -> str:
    return "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _make_pdf(path: Path) -> None:
    """极小 PDF：/Info + XMP + /ID + 批注（/T /NM /M）+ 页级 /LastModified + 可提取文本。"""
    content = b"BT /F1 12 Tf 72 720 Td (Hello Tender) Tj ET"
    xmp = (
        "<?xpacket begin='' id='W5M0MpCehiHzreSzNTczkc9d'?>"
        "<x:xmpmeta xmlns:x='adobe:ns:meta/'>"
        "<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>"
        "<rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/'"
        " xmlns:xmp='http://ns.adobe.com/xap/1.0/' xmlns:pdf='http://ns.adobe.com/pdf/1.3/'>"
        "<dc:creator><rdf:Seq><rdf:li>Zhang San</rdf:li></rdf:Seq></dc:creator>"
        "<xmp:CreatorTool>Microsoft Word</xmp:CreatorTool>"
        "<pdf:Producer>TestProducer</pdf:Producer>"
        "</rdf:Description></rdf:RDF></x:xmpmeta>"
        "<?xpacket end='w'?>").encode("utf-8")
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R /Metadata 6 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R"
        b" /Resources << /Font << /F1 5 0 R >> >> /Annots [7 0 R]"
        b" /LastModified (D:20260101000000) >>",
        b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(xmp)).encode() + b" >>\nstream\n" + xmp + b"\nendstream",
        b"<< /Type /Annot /Subtype /Text /Rect [100 700 120 720] /T (Zhang San)"
        b" /NM (N-001) /M (D:20260101) /Contents (note) >>",
        b"<< /Author (Zhang San) /Creator (Microsoft Word) /Producer (TestProducer)"
        b" /Title (Internal Draft) /CreationDate (D:20250101000000) /ModDate (D:20260101000000) >>",
    ]
    buf = io.BytesIO()
    buf.write(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(buf.tell())
        buf.write(f"{i} 0 obj\n".encode() + body + b"\nendobj\n")
    xref_off = buf.tell()
    n = len(objs) + 1
    buf.write(f"xref\n0 {n}\n".encode())
    buf.write(b"0000000000 65535 f \n")
    for off in offsets:
        buf.write(f"{off:010d} 00000 n \n".encode())
    id_hex = "0123456789abcdef0123456789abcdef"
    buf.write(f"trailer\n<< /Size {n} /Root 1 0 R /Info 8 0 R"
              f" /ID [<{id_hex}> <{id_hex}>] >>\n".encode())
    buf.write(f"startxref\n{xref_off}\n%%EOF\n".encode())
    path.write_bytes(buf.getvalue())


# --------------------------------------------------------------------------
# 用例
# --------------------------------------------------------------------------

def _fields(rep_files):
    """把 JSON 报告的 findings 压成 (where, field) 集合，方便类目断言。"""
    out = set()
    for f in rep_files:
        for x in f["findings"]:
            out.add((x["where"], x["field"]))
    return out


def t01(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    before = sha(src)
    code, out, _all = cli("--check", "--json", str(src))
    rep = json.loads(out)
    check("S01a 体检退出码 1", code == 1, f"code={code}")
    got = _fields(rep["files"])
    for where, field in (("docProps/core.xml", "creator"), ("docProps/core.xml", "lastModifiedBy"),
                         ("docProps/app.xml", "Company"), ("docProps/app.xml", "Template"),
                         ("docProps/custom.xml", "KSOProductBuildVer"),
                         ("word/people.xml", "w:person（修订作者名录）"),
                         ("word/settings.xml", "w:attachedTemplate"),
                         ("word/settings.xml", "w:rsids")):
        check(f"S01b 体检点面 {where}/{field}", (where, field) in got, f"got={sorted(got)}")
    check("S01c 批注作者点名",
          any("赵六" in x["value"] for x in rep["files"][0]["findings"]
              if "author" in x["field"]), "")
    check("S01d rsid 属性点名", any("rsid" in f for _w, f in got))
    check("S01e 缩略图点名", any("thumbnail" in w for w, _f in got))
    check("S01f 原文件未动", sha(src) == before)
    check("S01g 体检不落盘", not any(p.name.endswith("_cleaned.docx") for p in work.iterdir()))


def t02(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    before = sha(src)
    fp_before = strip.ooxml_fingerprint(str(src))
    code, out, _all = cli("--json", str(src))
    rep = json.loads(out)
    check("S02a 清理退出码 0", code == 0, f"code={code} out={out[:200]}")
    dst = Path(rep["files"][0]["output"])
    check("S02b 输出 _cleaned.docx", dst.name == "标书_cleaned.docx" and dst.is_file())
    check("S02c 原文件未动", sha(src) == before)
    check("S02d 正文指纹一致", strip.ooxml_fingerprint(str(dst))["text_sha256"]
          == fp_before["text_sha256"])
    check("S02e 复查干净", rep["files"][0]["verification"]["rescan_ok"] is True,
          str(rep["files"][0]["verification"].get("rescan_findings")))
    code2, _, _a2 = cli("--check", str(dst))
    check("S02f 输出复扫退出码 0", code2 == 0, f"code={code2}")


def t03(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    code, out, _all = cli("--basic", "--json", str(src))
    rep = json.loads(out)
    check("S03a basic 清理退出码 0", code == 0, f"code={code}")
    dst = Path(rep["files"][0]["output"])
    code2, out2, _all2 = cli("--check", "--json", str(dst))
    rep2 = json.loads(out2)
    got = _fields(rep2["files"])
    check("S03b basic 后复扫退出码 1", code2 == 1, f"code={code2}")
    check("S03c rsid 还在（basic 不清）", any("rsid" in f for _w, f in got))
    check("S03d core 作者已清", not any(w == "docProps/core.xml" and "creator" in f
                                        for w, f in got))


def t04(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    code, out, _all = cli("--keep-dates", "--json", str(src))
    rep = json.loads(out)
    check("S04a keep-dates 退出码 0", code == 0, f"code={code}")
    dst = Path(rep["files"][0]["output"])
    with zipfile.ZipFile(dst) as z:
        core = z.read("docProps/core.xml").decode("utf-8")
    check("S04b 日期保留", "2026-01-02T03:04:05Z" in core and "2026-03-04T05:06:07Z" in core)
    check("S04c 作者清空", "张三" not in core and "李四" not in core)


def t05(work: Path):
    if not HAS_PYPDF:
        FAIL.append("S05 — 缺 pypdf（PDF 用例记失败不跳过）")
        return
    src = work / "报价.pdf"
    _make_pdf(src)
    before = sha(src)
    code, out, _all = cli("--check", "--json", str(src))
    rep = json.loads(out)
    check("S05a PDF 体检退出码 1", code == 1, f"code={code} out={out[:300]}")
    got = _fields(rep["files"])
    check("S05b /Info 作者", any(f == "/Author" for _w, f in got))
    check("S05c XMP 点名", any("/Metadata" in w for w, _f in got))
    check("S05d 批注作者", any("/T" in f for _w, f in got))
    check("S05e 页级 /LastModified", any("LastModified" in f for _w, f in got))
    fp_before = strip.pdf_fingerprint(str(src))
    code, out, _all = cli("--json", str(src))
    rep = json.loads(out)
    check("S05f PDF 清理退出码 0", code == 0, f"code={code} out={out[:300]}")
    dst = Path(rep["files"][0]["output"])
    check("S05g 原文件未动", sha(src) == before)
    fp_after = strip.pdf_fingerprint(str(dst))
    check("S05h 页数与尺寸不变", fp_after["n_pages"] == fp_before["n_pages"]
          and fp_after["page_sizes"] == fp_before["page_sizes"])
    check("S05i 提取文本不变", fp_after["text_sha256"] == fp_before["text_sha256"],
          f'{fp_before["text_sha256"][:12]}… vs {fp_after["text_sha256"][:12]}…')
    code2, _, _a2 = cli("--check", str(dst))
    check("S05j 输出复扫退出码 0", code2 == 0)
    import pypdf
    r = pypdf.PdfReader(str(dst))
    root = r.trailer["/Root"]
    md = dict(r.metadata) if r.metadata else {}
    check("S05k /Info 已删", not (set(md) - {"/CreationDate", "/ModDate"}), str(md))
    check("S05l XMP 已删", "/Metadata" not in root)
    page = root["/Pages"].get_object()["/Kids"][0].get_object()
    annots = page.get("/Annots")
    a0 = annots[0].get_object() if annots else {}
    check("S05m 批注 /T /NM /M 已删", not any(k in a0 for k in ("/T", "/NM", "/M")))


def t16(work: Path):
    """S16 用户豁免项不判残留：--keep-attachments 清理后退出码 0、附件真被保留
    （真机回放坐实的缺陷：附件按用户口径保留，复扫却把它当残留判失败）。"""
    if not HAS_PYPDF:
        FAIL.append("S16 — 缺 pypdf（PDF 用例记失败不跳过）")
        return
    import pypdf
    src = work / "带附件.pdf"
    _make_pdf(src)
    w = pypdf.PdfWriter(clone_from=pypdf.PdfReader(str(src)))
    w.add_attachment("报价明细.xlsx", b"\x50\x4b\x03\x04" + b"\x00" * 64)
    with open(src, "wb") as fh:
        w.write(fh)
    # 默认：附件删除并退出码 0
    code, out, _all = cli("--json", str(src))
    rep = json.loads(out)
    check("S16a 默认删附件退出码 0", code == 0, f"code={code} out={out[:200]}")
    rd = pypdf.PdfReader(str(rep["files"][0]["output"]))
    check("S16b 默认附件已删", len(list(rd.attachments)) == 0)
    # 保留：附件保留且退出码 0（豁免项不算残留）
    src2 = work / "带附件2.pdf"
    _make_pdf(src2)
    w2 = pypdf.PdfWriter(clone_from=pypdf.PdfReader(str(src2)))
    w2.add_attachment("报价明细.xlsx", b"\x50\x4b\x03\x04" + b"\x00" * 64)
    with open(src2, "wb") as fh:
        w2.write(fh)
    code2, out2, _all2 = cli("--keep-attachments", "--json", str(src2))
    rep2 = json.loads(out2)
    check("S16c keep-attachments 退出码 0", code2 == 0, f"code={code2} out={out2[:200]}")
    check("S16d 附件真被保留", rep2["files"][0]["output"] is not None
          and len(list(pypdf.PdfReader(str(rep2["files"][0]["output"])).attachments)) == 1)


def t06(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    code, out, _all = cli("--json", str(src))
    rep = json.loads(out)
    check("S06a json 字段", all(k in rep for k in ("tool", "mode", "summary", "files")))
    check("S06b summary 对账", rep["summary"]["files"] == 1
          and rep["summary"]["findings"] == len(rep["files"][0]["findings"]))
    check("S06c actions 非空", len(rep["files"][0]["actions"]) > 0)


def t07(work: Path):
    code, out, _all = cli("--json", str(work / "不存在.docx"))
    check("S07 缺失文件退出码 2", code == 2, f"code={code}")


def t08(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    orig = sha(src)
    code, out, _all = cli("--in-place", "--backup", "--json", str(src))
    rep = json.loads(out)
    check("S08a 就地覆盖退出码 0", code == 0, f"code={code} out={out[:200]}")
    check("S08b 输出即原路径", rep["files"][0]["output"] == str(src))
    bak = src.with_suffix(".bak.docx")
    check("S08c 备份与原始一致", bak.is_file() and sha(bak) == orig)
    check("S08d 原路径已是干净件", sha(src) != orig)


def t09(work: Path):
    real = work / "真件.docx"
    _make_docx(real)
    (work / "~$真件.docx").write_bytes(b"not a zip")
    (work / ".~真件.docx").write_bytes(b"not a zip")
    (work / "._真件.docx").write_bytes(b"not a zip")
    files, missing = strip.collect_files([str(work)], recursive=False)
    names = [os.path.basename(f) for f in files]
    check("S09a 只收真件", names == ["真件.docx"], str(names))
    check("S09b 无 missing", missing == [])


def t10(work: Path):
    (work / "a").mkdir()
    (work / "b").mkdir()
    _make_docx(work / "a" / "同名.docx")
    _make_docx(work / "b" / "同名.docx")
    outdir = work / "out"
    code, out, _all = cli("-r", "-o", str(outdir), "--json", str(work))
    rep = json.loads(out)
    outs = sorted(Path(f["output"]).name for f in rep["files"] if f["output"])
    check("S10a 退出码 0", code == 0, f"code={code}")
    check("S10b 两份输出互不覆盖", outs == ["同名_cleaned-2.docx", "同名_cleaned.docx"], str(outs))


def t11(_work=None):
    spec = importlib.util.spec_from_file_location("docx_toolkit_for_smoke", TOOLKIT)
    toolkit = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(toolkit)
    check("S11 _find_soffice 与 docx_toolkit 同口径",
          strip._find_soffice() == toolkit._find_soffice(),
          f'{strip._find_soffice()} vs {toolkit._find_soffice()}')


def t12(work: Path):
    fake = work / "老件.doc"
    fake.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 64)  # OLE 头 + 垃圾体
    code, out, _all = cli("--json", str(fake))
    err = json.loads(out)["files"][0].get("error") or ""
    check("S12a 假 .doc 退出码 2", code == 2, f"code={code} err={err[:120]}")
    check("S12b 报在转换/找不到 LibreOffice",
          ("LibreOffice" in err) or ("docx" in err), err[:160])
    check("S12c 无输出文件", not any(p.suffix == ".docx" for p in work.iterdir()))
    code2, out2, _all2 = cli("--in-place", "--json", str(fake))
    err2 = json.loads(out2)["files"][0].get("error") or ""
    check("S12d .doc 就地覆盖直接拒绝", code2 == 2 and "就地覆盖" in err2, err2[:160])
    # 转换走不通（本机有 LibreOffice 时是转换失败，没有时是找不到）都要给出换办法的出路
    check("S12e 转换走不通时提示给出路（换一种办法 / 另存为 .docx）",
          ("换一种" in err) or ("另存为 .docx" in err), err[:200])
    check("S12f 失败回显不带系统噪音行（macOS 的 soffice 每次先吐一行 Task policy）",
          "Task policy" not in err, err[:200])


def t17(work: Path):
    """S17 没有 LibreOffice 时把「.doc → .docx」交给模型选办法（用户 2026-11-27 拍板：
    LibreOffice 体积大、不少用户装不了，不能是前提）。脚本只内置 LibreOffice 一条自动
    转换（本机唯一实测过的）；没有它就探出已装的 Office 套件、连同「转好后怎么接」
    一起回显后退出，不替模型调用没验证过的自动化。"""
    import types
    from contextlib import redirect_stdout
    from unittest import mock

    # ---- 探测：macOS 按应用包，Windows 按自动化接口的注册（模拟注册表）----
    mac_hits = {"/Applications/wpsoffice.app", "/Applications/Pages.app",
                os.path.join(os.path.expanduser("~"), "Applications", "Pages.app")}
    real_isdir = os.path.isdir
    with mock.patch.object(strip.sys, "platform", "darwin"), \
            mock.patch.object(strip.os.path, "isdir", lambda p: p in mac_hits or real_isdir(p)):
        got_mac = strip._installed_office()
    check("S17a macOS 探测：装了的套件点名、没装的不点名、去重保序",
          got_mac == ["WPS Office", "Pages"], str(got_mac))

    class _Key:
        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    opened = []

    def fake_open(hive, subkey, *a, **k):
        opened.append(subkey)
        if subkey == "KWPS.Application":
            return _Key()
        raise OSError("no such key")

    fake_winreg = types.SimpleNamespace(HKEY_CLASSES_ROOT="HKCR", OpenKey=fake_open)
    with mock.patch.object(strip.sys, "platform", "win32"), \
            mock.patch.dict(sys.modules, {"winreg": fake_winreg}):
        got_win = strip._installed_office()
    check("S17b Windows 探测：按自动化接口注册（模拟注册表，只装 WPS）",
          got_win == ["WPS Office"] and "Word.Application" in opened, f"{got_win} {opened}")
    boom = types.SimpleNamespace(HKEY_CLASSES_ROOT="HKCR",
                                 OpenKey=lambda *a, **k: (_ for _ in ()).throw(RuntimeError("注册表读不了")))
    with mock.patch.object(strip.sys, "platform", "win32"), \
            mock.patch.dict(sys.modules, {"winreg": boom}):
        got_err = strip._installed_office()
    check("S17c 探测自身出错不抛，只是少一条线索", got_err == [], str(got_err))

    # ---- 没有 LibreOffice：把转换交给模型，回退指引要说清怎么办 ----
    fake_doc = work / "无转换器.doc"
    fake_doc.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 64)
    before = fake_doc.read_bytes()

    def run_doc(found, *flags):
        tmp = work / "转换临时目录"
        tmp.mkdir()
        args = strip.build_parser().parse_args([*flags, str(fake_doc)])
        with mock.patch.object(strip, "_find_soffice", lambda: None), \
                mock.patch.object(strip, "_installed_office", lambda: list(found)), \
                mock.patch.object(strip.tempfile, "mkdtemp", lambda **k: str(tmp)):
            rep = strip.handle_file(str(fake_doc), args)
        return rep, args, tmp

    rep, _args, tmp = run_doc(["Microsoft Word", "WPS Office"])
    err = rep.get("error") or ""
    check("S17d 没找到 LibreOffice 时报错点明：.doc 要先转成 .docx、脚本只内置了 LibreOffice",
          ".docx" in err and "LibreOffice" in err, err[:240])
    check("S17e 回退指引点名检测到的套件（模型据此选办法）",
          "Microsoft Word" in err and "WPS Office" in err, err[:240])
    check("S17f 回退指引说清怎么接：对转出来的 .docx 运行同一条命令", "同一条命令" in err, err[:240])
    check("S17g 不再劝人安装 LibreOffice（用户拍板：它不是前提）",
          "装 LibreOffice" not in err and "安装 LibreOffice" not in err, err[:240])
    check("S17h 原 .doc 逐字节未动、没产出任何文件", fake_doc.read_bytes() == before
          and rep.get("output") is None and not any(p.name.endswith("_cleaned.docx") for p in work.iterdir()))
    check("S17i 转换用的临时目录用毕即删", not tmp.exists())
    rep2, args2, tmp2 = run_doc(["Microsoft Word", "WPS Office"], "--check")
    check("S17j 体检模式（--check）同样给回退指引、同样不落盘",
          "同一条命令" in (rep2.get("error") or "") and rep2.get("output") is None and not tmp2.exists())
    rep3, _a3, _t3 = run_doc([])
    err3 = rep3.get("error") or ""
    check("S17k 一件 Office 都没检测到时如实说，并给出请用户另存的兜底",
          "没检测到" in err3 and "另存为 .docx" in err3, err3[:240])
    buf = io.StringIO()
    with redirect_stdout(buf):
        strip._print_report([rep], [], args2)
    raw = buf.getvalue().splitlines()
    at = next((i for i, ln in enumerate(raw) if "[失败]" in ln), None)
    cont = []                      # 失败行之后、下一个空行之前的续行（报告尾部的汇总行不算）
    for ln in raw[at + 1:] if at is not None else []:
        if not ln.strip():
            break
        cont.append(ln)
    check("S17l 文本报告里多行指引的续行都带缩进（不散落成正文）",
          len(cont) >= 2 and all(ln.startswith("          ") for ln in cont), buf.getvalue()[:500])
    r = subprocess.run([sys.executable, "-S", str(STRIP), "--help"],
                       capture_output=True, text=True, errors="replace", timeout=60)
    check("S17m --help 写明没有 LibreOffice 时由模型选办法",
          r.returncode == 0 and "由你选" in (r.stdout + r.stderr), (r.stdout + r.stderr)[-300:])


def t18(work: Path):
    """S18 超过 5MB 的 .doc 不自动转换（用户 2026-11-27 定：超过 5M 的提醒用户自己转换。
    起因是真机上一份 26.7MB 的 .doc 让 LibreOffice 转满 300 秒超时，模型随后拼了一份丢光
    版式的 docx 交回来）：有 LibreOffice 也不转、不请模型另想办法，直接请用户在 Word/WPS
    里另存；原件不动、不建临时目录；正好 5MB 不算超。另两处通道见 engine_smoke E79。"""
    import types
    from contextlib import redirect_stdout
    from unittest import mock

    limit = getattr(strip, "_DOC_AUTO_LIMIT", 5_000_000)
    big = work / "超大.doc"
    with open(big, "wb") as fh:
        fh.write(b"\xd0\xcf\x11\xe0")
        fh.truncate(26_739_022)                    # 稀疏文件：真机那份的大小，不占盘
    before = (big.stat().st_size, open(big, "rb").read(4096))
    ran, made = [], []

    def run_doc(path, *flags):
        """有 LibreOffice 的前提下处理单个文件；soffice 与转换临时目录一旦被用到就记下来
        （临时目录按真实 mkdtemp 的行为：被调用时才建）。"""
        tmp = work / "转换临时目录"

        def fake_mkdtemp(**k):
            made.append(1)
            tmp.mkdir()
            return str(tmp)

        def fake_run(cmd, seconds):                 # 顶替「起 LibreOffice 并等它结束」：记一笔，转不成（没产出文件）
            ran.append(1)
            return types.SimpleNamespace(returncode=1, stdout="Error: source file could not be loaded", stderr="")

        args = strip.build_parser().parse_args([*flags, str(path)])
        with mock.patch.object(strip, "_find_soffice", lambda: "/opt/libreoffice/soffice"), \
                mock.patch.object(strip, "_run_lo", fake_run), \
                mock.patch.object(strip.tempfile, "mkdtemp", fake_mkdtemp):
            rep = strip.handle_file(str(path), args)
        return rep, args, tmp

    rep, args, tmp = run_doc(big)
    err = rep.get("error") or ""
    check("S18a 超过 5MB 的 .doc：有 LibreOffice 也不转（soffice 没被调用）",
          not ran and "超过 5MB" in err and "26.7MB" in err, f"转换调用={len(ran)} {err[:240]}")
    check("S18b 提示请用户在 Word/WPS 里另存为 .docx、说清对那份 .docx 重跑同一条命令，不请模型另想办法",
          "请用户" in err and "另存为 .docx" in err and "同一条命令" in err and "选办法" not in err,
          err[:300])
    check("S18c 原 .doc 未动、没产出文件、没建转换临时目录",
          (big.stat().st_size, open(big, "rb").read(4096)) == before and rep.get("output") is None
          and not made and not tmp.exists()
          and not any(p.name.endswith("_cleaned.docx") for p in work.iterdir()))
    rep2, args2, tmp2 = run_doc(big, "--check")
    check("S18d 体检模式（--check）同样不转、同样给提示、不落盘",
          "超过 5MB" in (rep2.get("error") or "") and rep2.get("output") is None and not made and not ran,
          (rep2.get("error") or "")[:240])
    buf = io.StringIO()
    with redirect_stdout(buf):
        strip._print_report([rep], [], args)
    raw = buf.getvalue().splitlines()
    at = next((i for i, ln in enumerate(raw) if "[失败]" in ln), None)
    cont = []                      # 失败行之后、下一个空行之前的续行
    for ln in raw[at + 1:] if at is not None else []:
        if not ln.strip():
            break
        cont.append(ln)
    check("S18e 文本报告里多行提示的续行都带缩进（不散落成正文）",
          len(cont) >= 2 and all(ln.startswith("          ") for ln in cont), buf.getvalue()[:500])

    # ---- 正好 5MB 不算超：照常去转（压线的假文件转不成，报的是转换失败而不是超限）----
    edge = work / "压线.doc"
    with open(edge, "wb") as fh:
        fh.write(b"\xd0\xcf\x11\xe0")
        fh.truncate(limit)
    ran.clear()
    made.clear()
    rep3, _a3, _t3 = run_doc(edge)
    err3 = rep3.get("error") or ""
    check("S18f 正好 5MB 不算超：照常去转（建了临时目录、调了转换），报的是转换失败而不是超限",
          ran and made and "超过 5MB" not in err3 and "失败" in err3 and "请用户" in err3,
          f"转换调用={len(ran)} 临时目录={len(made)} {err3[:240]}")

    # ---- 命令行（真实入口）与 --help ----
    code, out, _all = cli("--json", str(big))
    err_cli = (json.loads(out)["files"][0].get("error") or "") if out.strip().startswith("{") else _all
    check("S18g 命令行 --json：退出码 2，错误栏是超限提示", code == 2 and "超过 5MB" in err_cli,
          f"code={code} {err_cli[:240]}")
    r = subprocess.run([sys.executable, "-S", str(STRIP), "--help"],
                       capture_output=True, text=True, errors="replace", timeout=60)
    check("S18h --help 写明超过 5MB 的 .doc 不自动转", r.returncode == 0 and "5MB" in (r.stdout + r.stderr),
          (r.stdout + r.stderr)[-300:])


def _hanging_soffice(work: Path, name: str):
    """假的 soffice：自己再拉起一个子进程，然后父子俩都不结束；把两个 pid 写进文件。"""
    import time
    pids, script = work / f"{name}.pids", work / f"{name}.py"
    script.write_text(
        f"#!{sys.executable}\n"
        "import os, subprocess, sys, time\n"
        "child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(3600)'])\n"
        f"open({str(pids)!r}, 'w').write(f'{{os.getpid()}} {{child.pid}}')\n"
        "time.sleep(3600)\n", encoding="utf-8")
    script.chmod(0o755)
    return script, pids


def _alive(pid: int) -> bool:
    if sys.platform == "win32":
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True, errors="replace")
        return str(pid) in (r.stdout or "")
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _wait_dead(pids: list, seconds: float = 8.0) -> list:
    """等这些 pid 都消失；返回到点还活着的（调用方负责收尾杀掉）。"""
    import time
    end = time.time() + seconds
    while True:
        alive = [p for p in pids if _alive(p)]
        if not alive or time.time() > end:
            return alive
        time.sleep(0.2)


def t19(work: Path):
    """S19 LibreOffice 转换卡死时：时限一到就把整棵进程树杀掉、临时目录即删、提示请用户自己另存
    （用户 2026-11-27 定：转换超时就 kill 掉，然后提醒用户）。另两处通道见 engine_smoke E80。"""
    import time
    from unittest import mock
    fake = work / "转不动的小文件.doc"
    with open(fake, "wb") as fh:
        fh.write(b"\xd0\xcf\x11\xe0")
        fh.truncate(4_000_000)                     # 5MB 以内，不会被大小线拦下
    runner, pids = _hanging_soffice(work, "卡死的soffice")
    real_run = subprocess.run

    def clamp(*a, **k):                            # 改前的代码把 300 秒写死：夹到 2 秒
        return real_run(*a, **{**k, "timeout": min(k.get("timeout", 2), 2)})

    args = strip.build_parser().parse_args([str(fake)])
    tmp = work / "转换临时目录"

    def fake_mkdtemp(**k):                         # 转换临时目录落在测试目录里，才看得出「用毕即删」
        tmp.mkdir()
        return str(tmp)

    rep, left, t0 = None, [], time.time()
    try:
        with mock.patch.object(strip, "_find_soffice", lambda: str(runner)), \
                mock.patch.object(strip, "_LO_TIMEOUT", 2, create=True), \
                mock.patch.object(strip.subprocess, "run", clamp), \
                mock.patch.object(strip.tempfile, "mkdtemp", fake_mkdtemp):
            rep = strip.handle_file(str(fake), args)
    finally:
        if pids.is_file():
            left = _wait_dead([int(x) for x in pids.read_text().split()])
            for p in left:
                try:
                    os.kill(p, 9)
                except OSError:
                    pass
    err = (rep or {}).get("error") or ""
    check("S19a 转换卡死时时限一到就返回，报错是「已终止、请用户自己另存」的提示",
          time.time() - t0 < 20 and "超时" in err and "已终止" in err and "请用户" in err
          and "另存为 .docx" in err and "换一种" not in err, f"{err[:300]}")
    check("S19b 转换进程杀干净（父子都不留）", pids.is_file() and not left, f"假 soffice 起没起={pids.is_file()} 残留 {left}")
    check("S19c 原 .doc 未动、没产出文件、没留下转换临时目录",
          fake.stat().st_size == 4_000_000 and (rep or {}).get("output") is None and not tmp.exists(),
          f"临时目录还在={tmp.exists()}")


def t13(_work=None):
    r = subprocess.run([sys.executable, "-S", str(STRIP), "--help"],
                       capture_output=True, text=True, errors="replace", timeout=60)
    check("S13 纯净解释器 --help", r.returncode == 0, (r.stderr or "").strip()[-160:])


def t14(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    code, out, _all = cli("--json", str(src))
    rep = json.loads(out)
    dst = Path(rep["files"][0]["output"])
    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        doc = z.read("word/document.xml").decode("utf-8")
        comments = z.read("word/comments.xml").decode("utf-8")
        settings = z.read("word/settings.xml").decode("utf-8")
        people = z.read("word/people.xml").decode("utf-8")
        core = z.read("docProps/core.xml").decode("utf-8")
        app = z.read("docProps/app.xml").decode("utf-8")
        custom = z.read("docProps/custom.xml").decode("utf-8")
    check("S14a 缩略图部件消失", not any(n.lower().startswith("docprops/thumbnail") for n in names))
    check("S14b core 清空", "张三" not in core and "李四" not in core and "内部讨论稿" not in core)
    check("S14c app 清空", "某某科技" not in app and "zhangsan" not in app
          and "Microsoft Office Word" not in app)
    check("S14d custom 清空", "KSOProductBuildVer" not in custom and "ICV" not in custom)
    check("S14e settings 无模板路径与会话表", "attachedTemplate" not in settings
          and "w:rsids" not in settings and "rsidRoot" not in settings)
    check("S14f 正文 rsid 属性清零", "w:rsid" not in doc)
    check("S14g 修订作者置空不删属性", 'w:author=""' in doc and 'w:author="王五"' not in doc)
    check("S14h 批注作者置空正文保留", 'w:author=""' in comments and "批注正文" in comments)
    check("S14i people 名录清空", "<w:person" not in people)
    check("S14j 正文文字原样", "这是正文第一段，清洗后必须原样保留。" in doc
          and "修订插入的一句" in doc)


def t15(work: Path):
    src = work / "标书.docx"
    _make_docx(src)
    fp1 = strip.ooxml_fingerprint(str(src))
    tampered = work / "改字.docx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tampered, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = data.replace("这是正文第一段".encode("utf-8"), "这是被改过的段".encode("utf-8"))
            zout.writestr(item, data)
    fp2 = strip.ooxml_fingerprint(str(tampered))
    check("S15 改一个 w:t 指纹必变", fp1["text_sha256"] != fp2["text_sha256"])


def main() -> int:
    cases = [t01, t02, t03, t04, t05, t06, t07, t08, t09, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19]
    for i, fn in enumerate(cases, 1):
        with tempfile.TemporaryDirectory(prefix=f"sanitize_smoke_{i:02d}_") as td:
            run(f"S{i:02d}", lambda fn=fn, td=td: fn(Path(td)))
    print(f"sanitize_smoke：{len(PASS)} 断言通过，{len(FAIL)} 失败")
    for f in FAIL:
        print(f"  [失败] {f}")
    if not HAS_PYPDF:
        print("  [环境] 未装 pypdf，S05 按失败记（PDF 维度必须在装了 pypdf 的环境全绿）")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
