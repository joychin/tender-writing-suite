#!/usr/bin/env python3
r"""material_smoke.py — 公司资料库（material.py）回归电池 M01–M12。

素材复用的登记端：入库=脱敏副本+解析+索引，按主体分目录；检索输出素材列记法；
旧名候选供 check_writing 残留检查自动载入（其他主体/历史项目痕迹）。
所有用例一律把 TENDER_MATERIAL_HOME 指到临时目录——本机真实的登记目录
（~/.lingyan/tender-materials/）不进电池。新增行为先在这里写成失败用例再改 material.py。

M01 入库（docx）：脱敏副本+解析产物+索引齐备，元数据作者清零，索引指纹=库内文件指纹
M02 幂等：同一文件重复入库 → [已登记] 跳过，索引不增
M03 同名不同物：拒绝（退出码 2），指路 forget 或改名
M04 缺 --entity：拒绝并点破「资料按主体登记」（不替用户猜主体）
M05 两主体分目录：各存各的副本与解析产物，list 双主体可见
M06 旧名候选：项目名称/招标编号/采购人标签行的值进索引
M07 search：素材列记法输出；--entity 限定本主体、他主体命中单独点名；未指定主体给提醒
M08 show：章节树带行号 + copy-from 用法（--parse 指库内解析目录）
M09 forget：索引移除、库内副本与解析产物删除、原件不动
M10 目录入库：~$/.~ 锁文件与不支持的后缀跳过且点名
M11 入库（pdf）：脱敏副本进库记图源条目，不产解析产物（缺 pypdf 记失败不跳过）
M12 索引损坏：按空库重建并提醒，不炸
"""
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
MP = [sys.executable, "-B", str(ROOT / "tender-analysis" / "scripts" / "material.py")]

PASS, FAIL = [], []
WORK = Path(tempfile.mkdtemp(prefix="material-smoke-"))
LIB = WORK / "资料库"


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(f"{name}{(' — ' + detail) if detail and not cond else ''}")


def run(name, fn):
    """一条用例出错只记它自己失败，后面的用例照跑。"""
    try:
        fn()
    except Exception as exc:  # noqa: BLE001
        FAIL.append(f"{name} — 用例异常 {type(exc).__name__}: {exc}")


def mat(*a):
    e = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", TENDER_MATERIAL_HOME=str(LIB))
    r = subprocess.run(MP + [str(x) for x in a], capture_output=True, text=True, errors="replace", env=e)
    return r.returncode, (r.stdout + r.stderr).strip()


def _sha(p) -> str:
    h = hashlib.sha256()
    h.update(Path(p).read_bytes())
    return h.hexdigest()


def _idx() -> dict:
    return json.loads((LIB / "索引.json").read_text(encoding="utf-8"))


def _doc_fixture(path: Path):
    d = Document()
    cp = d.core_properties
    cp.author = "张三"
    cp.last_modified_by = "张三"
    d.add_paragraph("项目名称：云澜市智慧园区建设工程")
    d.add_paragraph("招标编号：YLCG-2026-GK-0387")
    d.add_paragraph("采购人：云澜市大数据管理局")
    d.add_heading("售后服务方案", level=1)
    d.add_paragraph("故障响应时间为 2 小时，提供 7×24 小时热线支持。")
    d.add_heading("培训方案", level=1)
    d.add_paragraph("提供两场现场培训，每场半天，覆盖平台操作与日常运维。")
    d.save(str(path))


def _doc_fixture2(path: Path):
    d = Document()
    d.add_paragraph("星灵创科售后承诺书")
    d.add_heading("服务承诺", level=1)
    d.add_paragraph("一般故障 4 小时内到场。")
    d.save(str(path))


def _pdf_fixture(path: Path):
    xmp = (b"<?xpacket begin='' id='W5M0MpCehiHzreSzNTczkc9d'?>"
           b"<x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>"
           b"<rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/'>"
           b"<dc:creator><rdf:Seq><rdf:li>Zhang San</rdf:li></rdf:Seq></dc:creator>"
           b"</rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end='w'?>")
    content = b"BT /F1 12 Tf 72 720 Td (Certificate) Tj ET"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
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


def m01():
    src = WORK / "公司简介.docx"
    _doc_fixture(src)
    code, out = mat("add", src, "--entity", "奥哲")
    idx = _idx()
    ents = [m for m in idx["materials"] if m["entity"] == "奥哲"]
    stored = LIB / "奥哲" / "公司简介.docx"
    pdir = LIB / "奥哲" / "解析" / "公司简介.docx"
    core_xml = zipfile.ZipFile(stored).read("docProps/core.xml").decode("utf-8") if stored.is_file() else ""
    check("M01 入库：脱敏副本+解析产物+索引齐备，元数据作者清零，索引指纹=库内文件指纹",
          code == 0 and len(ents) == 1 and ents[0]["cleaned"] is True and ents[0]["kind"] == "docx"
          and stored.is_file() and "张三" not in core_xml and ents[0]["sha256"] == _sha(stored)
          and (pdir / "公司简介.meta.json").is_file()
          and (pdir / "公司简介.outline.json").is_file()
          and (pdir / "公司简介.elements.json").is_file(),
          f"exit={code} out={out[:180]}")


def m02():
    code, out = mat("add", WORK / "公司简介.docx", "--entity", "奥哲")
    n = len([m for m in _idx()["materials"] if m["entity"] == "奥哲"])
    check("M02 幂等：同一文件重复入库 → [已登记] 跳过，索引不增",
          code == 0 and "已登记" in out and n == 1, f"exit={code} 条目数={n} out={out[:140]}")


def m03():
    src = WORK / "公司简介.docx"
    d = Document(str(src))
    d.add_paragraph("改过的一行，内容与上次入库不同。")
    d.save(str(src))
    code, out = mat("add", src, "--entity", "奥哲")
    n = len([m for m in _idx()["materials"] if m["entity"] == "奥哲"])
    check("M03 同名不同物：拒绝（退出码 2）并指路 forget 或改名",
          code == 2 and "同名" in out and "forget" in out and n == 1,
          f"exit={code} 条目数={n} out={out[:160]}")


def m04():
    code, out = mat("add", WORK / "公司简介.docx")
    check("M04 缺 --entity：拒绝（退出码 2）并点破「资料按主体登记」",
          code == 2 and "主体" in out, f"exit={code} out={out[:140]}")


def m05():
    src2 = WORK / "售后承诺.docx"
    _doc_fixture2(src2)
    code, out = mat("add", src2, "--entity", "星灵")
    ents = _idx()["materials"]
    ekeys = sorted({m["entity"] for m in ents})
    code2, out2 = mat("list")
    check("M05 两主体分目录：条目各归各主体，list 双主体可见",
          code == 0 and ekeys == ["奥哲", "星灵"] and (LIB / "星灵" / "售后承诺.docx").is_file()
          and "奥哲" in out2 and "星灵" in out2,
          f"exit={code} 主体={ekeys} list片段={out2[:120]}")


def m06():
    ents = {m["file"]: m for m in _idx()["materials"] if m["entity"] == "奥哲"}
    names = ents["公司简介.docx"].get("old_names") or []
    check("M06 旧名候选：项目名称/招标编号/采购人标签行的值进索引",
          "云澜市智慧园区建设工程" in names and "YLCG-2026-GK-0387" in names
          and "云澜市大数据管理局" in names, f"old_names={names}")


def m07():
    code, out = mat("search", "培训", "--entity", "奥哲")
    ok_primary = code == 0 and "公司简介.docx L" in out and "培训方案" in out
    code2, out2 = mat("search", "培训", "--entity", "星灵")
    ok_other = code2 == 0 and "他主体" in out2 and "奥哲" in out2 and "公司简介.docx" in out2
    code3, out3 = mat("search", "培训")
    check("M07 search：素材列记法输出；--entity 限定本主体、他主体命中单独点名；未指定主体给提醒",
          ok_primary and ok_other and code3 == 0 and "未指定主体" in out3,
          f"本主体exit={code} out={out[:140]} | 他主体exit={code2} out2={out2[:140]}")


def m08():
    code, out = mat("show", "公司简介.docx", "--entity", "奥哲")
    check("M08 show：章节树带行号 + copy-from 用法（--parse 指库内解析目录）",
          code == 0 and "L" in out and "copy-from" in out and "--parse" in out and "旧名" in out,
          f"exit={code} out={out[:180]}")


def m09():
    src = WORK / "公司简介.docx"
    before = len([m for m in _idx()["materials"] if m["entity"] == "奥哲"])
    code, out = mat("forget", "公司简介.docx", "--entity", "奥哲")
    after = len([m for m in _idx()["materials"] if m["entity"] == "奥哲"])
    check("M09 forget：索引移除、库内副本与解析产物删除、原件不动",
          code == 0 and before == 1 and after == 0
          and not (LIB / "奥哲" / "公司简介.docx").exists()
          and not (LIB / "奥哲" / "解析" / "公司简介.docx").exists()
          and src.is_file(),
          f"exit={code} before={before} after={after}")


def m10():
    d = WORK / "丢进来的"
    d.mkdir(exist_ok=True)
    _doc_fixture(d / "另一份.docx")
    (d / "~$另一份.docx").write_bytes(b"lock")
    (d / "说明.txt").write_text("说明", encoding="utf-8")
    code, out = mat("add", d, "--entity", "奥哲")
    ents = [m["file"] for m in _idx()["materials"] if m["entity"] == "奥哲"]
    check("M10 目录入库：docx 收、~$ 锁文件与 txt 跳过且点名",
          code == 0 and ents == ["另一份.docx"] and "说明.txt" in out and "~$" in out,
          f"exit={code} 条目={ents} out={out[:180]}")


def m11():
    try:
        import pypdf  # noqa: F401
    except Exception:
        check("M11 入库（pdf）：脱敏副本进库记图源条目，不产解析产物", False, "缺 pypdf——按电池口径记失败不跳过")
        return
    src = WORK / "资质证书.pdf"
    _pdf_fixture(src)
    code, out = mat("add", src, "--entity", "奥哲")
    ents = [m for m in _idx()["materials"] if m["entity"] == "奥哲" and m["file"] == "资质证书.pdf"]
    stored = LIB / "奥哲" / "资质证书.pdf"
    check("M11 入库（pdf）：脱敏副本进库记图源条目，不产解析产物",
          code == 0 and len(ents) == 1 and ents[0]["kind"] == "pdf" and stored.is_file()
          and not (LIB / "奥哲" / "解析" / "资质证书.pdf").exists()
          and "图源" in (ents[0].get("note") or ""),
          f"exit={code} out={out[:160]}")


def m12():
    (LIB / "索引.json").write_text("这不是 JSON", encoding="utf-8")
    src = WORK / "丢进来的" / "第三份.docx"
    _doc_fixture(src)
    code, out = mat("add", src, "--entity", "奥哲")
    idx_ok = isinstance(_idx().get("materials"), list)
    check("M12 索引损坏：按空库重建并提醒，不炸",
          code == 0 and "重建" in out and idx_ok
          and any(m["file"] == "第三份.docx" for m in _idx()["materials"]),
          f"exit={code} out={out[:160]}")


if __name__ == "__main__":
    for name, fn in (("M01", m01), ("M02", m02), ("M03", m03), ("M04", m04),
                     ("M05", m05), ("M06", m06), ("M07", m07), ("M08", m08),
                     ("M09", m09), ("M10", m10), ("M11", m11), ("M12", m12)):
        run(name, fn)
    print(f"\n{len(PASS)} 通过 / {len(FAIL)} 失败")
    for f_ in FAIL:
        print("  FAIL:", f_)
    if FAIL:
        print(f"  现场保留：{WORK}")
    else:
        shutil.rmtree(WORK, ignore_errors=True)
    sys.exit(1 if FAIL else 0)
