#!/usr/bin/env python3
"""docx_toolkit.py — 标书正文 docx 全生命周期工具（建节/读视图/修订/批注/插图/合册）。

子命令：
    create     建节（标题 Heading1 + body 块序列：段落/表格）
    read       文本视图（P 段号 + T 表格号 + 批注提示）
    revise     定向修订（段落 replace/insert_after/delete、表格格 fill）
    comment    加 Word 批注（缺料/待澄清的唯一正规落点）
    image      插入图片（全宽居中）
    render-page PDF 页转 PNG（扫描件贴图用）
    copy-from  把已解析 docx 的行号区间元素级拷进节文件（素材/格式件保真）
    assemble   按目录树合并节文件为整本（章节编号/目录页/小标题降级/批注迁移）
    toc-pages  给合册文件的静态目录条目写入实测页码（需 LibreOffice）

依赖：python-docx（assemble 另需 docxcompose；render-page 需 pypdfium2+Pillow）。
模型不直接改 docx 二进制——内容修改一律走本工具。
"""

from __future__ import annotations

import argparse
import io
import json
import random
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
RESERVED_NAMES = {"写作指引.md", "关键事实与承诺.md"}
SELF_NUMBERED = re.compile(r"^\s*[（(]?[一二三四五六七八九十百\d]+\s*[、.．）)]")
ILLEGAL = re.compile(r'[/\\:*?"<>|\n\r]')


def die(msg: str, code: int = 1):
    print(f"[失败] {msg}")
    sys.exit(code)


def sanitize_name(title: str) -> str:
    name = ILLEGAL.sub(" ", title)
    name = re.sub(r"\s+", " ", name).strip().lstrip(".")
    return name[:60].strip()


def _cn(n: int) -> str:
    digits = "零一二三四五六七八九"
    if n <= 10:
        return ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"][n]
    if n < 20:
        return "十" + (digits[n % 10] if n % 10 else "")
    if n < 100:
        return digits[n // 10] + "十" + (digits[n % 10] if n % 10 else "")
    return str(n)


def set_east_asia(rpr_holder, east: str) -> None:
    rpr = rpr_holder.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), east)


def style_run(run, east="宋体", size=None, bold=None):
    run.font.name = east
    set_east_asia(run._element, east)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


# ---------------------------------------------------------------- 基础版式

def _set_body_indent(p) -> None:
    """正文段 2 字符首行缩进（firstLineChars 而非固定磅，随字号缩放）。"""
    ind = p._p.get_or_add_pPr().get_or_add_ind()
    ind.set(qn("w:firstLineChars"), "200")
    if ind.get(qn("w:firstLine")):
        del ind.attrib[qn("w:firstLine")]


def _clear_first_line_indent(p) -> None:
    """清掉段落继承的首行缩进（表格单元格/标题/题注用）。"""
    p.paragraph_format.first_line_indent = Pt(0)
    ind = p._p.get_or_add_pPr().find(qn("w:ind"))
    if ind is not None and ind.get(qn("w:firstLineChars")):
        del ind.attrib[qn("w:firstLineChars")]


def setup_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"  # 西文/数字按标书惯例用 Times，汉字走 eastAsia 宋体
    normal.font.size = Pt(12)
    set_east_asia(normal.element, "宋体")
    # 注意：首行缩进不放 Normal——表格单元格/题注/目录条目都会继承，
    # 窄列里会退化成一字一行。缩进由 _set_body_indent 按正文段逐段设置。
    normal.paragraph_format.line_spacing = 1.5

    specs = [("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 14), ("Heading 4", 12)]
    for name, size in specs:
        st = doc.styles[name]
        st.font.name = "黑体"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = None
        set_east_asia(st.element, "黑体")

    for name, size in (("Cover Title", 26), ("Cover Sub", 15)):
        try:
            st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        except ValueError:
            st = doc.styles[name]
        st.font.name = "黑体"
        st.font.size = Pt(size)
        st.font.bold = True
        set_east_asia(st.element, "黑体")
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.first_line_indent = Pt(0)

    try:
        cap = doc.styles.add_style("Table Caption", WD_STYLE_TYPE.PARAGRAPH)
    except ValueError:
        cap = doc.styles["Table Caption"]
    cap.font.name = "宋体"
    cap.font.size = Pt(9)
    cap.font.bold = True
    set_east_asia(cap.element, "宋体")
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Pt(0)


DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "base-template.docx"


def _strip_body_content(doc: Document) -> None:
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def _ensure_builtin_styles(doc: Document) -> None:
    """用户模板缺标准标题样式时补建（编号/目录依赖它们）。"""
    for name in ("Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        try:
            doc.styles[name]
        except KeyError:
            try:
                doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
            except Exception:
                pass


def _get_style(doc: Document, name: str):
    try:
        return doc.styles[name]
    except KeyError:
        return None


def _rewrite_para_text(p, new_text: str) -> None:
    """整段换文本，保留首个 run 的字体格式（用户模板的标题样式不被覆盖）。"""
    fmt = None
    if p.runs:
        r0 = p.runs[0]
        fmt = (r0.font.name, r0.font.size, r0.font.bold)
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(new_text)
    if fmt and fmt[0]:
        run.font.name = fmt[0]
    if fmt and fmt[1]:
        run.font.size = fmt[1]
    if fmt and fmt[2] is not None:
        run.font.bold = fmt[2]


def base_document(template: str | None = None) -> tuple[Document, bool]:
    """从模板起建并清空正文（保留样式与页面设置）。

    返回 (doc, user_template)：user_template=True 时不改页面设置与页眉页脚。
    默认用包内基准模板；缺文件时回落程序内置样式。
    """
    user_template = template is not None
    path = Path(template) if template is not None else DEFAULT_TEMPLATE
    if path.is_file():
        doc = Document(str(path))
        _strip_body_content(doc)
        _ensure_builtin_styles(doc)
        return doc, user_template
    if user_template:
        die(f"模板文件不存在：{path}")
    doc = Document()
    setup_styles(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.18)
    return doc, False


def _page_field(paragraph) -> None:
    for tag, attrs, text in (
        ("w:fldChar", {"w:fldCharType": "begin"}, None),
        ("w:instrText", {"xml:space": "preserve"}, " PAGE "),
        ("w:fldChar", {"w:fldCharType": "end"}, None),
    ):
        run = paragraph.add_run()
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(qn(k), v)
        if text is not None:
            el.text = text
        run._element.append(el)
    for run in paragraph.runs:
        style_run(run, "Times New Roman", 9)


def enable_footer_pagenum(doc: Document, different_first: bool) -> None:
    sec = doc.sections[0]
    sec.different_first_page_header_footer = different_first
    footer = sec.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # 幂等：先清空段内既有内容再加域（基准模板已带页码域，重复添加会渲染成 11/22）
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    _page_field(p)
    if different_first:
        fp = sec.first_page_footer
        if fp.paragraphs:
            fp.paragraphs[0].text = ""


# ---------------------------------------------------------------- 批注（原生 Word comments part）

def _comments_part(doc: Document):
    part = doc.part
    for rel in list(part.rels.values()):
        if rel.reltype.endswith("/comments"):
            return rel.target_part
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:comments xmlns:w="{W_NS}"/>').encode("utf-8")
    p = Part(PackURI("/word/comments.xml"), COMMENTS_CT, xml, part.package)
    part.relate_to(p, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
    return p


def _load_comments(part) -> dict:
    try:
        from docx.oxml import parse_xml

        root = parse_xml(part.blob)
        return root
    except Exception:
        from docx.oxml import parse_xml

        return parse_xml(f'<w:comments xmlns:w="{W_NS}"/>')


def add_comment(doc: Document, para, text: str) -> int:
    part = _comments_part(doc)
    root = _load_comments(part)
    cid = random.randint(100000, 999999)
    while any(c.get(qn("w:id")) == str(cid) for c in root.findall(qn("w:comment"))):
        cid = random.randint(100000, 999999)
    c = OxmlElement("w:comment")
    c.set(qn("w:id"), str(cid))
    c.set(qn("w:author"), "灵燕智能")
    c.set(qn("w:initials"), "AI")
    c.set(qn("w:date"), "2026-01-01T00:00:00Z")
    cp = OxmlElement("w:p")
    cr = OxmlElement("w:r")
    ct = OxmlElement("w:t")
    ct.text = text
    cr.append(ct)
    cp.append(cr)
    c.append(cp)
    root.append(c)
    part._blob = _serialize_comments(root)

    start = OxmlElement("w:commentRangeStart")
    start.set(qn("w:id"), str(cid))
    ppr = para._p.find(qn("w:pPr"))
    para._p.insert(list(para._p).index(ppr) + 1 if ppr is not None else 0, start)
    end = OxmlElement("w:commentRangeEnd")
    end.set(qn("w:id"), str(cid))
    para._p.append(end)
    ref_run = OxmlElement("w:r")
    ref = OxmlElement("w:commentReference")
    ref.set(qn("w:id"), str(cid))
    ref_run.append(ref)
    para._p.append(ref_run)
    return cid


def _serialize_comments(root) -> bytes:
    from lxml import etree

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_comments(part) -> dict[str, str]:
    try:
        root = _load_comments(part)
    except Exception:
        return {}
    out = {}
    for c in root.findall(qn("w:comment")):
        cid = c.get(qn("w:id"))
        text = "".join(t.text or "" for t in c.iter(qn("w:t")))
        out[cid] = text
    return out


def doc_comments(doc: Document) -> dict[str, str]:
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/comments"):
            return read_comments(rel.target_part)
    return {}


# ---------------------------------------------------------------- 块枚举

def iter_blocks(doc: Document):
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def para_images(p) -> int:
    try:
        return len(p._p.findall(".//" + qn("a:blip"))) + len(p._p.findall(".//" + qn("w:pict")))
    except Exception:
        return 0


# ---------------------------------------------------------------- create

def cmd_create(args) -> None:
    path = Path(args.path)
    if path.exists() and not args.replace:
        die(f"节文件已存在：{path}（重写请加 --replace）")
    doc, _ = base_document(args.template)
    if getattr(args, "landscape", False):
        sec = doc.sections[0]
        w, h = sec.page_width, sec.page_height
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = max(w, h), min(w, h)
    doc.add_heading(args.title, level=1)
    n_p = n_t = 0
    if args.body:
        blocks = json.loads(Path(args.body).read_text(encoding="utf-8"))
        for blk in blocks:
            if blk.get("type") == "table":
                n_t += 1
                _emit_table(doc, blk)
            else:
                n_p += 1
                p = doc.add_paragraph(blk.get("text", ""))
                st = blk.get("style")
                if st in ("cover-title", "cover-sub"):
                    style_name = "Cover Title" if st == "cover-title" else "Cover Sub"
                    st_obj = _get_style(doc, style_name)
                    if st_obj is not None:
                        p.style = st_obj
                    else:  # 用户模板无封面样式 → 直接格式
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.first_line_indent = Pt(0)
                        for r in p.runs:
                            style_run(r, "黑体", 26 if st == "cover-title" else 15, True)
                elif st in ("h2", "h3", "h4"):
                    _ensure_builtin_styles(doc)
                    st_obj = _get_style(doc, f"Heading {st[1]}")
                    if st_obj is not None:
                        p.style = st_obj
                    p.paragraph_format.first_line_indent = Pt(0)
                elif st:
                    die(f"未知 style：{st}（支持 cover-title / cover-sub / h2 / h3 / h4；"
                        f"节标题走 --title，body 内禁 h1——合册会出双章标题）")
                else:
                    _set_body_indent(p)
        # 尾部不补空段：表格引擎已保证表后有段落（防相邻表格互吞），段落结尾
        # 补空段会在合册分页后顶出空白页（真机实测一册 3 处空白页皆源于此）
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    tag = "已覆盖" if args.replace and path.exists() else "已建节"
    print(f"[{tag}] {path}（段落 {n_p}、表格 {n_t}）")


# ---------------------------------------------------------------- 表格引擎 v2
#
# cell 形态：字符串 | {"t": 值, "span"?: 水平跨列数, "vm"?: "restart"|"continue"}
# （纯字符串 == {"t": s}；span→gridSpan，vm→vMerge 竖向合并）。
# 行内含 dict cell 时进入严格模式：每行 Σspan 必须等于网格宽，vmerge-continue
# 必须对上上一行的 restart 列位，不符即报错不产表。纯字符串表保持宽松（补空格）。
# 招标格式件原表仍走 copy-from 保真拷贝——本引擎只服务自建表。

def _cell_spec(c) -> tuple[str, int, str | None]:
    if isinstance(c, dict):
        t = str(c.get("t", ""))
        try:
            span = int(c.get("span", 1) or 1)
        except (TypeError, ValueError):
            span = 1
        vm = c.get("vm")
        return t, max(1, span), (vm if vm in ("restart", "continue") else None)
    return str(c), 1, None


def _content_width_twips(doc: Document) -> int:
    try:
        sec = doc.sections[0]
        w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
        if w > 2000:
            return w
    except Exception:
        pass
    return 8505  # 15cm 兜底


def _content_col_weights(hspecs, rspecs, grid_w: int) -> list[int] | None:
    """无 widths 时的自动列宽权重：各列最长单元格字数（截到 6-50）。
    合并格（span>1 / vm）不参与——跨列长度无法归属单列。全列无数据返回 None。"""
    lens = [0] * grid_w
    for specs in ([hspecs] if hspecs else []) + list(rspecs):
        pos = 0
        for text, span, vm in specs:
            if span == 1 and vm is None and pos < grid_w:
                lens[pos] = max(lens[pos], len(text.strip()))
            pos += span
    if not any(lens):
        return None
    return [max(6, min(50, n)) for n in lens]


def _grid_col_widths(blk: dict, grid_w: int, content_tw: int,
                     auto_w: list[int] | None = None) -> list[int]:
    weights = blk.get("widths")
    if weights:
        if len(weights) != grid_w:
            raise ValueError(f"widths 数量（{len(weights)}）与表格网格列数（{grid_w}）不符")
        wsum = sum(float(x) for x in weights)
        if wsum <= 0:
            weights = None
        else:
            cols = [max(300, int(content_tw * float(x) / wsum)) for x in weights]
            cols[-1] += content_tw - sum(cols)
            return cols
    if auto_w and len(auto_w) == grid_w and len(set(auto_w)) > 1:
        # 无 widths：按各列内容长度比例分配（长内容列变宽，评委好读）
        wsum = sum(auto_w)
        cols = [max(300, int(content_tw * w / wsum)) for w in auto_w]
        cols[-1] += content_tw - sum(cols)
        return cols
    base = max(300, content_tw // grid_w)
    cols = [base] * grid_w
    cols[-1] += content_tw - sum(cols)
    return cols


def _make_tr(specs: list[tuple[str, int, str | None]], col_w: list[int], *,
             bold: bool, size: float, header: bool, shade: str | None) -> "OxmlElement":
    tr = OxmlElement("w:tr")
    trPr = OxmlElement("w:trPr")
    cant = OxmlElement("w:cantSplit")           # 行不跨页拆分
    trPr.append(cant)
    if header:
        trPr.append(OxmlElement("w:tblHeader"))  # 表头跨页重复
    tr.append(trPr)
    pos = 0
    for text, span, vm in specs:
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        tcW = OxmlElement("w:tcW")               # tcPr 子元素按 schema 序：tcW→gridSpan→vMerge→shd→vAlign
        tcW.set(qn("w:w"), str(sum(col_w[pos:pos + span])))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)
        if span > 1:
            gs = OxmlElement("w:gridSpan")
            gs.set(qn("w:val"), str(span))
            tcPr.append(gs)
        if vm:
            v = OxmlElement("w:vMerge")
            v.set(qn("w:val"), "restart" if vm == "restart" else "continue")
            tcPr.append(v)
        if shade:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), shade)
            tcPr.append(shd)
        va = OxmlElement("w:vAlign")
        va.set(qn("w:val"), "center")
        tcPr.append(va)
        tc.append(tcPr)
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:firstLineChars"), "0")
        ppr.append(ind)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center" if header else "left")
        ppr.append(jc)
        p.append(ppr)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), "Times New Roman")
        rfonts.set(qn("w:eastAsia"), "宋体")
        rpr.append(rfonts)
        if bold or header:
            b = OxmlElement("w:b")
            rpr.append(b)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rpr.append(sz)
        szcs = OxmlElement("w:szCs")
        szcs.set(qn("w:val"), str(int(size * 2)))
        rpr.append(szcs)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text[:200]
        r.append(t)
        p.append(r)
        tc.append(p)
        tr.append(tc)
        pos += span
    return tr


def _emit_table(doc: Document, blk: dict) -> None:
    try:
        _emit_table_inner(doc, blk)
    except ValueError as exc:
        die(str(exc))


def _emit_table_inner(doc: Document, blk: dict) -> None:
    header_raw, rows_raw = blk.get("header", []), blk.get("rows", [])
    if not header_raw and not rows_raw:
        return
    if len(rows_raw) > 60:
        print(f"  ⚠️ 表格行数超限（{len(rows_raw)} 行，上限 60），已按原样生成")
    hspecs = [_cell_spec(c) for c in header_raw]
    rspecs = [[_cell_spec(c) for c in row] for row in rows_raw]
    strict = any(isinstance(c, dict) for row in [*header_raw, *rows_raw] for c in row)
    if strict:
        for spec in hspecs:
            if spec[2]:
                raise ValueError("表头单元格不支持 vm（竖向合并）")
    if hspecs:
        grid_w = sum(s for _, s, _ in hspecs)
    elif rspecs:
        grid_w = sum(s for _, s, _ in rspecs[0])
    else:
        return
    if grid_w > 16:
        raise ValueError(f"表格列数超限（{grid_w} 列，上限 16）——列更多的表请拆表或精简列")
    if strict:
        for rn, specs in enumerate(rspecs, start=2 if hspecs else 1):
            if sum(s for _, s, _ in specs) != grid_w:
                raise ValueError(f"表格第 {rn} 行网格（{sum(s for _, s, _ in specs)} 列）"
                                 f"与表头（{grid_w} 列）不符——核对该行各 cell 的 span")
    else:
        # 宽松模式（纯字符串表）：短行补空格到网格宽，保持旧契约
        hspecs = hspecs + [("", 1, None) for _ in range(grid_w - len(hspecs))]
        rspecs = [row + [("", 1, None)] * (grid_w - len(row)) for row in rspecs]
    # vmerge 列位校验：continue 必须落在上一行的 restart/continue 覆盖区间
    if strict:
        open_cols: set[int] = set()
        for rn, specs in enumerate([*([hspecs] if hspecs else []), *rspecs]):
            new_open: set[int] = set()
            pos = 0
            for _, span, vm in specs:
                cols = set(range(pos, pos + span))
                if vm == "continue" and not cols <= open_cols:
                    raise ValueError(f"表格第 {rn + 1} 行 vMerge continue（第 {pos + 1} 列起）"
                                     f"上方没有对应的合并起始")
                if vm in ("restart", "continue"):
                    new_open |= cols
                pos += span
            open_cols = new_open

    cap = blk.get("caption")
    if cap:
        cap_style = _get_style(doc, "Table Caption")
        if cap_style is not None:
            cp = doc.add_paragraph(cap, style=cap_style)
        else:
            cp = doc.add_paragraph(cap)
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.first_line_indent = Pt(0)
            for r in cp.runs:
                style_run(r, "宋体", 9, True)
        cp.paragraph_format.first_line_indent = Pt(0)

    content_tw = _content_width_twips(doc)
    col_w = _grid_col_widths(blk, grid_w, content_tw,
                             _content_col_weights(hspecs, rspecs, grid_w))
    size = 10.5 if grid_w <= 6 else (9.5 if grid_w <= 8 else 9.0)
    compact = grid_w >= 9

    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")  # 不挂 tblStyle：显式边框自足，不依赖样式表（用户模板可能无 TableGrid）
    tblW = OxmlElement("w:tblW")  # tblPr 子元素按 schema 序：tblW→tblBorders→tblLayout→tblCellMar
    tblW.set(qn("w:w"), str(content_tw))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tblPr.append(borders)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")  # 固定布局，杜绝 Word 自适应挤压
    tblPr.append(layout)
    if compact:
        mar = OxmlElement("w:tblCellMar")  # 宽表紧凑边距（60 twips ≈ 0.11cm）
        for side, v in (("top", "15"), ("left", "60"), ("bottom", "15"), ("right", "60")):
            m = OxmlElement(f"w:{side}")
            m.set(qn("w:w"), v)
            m.set(qn("w:type"), "dxa")
            mar.append(m)
        tblPr.append(mar)
    tbl.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for w in col_w:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    tbl.append(grid)
    if hspecs:
        tbl.append(_make_tr(hspecs, col_w, bold=True, size=size,
                            header=True, shade="D9D9D9"))
    for specs in rspecs:
        tbl.append(_make_tr(specs, col_w, bold=False, size=size,
                            header=False, shade=None))
    _body_insert_before_sectpr(doc, tbl)
    spacer = OxmlElement("w:p")  # 表后空段：防相邻表格互吞
    _body_insert_before_sectpr(doc, spacer)


# ---------------------------------------------------------------- read

def cmd_read(args) -> None:
    doc = Document(args.path)
    comments = doc_comments(doc)
    pi = ti = 0
    for block in iter_blocks(doc):
        if hasattr(block, "rows"):
            ti += 1
            print(f"[T{ti}]")
            prev_row_tcs: list | None = None
            for ri, row in enumerate(block.rows, 1):
                cells: list[str] = []
                tcs: list = []
                prev_tc = None
                for cell in row.cells:
                    tc = cell._tc
                    if tc is prev_tc:
                        cells.append("同左")
                    elif prev_row_tcs is not None and any(tc is t for t in prev_row_tcs):
                        cells.append("同上")
                    else:
                        cells.append(" ".join(cell.text.split())[:60])
                    prev_tc = tc
                    tcs.append(tc)
                prev_row_tcs = tcs
                print(f"  R{ri}：" + "  ".join(f"C{i+1}={c}" for i, c in enumerate(cells)))
            continue
        pi += 1
        imgs = para_images(block)
        cids = [r.get(qn("w:id")) for r in block._p.findall(".//" + qn("w:commentReference"))]
        marks = ""
        if imgs:
            marks += f"〔图×{imgs}〕"
        for cid in cids:
            if cid in comments:
                marks += f"〔批注：{comments[cid]}〕"
        try:
            sname = block.style.name if block.style is not None else ""
        except Exception:
            sname = ""
        tag = f"（{sname}）" if sname.startswith(("Heading", "Cover", "标题")) else ""
        print(f"[P{pi}]{tag}{marks}{block.text}")


# ---------------------------------------------------------------- revise

def _find_para(doc: Document, num: int):
    pi = 0
    for block in iter_blocks(doc):
        if hasattr(block, "rows"):
            continue
        pi += 1
        if pi == num:
            return block
    die(f"段落 P{num} 不存在（read 视图核对段号）")


def _find_table(doc: Document, num: int):
    ti = 0
    for block in iter_blocks(doc):
        if hasattr(block, "rows"):
            ti += 1
            if ti == num:
                return block
    die(f"表格 T{num} 不存在")


def cmd_revise(args) -> None:
    doc = Document(args.path)
    if args.cell:
        m = re.match(r"T(\d+)R(\d+)C(\d+)$", args.cell, re.I)
        if not m:
            die(f"--cell 格式应为 T1R2C3：{args.cell}")
        t, r, c = (int(x) for x in m.groups())
        tbl = _find_table(doc, t)
        try:
            cell = tbl.cell(r - 1, c - 1)
        except Exception:
            die(f"格 {args.cell} 越界（该表 {len(tbl.rows)} 行 × {len(tbl.columns)} 列）")
        cell.text = ""
        run = cell.paragraphs[0].add_run(args.text or "")
        style_run(run, "宋体", 10.5, False)
        doc.save(args.path)
        print(f"[已修订] {args.path}（{args.cell} ← {args.text[:30]}）")
        return
    para = _find_para(doc, args.para)
    if args.mode == "delete":
        para._p.getparent().remove(para._p)
        doc.save(args.path)
        print(f"[已修订] {args.path}（P{args.para} 已删除）")
        print("  ⚠️ 段号已变化：后续 revise/image/comment 前先重新 read 取段号")
        return
    if args.mode == "insert_after":
        new_p = OxmlElement("w:p")
        para._p.addnext(new_p)
        from docx.text.paragraph import Paragraph

        np = Paragraph(new_p, para._parent)
        if args.style in ("cover-title", "cover-sub"):
            style_name = "Cover Title" if args.style == "cover-title" else "Cover Sub"
            st_obj = _get_style(doc, style_name)
            if st_obj is not None:
                np.style = st_obj
            else:
                np.alignment = WD_ALIGN_PARAGRAPH.CENTER
            np.paragraph_format.first_line_indent = Pt(0)
        else:
            _set_body_indent(np)
        run = np.add_run(args.text or "　")
        style_run(run, "黑体" if args.style else "宋体",
                  {"cover-title": 26, "cover-sub": 15}.get(args.style))
        doc.save(args.path)
        print(f"[已修订] {args.path}（P{args.para} 后插入）")
        print("  ⚠️ 段号已变化：后续 revise/image/comment 前先重新 read 取段号")
        return
    # replace
    if not args.find:
        die("replace 需要 --find 旧文本")
    old_text = para.text
    if args.find not in old_text:
        die(f"P{args.para} 中找不到「{args.find}」（read 视图核对原文）")
    new_text = old_text.replace(args.find, args.text or "")
    for r in list(para.runs):
        r._element.getparent().remove(r._element)
    run = para.add_run(new_text)
    style_run(run, "宋体", 12, False)
    doc.save(args.path)
    print(f"[已修订] {args.path}（P{args.para}：{args.find[:20]} → {(args.text or '')[:20]}）")


# ---------------------------------------------------------------- comment / image / render-page

def cmd_comment(args) -> None:
    doc = Document(args.path)
    if args.cell and args.para:
        die("--para 与 --cell 只能选一个")
    if args.cell:
        m = re.match(r"T(\d+)R(\d+)C(\d+)$", args.cell, re.I)
        if not m:
            die(f"--cell 格式应为 T1R2C3：{args.cell}")
        t, r, c = (int(x) for x in m.groups())
        tbl = _find_table(doc, t)
        try:
            cell = tbl.cell(r - 1, c - 1)
        except Exception:
            die(f"格 {args.cell} 越界（该表 {len(tbl.rows)} 行 × {len(tbl.columns)} 列）")
        para = cell.paragraphs[0]
        loc = args.cell
    else:
        if not args.para:
            die("--para P# 与 --cell T#R#C# 二选一必填")
        para = _find_para(doc, args.para)
        loc = f"P{args.para}"
    cid = add_comment(doc, para, args.text)
    doc.save(args.path)
    print(f"[已批注] {args.path}（{loc}，批注 #{cid}）：{args.text[:50]}")


def _normalized_image(img: Path):
    """EXIF 方向归一（手机拍摄件常带 Orientation，不转会横/竖倒置插入）。
    返回可喂 add_picture 的流；环境缺 Pillow 或图无 EXIF 时返回原路径。"""
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return None
    try:
        with Image.open(img) as im:
            if not getattr(im, "getexif", None):
                return None
            exif = im.getexif()
            if not exif or 274 not in exif:  # 274 = Orientation
                return None
            fixed = ImageOps.exif_transpose(im)
            fmt = im.format if im.format in ("PNG", "JPEG", "BMP", "GIF") else "PNG"
            buf = io.BytesIO()
            fixed.save(buf, format=fmt)
            buf.seek(0)
            return buf
    except Exception:
        return None


def cmd_image(args) -> None:
    img = Path(args.img)
    if not img.is_file():
        die(f"图片不存在：{img}")
    doc = Document(args.path)
    anchor = _find_para(doc, args.after_para) if args.after_para else None
    content_cm = 15.0
    try:
        sec = doc.sections[0]
        content_cm = (sec.page_width - sec.left_margin - sec.right_margin).cm
    except Exception:
        pass
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.first_line_indent = Pt(0)
    if anchor is not None:
        anchor._p.addnext(para._p)
    run = para.add_run()
    stream = _normalized_image(img)
    if stream is not None:
        shape = run.add_picture(stream, width=Cm(content_cm))
        note = "，已按 EXIF 纠正方向"
    else:
        shape = run.add_picture(str(img), width=Cm(content_cm))
        note = "，未检测到方向信息"
    if shape.height.cm > 18:
        ratio = 18 / shape.height.cm
        shape.height = Cm(18)
        shape.width = Cm(shape.width.cm * ratio)
    if args.caption:
        cap_para = doc.add_paragraph(args.caption)
        cap_style = _get_style(doc, "Table Caption")
        if cap_style is not None:
            cap_para.style = cap_style
        else:
            cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in cap_para.runs:
                style_run(r, "宋体", 9, True)
        cap_para.paragraph_format.first_line_indent = Pt(0)
        if anchor is not None:
            para._p.addnext(cap_para._p)  # 题注跟在图段后
        note += "，题注已加"
    doc.save(args.path)
    print(f"[已插图] {args.path}（{img.name}，全宽居中{note}）")


def cmd_render_page(args) -> None:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        die("缺少依赖：pypdfium2、Pillow。请先执行：pip install pypdfium2 Pillow", 2)
    pdf = pdfium.PdfDocument(args.pdf)
    if not (1 <= args.page <= len(pdf)):
        die(f"页码越界：{args.page}（该 PDF 共 {len(pdf)} 页）")
    bitmap = pdf[args.page - 1].render(scale=2)
    img = bitmap.to_pil()
    img.save(args.out)
    print(f"[已导出] {args.out}（第 {args.page} 页，{img.width}×{img.height}）")


# ---------------------------------------------------------------- copy-from

def cmd_copy_from(args) -> None:
    src_path = Path(args.source)
    tgt_path = Path(args.target)
    if not src_path.is_file():
        die(f"源文件不存在：{src_path}")
    if not tgt_path.is_file():
        die(f"目标节文件不存在：{tgt_path}（先 create 建节）")
    if src_path.resolve() == tgt_path.resolve():
        die("源与目标是同一个文件")

    parse_dir = Path(args.parse) if args.parse else Path("work/parse") / src_path.name
    elem_path = parse_dir / (src_path.stem + ".elements.json")
    if not elem_path.is_file():
        die(f"源文件未解析或缺少元素映射：{elem_path}\n"
            f"先用 parse_document.py 解析源文件，再用 --parse 指向其解析目录")
    elements = json.loads(elem_path.read_text(encoding="utf-8"))

    if args.lines:
        m = re.match(r"L(\d+)(?:-L?(\d+))?$", args.lines, re.I)
        if not m:
            die(f"--lines 格式应为 L120-L260 或 L120：{args.lines}")
        lo = int(m.group(1))
        hi = int(m.group(2) or m.group(1))
    else:
        lo, hi = 1, 10 ** 9
    selected = [e for e in elements if e[2] >= lo and e[1] <= hi]
    if not selected:
        die(f"区间 L{lo}-L{hi} 未命中任何元素（源文件共 {max((e[2] for e in elements), default=0)} 行，"
            f"核对解析 md 的行号）")

    src_doc = Document(str(src_path))
    tgt_doc = Document(str(tgt_path))
    blocks = list(iter_blocks(src_doc))
    src_stem_map = {s.style_id: s for s in src_doc.styles if getattr(s, "style_id", None)}
    tgt_style_ids = {s.style_id for s in tgt_doc.styles if getattr(s, "style_id", None)}
    body = tgt_doc.element.body
    sectpr = body.find(qn("w:sectPr"))
    copied = 0
    warn_numpr = False
    img_copied = 0
    img_skipped: list[str] = []
    for e in selected:
        if e[0] >= len(blocks):
            continue
        blk = blocks[e[0]]
        el = blk._p if hasattr(blk, "_p") else blk._tbl
        clone, n_img, skipped = _clone_with_images(el, src_doc, tgt_doc)
        img_copied += n_img
        img_skipped.extend(skipped)
        for pstyle in clone.findall(".//" + qn("w:pStyle")):
            sid = pstyle.get(qn("w:val"))
            if sid and sid not in tgt_style_ids and sid in src_stem_map:
                body_styles = tgt_doc.styles.element
                body_styles.append(_deepcopy(src_stem_map[sid].element))
                tgt_style_ids.add(sid)
        if clone.findall(".//" + qn("w:numPr")):
            warn_numpr = True
        if sectpr is not None:
            sectpr.addprevious(clone)
        else:
            body.append(clone)
        copied += 1
    tgt_doc.save(str(tgt_path))
    print(f"[已拷贝] {copied} 个元素 → {tgt_path}（{src_path.name} L{lo}-L{hi}，段落/表格/图片保真）")
    if img_skipped:
        shown = "、".join(img_skipped[:6]) + ("…" if len(img_skipped) > 6 else "")
        print(f"  ⚠️ {len(img_skipped)} 个图形未能拷入（EMF/WMF 等格式不支持，或图形缺少关系 ID）：{shown}")
        print("     处置：用 render-page 把对应 PDF 页转 PNG 后 image 插入，或人工贴图；"
              "该位置已留空，不会产生坏图引用")
    if warn_numpr:
        print("  ⚠️ 拷入内容含 Word 自动编号段落——标题类编号由合册剥除，正文列表编号保留")


def _deepcopy(el):
    import copy

    return copy.deepcopy(el)


V_NS = "urn:schemas-microsoft-com:vml"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _add_image_part(doc, blob: bytes):
    """挂图片到目标文档，兼容 python-docx 0.8 与 1.0+（1.0 起 get_or_add_image_part 改名为
    get_or_add_image）。两代 API 都返回 (rId, Image)。"""
    part = doc.part
    for name in ("get_or_add_image", "get_or_add_image_part"):
        fn = getattr(part, name, None)
        if fn is not None:
            return fn(io.BytesIO(blob))
    raise AttributeError("当前 python-docx 无 get_or_add_image / get_or_add_image_part")


def _drop_shape(node) -> None:
    """图片导不进目标文档时摘掉整个图形容器，避免留下指向不存在关系的坏图。"""
    for holder in node.iterancestors():
        if holder.tag in (qn("w:drawing"), qn("w:pict"), qn("w:object")):
            holder.getparent().remove(holder)
            return


def _clone_with_images(el, src_doc, tgt_doc):
    """深拷贝元素并把图片重挂到目标文档；返回 (clone, 成功图片数, 跳过清单)。
    DrawingML（a:blip）与 VML（v:imagedata）双通道；EMF/WMF 等 python-docx 不认的
    格式按张跳过并记录，不中断整条拷贝。"""
    import copy

    clone = copy.deepcopy(el)
    ok = 0
    skipped: list[str] = []
    for node, attr in [(n, qn("r:embed")) for n in clone.findall(".//" + qn("a:blip"))] + \
                      [(n, R_ID) for n in clone.findall(".//" + "{%s}imagedata" % V_NS)]:
        rid = node.get(attr)
        if not rid:
            # VML/OLE 图形可能只有 src 没有关系 ID——这类元素在合册时会让
            # docxcompose 崩在 rels[None]，拷贝阶段直接摘掉并记录
            skipped.append("图形缺少关系 ID")
            _drop_shape(node)
            continue
        try:
            part = src_doc.part.related_parts[rid]
        except KeyError:
            skipped.append("图片关系缺失")
            _drop_shape(node)
            continue
        try:
            new_rid, _ = _add_image_part(tgt_doc, part.blob)
        except Exception as exc:  # EMF/WMF/SVG 等 python-docx 不支持的位图
            skipped.append(f"{Path(str(part.partname)).name}（{type(exc).__name__}）")
            _drop_shape(node)
            continue
        node.set(attr, new_rid)
        ok += 1
    # SmartArt（dgm:relIds 指向独立的 diagram 部件）无法随元素搬移：整块摘掉并记录，
    # 否则合册时 docxcompose 会在 rels[dm_rid] 上抛 KeyError
    for rel_ids in clone.findall(".//" + "{%s}relIds" % DGM_NS):
        skipped.append("SmartArt 图形")
        _drop_shape(rel_ids)
    return clone, ok, skipped


# ---------------------------------------------------------------- diagram（matplotlib 真图，字体探测回落）

def _cjk_font() -> str | None:
    from matplotlib import font_manager

    known = ["PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "SimHei", "SimSun",
             "Noto Sans CJK SC", "Noto Sans SC", "Source Han Sans SC", "Source Han Sans CN",
             "WenQuanYi Zen Hei", "WenQuanYi Micro Hei", "Heiti SC", "STHeiti", "STSong", "KaiTi"]
    names = {f.name for f in font_manager.fontManager.ttflist}
    for k in known:
        if k in names:
            return k
    for n in sorted(names):
        if any("\u4e00" <= ch <= "\u9fff" for ch in n):
            return n
    return None


def _matplotlib_with_cjk():
    try:
        import matplotlib
    except ImportError:
        die("环境无 matplotlib——本环境图示请用表格表达，或先执行：pip install matplotlib", 2)
    font = _cjk_font()
    if font is None:
        die("未检测到中文字体——真图会显示乱码，本环境图示请改用表格表达")
    matplotlib.rcParams["font.sans-serif"] = [font]
    matplotlib.rcParams["axes.unicode_minus"] = False
    return font


def _render_gantt(spec: dict, out: Path) -> None:
    import matplotlib.pyplot as plt

    tasks = spec.get("tasks", [])
    if not tasks:
        die("gantt spec 缺 tasks")
    fig, ax = plt.subplots(figsize=(10, 0.7 + 0.55 * len(tasks)), dpi=150)
    shades = ["#4d4d4d", "#7f7f7f", "#a6a6a6", "#c9c9c9"]
    for i, t in enumerate(tasks):
        y = len(tasks) - 1 - i
        s, e = float(t.get("start", 0)), float(t.get("end", t.get("start", 0)))
        if e > s:
            ax.broken_barh([(s, e - s)], (y - 0.28, 0.56),
                           facecolors=shades[i % 4], edgecolor="#333333")
        if t.get("milestone"):
            ax.plot((s + e) / 2, y, "D", color="#1a1a1a", markersize=7)
    ax.set_yticks(range(len(tasks)))
    ax.set_yticklabels([str(t.get("name", "")) for t in reversed(tasks)], fontsize=10)
    ax.set_xlabel(spec.get("unit", "周"), fontsize=10)
    ax.set_xlim(left=0)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.tick_params(axis="x", labelsize=9)
    if spec.get("title"):
        ax.set_title(spec["title"], fontsize=13)
    fig.tight_layout()
    fig.savefig(out)
    plt.close(fig)


def _render_layered(spec: dict, out: Path) -> None:
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    layers = spec.get("layers", [])
    if not layers:
        die("layered spec 缺 layers")
    n = len(layers)
    H = n * 1.1 + 0.3
    fig, ax = plt.subplots(figsize=(10, H), dpi=150)
    ax.axis("off")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, H)
    shades = ["#d9d9d9", "#c4c4c4", "#adadad", "#969696"]
    for li, layer in enumerate(layers):
        y0 = H - (li + 1) * 1.1 + 0.1
        ax.add_patch(FancyBboxPatch((0.3, y0), 9.4, 0.9,
                                    boxstyle="round,pad=0.02",
                                    facecolor=shades[min(li, 3)], edgecolor="#595959"))
        ax.text(0.65, y0 + 0.45, str(layer.get("name", "")), ha="left", va="center",
                fontsize=11, fontweight="bold")
        items = [str(x) for x in layer.get("items", [])]
        if items:
            w = 7.4 / len(items)  # items 区落在色带内（0.3–9.7，标签占左侧 2.1）
            for ii, it in enumerate(items):
                x0 = 2.1 + ii * w
                ax.add_patch(FancyBboxPatch((x0 + 0.06, y0 + 0.18), w - 0.12, 0.54,
                                            boxstyle="round,pad=0.02",
                                            facecolor="white", edgecolor="#7f7f7f"))
                ax.text(x0 + w / 2, y0 + 0.45, it, ha="center", va="center", fontsize=9)
        if li < n - 1:
            ax.annotate("", xy=(5, y0 - 0.18), xytext=(5, y0 - 0.02),
                        arrowprops=dict(arrowstyle="-|>", color="#595959", lw=1.2))
    if spec.get("title"):
        ax.set_title(spec["title"], fontsize=13)
    fig.tight_layout()
    fig.savefig(out)
    plt.close(fig)


def cmd_diagram(args) -> None:
    _matplotlib_with_cjk()
    out = Path(args.out)
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    if args.kind == "gantt":
        _render_gantt(spec, out)
        print(f"[已出图] {out}（甘特图，{len(spec.get('tasks', []))} 项任务——"
              f"用 image 子命令插入节文件）")
    else:
        _render_layered(spec, out)
        print(f"[已出图] {out}（分层图，{len(spec.get('layers', []))} 层——"
              f"用 image 子命令插入节文件）")


# ---------------------------------------------------------------- assemble

def parse_outline(md_text: str) -> list[dict]:
    """解析 tender-response-docs.md → [{name, project_info, nodes:[{text,depth,children}]}]"""
    volumes = []
    cur = None
    mode = None
    for raw in md_text.splitlines():
        line = raw.rstrip()
        if line.startswith("# 响应文件："):
            cur = {"name": line[len("# 响应文件："):].strip(), "nodes": []}
            volumes.append(cur)
            mode = None
            continue
        if cur is None:
            continue
        if line.startswith("## "):
            mode = line[3:].strip()
            continue
        if mode == "目录" and line.lstrip().startswith("- "):
            indent = len(line) - len(line.lstrip(" "))
            depth = indent // 2 + 1
            text = line.strip()[2:].strip()
            cur["nodes"].append({"text": text, "depth": depth, "children": []})
    return volumes


def build_tree(nodes: list[dict]) -> list[dict]:
    roots, stack = [], []
    for n in nodes:
        while stack and stack[-1]["depth"] >= n["depth"]:
            stack.pop()
        (stack[-1]["children"] if stack else roots).append(n)
        stack.append(n)
    return roots


def _apply_body_indents(master: Document) -> None:
    """合册收尾：给无显式缩进的正文段补 2 字符首行缩进。

    1.0.6 起缩进在建节时逐段写入；本后处理兜底旧版创建的节文件
    （其正文段靠 Normal 继承缩进，新模板 Normal 已不带）。
    跳过：标题/封面/题注样式、空段、已有显式 w:ind 的段（含 Pt(0) 的
    目录条目；表格单元格不在 doc.paragraphs 里，天然不受影响）。"""
    skip_prefix = ("Heading", "标题", "Cover", "Table Caption")
    for p in master.paragraphs:
        try:
            sname = p.style.name if p.style is not None else ""
        except Exception:
            sname = ""
        if sname.startswith(skip_prefix):
            continue
        if not p.text.strip():
            continue
        ppr = p._p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:ind")) is not None:
            continue
        _set_body_indent(p)


def _clear_table_cell_indents(master: Document) -> None:
    """合册收尾：清掉表格单元格段落继承/残留的首行缩进。

    旧版节文件与 copy-from 拷入的格式表可能带 firstLineChars 残留，
    窄列里会退化成一字一行。doc.paragraphs 不含表格内段落，
    需专门遍历（含嵌套表格）。"""
    stack = list(master.tables)
    while stack:
        tbl = stack.pop()
        for row in tbl.rows:
            for cell in row.cells:
                for cp in cell.paragraphs:
                    _clear_first_line_indent(cp)
                stack.extend(cell.tables)


def _clone_break_sectpr(master: Document, landscape: bool):
    """克隆终节属性做段内分节符：剥 titlePg 与 first/even 引用（default 保留，
    页码跨节连续——titlePg 节后跟普通节在 LibreOffice 下会页码错位，实测矩阵）。
    pgNumType 一并剥除——正文中途的横版区段分节绝不能继承「自 1 重起」。"""
    final = master.element.body.find(qn("w:sectPr"))
    if final is None:
        return None
    sect = _deepcopy(final)
    for ref in sect.findall(qn("w:headerReference")) + sect.findall(qn("w:footerReference")):
        if ref.get(qn("w:type")) in ("first", "even"):
            sect.remove(ref)
    for tag in ("w:titlePg", "w:pgNumType"):
        tp = sect.find(qn(tag))
        if tp is not None:
            sect.remove(tp)
    sz = sect.find(qn("w:pgSz"))
    if sz is not None:
        try:
            w, h = int(sz.get(qn("w:w"))), int(sz.get(qn("w:h")))
        except (TypeError, ValueError):
            return sect
        if landscape:
            sz.set(qn("w:w"), str(max(w, h)))
            sz.set(qn("w:h"), str(min(w, h)))
            sz.set(qn("w:orient"), "landscape")
        else:
            sz.set(qn("w:w"), str(min(w, h)))
            sz.set(qn("w:h"), str(max(w, h)))
            if sz.get(qn("w:orient")):
                del sz.attrib[qn("w:orient")]
    return sect


def _para_with_sect(sect) -> "OxmlElement":
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(sect)
    p.append(ppr)
    return p


def _find_chapter_heading(master: Document, title: str):
    for p in master.paragraphs:
        try:
            sname = p.style.name if p.style is not None else ""
        except Exception:
            sname = ""
        if sname.startswith(("Heading", "标题")) and _chap_match(p.text, title):
            return p
    return None


def _apply_landscape_regions(master: Document, chapter_plan: list, landscape_titles: set) -> None:
    """横向节文件经 docxcompose 合并会丢朝向——对横向章（含连续区段）补段内分节符：
    区段首章标题前插竖版分节（结束前一节），区段末插横版分节（结束横向节）。
    页脚引用随克隆延续，页码跨节连续。"""
    if not landscape_titles:
        return
    flags = [t in landscape_titles for (t, _, _, _) in chapter_plan]
    if not any(flags):
        return
    final_sectpr = master.element.body.find(qn("w:sectPr"))

    def _break_before(el, landscape: bool):
        sect = _clone_break_sectpr(master, landscape)
        if sect is not None:
            el.addprevious(_para_with_sect(sect))

    i = 0
    while i < len(chapter_plan):
        if not flags[i]:
            i += 1
            continue
        j = i
        while j + 1 < len(chapter_plan) and flags[j + 1]:
            j += 1
        head_i = _find_chapter_heading(master, chapter_plan[i][0])
        if head_i is None:
            print(f"  ⚠️ 横向节「{chapter_plan[i][0]}」未找到章节标题，朝向未应用——"
                  f"核对节文件标题与目录逐字一致后重新合册")
            i = j + 1
            continue
        _break_before(head_i._p, landscape=False)  # 结束前一节（竖版）
        if j + 1 < len(chapter_plan):
            head_next = _find_chapter_heading(master, chapter_plan[j + 1][0])
            if head_next is not None:
                _break_before(head_next._p, landscape=True)  # 结束横向区段
                i = j + 1
                continue
            print(f"  ⚠️ 下一章「{chapter_plan[j + 1][0]}」标题未定位，横向区段按到文末处理")
        # 区段到文末：终节就地转横版（不追加空段——那会多出一张竖版尾页）
        if final_sectpr is not None:
            tp = final_sectpr.find(qn("w:titlePg"))
            if tp is not None:
                final_sectpr.remove(tp)
            for ref in final_sectpr.findall(qn("w:headerReference")) + \
                    final_sectpr.findall(qn("w:footerReference")):
                if ref.get(qn("w:type")) in ("first", "even"):
                    final_sectpr.remove(ref)
            sz = final_sectpr.find(qn("w:pgSz"))
            if sz is not None:
                try:
                    w, h = int(sz.get(qn("w:w"))), int(sz.get(qn("w:h")))
                    sz.set(qn("w:w"), str(max(w, h)))
                    sz.set(qn("w:h"), str(min(w, h)))
                    sz.set(qn("w:orient"), "landscape")
                except (TypeError, ValueError):
                    pass
        i = j + 1


_PGNUMTYPE_AFTER = ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                    "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
                    "w:printerSettings", "w:sectPrChange")


def _restart_pagenum_at_one(sect) -> None:
    """正文节页码自 1 重起（w:pgNumType w:start=1，按 CT_SectPr 元素序插入——
    python-docx 未建模该元素，无 API）。幂等。"""
    if sect.find(qn("w:pgNumType")) is not None:
        return
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), "1")
    for tag in _PGNUMTYPE_AFTER:
        succ = sect.find(qn(tag))
        if succ is not None:
            succ.addprevious(pg)
            return
    sect.append(pg)


def _front_section_no_pagenum(master: Document, toc_end, toc_marker) -> bool:
    """前置整节（封面/前置/目录页）零页眉页脚引用、页码不编；正文节（终节）
    页码自 1 重起——行业惯例（2026-11-20 拍板：标书正文第一页=第 1 页）。
    分节符插在目录块结束锚点之后：封面/前置/目录同属前置节，目录页也不带
    页码。零引用前置节 + default 页脚正文节的组合沿自实测矩阵 m5（比 titlePg
    稳：LibreOffice 下 titlePg 节后跟普通节会页码错位 m1/m3）。无目录锚点时
    返回 False（调用方退回 titlePg 方案，此时页码从封面连续编）。"""
    anchor = toc_end if toc_end is not None else toc_marker
    if anchor is None:
        return False
    final = master.element.body.find(qn("w:sectPr"))
    if final is None:
        return False
    sect = _deepcopy(final)
    sz = sect.find(qn("w:pgSz"))
    if sz is not None:
        try:
            w, h = int(sz.get(qn("w:w"))), int(sz.get(qn("w:h")))
            sz.set(qn("w:w"), str(min(w, h)))
            sz.set(qn("w:h"), str(max(w, h)))
            if sz.get(qn("w:orient")):
                del sz.attrib[qn("w:orient")]
        except (TypeError, ValueError):
            pass
    for ref in sect.findall(qn("w:headerReference")) + sect.findall(qn("w:footerReference")):
        sect.remove(ref)
    tp = sect.find(qn("w:titlePg"))
    if tp is not None:
        sect.remove(tp)
    anchor.addnext(_para_with_sect(sect))
    tp = final.find(qn("w:titlePg"))
    if tp is not None:
        final.remove(tp)
    for ref in final.findall(qn("w:headerReference")) + final.findall(qn("w:footerReference")):
        if ref.get(qn("w:type")) in ("first", "even"):
            final.remove(ref)
    # 页码重起落在**正文首节**的 sectPr 上=文档序中紧随前置分节符之后的第一个
    # sectPr。无横版区段时即终节；有横版区段时横版补分节已在正文首节与终节
    # 之间再切节——落在终节会让重起被中途分节吞掉（其后各段仍从封面连续编）
    seen = False
    for sp in master.element.body.iter(qn("w:sectPr")):
        if sp is sect:
            seen = True
            continue
        if seen:
            _restart_pagenum_at_one(sp)
            break
    return True


def _fix_first_page(master: Document) -> None:
    """封面首页不带页码：titlePg + 独立空首页页脚，且删除首页页眉引用。
    两个 LibreOffice 实测坑（Word/WPS 均无此问题）：
    ① 完全空的首页页脚部件会回退默认页脚——页脚里塞一个空格 run；
    ② 存在 headerReference type="first" 时 titlePg 首页页脚被整体忽略
       （docxcompose 合并会引入它）——直接删掉首页页眉引用。"""
    sec = master.sections[0]
    sec.different_first_page_header_footer = True
    sect = sec._sectPr
    for ref in sect.findall(qn("w:headerReference")):
        if ref.get(qn("w:type")) == "first":
            sect.remove(ref)
    fpf = sec.first_page_footer
    if fpf.is_linked_to_previous:
        fpf.is_linked_to_previous = False
    for p in fpf.paragraphs:
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.text = ""
    if fpf.paragraphs:
        fpf.paragraphs[0].add_run(" ")

def _number_scheme_leaf(depth: int, counters: dict, scheme: str) -> str:
    if scheme == "none":
        return ""
    counters[depth] = counters.get(depth, 0) + 1
    for k in [k for k in counters if k > depth]:
        counters[k] = 0
    # 目录树含父节点（容器，不产章节）时，上层计数可能不存在——补 1 防 KeyError
    for k in range(1, depth + 1):
        counters.setdefault(k, 1)
    if scheme == "chapter":
        return f"第{_cn(counters[1])}章　" if depth == 1 else \
            ".".join(str(counters[k]) for k in range(1, depth + 1)) + " "
    if scheme == "gov":
        if depth == 1:
            return f"{_cn(counters[1])}、"
        if depth == 2:
            return f"（{_cn(counters[2])}）"
        return f"{counters[depth]}. "
    # decimal
    return ".".join(str(counters[k]) for k in range(1, depth + 1)) + " "


def cmd_assemble(args) -> None:
    try:
        from docxcompose.composer import Composer
    except ImportError:
        die("缺少依赖：docxcompose。请先执行：pip install docxcompose", 2)
    work = Path(args.work)
    outline_path = work / "outline" / "tender-response-docs.md"
    if not outline_path.is_file():
        die(f"目录文件不存在：{outline_path}（先完成目录树）")
    volumes = parse_outline(outline_path.read_text(encoding="utf-8"))
    if not volumes:
        die("目录文件未解析出任何响应文件（核对「# 响应文件：X」标题与树格式）")
    body_dir = work / "body"
    multi = len(volumes) > 1
    warnings: list[str] = []

    for vol in volumes:
        vol["tree"] = build_tree(vol["nodes"])

    master = None
    composer = None
    toc_entries: list[tuple[str, int, bool]] = []  # (编号+标题, 缩进层级, 是否另附)

    for vol in volumes:
        used_files: set[Path] = set()  # 每册重置——防上一册节文件混进本册批注迁移/孤儿判定
        out_name = f"整本-{sanitize_name(vol['name'])}.docx"
        master, user_tpl = base_document(args.template)
        if not user_tpl:
            enable_footer_pagenum(master, different_first=True)
        composer = Composer(master)
        counters: dict[int, int] = {}
        chapter_plan: list[tuple[str, int, str, bool]] = []  # (title, depth, num, in_front)
        landscape_titles: set[str] = set()  # 本册横向节文件的章标题（合册后补分节重排）
        in_front = True  # 目录节点之前=前置区
        toc_marker = None
        merged = 0

        def _emit_node(node, depth):
            nonlocal in_front, toc_marker, merged
            title = node["text"]
            if title == "封面":
                sf = _section_file(body_dir, vol["name"], title, multi)
                if sf and sf.is_file():
                    composer.append(Document(str(sf)))
                    used_files.add(sf)
                    _strip_first_heading(master, title)
                else:
                    _append_simple_cover(master, vol["name"])
                merged += 1
                return
            if title == "目录":
                in_front = False
                toc_marker = _append_toc_marker(master)
                merged += 1
                return
            num = "" if in_front else _number_scheme_leaf(depth, counters, args.numbering)
            if node["children"]:  # 容器章：占章号、产章标题（无节文件）
                if in_front:
                    # 前置区分组无章标题——不入编号与目录，仅分组提示
                    warnings.append(f"前置区容器「{title}」不占章号，其子节按前置区处理")
                    for child in node["children"]:
                        _emit_node(child, depth + 1)
                    return
                _append_container_heading(master, title, num,
                                          page_break=(depth == 1 and bool(num)))
                chapter_plan.append((title, depth, num, in_front))
                toc_entries.append(((num + title).strip(), depth, False))
                merged += 1
                for child in node["children"]:
                    _emit_node(child, depth + 1)
                return
            sf = _section_file(body_dir, vol["name"], title, multi)
            exists = sf and sf.is_file()
            if exists:
                sub_doc = Document(str(sf))
                try:
                    _sec0 = sub_doc.sections[0]
                    if _sec0.page_width > _sec0.page_height:
                        landscape_titles.add(title)
                except Exception:
                    pass
                composer.append(sub_doc)
                used_files.add(sf)
            else:
                warnings.append(f"缺失节：{title}（{'前置页' if in_front else '章节'}，整本中留空位）")
            if not _apply_chapter_number(master, title, num,
                                         page_break=(depth == 1 and not in_front and bool(num))):
                if exists:
                    warnings.append(f"节文件内标题与目录不一致（目录：{title}）——编号未应用，"
                                    f"改节文件标题与目录逐字一致后重新合册")
            chapter_plan.append((title, depth, num, in_front))
            toc_entries.append(((num + title).strip(), depth, not exists))
            merged += 1

        for top in vol["tree"]:
            _emit_node(top, 1)
        toc_end = None
        if toc_marker is not None:
            toc_end = _fill_toc(master, toc_marker, toc_entries)
        _restyle_all(master, chapter_plan, args.numbering)
        _apply_body_indents(master)
        _clear_table_cell_indents(master)
        _apply_landscape_regions(master, chapter_plan, landscape_titles)
        # 孤儿节文件点名（多册只扫本册子目录，单册扫 body 根）
        if not any(n["text"] == "目录" for n in vol["nodes"]):
            warnings.append("目录树中无「目录」节点——未生成目录页，且全文按前置区处理（无章节编号）")
        scan_dir = body_dir / sanitize_name(vol["name"]) if multi else body_dir
        for f in sorted(scan_dir.glob("*.docx")):
            if f.name.startswith("整本-") or f in used_files:
                continue
            warnings.append(f"目录外文件：{f.as_posix()}（不在目录树中，未并入）")
        # 批注迁移 + 打开自动刷新域（目录出真页码）+ 前置整节不编页码（封面/
        # 前置/目录页），正文节页码自 1 重起（2026-11-20 拍板：行业惯例）
        n_comments = _merge_comments(master, used_files)
        _enable_update_fields(master)
        if not _front_section_no_pagenum(master, toc_end, toc_marker):
            _fix_first_page(master)  # 无目录锚点的退化路径：退回 titlePg 方案
        _shrink_trailing_blank(master)
        out_path = body_dir / out_name
        composer.save(str(out_path))
        print(f"[已合册] {out_path}（合并 {merged} 节，编号 {args.numbering}，"
              f"批注 {n_comments} 条待处理）")
        for w in warnings:
            print(f"  ⚠️ {w}")
        warnings = []
        toc_entries = []


def _shrink_trailing_blank(master: Document) -> None:
    """册尾空段压近零：表格结尾时 Word 规范强制保留一个段落，满行高会在内容
    恰好填满时挤出一张空白尾页（交付观感）。压成 2pt 定行高空段，视觉上消失。"""
    body = master.element.body
    for child in reversed(list(body)):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag != qn("w:p"):
            return  # 最后一个实体元素不是段落——无册尾空段问题
        if any((t.text or "").strip() for t in child.findall(".//" + qn("w:t"))):
            return  # 末段有内容
        if child.findall(".//" + qn("w:drawing")) or child.findall(".//" + qn("w:commentReference")):
            return  # 末段带图/批注引用，不动
        ppr = OxmlElement("w:pPr")  # 整体重建 pPr：spacing→rPr 顺序 schema 安全
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "40")  # 2pt 定值行高
        spacing.set(qn("w:lineRule"), "exact")
        ppr.append(spacing)
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "4")  # 2pt
        rpr.append(sz)
        szcs = OxmlElement("w:szCs")
        szcs.set(qn("w:val"), "4")
        rpr.append(szcs)
        ppr.append(rpr)
        old = child.find(qn("w:pPr"))
        if old is not None:
            child.remove(old)
        child.insert(0, ppr)
        return


def _section_file(body_dir: Path, vol_name: str, title: str, multi: bool) -> Path | None:
    name = sanitize_name(title)
    if multi:
        p = body_dir / sanitize_name(vol_name) / f"{name}.docx"
        return p if p.is_file() else None
    p = body_dir / f"{name}.docx"
    return p if p.is_file() else None


def _append_container_heading(master: Document, title: str, num: str, page_break: bool) -> None:
    """容器章（目录树中有子节、无节文件的分组节点）：产章标题占位。"""
    try:
        p = master.add_paragraph(num + title, style=master.styles["Heading 1"])
    except KeyError:
        p = master.add_paragraph(num + title)
    if page_break:
        p.paragraph_format.page_break_before = True


def _strip_first_heading(master: Document, title: str) -> None:
    for p in master.paragraphs:
        try:
            if p.style.name in ("Heading 1", "标题 1") and p.text.strip() == title:
                p._p.getparent().remove(p._p)
                return
        except Exception:
            continue


def _append_simple_cover(master: Document, vol_name: str) -> None:
    master.add_paragraph("　")
    master.add_paragraph("　")
    p = master.add_paragraph(vol_name)
    cover_style = _get_style(master, "Cover Title")
    if cover_style is not None:
        p.style = cover_style
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            style_run(r, "黑体", 26, True)
    p.paragraph_format.first_line_indent = Pt(0)
    master.add_paragraph("　")
    master.add_paragraph("　")
    master.add_paragraph("　")
    master.add_paragraph("　")


def _body_insert_before_sectpr(master: Document, el) -> None:
    body = master.element.body
    sectpr = body.find(qn("w:sectPr"))
    if sectpr is not None:
        sectpr.addprevious(el)
    else:
        body.append(el)


def _append_toc_marker(master: Document):
    """目录占位标题段（条目在全部节合并后回填）；段前分页——目录不与封面同页。"""
    from docx.text.paragraph import Paragraph

    h = OxmlElement("w:p")
    _body_insert_before_sectpr(master, h)
    hp = Paragraph(h, master)
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.first_line_indent = Pt(0)
    hp.paragraph_format.page_break_before = True
    r = hp.add_run("目录")
    style_run(r, "黑体", 16, True)
    return h


def _enable_update_fields(doc: Document) -> None:
    """settings 加 updateFields：Word/WPS 打开时自动刷新域（目录出真页码）。"""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.insert(0, uf)


def _fill_toc(master: Document, marker, entries: list[tuple[str, int, bool]]):
    """在目录标题段之后插入静态条目（TOC 域包裹；条目带点线制表位——
    未更新域前也有目录观感，Word/WPS 打开经 updateFields 刷新出真页码）。
    返回目录块结束锚点（域 end 段）——合册分节符插在它之后，目录页归
    前置节、不带页码。"""
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.text.paragraph import Paragraph

    if marker is None or not entries:
        return None
    prev = marker

    def _insert_after(el, ref):
        ref.addnext(el)
        return el

    endp = _insert_after(OxmlElement("w:p"), prev)
    endrun = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    endrun.append(fc)
    endp.append(endrun)
    anchor = endp
    for text, depth, missing in reversed(entries):
        p = _insert_after(OxmlElement("w:p"), marker)
        entry = Paragraph(p, master)
        entry.paragraph_format.first_line_indent = Pt(0)
        entry.paragraph_format.left_indent = Cm(0.74 * max(0, depth - 1))
        entry.paragraph_format.tab_stops.add_tab_stop(
            Cm(14.4), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        label = text + ("（另附）" if missing else "") + "\t"
        run = entry.add_run(label)
        style_run(run, "宋体", 12, depth == 1)
        anchor = p
    # TOC 域包裹静态清单：begin+instr+separate 插到第一条目的 pPr 之后
    first = anchor
    base = 1 if first.find(qn("w:pPr")) is not None else 0
    fld = OxmlElement("w:r")
    fldc = OxmlElement("w:fldChar")
    fldc.set(qn("w:fldCharType"), "begin")
    fld.append(fldc)
    instr = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = ' TOC \\o "1-3" \\h \\z \\u '
    instr.append(it)
    sep = OxmlElement("w:r")
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "separate")
    sep.append(fc2)
    first.insert(base, sep)
    first.insert(base, instr)
    first.insert(base, fld)
    return endp


def _heading_para(master: Document, title: str):
    for p in master.paragraphs:
        try:
            if p.style is not None and p.style.name in ("Heading 1", "标题 1") \
                    and p.text.strip() == title.strip():
                return p
        except Exception:
            continue
    return None


def _apply_chapter_number(master: Document, title: str, num: str, page_break: bool) -> bool:
    p = _heading_para(master, title)
    if p is None:
        return False
    if page_break:
        p.paragraph_format.page_break_before = True
    if not num:
        return True
    _rewrite_para_text(p, num + p.text)
    return True


def _chap_match(text: str, title: str) -> bool:
    t = text.strip()
    return t == title.strip() or t.endswith(title.strip())


def _restyle_all(master: Document, chapters: list[tuple[str, int, str, bool]],
                 scheme: str) -> None:
    """单遍重排：depth≥2 章标题降一级；节内小标题按树深度降级 + 接续节号 +
    摘出导航大纲（outlineLvl 9，不进目录/导航窗格）。"""
    idx = 0
    cur: tuple[str, int, bool] | None = None  # (章号, 树深, 是否前置区)
    last_chap = -1
    sub_counters: dict[int, int] = {}
    for p in master.paragraphs:
        try:
            sname = p.style.name if p.style is not None else ""
        except Exception:
            sname = ""
        if sname in ("Heading 1", "标题 1"):
            if idx < len(chapters) and _chap_match(p.text, chapters[idx][0]):
                title, depth, num, in_front = chapters[idx]
                if idx != last_chap:
                    sub_counters = {}
                    last_chap = idx
                idx += 1
                cur = (num, depth, in_front)
                if depth >= 2:
                    try:
                        p.style = master.styles["Heading 2"]
                    except KeyError:
                        pass
                continue
            cur = None
            continue
        if cur is None or sname not in ("Heading 2", "Heading 3", "Heading 4",
                                        "标题 2", "标题 3", "标题 4"):
            continue
        num, depth, in_front = cur
        src_level = int(sname.split()[-1]) if sname[-1].isdigit() else 2
        target = max(2, min(depth + src_level - 1, 9))
        try:
            p.style = master.styles[f"Heading {min(target, 4)}"]
        except KeyError:
            pass
        ppr = p._p.get_or_add_pPr()
        for old_lvl in ppr.findall(qn("w:outlineLvl")):
            ppr.remove(old_lvl)
        lvl = OxmlElement("w:outlineLvl")
        lvl.set(qn("w:val"), "8")
        ppr.append(lvl)
        if in_front or scheme == "none" or target > 4 or depth < 2 \
                or SELF_NUMBERED.match(p.text):
            continue
        sub_counters[target] = sub_counters.get(target, 0) + 1
        for k in [k for k in sub_counters if k > target]:
            sub_counters[k] = 0
        sub = ".".join(str(sub_counters[k]) for k in range(depth + 1, target + 1))
        _rewrite_para_text(p, f"{num.strip()}.{sub} {p.text}" if sub else f"{num.strip()} {p.text}")


def _merge_comments(master: Document, used_files: set[Path]) -> int:
    """docxcompose 迁移正文批注引用；comments part 内容在此合并（id 创建时全局唯一）。"""
    existing = doc_comments(master)
    for f in used_files:
        try:
            others = doc_comments(Document(str(f)))
        except Exception:
            continue
        for cid, text in others.items():
            if cid in existing:
                continue
            existing[cid] = text
    if not existing:
        return 0
    part = _comments_part(master)
    root = _load_comments(part)
    for c in root.findall(qn("w:comment")):
        existing.pop(c.get(qn("w:id")), None)
    for cid, text in existing.items():
        c = OxmlElement("w:comment")
        c.set(qn("w:id"), str(cid))
        c.set(qn("w:author"), "灵燕智能")
        c.set(qn("w:initials"), "AI")
        cp = OxmlElement("w:p")
        cr = OxmlElement("w:r")
        ct = OxmlElement("w:t")
        ct.text = text
        cr.append(ct)
        cp.append(cr)
        c.append(cp)
        root.append(c)
    part._blob = _serialize_comments(root)
    return len(doc_comments(master))


# ---------------------------------------------------------------- toc-pages（静态目录条目补实测页码）

def _find_soffice() -> str | None:
    import shutil
    for cand in (shutil.which("soffice"), shutil.which("libreoffice"),
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 "/usr/bin/soffice", "/opt/homebrew/bin/soffice",
                 "/opt/homebrew/bin/libreoffice"):
        if cand and Path(cand).exists():
            return cand
    return None


def _lo_convert(soffice: str, docx_path: Path, outdir: Path, profile: Path) -> Path:
    import subprocess

    cmd = [soffice, "--headless",
           f"-env:UserInstallation=file://{profile}",
           "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    except subprocess.TimeoutExpired:
        die("LibreOffice 转 PDF 超时（300s）——可跳过本步，交付仍以 Word/WPS 导出为准")
    pdf = outdir / (docx_path.stem + ".pdf")
    if not pdf.is_file():
        die(f"LibreOffice 转 PDF 失败：{(r.stdout or r.stderr).strip()[:200]}")
    return pdf


def _norm_label(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


_TOC_LEADERS = ("dot", "middleDot")
_DOT_LEADER = re.compile(r"\.{5,}")


def _toc_entries(doc) -> list:
    """目录条目段落——**按 TOC 域范围取，不做全文 tab 扫描**。
    真实标书的格式件里既有大量「投 标人：\\t（盖单位章）」这类带 tab 的正文段落，
    也有正文承诺段自带点线制表位；两种误判都会把页码写进正文。
    结构判据：TOC 域 begin→end 之间的段落（assemble 生成的静态条目与 Word/WPS
    刷新域后的真条目都在其中）；文件里没有域（被工具重新保存丢了域）时，退回
    ② TOC 样式族 ③ 点线右对齐制表位 两条弱判据。"""
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    order: list = []
    in_toc: dict = {}
    cur_p = None
    stack: list[dict] = []
    cur_toc = False
    for el in body.iter():
        tag = el.tag
        if tag == qn("w:p"):
            cur_p = el
            order.append(el)
            in_toc[el] = in_toc.get(el, False) or cur_toc
        elif tag == qn("w:fldChar"):
            kind = el.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"type": None})
            elif kind == "end" and stack:
                stack.pop()
            cur_toc = any(f["type"] == "toc" for f in stack)
            if cur_toc and cur_p is not None:
                in_toc[cur_p] = True
        elif tag == qn("w:instrText"):
            if stack and stack[-1]["type"] is None:
                instr = (el.text or "").strip().upper()
                if instr.startswith("TOC"):
                    stack[-1]["type"] = "toc"
                elif instr:
                    stack[-1]["type"] = "other"
            cur_toc = any(f["type"] == "toc" for f in stack)
            if cur_toc and cur_p is not None:
                in_toc[cur_p] = True

    entries = [Paragraph(p, doc) for p in order if in_toc.get(p)]
    entries = [p for p in entries if "\t" in p.text]
    if entries:
        return entries

    for p in doc.paragraphs:  # 兜底：域丢失时的弱判据
        try:
            sname = (p.style.name if p.style is not None else "") or ""
        except Exception:
            sname = ""
        if sname.replace(" ", "").upper().startswith("TOC"):
            entries.append(p)
            continue
        pPr = p._p.pPr
        tabs = pPr.find(qn("w:tabs")) if pPr is not None else None
        if tabs is not None and any(t.get(qn("w:leader")) in _TOC_LEADERS
                                    for t in tabs.findall(qn("w:tab"))):
            if "\t" in p.text:
                entries.append(p)
    return entries


def _entry_label(p) -> str:
    text = p.text
    idx = text.rfind("\t")
    left = text[:idx] if idx >= 0 else text
    return re.sub(r"（另附）$", "", left.strip())


def _page_has_heading(lines: list[str], want: str) -> bool:
    """整行命中优先；标题被排版折行时用相邻 2–3 行拼接再比（长标题跨行是常态）。"""
    for j, l in enumerate(lines):
        if not l:
            continue
        if l == want or (l.startswith(want) and len(l) - len(want) <= 4):
            return True
        for k in (2, 3):
            joined = "".join(x for x in lines[j:j + k] if x)
            if joined == want or joined.startswith(want):
                return True
    return False


def _measure_heading_pages(pdf_path: Path, labels: list[str],
                           restart: bool = False) -> dict[str, int]:
    """按渲染 PDF 实测各目录条目标题所在页。restart=整本带页码重起（前置整节
    不编页码、正文自 1 重起）：条目写正文显示页号=物理页号−前置页数（首个
    非目录页即正文首页，显示为 1）；否则写物理页号（封面计 1，旧口径）。
    从目录页之后开始找（目录页本身含同名条目，不排除会把自己的页码当命中）。"""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(pdf_path))
    pages_lines: list[list[str]] = []
    has_leader: list[bool] = []
    for i in range(len(pdf)):
        txt = pdf[i].get_textpage().get_text_range() or ""
        lines = [l.strip() for l in txt.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        pages_lines.append([_norm_label(l) for l in lines])
        has_leader.append(any(_DOT_LEADER.search(l) for l in lines))
    # 目录可能跨页（末页只剩两条也照样是目录页）：从首个点线页起取连续段
    start = 0
    first = next((i for i, f in enumerate(has_leader) if f), None)
    if first is not None:
        end = first
        while end + 1 < len(has_leader) and has_leader[end + 1]:
            end += 1
        start = end + 1
    off = start if restart else 0  # 前置页数=正文首页的 0 基下标（封面/前置/目录）
    mapping: dict[str, int] = {}
    for label in labels:
        want = _norm_label(label)
        if not want:
            continue
        for i in range(start, len(pages_lines)):
            if _page_has_heading(pages_lines[i], want):
                mapping[label] = i + 1 - off
                break
    pdf.close()
    return mapping


def _has_pagenum_restart(docx_path: Path) -> bool:
    """整本是否带页码重起（正文节 pgNumType start）——决定目录条目写
    正文显示页号还是物理页号；旧版合册文件（连续编页）自动走旧口径。"""
    doc = Document(str(docx_path))
    for sect in doc.element.body.iter(qn("w:sectPr")):
        if sect.find(qn("w:pgNumType")) is not None:
            return True
    return False


def _set_entry_page(p, page: int) -> bool:
    """把条目制表符之后的整段内容重写为页码（不是追加）——重复执行结果恒定。
    Word 刷新域后页码可能落在制表符之后的另一个 run 里，一并清掉。"""
    runs = p.runs
    text = p.text
    idx = text.rfind("\t")
    if idx < 0:
        return False
    pos = 0
    for i, r in enumerate(runs):
        n = len(r.text)
        if pos <= idx < pos + n:
            r.text = r.text[: idx - pos + 1] + str(page)
            for later in runs[i + 1:]:
                later.text = ""
            return True
        pos += n
    return False


def _toc_entry_labels(docx_path: Path) -> list[str]:
    labels: list[str] = []
    for p in _toc_entries(Document(str(docx_path))):
        label = _entry_label(p)
        if label and label not in labels:
            labels.append(label)
    return labels


def _toc_entry_pages(docx_path: Path) -> dict[str, str]:
    pages: dict[str, str] = {}
    for p in _toc_entries(Document(str(docx_path))):
        text = p.text
        idx = text.rfind("\t")
        pages[_entry_label(p)] = text[idx + 1:].strip() if idx >= 0 else ""
    return pages


def _body_signature(docx_path: Path) -> list[str]:
    """目录条目之外的段落文本快照——toc-pages 运行前后必须逐段一致。"""
    doc = Document(str(docx_path))
    skip = {p._p for p in _toc_entries(doc)}
    return [p.text for p in doc.paragraphs if p._p not in skip]


def _write_toc_numbers(docx_path: Path, mapping: dict[str, int]) -> tuple[int, int]:
    """返回 (已改条目数, 制表位写不进去的条目数)。只有真的改了才落盘。"""
    doc = Document(str(docx_path))
    written = failed = 0
    for p in _toc_entries(doc):
        page = mapping.get(_entry_label(p))
        if page is None:
            continue
        text = p.text
        idx = text.rfind("\t")
        if idx >= 0 and text[idx + 1:].strip() == str(page):
            continue  # 已是目标页码，不写盘不计改动
        if _set_entry_page(p, page):
            written += 1
        else:
            failed += 1
    if written:
        doc.save(str(docx_path))
    return written, failed


def cmd_toc_pages(args) -> None:
    docx_path = Path(args.docx)
    if not docx_path.is_file():
        die(f"文件不存在：{docx_path}")
    soffice = args.soffice or _find_soffice()
    if not soffice or not Path(soffice).exists():
        print("[目录页码跳过] 环境未找到 LibreOffice——目录页码由 Word/WPS 打开合册文件时"
              "自动刷新，交付 PDF 用 Word/WPS 导出即可，本步可忽略")
        return
    labels = _toc_entry_labels(docx_path)
    if not labels:
        die("未找到目录条目（先 assemble 生成目录页）")
    original = docx_path.read_bytes()
    body_before = _body_signature(docx_path)

    def _rollback(msg: str) -> None:
        docx_path.write_bytes(original)
        die(msg)

    import tempfile

    restart = _has_pagenum_restart(docx_path)
    if restart:
        print("  页码口径：前置（封面/目录）不编页码，目录条目写正文页号（自 1 起）")
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        profile = tmp / "profile"
        mapping_prev: dict[str, int] | None = None
        for _ in range(3):
            pdf = _lo_convert(soffice, docx_path, tmp, profile)
            mapping = _measure_heading_pages(pdf, labels, restart)
            if mapping == mapping_prev:
                break
            mapping_prev = mapping
            _, failed = _write_toc_numbers(docx_path, mapping)
            if failed:
                _rollback(f"自检失败：{failed} 条目录条目制表位无法重写，已回滚到运行前内容")
        if mapping_prev is None:
            mapping_prev = {}
        # 硬门禁一：正文零改动（目录条目之外的任何段落文本都不得变）
        if _body_signature(docx_path) != body_before:
            _rollback("自检失败：本次运行改动了正文段落（非目录条目），已回滚到运行前内容")
        # 硬门禁二：所有定位到页码的条目都已写入目标值
        pages_now = _toc_entry_pages(docx_path)
        stale = [l for l, pg in mapping_prev.items() if pages_now.get(l) != str(pg)]
        if stale:
            _rollback(f"自检失败：{len(stale)} 条目录条目页码未写入，已回滚"
                      f"（{'、'.join(stale[:3])}）")
        not_found = [l for l in labels if l not in mapping_prev]
        if not_found:
            print(f"  ⚠️ {len(not_found)} 条未定位到正文标题（未写页码）：{'、'.join(not_found[:4])}")
        try:
            import pypdfium2 as pdfium

            print(f"[页数] 整本共 {len(pdfium.PdfDocument(str(pdf)))} 页（{docx_path.name}）")
        except Exception:
            pass
        print(f"[目录页码] 已按渲染实测写入 {len(labels) - len(not_found)}/{len(labels)} 条")
        if getattr(args, "map", None):
            map_path = Path(args.map)
            map_path.parent.mkdir(parents=True, exist_ok=True)
            ordered = {k: mapping_prev[k] for k in
                       sorted(mapping_prev, key=lambda x: (mapping_prev[x], x))}
            map_path.write_text(json.dumps(ordered, ensure_ascii=False, indent=1),
                                encoding="utf-8")
            print(f"[已导出] {map_path}（标题→页码映射，评分索引表等页码回填用）")
        if args.pdf:
            out = Path(args.pdf)
            out.parent.mkdir(parents=True, exist_ok=True)
            _lo_convert(soffice, docx_path, out.parent, profile)
            if out.name != docx_path.stem + ".pdf":
                (out.parent / (docx_path.stem + ".pdf")).replace(out)
            print(f"[已导出] {out}（预览用；交付 PDF 仍以 Word/WPS 导出为准）")


# ---------------------------------------------------------------- CLI

def _para_no(v: str) -> int:
    """段落号参数：接受 3 与 P3 两种写法（read 视图显示 P3，文档示例也用 P3）。"""
    s = str(v).strip()
    if s[:1] in ("P", "p"):
        s = s[1:]
    try:
        n = int(s)
    except ValueError:
        raise argparse.ArgumentTypeError(f"段落号应为数字或 P#（如 3 / P3）：{v}")
    if n < 1:
        raise argparse.ArgumentTypeError(f"段落号从 1 起：{v}")
    return n


def main() -> None:
    ap = argparse.ArgumentParser(description="标书正文 docx 工具族")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sp = sub.add_parser("create", help="建节")
    sp.add_argument("path")
    sp.add_argument("--title", required=True)
    sp.add_argument("--body", help="body 块序列 JSON 文件（段落/表格混排；段落可带 "
                                   "style: cover-title/cover-sub/h2/h3/h4）")
    sp.add_argument("--template", help="版式模板 docx（缺省用包内基准模板）")
    sp.add_argument("--landscape", action="store_true",
                    help="整节横向页（A4 横版）——宽表节（9 列以上团队/资信表）用")
    sp.add_argument("--replace", action="store_true", help="已存在时覆盖重写")
    sp.set_defaults(func=cmd_create)

    sp = sub.add_parser("read", help="文本视图")
    sp.add_argument("path")
    sp.set_defaults(func=cmd_read)

    sp = sub.add_parser("revise", help="定向修订（直接生效，无修订标记）")
    sp.add_argument("path")
    sp.add_argument("--para", type=_para_no, help="段落号 P#（也接受 3）")
    sp.add_argument("--cell", help="表格格 T1R2C3")
    sp.add_argument("--mode", choices=["replace", "insert_after", "delete"], default="replace")
    sp.add_argument("--find", help="replace：要替换的旧文本")
    sp.add_argument("--text", default="")
    sp.add_argument("--style", choices=["cover-title", "cover-sub"], help="insert_after 版式")
    sp.set_defaults(func=cmd_revise)

    sp = sub.add_parser("comment", help="加 Word 批注（--para 段落 / --cell 表格格）")
    sp.add_argument("path")
    sp.add_argument("--para", type=_para_no, default=None, help="段落号 P#（也接受 3）")
    sp.add_argument("--cell", help="表格格 T1R2C3（批注落该格首段）")
    sp.add_argument("--text", required=True)
    sp.set_defaults(func=cmd_comment)

    sp = sub.add_parser("image", help="插入图片（全宽居中）")
    sp.add_argument("path")
    sp.add_argument("--img", required=True)
    sp.add_argument("--after-para", type=_para_no, default=None)
    sp.add_argument("--caption", default=None, help="图下题注（居中 9pt，Table Caption 样式）")
    sp.set_defaults(func=cmd_image)

    sp = sub.add_parser("render-page", help="PDF 页转 PNG")
    sp.add_argument("pdf")
    sp.add_argument("page", type=int)
    sp.add_argument("-o", "--out", required=True)
    sp.set_defaults(func=cmd_render_page)

    sp = sub.add_parser("copy-from", help="已解析 docx 行号区间元素级拷贝")
    sp.add_argument("source")
    sp.add_argument("target")
    sp.add_argument("--lines", help="行号区间 L120-L260（缺省整文件）")
    sp.add_argument("--parse", help="源文件解析目录（缺省 work/parse/<源文件名>）")
    sp.set_defaults(func=cmd_copy_from)

    sp = sub.add_parser("diagram", help="甘特/分层真图 PNG（需 matplotlib+中文字体；无则表格表达）")
    sp.add_argument("out", help="输出 PNG 路径")
    sp.add_argument("--kind", choices=["gantt", "layered"], required=True)
    sp.add_argument("--spec", required=True,
                    help="spec JSON：gantt={title,unit?,tasks:[{name,start,end,milestone}]}（unit 默认 周，start/end 单位随 unit）；"
                         "layered={title,layers:[{name,items:[]}]}（自上而下）")
    sp.set_defaults(func=cmd_diagram)

    sp = sub.add_parser("toc-pages", help="给合册文件的静态目录条目写入实测页码（需 LibreOffice）")
    sp.add_argument("docx", help="合册产物 docx（整本-<册名>.docx）")
    sp.add_argument("--pdf", default=None, help="顺带导出预览 PDF 到该路径（可选）")
    sp.add_argument("--map", default=None,
                    help="导出「标题→页码」JSON 映射（评分索引表等页码回填用）")
    sp.add_argument("--soffice", default=None,
                    help="LibreOffice 可执行文件路径（缺省自动探测；找不到则优雅跳过）")
    sp.set_defaults(func=cmd_toc_pages)

    sp = sub.add_parser("assemble", help="按目录树合册")
    sp.add_argument("--work", default="work")
    sp.add_argument("--numbering", choices=["chapter", "decimal", "gov", "none"],
                    default="chapter")
    sp.add_argument("--template", help="版式模板 docx（缺省用包内基准模板；用户模板的"
                                       "页面设置与页眉页脚原样保留）")
    sp.set_defaults(func=cmd_assemble)

    args = ap.parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except Exception as exc:
        die(f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
