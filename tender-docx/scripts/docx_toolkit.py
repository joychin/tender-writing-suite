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
import os
import random
import re
import sys
from pathlib import Path

# 第三方依赖惰性加载：顶层不 import python-docx，缺依赖时本模块仍可导入、
# 八个脚本的 --help 仍可用；真正要用 Word 能力时才由 require_docx() 加载并给出
# 可执行的中文提示。顶层 import 会让未装依赖的环境（沙箱评测、纯净解释器、
# 别人的电脑）在启动第一步就 ModuleNotFoundError 崩掉。
Document = None  # type: ignore[assignment]
WD_ORIENT = WD_STYLE_TYPE = WD_ALIGN_PARAGRAPH = None
OxmlElement = qn = None
Cm = Pt = None
DEP_INSTALL_HINT = (
    "缺依赖：python-docx 未安装（import docx 失败）。\n"
    "本技能的 Word 读写都要用它，请先在本机装一次：\n"
    "    pip install python-docx pypdfium2 docxcompose Pillow\n"
    "这是本地安装，脚本自身零网络请求；pip 走不通时请换企业/系统镜像源重试。"
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
RESERVED_NAMES = {"写作指引.md", "关键事实与承诺.md"}
SELF_NUMBERED = re.compile(r"^\s*(?:[（(]?[一二三四五六七八九十百\d]+\s*[、.．）)]|[①-⑳])")
ILLEGAL = re.compile(r'[/\\:*?"<>|\n\r]')


def die(msg: str, code: int = 1):
    print(f"[失败] {msg}")
    sys.exit(code)


def require_docx() -> None:
    """加载 python-docx 及其子模块（幂等）；缺失时给可执行提示后退出，不抛 traceback。"""
    global Document, WD_ORIENT, WD_STYLE_TYPE, WD_ALIGN_PARAGRAPH
    global OxmlElement, qn, Cm, Pt
    if Document is not None:
        return
    try:
        from docx import Document as _Document
        from docx.enum.section import WD_ORIENT as _WD_ORIENT
        from docx.enum.style import WD_STYLE_TYPE as _WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement as _OxmlElement
        from docx.oxml.ns import qn as _qn
        from docx.shared import Cm as _Cm, Pt as _Pt
    except ImportError:
        die(DEP_INSTALL_HINT, 2)
    Document, WD_ORIENT = _Document, _WD_ORIENT
    WD_STYLE_TYPE, WD_ALIGN_PARAGRAPH = _WD_STYLE_TYPE, _WD_ALIGN_PARAGRAPH
    OxmlElement, qn, Cm, Pt = _OxmlElement, _qn, _Cm, _Pt


def _build_review_cmd(work: Path) -> str:
    """可复制运行的 build_review.py 真实路径（五技能平级形态或单包扁平
    scripts/ 同目录，两种安装形态都可覆盖；兜底按部件拼装）。"""
    self_scripts = Path(__file__).resolve().parent
    for cand in (self_scripts.parent.parent / "tender-review" / "scripts" / "build_review.py",
                 self_scripts / "build_review.py"):
        if cand.is_file():
            return f"python3 {cand} --work {work}"
    return f"python3 {Path('tender-review', 'scripts', 'build_review.py')} --work {work}"


def _body_write_gate(target) -> None:
    """第五步写节前置硬闸：body/ 下的节 docx 首次写入前，写作进度面板必须已首开
    （build_review.py --open-progress 落 work/review/.progress-opened 标记）——
    监视面倒置为事后交付物是真机事故（26 节写完链接才首次出现），措辞层拦不住，
    闸挂在写节必经通道上。评审台也不存在时按红线「评审台未首建禁建节」拦：
    回 G3 补门，--open-progress 不替代门。整本-*.docx（合册产物）与非 body/
    路径（素材件、模板）不受此闸。"""
    if not target:
        return
    t = Path(target)
    body_dir = next((p for p in t.parents if p.name == "body"), None)
    if body_dir is None or t.stem.startswith("整本-"):
        return
    review = body_dir.parent / "review"
    if not (review / "评审台.html").is_file():
        print("[写节被拦] 评审台未首建禁建节（入口红线）——回入口技能补 G3："
              "目录确认 → 首建评审台 → 发链接 → 停轮，再进第五步。")
        sys.exit(2)
    if not (review / ".progress-opened").is_file():
        print("[写节被拦] 写作进度面板未首开——监视面必须先给到用户，开写第一动作：\n"
              f"  {_build_review_cmd(body_dir.parent)} --open-progress\n"
              "把回显的面板链接以 Markdown 发给用户，再写节（标记落盘后放行，只拦首节前）。")
        sys.exit(2)


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


_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _gov_label(depth: int, c: dict) -> str:
    """gov 形态各层编号（公文惯例：一、/（一）/1./（1）/①，第 5 层起封顶 ⑳）。
    chapter/decimal 的深层都是「.」拼接形态，gov 深层若同形（第 3/4 层都出「1.」）
    层级就不可读；节内小标题按同函数续编，不再做「num.strip() + '.' + sub」的
    字符串拼接（曾拼出「2..1」「（二）.1」）。"""
    if depth == 1:
        return f"{_cn(c.get(1, 1))}、"
    if depth == 2:
        return f"（{_cn(c.get(2, 1))}）"
    if depth == 3:
        return f"{c.get(3, 1)}. "
    if depth == 4:
        return f"（{c.get(4, 1)}）"
    return _CIRCLED[min(c.get(depth, 1), 20) - 1]


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


# ---------------------------------------------------------------- 中文字体归一

# 主题字体属性（w:asciiTheme 等）按 OOXML 规定优先于同名显式字体属性：样式里
# 两套并存时，显式设的「黑体」会被顶掉，实际取主题字体。主题东亚分支又由
# settings.xml 的 themeFontLang/@eastAsia 决定——python-docx 自带模板写的是
# ja-JP，于是标题落到日文分支（ＭＳ 明朝/ゴシック），简体缺字再逐字回退，
# 一行标题混排两种字体。开发机往往没装日文字体、渲染时静默回退，肉眼看不出来，
# 所以加载模板后一律归一到中文主题语言并摘掉越权的主题字体属性。
THEME_FONT_ATTRS = (("w:asciiTheme", "w:ascii"), ("w:hAnsiTheme", "w:hAnsi"),
                   ("w:eastAsiaTheme", "w:eastAsia"), ("w:cstheme", "w:cs"))
CJK_THEME_LANG = "zh-CN"
# themeFontLang 在 CT_Settings 里是定序元素：缺该元素时插在第一个后继之前。
# 后继清单按 schema 顺序，取本模块可能遇到的几个（中间的 caption/smartTagType 等
# 极少出现，缺了它们落到 append 也更安全）。
_AFTER_THEME_LANG = (
    "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
    "w:forceUpgrade", "w:captions", "w:readModeInkLockDown", "w:smartTagType",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)


def clear_theme_fonts(rpr_holder) -> bool:
    """摘掉被显式字体顶替的主题字体属性（返回是否改动）。

    逐对处理：`w:eastAsia` 在就删 `w:eastAsiaTheme`——显式值是明确意图，主题属性
    只是同层越权覆盖它。只有主题属性、没有显式值的样式（Heading 5、Title 之类）
    原样保留：它们本就该跟主题走，而主题语言已归一到简体中文分支。
    """
    rpr = rpr_holder.find(qn("w:rPr"))
    if rpr is None:
        return False
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return False
    changed = False
    for theme_attr, explicit_attr in THEME_FONT_ATTRS:
        if rfonts.get(qn(explicit_attr)) is not None and rfonts.get(qn(theme_attr)) is not None:
            del rfonts.attrib[qn(theme_attr)]
            changed = True
    return changed


def _set_theme_font_lang(settings_el) -> str | None:
    """把 settings 的主题语言归一到 zh-CN，返回原值（无改动返回 None）。"""
    el = settings_el.find(qn("w:themeFontLang"))
    if el is None:
        el = OxmlElement("w:themeFontLang")
        el.set(qn("w:val"), "en-US")
        anchor = None
        for tag in _AFTER_THEME_LANG:
            anchor = settings_el.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(el)
        else:
            settings_el.append(el)
    old = el.get(qn("w:eastAsia"))
    if old == CJK_THEME_LANG:
        return None
    el.set(qn("w:eastAsia"), CJK_THEME_LANG)
    return old or "(未设)"


def _set_doc_lang(styles_el) -> bool:
    """docDefaults 的校对语言改中文（不影响字体，只影响拼写检查与断词）。"""
    lang = styles_el.find(f'{qn("w:docDefaults")}/{qn("w:rPrDefault")}/'
                          f'{qn("w:rPr")}/{qn("w:lang")}')
    if lang is None or lang.get(qn("w:eastAsia")) == CJK_THEME_LANG:
        return False
    lang.set(qn("w:eastAsia"), CJK_THEME_LANG)
    return True


def ensure_cjk_theme(doc: Document) -> list[str]:
    """把主题东亚语言归一到简体中文，并摘掉显式样式上越权的主题字体属性。

    内置模板与用户模板都要过一遍：用户模板多半也源自 python-docx 或英文版 Word，
    同样带着 ja-JP 的 themeFontLang。返回的说明行供调用方回显。
    """
    notes: list[str] = []
    old = _set_theme_font_lang(doc.settings.element)
    if old:
        notes.append(f"主题东亚语言 {old} → {CJK_THEME_LANG}")
    cleared: list[str] = []
    for st in doc.styles.element.findall(qn("w:style")):
        name = st.get(qn("w:styleId")) or "?"
        if clear_theme_fonts(st):
            cleared.append(name)
    if cleared:
        notes.append("摘除主题字体属性：" + "、".join(cleared))
    if _set_doc_lang(doc.styles.element):
        notes.append(f"校对语言 → {CJK_THEME_LANG}")
    return notes


_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _theme_parts(doc: Document) -> list:
    """package 里的主题部件（word/theme/theme*.xml；themeFontLang 管不到的层）。"""
    try:
        return [p for p in doc.part.package.iter_parts()
                if str(getattr(p, "partname", "")).startswith("/word/theme/")]
    except Exception:
        return []


def _set_theme_ea(doc: Document) -> int:
    """theme 的东亚字形显式化：majorFont→黑体、minorFont→宋体。

    东亚分支为空 typeface 时，西文主题仍会兜住东亚解析（真机三层根因之二），
    光归一 themeFontLang 压不住。返回改动的 theme 部件数。"""
    n = 0
    for part in _theme_parts(doc):
        try:
            from lxml import etree

            root = etree.fromstring(part.blob)
        except Exception:
            continue
        changed = False
        for scheme_tag, face in ((_A_NS + "majorFont", "黑体"), (_A_NS + "minorFont", "宋体")):
            for scheme in root.iter(scheme_tag):
                ea = scheme.find(_A_NS + "ea")
                if ea is None:
                    ea = etree.SubElement(scheme, _A_NS + "ea")
                    latin = scheme.find(_A_NS + "latin")
                    if latin is not None:  # schema 定序：ea 跟在 latin 之后
                        latin.addnext(ea)
                if not (ea.get("typeface") or "").strip():
                    ea.set("typeface", face)
                    changed = True
        if changed:
            part._blob = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
            n += 1
    return n


# run 级东亚字体固化的样式族映射：标题/封面族→黑体，其余→宋体（setup_styles 同一套口径）
_HEADING_FONT_KEYS = ("Heading", "标题", "Cover", "Title")


def _iter_all_paragraphs(doc: Document):
    """正文段 + 表格格段（含嵌套表）+ 页眉页脚段——run 级后处理要覆盖的全部段落。"""

    def _tbl(tbl):
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for nested in cell.tables:
                    yield from _tbl(nested)

    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        yield from _tbl(t)
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                   sec.even_page_header, sec.even_page_footer):
            try:
                for p in hf.paragraphs:
                    yield p
            except Exception:
                continue


def _solidify_fonts(doc: Document) -> list[str]:
    """合册终稿字体固化（无条件执行）：theme 东亚字形补显式 + 样式层归一 + run 级固化。

    模板会被 skill 更新/重装覆盖，单靠模板归一不保险；copy-from 拷入的历史标书
    run 还带着 w:*Theme 越权属性或缺失 eastAsia，主题解析最终仍可能落到日文分支
    （真机整本混排 ＭＳ 明朝 的三层根因）。终稿后处理三层齐下，让字体问题在
    交付面上不可能复发。返回说明行供回显。"""
    notes: list[str] = []
    n_theme = _set_theme_ea(doc)
    if n_theme:
        notes.append(f"主题东亚字形补显式 {n_theme} 处")
    notes += ensure_cjk_theme(doc)
    stripped = filled = 0
    for p in _iter_all_paragraphs(doc):
        try:
            sname = (p.style.name if p.style is not None else "") or ""
        except Exception:
            sname = ""
        face = "黑体" if any(k in sname for k in _HEADING_FONT_KEYS) else "宋体"
        for r in p.runs:
            rpr = r._element.rPr
            if rpr is None:
                continue  # 无 rPr 的 run 全靠样式继承——归一过的样式层已安全
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                continue
            for theme_attr, _ in THEME_FONT_ATTRS:
                if rf.get(qn(theme_attr)) is not None:
                    del rf.attrib[qn(theme_attr)]  # 主题属性优先于显式字体，终稿一律剥掉
                    stripped += 1
            if rf.get(qn("w:eastAsia")) is None:
                rf.set(qn("w:eastAsia"), face)
                filled += 1
    if stripped:
        notes.append(f"摘 run 级主题字体属性 {stripped} 处")
    if filled:
        notes.append(f"run 级东亚字体固化 {filled} 处")
    return notes


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

    # 目录条目独立样式：单倍行距是硬要求——目录挂 Normal（1.5 倍距）会被行距类
    # 版式调整整页整页地放大（真机 15 页→12 页的教训）；独立样式也让版式压缩
    # 有明确的跳过目标。按级左缩进，一级加粗。
    for lvl in range(1, 5):
        name = f"TOC {lvl}"
        try:
            toc_st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        except ValueError:
            toc_st = doc.styles[name]
        toc_st.font.name = "宋体"
        toc_st.font.size = Pt(12)
        toc_st.font.bold = (lvl == 1)
        set_east_asia(toc_st.element, "宋体")
        toc_st.paragraph_format.line_spacing = 1.0
        toc_st.paragraph_format.first_line_indent = Pt(0)
        toc_st.paragraph_format.left_indent = Cm(0.74 * (lvl - 1))

    ensure_cjk_theme(doc)
    _set_theme_ea(doc)


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
        notes = ensure_cjk_theme(doc)
        if notes and user_template:
            print("[模板中文字体] " + "；".join(notes)
                  + f"（模板 {path.name} 的主题字体语言会顶掉显式中文字体，已归一）")
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
    root = _comments_root(part)
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
    _comments_flush(part, root)

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


def _comments_root(part):
    """comments part 的可写根。python-docx ≥1.2 把磁盘上已有批注的 part 载成
    CommentsPart（blob 每次从 _element 现序列化，改 _blob 是死存储——跨进程第二次
    加批注 body 静默丢、锚点残留成悬空引用）：有 _element 改活树；裸 Part（本工具
    自建形态）改 _blob。批注读写统一走这对助手，禁止直摸 part._blob。"""
    el = getattr(part, "_element", None)
    if el is not None:
        return el
    return _load_comments(part)


def _comments_flush(part, root) -> None:
    if getattr(part, "_element", None) is None:
        part._blob = _serialize_comments(root)


def _anchor_ids(el) -> set[str]:
    """正文里实际挂着引用的批注 id 集（commentReference 是 Word 展示批注的依据）。"""
    return {e.get(qn("w:id")) for e in el.findall(".//" + qn("w:commentReference"))}


def _purge_orphan_bodies(doc: Document) -> int:
    """删掉 comments part 里无任何正文锚点的批注体。revise 删段/重写会把锚点抹掉，
    留下的体在 Word 里不可见、却随合册 _merge_comments 迁移污染整本（真机实测
    评分索引表被页码回填抹掉 11 个锚，11 条孤儿体进了整本）。"""
    anchored = _anchor_ids(doc.element.body)
    removed = 0
    for rel in doc.part.rels.values():
        if not rel.reltype.endswith("/comments"):
            continue
        root = _comments_root(rel.target_part)
        for c in list(root.findall(qn("w:comment"))):
            if c.get(qn("w:id")) not in anchored:
                root.remove(c)
                removed += 1
        _comments_flush(rel.target_part, root)
    return removed


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

def para_text_full(p_el) -> str:
    """段落全文：含超链接、修订插入（w:ins）里的文字，不含修订删除（w:del）里的文字。

    python-docx 的 paragraph.text 只遍历段落的直接 w:r——超链接里的 run 和修订
    插入容器里的 run 都不在其中，read/revise/校验全部因此看漏文字（历史标书的
    修订痕迹段落曾整句不可见）。全文口径统一走本函数。"""
    out = []
    for t in p_el.iter(qn("w:t")):
        if any(a.tag in (qn("w:del"), qn("w:moveFrom")) for a in t.iterancestors()):
            continue
        out.append(t.text or "")
    return "".join(out)


def iter_blocks(doc: Document):
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def cell_text_full(cell) -> str:
    """单元格全文（含超链接/修订插入文字）。"""
    return "\n".join(para_text_full(p._p) for p in cell.paragraphs)


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
    existed = path.exists()
    dropped_comments = 0
    if existed and args.replace:
        try:
            old = Document(str(path))
        except Exception:
            old = None
        if old is not None:
            n_comments = len(doc_comments(old))
            if n_comments and not args.drop_comments:
                die(f"节文件已有 {n_comments} 条批注，覆盖会静默丢掉——确认后加 --drop-comments 重跑，"
                    f"或先用 comment-remove 逐条销号", 2)
            dropped_comments = n_comments
            if old.element.body.findall(".//" + qn("a:blip")):
                print("  ⚠️ 覆盖会丢弃节内已插图——重写后需 image 重新插入")
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
    tag = "已覆盖" if existed else "已建节"
    if dropped_comments:
        print(f"  （--drop-comments：已按指示丢弃原批注 {dropped_comments} 条）")
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


def _col_profiles(hspecs, rspecs, grid_w: int) -> tuple[list[int], list[int]]:
    """各列 (最长字数, 最长显示宽[半字符单位：全角2/半角1])。
    合并格（span>1 / vm）不参与——跨列长度无法归属单列。"""
    chars = [0] * grid_w
    disp = [0] * grid_w
    for specs in ([hspecs] if hspecs else []) + list(rspecs):
        pos = 0
        for text, span, vm in specs:
            if span == 1 and vm is None and pos < grid_w:
                t = text.strip()
                chars[pos] = max(chars[pos], len(t))
                disp[pos] = max(disp[pos], sum(2 if ord(ch) > 0x2E80 else 1 for ch in t))
            pos += span
    return chars, disp


def _label_min_tw(disp_half: int, size: float, content_tw: int) -> int:
    """短标签列（最长内容 ≤8 全角=16 半）保一行放下：字宽×字数+单元格边距。
    真机教训：按字数给比例、300 dxa 兜底，会把「渲染比对」拆成「渲染比/对」；
    长内容列按 16 半兜底而非 300——孤字窄列（「…初审专/家」）比合理折行更伤。
    上限 content_tw//4 防多列标签挤爆版心（标签宁可折行不撑破页）。"""
    if disp_half <= 0:
        return 300
    eff = min(disp_half, 16)
    return min(eff * int(size * 10) + 260, max(300, content_tw // 4))


def _fit_widths(cols: list[int], content_tw: int,
                mins: list[int] | None = None) -> list[int]:
    """列宽收口。mins=各列下限（短标签保一行），缺省 300 兜底。
    流程：下限总和超版心先比例缩（标签宁可折行不撑破页）→ 钳到下限 →
    总宽超版心时从非末列按比例收回（保下限）→ 仍收不齐接受表宽略超版心
    （负 dxa 才是 corrupt）→ 有余贴末列。"""
    if not mins or len(mins) != len(cols):
        mins = [300] * len(cols)
    mins = [max(300, m) for m in mins]
    if sum(mins) > content_tw:
        f = content_tw / sum(mins)
        mins = [max(300, int(m * f)) for m in mins]
    cols = [max(c, m) for c, m in zip(cols, mins)]
    diff = content_tw - sum(cols)
    if diff >= 0:
        cols[-1] += diff  # 版心有余：余数贴末列
        return cols
    over = -diff  # 超出量：从有富余的列按比例收回（含末列——钳到下限后，
    # 富余常集中在长内容末列；mins 钳过不存在「末列欠债」形态）
    flexible = [i for i in range(len(cols)) if cols[i] > mins[i]]
    pool = sum(cols[i] - mins[i] for i in flexible)
    if pool:
        for i in flexible:
            slack = cols[i] - mins[i]
            take = min(slack, int(over * slack / pool))
            cols[i] -= take
            over -= take
    for i in flexible:  # 按比例取整的零头顺次补收
        if over <= 0:
            break
        take = min(over, cols[i] - mins[i])
        cols[i] -= take
        over -= take
    return cols  # 仍收不齐（全顶下限）= 接受表宽略超版心


def _grid_col_widths(blk: dict, grid_w: int, content_tw: int,
                     auto_w: list[int] | None = None,
                     mins: list[int] | None = None) -> list[int]:
    weights = blk.get("widths")
    if weights:
        if len(weights) != grid_w:
            raise ValueError(f"widths 数量（{len(weights)}）与表格网格列数（{grid_w}）不符")
        wsum = sum(float(x) for x in weights)
        if wsum <= 0:
            weights = None
        else:
            return _fit_widths(
                [max(300, int(content_tw * float(x) / wsum)) for x in weights],
                content_tw, mins)
    if auto_w and len(auto_w) == grid_w and len(set(auto_w)) > 1:
        # 无 widths：按各列内容长度比例分配（长内容列变宽，评委好读）
        wsum = sum(auto_w)
        return _fit_widths([max(300, int(content_tw * w / wsum)) for w in auto_w],
                           content_tw, mins)
    return _fit_widths([max(300, content_tw // grid_w)] * grid_w, content_tw, mins)


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
        t.text = text  # 不截断：静默丢字是标书事故，超长提示在 _emit_table_inner 里给出
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
    # 超长单元格点名（旧版 t.text=text[:200] 静默截断是标书事故，200 字以上提示行列位）
    for rn, specs in enumerate([*([hspecs] if hspecs else []), *rspecs], 1):
        pos = 0
        for t, span, vm in specs:
            if span == 1 and vm is None and len(t) > 200:
                print(f"  ⚠️ 单元格超长（第 {rn} 行第 {pos + 1} 列，共 {len(t)} 字）"
                      f"——已按全文生成，窄列会折多行")
            pos += span
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
    size = 10.5 if grid_w <= 6 else (9.5 if grid_w <= 8 else 9.0)
    chars, disp = _col_profiles(hspecs, rspecs, grid_w)
    auto_w = [max(6, min(50, c)) for c in chars] if any(chars) else None
    mins = [_label_min_tw(d, size, content_tw) for d in disp]
    col_w = _grid_col_widths(blk, grid_w, content_tw, auto_w, mins)
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
                        norm = " ".join(cell_text_full(cell).split())
                        # 截到 200 字防刷屏，但标出总字数——60 字静默截断曾让
                        # 长参数单元格在 read 视图里显得完整，模型无从发现丢失
                        cells.append(norm[:200] + (f"（共 {len(norm)} 字）" if len(norm) > 60 else ""))
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
        print(f"[P{pi}]{tag}{marks}{para_text_full(block._p)}")


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


def _rewrite_cell_text(cell, text: str) -> None:
    """清格重写但保留批注锚点。直接 `cell.text = ""` 会连 rangeStart/End/引用一起
    抹掉（真机实测：评分索引表页码回填抹掉 11 个格上的批注锚，11 条孤儿体进了
    整本）——这里只摘 run，rangeStart/End 原位保留，新文字插回批注范围之内，
    引用 run 补回 rangeEnd 之后。"""
    p0 = cell.paragraphs[0]
    ids = [el.get(qn("w:id")) for el in p0._p.findall(qn("w:commentRangeStart"))]
    keep_rpr = None  # 该格首个显式格式：格式件填空沿用原格式（仿宋 14pt 加粗下划线的
    # 空格曾被写成宋体 10.5，同一张表里填过的和没填的字体不一致）；空格无 run 就不写
    for p in cell.paragraphs:
        for r_el in p._p.findall(qn("w:r")):
            rp = r_el.find(qn("w:rPr"))
            if rp is not None:
                keep_rpr = _deepcopy(rp)
                break
        if keep_rpr is not None:
            break
    for p in cell.paragraphs:
        for r_el in p._p.findall(qn("w:r")):
            p._p.remove(r_el)
    for extra in list(cell.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)
    run = p0.add_run(text)
    if keep_rpr is not None:
        run._element.insert(0, keep_rpr)
    range_end = p0._p.find(qn("w:commentRangeEnd"))
    if range_end is not None:
        p0._p.remove(run._r)
        range_end.addprevious(run._r)  # 新文字落回批注范围之内
    for cid in ids:
        ref_run = OxmlElement("w:r")
        ref = OxmlElement("w:commentReference")
        ref.set(qn("w:id"), cid)
        ref_run.append(ref)
        if range_end is not None:
            range_end.addnext(ref_run)
        else:
            p0._p.append(ref_run)


def cmd_revise(args) -> None:
    if not args.cell and not args.para:
        die("需要 --para P# 或 --cell T#R#C#（read 视图核对坐标）")
    doc = Document(args.path)
    if args.cell:
        m = re.match(r"T(\d+)R(\d+)C(\d+)$", args.cell, re.I)
        if not m:
            die(f"--cell 格式应为 T1R2C3：{args.cell}")
        t, r, c = (int(x) for x in m.groups())
        tbl = _find_table(doc, t)
        if r > len(tbl.rows):
            die(f"格 {args.cell} 越界（该表 {len(tbl.rows)} 行 × {len(tbl.columns)} 列）")
        if args.mode == "insert_after":
            # 在该格所在行之后插一个空行：克隆该行网格（列宽/跨列随行保留）、清空文字。
            # 格式件模板空行不够时的正路——不校验查重；插行后行号变化须重读坐标。
            tr = tbl.rows[r - 1]._tr
            for tc_el in tr.findall(qn("w:tc")):
                tcpr = tc_el.find(qn("w:tcPr"))
                if tcpr is not None and tcpr.find(qn("w:vMerge")) is not None:
                    die("该行含竖向合并格——插行请选普通数据行（克隆合并行会带出悬挂续行）")
            new_tr = _deepcopy(tr)
            for tc_el in new_tr.findall(qn("w:tc")):
                ps = tc_el.findall(qn("w:p"))
                for p_el in ps:
                    for r_el in p_el.findall(qn("w:r")):
                        p_el.remove(r_el)
                for extra in ps[1:]:
                    tc_el.remove(extra)
            tr.addnext(new_tr)
            doc.save(args.path)
            print(f"[已插行] {args.path}（T{t}R{r} 之后新增空行，网格随原行克隆）")
            print("  ⚠️ 行号已变化：后续 revise/comment 前先重新 read 取坐标")
            return
        try:
            tc = tbl.cell(r - 1, c - 1)._tc
        except Exception:
            die(f"格 {args.cell} 越界（该表 {len(tbl.rows)} 行 × {len(tbl.columns)} 列）")
        # 合并区坐标解析回承载格后写入并点名（旧版「延续位」守卫是死代码——
        # tbl.cell() 已把延续位解析回起始格，die 永远不触发）
        anch = _grid_anchor(tbl, r, c)
        if anch is not None and (anch[0] != r or anch[1] != c):
            print(f"  （{args.cell} 位于合并区，已写入承载格 T{t}R{anch[0]}C{anch[1]}）")
        # 合并格守卫：跨列合并行的延续位写入会静默落到锚点格（实测页码错进组行），
        # 竖向合并延续位写入会落进看不见的续格——都直接拦下并指路
        row_tcs = tbl.rows[r - 1].cells
        first_idx = next((i for i, cc in enumerate(row_tcs) if cc._tc is tc), c - 1)
        if first_idx != c - 1:
            die(f"{args.cell} 是跨列合并行的延续位（整行一个格子）——"
                f"组行写入请用锚点格 T{t}R{r}C{first_idx + 1}，数据行请换行")
        tcpr = tc.find(qn("w:tcPr"))
        vm = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
        if vm is not None and vm.get(qn("w:val")) != "restart":
            die(f"{args.cell} 是竖向合并的延续位（内容在上方合并起始格）——"
                f"请用 read 视图核对该表合并结构后改写起始格")
        cell = tbl.cell(r - 1, c - 1)
        _rewrite_cell_text(cell, args.text or "")
        doc.save(args.path)
        print(f"[已修订] {args.path}（{args.cell} ← {args.text[:30]}）")
        return
    para = _find_para(doc, args.para)
    if args.mode == "delete":
        para._p.getparent().remove(para._p)
        n_orphan = _purge_orphan_bodies(doc)
        doc.save(args.path)
        print(f"[已修订] {args.path}（P{args.para} 已删除）")
        if n_orphan:
            print(f"  （随段删除的批注体 {n_orphan} 条已一并清掉，不留孤儿体）")
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
    new_s = args.text or ""
    hit_runs = [r for r in para.runs if args.find in r.text]
    if hit_runs:
        # 命中落在直接 run 内：只改该 run——同段其余 run（加粗引题等）的格式原样，
        # 本 run 的 rPr 也不动（旧版整段重写成宋体 12pt，标题被顶成正文样式）
        for r in hit_runs:
            r.text = r.text.replace(args.find, new_s)
        doc.save(args.path)
        print(f"[已修订] {args.path}（P{args.para}：{args.find[:20]} → {new_s[:20]}）")
        return
    full = para_text_full(para._p)
    if args.find not in full:
        die(f"P{args.para} 中找不到「{args.find}」（read 视图核对原文）")
    new_text = full.replace(args.find, new_s)
    ref_ids = [el.get(qn("w:id")) for el in para._p.findall(".//" + qn("w:commentReference"))]
    fmt_rpr = None  # 首个「文字 run」的显式格式：跳过只装批注引用的 run 与已删除文字；
    # 取「首个带格式的 run」会把整段染成该 run 的格式（实测加粗扩散），故按文字 run 取
    for r_el in para._p.iter(qn("w:r")):
        if r_el.find(qn("w:t")) is None or \
                any(a.tag in (qn("w:del"), qn("w:moveFrom")) for a in r_el.iterancestors()):
            continue
        rp = r_el.find(qn("w:rPr"))
        fmt_rpr = _deepcopy(rp) if rp is not None else None
        break
    for r_el in list(para._p.iter(qn("w:r"))):
        r_el.getparent().remove(r_el)
    for el in list(para._p):
        # 摘掉超链接/修订容器等全部非文字结构（旧版漏摘 w:hyperlink，新旧文字重复）；
        # 批注范围标记保留，引用 run 在下方补回
        if el.tag not in (qn("w:pPr"), qn("w:commentRangeStart"), qn("w:commentRangeEnd")):
            para._p.remove(el)
    run = para.add_run(new_text)
    if fmt_rpr is not None:  # 无显式格式就不写 rPr——继承段落样式，标题不被抹成宋体 12
        run._element.insert(0, fmt_rpr)
    for cid in ref_ids:  # 引用 run 随旧文字被摘，补回段尾使批注仍挂在本段
        ref_run = OxmlElement("w:r")
        ref = OxmlElement("w:commentReference")
        ref.set(qn("w:id"), cid)
        ref_run.append(ref)
        para._p.append(ref_run)
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
    head = args.text[:50]
    tail = "……（回显截断，批注全文已落盘）" if len(args.text) > len(head) else ""
    print(f"[已批注] {args.path}（{loc}，批注 #{cid}）：{head}{tail}")


def _row_grid(row):
    """一行的物理格清单 [(tc, 起始网格列1基, 跨列数, vMerge)]——合并行的坐标解析基准。"""
    grid = []
    pos = 0
    for tc in row._tr.findall(qn("w:tc")):
        tcpr = tc.find(qn("w:tcPr"))
        span = 1
        vm = None
        if tcpr is not None:
            gs = tcpr.find(qn("w:gridSpan"))
            if gs is not None:
                span = int(gs.get(qn("w:val")) or 1)
            v = tcpr.find(qn("w:vMerge"))
            if v is not None:
                vm = v.get(qn("w:val")) or "continue"
        grid.append((tc, pos + 1, span, vm))
        pos += span
    return grid


def _grid_anchor(tbl, r: int, c: int) -> tuple[int, int] | None:
    """请求格坐标 (1基) → 承载内容的物理格坐标 (1基行, 1基起始网格列)。

    合并区（横向/竖向）的任意坐标都解析回锚点格——与 comment --cell 写入时的
    tbl.cell() 解析同一语义；旧版 comment-list 按「唯一格序」报坐标，横向合并行
    会差一位、竖向合并会落到最下面的延续行，销号时对不上。行越界返回 None。"""
    if r < 1 or r > len(tbl.rows):
        return None
    for tc, start, span, vm in _row_grid(tbl.rows[r - 1]):
        if start <= c < start + span:
            if vm == "continue":  # 竖向合并延续格：向上找同列的 restart 行
                for rr in range(r - 1, 0, -1):
                    for tc2, start2, span2, vm2 in _row_grid(tbl.rows[rr - 1]):
                        if start2 <= c < start2 + span2:
                            if vm2 == "restart":
                                return rr, start2
                            break  # 同列非 restart（畸形）：就地返回请求行
                return r, start
            return r, start
    return r, c  # 列越界：原样返回，交由上层报「上没有批注」


def _comment_locs(doc: Document) -> tuple[dict[str, str], dict[str, tuple[int, int]]]:
    """批注 id → (定位串, 文档位次)。定位串与 comment --cell 同一记法（合并区解析回
    锚点格的起始网格列），comment-remove 按此回删；文档位次按块序（P 与 T 按
    真实文档顺序交错），comment-list 排序用——旧版先段落后表格的「位次」与文档
    顺序不一致。"""
    locs: dict[str, str] = {}
    order: dict[str, tuple[int, int, int]] = {}
    pi = ti = 0
    for bi, block in enumerate(iter_blocks(doc), 1):
        if hasattr(block, "rows"):  # Table
            ti += 1
            for ri, row in enumerate(block.rows, 1):
                for gi, (tc, start, span, vm) in enumerate(_row_grid(row)):
                    for p in tc.findall(qn("w:p")):
                        for el in p.findall(qn("w:commentRangeStart")):
                            cid = el.get(qn("w:id"))
                            locs[cid] = f"T{ti}R{ri}C{start}"
                            order[cid] = (bi, ri, gi)  # 块序+行序+格序：跨行位次可比较
        else:
            pi += 1
            for el in block._p.findall(qn("w:commentRangeStart")):
                cid = el.get(qn("w:id"))
                locs[cid] = f"P{pi}"
                order[cid] = (bi, 0, 0)
    return locs, order


def cmd_comment_list(args) -> None:
    doc = Document(args.path)
    comments = doc_comments(doc)
    if not comments:
        print(f"[无批注] {args.path}")
        return
    locs, order = _comment_locs(doc)
    print(f"[批注清单] {args.path}（共 {len(comments)} 条，按文档位次）")

    for cid in sorted(comments, key=lambda c: (order.get(c, (10 ** 9,)), c)):
        print(f"  #{cid} [{locs.get(cid, '?')}] {comments[cid]}")
    print("销号回删：comment-remove <path> --id N（可重复）/ --para P# / "
          "--cell T#R#C# / --all；孤儿体（锚点被 revise 抹掉的批注）用 --orphan 清；"
          "节文件删后需重新 assemble 迁移")


def _remove_comments(doc: Document, cids: set[str]) -> None:
    """摘掉正文锚点（range 标记 + 引用）与 comments part 条目。引用所在 run 可能
    携带正文文字（外来批注），只摘引用元素不删整个 run，避免误伤正文。"""
    body = doc.element.body
    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
        for el in list(body.findall(".//" + qn(tag))):
            if el.get(qn("w:id")) in cids:
                el.getparent().remove(el)
    for ref in list(body.findall(".//" + qn("w:commentReference"))):
        if ref.get(qn("w:id")) in cids:
            ref.getparent().remove(ref)
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/comments"):
            root = _comments_root(rel.target_part)
            for c in list(root.findall(qn("w:comment"))):
                if c.get(qn("w:id")) in cids:
                    root.remove(c)
            _comments_flush(rel.target_part, root)


def cmd_comment_remove(args) -> None:
    pickers = [bool(args.id), bool(args.para), bool(args.cell), args.all, args.orphan]
    if sum(pickers) == 0:
        die("--id / --para / --cell / --all / --orphan 至少给一个（comment-list 核对批注号）")
    if sum(pickers) > 1:
        die("--id/--para/--cell/--all/--orphan 互斥——一次只给一个选择器")
    doc = Document(args.path)
    comments = doc_comments(doc)
    if not comments:
        die(f"无批注可删：{args.path}")
    locs, _order = _comment_locs(doc)
    if args.orphan:
        targets = set(comments) - _anchor_ids(doc.element.body)
        if not targets:
            die("无孤儿批注体（批注体都有正文锚点）——不需要清理")
    elif args.all:
        targets = set(comments)
    elif args.id:
        targets = set()
        for spec in args.id:
            m = re.fullmatch(r"#?(\d+)", spec.strip())
            if not m:
                die(f"--id 格式应为数字（comment-list 回显的 #N）：{spec}")
            targets.add(m.group(1))
    else:
        if args.para:
            key = f"P{args.para}"
        else:
            m = re.match(r"T(\d+)R(\d+)C(\d+)$", args.cell or "", re.I)
            if not m:
                die(f"--cell 格式应为 T1R2C3：{args.cell}")
            t, r, c = (int(x) for x in m.groups())
            tbl = _find_table(doc, t)
            anch = _grid_anchor(tbl, r, c)
            if anch is None:
                die(f"行越界：{args.cell}（该表 {len(tbl.rows)} 行）")
            key = f"T{t}R{anch[0]}C{anch[1]}"  # 合并区坐标解析回锚点格再查
        hits = [cid for cid, loc in locs.items() if loc == key]
        if not hits:
            die(f"{key} 上没有批注")
        targets = set(hits)
    unknown = sorted(targets - set(comments))
    if unknown:
        die(f"批注号不存在：{'、'.join('#' + c for c in unknown)}（comment-list 核对）")
    _remove_comments(doc, targets)
    doc.save(args.path)
    print(f"[已删批注] {args.path}（{len(targets)} 条："
          f"{'、'.join('#' + c for c in sorted(targets))}）；剩余 "
          f"{len(comments) - len(targets)} 条")


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


# ---------------------------------------------------------------- 版式压缩（compact）

def max_font_pt(p, default: float = 12.0) -> float:
    """段内最大字号（磅）——run 未显式设字号按 Normal 缺省 12。"""
    best = 0.0
    for r in p.runs:
        sz = r.font.size
        best = max(best, sz.pt if sz is not None else default)
    return best or default


def has_image(p) -> bool:
    """段内是否含图——exact 行距会把图裁成一条，压缩的硬保护判据。"""
    return para_images(p) > 0


def cmd_compact(args) -> None:
    """正文段固定行距压缩。两条保护常开不可关（可预测可预防的事故不靠事后补救）：
    大字号段（字号×1.2 超行距即压扁风险，36pt 封面标题事故）与含图段一律跳过，
    TOC 样式段（目录单倍行距是独立样式族的硬约定）跳过。表格与页眉页脚不动。"""
    from docx.enum.text import WD_LINE_SPACING

    doc = Document(args.docx)
    applied = big = img = toc_skipped = 0
    for p in doc.paragraphs:
        sname = (p.style.name if p.style is not None else "") or ""
        if sname.replace(" ", "").upper().startswith("TOC"):
            toc_skipped += 1
            continue
        if not p.text.strip() and not has_image(p):
            continue  # 空段是格式件的结构间隔，压了无收益还添乱
        if has_image(p):
            img += 1
            continue
        if max_font_pt(p) * 1.2 > args.line_pt:
            big += 1
            continue
        pf = p.paragraph_format
        pf.line_spacing = Pt(args.line_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        applied += 1
    doc.save(args.docx)
    print(f"[版式压缩] 固定行距 {args.line_pt:g}pt：正文段 {applied}；保护跳过——"
          f"大字号 {big} 段、含图 {img} 段、目录 {toc_skipped} 段（保护常开不可关）")


# ---------------------------------------------------------------- 图片体检（image-scan / image-rotate）

def _image_rels(doc: Document) -> list:
    """文档主部件的图片关系，按媒体文件名排序——scan 与 rotate 的序号基准一致。"""
    return sorted((rel for rel in doc.part.rels.values() if rel.reltype.endswith("/image")),
                  key=lambda r: str(r.target_part.partname))


def _blip_ref_counts(doc: Document) -> dict:
    """各图片关系被引用次数（正文 + 页眉页脚的 a:blip r:embed / v:imagedata r:id）。"""
    from collections import Counter

    counts = Counter()
    parts = [doc.part]
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                   sec.even_page_header, sec.even_page_footer):
            try:
                parts.append(hf.part)
            except Exception:
                continue
    for part in parts:
        for el in part.element.iter():
            if el.tag == qn("a:blip"):
                rid = el.get(qn("r:embed"))
            elif el.tag == f"{{{V_NS}}}imagedata":  # VML 前缀不在 python-docx nsmap 里，qn() 会 KeyError
                rid = el.get(qn("r:id"))
            else:
                continue
            if rid:
                counts[rid] += 1
    return counts


def cmd_image_scan(args) -> None:
    """导出全部内嵌图片的缩略图网格：证件方向自动判定不可靠（行投影分不出
    0°/180°，OSD 在证件上误判率高），人眼扫网格最快——63 张证件 9 张错向的
    人工成本大头。网格文件名带序号，错了记序号走 image-rotate 修。"""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        die("缺少依赖：Pillow。请先执行：pip install Pillow", 2)
    docx_path = Path(args.docx)
    if not docx_path.is_file():
        die(f"文件不存在：{docx_path}")
    doc = Document(str(docx_path))
    rels = _image_rels(doc)
    if not rels:
        die("文档内没有图片（先 image 插入，或经 copy-from 带图拷贝）")
    counts = _blip_ref_counts(doc)
    infos = []
    for i, rel in enumerate(rels, 1):
        part = rel.target_part
        w = h = 0
        try:
            with Image.open(io.BytesIO(part.blob)) as im:
                w, h = im.size
        except Exception:
            pass
        infos.append((i, str(part.partname).rsplit("/", 1)[-1], w, h,
                      counts.get(rel.rId, 0), part))
    print(f"[图片清单] {docx_path.name} 共 {len(infos)} 张（序号供 image-rotate 用）：")
    for i, fname, w, h, n, _part in infos:
        mark = "" if n else "（未被引用）"
        print(f"  #{i} {fname} {w}x{h} 引用 {n} 次{mark}")
    out_dir = Path(args.out) if args.out else docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    thumb, cols, pad, label_h = args.thumb, args.cols, 10, 22
    per_sheet = cols * 6
    pages = [infos[i:i + per_sheet] for i in range(0, len(infos), per_sheet)]
    made = 0
    for pi, page in enumerate(pages, 1):
        rows = (len(page) + cols - 1) // cols
        cw, ch = thumb + pad * 2, thumb + pad * 2 + label_h
        sheet = Image.new("RGB", (cw * cols, ch * rows), "white")
        draw = ImageDraw.Draw(sheet)
        font = ImageFont.load_default()  # 位图字体：标签是 ASCII（#1 400x200），无需中文字体
        for j, (i, fname, w, h, n, part) in enumerate(page):
            x0, y0 = (j % cols) * cw, (j // cols) * ch
            try:
                with Image.open(io.BytesIO(part.blob)) as im:
                    im.thumbnail((thumb, thumb))
                    if im.mode not in ("RGB", "L"):
                        im = im.convert("RGB")
                    sheet.paste(im, (x0 + (cw - im.width) // 2,
                                     y0 + (ch - label_h - im.height) // 2))
            except Exception:
                draw.text((x0 + pad, y0 + pad), "N/A", fill="red", font=font)
            draw.text((x0 + pad, y0 + ch - label_h - 2), f"#{i} {w}x{h}",
                      fill="black", font=font)
        sheet.save(out_dir / f"图片清单-{docx_path.stem}-{pi:02d}.png")
        made += 1
    print(f"[已导出] {made} 张缩略图网格 → {out_dir}（人工核对方向，记序号走 image-rotate 修）")


def _swap_extents(doc: Document, rids: set, swap: bool) -> int:
    """引用给定图片关系的图形尺寸改为旋转后口径：90°/270° 交换 wp:extent 与
    a:xfrm/a:ext 的 cx/cy（同一媒体被多处引用全部重算）；180° 尺寸不变只计数。
    死坑：旋转后不按新比例重算 extent，Word 会按旧宽高渲染——拉伸变形。"""
    n = 0
    parts = [doc.part]
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                   sec.even_page_header, sec.even_page_footer):
            try:
                parts.append(hf.part)
            except Exception:
                continue
    for part in parts:
        for blip in part.element.iter(qn("a:blip")):
            if blip.get(qn("r:embed")) not in rids:
                continue
            anchor = blip
            while anchor is not None and anchor.tag not in (qn("wp:inline"), qn("wp:anchor")):
                anchor = anchor.getparent()
            if anchor is None:
                continue
            exts = [anchor.find(qn("wp:extent"))]
            exts += anchor.findall(f".//{qn('a:xfrm')}/{qn('a:ext')}")
            for ext in exts:
                if ext is None:
                    continue
                cx, cy = ext.get("cx"), ext.get("cy")
                if swap and cx and cy:
                    ext.set("cx", cy)
                    ext.set("cy", cx)
                n += 1
    return n


def cmd_image_rotate(args) -> None:
    """按 image-scan 序号就地旋转图片（证件方向人工核出错向后的修复动作）。"""
    try:
        from PIL import Image
    except ImportError:
        die("缺少依赖：Pillow。请先执行：pip install Pillow", 2)
    docx_path = Path(args.docx)
    doc = Document(str(docx_path))
    rels = _image_rels(doc)
    if not (1 <= args.index <= len(rels)):
        die(f"图片序号越界：#{args.index}（共 {len(rels)} 张——序号以 image-scan 输出为准）")
    rel = rels[args.index - 1]
    part = rel.target_part
    fname = str(part.partname).rsplit("/", 1)[-1]
    try:
        im = Image.open(io.BytesIO(part.blob))
    except Exception:
        die(f"无法解码 {fname}（WMF/EMF 矢量图不支持就地旋转）——导出后外部处理再插回")
    fmt = im.format if im.format in ("PNG", "JPEG", "BMP", "GIF") else None
    if fmt is None:
        die(f"图片格式 {im.format or '未知'} 不支持就地旋转（支持 PNG/JPEG/BMP/GIF）")
    fixed = im.transpose({90: Image.ROTATE_90, 180: Image.ROTATE_180,
                          270: Image.ROTATE_270}[args.deg])
    buf = io.BytesIO()
    fixed.save(buf, format=fmt)
    part._blob = buf.getvalue()
    # 同一媒体可能被多处引用：所有指向该部件的关系都要重算 extent
    rids = {r.rId for r in doc.part.rels.values()
            if r.reltype.endswith("/image") and r.target_part is part}
    n_ext = _swap_extents(doc, rids, swap=(args.deg in (90, 270)))
    doc.save(str(docx_path))
    print(f"[已旋转] #{args.index} {fname} {args.deg}°（{im.width}x{im.height} → "
          f"{fixed.width}x{fixed.height}，重算引用处 extent {n_ext} 处——"
          f"不按新比例重算会拉伸变形）")


def _parse_pages(spec: str, n_pages: int) -> list[int]:
    """'3-7' / '3,5,7' / '7' → 去重升序页号列表（1 基，越界报错）。"""
    got: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, hi = (int(x) for x in part.split("-", 1))
        else:
            lo = hi = int(part)
        if lo < 1 or hi < lo:
            die(f"--pages 区间不合法：{part}")
        got.update(range(lo, hi + 1))
    pages = sorted(got)
    if not pages:
        die("--pages 为空")
    if pages[-1] > n_pages:
        die(f"--pages 越界：{spec}（该 PDF 共 {n_pages} 页）")
    return pages


def cmd_render_page(args) -> None:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        die("缺少依赖：pypdfium2、Pillow。请先执行：pip install pypdfium2 Pillow", 2)
    pdf = pdfium.PdfDocument(args.pdf)
    if getattr(args, "pages", ""):
        pages = _parse_pages(args.pages, len(pdf))
        out = Path(args.out)
        made = []
        for p in pages:
            img = pdf[p - 1].render(scale=2).to_pil()
            path = out.parent / f"{out.stem}-p{p:03d}.png"
            img.save(path)
            made.append(f"{path}（第 {p} 页，{img.width}×{img.height}）")
        print(f"[已导出] {len(pages)} 页：\n  " + "\n  ".join(made))
        return
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
    clean: dict[str, int] = {}
    for e in selected:
        if e[0] >= len(blocks):
            continue
        blk = blocks[e[0]]
        el = blk._p if hasattr(blk, "_p") else blk._tbl
        clone, n_img, skipped, st = _clone_with_images(el, src_doc, tgt_doc)
        img_copied += n_img
        img_skipped.extend(skipped)
        for k, v in st.items():
            clean[k] = clean.get(k, 0) + v
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
    if any(clean.values()):
        print(f"  （已清理源文件痕迹：{_fmt_stats(clean)}）")


def _deepcopy(el):
    import copy

    return copy.deepcopy(el)


V_NS = "urn:schemas-microsoft-com:vml"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# ---------------------------------------------------------------- 痕迹清理（copy-from / 合册共用）
#
# 历史标书的段落常带修订痕迹（w:ins/w:del 带作者日期）、批注锚点、超链接——原样
# 拷进节文件后：修订痕迹直接进递交稿；锚点失去批注体成悬空引用；超链接的关系号
# 不重挂会撞上目标文件同号关系（实测指向页脚部件）。copy-from 拷贝与 assemble
# 打开节文件都走同一套清理。

_COPY_UNWRAP = ("w:ins", "w:moveTo")  # 接受插入：解包容器、内容原位保留
_COPY_DROP = (                        # 接受删除 / 摘锚点 / 摘修订元数据
    "w:del", "w:moveFrom",
    "w:moveFromRangeStart", "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd",
    "w:commentRangeStart", "w:commentRangeEnd", "w:commentReference",
    "w:footnoteReference", "w:endnoteReference",
    "w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange", "w:tblGridChange",
    "w:trPrChange", "w:tcPrChange", "w:numberingChange", "w:cellIns", "w:cellDel",
    "w:cellMerge",
    "w:customXmlInsRangeStart", "w:customXmlInsRangeEnd",
    "w:customXmlDelRangeStart", "w:customXmlDelRangeEnd",
)


def _sanitize_element(el, src_doc: Document, tgt_doc: Document,
                      keep_anchors: set[str] | None = None) -> dict[str, int]:
    """对一棵元素树接受修订、摘批注/脚注锚、重挂超链接。返回各项计数。

    keep_anchors=None：锚点全摘（copy-from 场景——批注体不随片段走，留下的都是悬空锚）；
    传 id 集合：集合内的锚保留（assemble 场景——节文件自带的待办批注锚+体同在，要照常迁移）。"""
    n: dict[str, int] = {"接受插入": 0, "接受删除": 0, "摘批注锚": 0,
                         "重挂外链": 0, "摘坏引用": 0}
    for tag in _COPY_UNWRAP:
        for holder in list(el.iter(qn(tag))):
            parent = holder.getparent()
            if parent is None:
                continue
            idx = parent.index(holder)
            for child in list(holder):
                parent.insert(idx, child)
                idx += 1
            parent.remove(holder)
            n["接受插入"] += 1
    for tag in _COPY_DROP:
        for x in list(el.iter(qn(tag))):
            if keep_anchors is not None and tag.startswith("w:comment") \
                    and x.get(qn("w:id")) in keep_anchors:
                continue
            x.getparent().remove(x)
            if tag == "w:del":
                n["接受删除"] += 1
            elif tag in ("w:commentRangeStart", "w:commentReference"):
                n["摘批注锚"] += 1
    for r_el in list(el.iter(qn("w:r"))):  # 摘引用后只剩 rPr 的空 run
        if r_el.find(qn("w:t")) is None and r_el.find(qn("w:br")) is None \
                and r_el.find(qn("w:tab")) is None \
                and not r_el.findall(qn("w:drawing")) \
                and all(c.tag == qn("w:rPr") for c in r_el):
            r_el.getparent().remove(r_el)
    for hl in list(el.iter(qn("w:hyperlink"))):
        rid = hl.get(R_ID)
        if not rid:
            continue
        try:
            rel = src_doc.part.rels[rid]
        except KeyError:
            hl.attrib.pop(R_ID, None)
            n["摘坏引用"] += 1
            continue
        if rel.is_external:
            hl.set(R_ID, tgt_doc.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
            n["重挂外链"] += 1
        else:
            hl.attrib.pop(R_ID, None)  # 文档内链走 w:anchor，不依赖关系号
    return n


def _sanitize_doc(doc: Document) -> dict[str, int]:
    """整文档接受修订；批注锚只摘「无批注体」的悬空锚——文件自带的待办批注
    （锚+体同在）保留，照常迁移进整本。改的是内存树，落盘由调用方定。"""
    anchored = set(doc_comments(doc))
    return _sanitize_element(doc.element.body, doc, doc, keep_anchors=anchored)


def _fmt_stats(n: dict[str, int]) -> str:
    return "、".join(f"{k} {v}" for k, v in n.items() if v)


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
    """深拷贝元素并把图片重挂到目标文档；返回 (clone, 成功图片数, 跳过清单, 清理计数)。
    DrawingML（a:blip）与 VML（v:imagedata）双通道；EMF/WMF 等 python-docx 不认的
    格式按张跳过并记录，不中断整条拷贝。拷贝后统一接受修订/摘锚/重挂链接。"""
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
    stats = _sanitize_element(clone, src_doc, tgt_doc)
    return clone, ok, skipped, stats


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
        die("环境无 matplotlib——真图阶梯改走宿主渲染（HTML/浏览器截图出 PNG 再 image 插入），"
            "或先执行：pip install matplotlib", 2)
    font = _cjk_font()
    if font is None:
        die("未检测到中文字体——真图会显示乱码，真图阶梯改走宿主渲染"
            "（HTML 渲染中文不依赖本机字体）或表格表达")
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


# ---------------------------------------------------------------- 格式件保真核对（红线）

# 招标文件给了格式的表只能 copy-from 拷原件、往空格里填：行列与合并格一个不动。
# 真机事故（之江实验室项目）：12 网格列的「其他资信资料」被重排成两列键值表、
# 「评分索引表」被改成五列，写作指引备注里一句「改为两列」就放行——措辞层的
# 「表格/函件不会重排走样」拦不住（模型嫌填值后版面难看就自己重建表）。这里拿
# 节内表格比对原件：网格列数一致、原件标签文字还在、整张表没丢。擅自改格式件
# 是废标级，所以单节自查、全局校验、合册三道都跑这同一份核对。

_FF_TPL_ANY_RE = re.compile(r"TPL-0*(\d{1,3})", re.I)
_FF_RANGE_RE = re.compile(r"L\s*(\d{1,6})(?:\s*[-–~]\s*L?\s*(\d{1,6}))?")
_FF_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
_FF_HDR_SOURCE_RE = re.compile(r"^>\s*来源文件[：:]\s*`?([^`（(\n]+?)`?\s*(?:[（(。；]|$)")
_FF_MATERIAL_RE = re.compile(r"^(.+\.docx)\s+L(\d+)(?:-L?(\d+))?$", re.I)
_FF_MD_SEP_RE = re.compile(r"^:?-{2,}:?$")
_FF_FIX = ("处置：删掉重排的表，用 copy-from 把招标原件对应区间拷进节文件，再 revise 只往空格里填；"
           "版面难看不改表，加 Word 批注请用户手动处理；"
           "用户已在对话里明确同意改结构的，加 --allow-format-change \"<节名>\" 放行。")


def allow_format_set(raw: str) -> set[str]:
    """--allow-format-change 的节名清单 → 归一集合（节文件名主干，中英文逗号分号都认）。"""
    return {re.sub(r"\s+", "", x) for x in re.split(r"[,，;；]", raw or "") if x.strip()}


def _ff_md_cells(line: str) -> list[str]:
    s = line.strip()
    if len(s) < 2 or not (s.startswith("|") and s.endswith("|")):
        return []
    return [c.strip() for c in re.split(r"(?<!\\)\|", s[1:-1])]


def _ff_md_rows(text: str) -> list[list[str]]:
    return [cells for cells in (_ff_md_cells(ln) for ln in text.splitlines())
            if cells and not all(_FF_MD_SEP_RE.match(c) for c in cells if c)]


def _ff_md_tables(lines: list[str]) -> list[dict]:
    """解析 md 里的管道表：[{lo, hi, cols, rows, labels}]（行号 1 基，rows 含表头行）。
    合并格在解析 md 里按网格位重复出现（「电话 | 电话」），比对前去重即标签集合。"""
    tables: list[dict] = []
    cur: dict | None = None
    for i, raw in enumerate(lines, 1):
        cells = _ff_md_cells(raw)
        if not cells:
            cur = None
            continue
        if any(cells) and all(_FF_MD_SEP_RE.match(c) for c in cells if c):
            continue  # 表头分隔行；整行皆空的是空行数据行，要计入跨度与行数
        if cur is None:
            cur = {"lo": i, "hi": i, "cols": 0, "rows": 0, "labels": []}
            tables.append(cur)
        cur["hi"] = i
        cur["cols"] = max(cur["cols"], len(cells))
        cur["rows"] += 1
        for c in cells:
            t = re.sub(r"\s+", "", c)
            if len(t) >= 2 and _FF_CJK_RE.search(t):
                cur["labels"].append(t)
    return tables


def _ff_main_source(work: Path) -> str:
    try:
        data = json.loads((work / "parse" / "sources.json").read_text(encoding="utf-8"))
        main = data.get("main")
        if isinstance(main, list) and main:
            main = main[0]
        if isinstance(main, dict):
            main = main.get("path") or main.get("file") or ""
        return Path(str(main)).name if main else ""
    except Exception:
        return ""


def _ff_sole_source(work: Path) -> str:
    """解析产物只有一份时，它就是「同文件」所指（单来源项目没有来源确认单也成立）。"""
    try:
        found = [d.name for d in (work / "parse").iterdir()
                 if d.is_dir() and (d / (Path(d.name).stem + ".md")).is_file()]
    except Exception:
        return ""
    return found[0] if len(found) == 1 else ""


def ff_tpl_registry(work: Path) -> dict[str, dict]:
    """分析登记表模板表：TPL 编号 → {name, file, ranges}。

    编号=模板表内的行序（与 check_writing 的登记表读取同一约定：首列写「TPL-09」
    还是纯序号都按行序编号）。出处里的「同文件」沿用表内最近出现的文件名（种子=
    登记表头部「> 来源文件：」声明）；都没有时依次退到 sources.json 的主文件、唯一
    的那份解析产物——真机工作区实测有来源确认单缺失、头部没写文件名的，「同文件」
    解析不出文件会让整份核对静默失效。"""
    path = work / "analysis" / "requirements-format.md"
    if not path.is_file():
        return {}
    text = path.read_text(encoding="utf-8", errors="replace")
    seed = ""
    for ln in text.splitlines()[:12]:
        m = _FF_HDR_SOURCE_RE.match(ln.strip())
        if m:
            seed = m.group(1).strip()
            break
    fallback = _ff_main_source(work) or _ff_sole_source(work)
    out: dict[str, dict] = {}
    header: list[str] | None = None
    src_idx = name_idx = 0
    last_file = seed
    for r in _ff_md_rows(text):
        joined = "".join(r)
        has_src = any("出处" in c or "线索" in c for c in r)
        if "模板名称" in joined and has_src:
            header = r
            src_idx = next((i for i, c in enumerate(r) if "出处" in c or "线索" in c), len(r) - 1)
            name_idx = next((i for i, c in enumerate(r) if "模板名称" in c), 0)
            continue
        if header is None:
            continue
        if has_src and ("序号" in joined or len(r) != len(header)):
            header = None  # 另一张登记表的表头：关表，后面的行不属于模板表
            continue
        r = (r + [""] * len(header))[:len(header)]
        src = r[src_idx]
        fm = re.match(r"^(\S+?\.(?:docx|pdf|txt|md))(?=\s|[（(]|$)", src)
        if fm:
            last_file = fm.group(1)
        out[f"TPL-{len(out) + 1:02d}"] = {
            "name": r[name_idx],
            "file": last_file or fallback,
            "ranges": [(int(a), int(b or a)) for a, b in _FF_RANGE_RE.findall(src)],
        }
    return out


def ff_outline_tpl(md_text: str) -> list[tuple[str, str, list[str]]]:
    """目录「来源标注」：[(响应文件名, 叶子标题, [TPL 编号])]——用户在 G3 确认过的树，
    指引依据列漏记 TPL 时（真机事故里评分索引表就是）仍能查到节与原件的对应。"""
    out: list[tuple[str, str, list[str]]] = []
    vol, sec = "", None
    for raw in md_text.splitlines():
        line = raw.rstrip()
        if line.startswith("# 响应文件："):
            vol, sec = line[len("# 响应文件："):].strip(), None
        elif line.startswith("## "):
            sec = line[3:].strip()
        elif sec == "来源标注" and line.lstrip().startswith("- ") and " :: " in line:
            name, rest = line.strip()[2:].split(" :: ", 1)
            ids = [f"TPL-{int(n):02d}" for n in _FF_TPL_ANY_RE.findall(rest)]
            if ids:
                out.append((vol, name.strip(), ids))
    return out


def _ff_same_leaf(a: str, b: str) -> bool:
    return re.sub(r"\s+", "", a) == re.sub(r"\s+", "", b) or sanitize_name(a) == sanitize_name(b)


def format_fidelity(work: Path, allow: set[str] | None = None, only: Path | None = None) -> list[str]:
    """格式件保真核对：挂了 TPL 的节，节内表格须与招标原件的格式表同一网格。
    返回问题清单（空=通过）；只读不落盘。原件比对依据是解析 md（不依赖招标 docx
    还在不在）：素材列给了同文件更窄区间时按交集核对（一个 TPL 登记多个备选模板、
    本节只取其一时不误报）；解析产物不在时无从比对，跳过不拦。"""
    work = Path(work)
    body_dir = work / "body"
    outline_path = work / "outline" / "tender-response-docs.md"
    if not body_dir.is_dir() or not outline_path.is_file():
        return []
    registry = ff_tpl_registry(work)
    if not registry:
        return []
    outline_text = outline_path.read_text(encoding="utf-8", errors="replace")
    volumes = parse_outline(outline_text)
    single = len(volumes) == 1
    out_rows = ff_outline_tpl(outline_text)
    guide_rows = [r for r in _ff_md_rows(_ff_read(body_dir / "写作指引.md")) if r[0].strip() != "节"]
    allow = allow or set()
    md_tables: dict[str, list[dict]] = {}
    issues: list[str] = []
    if only is not None:
        sections = [Path(only)]
    else:
        sections = sorted(p for p in list(body_dir.glob("*.docx")) + list(body_dir.glob("*/*.docx"))
                          if not p.name.startswith(("整本-", "~$", ".~")))
    for p in sections:
        if not p.is_file():
            continue
        stem = p.stem
        in_vol = p.parent != body_dir
        vol_dir = p.parent.name if in_vol else ""
        if re.sub(r"\s+", "", stem) in allow or re.sub(r"\s+", "", f"{vol_dir}/{stem}") in allow:
            continue
        row = next((r for r in guide_rows if _ff_same_leaf(r[0].rsplit("/", 1)[-1], stem)), None)
        # 节与原件的对应只认目录「来源标注」（用户在目录确认门确认过的树；模板编号没被
        # 任何节引用时 check_outline 会报警）。指引依据列是模型写的，错挂 TPL 会把无关
        # 的节拖进来误报（测试夹具的封面行挂了评分索引表的模板）
        ids: list[str] = []
        for ovol, otitle, oids in out_rows:
            if ((single and not in_vol) or (not single and sanitize_name(ovol) == vol_dir)) \
                    and _ff_same_leaf(otitle, stem):
                ids += oids
        hits: list[tuple[dict, str, str]] = []  # (原件表, TPL 编号, 模板名)
        for tid in dict.fromkeys(ids):
            entry = registry.get(tid)
            if not entry or not entry["ranges"] or not entry["file"]:
                continue
            fname = entry["file"]
            md = work / "parse" / fname / (Path(fname).stem + ".md")
            if not md.is_file():
                continue
            if fname not in md_tables:
                md_tables[fname] = _ff_md_tables(md.read_text(encoding="utf-8", errors="replace").splitlines())
            material = []
            for tok in ((row[3] if row is not None else "") or "").split("，"):
                mm = _FF_MATERIAL_RE.match(tok.strip())
                if mm and re.split(r"[/\\]", mm.group(1))[-1] == fname:
                    material.append((int(mm.group(2)), int(mm.group(3) or mm.group(2))))
            eff: list[tuple[int, int]] = []
            for lo, hi in entry["ranges"]:
                cuts = [(max(lo, a), min(hi, b)) for a, b in material if a <= hi and b >= lo]
                eff += cuts or [(lo, hi)]
            for t in md_tables[fname]:
                for lo, hi in eff:
                    # 表头行落在区间内（留一行余量），或区间盖过半张表——算这张原件表的账
                    covered = min(t["hi"], hi) - max(t["lo"], lo) + 1
                    if covered >= 1 and (t["lo"] >= lo - 1 or covered * 2 >= t["hi"] - t["lo"] + 1):
                        hits.append((t, tid, entry["name"]))
                        break
        if not hits:
            continue
        doc = Document(str(p))
        sec_tabs: list[tuple[int, int]] = []
        sec_texts: set[str] = set()
        for t in doc.tables:
            grid = t._tbl.find(qn("w:tblGrid"))
            try:
                cols = len(grid.findall(qn("w:gridCol"))) if grid is not None else len(t.columns)
            except Exception:
                cols = 0
            sec_tabs.append((cols, len(t.rows)))
            seen: list = []
            for table_row in t.rows:
                for cell in table_row.cells:
                    if any(cell._tc is tc for tc in seen):
                        continue
                    seen.append(cell._tc)
                    sec_texts.add(re.sub(r"\s+", "", cell_text_full(cell)))
        tag = "、".join(dict.fromkeys(f"{tid} {name}".strip() for _, tid, name in hits))
        origins = "、".join(f"L{t['lo']}" + (f"-L{t['hi']}" if t["hi"] != t["lo"] else "")
                           + f"（{t['rows']} 行 × {t['cols']} 列）" for t, _, _ in hits)
        if not sec_tabs:
            issues.append(f"{p.name} 是招标格式件（{tag}），原件 {origins} 有表格，节内没有表格"
                          f"——格式件的表只许 copy-from 拷原件后往空格里填，不许删掉或换成段落、图片")
            continue
        shape = "、".join(f"{r} 行 × {c} 列" for c, r in sec_tabs)
        for t, tid, name in hits:
            if any(c == t["cols"] for c, _ in sec_tabs):
                labels = list(dict.fromkeys(t["labels"]))
                kept = [x for x in labels if any(x in s for s in sec_texts)]
                if labels and len(kept) * 2 < len(labels):
                    issues.append(f"{p.name} 招标格式件表格列数与原件一致，但原件标签只保留 "
                                  f"{len(kept)}/{len(labels)}（{tid} {name}）——核对是否重建了表")
                continue
            lost = [x for x in dict.fromkeys(t["labels"]) if not any(x in s for s in sec_texts)]
            issues.append(f"{p.name} 招标格式件表格与原件不一致（{tid} {name}）：原件 L{t['lo']}-L{t['hi']} "
                          f"{t['rows']} 行 × {t['cols']} 列，节内表格 {shape}"
                          + (f"；原件标签缺失：{'、'.join(lost[:8])}" if lost else "")
                          + "——行列与合并格一个不动，只许 copy-from 拷原件后往空格里填")
    return issues


def _ff_read(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace") if path.is_file() else ""
    except Exception:
        return ""


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
    flags = [t in landscape_titles for (t, _, _, _, _) in chapter_plan]
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
    分节符并入目录块结束锚点段的 pPr（不另插独立空段——空段会在目录页近满时
    溢出成无页脚空白页，既多一张白页又把 toc-pages 的页码推算顶歪一位，实测
    差 1 缺陷根因）：封面/前置/目录同属前置节，目录页也不带页码。零引用前置
    节 + default 页脚正文节的组合沿自实测矩阵 m5（比 titlePg 稳：LibreOffice
    下 titlePg 节后跟普通节会页码错位 m1/m3）。无目录锚点时返回 False（调用
    方退回 titlePg 方案，此时页码从封面连续编）。"""
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
    ppr = anchor.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        anchor.insert(0, ppr)
    ppr.append(sect)
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
        return _gov_label(depth, counters)
    # decimal
    return ".".join(str(counters[k]) for k in range(1, depth + 1)) + " "


def _remap_section_ids(doc: Document, taken: set[str]) -> tuple[set[str], int]:
    """合册前把与本册已占用批注号冲突的锚就地换号（复制节文件当模板的撞号形态：
    两个节文件同号时 _merge_comments 只并一条体，第二节锚挂着第一节的文字）。
    锚与批注体同改、落回节文件，_merge_comments 读盘自然一致。
    返回 (更新后的占用集, 换号数)。"""
    ids = {e.get(qn("w:id")) for e in doc.element.body.iter(qn("w:commentReference"))}
    mapping: dict[str, str] = {}
    for cid in sorted(ids):
        if cid in taken:
            new = str(random.randint(1_000_000, 99_999_999))
            while new in taken or new in ids or new in mapping.values():
                new = str(random.randint(1_000_000, 99_999_999))
            mapping[cid] = new
    if mapping:
        for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
            for e in doc.element.body.iter(qn(tag)):
                if e.get(qn("w:id")) in mapping:
                    e.set(qn("w:id"), mapping[e.get(qn("w:id"))])
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/comments"):
                for cbody in _comments_root(rel.target_part).findall(qn("w:comment")):
                    if cbody.get(qn("w:id")) in mapping:
                        cbody.set(qn("w:id"), mapping[cbody.get(qn("w:id"))])
    taken |= ids - set(mapping)
    taken |= set(mapping.values())
    return taken, len(mapping)


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
        die(f"目录文件未解析出任何响应文件（核对「# 响应文件：X」标题与树格式）")
    body_dir = work / "body"
    # 格式件保真（入口红线 8）：招标给了格式的表只许拷原件填空——节内表格与原件对不上
    # 就拒绝合册。缺料类问题带批注照常合册，废标级红线不在此列；核对自身出错只告警
    # （保真核对是防线，不许反过来变成交付故障点）
    try:
        broken = format_fidelity(work, allow=allow_format_set(getattr(args, "allow_format_change", "")))
    except Exception as exc:  # noqa: BLE001
        broken = []
        print(f"  ⚠️ 格式件保真核对未能执行（{type(exc).__name__}），请人工确认格式件表格与招标原件一致")
    if broken:
        print("[合册被拦] 招标格式件表格与原件不一致——擅自改格式件是废标级红线，先修复再合册：")
        for b in broken:
            print(f"- {b}")
        print(f"  {_FF_FIX}")
        sys.exit(2)
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
        container_files: set[Path] = set()  # 容器章同名节文件（容器分支已点名，孤儿扫描不再重复报）
        taken_ids: set[str] = set()  # 本册已占用的批注号（跨节撞号就地重排）
        # 同名叶子点名：不同章节下的同名节共用同一节文件（body/ 按 <节名>.docx 找），
        # 内容会在整本里重复出现——目录树设计常见误区，点名让用户确认
        name_counts: dict[str, int] = {}
        for n in vol["nodes"]:
            if not n["children"]:
                name_counts[n["text"]] = name_counts.get(n["text"], 0) + 1
        for k, v in name_counts.items():
            if v > 1:
                warnings.append(f"目录树中有 {v} 个同名节「{k}」——不同章节下的同名节共用同一节文件，"
                                f"内容会重复出现，确认目录树设计是否如此")
        in_front = True  # 目录节点之前=前置区
        toc_marker = None
        merged = 0

        def _emit_node(node, depth):
            nonlocal in_front, toc_marker, merged, taken_ids
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
                # 容器章不并入节文件——同名节文件存在即为正文丢失前兆，专项点名
                # （只报「节点是容器」，不报「文件不在树中」：文件与目录逐字一致时
                # 走孤儿分支会把排查引向文件名，真机实测绕了六轮）
                sf0 = _section_file(body_dir, vol["name"], title, multi)
                if sf0:
                    container_files.add(sf0)
                    warnings.append(f"容器节点「{title}」存在同名节文件 {sf0.name}——"
                                    f"容器章不并入节文件，该节正文会丢：把其子节提升为"
                                    f"平级使本节点成叶子，或移除该节文件")
                if in_front:
                    # 前置区分组无章标题——不入编号与目录，仅分组提示
                    warnings.append(f"前置区容器「{title}」不占章号，其子节按前置区处理")
                    for child in node["children"]:
                        _emit_node(child, depth + 1)
                    return
                _append_container_heading(master, title, num,
                                          page_break=(depth == 1 and bool(num)))
                chapter_plan.append((title, depth, num, in_front,
                                     str(counters[1]) if counters.get(1) else ""))
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
                # 节文件可能被人工在 Word 里开着修订改过：整本是递交稿，痕迹与悬空锚
                # 在此统一接受/清理并落回节文件（幂等——重新合册不再重复点名）
                st = _sanitize_doc(sub_doc)
                # 跨节批注号撞号：冲突锚就地换号（体同改），落盘后 _merge_comments 读盘一致
                n_remap = 0
                if sub_doc.element.body.findall(".//" + qn("w:commentReference")):
                    taken_ids, n_remap = _remap_section_ids(sub_doc, taken_ids)
                if any(st.values()) or n_remap:
                    sub_doc.save(str(sf))
                if any(st.values()):
                    warnings.append(f"节文件「{title}」含修订痕迹/批注锚——已接受修订并清理"
                                    f"（{_fmt_stats(st)}），整本按清理后内容合并")
                composer.append(sub_doc)
                used_files.add(sf)
            else:
                warnings.append(f"缺失节：{title}（{'前置页' if in_front else '章节'}，整本中留空位）")
            # 前置节（封面之后的评分索引表/报价函等格式件）每节独立起页——
            # Word 与 LibreOffice 度量不同，封面溢行时格式件会贴到上一页
            # （2026-11-21 实测），段前分页对齐正文章的独立起页惯例
            if not _apply_chapter_number(master, title, num,
                                         page_break=(in_front and merged > 0)
                                                     or (depth == 1 and not in_front and bool(num))):
                if exists:
                    warnings.append(f"节文件内标题与目录不一致（目录：{title}）——编号未应用，"
                                    f"改节文件标题与目录逐字一致后重新合册")
            chapter_plan.append((title, depth, num, in_front,
                                 str(counters[1]) if counters.get(1) else ""))
            toc_entries.append(((num + title).strip(), depth, not exists))
            merged += 1

        for top in vol["tree"]:
            _emit_node(top, 1)
        toc_end = None
        if toc_marker is not None:
            toc_end = _fill_toc(master, toc_marker, toc_entries, args.toc_depth)
        _restyle_all(master, chapter_plan, args.numbering, args.flat_outline)
        _apply_body_indents(master)
        _clear_table_cell_indents(master)
        font_notes = _solidify_fonts(master)
        _apply_landscape_regions(master, chapter_plan, landscape_titles)
        # 孤儿节文件点名（多册只扫本册子目录，单册扫 body 根）
        if not any(n["text"] == "目录" for n in vol["nodes"]):
            warnings.append("目录树中无「目录」节点——未生成目录页，且全文按前置区处理（无章节编号）")
        scan_dir = body_dir / sanitize_name(vol["name"]) if multi else body_dir
        for f in sorted(scan_dir.glob("*.docx")):
            # ~$节.docx / .~节.docx 是 Word/WPS 打开节文件时落的占用锁文件（隐藏），不是产物
            if f.name.startswith(("整本-", "~$", ".~")) or f in used_files or f in container_files:
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
        if font_notes:
            print("  [字体固化] " + "；".join(font_notes))
        print("  对用户口径：整本 Word 文档已生成（「合册」是内部动作，"
              "不进对话、不作提问——reply-style 硬规则 9）。")
        for w in warnings:
            print(f"  ⚠️ {w}")
        warnings = []
        toc_entries = []


def _shrink_trailing_blank(master: Document) -> None:
    """册尾连续空段全压成 2pt 定行高：表格结尾时 Word 规范强制保留一个段落，
    满行高会在内容恰好填满时挤出空白尾页（交付观感）。素材尾部常带连续空段，
    只压末段压不干净——逐个向前压，遇内容/图/批注引用停；带分节符的段不压
    （重建 pPr 会把 w:sectPr 一起剥掉、吞分节）。"""
    body = master.element.body

    def shrink(p_el) -> None:
        ppr = OxmlElement("w:pPr")  # 整体重建 pPr：spacing→rPr 顺序 schema 安全
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "40")  # 2pt 定值行高
        spacing.set(qn("w:lineRule"), "exact")
        ppr.append(spacing)
        rpr = OxmlElement("w:rPr")
        for tag in ("w:sz", "w:szCs"):
            el = OxmlElement(tag)
            el.set(qn("w:val"), "4")  # 2pt
            rpr.append(el)
        ppr.append(rpr)
        old = p_el.find(qn("w:pPr"))
        if old is not None:
            p_el.remove(old)
        p_el.insert(0, ppr)

    for child in reversed(list(body)):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag != qn("w:p"):
            return  # 最后的实体元素不是段落——无册尾空段问题
        if any((t.text or "").strip() for t in child.findall(".//" + qn("w:t"))):
            return  # 末段有内容
        if child.findall(".//" + qn("w:drawing")) or child.findall(".//" + qn("w:commentReference")):
            return  # 末段带图/批注引用，不动
        old_ppr = child.find(qn("w:pPr"))
        if old_ppr is not None and old_ppr.find(qn("w:sectPr")) is not None:
            return  # 分节符段落：动它=吞分节
        shrink(child)  # 压掉后继续向前，连续空段全压


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


def _fill_toc(master: Document, marker, entries: list[tuple[str, int, bool]],
              depth_limit: int = 3):
    """在目录标题段之后插入静态条目（TOC 域包裹；条目带点线制表位——
    未更新域前也有目录观感，Word/WPS 打开经 updateFields 刷新出真页码）。
    depth_limit 控制条目层级与域刷新口径（\\o "1-N"，静态条目同深度截断）。
    返回目录块结束锚点=视觉末条目段（域 end 并入该段）——合册分节符直接挂在
    它身上，目录页归前置节、不带页码；不产生任何尾随空段（空段会在目录页
    近满时溢出成无页脚空白页，把 toc-pages 的页码推算顶歪一位——实测教训）。"""
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.text.paragraph import Paragraph

    if marker is None or not entries:
        return None
    entries = [e for e in entries if e[1] <= depth_limit]
    if not entries:
        return None

    def _insert_after(el, ref):
        ref.addnext(el)
        return el

    p_first = p_last = None  # 循环按倒序插入：末次插入=视觉首条目，首次插入=视觉末条目
    for text, depth, missing in reversed(entries):
        p = _insert_after(OxmlElement("w:p"), marker)
        entry = Paragraph(p, master)
        toc_st = _get_style(master, f"TOC {depth}")
        if toc_st is not None:
            entry.style = toc_st  # 独立样式：单倍行距+按级缩进，版式压缩跳过它
        else:  # 用户模板无 TOC 样式族 → 直排格式兜底
            entry.paragraph_format.first_line_indent = Pt(0)
            entry.paragraph_format.left_indent = Cm(0.74 * max(0, depth - 1))
        entry.paragraph_format.tab_stops.add_tab_stop(
            Cm(14.4), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        label = text + ("（另附）" if missing else "") + "\t"
        run = entry.add_run(label)
        if toc_st is None:
            style_run(run, "宋体", 12, depth == 1)
        if p_last is None:
            p_last = p
        p_first = p
    # TOC 域包裹静态清单：begin+instr+separate 插到第一条目的 pPr 之后。
    # 域指令不带 \\h——LibreOffice/WPS 会把目录条目渲染成下划线超链接（真机坑）；
    # 刷新口径与静态条目同截断到 depth_limit，F9 刷新不丢已显示的层级。
    first = p_first
    base = 1 if first.find(qn("w:pPr")) is not None else 0
    fld = OxmlElement("w:r")
    fldc = OxmlElement("w:fldChar")
    fldc.set(qn("w:fldCharType"), "begin")
    fld.append(fldc)
    instr = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f' TOC \\o "1-{depth_limit}" \\z \\u '
    instr.append(it)
    sep = OxmlElement("w:r")
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "separate")
    sep.append(fc2)
    first.insert(base, sep)
    first.insert(base, instr)
    first.insert(base, fld)
    # 域 end 并入视觉末条目段（不再产独立空段——见函数 docstring）
    endrun = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "end")
    endrun.append(fc3)
    p_last.append(endrun)
    return p_last


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


def _restyle_all(master: Document, chapters: list, scheme: str,
                 flat_outline: bool = False) -> None:
    """单遍重排：depth≥2 章标题降一级；节内小标题按树深度降级 + 接续节号。
    导航大纲默认继承 Heading 样式层级（导航窗格出树形、TOC 域刷新不丢子级）；
    flat_outline=True 时直写 outlineLvl=9 摘出导航（导航降噪的特殊场景，显式
    选入——代价是导航全扁平、F9 刷目录丢二级以下条目）。扁平目录（节直接挂
    册下，depth=1）的节内小标题同样编号（10.1 式——chapter 方案的章号是
    「第十章」汉字形态，前缀须换 counters 的阿拉伯章号拼）。章标题与计划错位
    （节文件缺失/被删，整本中留空位）时按标题在剩余计划内向后对位，跳过缺章
    继续接续小标题编号；完全对不上才中断该章的小标题处理。"""
    idx = 0
    cur: tuple[str, int, bool, str] | None = None  # (章号, 树深, 前置区, 阿拉伯章号)
    last_chap = -1
    sub_counters: dict[int, int] = {}
    for p in master.paragraphs:
        try:
            sname = p.style.name if p.style is not None else ""
        except Exception:
            sname = ""
        if sname in ("Heading 1", "标题 1"):
            hit = -1
            if idx < len(chapters) and _chap_match(p.text, chapters[idx][0]):
                hit = idx
            else:
                for j in range(idx + 1, len(chapters)):
                    if _chap_match(p.text, chapters[j][0]):
                        hit = j
                        break
            if hit >= 0:
                title, depth, num, in_front, arab = chapters[hit]
                if hit != last_chap:
                    sub_counters = {}
                    last_chap = hit
                idx = hit + 1
                cur = (num, depth, in_front, arab)
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
        num, depth, in_front, arab = cur
        src_level = int(sname.split()[-1]) if sname[-1].isdigit() else 2
        target = max(2, min(depth + src_level - 1, 9))
        try:
            p.style = master.styles[f"Heading {min(target, 4)}"]
        except KeyError:
            pass
        ppr = p._p.get_or_add_pPr()
        for old_lvl in ppr.findall(qn("w:outlineLvl")):
            ppr.remove(old_lvl)  # copy-from 带入的直写层级一并清掉，统一继承样式
        if flat_outline:
            lvl = OxmlElement("w:outlineLvl")
            lvl.set(qn("w:val"), "8")  # 0 基：第 9 级=正文级，不进导航窗格/TOC 域
            ppr.append(lvl)
        if in_front or scheme == "none" or target > 4 or SELF_NUMBERED.match(p.text):
            continue
        sub_counters[target] = sub_counters.get(target, 0) + 1
        for k in [k for k in sub_counters if k > target]:
            sub_counters[k] = 0
        if scheme == "gov":
            # gov 小标题按公文层级续编（第 depth 层章下的标题取 gov 第 target 层形态）
            _rewrite_para_text(p, f"{_gov_label(target, sub_counters)}{p.text}")
            continue
        sub = ".".join(str(sub_counters[k]) for k in range(depth + 1, target + 1))
        prefix = arab if (depth == 1 and scheme == "chapter" and arab) else num.strip()
        _rewrite_para_text(p, f"{prefix}.{sub} {p.text}" if sub else f"{prefix} {p.text}")


def _merge_comments(master: Document, used_files: set[Path]) -> int:
    """docxcompose 迁移正文批注引用；comments part 在此合并（id 创建时全局唯一）。
    只并有锚点的批注体——源文件里锚点已被 revise 抹掉的孤儿体不迁移（真机实测
    11 条孤儿体由此进整本），孤儿清理走 comment-remove --orphan 在节文件层做。"""
    existing = doc_comments(master)
    for f in used_files:
        try:
            fdoc = Document(str(f))
        except Exception:
            continue
        anchored = _anchor_ids(fdoc.element.body)
        for cid, text in doc_comments(fdoc).items():
            if cid in existing or cid not in anchored:
                continue
            existing[cid] = text
    if not existing:
        return 0
    part = _comments_part(master)
    root = _comments_root(part)
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
    _comments_flush(part, root)
    return len(doc_comments(master))


# ---------------------------------------------------------------- toc-pages（静态目录条目补实测页码）

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


def _lo_convert(soffice: str, docx_path: Path, outdir: Path, profile: Path) -> Path:
    import subprocess

    # 配置目录地址必须由 as_uri 生成：手工拼 file://+路径 在 Windows 上是坏地址（盘符
    # 冒号、反斜杠），路径含空格/# 也不转义——地址坏了 LibreOffice 就退回默认配置，
    # 界面版开着时转换被转交给它、本进程拿不到产物
    cmd = [soffice, "--headless", f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
           "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)]
    try:
        # errors="replace"：text=True 缺省严格解码，输出含本机代码页外的字节就抛异常
        r = subprocess.run(cmd, capture_output=True, text=True, errors="replace", timeout=300)
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
    try:
        return _heading_pages_in(pdf, labels, restart)
    finally:
        pdf.close()  # 异常路径也关：Windows 下留着句柄，临时目录清理会失败并盖掉真正的报错


def _heading_pages_in(pdf, labels: list[str], restart: bool) -> dict[str, int]:
    pages_lines: list[list[str]] = []
    has_leader: list[bool] = []
    for i in range(len(pdf)):
        tp = pdf[i].get_textpage()
        # get_text_range 缺省参数会被 pypdfium2 重定向并打 UserWarning（stderr 噪声）；
        # 新版显式走 get_text_bounded，旧版没有该方法再回落
        txt = (tp.get_text_bounded() if hasattr(tp, "get_text_bounded") else tp.get_text_range()) or ""
        tp.close()
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
    # 目录后的无字空白页（前置节尾空段溢出残留）不计入正文页号——正文首页是
    # 目录之后第一张有内容的页；防御式跳过，正常经 _fill_toc 修复后无空白页
    while start < len(pages_lines) and not any(pages_lines[start]):
        start += 1
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


_CHAP_REF = re.compile(r"第([0-9零一二三四五六七八九十百]+)章")
_CN_DIGITS = "零一二三四五六七八九"


def _zh_int(s: str) -> int | None:
    """章号归一：第一章/第1章/第二十一章 → 1/1/21（百以内汉字 + 阿拉伯数字）。"""
    s = s.strip()
    if s.isdigit():
        return int(s) if int(s) <= 9999 else None
    if not s or any(ch not in _CN_DIGITS + "十百" for ch in s):
        return None
    if "百" in s:
        h, _, rest = s.partition("百")
        head = _zh_int(h) if h else 1
        tail = _zh_int(rest) if rest else 0
        return head * 100 + tail if head is not None and tail is not None else None
    if "十" in s:
        t, _, rest = s.partition("十")
        head = _zh_int(t) if t else 1
        tail = _zh_int(rest) if rest else 0
        return head * 10 + tail if head is not None and tail is not None else None
    return _CN_DIGITS.index(s) if len(s) == 1 else None


def _xref_audit(docx_path: Path, mapping: dict[str, int]) -> None:
    """表格交叉引用核对（只报事实不改文件）：表内「第X章」与目录实测章号对账。

    评分索引表/业绩表引用正文章号，正文一动全要手改；章号写错或章节增删后未
    同步，引用会指向不存在的章（真机：表写第 20 章，实际第 21 章起）。目录实测
    映射在手，顺手对账。章号 scheme 非 chapter 时映射里无「第X章」前缀，跳过。"""
    chapters: dict[int, str] = {}
    for label in mapping:
        m = _CHAP_REF.match(label)
        if m:
            n = _zh_int(m.group(1))
            if n is not None:
                chapters.setdefault(n, label)
    if not chapters:
        return
    doc = Document(str(docx_path))
    hits: set[tuple[int, str]] = set()
    for ti, tbl in enumerate(doc.tables, 1):
        for row in tbl.rows:
            seen: set[int] = set()
            for cell in row.cells:
                if id(cell._tc) in seen:  # 合并格只算一次
                    continue
                seen.add(id(cell._tc))
                for m in _CHAP_REF.finditer(cell_text_full(cell)):
                    n = _zh_int(m.group(1))
                    if n is None or n in chapters:
                        continue
                    hits.add((ti, m.group(0)))
    for ti, ref in sorted(hits):
        m = _CHAP_REF.match(ref)
        ref_n = _zh_int(m.group(1)) if m else 0
        near = min(chapters, key=lambda c: abs(c - ref_n))
        label = chapters[near]
        print(f"  ⚠️ 表格 T{ti} 引用「{ref}」——目录实测无此章"
              f"（最近：{label}，页 {mapping.get(label, '?')}），请核对")


def _tempdir():
    """工作用临时目录。Windows 下目录里只要还有一个没关的文件句柄（或刚退出的
    LibreOffice 还占着配置文件），清理就抛 PermissionError（WinError 32）并盖掉
    真正的报错——清理失败只留目录不抛，系统临时目录自会回收。"""
    import tempfile
    return tempfile.TemporaryDirectory(ignore_cleanup_errors=True)


def cmd_toc_pages(args) -> None:
    docx_path = Path(args.docx)
    if not docx_path.is_file():
        die(f"文件不存在：{docx_path}")
    soffice = args.soffice or _find_soffice()
    if not soffice or not Path(soffice).exists():
        print("[目录页码跳过] 环境未找到 LibreOffice——目录页码由 Word/WPS 打开合册文件时"
              "自动刷新，交付 PDF 用 Word/WPS 导出即可，本步可忽略"
              "（已安装但装在非默认位置：加 --soffice 指到 soffice 可执行文件）")
        return
    labels = _toc_entry_labels(docx_path)
    if not labels:
        die("未找到目录条目（先 assemble 生成目录页）")
    original = docx_path.read_bytes()
    body_before = _body_signature(docx_path)

    def _rollback(msg: str) -> None:
        docx_path.write_bytes(original)
        die(msg)

    restart = _has_pagenum_restart(docx_path)
    if restart:
        print("  页码口径：前置（封面/目录）不编页码，目录条目写正文页号（自 1 起）")
    with _tempdir() as td:
        tmp = Path(td)
        profile = tmp / "profile"
        # 收敛闭环：实测→回填→复测。页码回填会改目录占行，正文跟着再漂——
        # 两轮实测一致才算收敛；轮内漂移回显条数；写后复测仍漂移即回滚报警，
        # 不静默输出错页码（真机来回漂三次才收敛，撞 3 轮上限曾无声通过）。
        mapping_prev: dict[str, int] | None = None
        converged = False
        rounds = 0
        for i in range(args.max_rounds):
            rounds = i + 1
            pdf = _lo_convert(soffice, docx_path, tmp, profile)
            mapping = _measure_heading_pages(pdf, labels, restart)
            if mapping == mapping_prev:
                converged = True
                break
            if mapping_prev is not None:
                drift = sum(1 for k in mapping if mapping_prev.get(k) != mapping[k])
                print(f"  第 {rounds} 轮实测：{drift} 条页码仍在漂移，回填后再测")
            mapping_prev = mapping
            _, failed = _write_toc_numbers(docx_path, mapping)
            if failed:
                _rollback(f"自检失败：{failed} 条目录条目制表位无法重写，已回滚到运行前内容")
        if not converged and mapping_prev is not None:
            pdf = _lo_convert(soffice, docx_path, tmp, profile)  # 末轮回填后的写后复测
            converged = _measure_heading_pages(pdf, labels, restart) == mapping_prev
            rounds += 1
        if not converged:
            _rollback(f"页码未收敛（{rounds} 轮实测仍在漂移）——已回滚到运行前内容；"
                      f"先固定目录长度（目录条目数/版式稳定）再重跑，"
                      f"或评估加大 --max-rounds（当前 {args.max_rounds}）")
        if mapping_prev is None:
            mapping_prev = {}
        print(f"[目录页码] 实测收敛（{rounds} 轮）")
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

            _d = pdfium.PdfDocument(str(pdf))
            print(f"[页数] 整本共 {len(_d)} 页（{docx_path.name}）")
            _d.close()  # 不留给解释器回收（回收时打 Cannot close 噪声）
        except Exception:
            pass
        print(f"[目录页码] 已按渲染实测写入 {len(labels) - len(not_found)}/{len(labels)} 条")
        _xref_audit(docx_path, mapping_prev)
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
        raise argparse.ArgumentTypeError(f"段落号须为整数：{v}")
    if n < 1:
        raise argparse.ArgumentTypeError(f"段落号从 1 起：{v}")
    return n


def _positive_int(v: str) -> int:
    """正整数参数（收敛轮数等——0 轮无从谈起）。"""
    try:
        n = int(v)
    except ValueError:
        raise argparse.ArgumentTypeError(f"须为整数：{v}")
    if n < 1:
        raise argparse.ArgumentTypeError(f"至少 1：{v}")
    return n


def main() -> None:
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(errors="replace")  # 中文 Windows 管道（GBK）下 ⚠️ 会编码崩，降级输出
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
    sp.add_argument("--drop-comments", action="store_true",
                    help="覆盖时丢弃已有批注（缺省拦截：批注是待办凭证，不许静默丢）")
    sp.set_defaults(func=cmd_create)

    sp = sub.add_parser("read", help="文本视图")
    sp.add_argument("path")
    sp.set_defaults(func=cmd_read)

    sp = sub.add_parser("revise", help="定向修订（直接生效，无修订标记）")
    sp.add_argument("path")
    sp.add_argument("--para", type=_para_no, help="段落号 P#（也接受 3）")
    sp.add_argument("--cell", help="表格格 T1R2C3；配 insert_after 在该格所在行后插空行（克隆网格，格式件空行不够时增行用）")
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

    sp = sub.add_parser("comment-list", help="列出批注（id + 定位 + 全文，销号回删前核对）")
    sp.add_argument("path")
    sp.set_defaults(func=cmd_comment_list)

    sp = sub.add_parser("comment-remove", help="删批注（销号）——--id 可重复 / --para / --cell / --all / --orphan")
    sp.add_argument("path")
    sp.add_argument("--id", action="append", help="批注号 #N（comment-list 回显），可重复")
    sp.add_argument("--para", type=_para_no, default=None, help="删该段上的批注")
    sp.add_argument("--cell", help="删该格上的批注（T1R2C3）")
    sp.add_argument("--all", action="store_true", help="删全部批注")
    sp.add_argument("--orphan", action="store_true", help="只删孤儿批注体（正文锚点已被 revise 抹掉的）")
    sp.set_defaults(func=cmd_comment_remove)

    sp = sub.add_parser("image", help="插入图片（全宽居中）")
    sp.add_argument("path")
    sp.add_argument("--img", required=True)
    sp.add_argument("--after-para", type=_para_no, default=None)
    sp.add_argument("--caption", default=None, help="图下题注（居中 9pt，Table Caption 样式）")
    sp.set_defaults(func=cmd_image)

    sp = sub.add_parser("render-page", help="PDF 页转 PNG（--pages 批量）")
    sp.add_argument("pdf")
    sp.add_argument("page", nargs="?", type=int, default=0)
    sp.add_argument("--pages", default="",
                    help="批量：'3-7' 或 '3,5,7'——输出 <out主干>-pNNN.png")
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
                    help="spec JSON：gantt={title?,unit?,tasks:[{name,start,end,milestone}]}（title 可省——"
                         "配 image --caption 时留空防图内外标题重复；unit 默认 周，start/end 单位随 unit）；"
                         "layered={title?,layers:[{name,items:[]}]}（自上而下）")
    sp.set_defaults(func=cmd_diagram)

    sp = sub.add_parser("toc-pages", help="给合册文件的静态目录条目写入实测页码（需 LibreOffice）")
    sp.add_argument("docx", help="合册产物 docx（整本-<册名>.docx）")
    sp.add_argument("--pdf", default=None, help="顺带导出预览 PDF 到该路径（可选）")
    sp.add_argument("--map", default=None,
                    help="导出「标题→页码」JSON 映射（评分索引表等页码回填用）")
    sp.add_argument("--max-rounds", type=_positive_int, default=5,
                    help="页码收敛上限（实测→回填→复测的闭环轮数；未收敛回滚报警不静默）")
    sp.add_argument("--soffice", default=None,
                    help="LibreOffice 可执行文件路径（缺省自动探测；找不到则优雅跳过）")
    sp.set_defaults(func=cmd_toc_pages)

    sp = sub.add_parser("assemble", help="按目录树合册")
    sp.add_argument("--work", default="work")
    sp.add_argument("--numbering", choices=["chapter", "decimal", "gov", "none"],
                    default="chapter",
                    help="章节编号形态：chapter=第X章/1.1；decimal=1/1.1；"
                         "gov=一、/（一）/1.；none=不编")
    sp.add_argument("--toc-depth", type=int, choices=[1, 2, 3, 4], default=3,
                    help="目录层级（静态条目与 TOC 域刷新口径一致，默认 3 级）")
    sp.add_argument("--flat-outline", action="store_true",
                    help="节内小标题摘出导航大纲（outlineLvl=9）：导航窗格降噪，代价是"
                         "导航全扁平、F9 刷目录丢二级以下条目——默认关闭")
    sp.add_argument("--template", help="版式模板 docx（缺省用包内基准模板；用户模板的"
                                       "页面设置与页眉页脚原样保留）")
    sp.add_argument("--allow-format-change", default="",
                    help="用户已在对话里明确同意改表格式的节名，逗号分隔（节文件名主干）——"
                         "格式件保真核对对这些节放行；擅自改招标格式件是废标级红线，无用户同意不得使用")
    sp.set_defaults(func=cmd_assemble)

    sp = sub.add_parser("compact", help="正文段固定行距压缩（大字号/含图/目录段保护常开）")
    sp.add_argument("--docx", required=True, help="目标 docx（整本或节文件）")
    sp.add_argument("--line-pt", type=float, default=22.0,
                    help="固定行距磅值（默认 22pt；exact 行距，正文页数压缩用）")
    sp.set_defaults(func=cmd_compact)

    sp = sub.add_parser("image-scan", help="导出内嵌图片缩略图网格（人工核方向，记序号）")
    sp.add_argument("--docx", required=True, help="目标 docx")
    sp.add_argument("--out", default=None, help="网格图输出目录（缺省 docx 同目录）")
    sp.add_argument("--thumb", type=int, default=160, help="缩略图最大边像素（默认 160）")
    sp.add_argument("--cols", type=int, default=6, help="每行列数（默认 6）")
    sp.set_defaults(func=cmd_image_scan)

    sp = sub.add_parser("image-rotate", help="按序号旋转图片并重算引用处 extent（序号见 image-scan）")
    sp.add_argument("--docx", required=True, help="目标 docx")
    sp.add_argument("--index", type=_positive_int, required=True, help="图片序号 #N（image-scan 回显）")
    sp.add_argument("--deg", type=int, choices=[90, 180, 270], required=True, help="顺时针旋转度数")
    sp.set_defaults(func=cmd_image_rotate)

    args = ap.parse_args()
    require_docx()  # --help 已在 parse_args 内退出，这里才真正需要 python-docx
    if args.cmd in ("create", "revise", "comment", "image", "copy-from", "compact", "image-rotate"):
        _body_write_gate(getattr(args, "target", None) or getattr(args, "path", None)
                         or getattr(args, "docx", None))
    try:
        args.func(args)
    except SystemExit:
        raise
    except Exception as exc:
        if os.environ.get("TENDER_TRACEBACK"):
            raise  # 调试开关：带完整栈定位；缺省一行式干净（模型读回显用）
        die(f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
