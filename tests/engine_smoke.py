#!/usr/bin/env python3
r"""engine_smoke.py — docx_toolkit 引擎回归电池（合成夹具，运行时生成，不含真实项目数据）。

每条用例对应一个已坐实过的缺陷，走真实命令行（同模型使用方式）。发版前与 fp 两电池
一起跑；新增引擎行为先在这里写成失败用例再改代码。

第 1 批 内容保真
E01 单元格超长按全文落盘并提示        E02 read 视图标出单元格总字数
E03 copy-from 接受修订、摘批注锚      E04 copy-from 重挂超链接
E05 read 能看到修订插入的文字          E06 replace 不重复超链接文字
E07 replace 保标题样式                 E08 replace 命中单 run 只改该 run
E09 跨 run 重写不继承无关格式          E10 表格格填空沿用原格式
E11 空格填空不写显式格式               E12 create --replace 带批注被拦
E13 --drop-comments 放行并回显         E14 create --replace 新建回显「已建节」
E15 合册点名节文件修订痕迹             E16 GBK 输出下建节/合册不崩
E17 check_writing 看得到修订插入文字   E18 批注格填空锚点完好（既有行为回归）

第 2 批 批注与编号
E19 合并格批注坐标三方一致             E20 撞号批注合册体锚对应
E21 gov 深层编号不畸形                 E22 comment-list 按文档位次
E23 同名叶子共用节文件被点名

第 3 批 跨平台与发版闸
E24 toc-pages stderr 无第三方告警      E25 TENDER_TRACEBACK 带完整栈
E26 revise 缺坐标参数给可执行提示

第 4 批 真机复盘（终稿字体/导航大纲/页码收敛/多级目录/版式压缩/图片/完整性）
E27 合册终稿 run 级字体固化            E28 outlineLvl=8 默认摘除（--flat-outline 恢复）
E29 toc-pages 收敛闭环与 --max-rounds  E30 --toc-depth 多级目录 + TOC 独立样式 + 摘 \h
E31 compact 双保护（大字号/含图不压）   E32 image-scan 网格 + image-rotate 重算 extent
E33 偏离表完整性提示（空响应/REQ 对账） E34 表格交叉引用「第X章」核对

第 5 批 Windows 兼容（全量复查）
E35 soffice 认 Windows 默认安装位置    E36 LibreOffice 配置目录 URL 走 as_uri
E37 PDF 句柄必关 + 临时目录清理不抛   E38 Word/WPS 占用锁文件不进枚举
E39 带 BOM 的 txt/md 首行标题不丢

第 6 批 格式件保真（红线）
E40 招标格式件表格只拷不改（重排被拦 / 保真不误报 / 合册拒绝 / --allow-format-change 放行）

第 7 批 文档检查（只读；结果=文字结论+表，检查自身不拦流程）
E41 金额行验算/列验算（单价×数量、合计=明细之和、多级小计）
E42 大写小写一致 + 跨表总价一致（同格/表尾合并行两种真实形态）
E43 口径换算不误报（万元写法、带量词数量、单位后缀）
E44 读不出/空着/零负数/税额 + 「没验算成」兜底
E45 内部对账编号外泄 + 占位写法
E46 表格超宽 / 列太窄 / 固定行高裁字
E47 图片拉伸 / 超宽 / 损坏 / 低分辨率
E48 行距压字 / 域报错 / 未接受的修订
E49 文档属性里的工具痕迹（点名 + 工具建节与合册时不带）
E50 check_writing 接入（非阻断、落 JSON、红线照旧拦）
E51 整本对账（整本比节文件旧 / 被手工改过 / 批注数对不上）
E52 inspect 命令行（只读、结论先行、已检查范围、缺文件、GBK）
E53 真实工作区回放后的校准（之江「叁」／项目总价整行合并／金额整列空着／报价没填／
    人员数量区间／只有表头的窄列／图片阈值／同类并成一行／节文件不看文档属性）
E54 写节闸的指路话（进度面板未首开时给命令，不抛报错栈）
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import types
import zipfile
from pathlib import Path
from unittest import mock

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
TK = [sys.executable, "-B", str(ROOT / "tender-docx" / "scripts" / "docx_toolkit.py")]
CW = [sys.executable, "-B", str(ROOT / "tender-writing" / "scripts" / "check_writing.py")]
RT_HL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
GBK = {"PYTHONIOENCODING": "gbk"}  # 中文 Windows 下宿主以管道捕获输出时的编码环境

PASS, FAIL = [], []
WORK = Path(tempfile.mkdtemp(prefix="engine-smoke-"))
LONG = ("本项目要求投标人具备完善的质量保证体系，并提供近三年同类项目实施案例证明材料，"
        "每个案例须附合同关键页复印件与用户验收证明，所有证明材料须加盖投标人公章。") * 3 + "未尽事宜按招标文件执行。"


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(f"{name}{(' — ' + detail) if detail and not cond else ''}")


def run(name, fn):
    """一条用例出错只记它自己失败，后面的用例照跑。"""
    try:
        fn()
    except Exception as exc:  # noqa: BLE001
        FAIL.append(f"{name} — 用例异常 {type(exc).__name__}: {exc}")


def tk(*a, env=None):
    """跑 docx_toolkit 命令行，返回 (退出码, 合并输出)。"""
    e = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
    if env:
        e.update(env)
    r = subprocess.run(TK + [str(x) for x in a], capture_output=True, text=True, errors="replace", env=e)
    return r.returncode, (r.stdout + r.stderr).strip()


def body(name, blocks) -> Path:
    p = WORK / f"{name}.json"
    p.write_text(json.dumps(blocks, ensure_ascii=False), encoding="utf-8")
    return p


def xml(docx, name) -> str:
    z = zipfile.ZipFile(docx)
    return z.read(name).decode("utf-8") if name in z.namelist() else ""


def rels(docx) -> dict:
    out = {}
    for m in re.finditer(r"<Relationship [^>]*?Id=\"(rId\d+)\"[^>]*/>", xml(docx, "word/_rels/document.xml.rels")):
        tag = m.group(0)
        out[m.group(1)] = (re.search(r'Type="[^"]*/([^"/]+)"', tag).group(1),
                           re.search(r'Target="([^"]+)"', tag).group(1))
    return out


def runs_fmt(paragraph):
    out = []
    for r in paragraph.runs:
        rpr = r._element.rPr
        ea = rpr.rFonts.get(qn("w:eastAsia")) if (rpr is not None and rpr.rFonts is not None) else None
        out.append((r.text, r.font.name, ea, r.font.size.pt if r.font.size else None, r.font.bold))
    return out


def tracked(paragraph, before: str, inserted: str, deleted: str, after: str) -> None:
    """给段落造「修订插入 + 修订删除」痕迹（历史标书里的常见形态）。"""
    paragraph.add_run(before)
    ins = OxmlElement("w:ins")
    for k, v in (("w:id", "901"), ("w:author", "历史同事甲"), ("w:date", "2024-05-06T08:00:00Z")):
        ins.set(qn(k), v)
    ri = OxmlElement("w:r")
    ti = OxmlElement("w:t")
    ti.text = inserted
    ri.append(ti)
    ins.append(ri)
    paragraph._p.append(ins)
    dele = OxmlElement("w:del")
    for k, v in (("w:id", "902"), ("w:author", "历史同事甲"), ("w:date", "2024-05-06T08:00:00Z")):
        dele.set(qn(k), v)
    rd = OxmlElement("w:r")
    td = OxmlElement("w:delText")
    td.text = deleted
    rd.append(td)
    dele.append(rd)
    paragraph._p.append(dele)
    paragraph.add_run(after)


def hyperlink(paragraph, text: str, url: str, doc) -> None:
    r_id = doc.part.relate_to(url, RT_HL, is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    hr = OxmlElement("w:r")
    ht = OxmlElement("w:t")
    ht.text = text
    hr.append(ht)
    hl.append(hr)
    paragraph._p.append(hl)


def make_work(name, outline, leaves=(), with_body=False):
    """搭一个可合册的工作区（评审台/进度面板标记仅为通过写节闸）。返回 (work, 整本路径)。"""
    work = WORK / name
    for d in ("outline", "body", "review"):
        (work / d).mkdir(parents=True)
    (work / "review" / "评审台.html").write_text("<html></html>", encoding="utf-8")
    (work / "review" / ".progress-opened").write_text("", encoding="utf-8")
    (work / "outline" / "tender-response-docs.md").write_text(outline, encoding="utf-8")
    for leaf in leaves:
        args = ["create", work / "body" / f"{leaf}.docx", "--title", leaf]
        if with_body:
            args += ["--body", body(f"b-{leaf}", [{"text": "本项目团队共  人。"}])]
        code, out = tk(*args)
        assert code == 0, out
    return work, next((work / "body").glob("整本-*.docx"), None)


SRC = WORK / "历史标书.docx"      # E03–E06 共用：修订痕迹 + 批注 + 超链接
TGT = WORK / "拷入节.docx"


def build_src():
    d = Document()
    d.add_paragraph("历史标书第一段。")
    p = d.add_paragraph()
    tracked(p, "本项目团队共 ", "十二", "三", " 人。")
    hyperlink(p, "点击查看公告原文", "http://example.com/tender-notice", d)
    d.add_paragraph("历史标书第三段。")
    d.save(str(SRC))
    code, out = tk("comment", SRC, "--para", "P2", "--text", "历史批注：人数要不要更新")
    assert code == 0, out
    parse = WORK / "parse"
    parse.mkdir()
    (parse / "历史标书.elements.json").write_text(json.dumps([[1, 2, 2]]), encoding="utf-8")
    code, out = tk("create", TGT, "--title", "团队介绍")
    assert code == 0, out
    code, out = tk("copy-from", SRC, TGT, "--lines", "L2-L2", "--parse", parse)
    assert code == 0, out


def e01():
    code, out = tk("create", WORK / "长表.docx", "--title", "技术方案", "--body", body("长表", [
        {"type": "table", "header": ["序号", "技术参数要求"], "rows": [["1", LONG]]}]))
    cell = Document(str(WORK / "长表.docx")).tables[0].cell(1, 1).text
    check(f"E01 单元格全文落盘（{len(cell)}/{len(LONG)} 字）并提示",
          code == 0 and len(cell) == len(LONG) and "超长" in out and str(len(LONG)) in out, out[:160])


def e02():
    _, view = tk("read", WORK / "长表.docx")
    check(f"E02 read 标出单元格总字数（共 {len(LONG)} 字）", f"（共 {len(LONG)} 字）" in view, view[:200])


def e03():
    build_src()
    dx = xml(TGT, "word/document.xml")
    clean = (dx.count("<w:ins ") == 0 and dx.count("<w:del ") == 0 and "delText" not in dx
             and "commentRangeStart" not in dx and "commentReference" not in dx)
    check("E03 copy-from 接受修订并摘批注锚", clean, f"ins={dx.count('<w:ins ')} del={dx.count('<w:del ')} "
          f"delText={'delText' in dx} anchor={dx.count('commentRangeStart')}")


def e04():
    dx = xml(TGT, "word/document.xml")
    rm = rels(TGT)
    hl_ids = re.findall(r'<w:hyperlink [^>]*r:id="(rId\d+)"', dx)
    got = rm.get(hl_ids[0]) if len(hl_ids) == 1 else None
    check("E04 copy-from 重挂超链接", got == ("hyperlink", "http://example.com/tender-notice"),
          f"rId={hl_ids} → {got}")


def e05():
    _, view = tk("read", TGT)
    check("E05 read 看得到修订插入的文字", "共 十二 人" in view, view[:200])


def e06():
    code, out = tk("revise", TGT, "--para", "P2", "--find", "公告", "--text", "通知")
    ptxt = Document(str(TGT)).paragraphs[1].text
    check("E06 replace 不重复超链接文字",
          code == 0 and ptxt.count("点击查看") == 1 and "点击查看通知原文" in ptxt, f"exit={code} {ptxt!r}")


FMT = WORK / "格式节.docx"


def build_fmt():
    """标题 + 一个「普通 run / 加粗 run / 普通 run」三段式段落。"""
    tk("create", FMT, "--title", "技术方案", "--body", body("格式节", [
        {"style": "h2", "text": "二 项目背景与现状"}, {"text": "占位。"}]))
    d = Document(str(FMT))
    p3 = d.paragraphs[2]
    for r in list(p3.runs):
        r._element.getparent().remove(r._element)
    p3.add_run("第一段：含")
    p3.add_run("加粗引题").font.bold = True
    p3.add_run("，待核对项目名称。")
    d.save(str(FMT))


def e07():
    build_fmt()
    code, out = tk("revise", FMT, "--para", "P2", "--find", "项目", "--text", "工程")
    h2 = Document(str(FMT)).paragraphs[1]
    check("E07 replace 保标题样式", code == 0 and h2.style.name == "Heading 2" and h2.text == "二 工程背景与现状"
          and h2.runs[0]._element.rPr is None, f"style={h2.style.name} text={h2.text!r}")


def e08():
    code, out = tk("revise", FMT, "--para", "P3", "--find", "项目名称", "--text", "工程名称")
    rs = runs_fmt(Document(str(FMT)).paragraphs[2])
    check("E08 replace 命中单 run 只改该 run",
          code == 0 and len(rs) == 3 and rs[1][0] == "加粗引题" and rs[1][4] is True and rs[2][0] == "，待核对工程名称。",
          str(rs))


def e09():
    code, out = tk("revise", FMT, "--para", "P3", "--find", "含加粗引题，待核对", "--text", "含加粗引题，已核对")
    rs = runs_fmt(Document(str(FMT)).paragraphs[2])
    check("E09 跨 run 重写沿用首个文字 run 的格式",
          code == 0 and len(rs) == 1 and rs[0][4] is None and rs[0][0] == "第一段：含加粗引题，已核对工程名称。", str(rs))


FF = WORK / "格式件.docx"


def build_ff():
    d = Document()
    t = d.add_table(rows=2, cols=2)
    for cell, text in ((t.cell(0, 0), "投标人名称（盖章）："), (t.cell(0, 1), "＿＿＿＿＿＿＿＿"),
                       (t.cell(1, 0), "法定代表人："), (t.cell(1, 1), "")):
        if text:
            r = cell.paragraphs[0].add_run(text)
            r.font.name = "仿宋"
            r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "仿宋")
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.underline = True
    d.save(str(FF))


def e10():
    build_ff()
    code, out = tk("revise", FF, "--cell", "T1R1C2", "--text", "某某科技有限公司")
    r = Document(str(FF)).tables[0].cell(0, 1).paragraphs[0].runs[0]
    got = (r.text, r.font.name, r.font.size.pt, r.font.bold, r.font.underline)
    check("E10 表格格填空沿用原格式", code == 0 and got == ("某某科技有限公司", "仿宋", 14.0, True, True), str(got))


def e11():
    code, out = tk("revise", FF, "--cell", "T1R2C2", "--text", "张三")
    r = Document(str(FF)).tables[0].cell(1, 1).paragraphs[0].runs[0]
    check("E11 空格填空不写显式格式", code == 0 and r.text == "张三" and r._element.rPr is None, str(r._element.rPr))


G = WORK / "带批注节.docx"
B2 = None


def build_g():
    global B2
    B2 = body("两段", [{"text": "第一段正文内容。"}, {"text": "第二段正文内容。"}])
    tk("create", G, "--title", "某节", "--body", B2, "--replace")
    code, out = tk("comment", G, "--para", "P2", "--text", "待澄清：此处金额待核实")
    assert code == 0, out


def e12():
    build_g()
    code, out = tk("create", G, "--title", "某节", "--body", B2, "--replace")
    _, lst = tk("comment-list", G)
    check("E12 create --replace 带批注被拦",
          code == 2 and "--drop-comments" in out and "待澄清" in lst, f"exit={code} {out[:160]}")


def e13():
    code, out = tk("create", G, "--title", "某节", "--body", B2, "--replace", "--drop-comments")
    _, lst = tk("comment-list", G)
    check("E13 --drop-comments 放行并回显", code == 0 and "批注 1 条" in out and "[无批注]" in lst, f"exit={code} {out[:160]}")


def e14():
    code, out = tk("create", WORK / "全新节.docx", "--title", "新节", "--body", B2, "--replace")
    check("E14 create --replace 新建回显「已建节」", code == 0 and "[已建节]" in out and "已覆盖" not in out, out[:120])


def e15():
    work, _ = make_work("w-rev", "# 响应文件：投标文件\n## 目录\n- 目录\n- 团队介绍\n", ["团队介绍"], with_body=True)
    sec = work / "body" / "团队介绍.docx"
    d = Document(str(sec))
    p = d.paragraphs[1]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    tracked(p, "本项目团队共 ", "十二", "三", " 人。")
    d.save(str(sec))
    code, out = tk("assemble", "--work", work, "--numbering", "chapter")
    volume = next((work / "body").glob("整本-*.docx"))
    bx = xml(volume, "word/document.xml")
    body_text = "\n".join(pp.text for pp in Document(str(volume)).paragraphs)  # 接受后无 ins，.text 即全文
    check("E15 合册点名节文件修订痕迹且整本不带痕迹",
          code == 0 and "修订痕迹" in out and bx.count("<w:ins ") == 0 and "共 十二 人" in body_text,
          f"exit={code} ins={bx.count('<w:ins ')} {out[:200]}")


def e16():
    big = body("大表", [{"type": "table", "header": ["序号", "内容"],
                         "rows": [[str(i), f"第 {i} 行内容"] for i in range(1, 62)]}])
    code, out = tk("create", WORK / "大表节.docx", "--title", "大表节", "--body", big, env=GBK)
    check("E16a GBK 下建节不崩且落盘", code == 0 and (WORK / "大表节.docx").is_file(), f"exit={code} {out[:120]}")
    work, _ = make_work("w-gbk", "# 响应文件：投标文件\n## 目录\n- 目录\n- 已写节\n- 未写节\n", ["已写节"])
    code, out = tk("assemble", "--work", work, "--numbering", "chapter", env=GBK)
    check("E16b GBK 下合册缺节不报失败", code == 0 and next((work / "body").glob("整本-*.docx")).is_file(),
          f"exit={code} {out[:120]}")


def e17():
    sys.path.insert(0, str(ROOT / "tender-writing" / "scripts"))
    import check_writing
    check_writing._load_docx()
    full, _ = check_writing.docx_text(TGT)
    check("E17 check_writing 文本口径含修订插入文字", "共 十二 人" in full, full[:120])


def e18():
    q = WORK / "报价.docx"
    tk("create", q, "--title", "报价", "--body", body("报价", [
        {"type": "table", "header": ["项目", "金额"], "rows": [["合计", ""]]}]))
    tk("comment", q, "--cell", "T1R2C2", "--text", "待确认：合计金额")
    tk("revise", q, "--cell", "T1R2C2", "--text", "人民币壹拾贰万元整")
    qx = xml(q, "word/document.xml")
    cellx = re.search(r"<w:tc>.*?</w:tc>", qx[qx.find("合计"):], re.S).group(0)
    order = [m.group(0) for m in re.finditer(
        r"commentRangeStart|<w:t>[^<]*</w:t>|commentRangeEnd|commentReference", cellx)]
    check("E18 批注格填空锚点完好",
          qx.count("commentRangeStart") == 1 and qx.count("commentRangeEnd") == 1 and qx.count("commentReference") == 1
          and order == ["commentRangeStart", "<w:t>人民币壹拾贰万元整</w:t>", "commentRangeEnd", "commentReference"],
          str(order))


# ---------------------------------------------------------------- 第 2 批：批注与编号（B1–B5）

MERGED = body("合并表", [
    {"type": "table", "header": ["A", "B", "C"],
     "rows": [[{"t": "甲", "span": 2}, "乙"],
              [{"t": "上格", "vm": "restart"}, "p", "x"],
              [{"t": "", "vm": "continue"}, "q", "y"]]}])


def e19():
    """B1：合并格批注坐标三处一致（写入 = 清单 = 销号），延续位坐标也能销号。"""
    q = WORK / "合并节.docx"
    tk("create", q, "--title", "合并节", "--body", MERGED)
    c1, o1 = tk("comment", q, "--cell", "T1R2C3", "--text", "批注乙（横向合并行第3网格列）")
    c2, o2 = tk("comment", q, "--cell", "T1R3C1", "--text", "批注上格（竖向合并起始格）")
    assert c1 == 0 and c2 == 0, o1 + o2
    _, lst = tk("comment-list", q)
    ok_list = "[T1R2C3]" in lst and "[T1R3C1]" in lst
    c3, o3 = tk("comment-remove", q, "--cell", "T1R2C3")   # 按写入坐标销号（旧版报 C2、删 C3 失败）
    c4, o4 = tk("comment-remove", q, "--cell", "T1R4C1")   # 延续行坐标也能解析到起始格销号
    _, lst2 = tk("comment-list", q)
    check("E19 合并格批注坐标写入/清单/销号一致",
          ok_list and c3 == 0 and c4 == 0 and "共 2 条" in lst and "[无批注]" in lst2,
          f"list={lst!r} rm3={o3[:80]} rm4={o4[:80]}")


def e20():
    """B2：跨节文件批注号相同时合册不得串批注（原缺陷 2 锚点只剩 1 体）。"""
    work, _ = make_work("w-collide2", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n- 乙节\n",
                        ["甲节", "乙节"], with_body=True)
    fa, fb = work / "body" / "甲节.docx", work / "body" / "乙节.docx"
    tk("comment", fa, "--para", "P2", "--text", "甲节的批注内容")
    tk("comment", fb, "--para", "P2", "--text", "乙节的批注内容")
    ida = re.search(r'w:id="(\d+)"', xml(fa, "word/comments.xml")).group(1)
    idb = re.search(r'w:id="(\d+)"', xml(fb, "word/comments.xml")).group(1)
    tmp = fb.with_suffix(".tmp.docx")   # 把乙节批注号改成与甲节相同（模拟复制节文件当模板）
    with zipfile.ZipFile(fb) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in ("word/comments.xml", "word/document.xml"):
                data = data.decode("utf-8").replace(f'w:id="{idb}"', f'w:id="{ida}"').encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(fb)
    code, out = tk("assemble", "--work", work, "--numbering", "chapter")
    vol = next((work / "body").glob("整本-*.docx"))
    cx, dx = xml(vol, "word/comments.xml"), xml(vol, "word/document.xml")
    bodies = dict(zip(re.findall(r'<w:comment [^>]*w:id="(\d+)"', cx),
                      re.findall(r"<w:t>([^<]*)</w:t>", cx)))
    anchors = set(re.findall(r'<w:commentReference w:id="(\d+)"', dx))
    check("E20 撞号批注合册后体锚一一对应",
          code == 0 and len(bodies) == 2 and anchors == set(bodies)
          and set(bodies.values()) == {"甲节的批注内容", "乙节的批注内容"},
          f"exit={code} bodies={bodies} anchors={anchors}")


def e21():
    """B3：gov 编号深层不畸形（第4层（1）、第5层①；节内小标题不出现 .. / ）. 拼接）。"""
    work, _ = make_work(
        "w-gov2",
        "# 响应文件：投标文件\n## 目录\n- 目录\n- 技术方案\n  - 项目理解\n    - 现状分析\n"
        "      - 细分要点\n        - 明细清单\n    - 需求清单\n  - 实施计划\n- 商务部分\n",
        ["明细清单", "需求清单", "实施计划", "商务部分"], with_body=True)
    # 「需求清单」是第 3 层叶子：节内 h2 小标题应续编为 gov 第 4 层「（1）」形态
    tk("create", work / "body" / "需求清单.docx", "--title", "需求清单", "--replace",
       "--body", body("b-h2", [{"style": "h2", "text": "节内小标题甲"}, {"text": "正文。"}]))
    code, out = tk("assemble", "--work", work, "--numbering", "gov")
    vol = next((work / "body").glob("整本-*.docx"))
    headings = [pp.text for pp in Document(str(vol)).paragraphs
                if pp.style is not None and pp.style.name.startswith("Heading")]
    bad = [h for h in headings if ".." in h or re.search(r"[）)].\d", h)]
    ok_shapes = any("（1）" in h for h in headings) and any("①" in h for h in headings)
    check("E21 gov 深层编号不畸形且形态齐全（含节内小标题续编）",
          code == 0 and not bad and ok_shapes and any(h.startswith("（1）节内小标题甲") for h in headings),
          f"exit={code} headings={headings} bad={bad}")


def e22():
    """B4：comment-list 按真实文档位次排序（表格夹在两段之间时，表格批注排中间）。"""
    q = WORK / "位次节.docx"
    tk("create", q, "--title", "位次节", "--body", body("位次", [
        {"text": "开头段。"},
        {"type": "table", "header": ["A"], "rows": [["1"]]},
        {"text": "结尾段。"}]))
    tk("comment", q, "--para", "P4", "--text", "结尾段批注")
    tk("comment", q, "--para", "P1", "--text", "开头段批注")
    tk("comment", q, "--cell", "T1R2C1", "--text", "表格批注")
    _, lst = tk("comment-list", q)
    seq = [l for l in lst.splitlines() if l.strip().startswith("#")]
    order = ["开头段批注" in s for s in seq].index(True) if any("开头段批注" in s for s in seq) else -1
    mid = next((i for i, s in enumerate(seq) if "表格批注" in s), -1)
    tail = next((i for i, s in enumerate(seq) if "结尾段批注" in s), -1)
    check("E22 comment-list 按文档位次（表格批注排中间）",
          0 <= order < mid < tail, f"seq={seq}")


def e23():
    """B5：不同章节下的同名叶子共用节文件时合册必须点名。"""
    work, _ = make_work(
        "w-dup", "# 响应文件：投标文件\n## 目录\n- 目录\n- 技术方案\n  - 其他说明\n- 商务部分\n  - 其他说明\n",
        ["其他说明"], with_body=True)
    code, out = tk("assemble", "--work", work, "--numbering", "chapter")
    check("E23 同名叶子共用节文件被点名",
          code == 0 and "同名节" in out, f"exit={code} {out[:200]}")


# ---------------------------------------------------------------- 第 3 批：跨平台与发版闸（C2–C4）

def e24():
    """C2：toc-pages 的 stderr 必须干净（pypdfium2 的弃用告警与对象回收噪声曾污染输出）。"""
    soffice = _soffice_path()
    if not soffice:
        raise NotImplementedError("本机无 LibreOffice，跳过")
    work, _ = make_work("w-toc", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", ["甲节"], with_body=True)
    code, out = tk("assemble", "--work", work, "--numbering", "chapter")
    assert code == 0, out
    vol = next((work / "body").glob("整本-*.docx"))
    import subprocess as _sp
    r = _sp.run(TK + ["toc-pages", str(vol)], capture_output=True, text=True, timeout=300)
    check("E24 toc-pages stderr 干净",
          r.returncode == 0 and "UserWarning" not in r.stderr and "Cannot close" not in r.stderr
          and "目录页码" in r.stdout,
          f"exit={r.returncode} stderr={r.stderr[:200]}")


def e25():
    """C3：TENDER_TRACEBACK=1 时异常带完整栈（缺省一行式给模型读）。"""
    bad = WORK / "坏文件.docx"
    bad.write_bytes(b"not a docx")
    rc1, out1 = tk("read", bad)
    import subprocess as _sp
    import os as _os
    r = _sp.run(TK + ["read", str(bad)], capture_output=True, text=True,
                env=dict(_os.environ, TENDER_TRACEBACK="1"))
    check("E25 缺省一行式、开关下带完整栈",
          rc1 != 0 and "Traceback" not in out1 and r.returncode != 0
          and "Traceback" in (r.stderr + r.stdout), f"缺省={out1[:80]} 开关={r.stderr[:120]}")


def e26():
    """C4：revise 缺 --para/--cell 时给可执行提示（旧版报「段落 PNone 不存在」）。"""
    q = WORK / "提示节.docx"
    tk("create", q, "--title", "提示节", "--body", body("提示", [{"text": "正文。"}]))
    rc, out = tk("revise", q, "--find", "正文", "--text", "改")
    check("E26 revise 缺坐标参数给可执行提示",
          rc != 0 and "--para" in out and "PNone" not in out, f"exit={rc} {out[:120]}")


# ---------------------------------------------------------------- 第 4 批：真机复盘


def cw(*a):
    """跑 check_writing 命令行，返回 (退出码, 合并输出)。"""
    r = subprocess.run(CW + [str(x) for x in a], capture_output=True, text=True, errors="replace",
                       env=dict(os.environ, PYTHONDONTWRITEBYTECODE="1"))
    return r.returncode, (r.stdout + r.stderr).strip()


def _assemble(work, *extra):
    code, out = tk("assemble", "--work", work, *extra)
    assert code == 0, out
    return next((work / "body").glob("整本-*.docx"))


def _soffice_path():
    """与发货版同一套探测（含 Windows 默认安装位置）——装了 LibreOffice 的开发机不该误记失败。"""
    return _load("tender-docx/scripts/docx_toolkit.py")._find_soffice()


def e27():
    """P0-1：合册终稿 run 级字体固化——主题属性剥离、eastAsia 显式化、theme a:ea 非空。

    根因三层：themeFontLang=ja-JP（2.1.6 已归一）+ theme 空 a:ea + run 级 w:*Theme
    越权。终稿后处理补后两层：模板会被 skill 更新覆盖，assemble 产物才是交付面。"""
    work, _ = make_work("w-font", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", ["甲节"], with_body=True)
    sec = work / "body" / "甲节.docx"
    d = Document(str(sec))
    p = d.add_paragraph()
    r1 = p.add_run("带主题属性的脏 run")
    rf1 = r1._element.get_or_add_rPr().get_or_add_rFonts()
    rf1.set(qn("w:asciiTheme"), "minorHAnsi")
    rf1.set(qn("w:eastAsiaTheme"), "minorEastAsia")
    r2 = p.add_run("无东亚字体的 run")
    rf2 = r2._element.get_or_add_rPr().get_or_add_rFonts()
    rf2.set(qn("w:ascii"), "Times New Roman")
    d.save(str(sec))
    vol = _assemble(work)
    dx = xml(vol, "word/document.xml")
    theme = xml(vol, "word/theme/theme1.xml")
    all_rfonts = re.findall(r"<w:rFonts [^>]*>", dx)
    check("E27 合册终稿字体固化",
          "w:eastAsiaTheme" not in dx and "w:asciiTheme" not in dx
          and all_rfonts and all('w:eastAsia=' in tag for tag in all_rfonts)
          and re.search(r'<a:ea typeface="[^"]+"', theme) is not None,
          f"themeAttrs在={'eastAsiaTheme' in dx or 'asciiTheme' in dx} "
          f"缺eastAsia的rFonts={sum(1 for t in all_rfonts if 'w:eastAsia=' not in t)} "
          f"themeEa空={'<a:ea typeface=\"\"' in theme}")


def e28():
    """P0-2：outlineLvl=8 默认摘除（导航窗格恢复树形、F9 刷目录不丢子级）；--flat-outline 显式恢复。"""
    work, _ = make_work("w-flat", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n  - 乙节\n",
                        ["乙节"], with_body=True)
    code, out = tk("create", work / "body" / "乙节.docx", "--title", "乙节", "--replace", "--body",
                   body("b-flat", [{"text": "小标题一", "style": "h2"}, {"text": "正文段。"},
                                   {"text": "小标题二", "style": "h3"}]))
    assert code == 0, out
    vol = _assemble(work)
    check("E28a 默认不写 outlineLvl（继承标题样式层级）",
          "w:outlineLvl" not in xml(vol, "word/document.xml"))
    vol2 = _assemble(work, "--flat-outline")
    check("E28b --flat-outline 保留摘出导航", 'w:val="8"' in xml(vol2, "word/document.xml"))


def e29():
    """P0-3：toc-pages 收敛闭环——写后复测、--max-rounds 参数（0 拒绝）、收敛回报。

    页码回填本身会改目录占行→正文再漂：两轮实测一致才收敛；轮内漂移要回显；
    未收敛回滚报警，不静默输出错页码。"""
    if _soffice_path() is None:
        raise NotImplementedError("本机无 LibreOffice，跳过")
    work, _ = make_work("w-conv", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n- 乙节\n",
                        ["甲节", "乙节"], with_body=True)
    vol = _assemble(work)
    r = subprocess.run(TK + ["toc-pages", str(vol), "--max-rounds", "1"],
                       capture_output=True, text=True, timeout=300)
    r2 = subprocess.run(TK + ["toc-pages", str(vol), "--max-rounds", "0"],
                        capture_output=True, text=True, timeout=60)
    check("E29 toc-pages 收敛闭环与 --max-rounds",
          r.returncode == 0 and "收敛" in r.stdout and r2.returncode != 0,
          f"exit={r.returncode} out={r.stdout[:150]} badExit={r2.returncode}")


def e30():
    """P1-4：--toc-depth 控制目录层级；TOC 1–4 独立样式（单倍行距防版式压缩误伤）；域指令摘 \\h。"""
    outline = "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n  - 乙节\n  - 丙组\n    - 丁小节\n"
    work, _ = make_work("w-tocd", outline, ["乙节", "丁小节"], with_body=True)
    vol = _assemble(work)
    dx = xml(vol, "word/document.xml")
    instr = re.search(r"<w:instrText[^>]*>([^<]*)</w:instrText>", dx)
    d = Document(str(vol))
    toc_names = sorted({(p.style.name or "") for p in d.paragraphs
                        if (p.style.name or "").startswith("TOC")})
    sp = d.styles["TOC 1"].paragraph_format.line_spacing if "TOC 1" in toc_names else None
    check("E30a 默认三级目录 + TOC 样式单倍行距 + 摘 \\h",
          instr is not None and '\\o "1-3"' in instr.group(1) and "\\h" not in instr.group(1)
          and toc_names == ["TOC 1", "TOC 2", "TOC 3"] and sp == 1.0,
          f"instr={instr.group(1) if instr else None} styles={toc_names} spacing={sp}")
    vol2 = _assemble(work, "--toc-depth", "1")
    dx2 = xml(vol2, "word/document.xml")
    instr2 = re.search(r"<w:instrText[^>]*>([^<]*)</w:instrText>", dx2)
    d2 = Document(str(vol2))
    toc_names2 = sorted({(p.style.name or "") for p in d2.paragraphs
                         if (p.style.name or "").startswith("TOC")})
    check("E30b --toc-depth 1 只出一级条目",
          instr2 is not None and '\\o "1-1"' in instr2.group(1) and toc_names2 == ["TOC 1"],
          f"instr={instr2.group(1) if instr2 else None} styles={toc_names2}")


def e31():
    """P1-5：compact 双保护常开——36pt 大字号段与含图段不压（exact 会压扁/裁条），普通段压固定行距。"""
    from PIL import Image
    work, _ = make_work("w-compact", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n",
                        ["甲节"], with_body=True)
    sec = work / "body" / "甲节.docx"
    d = Document(str(sec))
    big = d.add_paragraph()
    big.add_run("三十六磅封面标题").font.size = Pt(36)
    tiny = WORK / "tiny.png"
    Image.new("RGB", (40, 20), "red").save(tiny)
    pic = d.add_paragraph()
    pic.add_run().add_picture(str(tiny), width=Cm(5))
    d.add_paragraph("普通正文段落，应当被压缩。")
    d.save(str(sec))
    code, out = tk("compact", "--docx", sec, "--line-pt", "22")
    assert code == 0, out

    def _spacing(pred):
        d2 = Document(str(sec))
        for pp in d2.paragraphs:
            if pred(pp):
                sp = pp._p.pPr.find(qn("w:spacing")) if pp._p.pPr is not None else None
                return (sp.get(qn("w:line")), sp.get(qn("w:lineRule"))) if sp is not None else None
        return None

    def _has_pic(pp):
        return bool(pp._p.findall(".//" + qn("a:blip")))

    check("E31 compact 双保护（大字号/含图不压）",
          _spacing(lambda pp: pp.text.startswith("普通正文")) == ("440", "exact")
          and _spacing(lambda pp: "三十六磅" in pp.text) is None
          and _spacing(lambda pp: _has_pic(pp) and not pp.text.strip()) is None
          and "保护" in out,
          f"norm={_spacing(lambda pp: pp.text.startswith('普通正文'))} out={out[:120]}")


def e32():
    """P1-6：image-scan 出缩略图网格 + image-rotate 旋转并按新比例重算 wp:extent（不重算即拉伸）。"""
    from PIL import Image
    work, _ = make_work("w-img", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n",
                        ["甲节"], with_body=True)
    sec = work / "body" / "甲节.docx"
    cert = WORK / "cert.png"
    Image.new("RGB", (400, 200), "blue").save(cert)
    code, out = tk("image", sec, "--img", cert)
    assert code == 0, out
    before = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', xml(sec, "word/document.xml"))
    code, scan_out = tk("image-scan", "--docx", sec)
    assert code == 0, scan_out
    grids = list((work / "body").glob("图片清单-*.png"))
    code, out = tk("image-rotate", "--docx", sec, "--index", "1", "--deg", "90")
    assert code == 0, out
    after = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', xml(sec, "word/document.xml"))
    with zipfile.ZipFile(sec) as z:
        media = z.read("word/media/image1.png")
    with Image.open(io.BytesIO(media)) as im:
        wh = im.size
    check("E32 image-scan 网格 + image-rotate 重算 extent",
          bool(grids) and "#1" in scan_out and "400x200" in scan_out
          and before is not None and after is not None
          and (after.group(1), after.group(2)) == (before.group(2), before.group(1))
          and wh == (200, 400) and "extent" in out,
          f"grids={bool(grids)} scan={scan_out[:80]} extent={before and before.groups()}→"
          f"{after and after.groups()} media={wh}")


def e33():
    """P2-7：偏离表完整性提示——响应列空行点名 + 与 REQ 登记条数对账（缺商务条款靠人工发现的事故）。"""
    work, _ = make_work("w-dev", "# 响应文件：投标文件\n## 目录\n- 目录\n- 偏离表\n",
                        ["偏离表"], with_body=True)
    ana = work / "analysis"
    ana.mkdir()
    (ana / "requirements-business.md").write_text(
        "| ID | 要求 | coverage |\n|---|---|---|\n"
        "| REQ-001 | 交付期 30 天 | 已承接 |\n| REQ-002 | 质保 24 个月 | 已承接 |\n",
        encoding="utf-8")
    code, out = tk("create", work / "body" / "偏离表.docx", "--title", "偏离表", "--replace", "--body",
                   body("b-dev", [{"type": "table",
                                   "header": ["序号", "招标要求", "响应说明", "偏离"],
                                   "rows": [["1", "交付期", "", "无"]]}]))
    assert code == 0, out
    code, out = cw("--work", work)
    check("E33 偏离表完整性提示",
          "偏离表" in out and "REQ 登记表 2 条" in out and "仅 1 行" in out and "未填写" in out,
          f"exit={code} out={out[:200]}")


def e34():
    """P2-8：toc-pages 收敛后交叉引用核对——表格里「第9章」对不上实测章号被点名（章号写错/增删未同步）。"""
    if _soffice_path() is None:
        raise NotImplementedError("本机无 LibreOffice，跳过")
    work, _ = make_work("w-xref", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n- 乙章\n",
                        ["甲章", "乙章"], with_body=True)
    code, out = tk("create", work / "body" / "甲章.docx", "--title", "甲章", "--replace", "--body",
                   body("b-xref", [{"text": "评分索引见下表。"},
                                   {"type": "table", "header": ["评分项", "对应章节"],
                                    "rows": [["团队配置", "详见第9章"]]}]))
    assert code == 0, out
    vol = _assemble(work, "--numbering", "chapter")
    r = subprocess.run(TK + ["toc-pages", str(vol)], capture_output=True, text=True, timeout=300)
    check("E34 表格交叉引用「第X章」核对",
          r.returncode == 0 and "第9章" in r.stdout and "请核对" in r.stdout,
          f"exit={r.returncode} out={r.stdout[:200]}")


# ---------------------------------------------------------------- 第 5 批 Windows 兼容

def _load(rel: str):
    """按路径载入环节脚本当模块用（走真实函数，不打命令行；同名的只载一次）。"""
    import importlib.util
    name = Path(rel).stem
    if name not in sys.modules:
        spec = importlib.util.spec_from_file_location(name, ROOT / rel)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
    return sys.modules[name]


def e35():
    """Windows 默认安装的 LibreOffice 不进 PATH：两个脚本都得认 Program Files 约定位置，
    且优先 soffice.com（控制台版：转换完才返回、报错有输出可回显）。"""
    program = WORK / "e35" / "Program Files"
    (program / "LibreOffice" / "program").mkdir(parents=True)
    for exe in ("soffice.exe", "soffice.com"):
        (program / "LibreOffice" / "program" / exe).write_bytes(b"")
    want = str(program / "LibreOffice" / "program" / "soffice.com")
    env = {"ProgramFiles": str(program), "ProgramFiles(x86)": "", "ProgramW6432": "", "LOCALAPPDATA": ""}
    with mock.patch.dict(os.environ, env), mock.patch("shutil.which", return_value=None):
        got_tk = _load("tender-docx/scripts/docx_toolkit.py")._find_soffice()
        got_pd = _load("tender-analysis/scripts/parse_document.py")._find_soffice()
    check("E35a docx_toolkit 认 Program Files 下的 LibreOffice（优先 soffice.com）",
          got_tk == want, f"got={got_tk}")
    check("E35b parse_document 认同一位置", got_pd == want, f"got={got_pd}")


def e36():
    """LibreOffice 隔离配置目录的 file URL 须由 Path.as_uri() 生成：手工拼 file://+路径 在
    Windows 上是坏地址（盘符冒号、反斜杠），任何平台上路径含空格/# 都不转义。转换输出须
    宽松解码（text=True 缺省严格，输出含本机代码页外的字节就抛 UnicodeDecodeError）。"""
    seen = []

    def fake_run(cmd, **kw):
        seen.append((list(cmd), kw))
        fmt = cmd[cmd.index("--convert-to") + 1]
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        (outdir / (Path(cmd[-1]).stem + "." + fmt)).write_bytes(b"%PDF-1.4\n%%EOF\n")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    base = WORK / "e36 空格 #目录"
    (base / "out").mkdir(parents=True)
    docx = base / "整本-测试.docx"
    docx.write_bytes(b"PK")
    src_doc = base / "out" / "招标文件.doc"
    src_doc.write_bytes(b"\xd0\xcf\x11\xe0")
    tk = _load("tender-docx/scripts/docx_toolkit.py")
    pd = _load("tender-analysis/scripts/parse_document.py")
    with mock.patch("subprocess.run", side_effect=fake_run), \
            mock.patch.object(pd, "_find_soffice", return_value="/opt/soffice"):
        tk._lo_convert("/opt/soffice", docx, base, base / "profile")
        pd._doc_to_docx(src_doc, base / "out")
    cmd_tk, kw_tk = seen[0]
    url_tk = cmd_tk[2].split("=", 1)[1]
    url_pd = next(a for a in seen[1][0] if a.startswith("-env:UserInstallation=")).split("=", 1)[1]
    check("E36a toc-pages 配置目录 URL 合法（as_uri）",
          url_tk == (base / "profile").resolve().as_uri() and " " not in url_tk and "#" not in url_tk,
          f"url={url_tk}")
    check("E36b .doc 转换配置目录 URL 合法（as_uri）",
          url_pd == (base / "out" / ".lo-profile").resolve().as_uri() and " " not in url_pd,
          f"url={url_pd}")
    check("E36c 转换输出宽松解码", kw_tk.get("errors") == "replace", f"kw={kw_tk}")
    check("E36d .doc 转换产物落 _converted.docx 且中间件清掉",
          (base / "out" / "_converted.docx").is_file() and not (base / "out" / "招标文件.docx").exists())


def e37():
    """Windows 下未关的文件句柄会让临时目录清理抛 PermissionError（WinError 32），盖掉真正的
    报错：页码测量在异常路径也必须关 PDF；临时目录清理失败只留目录不抛。"""
    state = {}

    class Doc:
        def __init__(self, *_a, **_k):
            state["doc"] = self
            self.closed = False

        def __len__(self):
            return 1

        def __getitem__(self, _i):
            tp = types.SimpleNamespace(close=lambda: None, get_text_range=lambda: "正文")
            return types.SimpleNamespace(get_textpage=lambda: tp)

        def close(self):
            self.closed = True

    tk = _load("tender-docx/scripts/docx_toolkit.py")
    fake = types.SimpleNamespace(PdfDocument=Doc)
    with mock.patch.dict(sys.modules, {"pypdfium2": fake}), \
            mock.patch.object(tk, "_page_has_heading", side_effect=RuntimeError("测量出错")):
        raised = ""
        try:
            tk._measure_heading_pages(WORK / "x.pdf", ["甲章"])
        except RuntimeError as exc:
            raised = str(exc)
    closed = bool(state.get("doc") and state["doc"].closed)
    check("E37a 页码测量抛错时 PDF 句柄已关", raised == "测量出错" and closed,
          f"raised={raised!r} closed={closed}")
    td = tk._tempdir()
    try:
        check("E37b 临时目录清理失败不抛（ignore_cleanup_errors）",
              getattr(td, "_ignore_cleanup_errors", False) is True and Path(td.name).is_dir())
    finally:
        td.cleanup()


def e38():
    """Word/WPS 打开节文件时在同目录落占用锁文件（~$节.docx / .~节.docx，隐藏）：按 *.docx
    枚举 body/ 的三处都得跳过——合册孤儿扫描误报「目录外文件」、build_dispatch 同级摘要被
    抢走名额、check_writing 全局校验直接 PackageNotFoundError 崩溃（收尾必跑）。"""
    work, _ = make_work("w-lock",
                        "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n- 乙章\n- 丙章\n",
                        ["甲章", "乙章", "丙章"])
    for name in ("~$甲章.docx", ".~乙章.docx"):
        p = work / "body" / name
        p.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")  # Office 锁文件：非 zip 的小文件
        os.utime(p, (time.time() + 5, time.time() + 5))  # 锁文件最新（打开时才建）
    code, out = tk("assemble", "--work", work)
    check("E38a 合册孤儿扫描不点名占用锁文件",
          code == 0 and "目录外文件" not in out and "合并 4 节" in out, f"exit={code} out={out[:300]}")
    code, out = cw("--work", work)
    check("E38b 全局校验遇锁文件不崩",
          "PackageNotFound" not in out and "Traceback" not in out, f"exit={code} out={out[:300]}")
    got = _load("tender-writing/scripts/build_dispatch.py")._sibling_summaries(work, "", set())
    check("E38c 同级摘要跳过锁文件", len(got) == 3 and not any("~$" in s or ".~" in s for s in got),
          f"got={got}")


def e39():
    """Windows 记事本/PowerShell 5.1 另存的 UTF-8 带 BOM：首行标题不能丢、产物首行不带 BOM 字符。"""
    src = WORK / "e39-bom.md"
    src.write_bytes("\ufeff# 第一章 总则\n\n正文。\n".encode("utf-8"))
    pd = _load("tender-analysis/scripts/parse_document.py")
    L = pd.Lines()
    pd.parse_passthrough(src, L)
    check("E39a 带 BOM 的 md 首行标题被识别", [h[1] for h in L.headings] == ["第一章 总则"],
          f"headings={L.headings}")
    check("E39b 产物首行不带 BOM 字符", L.lines[0] == "# 第一章 总则", f"first={L.lines[0]!r}")


# ---------------------------------------------------------------- 第 6 批：格式件保真（红线）


def _fmt_work(name: str) -> tuple[Path, Path]:
    """搭「招标原件 → 真实解析 → 登记 → 指引 → 节文件」最小工作区，返回 (work, 招标原件)。

    原件三张表：四网格列合并格格式表（L7-L10）、三列评分索引表（L16-L19）、一格方框
    （L23）。四个叶子里评分索引表走 copy-from 拷原件（正路）；其他资信资料用 table 块
    重排成两列键值表，且指引依据列漏记 TPL——只有目录「来源标注」挂着（真机事故形态）；
    封面区间无表；授权委托书把方框换成了段落。"""
    base = WORK / name
    work = base / "work"
    for d in ("outline", "analysis", "body", "review"):
        (work / d).mkdir(parents=True)
    (work / "review" / "评审台.html").write_text("<html></html>", encoding="utf-8")
    (work / "review" / ".progress-opened").write_text("", encoding="utf-8")
    tender = base / "招标文件.docx"
    d = Document()
    d.add_paragraph("第七章 响应文件格式")
    d.add_paragraph("五、其他资信资料")
    d.add_paragraph("其他资信资料")
    t = d.add_table(rows=3, cols=4)
    t.cell(0, 0).text = "单位名称"
    t.cell(0, 2).text = "电话"
    t.cell(0, 2).merge(t.cell(0, 3))
    t.cell(1, 0).text = "单位概况"
    t.cell(1, 0).merge(t.cell(2, 0))
    t.cell(1, 1).text = "统一社会信用代码"
    t.cell(1, 2).merge(t.cell(1, 3))
    t.cell(2, 1).text = "职工总数"
    t.cell(2, 2).text = "共 人"
    t.cell(2, 2).merge(t.cell(2, 3))
    d.add_paragraph("填写说明：资格审查资料中已提供的无需重复提供。")
    d.add_paragraph("评分索引表")
    t2 = d.add_table(rows=3, cols=3)
    for i, s in enumerate(("序号", "评分因素", "页码")):
        t2.cell(0, i).text = s
    t2.cell(1, 0).text = "一、履约能力"
    t2.cell(1, 0).merge(t2.cell(1, 2))
    t2.cell(2, 0).text = "1"
    d.add_paragraph("法定代表人授权委托书")
    t3 = d.add_table(rows=1, cols=1)
    t3.cell(0, 0).text = "授权代表身份证复印件"
    d.save(str(tender))
    r = subprocess.run([sys.executable, "-B", str(ROOT / "tender-analysis" / "scripts" / "parse_document.py"),
                        str(tender), "-o", str(work / "parse")], capture_output=True, text=True, errors="replace")
    assert r.returncode == 0, r.stdout + r.stderr
    (work / "parse" / "sources.json").write_text(
        json.dumps({"main": tender.name, "supplements": [], "excluded": []}), encoding="utf-8")
    (work / "analysis" / "requirements-format.md").write_text(
        "> 来源文件：`招标文件.docx`。下表「同文件」即指该文件。\n\n"
        "## 三、招标方给出的章节示例 / 模板（templates）\n\n"
        "| 序号 | 模板名称 | 归属(bidPart) | 对应必须章节 | 原文出处线索 |\n"
        "|---|---|---|---|---|\n"
        "| TPL-01 | 其他资信资料（单位概况表） | technical | 3 | 同文件 五（L3-L12） |\n"
        "| TPL-02 | 评分索引表 | technical | 1 | 同文件（L14-L19） |\n"
        "| TPL-03 | 封面样例 | technical | 0 | 同文件（L1-L1） |\n"
        "| TPL-04 | 授权委托书 | technical | 2 | 同文件（L21-L23） |\n",
        encoding="utf-8")
    (work / "outline" / "tender-response-docs.md").write_text(
        "# 响应文件：投标文件\n\n## 目录\n- 封面\n- 目录\n- 评分索引表\n- 其他资信资料\n- 授权委托书\n\n"
        "## 来源标注\n- 封面 :: TPL-03\n- 评分索引表 :: TPL-02\n- 其他资信资料 :: MAND-01, TPL-01\n"
        "- 授权委托书 :: TPL-04\n\n"
        "## 目录说明\n- 封面 :: 交付形态=模板或附件填充 | 归位理由=样例 | 理由来源=TPL-03 | 概述=封面\n"
        "- 评分索引表 :: 交付形态=模板或附件填充 | 归位理由=格式件 | 理由来源=TPL-02 | 概述=评分索引\n"
        "- 其他资信资料 :: 交付形态=模板或附件填充 | 归位理由=格式件 | 理由来源=MAND-01 | 概述=单位概况表\n"
        "- 授权委托书 :: 交付形态=模板或附件填充 | 归位理由=格式件 | 理由来源=TPL-04 | 概述=授权书\n",
        encoding="utf-8")
    (work / "body" / "写作指引.md").write_text(
        "| 节 | 模式 | 依据 | 素材 | 图示 | 缺口/备注 |\n|---|---|---|---|---|---|\n"
        "| 封面 | — | TPL-03 | — | — | 格式件：按封面样例排版 |\n"
        "| 目录 | — | — | — | — | 目录页：合册机械生成，不建节 |\n"
        "| 评分索引表 | 格式跟随 | TPL-02 | 招标文件.docx L16-L19 | — | — |\n"
        "| 其他资信资料 | 格式跟随 | MAND-01 | — | — | 招标原表为四列合并单元格，填入后版面失衡，"
        "改为两列项目/内容表逐项填报 |\n"
        "| 授权委托书 | 格式跟随 | TPL-04 | — | — | — |\n",
        encoding="utf-8")
    body_dir = work / "body"
    code, out = tk("create", body_dir / "封面.docx", "--title", "封面", "--body",
                   body("b-fmt-cover", [{"text": "科学数据标注服务项目", "style": "cover-title"}]))
    assert code == 0, out
    sec_b = body_dir / "评分索引表.docx"
    code, out = tk("create", sec_b, "--title", "评分索引表", "--body",
                   body("b-fmt-index", [{"text": "本表按评审细则逐项标注响应位置。"},
                                        {"type": "table", "header": ["备注项", "说明"],
                                         "rows": [["修订记录", "无"]]}]))
    assert code == 0, out
    code, out = tk("copy-from", tender, sec_b, "--lines", "L16-L19",
                   "--parse", work / "parse" / tender.name)
    assert code == 0, out
    # 拷入的表是 T2：往空行填页码、再插一行（格式件只填不改的正路）
    code, out = tk("revise", sec_b, "--cell", "T2R3C3", "--text", "12")
    assert code == 0, out
    code, out = tk("revise", sec_b, "--cell", "T2R3C1", "--mode", "insert_after")
    assert code == 0, out
    code, out = tk("create", body_dir / "其他资信资料.docx", "--title", "其他资信资料", "--body",
                   body("b-fmt-rebuilt", [{"text": "本表按采购文件要求逐项填报。"},
                                          {"type": "table", "caption": "其他资信资料",
                                           "header": ["项目", "内容"],
                                           "rows": [["单位名称", "某某科技有限公司"],
                                                    ["统一社会信用代码", "91440300000000000X"],
                                                    ["职工总数", "共 12 人"]]}]))
    assert code == 0, out
    code, out = tk("create", body_dir / "授权委托书.docx", "--title", "授权委托书", "--body",
                   body("b-fmt-auth", [{"text": "附："}, {"text": "授权代表身份证复印件："}]))
    assert code == 0, out
    return work, tender


def e40():
    """红线：招标格式件表格只拷不改——节内表格与原件格式表对不上，单节自查、全局校验、
    合册三道都拦；用户明确同意的节按参数放行；核对自身出错只告警。"""
    work, tender = _fmt_work("w-fmt")
    sec_a = work / "body" / "其他资信资料.docx"
    sec_b = work / "body" / "评分索引表.docx"
    code, out = cw("--work", work, "--section", sec_a)
    check("E40a 重排的格式表被单节自查拦下（原件 4 列、节内 2 列，点名丢失标签）",
          code == 1 and "招标格式件" in out and "4 列" in out and "2 列" in out and "单位概况" in out,
          f"exit={code} out={out[:300]}")
    code, out = cw("--work", work, "--section", sec_b)
    check("E40b 拷原件填空+插行+另有自建表的节通过", code == 0 and "[校验通过]" in out,
          f"exit={code} out={out[:300]}")
    code, out = cw("--work", work)
    check("E40c 全局校验点名两处被改的格式件（重排、方框被换成段落）并给红线指令",
          code == 1 and "其他资信资料.docx" in out and "授权委托书.docx" in out
          and "没有表格" in out and "[红线指令]" in out and "[G6 指令]" not in out,
          f"exit={code} out={out[:400]}")
    check("E40d 保真的节与区间无表的节不被误报",
          "评分索引表.docx" not in out and "封面.docx" not in out, out[:400])
    code, out = cw("--work", work, "--allow-format-change", "其他资信资料，授权委托书")
    check("E40e 用户同意的节按参数放行后校验通过", code == 0 and "[校验通过]" in out and "[G6 指令]" in out,
          f"exit={code} out={out[:300]}")
    code, out = cw("--work", work, "--guide")
    check("E40f 指引备注里的结构变更措辞被提示",
          "格式件结构变更" in out and "其他资信资料" in out, f"exit={code} out={out[:300]}")
    code, out = tk("assemble", "--work", work)
    check("E40g 合册拒绝带被改格式件的稿子（退出码 2、不产整本）",
          code == 2 and "[合册被拦]" in out and "其他资信资料.docx" in out
          and not list((work / "body").glob("整本-*.docx")), f"exit={code} out={out[:300]}")
    code, out = tk("assemble", "--work", work, "--allow-format-change", "其他资信资料,授权委托书")
    check("E40h 用户同意后合册放行", code == 0 and list((work / "body").glob("整本-*.docx")),
          f"exit={code} out={out[:300]}")
    # 核对自身出错：校验只告警、合册照常——保真核对是防线，不许反过来变成交付故障点
    cw_mod = _load("tender-writing/scripts/check_writing.py")
    cw_mod._load_docx()
    with mock.patch.object(cw_mod, "format_fidelity", side_effect=RuntimeError("核对出错")):
        issues, warnings = cw_mod.check_global(work, [])
    check("E40i 保真核对出错时全局校验只告警",
          not issues and any("格式件保真核对未能执行" in w for w in warnings), f"{issues} {warnings}")
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    tk_mod.require_docx()
    args = types.SimpleNamespace(work=str(work), numbering="chapter", toc_depth=3, flat_outline=False,
                                 template=None, allow_format_change="")
    with mock.patch.object(tk_mod, "format_fidelity", side_effect=RuntimeError("核对出错")):
        tk_mod.cmd_assemble(args)
    check("E40j 保真核对出错时合册照常完成", list((work / "body").glob("整本-*.docx")))
    # 解析产物不在时无从比对：跳过不拦、不崩
    shutil.move(str(work / "parse" / tender.name), str(work / "parse" / "gone"))
    code, out = cw("--work", work, "--section", sec_a)
    check("E40k 解析产物缺失时跳过核对", code == 0, f"exit={code} out={out[:200]}")
    # 真机工作区两种登记形态都得认：首列写纯序号（编号=行序）、出处写「同文件」而
    # 来源确认单与登记表头部声明都没有（退到唯一的那份解析产物）——漏一种整份核对就静默失效
    shutil.move(str(work / "parse" / "gone"), str(work / "parse" / tender.name))
    tk_mod, cw_mod = _load("tender-docx/scripts/docx_toolkit.py"), _load("tender-writing/scripts/check_writing.py")
    mine = {k: v["ranges"][0][0] for k, v in tk_mod.ff_tpl_registry(work).items()}
    theirs, _s = cw_mod._registry_payload(work)
    theirs = {k: v["line"] for k, v in theirs.items() if k.startswith("TPL-")}
    check("E40l 模板登记读取与 check_writing 同口径（编号按行序、起始行一致）",
          mine == theirs == {"TPL-01": 3, "TPL-02": 14, "TPL-03": 1, "TPL-04": 21}, f"{mine} vs {theirs}")
    (work / "parse" / "sources.json").unlink()
    ana = work / "analysis" / "requirements-format.md"
    text = ana.read_text(encoding="utf-8").split("\n", 1)[1]
    for n in range(1, 5):
        text = text.replace(f"| TPL-0{n} |", f"| {n} |")
    ana.write_text(text, encoding="utf-8")
    code, out = cw("--work", work, "--section", sec_a)
    check("E40m 首列纯序号 + 「同文件」+ 无来源确认单时仍核对得到",
          code == 1 and "招标格式件" in out, f"exit={code} out={out[:300]}")


# ---------------------------------------------------------------- 第 7 批 文档检查（只读）

TABLE_HEAD = "| 位置 | 问题 | 依据 | 确定/疑似 |"


def ins(*a, env=None):
    """跑 docx_toolkit inspect 命令行（只读），返回 (退出码, 合并输出)。"""
    return tk("inspect", *a, env=env)


def _doc():
    """不带 python-docx 作者属性的空白文档——文档属性是 E49 专门验的，别的用例不牵扯。"""
    d = Document()
    d.core_properties.author = d.core_properties.last_modified_by = d.core_properties.comments = ""
    return d


def _table(d, header, rows, widths_cm=None):
    """建一张表（表头 + 数据行）；给了 widths_cm 就固定布局并逐列设宽。"""
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    for i, h in enumerate(header):
        t.cell(0, i).text = h
    for ri, row in enumerate(rows, 1):
        for ci, v in enumerate(row):
            t.cell(ri, ci).text = v
    if widths_cm:
        t.autofit = False
        for ci, w in enumerate(widths_cm):
            t.columns[ci].width = Cm(w)
            for cell in t.columns[ci].cells:
                cell.width = Cm(w)
    return t


QTY_HDR = ["名称", "数量", "单位", "单价（元/万条）", "小计（元）", "备注"]
GOOD_ROWS = [("标注服务甲", "300", "万条", "12,650.00", "3,795,000.00", "含税，包干价"),
             ("标注服务乙", "900", "万条", "1,000.00", "900,000.00", "")]
BAD_ROW = [("标注服务甲", "300", "万条", "12,650.00", "3,795,000.00", ""),
           ("标注服务乙", "900", "万条", "1,000.00", "800,000.00", "")]


def _quote(name, rows, total, header=QTY_HDR):
    """之江实验室项目「初次报价组成明细表」形态：单价按元/万条计、合计行跨合并格只有金额。"""
    d = _doc()
    d.add_heading("初次报价组成明细表", 1)
    d.add_paragraph("（价格单位：元人民币，含税）")
    t = _table(d, header, list(rows) + [["合计（以上费用之和）", "", "", "", total, ""]])
    last = t.rows[-1]
    last.cells[0].merge(last.cells[3])
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _sheet(name, upper, lower="¥ 3,795,000.00"):
    """「初次报价一览表」形态：大写小写在同一格。"""
    d = _doc()
    d.add_heading("初次报价一览表", 1)
    d.add_paragraph("（价格单位：元人民币）")
    _table(d, ["序号", "项目名称", "投标总价（元）"],
           [["1", "科学数据标注服务", f"小写：{lower}\n大写：{upper}"]])
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _amount_work(name):
    """带一张算错的报价表的工作区（评审台与进度面板标记齐备，能过写节闸）。"""
    work, _ = make_work(name, "# 响应文件：投标文件\n## 目录\n- 目录\n- 报价表\n- 技术方案\n",
                        ["报价表", "技术方案"])
    sec = work / "body" / "报价表.docx"
    d = Document(str(sec))
    _table(d, ["名称", "数量", "单价（元）", "合价（元）"], [["设备甲", "2", "100.00", "150.00"]])
    d.save(str(sec))
    return work, sec


def e41():
    """金额：单价×数量≠合价、合计行≠明细之和被点名并带算式与格坐标；多级小计逐级验；对得上的零发现。"""
    code, out = ins(_quote("报价-行错", BAD_ROW, "4,595,000.00"))
    check("E41a 单价×数量≠合价被点名（带算式与格坐标，合计没错就不提合计）",
          code == 0 and "单价 × 数量 ≠ 合价" in out and "1,000.00 × 900 = 900,000.00" in out
          and "表里写的是 800,000.00" in out and "T1R3C5" in out and "明细之和" not in out,
          f"exit={code} out={out[:500]}")
    check("E41a' 表格四列齐（位置/问题/依据/确定还是疑似），行验算算确定",
          TABLE_HEAD in out and out.count("| 确定 |") == 1, out[:500])
    code, out = ins(_quote("报价-列错", GOOD_ROWS, "4,700,000.00"))
    check("E41b 合计行≠明细之和被点名（两个数与差额都写出来）",
          "合计行 ≠ 明细之和" in out and "4,700,000.00" in out and "4,695,000.00" in out
          and "差 5,000.00" in out and "单价 × 数量" not in out, out[:500])
    code, out = ins(_quote("报价-对", GOOD_ROWS, "4,695,000.00"))
    check("E41c 算术全对的报价表零发现（阴性对照）", "未发现问题" in out, out[:400])
    lv = [["设备甲", "1", "100.00", "100.00"], ["设备乙", "1", "200.00", "200.00"], ["小计", "", "", "300.00"],
          ["服务丙", "1", "50.00", "50.00"], ["服务丁", "1", "50.00", "50.00"], ["小计", "", "", "100.00"],
          ["合计", "", "", "400.00"]]
    hdr = ["名称", "数量", "单价（元）", "合价（元）"]

    def multi(total):
        d = _doc()
        d.add_heading("分项报价表", 1)
        _table(d, hdr, lv[:-1] + [["合计", "", "", total]])
        p = WORK / f"分项-{total}.docx"
        d.save(str(p))
        return p
    code, out = ins(multi("400.00"))
    check("E41d 多级小计与合计都对得上 → 零发现", "未发现问题" in out, out[:400])
    code, out = ins(multi("450.00"))
    check("E41e 合计错了只点合计、不点小计", out.count("≠ 明细之和") == 1 and "小计行" not in out, out[:500])


def e42():
    """大写小写一致、跨表总价一致——同格（之江实验室项目真实数字，大写带「叁」）与表尾合并行
    （qszbh260186 项目）两种真实形态。"""
    detail = _quote("明细-之江", [("科学数据标注服务（数学）", "300", "万条", "12,650.00", "3,795,000.00", "含税，包干价")],
                    "3,795,000.00")
    code, out = ins(_sheet("一览-之江", "人民币叁佰柒拾玖万伍仟元整"), detail)
    check("E42a 大写小写同格且一致、跨表总价一致 → 零发现", "未发现问题" in out, out[:500])
    code, out = ins(_sheet("一览-大写错", "人民币肆佰万零壹仟元整"), detail)
    check("E42b 大写与小写不一致被点名（两个数都写出来）",
          "大写与小写不一致" in out and "4,001,000" in out and "3,795,000.00" in out, out[:500])
    code, out = ins(_sheet("一览-总价错", "人民币叁佰柒拾玖万伍仟捌佰元整", "¥ 3,795,800.00"), detail)
    check("E42c 两张表的总价不一致被点名（一览表与明细表各是多少）",
          "各处写的总价不一致" in out and "3,795,800.00" in out and "3,795,000.00" in out
          and "大写与小写不一致" not in out, out[:600])
    d = _doc()
    d.add_heading("初次报价一览表", 1)
    t = _table(d, ["序号", "标的名称", "规格型号（或具体服务）", "数量", "单价（元）", "合计（元）"],
               [["1", "公式标注", "最终有效公式", "30万", "2.2", "660000"]])
    cells = t.add_row().cells
    cells[0].merge(cells[5]).text = "响应报价（计价单位：人民币元） 小写：660000 元 大写：陆拾陆万元整"
    d.add_paragraph("说明：响应报价仅作为价格分评审依据。")
    p = WORK / "一览-qszbh.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E42d 数量带「万」、小写带单位、表尾合并行里大小写同行 → 零发现", "未发现问题" in out, out[:500])


def e43():
    """口径：万元写法、数量带单位、金额带括注都不误报。"""
    detail = _quote("明细-口径", [("科学数据标注服务（数学）", "300", "万条", "12,650.00", "3,795,000.00", "")],
                    "3,795,000.00")
    code, out = ins(_sheet("一览-万元", "人民币叁佰柒拾玖万伍仟元整", "379.5 万元"), detail)
    check("E43a 小写写成「379.5 万元」与明细表合计对得上 → 零发现", "未发现问题" in out, out[:500])
    d = _doc()
    d.add_heading("单价表", 1)
    _table(d, ["名称", "数量", "单价", "合计"], [["公式标注", "30万条", "2.2元/条", "660,000.00"],
                                              ["校验标注", "1 批", "1,000 元", "1,000.00（含税）"]])
    p = WORK / "口径-单位.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E43b 数量带单位、单价带单位、金额带括注 → 行验算照常做且零发现", "未发现问题" in out, out[:500])


def e44():
    """兜底：读不出、空着、零与负数、税额都要点名；认出是报价表却一行都没验算成也要点名。"""
    d = _doc()
    d.add_heading("报价表", 1)
    _table(d, QTY_HDR, [["服务甲", "1", "批", "100.00", "面议", ""],
                        ["服务乙", "1", "批", "", "100.00", ""]])
    p = WORK / "兜底-读不出.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E44a 合价格写成「面议」→ 点名读不出", "金额读不出来" in out and "面议" in out, out[:500])
    check("E44b 单价格空着 → 点名空着", "单价空着" in out, out[:500])
    d = _doc()
    d.add_heading("报价表", 1)
    _table(d, QTY_HDR, [["服务甲", "1", "批", "0", "0", ""], ["服务乙", "1", "批", "-100.00", "-100.00", ""]])
    p = WORK / "兜底-零负.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E44c 单价为 0、为负 → 点名", "为 0 或负数" in out and out.count("为 0 或负数") == 2, out[:500])
    d = _doc()
    d.add_heading("报价表", 1)
    _table(d, ["名称", "数量", "单价", "合计"], [["服务甲", "壹批", "拾元", "拾元"]])
    p = WORK / "兜底-没验算.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E44d 表头像报价表、却一行都没验算成 → 点名（不静默跳过）",
          "没有一行能验算" in out, out[:500])
    tax = ["名称", "金额（元）", "税率", "税额（元）", "价税合计（元）"]

    def taxed(rows):
        d = _doc()
        d.add_heading("含税报价表", 1)
        _table(d, tax, rows)
        p = WORK / f"税-{len(rows)}-{abs(hash(tuple(map(tuple, rows)))) % 997}.docx"
        d.save(str(p))
        return p
    code, out = ins(taxed([["甲", "1,000.00", "13%", "130.00", "1,130.00"],
                           ["乙", "333.33", "13%", "43.33", "376.66"]]))
    check("E44e 税额与价税合计都对得上（含四舍五入到分）→ 零发现", "未发现问题" in out, out[:500])
    code, out = ins(taxed([["甲", "1,000.00", "13%", "150.00", "1,150.00"],
                           ["乙", "1,000.00", "13%", "130.00", "1,200.00"]]))
    check("E44f 税额≠金额×税率、价税合计≠金额+税额 各点一处",
          "税额与金额 × 税率对不上" in out and "价税合计 ≠ 金额 + 税额" in out
          and "130.00" in out and "1,130.00" in out, out[:600])


def e45():
    """内部对账编号不进正文；贴入位与其他占位写法点名；正常的条款引用不误报。"""
    d = _doc()
    d.add_heading("技术方案", 1)
    d.add_paragraph("本章响应 REQ-03 与 SCORE-12 的要求。")
    d.add_paragraph("详见第三章 2.3 条。")
    d.add_paragraph("（此处贴入：营业执照复印件，加盖公章）")
    d.add_paragraph("投标人：×××")
    d.add_paragraph("联系电话：此处填写")
    p = WORK / "编号与占位.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E45a 正文里的内部对账编号被点名（两个都点到）",
          "内部对账编号" in out and "REQ-03" in out and "SCORE-12" in out, out[:600])
    check("E45b 「详见第三章 2.3 条」不误报", "第三章" not in out, out[:600])
    check("E45c 贴入位还在、×××、此处填写 各点一处",
          "贴入位还在" in out and "×××" in out and "此处填写" in out, out[:600])


def e46():
    """表格几何：超宽、列太窄、固定行高裁字；正常的表不误报。"""
    def tbl(name, header, rows, widths):
        d = _doc()
        d.add_heading("团队配置", 1)
        _table(d, header, rows, widths)
        p = WORK / f"{name}.docx"
        d.save(str(p))
        return p
    code, out = ins(tbl("表-超宽", ["岗位", "姓名", "职责", "备注"], [["项目经理", "张三", "负责整体交付", ""]], [5, 5, 5, 5]))
    check("E46a 表格超出可用宽度被点名（两个宽度都写出来）",
          "表格超出可用宽度" in out and "20.0 厘米" in out and "15.2 厘米" in out, out[:500])
    code, out = ins(tbl("表-正常", ["岗位", "姓名", "职责", "备注"], [["项目经理", "张三", "负责整体交付", ""]], [3, 3, 6, 3]))
    check("E46b 宽度合适的表零发现", "未发现问题" in out, out[:400])
    code, out = ins(tbl("表-窄列", ["联系电话", "联系人", "备注"], [["13800000000", "张三", ""]], [0.6, 6, 6]))
    check("E46c 列太窄被点名（按列点一次、带列头名）",
          "一字一行" in out and "联系电话" in out and out.count("一字一行") == 1 and "| 疑似 |" in out, out[:500])
    clip = WORK / "表-裁字.docx"
    d = _doc()
    d.add_heading("说明", 1)
    t = _table(d, ["项目", "内容"], [["概述", "这是一段需要两行才能放下的说明文字，用来验证固定行高裁字。"]], [3, 9])
    trPr = t.rows[1]._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), "200")
    h.set(qn("w:hRule"), "exact")
    trPr.append(h)
    d.save(str(clip))
    code, out = ins(clip)
    check("E46d 固定行高小于一行字被点名（行高与字号都写出来）", "固定行高" in out and "10 磅" in out, out[:500])


def e47():
    """图片：拉伸变形、超出可用宽度、打不开、分辨率偏低；等比不超宽的图零发现。"""
    from PIL import Image
    img = WORK / "pic.png"
    Image.new("RGB", (400, 200), "blue").save(img)

    def pic(name, source=img, **kw):
        d = _doc()
        d.add_heading("架构图", 1)
        d.add_paragraph("系统总体架构如下图所示：")
        d.add_picture(str(source), **kw)
        p = WORK / f"{name}.docx"
        d.save(str(p))
        return p
    code, out = ins(pic("图-拉伸", width=Cm(10), height=Cm(10)))
    check("E47a 图片拉伸变形被点名（原像素与页面尺寸）",
          "拉伸变形" in out and "400×200" in out and "10.0×10.0" in out, out[:500])
    code, out = ins(pic("图-超宽", width=Cm(20)))
    check("E47b 图片超出可用宽度被点名", "图片超出可用宽度" in out, out[:500])
    code, out = ins(pic("图-正常", width=Cm(8)))
    check("E47c 等比、不超宽、分辨率够的图零发现", "未发现问题" in out, out[:400])
    broken = pic("图-损坏", width=Cm(8))
    tmp = broken.with_suffix(".tmp")
    with zipfile.ZipFile(broken) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            zout.writestr(item, b"not an image" if item.filename.startswith("word/media/") else data)
    tmp.replace(broken)
    code, out = ins(broken)
    check("E47d 图片字节损坏被点名", "图片打不开" in out, out[:400])
    tiny = WORK / "tiny.png"
    Image.new("RGB", (60, 30), "white").save(tiny)
    code, out = ins(pic("图-模糊", source=tiny, width=Cm(8)))
    check("E47e 有效分辨率不足 100 DPI 被点名（疑似）", "分辨率偏低" in out and "| 疑似 |" in out, out[:500])


def e48():
    """行距比字小、域报错文字、未接受的修订；空段的定行高不误报。"""
    p = WORK / "版式.docx"
    d = _doc()
    d.add_heading("实施方案", 1)
    q = d.add_paragraph()
    q.add_run("这一段的字号比固定行距大，文字会被压掉。").font.size = Pt(14)
    q.paragraph_format.line_spacing = Pt(10)
    d.add_paragraph("详见下表（错误!未找到引用源）。")
    tracked(d.add_paragraph(), "本项目团队共 ", "十二", "三", " 人。")
    d.add_paragraph("").paragraph_format.line_spacing = Pt(2)
    d.save(str(p))
    code, out = ins(p)
    check("E48a 行距比字小被点名（两个磅值都写出来）", "行距比字小" in out and "10 磅" in out and "14 磅" in out, out[:500])
    check("E48b 域报错文字被点名", "域报错" in out and "未找到引用源" in out, out[:500])
    check("E48c 未接受的修订被点名（插入与删除各几处）", "没接受的修订" in out and "插入 1" in out and "删除 1" in out, out[:500])
    check("E48d 空段的 2 磅定行高不误报", out.count("行距比字小") == 1, out[:500])


def e49():
    """文档属性：工具痕迹被点名；工具建节与合册产出的文档不带。"""
    raw = WORK / "属性-原始.docx"
    Document().save(str(raw))
    code, out = ins(raw)
    check("E49a 文档属性里的工具痕迹被点名", "文档属性" in out and "python-docx" in out, out[:500])
    made = WORK / "属性-工具建的.docx"
    code, out = tk("create", made, "--title", "技术方案", "--body", body("b-props", [{"text": "正文。"}]))
    check("E49b 工具建的节不带工具痕迹", code == 0 and "python-docx" not in xml(made, "docProps/core.xml"),
          xml(made, "docProps/core.xml")[:300])
    work, _ = make_work("w-props", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", ["甲节"])
    code, out = tk("assemble", "--work", work)
    book = next((work / "body").glob("整本-*.docx"))
    code, out2 = ins(book)
    check("E49c 生成的整本不带工具痕迹", code == 0 and "python-docx" not in xml(book, "docProps/core.xml")
          and "工具痕迹" not in out2, xml(book, "docProps/core.xml")[-300:] + out2[:200])


def e50():
    """check_writing 接入：文档检查结果跟在自查与全局校验后面、不改变结论；落 JSON；红线照旧拦。"""
    work, sec = _amount_work("w-ins-cw")
    res = work / "review" / "check-result.json"
    code, out = cw("--work", work, "--section", sec, "--json", res)
    data = json.loads(res.read_text(encoding="utf-8"))
    check("E50a 单节自查带上文档检查结果但不阻断（结论仍是校验通过、退出码 0）",
          code == 0 and "[校验通过]" in out and "单价 × 数量 ≠ 合价" in out and "T1R2C4" in out
          and TABLE_HEAD in out, f"exit={code} out={out[:600]}")
    got = data.get("findings") or []
    check("E50b 检查结果落进 JSON（位置/问题/依据/确定还是疑似/坐标齐全）",
          bool(got) and got[0].get("anchor") == "T1R2C4"
          and {"level", "where", "what", "basis", "file"} <= set(got[0]), str(data)[:400])
    code, out = cw("--work", work, "--json", res)
    check("E50c 全局校验同样带上、退出码仍为 0", code == 0 and "单价 × 数量 ≠ 合价" in out and TABLE_HEAD in out,
          f"exit={code} out={out[:600]}")
    d = Document(str(sec))
    d.add_paragraph("【待补：交付时间】")
    d.save(str(sec))
    code, out = cw("--work", work, "--section", sec)
    check("E50d 原有红线（占位文字）照旧拦，检查结果照常列出",
          code == 1 and "占位文字" in out and "单价 × 数量 ≠ 合价" in out, f"exit={code} out={out[:600]}")
    clean = work / "body" / "技术方案.docx"
    code, out = cw("--work", work, "--section", clean)
    check("E50e 干净的节不多出检查表", code == 0 and TABLE_HEAD not in out, f"exit={code} out={out[:400]}")


def e51():
    """整本对账：节文件比整本新、整本被手工改过、批注数与各节之和对不上；引擎自己写整本后指纹同步。"""
    work, sec = _amount_work("w-ins-book")
    code, out = tk("assemble", "--work", work)
    book = next((work / "body").glob("整本-*.docx"))
    code, out = ins("--work", work)
    check("E51a 刚生成的整本：金额问题照常列出、三项对账无发现",
          code == 0 and "单价 × 数量 ≠ 合价" in out and "整本比节文件旧" not in out
          and "手工改过" not in out and "对不上" not in out, f"exit={code} out={out[:600]}")
    st = sec.stat()
    later = st.st_mtime + 120
    os.utime(sec, (later, later))
    code, out = ins("--work", work)
    check("E51b 节文件在整本之后改过 → 点名整本没重新生成", "整本比节文件旧" in out and "「报价表」" in out, out[:600])
    os.utime(sec, (st.st_atime, st.st_mtime))
    tk("assemble", "--work", work)
    Document(str(book)).save(str(book))  # 模拟在 Word 里打开后保存
    code, out = ins("--work", work)
    check("E51c 整本被手工改过 → 点名（重新生成会覆盖）", "手工改过" in out and "会覆盖" in out, out[:600])
    code, out = tk("assemble", "--work", work)
    check("E51d 重新生成时提醒现有整本被手工改过（不拒绝）", code == 0 and "手工改过" in out, f"exit={code} {out[:400]}")
    code, out = ins("--work", work)
    check("E51e 引擎自己重新生成后指纹同步、不再误报", "手工改过" not in out, out[:600])
    tk("comment", sec, "--para", "P1", "--text", "合价待核")
    code, out = ins("--work", work)
    check("E51f 各节批注数与整本里的对不上 → 点名（两边各几条）",
          "批注" in out and "对不上" in out and "各节相加 1 条" in out and "整本里 0 条" in out, out[:700])
    empty, _ = make_work("w-ins-nobook", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", ["甲节"])
    code, out = ins("--work", empty)
    check("E51g 还没有整本时说清楚、不报错", code == 0 and "还没有整本" in out, f"exit={code} {out[:300]}")
    check("E51i 已检查范围里写明整本对账（「未发现」只对已检查范围成立）", "整本对账" in out or "整本对账" in ins("--work", work)[1],
          ins("--work", work)[1][-260:])
    Path(str(book) + ".fingerprint.json").unlink()  # 旧版引擎生成的整本没有指纹记录
    code, out = ins("--work", work)
    check("E51h 没有指纹记录时说明「整本被手工改过」这一项查不了，表里不出这一行（不假装查过）",
          code == 0 and "没有指纹记录" in out and "整本指纹对账未做" in out and "| 整本被手工改过" not in out,
          f"exit={code} {out[-300:]}")


def e52():
    """inspect 命令行：只读（文件字节不变）、结论先行、零发现时交代已检查范围、缺文件一行式、GBK 管道不崩。"""
    import hashlib
    p = _quote("命令行-行错", BAD_ROW, "4,595,000.00")
    before = hashlib.sha256(p.read_bytes()).hexdigest()
    res = WORK / "inspect-result.json"
    code, out = ins(p, "--json", res)
    after = hashlib.sha256(p.read_bytes()).hexdigest()
    check("E52a 只读：文件字节不变、退出码 0", code == 0 and before == after, f"exit={code}")
    check("E52b 首行给结论并说明未改文件", "[文档检查]" in out and "未改任何文件" in out and "查出 1 处" in out, out[:400])
    data = json.loads(res.read_text(encoding="utf-8"))
    check("E52c --json 落盘结构化结果", isinstance(data.get("findings"), list) and len(data["findings"]) == 1
          and data["findings"][0]["level"] == "确定", str(data)[:300])
    code, out = ins(_quote("命令行-对", GOOD_ROWS, "4,695,000.00"))
    check("E52d 零发现时交代已检查范围", "未发现问题" in out and "已检查" in out and "金额" in out, out[:400])
    code, out = ins(WORK / "不存在.docx")
    check("E52e 文件不存在：一行式失败、退出码非 0、无 Traceback", code != 0 and "[失败]" in out and "Traceback" not in out,
          f"exit={code} {out[:200]}")
    code, out = ins(p, env=GBK)
    check("E52f GBK 管道下照常输出（退出码 0，算式里的数字还在）", code == 0 and "900,000.00" in out,
          f"exit={code} {out[:200]}")


def e53():
    """真实工作区回放后的校准——每种形态都来自本机真实项目（之江实验室、11-07、11-21 的工作区）。
    合成用例全绿的那一版，在真实文件上漏报一处（大写「叁」读不出）、误报三类（整行合并的合计行、
    人员数量区间、只有表头的窄列），这一批把它们逐个钉住。"""
    real = _quote("回放-之江明细", [("科学数据标注服务（数学）", "300", "万条", "12,650.00", "3,795,000.00", "含税，包干价")],
                  "3,795,000.00")
    code, out = ins(_sheet("回放-之江一览", "人民币叁佰柒拾玖万伍仟元整"), real)
    check("E53a 之江实验室真实报价（大写带「叁」、大小写同格、合计跨合并格）零发现", "未发现问题" in out, out[:600])
    hdr7 = ["序号", "分项名称", "单位", "数量", "单价", "合计金额", "备注"]
    d = _doc()
    d.add_heading("分项报价明细表", 1)
    t = _table(d, hdr7, [["1", "产品费用", "项", "1", "100.00", "100.00", ""],
                         ["2", "实施费用", "项", "1", "200.00", "200.00", ""]])
    cells = t.add_row().cells
    cells[0].merge(cells[3]).text = "项目总价"
    cells[5].text = "300.00"
    p = WORK / "回放-项目总价.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53b 整行合并的「项目总价」按合计行验，不误报成读不出来", "未发现问题" in out, out[:600])
    d = _doc()
    d.add_heading("分项报价明细表", 1)
    t = _table(d, hdr7, [["1", "产品费用", "项", "", "", "", "可按细分模块分别报价"],
                         ["2", "实施费用", "项", "", "", "", ""], ["3", "低代码平台", "套", "", "", "", ""]])
    cells = t.add_row().cells
    cells[0].merge(cells[6]).text = "项目总价"
    p = WORK / "回放-整列空着.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53c 报价表金额整列空着 → 点名「金额格都空着」，不说「没有一行能验算」也不说读不出来",
          "报价表的金额格都空着" in out and "产品费用" in out and "没有一行能验算" not in out
          and "读不出来" not in out, out[:600])
    d = _doc()
    d.add_heading("初次报价一览表", 1)
    t = _table(d, ["序号", "标的名称", "规格型号（或具体服务）", "数量", "单价（元）", "合计（元）"],
               [["1", "公式标注", "最终有效公式", "30万", "", ""]])
    cells = t.add_row().cells
    cells[0].merge(cells[5]).text = "响应报价（计价单位：人民币元）\n小写：_______\n大写：_______"
    p = WORK / "回放-报价没填.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53d 报价没填 → 点名单价与合价空着、总价还没填，不重复点「没有一行能验算」",
          "单价和合价都空着" in out and "总价还没填" in out and "没有一行能验算" not in out, out[:600])
    d = _doc()
    d.add_heading("人员需求", 1)
    _table(d, ["角色", "拟投入数量（人）"], [["标注员", "150~214"], ["审核员", "113~161"]])
    p = WORK / "回放-人员表.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53e 人员表里的数量区间不进金额验算", "未发现问题" in out, out[:400])
    d = _doc()
    d.add_heading("团队", 1)
    _table(d, ["性别", "姓名", "备注"], [["男", "张三", ""], ["女", "李四", ""]], [0.9, 6, 6])
    p = WORK / "回放-窄列只有表头.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53f 只有表头两个字的窄列不点名", "未发现问题" in out, out[:400])
    from PIL import Image
    big = WORK / "big.png"
    Image.new("RGB", (1000, 500), "green").save(big)
    small = WORK / "small.png"  # e53 自己的小图：本用例单独跑时不依赖 e47
    Image.new("RGB", (400, 200), "blue").save(small)

    def pic(name, src=big, **kw):
        d = _doc()
        d.add_heading("附图", 1)
        for i in range(kw.pop("n", 1)):
            d.add_paragraph(f"第 {i + 1} 部分：")
            d.add_picture(str(src), **kw)
        p = WORK / f"{name}.docx"
        d.save(str(p))
        return p
    code, out = ins(pic("回放-图-阈值", small, width=Cm(12.5)))    # 400px/12.5cm ≈ 81 DPI：真实图片的最低档
    check("E53g 81 DPI（本机真实图片的最低档）不点名", "分辨率偏低" not in out, out[:400])
    code, out = ins(pic("回放-图-轻拉伸", small, width=Cm(10), height=Cm(4.8)))  # 偏差 4%
    check("E53h 拉伸偏差 4% 不点名", "拉伸变形" not in out, out[:400])
    code, out = ins(pic("回放-图-超一点", width=Cm(15.6)))                      # 比可用宽度多 0.4 厘米
    check("E53i 超出可用宽度 0.4 厘米不点名", "图片超出可用宽度" not in out, out[:400])
    code, out = ins(pic("回放-图-超得多", width=Cm(16.5)))                      # 多 1.3 厘米
    check("E53j 超出可用宽度 1.3 厘米点名（写明多出多少）", "图片超出可用宽度" in out and "多出 1.3 厘米" in out, out[:400])
    bulk = pic("回放-图-成批", width=Cm(16.5), n=5)
    res = WORK / "bulk.json"
    code, out = ins(bulk, "--json", res)
    check("E53k 同一文件 5 张图都超宽 → 表里并成一行，写明共几处",
          out.count("图片超出可用宽度") == 1 and "共 5 处" in out, out[:700])
    check("E53k' 并成一行后逐处的位置仍在 --json 结果里（5 条，各带自己的位置）",
          len(json.loads(res.read_text(encoding="utf-8"))["findings"]) == 5,
          res.read_text(encoding="utf-8")[:300])
    raw = WORK / "回放-原始属性节.docx"
    work = WORK / "w-props-sec"
    for d_ in ("review", "body"):
        (work / d_).mkdir(parents=True, exist_ok=True)
    (work / "review" / "评审台.html").write_text("<html></html>", encoding="utf-8")
    (work / "review" / ".progress-opened").write_text("", encoding="utf-8")
    (work / "outline" / "tender-response-docs.md").parent.mkdir(parents=True, exist_ok=True)
    (work / "outline" / "tender-response-docs.md").write_text("# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", encoding="utf-8")
    d = Document()
    d.add_heading("甲节", 1)
    d.add_paragraph("正文。")
    d.save(str(work / "body" / "甲节.docx"))
    code, out = cw("--work", work, "--section", work / "body" / "甲节.docx")
    check("E53l 流水线里查节文件不看节文件的文档属性（交付的是整本）",
          code == 0 and "工具痕迹" not in out, f"exit={code} {out[:400]}")
    code, out = ins(work / "body" / "甲节.docx")
    check("E53m 单独检查这份文件时，文档属性里的工具痕迹照常点名", "工具痕迹" in out and "python-docx" in out, out[:400])
    # 引擎自己插的图不能比正文区宽。回放时发现：image 命令想按页面宽度算，但「页宽 − 左右边距」
    # 算出来是普通整数、没有 .cm，那一步总是抛异常被吞掉，宽度一直是兜底的 15 厘米——
    # A4 页三边距 3.18 厘米时正文区只有 14.64 厘米，每张图都比正文和表格宽出 0.36 厘米
    w2 = WORK / "w-img-width"
    for d_ in ("review", "body", "outline"):
        (w2 / d_).mkdir(parents=True, exist_ok=True)
    (w2 / "review" / "评审台.html").write_text("<html></html>", encoding="utf-8")
    (w2 / "review" / ".progress-opened").write_text("", encoding="utf-8")
    sec = w2 / "body" / "附图节.docx"
    code, out = tk("create", sec, "--title", "附图节", "--body", body("b-imgw", [{"text": "系统架构如下图："}]))
    code, out = tk("image", sec, "--img", small)
    mod = _load("tender-docx/scripts/docx_toolkit.py")
    mod.require_docx()
    d = mod.Document(str(sec))
    blk = list(d.element.body.iterchildren())
    usable = mod._block_widths_emu(d, blk)[id(blk[0])] / 360000
    blip = next(d.element.body.iter(mod.qn("a:blip")))
    holder = blip
    while holder.tag not in (mod.qn("wp:inline"), mod.qn("wp:anchor")):
        holder = holder.getparent()
    cx = int(holder.find(mod.qn("wp:extent")).get("cx")) / 360000
    check("E53n 引擎自己插的图宽度跟页面走，不超过正文区（页宽 − 左右边距）",
          code == 0 and cx <= usable + 0.02 and abs(cx - usable) < 0.05, f"图宽 {cx:.2f} 厘米，正文区 {usable:.2f} 厘米")
    # 表格比正文区宽 0.4 厘米（拷入的表常见，落进页边距里）不点名；宽 1 厘米以上才点名并写明多出多少
    d = _doc()
    d.add_heading("一览", 1)
    _table(d, ["序号", "名称", "备注", "说明"], [["1", "甲", "", ""]], [4, 4, 4, 3.6])
    p = WORK / "回放-表-超一点.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53o 表格超出可用宽度 0.4 厘米不点名", "表格超出可用宽度" not in out, out[:500])
    d = _doc()
    d.add_heading("一览", 1)
    _table(d, ["序号", "名称", "备注", "说明"], [["1", "甲", "", ""]], [4, 4, 4, 5.2])
    p = WORK / "回放-表-超得多.docx"
    d.save(str(p))
    code, out = ins(p)
    check("E53p 表格超出可用宽度 2 厘米点名（写明多出多少）", "表格超出可用宽度" in out and "多出 2.0 厘米" in out, out[:500])


def e54():
    """写节闸的指路话不能丢：评审台已建、写作进度面板未首开时，单节自查要拒查、退出码 2，
    并说清该跑哪个命令。统一三形态解析时 `_SENTINEL` 只留了两个键，而提示语里还要找
    tender-review——模型忘了先开面板时拿到的是一个 KeyError 报错栈（闸拦住了，路没指）。"""
    work = WORK / "w-gate"
    for d in ("review", "body"):
        (work / d).mkdir(parents=True)
    (work / "review" / "评审台.html").write_text("<html></html>", encoding="utf-8")
    sec = work / "body" / "甲节.docx"
    d = Document()
    d.add_heading("甲节", 1)
    d.add_paragraph("正文。")
    d.save(str(sec))
    code, out = cw("--work", work, "--section", sec)
    check("E54a 进度面板未首开：拒查、退出码 2、给出该跑的命令、无报错栈",
          code == 2 and "写作进度面板未首开" in out and "--open-progress" in out
          and "build_review.py" in out and "Traceback" not in out, f"exit={code} {out[:300]}")


if __name__ == "__main__":
    for name, fn in (("E01", e01), ("E02", e02), ("E03", e03), ("E04", e04), ("E05", e05), ("E06", e06),
                     ("E07", e07), ("E08", e08), ("E09", e09), ("E10", e10), ("E11", e11), ("E12", e12),
                     ("E13", e13), ("E14", e14), ("E15", e15), ("E16", e16), ("E17", e17), ("E18", e18),
                     ("E19", e19), ("E20", e20), ("E21", e21), ("E22", e22), ("E23", e23),
                     ("E24", e24), ("E25", e25), ("E26", e26),
                     ("E27", e27), ("E28", e28), ("E29", e29), ("E30", e30), ("E31", e31),
                     ("E32", e32), ("E33", e33), ("E34", e34),
                     ("E35", e35), ("E36", e36), ("E37", e37), ("E38", e38), ("E39", e39),
                     ("E40", e40), ("E41", e41), ("E42", e42), ("E43", e43), ("E44", e44),
                     ("E45", e45), ("E46", e46), ("E47", e47), ("E48", e48), ("E49", e49),
                     ("E50", e50), ("E51", e51), ("E52", e52), ("E53", e53), ("E54", e54)):
        run(name, fn)
    print(f"\n{len(PASS)} 通过 / {len(FAIL)} 失败")
    for f_ in FAIL:
        print("  FAIL:", f_)
    if FAIL:
        print(f"  现场保留：{WORK}")
    else:
        shutil.rmtree(WORK, ignore_errors=True)
    sys.exit(1 if FAIL else 0)
