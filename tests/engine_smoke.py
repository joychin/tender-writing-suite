#!/usr/bin/env python3
"""engine_smoke.py — docx_toolkit 引擎回归电池（合成夹具，运行时生成，不含真实项目数据）。

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
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

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
    import shutil as _sh
    soffice = _sh.which("soffice") or _sh.which("libreoffice") or \
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not Path(soffice).exists():
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
    import shutil as _sh
    for cand in (_sh.which("soffice"), _sh.which("libreoffice"),
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if cand and Path(cand).exists():
            return cand
    return None


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


if __name__ == "__main__":
    for name, fn in (("E01", e01), ("E02", e02), ("E03", e03), ("E04", e04), ("E05", e05), ("E06", e06),
                     ("E07", e07), ("E08", e08), ("E09", e09), ("E10", e10), ("E11", e11), ("E12", e12),
                     ("E13", e13), ("E14", e14), ("E15", e15), ("E16", e16), ("E17", e17), ("E18", e18),
                     ("E19", e19), ("E20", e20), ("E21", e21), ("E22", e22), ("E23", e23),
                     ("E24", e24), ("E25", e25), ("E26", e26),
                     ("E27", e27), ("E28", e28), ("E29", e29), ("E30", e30), ("E31", e31),
                     ("E32", e32), ("E33", e33), ("E34", e34)):
        run(name, fn)
    print(f"\n{len(PASS)} 通过 / {len(FAIL)} 失败")
    for f_ in FAIL:
        print("  FAIL:", f_)
    if FAIL:
        print(f"  现场保留：{WORK}")
    else:
        shutil.rmtree(WORK, ignore_errors=True)
    sys.exit(1 if FAIL else 0)
