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

第 8 批 公司版式读取（layout-read，只读；「按这个样式来做」场景的读取端）
E55 页面组（多节）：纸张边距取正文所在节、各节差异、网格类型、页眉页码跨节与位置
E56 正文组：字体字号行距缩进、对齐缺省=左、段前段后、中文版式真实读取、列表段
E57 直接格式形态（第X章自带编号）：编号形态/间隔/章起新页；混排疑似并列不截断
E58 目录组：标题/条目分级/层级/点线/缩进；目录条目不进标题候选
E59 标题组（命名样式六级、文字不带编号）：层级、五级及以上、编号形态来源、事实行
E60 表格与图：样式边框与表头底纹、表宽、单元格边距对齐、表题图题位置、图宽占比
E61 封面：止于分节符、文本框去重分行、封面表格、字段标记、页面设置差异
E62 落盘：--out 完整不截断、--json 多文件为列表、原件只读
E63 .doc 输入（LibreOffice；缺了记失败）；PDF 拒收
E64 段前段后：单边不崩、按行写、自动间距；绝对磅值缩进归一成字
E65 主题字体属性优先；奇偶页不同；标题对齐缺省
E66 纯文字目录条目（到分页符为止）；手打编号的条款行不算标题

第 9 批 分析阶段交付链接（build_review 回显；分析阶段交付面=报告页本身）
E67 报告构建（--pages report）的 [交付链接]＝报告页直链 report.html?v=
E68 报告构建无服务：[交付链接]＝报告页静态文件路径（不是「暂不可用」死胡同）
E69 全量构建：评审台 [交付链接] 照旧（带 #tab 锚点）+ 追加 [报告页链接] 一行

第 10 批 结构要求散文可读性（build_review 软提醒；报告页按源行分段，一行源文本=一个 <p>）
E70 超长段回显 [结构要求超长段]（段序+字数+按条款拆行指路；不拦流程）
E71 全部短段不提醒（阴性对照）

第 11 批 落盘元数据清零（用户拍板 2026-11-26「保存的时候不存入任何 metadata」，沉淀为引擎能力）
E72 统一落盘口（compact）：core 全字段/app.xml 应用名与模板名/custom.xml（WPS 的 ICV 身份码）/
    首页缩略图全清，正文照常压缩落盘
E73 合册交付稿（composer 路径）+ 混入的外部脏节文件：整本不带任何元数据
E74 inspect 点名残留（工具痕迹含 ICV/KSO、单位信息、自定义属性、缩略图）；清零后转阴性

第 12 批 落盘路径防呆（--json 是「结果写到哪个文件」不是开关；真机上模型当开关用，把一份真实标书写掉）
E75 四扇落盘门（inspect --json / layout-read --json、--out / check_writing --json）：与被读文件同路径、
    后缀不符一律拒绝且不写任何文件；inspect 没给文件报参数错；合法落盘照常

第 13 批 .doc 通道：没有 LibreOffice 时把转换交给模型（用户 2026-11-27 拍板；脱敏那处见 sanitize_smoke S17）
E76 parse_document 回退指引与 --converted 接力（来源登记／解析目录／转换件落点／产物与自动转换一致、
    幂等沿用、误用拒绝）；layout-read 回退指引；三份脚本的探测与指引逐字同口径

第 14 批 版式读取遇到不规整的真实文档（2026-11-27 真机：读两份真实文件整份崩，模型自己转出来的
    docx 又崩在页面上；每处都先在真实文件上复现，再写成合成夹具）
E77 没有一级标题（章挂在标题 2 上）不崩、按最浅一级照读；节里没写页面设置／没有节属性按 Word 缺省值
    读出并注明；带单位的纸张写法回归；某一组读不出来不拖垮整份（记一行、首行点明、上游没读出的组
    不白跑、TENDER_TRACEBACK=1 照旧抛出）；正常文档阴性对照

第 15 批 用户模板缺标题样式（2026-11-27 真机：模型照 SKILL 写的「缺则回落内置样式并点名」向用户转述；
    实测只补了个空样式、不回显，章标题跟正文一个样）
E78 模板没有 Heading 1-4 时按内置样式补建并回显 [模板标题样式] 点名；建节同路；模板自带标题样式时不回显不覆盖

第 16 批 超过 5MB 的 .doc 不自动转换（用户 2026-11-27 定；脱敏那处见 sanitize_smoke S18）
E79 版式读取与解析遇超过 5MB 的 .doc：不启动转换、退出码 2、请用户自己另存；正好 5MB 不算超；
    解析带 --converted 或已有完好的转换件时照常走；三份脚本的提示与大小线逐字同口径

第 17 批 LibreOffice 转换超时：杀干净，再提醒用户（用户 2026-11-27 定；脱敏那处见 sanitize_smoke S19）
E80 三份脚本共用的「起 LibreOffice 并等它结束」：超时后整棵进程树杀掉（父子孙都不留），Windows 分支
    按 taskkill /T 强杀（mac 上模拟）；三份脚本的超时提示逐字同口径、都是请用户自己另存；版式读取与
    解析端到端走一遍；转换失败走同一个出口；目录页码实测那条路文案照旧

第 18 批 版式读取复盘（2026-11-27 WorkBuddy 真机：读数是对的，模型拿着脚本说错、说漏的几句话
    把用户的版式说成「做得差」，自创一套标题推荐成默认，还越过了「本版只到读出规格表」的边界）
E81 事实行：命名标题样式 0 段时外观从哪来说哪，不写死「标题全是直接格式」
E82 页面设置差异：设置相同的节归成一组；一行装不下时按段收、写明另有几段，不拦腰截断
E83 页码依据：正文节沿用前面某节的页脚时写明沿用关系，不把那一节说成正文所在节
E84 目录缩进：缩进打在条目文字里的（全角空格起头）读出来
E85 回显末尾带下一步与取值纪律（光靠 SKILL.md 的散文拦不住模型手工搭模板）

第 19 批 公司版式：登记与自动采用（2026-11-27 晚，用户：「正常情况下，用户只会通过对话和 agent
    表达使用模板的意图」——真机上模型让用户「建节和合册时传入这个参数」，因为套件到读出规格表就停了）
E86 layout-adopt 登记／layout-show 查看／不是规格表的文件拒收
E87 建节自动按登记的版式出（不传 --template）：样式、缩进、表格、表题全跟版式走
E88 合册自动按登记的版式出：页面与网格、页码写法与位置、编号与间隔、章起新页、目录标题与缩进点线
E89 显式模板与 --ignore-layout 优先；没有登记时照旧
E90 往返一致：按登记的版式出样张，读回的规格表与登记的对得上（两种真实形态的夹具）
E91 对话修订（--set、直接改登记的规格表、改坏了）与 layout-forget；起过叫法的版式修订后叫法不变
E92 读数补洞：段前段后／行距带上文档默认值；目录段前段后单列一行；decimal 编号也读编号间隔；
    纯数字页码的字号看页码域自己写的

第 20 批 公司版式：封面整页与页眉落地（2026-11-27 深夜批——2.4.0 登记时明说「还没接上」的两项，
    要从历史标书里取件：登记时 --source 连源文件一起给，提取封面元素与页眉 logo 存进登记目录）
E93 layout-adopt --source：提取封面（封面.docx）与页眉（logo 图进 媒体/、文字规格进 layout.json）；
    不带 --source 时维持「还没接上」并给一句怎么接
E94 合册落地：封面节点无节文件时落公司封面为 body/封面.docx 并点名待改字段；正文页页眉带
    logo／公司名／横线，前置页不带；页码口径不变
E95 用户（或模型）已写 body/封面.docx 时不覆盖、不点名
E96 单元格边距跟版式（读到非缺省值才显式写）

第 21 批 公司版式真机复盘（2026-11-28 WorkBuddy 首次真实登记抓出的三个「首次/边界」缺陷——
    电池里登记目录总是先建好，盖不住真机上从零开始的形态；模型就地打了补丁，官方修复落仓库）
E97 首次登记：登记目录还不存在时 --source 提取照常成功（改前 FileNotFoundError、整条登记失败）
E98 节文件小标题跳级（直接 h3、没经过 h2）：合册不崩、编号按缺位续（改前 KeyError）
E99 封面提取保留空段——空段就是封面的竖向排版（改前内容全挤到页顶）

第 22 批 素材新鲜度闸（2026-11-28 设计讨论拍板「停」：指引素材列记「文件名+行号区间」只是计划，
    写作逐节进行、素材文件是活的——解析之后被改过，还没拷的节照旧区间拷错，且无声进标书；
    拷过的节独立不受影响，同一素材不同区间可分派多节）
E100 copy-from 源文件在解析后被改：拒绝（退出码 2）、点破「行号区间已不可信」、
    给重新解析+重核指引区间的处置；未改源时照常拷（阴性对照在同一用例里）
E101 重新解析（解析登记记下新指纹）后照常拷贝
E102 解析产物没记源指纹（旧版产物）：不拦、提醒建议重新解析、照常拷贝
E103 旧名名单.txt 隐形约定退场＋空名单明说「残留扫描未执行」（不许没查装查过）
E104 登记资料库旧名自动载入：他主体痕迹=疑似残留不拦、自己主体不扫、显式 --old-names 仍是红线、
    主体悬空（承诺清单没有投标主体行）不猜

第 23 批 之江真机两处崩溃入库（2026-11-28 真机就地补丁被 2.4.5 同步覆盖，官方修复落仓库）
E105 多册合册 + 登记版式封面：封面节点无节文件时按册落公司封面（改前 str/str TypeError 必崩）
E106 删整本重合册：指纹记录在、整本不在——正当处置不崩（改前 FileNotFoundError），指纹刷新
"""
import io
import itertools
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


NO_LAYOUT = WORK / "无登记"   # 版式登记目录的隔离值：用例缺省指向空目录，不读本机真实的公司版式登记


def tk(*a, env=None):
    """跑 docx_toolkit 命令行，返回 (退出码, 合并输出)。"""
    e = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", TENDER_LAYOUT_HOME=str(NO_LAYOUT))
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


def cw(*a, cwd=None, env=None):
    """跑 check_writing 命令行，返回 (退出码, 合并输出)。cwd 供 CWD 相对读取的行为用例
    指定工作目录；env 供登记资料库（TENDER_MATERIAL_HOME）相关用例指向临时登记目录。"""
    e = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
    if env:
        e.update(env)
    r = subprocess.run(CW + [str(x) for x in a], capture_output=True, text=True, errors="replace",
                       env=e, cwd=str(cwd) if cwd is not None else None)
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

    def fake_run(cmd, seconds):     # 顶替「起 LibreOffice 并等它结束」那一步：记下命令行，替它把产物写出来
        seen.append((list(cmd), seconds))
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
    sm = _load("tender-sanitize/scripts/strip_metadata.py")
    with mock.patch.object(tk, "_run_lo", fake_run), mock.patch.object(pd, "_run_lo", fake_run), \
            mock.patch.object(pd, "_find_soffice", return_value="/opt/soffice"):
        tk._lo_convert("/opt/soffice", docx, base, base / "profile")
        pd._doc_to_docx(src_doc, base / "out")
    cmd_tk, sec_tk = seen[0]
    url_tk = cmd_tk[2].split("=", 1)[1]
    url_pd = next(a for a in seen[1][0] if a.startswith("-env:UserInstallation=")).split("=", 1)[1]
    check("E36a toc-pages 配置目录 URL 合法（as_uri）",
          url_tk == (base / "profile").resolve().as_uri() and " " not in url_tk and "#" not in url_tk,
          f"url={url_tk}")
    check("E36b .doc 转换配置目录 URL 合法（as_uri）",
          url_pd == (base / "out" / ".lo-profile").resolve().as_uri() and " " not in url_pd,
          f"url={url_pd}")
    # 转换输出宽松解码：让进程吐出几个解不开的字节，三份脚本各自的 _run_lo 都不该抛，缺省用替换字符
    noisy = [sys.executable, "-c", "import sys; sys.stdout.buffer.write(b'\\xff\\xfex\\n'); "
                                   "sys.stderr.buffer.write(b'\\xffy\\n')"]
    loose, detail = True, ""
    for tag, mod in (("docx_toolkit", tk), ("parse_document", pd), ("strip_metadata", sm)):
        try:
            r = mod._run_lo(noisy, 60)
            ok = r.returncode == 0 and r.stdout == "\ufffd\ufffdx\n" and r.stderr == "\ufffdy\n"
            loose = loose and ok
            detail += f"{tag}:{ok} "
        except Exception as exc:  # noqa: BLE001
            loose = False
            detail += f"{tag}:{type(exc).__name__}:{exc} "
    check("E36c 转换输出宽松解码（输出里有本机代码页外的字节也不抛，缺省用替换字符）", loose, detail.strip())
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
    """不带任何元数据的空白文档——文档属性是 E49/第 11 批专门验的，别的用例不牵扯。
    python-docx 默认模板自带作者与首页缩略图，一并摘掉，阴性对照不被元数据发现打脏。"""
    d = Document()
    d.core_properties.author = d.core_properties.last_modified_by = d.core_properties.comments = ""
    rels = d.part.package.rels
    for rid in [rid for rid, rel in list(rels.items())
                if rel.reltype.rsplit("/", 1)[-1] in ("thumbnail", "thumbnail-picture",
                                                      "custom-properties")]:
        rels.pop(rid, None)
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
    # 模拟在 Word 里打开后保存：真实改一字（python-docx 落盘字节级确定，引擎产物空保存
    # 不变字节——落盘元数据清零后 core.xml 已归一，改前靠 core 往返不稳定制造指纹差）
    bd = Document(str(book))
    bd.paragraphs[0].add_run(" ")
    bd.save(str(book))
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
    check("E52b 首行给结论并说明未改被检查的文件（带了 --json 时它写过结果文件，说「未改任何文件」是假话）",
          "[文档检查]" in out and "未改被检查的文件" in out and "未改任何文件" not in out and "查出 1 处" in out,
          out[:400])
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


# ---- 第 8 批 公司版式读取（layout-read，只读）----
#
# 两份夹具对应两类真实文档（都在回放里坐实过）：
#   _company_doc  结构化形态（深圳奥哲技术标）：三节页面设置各不相同、封面标题在文本框里
#                 （mc:Choice 与 mc:Fallback 各一份）、封面信息表格、目录条目带 toc 样式和
#                 页码域（没有点线字符）、命名标题用到六级且文字不带编号、正文不设对齐
#                 （Word 缺省=左）、段前用「行」写、表格边框与表头底纹来自表格样式
#   _direct_doc   直接格式形态（真实招标文件）：零命名标题、外观全写在文字上、章标题自带
#                 「第X章」且段前分页、点线目录、直接设边框底纹的表格


def _shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _tbl_borders_single(tbl):
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tbl._tbl.tblPr.append(borders)


def _p_bottom_border(p):
    ppr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    pBdr.append(bottom)
    ppr.append(pBdr)


def _run_fmt(r, ea, pt, bold=False):
    r.font.name = "Times New Roman"
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), ea)
    r.font.size = Pt(pt)
    if bold:
        r.font.bold = True


def _field(p, instr, result, prefix="第 ", suffix=" 页"):
    """字面文字 + 域（begin/instr/separate/结果/end 五个 run）。"""
    p.add_run(prefix)
    r = p.add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "begin")
    r._element.append(fc)
    r2 = p.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r2._element.append(it)
    r3 = p.add_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3._element.append(sep)
    p.add_run(result)
    r5 = p.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r5._element.append(end)
    p.add_run(suffix)


_TXBX_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office"')


def _textbox_run(lines):
    """带文本框的 run：mc:Choice 里是 drawing、mc:Fallback 里是 VML，两边文字一样——
    封面标题的常见形态，不去重会把同一行读成两遍。命名空间手工声明（python-docx 的
    命名空间表里没有 mc / wps）。"""
    from docx.oxml import parse_xml

    inner = f'<w:p><w:r><w:t xml:space="preserve">{lines[0]}</w:t></w:r>' + "".join(
        f'<w:r><w:br/><w:t xml:space="preserve">{ln}</w:t></w:r>' for ln in lines[1:]) + "</w:p>"
    txbx = f"<w:txbxContent>{inner}</w:txbxContent>"
    return parse_xml(
        f'<w:r {_TXBX_NS}>'
        '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>1500000</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="5000000" cy="700000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="901" name="文本框"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="5000000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/></wps:spPr>'
        f'<wps:txbx>{txbx}</wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" '
        'anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
        f'<mc:Fallback><w:pict><v:shape id="文本框" o:spid="_x0000_s1026" type="#_x0000_t202" '
        f'style="width:300pt;height:40pt" filled="f" stroked="f">'
        f'<v:textbox inset="0,0,0,0"><w:txbxContent>{inner}</w:txbxContent></v:textbox>'
        '</v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>')


_LSENT = ("本项目要求覆盖数据采集、治理与服务各个环节，并与现有系统通过标准接口衔接，"
          "各层之间职责边界清晰、运行状态可监控可追溯。")


def _direct_doc(name):
    """直接格式形态（真实招标文件）。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

    d = _doc()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.5)

    def para(text, ea, pt, bold=False, indent_chars=None, line=None, align=None, pbb=False):
        p = d.add_paragraph()
        if pbb:
            p.paragraph_format.page_break_before = True
        if align is not None:
            p.alignment = align
        if line is not None:
            p.paragraph_format.line_spacing = line
        if indent_chars is not None:
            p._p.get_or_add_pPr().get_or_add_ind().set(qn("w:firstLineChars"), str(indent_chars))
        _run_fmt(p.add_run(text), ea, pt, bold)
        return p

    para("XX科技股份有限公司投标文件", "黑体", 26, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    para("项目名称：智慧园区数据平台建设项目", "仿宋", 16, True)
    para("项目编号：XNBH-2025-031", "仿宋", 16)
    para("投标人：XX科技股份有限公司（盖章）", "仿宋", 16)
    para("日期：2025 年 3 月", "仿宋", 16)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    para("目  录", "黑体", 16, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, t in enumerate(["第一章 项目概述", "1.1 建设背景", "第二章 技术方案"], 1):
        para(f"{t}{'.' * 30}{i}", "仿宋", 12)
    para("第一章 项目概述", "黑体", 16, True, pbb=True)
    para(_LSENT * 2, "仿宋", 12, indent_chars=200, line=1.5)
    para("1.1 建设背景", "黑体", 14, True)
    para("随着业务的快速发展，" + _LSENT, "仿宋", 12, indent_chars=200, line=1.5)
    para("第二章 技术方案", "黑体", 16, True, pbb=True)
    para("总体架构自下而上分为采集层、治理层与服务层，" + _LSENT, "仿宋", 12, indent_chars=200, line=1.5)
    para("表 3-1 人员配置表", "宋体", 9, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t = d.add_table(rows=2, cols=3)
    _tbl_borders_single(t)
    for ci, h in enumerate(("岗位", "人数", "职责")):
        c = t.cell(0, ci)
        _shade(c, "D9D9D9")
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run_fmt(p.add_run(h), "宋体", 9, True)
    for ci, v in enumerate(("项目经理", "1", "总体负责")):
        for p in t.cell(1, ci).paragraphs:
            _run_fmt(p.add_run(v), "宋体", 9)
    pic_p = d.add_paragraph()
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.add_run().add_picture(str(_png(name)), width=Cm(10))
    para("图 3-1 网络拓扑图", "宋体", 9, align=WD_ALIGN_PARAGRAPH.CENTER)
    hp = sec.header.paragraphs[0]
    hp.text = "XX科技股份有限公司"
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _p_bottom_border(hp)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _field(fp, "PAGE", "1")
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _png(name):
    import io
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (60, 40), "#4060A0").save(buf, format="PNG")
    path = WORK / f"{name}.png"
    path.write_bytes(buf.getvalue())
    return path


def _company_doc(name, cjk_off=0):
    """结构化形态（深圳奥哲技术标）。cjk_off=末尾几段关掉中西文间距与数字间距。"""
    from docx.enum.section import WD_SECTION_START
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    d = _doc()
    s1 = d.sections[0]
    s1.page_width, s1.page_height = Cm(21.0), Cm(29.7)
    s1.top_margin = s1.bottom_margin = Cm(2.54)
    s1.left_margin = s1.right_margin = Cm(3.17)

    def para(text, ea, pt, bold=False, line=None, indent_chars=None, align=None, cjk=False):
        p = d.add_paragraph()
        if align is not None:
            p.alignment = align
        if line is not None:
            p.paragraph_format.line_spacing = line
        if indent_chars is not None:
            p._p.get_or_add_pPr().get_or_add_ind().set(qn("w:firstLineChars"), str(indent_chars))
        if cjk:
            ppr = p._p.get_or_add_pPr()
            for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
                el = OxmlElement(tag)
                el.set(qn("w:val"), "0")
                ppr.insert(0, el)
        _run_fmt(p.add_run(text), ea, pt, bold)
        return p

    # —— 第 1 节：封面（文本框标题 + 信息表格）——
    for _ in range(3):
        d.add_paragraph()
    d.add_paragraph()._p.append(_textbox_run(["竞争性磋商采购响应文件技术标"]))
    d.add_paragraph()
    d.add_paragraph()._p.append(_textbox_run(["项目编号：XNBH-2025-031", "项目名称：客户关系管理系统建设项目"]))
    ct = d.add_table(rows=1, cols=2)
    for ci, v in enumerate(("供应商名称：深圳奥哲网络科技有限公司", "日期：2025 年 3 月")):
        _run_fmt(ct.cell(0, ci).paragraphs[0].add_run(v), "仿宋", 12)
    s2 = d.add_section(WD_SECTION_START.NEW_PAGE)

    # —— 第 2 节：前置页（评分索引表）——
    s2.left_margin = s2.right_margin = Cm(1.91)
    para("技术标评分索引表（分）", "宋体", 12, True)
    ft = d.add_table(rows=2, cols=2)
    for ri, row in enumerate((("序号", "评分项目"), ("1", "技术方案"))):
        for ci, v in enumerate(row):
            _run_fmt(ft.cell(ri, ci).paragraphs[0].add_run(v), "宋体", 10.5)
    s3 = d.add_section(WD_SECTION_START.NEW_PAGE)

    # —— 第 3 节：目录 + 正文（页面设置与前两节都不同）——
    s3.top_margin, s3.bottom_margin = Cm(3.0), Cm(2.54)
    s3.left_margin, s3.right_margin = Cm(2.6), Cm(2.4)
    grid = s3._sectPr.find(qn("w:docGrid"))
    if grid is None:
        grid = OxmlElement("w:docGrid")
        s3._sectPr.append(grid)
    grid.set(qn("w:type"), "lines")
    grid.set(qn("w:linePitch"), "312")

    para("目  录", "黑体", 16, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for lvl, text in ((1, "1 技术服务响应表"), (2, "1.1 补充说明"), (3, "1.1.1 说明明细")):
        try:
            st = d.styles.add_style(f"toc {lvl}", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
        except ValueError:
            st = d.styles[f"toc {lvl}"]
        p = d.add_paragraph(style=st)
        p.paragraph_format.left_indent = Cm(0.74 * (lvl - 1))
        p.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        _run_fmt(p.add_run(text), "黑体" if lvl == 1 else "宋体", 12, lvl == 1)
        _run_fmt(p.add_run("\t"), "宋体", 12)
        _field(p, "PAGEREF _Toc1", str(lvl), prefix="", suffix="")

    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)  # 目录后的分页段

    def heading(text, lvl, pt):
        h = d.add_heading(text, level=lvl)
        for r in h.runs:
            _run_fmt(r, "黑体", pt, True)
        return h

    heading("技术服务响应表", 1, 16)
    for i in range(3):
        para(_LSENT, "仿宋", 12, line=1.5, indent_chars=200)
    heading("补充说明", 2, 14)
    p = para(_LSENT, "仿宋", 12, line=1.5, indent_chars=200)
    p._p.get_or_add_pPr().get_or_add_spacing().set(qn("w:beforeLines"), "100")  # 段前按「行」写
    heading("说明明细", 3, 12)
    for i in range(2):
        para(_LSENT, "仿宋", 12, line=1.5, indent_chars=200, cjk=(i < cjk_off))
    # 自动编号的列表段（格式随编号走，不进正文候选）
    numbering = d.part.numbering_part.element
    abs_el = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
        '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>')
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        numbering.append(abs_el)
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
    lp = para(_LSENT, "仿宋", 12)
    npr = OxmlElement("w:numPr")
    for tag, val in (("w:ilvl", "0"), ("w:numId", "90")):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        npr.append(e)
    lp._p.get_or_add_pPr().insert(0, npr)
    heading("实施保障", 4, 12)
    heading("质量保证", 5, 12)
    heading("应急预案", 6, 12)

    # 表格样式定义边框、表头底纹与加粗；表本身不直接设
    d.styles.element.append(parse_xml(
        f'<w:style {nsdecls("w")} w:type="table" w:styleId="CompanyTable"><w:name w:val="Company Table"/>'
        '<w:tblPr><w:tblBorders>'
        + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                  for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
        + '</w:tblBorders><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'
        '</w:tblCellMar></w:tblPr>'
        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>'
        '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr></w:style>'))
    para("表 1-1 人员配置表", "宋体", 9, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t = d.add_table(rows=2, cols=3)
    t._tbl.tblPr.insert(0, parse_xml(f'<w:tblStyle {nsdecls("w")} w:val="CompanyTable"/>'))
    tw_el = t._tbl.tblPr.find(qn("w:tblW"))   # 表宽写成占版心的百分比
    tw_el.set(qn("w:type"), "pct")
    tw_el.set(qn("w:w"), "5000")
    t._tbl.tblPr.append(parse_xml(
        f'<w:tblLook {nsdecls("w")} w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" '
        'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'))
    for ci, h in enumerate(("岗位", "人数", "职责")):
        c = t.cell(0, ci)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run_fmt(c.paragraphs[0].add_run(h), "宋体", 9)
    for ci, v in enumerate(("项目经理", "1", "总体负责")):
        _run_fmt(t.cell(1, ci).paragraphs[0].add_run(v), "宋体", 9)
    heading("项目实施人员一览表", 1, 16)  # 第二个一级标题：前面没有分页
    para(_LSENT, "仿宋", 12, line=1.5, indent_chars=200)
    pic_p = d.add_paragraph()
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.add_run().add_picture(str(_png(name)), width=Cm(10))
    para("图 1-1 网络拓扑图", "宋体", 9, align=WD_ALIGN_PARAGRAPH.CENTER)

    s3.header.is_linked_to_previous = False
    hp = s3.header.paragraphs[0]
    hp.text = "深圳奥哲网络科技有限公司"
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _p_bottom_border(hp)
    s3.footer.is_linked_to_previous = False
    fp = s3.footer.paragraphs[0]
    fp.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_TAB_ALIGNMENT.RIGHT)
    _run_fmt(fp.add_run("\t"), "宋体", 9)
    _field(fp, "PAGE", "1")
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _plain_doc(name, fig_secs=12, big_ch=7):
    """纯文字形态（真实回放：一份技术标的正文加附图册，2026-11-27）。全篇一段样式都不挂，外观
    全部取自正文默认样式——标题跟正文同形；目录是手打的点线条目，二级条目以全角空格起头；
    纵向正文节与横向图页节交替；页脚只在第 1 节定义、后面的节沿用。big_ch 那一章正文段最多。"""
    from docx.enum.section import WD_ORIENT, WD_SECTION_START
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

    cn = ("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二")
    d = _doc()
    normal = d.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    def portrait(sec):
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.top_margin = sec.bottom_margin = Cm(2.0)
        sec.left_margin = sec.right_margin = Cm(2.0)

    def landscape(sec):
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
        sec.top_margin = sec.bottom_margin = Cm(1.5)
        sec.left_margin = sec.right_margin = Cm(1.5)

    def para(text):
        p = d.add_paragraph(text)
        p._p.get_or_add_pPr().get_or_add_ind().set(qn("w:firstLine"), "480")
        return p

    portrait(d.sections[0])
    for t in ("XX网络科技有限公司", "技术标投标文件", "日期：2026 年 9 月"):
        d.add_paragraph(t)
    fp = d.sections[0].footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _field(fp, "PAGE", "1")
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph("目录")
    d.add_paragraph("编制说明" + "." * 40 + "1")
    for t in ("一、文件构成与应答范围", "二、编制依据", "三、应答口径"):
        d.add_paragraph("\u3000\u3000" + t + "." * 30 + "2")
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    for k in range(1, fig_secs + 1):
        p = d.add_paragraph(f"第{cn[k - 1]}章\u3000项目理解与建设目标")
        p.paragraph_format.page_break_before = True
        d.add_paragraph("1.1\u3000项目理解")
        for _ in range(30 if k == big_ch else 6):
            para(_LSENT * 2)
        d.add_paragraph("1.1.1\u3000采购方现状")
        for _ in range(3):
            para(_LSENT * 2)
        d.add_section(WD_SECTION_START.NEW_PAGE)   # 上面的文字归一个纵向正文节
        landscape(d.sections[-1])                  # 横向图页节
        d.add_paragraph().add_run().add_picture(str(_png(name)), width=Cm(24))
        cap = d.add_paragraph(f"图 {k}-1\u3000业务流程")
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        d.add_section(WD_SECTION_START.NEW_PAGE)   # 图页节到此为止
        portrait(d.sections[-1])                   # 回到纵向节，接着写下一章
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _ragged_secs_doc(name):
    """十四节、每一节的页面设置都不一样（真实回放：一份招标文件的十个节里八个设置相同却逐节罗列；
    这里让它们各不相同，验「超过一行装得下的量」时的收法）。正文段第 7 节最多。"""
    from docx.enum.section import WD_ORIENT, WD_SECTION_START

    d = _doc()
    for i in range(1, 15):
        if i > 1:
            d.add_section(WD_SECTION_START.NEW_PAGE)
        sec = d.sections[-1]
        sec.orientation = WD_ORIENT.LANDSCAPE if i % 2 else WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(29.7 if i % 2 else 21.0), Cm(21.0 if i % 2 else 29.7)
        sec.top_margin = sec.bottom_margin = Cm(1.5 + 0.1 * i)
        sec.left_margin = sec.right_margin = Cm(1.0 + 0.1 * i)
        for _ in range(20 if i == 7 else 1):
            d.add_paragraph(_LSENT * 2)
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _row(out, item):
    for ln in out.splitlines():
        if ln.startswith(f"| {item} |"):
            return ln
    return ""


def _val(out, item):
    row = _row(out, item)
    return row.split("|")[2].strip() if row else ""


def e55():
    """页面组（多节文档）：纸张边距取正文所在节、各节差异、网格类型、页眉页码跨节与位置。"""
    p = _company_doc("公司标书")
    sha = p.read_bytes()
    code, out = tk("layout-read", p)
    check("E55a 退出码 0、结论先行、只读声明、原件未动", code == 0 and out.startswith("[版式读取]")
          and "只读" in out and p.read_bytes() == sha, f"exit={code} {out[:300]}")
    check("E55b 纸张 A4 纵向，依据写明正文所在第 3 节", "A4 纵向" in _val(out, "纸张")
          and "第 3 节" in _row(out, "纸张"), _row(out, "纸张"))
    check("E55c 页边距取正文节（不是封面节的左右 3.17）",
          _val(out, "页边距") == "上 3、下 2.54、左 2.6、右 2.4 厘米", _row(out, "页边距"))
    diff = _row(out, "页面设置差异")
    check("E55d 各节页面设置差异点名（第 1 节左右 3.17、第 2 节左右 1.91）",
          "第 1 节" in diff and "3.17" in diff and "第 2 节" in diff and "1.91" in diff, diff)
    check("E55e 文档网格按类型读（type=lines 才叫行网格）", "行网格" in _val(out, "文档网格")
          and "行" in _val(out, "文档网格"), _row(out, "文档网格"))
    hd = _row(out, "页眉")
    check("E55f 页眉：第 3 节有（文字、居中、横线），第 1–2 节无",
          "深圳奥哲网络科技有限公司" in hd and "居中" in hd and "横线" in hd and "第 1–2 节无" in hd, hd)
    pn = _row(out, "页码写法")
    check("E55g 页码：第 X 页、靠右制表位定位、第 3 节起", "第 X 页" in pn and "页脚右" in pn
          and "第 3 节" in pn, pn)
    check("E55h 对称页边距、页脚其他文字各一行", _val(out, "对称页边距") == "否"
          and _val(out, "页脚其他文字") == "无", out[:600])


def e56():
    """正文组：字体字号行距缩进对齐（缺省=左）、段前段后、中文版式真实读取、列表段。"""
    p = _company_doc("正文组")
    code, out = tk("layout-read", p)
    check("E56a 正文字体、字号、行距、缩进", _val(out, "正文字体") == "中文仿宋、西文 Times New Roman"
          and _val(out, "正文字号") == "小四（12 磅）" and _val(out, "行距") == "1.5 倍"
          and _val(out, "缩进") == "首行 2 字", out[:900])
    check("E56b 对齐没设就是左对齐（Word 缺省），不是两端", _val(out, "对齐") == "左"
          and "确定" in _row(out, "对齐"), _row(out, "对齐"))
    check("E56c 段前段后：样式链上没设、文档默认值里是段后 10 磅 → 段后 10 磅（确定）",
          _val(out, "段前段后") == "段后 10 磅" and "确定" in _row(out, "段前段后"), _row(out, "段前段后"))
    check("E56d 中文版式：多数标准 → 标准（确定）", _val(out, "中文版式") == "标准"
          and "确定" in _row(out, "中文版式"), _row(out, "中文版式"))
    ls = _row(out, "列表习惯")
    check("E56e 列表段不进正文候选；列表写法读出「编号 1.」", "7/7" in _row(out, "正文字号")
          and "1 段" in ls and "1." in ls, ls + " ｜ " + _row(out, "正文字号"))
    # 两段关掉中西文间距 → 与标准并存 → 疑似并列
    p2 = _company_doc("中文版式混排", cjk_off=2)
    code, out = tk("layout-read", p2)
    v = _val(out, "中文版式")
    check("E56f 中文版式真实读取：关了的读出来，两种并存记疑似",
          "疑似" in _row(out, "中文版式") and "标准" in v and "关闭了中西文间距" in v and "数字间距" in v,
          _row(out, "中文版式"))


def e57():
    """直接格式形态（第X章自带编号）：编号形态/间隔/章起新页/标题格式；混排疑似并列且不截断。"""
    p = _direct_doc("直接格式标书")
    code, out = tk("layout-read", p)
    check("E57a 编号形态 chapter、依据是标题文字自带编号", "chapter" in _val(out, "编号形态")
          and "自带编号" in _row(out, "编号形态") and "确定" in _row(out, "编号形态"), _row(out, "编号形态"))
    check("E57b 标题一级/二级读出（直接格式）", "中文黑体" in _val(out, "标题一级")
          and "三号（16 磅）" in _val(out, "标题一级") and "四号（14 磅）" in _val(out, "标题二级"),
          out[:1200])
    check("E57c 编号间隔 半角空格；章起新页 是（自身段前分页）", "半角空格" in _val(out, "编号间隔")
          and _val(out, "章起新页") == "是", _row(out, "编号间隔") + _row(out, "章起新页"))
    check("E57d 事实行：命名标题 0 段、全是直接格式", "命名标题样式 0/" in out and "全是直接格式" in out,
          out[-500:])
    # 行距混排 → 疑似并列
    d = Document(str(p))
    ps = [q for q in d.paragraphs if len(q.text) >= 30 and not re.search(r"\.{4,}", q.text)]
    ps[0].paragraph_format.line_spacing = 1.5
    ps[1].paragraph_format.line_spacing = Pt(28)
    d.save(str(p))
    code, out = tk("layout-read", p)
    v = _val(out, "行距")
    check("E57e 行距两种并存 → 疑似且并列两种", "疑似" in _row(out, "行距") and "1.5 倍" in v and "28 磅" in v, v)
    # 标题两种格式并存 → 并列的两种都要完整可见，不许截断
    d = Document(str(p))
    heads = [q for q in d.paragraphs if re.match(r"^第[一二三四五六七八九十]+章\s+\S", q.text)
             and not re.search(r"\.{3,}", q.text)]   # 只挑正文里的章标题，目录里的点线条目不算
    assert len(heads) == 2, f"章标题应有两处：{[q.text[:10] for q in heads]}"
    heads[1].runs[0].font.size = Pt(15)
    d.save(str(p))
    code, out = tk("layout-read", p)
    v = _val(out, "标题一级")
    check("E57f 疑似行的两种取值完整呈现（不被 70 字上限截掉）",
          "疑似" in _row(out, "标题一级") and "16 磅" in v and "15 磅" in v and "…" not in v, v)
    check("E57g 表格直接边框全线、表头底纹 D9D9D9 加粗居中", "全线" in _val(out, "表格边框")
          and "D9D9D9" in _val(out, "表头样式") and "表头加粗" in _val(out, "表头样式")
          and "表头居中" in _val(out, "表头样式"), out[-1500:])
    check("E57h 文档网格没写类型 = 无网格", "无网格" in _val(out, "文档网格"), _row(out, "文档网格"))


def e58():
    """目录组：标题格式、条目分级、层级、点线；目录条目不进标题候选。"""
    p = _company_doc("目录组")
    code, out = tk("layout-read", p)
    check("E58a 目录标题格式", "中文黑体" in _val(out, "目录标题") and "三号（16 磅）" in _val(out, "目录标题")
          and "居中" in _val(out, "目录标题"), _row(out, "目录标题"))
    check("E58b 目录条目分级读出", "加粗" in _val(out, "目录一级条目")
          and "加粗" not in _val(out, "目录二级条目") and _row(out, "目录三级条目") != "", out[1200:2400])
    check("E58c 目录层级 3 级、点线引导、缩进逐级", "3 级" in _val(out, "目录层级")
          and "点线" in _val(out, "目录点线") and "0.74" in _val(out, "目录缩进"),
          _row(out, "目录层级") + _row(out, "目录点线") + _row(out, "目录缩进"))
    scope = out[out.index("已检查："):]
    check("E58d 目录条目不进标题候选：标题候选 7 段、目录条目 3 条",
          "标题候选 7 段" in scope and "目录条目 3 条" in scope, scope)
    check("E58e 正文候选 7 段（目录条目不算正文）", "正文候选 7 段" in scope, scope)


def e59():
    """标题组（命名样式、用到六级、文字不带编号）：层级、五级及以上、编号形态来源、事实行。"""
    p = _company_doc("标题组")
    code, out = tk("layout-read", p)
    check("E59a 标题层级读到 6 级（不截到四级）", "6 级" in _val(out, "标题层级"), _row(out, "标题层级"))
    check("E59b 五级及以上单列一行", _row(out, "标题五级及以上") != ""
          and "中文黑体" in _val(out, "标题五级及以上"), _row(out, "标题五级及以上"))
    v, r = _val(out, "编号形态"), _row(out, "编号形态")
    check("E59c 标题文字不带编号：形态标疑似，依据说明来自目录条目",
          "疑似" in r and "不带编号" in r and "目录条目" in r, r)
    check("E59d 标题一级格式含对齐", "中文黑体" in _val(out, "标题一级") and "三号（16 磅）" in _val(out, "标题一级")
          and "左" in _val(out, "标题一级"), _row(out, "标题一级"))
    check("E59e 章起新页：首章前有分页、第二章没有 → 否", _val(out, "章起新页") == "否"
          and "2 处" in _row(out, "章起新页"), _row(out, "章起新页"))
    check("E59f 事实行：命名标题 7/7、走命名样式", "命名标题样式 7/7" in out and "走命名样式" in out, out[-500:])


def e60():
    """表格与图：样式边框与表头底纹、表宽、单元格边距与对齐、表题图题位置与编号、图宽占比。"""
    p = _company_doc("表格图")
    code, out = tk("layout-read", p)
    check("E60a 表格边框来自表格样式 → 全线 0.5 磅", "全线" in _val(out, "表格边框")
          and "0.5 磅" in _val(out, "表格边框"), _row(out, "表格边框"))
    hs = _val(out, "表头样式")
    check("E60b 表头底纹与加粗来自样式（表本身不直接设）", "D9D9D9" in hs and "表头加粗" in hs
          and "表头居中" in hs, hs)
    check("E60c 表格宽度占版心、单元格边距", "100%" in _val(out, "表格宽度")
          and "0.14" in _val(out, "单元格边距"), _row(out, "表格宽度") + _row(out, "表格宽度")
          + _row(out, "单元格边距"))
    check("E60d 表格文字含字号与对齐", "小五（9 磅）" in _val(out, "表格文字") and "左" in _val(out, "表格文字"),
          _row(out, "表格文字"))
    cap = _val(out, "表题")
    check("E60e 表题：位置读出在表上方、编号 N-M、居中", "表上方" in cap and "表 N-M" in cap and "居中" in cap, cap)
    cap = _val(out, "图题")
    check("E60f 图题：位置读出在图下方、编号 N-M", "图下方" in cap and "图 N-M" in cap, cap)
    check("E60g 图宽 10 厘米并给出占版心的比例（版心 16 厘米 → 62.5%，按缇算进位 63%）",
          "10 厘米" in _val(out, "图宽") and "63%" in _val(out, "图宽"), _row(out, "图宽"))
    # 表题放在表格下面：位置要跟着读出来
    d = Document(str(p))
    from docx.text.paragraph import Paragraph
    kids = list(d.element.body.iterchildren())
    ti = next(i for i, k in enumerate(kids) if k.tag == qn("w:tbl") and k.find(f".//{qn('w:tblStyle')}") is not None)
    cap_el = kids[ti - 1]
    kids[ti - 1].getparent().remove(cap_el)
    kids[ti].addnext(cap_el)
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E60h 表题挪到表下方后读出「表下方」（位置是读出来的，不是写死的）",
          "表下方" in _val(out, "表题"), _row(out, "表题"))


def e61():
    """封面：区域止于分节符；文本框去重分行；封面表格；字段标记；页面设置与正文不同。"""
    p = _company_doc("封面")
    code, out = tk("layout-read", p)
    block = out[out.index("封面文字"):] if "封面文字" in out else ""
    lines = [ln.strip() for ln in block.splitlines()[1:] if ln.strip().startswith(("→", "·"))]
    check("E61a 文本框文字只读一遍（兼容分支去重）",
          sum("竞争性磋商采购响应文件技术标" in ln for ln in lines) == 1
          and "竞争性磋商采购响应文件技术标竞争性磋商" not in block, block[:400])
    check("E61b 文本框内换行分成两行，各自完整", any("项目编号：XNBH-2025-031" in ln and "项目名称" not in ln
                                                       for ln in lines)
          and any("项目名称：客户关系管理系统建设项目" in ln for ln in lines), block[:500])
    check("E61c 封面表格里的文字也列出", any("供应商名称：深圳奥哲网络科技有限公司" in ln
                                          and "日期：2025 年 3 月" in ln for ln in lines), block[:500])
    check("E61d 封面区止于分节符：前置页标题不算封面", "技术标评分索引表" not in block, block[:500])
    check("E61e 字段标记：编号与日期行标 →，固定大字行标 ·",
          any("项目编号" in ln and ln.startswith("→") for ln in lines)
          and any("竞争性磋商采购响应文件技术标" in ln and ln.startswith("·") for ln in lines), block[:500])
    check("E61f 封面字段汇总", all(k in _val(out, "封面字段") for k in ("编号", "项目名称", "供应商", "日期")),
          _row(out, "封面字段"))
    check("E61g 封面版式清点文字行、表格、文本框", "表格 1 张" in _val(out, "封面版式")
          and "文本框 2 个" in _val(out, "封面版式"), _row(out, "封面版式"))
    check("E61h 封面页设置与正文不同并点名差异项", "与正文不同" in _val(out, "封面页设置")
          and "3.17" in _val(out, "封面页设置"), _row(out, "封面页设置"))


def e62():
    """落盘：--out 规格表完整不截断；--json 多文件为列表；原件只读。"""
    p = _direct_doc("落盘源")
    d = Document(str(p))
    heads = [q for q in d.paragraphs if re.match(r"^第[一二三四五六七八九十]+章\s+\S", q.text)
             and not re.search(r"\.{3,}", q.text)]   # 只挑正文里的章标题，目录里的点线条目不算
    assert len(heads) == 2, f"章标题应有两处：{[q.text[:10] for q in heads]}"
    heads[1].runs[0].font.size = Pt(15)
    d.save(str(p))
    sha = p.read_bytes()
    spec = WORK / "落盘-规格.md"
    code, out = tk("layout-read", p, "--out", spec)
    text = spec.read_text(encoding="utf-8") if spec.is_file() else ""
    row = next((ln for ln in text.splitlines() if ln.startswith("| 标题一级 |")), "")
    check("E62a --out 落盘规格表：四列表头、来源行、疑似行两种取值完整",
          code == 0 and "| 项目 | 取值 | 依据 | 确定/疑似 |" in text and "sha256" in text
          and "16 磅" in row and "15 磅" in row and "…" not in row, row or out[:300])
    p2 = _direct_doc("落盘源二")
    jp = WORK / "落盘-结果.json"
    code, out = tk("layout-read", p, p2, "--json", jp)
    data = json.loads(jp.read_text(encoding="utf-8")) if jp.is_file() else {}
    files = data.get("files", [])
    check("E62b --json 多文件：列表保留每一份，不以后一份覆盖前一份",
          code == 0 and len(files) == 2 and files[0]["file"] != files[1]["file"]
          and all("rows" in f for f in files), str(data)[:300])
    check("E62c 原件未被改动（逐字节一致）", p.read_bytes() == sha, "")


def e63():
    """.doc 输入：LibreOffice 转换后读取；缺 LibreOffice 记失败（与 E24 同口径）；PDF 拒收。"""
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    soffice = tk_mod._find_soffice()
    if not soffice:
        raise NotImplementedError("本机无 LibreOffice，跳过")
    src = _direct_doc("转档源")
    with tk_mod._tempdir() as td:
        r = subprocess.run([soffice, "--headless", f"-env:UserInstallation={Path(td).resolve().as_uri()}",
                            "--convert-to", "doc", "--outdir", td, str(src)],
                           capture_output=True, text=True, errors="replace", timeout=300)
        doc_path = Path(td) / (src.stem + ".doc")
        assert doc_path.is_file(), f"转 .doc 失败：{r.stdout[:200]} {r.stderr[:200]}"
        code, out = tk("layout-read", doc_path)
        check("E63a .doc 读取成功并回显转换", code == 0 and "已转" in out and "Traceback" not in out,
              f"exit={code} {out[:300]}")
        check("E63b 正文读出仿宋小四", "仿宋" in _val(out, "正文字体") and "12 磅" in _val(out, "正文字号"), out[:600])
        check("E63c 纸张 A4 纵向", "A4 纵向" in _val(out, "纸张"), out[:300])
    bad = WORK / "不收.pdf"
    bad.write_bytes(b"%PDF-1.4 fake")
    code, out = tk("layout-read", bad)
    check("E63d PDF 明确不收（提示给 Word 版，退出码非 0，无报错栈）",
          code != 0 and "Word" in out and "Traceback" not in out, f"exit={code} {out[:200]}")


def e64():
    """段前段后：单边设置不崩、按「行」写读成行、自动间距；绝对磅值缩进归一成字。"""
    p = WORK / "段距.docx"
    d = _doc()
    for kind in ("单边", "行", "自动", "单边", "行", "自动"):
        q = d.add_paragraph("这一段的段前段后写法各不相同，读取时应当逐段解析并在取多数时如实反映出来。")
        sp = OxmlElement("w:spacing")
        if kind == "单边":
            sp.set(qn("w:before"), "120")
        elif kind == "行":
            sp.set(qn("w:beforeLines"), "100")
        else:
            sp.set(qn("w:beforeAutospacing"), "1")
        ppr = q._p.get_or_add_pPr()
        ppr.append(sp)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), "480")  # 绝对磅值写的首行缩进
        ppr.append(ind)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "24")
        q.runs[0]._element.get_or_add_rPr().append(sz)
    d.save(str(p))
    code, out = tk("layout-read", p)
    row = _row(out, "段前段后")
    v = _val(out, "段前段后")
    check("E64a 退出码 0、无报错栈", code == 0 and "Traceback" not in out, f"exit={code} {out[:300]}")
    check("E64b 三种写法并存 → 疑似，前三种都列出来（磅/行/自动），None 不进格式化",
          "疑似" in row and "段前 6 磅" in v and "段前 1 行" in v and "段前自动" in v, row)
    # 绝对磅值首行缩进按生效字号归一（480 缇 @ 12 磅 = 2 字）
    check("E64c 480 缇@12磅 读成 首行 2 字", _val(out, "缩进") == "首行 2 字", _row(out, "缩进"))
    p2 = WORK / "段距行.docx"
    d2 = _doc()
    for _ in range(3):
        q = d2.add_paragraph("这一段的段前间距是按行写的，三段都一样，读取时应当确定地读成段前 1 行才对。")
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:beforeLines"), "100")
        sp.set(qn("w:after"), "120")
        q._p.get_or_add_pPr().append(sp)
    d2.save(str(p2))
    code, out = tk("layout-read", p2)
    check("E64d 段前 1 行、段后 6 磅（行与磅混写）", _val(out, "段前段后") == "段前 1 行、段后 6 磅",
          _row(out, "段前段后"))


def e65():
    """主题字体属性优先于显式字体；奇偶页页眉不同；标题样式的对齐缺省是左。"""
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    tk_mod.require_docx()
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    rpr = parse_xml(f'<w:rPr {nsdecls("w")}><w:rFonts w:eastAsia="宋体" w:eastAsiaTheme="majorEastAsia" '
                    'w:asciiTheme="minorHAnsi"/></w:rPr>')
    ea, asc = tk_mod._fonts_of([rpr], {"majorEa": "黑体", "minorLatin": "Calibri"}, {})
    check("E65a 同一 rFonts 里主题属性优先：读出黑体与 Calibri，不是宋体", ea == "黑体" and asc == "Calibri",
          f"{ea} {asc}")
    p = _company_doc("奇偶页")
    d = Document(str(p))
    d.settings.element.append(parse_xml(f'<w:evenAndOddHeaders {nsdecls("w")}/>'))
    s3 = d.sections[2]
    s3.even_page_footer.is_linked_to_previous = False
    s3.even_page_footer.paragraphs[0].text = "偶数页页脚"
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E65b 开了奇偶页不同并定义了偶数页页脚 → 点名", "奇偶页不同" in out, out[:900])
    h1 = _val(out, "标题一级")
    check("E65c 标题没设对齐 → 读成左，并显式写进行里", h1.endswith("左"), h1)


def e66():
    """招标文件形态：目录条目是不带页码、不带样式的纯文字行（到分页符为止）；手打编号的
    条款行（以句末标点收尾，或不加粗又超过 20 字）不算标题。"""
    from docx.enum.text import WD_BREAK
    p = WORK / "纯文字目录.docx"
    d = _doc()

    def para(text, ea, pt, bold=False):
        q = d.add_paragraph()
        _run_fmt(q.add_run(text), ea, pt, bold)
        return q

    para("公式标注竞争性磋商文件", "仿宋", 42, True)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    para("目  录", "黑体", 16, True)
    d.add_paragraph()
    for num, name in zip(("一", "二", "三", "四"), ("采购邀请", "采购需求", "供应商须知", "评审办法")):
        para(f"第{num}章  {name}", "仿宋", 15, True)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    para("第一章  采购邀请", "仿宋", 16, True)
    para("本项目旨在建设一套完整的平台，" + _LSENT, "仿宋", 10.5)
    para("2.投标人须具备独立法人资格，并提供有效营业执照复印件。", "仿宋", 10.5)   # 条款：句号收尾
    para("3.本项目不接受联合体投标，投标人须独自完成全部服务内容并承担相应责任", "仿宋", 10.5)  # 条款：不加粗超 20 字
    # 带编号的长条款（超过 40 字）：是正文不是标题——改写时曾把长度条件挪丢，整份招标文件里
    # 以数字开头的长段全被算成了标题（真实文档回放抓到的回归）
    para("4.供应商应当对所提交的响应文件的真实性、完整性、合法性承担法律责任，并接受采购人"
         "及监管部门的监督检查，如有弄虚作假行为将依法承担相应责任。", "仿宋", 10.5)
    para("第二章  采购需求", "仿宋", 16, True)
    para("总体架构说明如下。" + _LSENT, "仿宋", 10.5)
    d.save(str(p))
    code, out = tk("layout-read", p)
    scope = out[out.index("已检查："):]
    check("E66a 纯文字目录条目认出来：4 条、1 级", code == 0 and "目录条目 4 条" in scope
          and _val(out, "目录层级") == "1 级", scope)
    check("E66b 目录条目不进标题候选：一级标题 2/2", "2/2" in _row(out, "标题一级")
          and "标题候选 2 段" in scope, _row(out, "标题一级") + scope)
    check("E66c 手打编号的条款行（句号收尾／不加粗超 20 字／超过 40 字）都不算标题",
          "标题候选 2 段" in scope and "正文候选 4 段" in scope, scope)


# ---------------------------------------------------------------- 第 9 批：分析阶段交付链接

BR = [sys.executable, "-B", str(ROOT / "tender-review" / "scripts" / "build_review.py")]


def br(work, *a):
    """跑 build_review 命令行，返回 (退出码, 合并输出)。"""
    r = subprocess.run(BR + ["--work", str(work)] + [str(x) for x in a],
                       capture_output=True, text=True, errors="replace",
                       env=dict(os.environ, PYTHONDONTWRITEBYTECODE="1"))
    return r.returncode, (r.stdout + r.stderr).strip()


def _fake_service(work, port=55901, token="deadbeefdeadbeef"):
    """写一份假 serve.json——_serve_base 只读 port/token 回基址，不起真服务。"""
    sj = work / "review" / "serve.json"
    sj.parent.mkdir(parents=True, exist_ok=True)
    sj.write_text(json.dumps({"port": port, "token": token, "pid": 0,
                              "started": "2026-01-01 00:00", "work": str(work)}), encoding="utf-8")


def _echo_line(out, tag):
    return next((ln for ln in out.splitlines() if ln.startswith(tag)), "")


def e67():
    """真机事故（qszbh260186，2026-11-26）：分析环节回显的 [交付链接] 恒为评审台根，
    模型把评审台链接标成「打开招标文件分析报告」，用户点开落在评审台总览（待确认卡），
    报告全量页藏在入口卡后面——分析阶段（G1 异常门 / G2）的交付面就是报告页本身，
    [交付链接] 应为报告页直链。"""
    w = WORK / "br67"
    w.mkdir(parents=True, exist_ok=True)
    _fake_service(w)
    rc, out = br(w, "--pages", "report")
    line = _echo_line(out, "[交付链接]")
    check("E67 报告构建的交付链接=报告页直链（带 ?v=，非评审台根）",
          rc == 0 and "report.html?v=" in line and "deadbeefdeadbeef" in line,
          f"exit={rc} {line[:200]}")


def e68():
    """报告页自包含、可离线打开：--pages report 且服务未起时，[交付链接] 给报告页
    文件路径（静态交付），不许只回「服务未起」叫用户先起服务。"""
    w = WORK / "br68"
    w.mkdir(parents=True, exist_ok=True)
    rc, out = br(w, "--pages", "report")
    line = _echo_line(out, "[交付链接]")
    check("E68 无服务时交付链接=报告页文件路径（不是「暂不可用」死胡同）",
          rc == 0 and "招标文件分析报告.html" in line and "暂不可用" not in line,
          f"exit={rc} {line[:200]}")


def e69():
    """G3 起全量构建：评审台 [交付链接] 照旧（根地址+#tab 锚点提示），另回显
    [报告页链接] 一行——模型照抄回显，不自行把评审台链接改标签当报告交付。"""
    w = WORK / "br69"
    w.mkdir(parents=True, exist_ok=True)
    _fake_service(w)
    rc, out = br(w, "--pages", "all")
    link = _echo_line(out, "[交付链接]")
    rpt = _echo_line(out, "[报告页链接]")
    check("E69 全量构建交付链接仍是评审台并附报告页直链一行",
          rc == 0 and "/deadbeefdeadbeef/?v=" in link and "#tab=" in link
          and "report.html?v=" in rpt,
          f"exit={rc} {link[:120]} | {rpt[:120]}")


# ---------------------------------------------------------------- 第 10 批：结构要求散文可读性

def _fmt_md(work, prose_lines):
    """写一份最小 requirements-format.md——_fmt_payload 只解析「一、」节的散文与大纲。"""
    d = work / "analysis"
    d.mkdir(parents=True, exist_ok=True)
    lines = ["## 一、投标文件结构要求", ""] + list(prose_lines)
    lines += ["", "结构大纲：", "- 响应文件", "  - 资格文件", "",
              "## 二、必须有的章节（mandatory chapters）",
              "| 序号 | 章节名称 | 归属(bidPart) | 原文出处线索 |",
              "| 1 | 响应函 | commercial | 同文件 第六章（L1100-L1102） |"]
    (d / "requirements-format.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def e70():
    """真机事故（云栖工程，2026-11-26）：分析产物把约 10 个条款连排成一个 708 字物理行，
    报告页按源行分段（_fmt_payload 一行一条、模板一行一个 <p>），整段挤成一团没法逐条看。
    构建时对超长散文段回显软提醒（不拦流程、不改退出码），指路按条款拆行。"""
    w = WORK / "br70"
    long_line = ("响应文件由资格文件、报价文件、商务和技术文件三部分组成（同文件 第三章三、1.1（L552）），"
                 "三部分合并装订成册：正本一份、副本四份，封面注明「正本」「副本」字样；") \
        + "；".join(f"包装封面须注明供应商名称、地址、项目名称、项目编号及「开启时启封」字样并加盖公章"
                    f"（同文件 第三章三、2.2（L{560 + i}））" for i in range(6))
    _fmt_md(w, [long_line,
                "响应文件须在规定位置盖章并由供应商代表签署，不得涂改（同文件 第三章三、3.1（L568））。"])
    rc, out = br(w, "--pages", "report")
    line = _echo_line(out, "[结构要求超长段]")
    check("E70 超长段回显软提醒（段序+字数+拆行指路）",
          rc == 0 and line.startswith("[结构要求超长段]") and "第1段" in line
          and str(len(long_line)) in line and "拆行" in line,
          f"exit={rc} line={line[:220]}")


def e71():
    """阴性对照：全部散文段都在阈值内时不回显提醒——闸只打超长段，不误伤正常行文。"""
    w = WORK / "br71"
    _fmt_md(w, ["响应文件由资格文件、报价文件、商务和技术文件三部分组成（同文件 第三章三、1.1（L552））。",
                "▲活页装订（卡条、抽杆夹、订书机、散装）的响应文件无效（同文件 供应商须知前附表三、2.1（L318））。"])
    rc, out = br(w, "--pages", "report")
    check("E71 全部短段不提醒", rc == 0 and "[结构要求超长段]" not in out,
          f"exit={rc} {out[:200]}")


_CUSTOM_PROPS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ICV">'
    "<vt:lpwstr>88A1C0B7F5D94E21</vt:lpwstr></property>"
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="KSOProductBuildVer">'
    "<vt:lpwstr>12.1.0.19382</vt:lpwstr></property>"
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="客户编号">'
    "<vt:lpwstr>KH-009</vt:lpwstr></property></Properties>"
)


_DIRTY_CORE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    "<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/"
    "core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' "
    "xmlns:dcterms='http://purl.org/dc/terms/' "
    "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    "<dc:title>星灵创科投标文件</dc:title><dc:subject>YLCG-2026-GK-0387</dc:subject>"
    "<dc:creator>张三</dc:creator><cp:keywords>投标;响应文件</cp:keywords>"
    "<dc:description>投标文件初稿</dc:description><cp:lastModifiedBy>李四</cp:lastModifiedBy>"
    "<cp:revision>3</cp:revision>"
    "<dcterms:created xsi:type='dcterms:W3CDTF'>2024-05-06T08:00:00Z</dcterms:created>"
    "<dcterms:modified xsi:type='dcterms:W3CDTF'>2024-05-06T08:00:00Z</dcterms:modified>"
    "<cp:category>投标文件</cp:category></cp:coreProperties>"
)


def _dirty_docx(path, fresh=True):
    """带满格元数据的「历史标书」夹具：core 全字段脏值（zip 层整份替换 core.xml——
    E40 起进程内跑过 assemble 的用例经 docxcompose 改坏全局 nsmap["cp"]，此后属性
    API 建出的元素命名空间是错的，不能依赖）+ zip 层注入 docProps/custom.xml（WPS
    真实文件在这里写 ICV——云账号身份码）。fresh=True 现建一份——python-docx 默认
    模板自带 app.xml 的 Word 应用名与首页缩略图，正好当脏数据；False 时就地弄脏
    既有文件（保留原有样式骨架）。"""
    if fresh:
        d = Document()
        d.add_paragraph("正文一句话。")
        d.save(str(path))
    zin = zipfile.ZipFile(path)
    entries = [(i.filename, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    have_core = False
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries:
            if name == "docProps/core.xml":
                data, have_core = _DIRTY_CORE_XML.encode("utf-8"), True
            elif name == "[Content_Types].xml":
                data = data.decode("utf-8").replace(
                    "</Types>",
                    '<Override PartName="/docProps/custom.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.custom-properties+xml"'
                    "/></Types>").encode("utf-8")
            elif name == "_rels/.rels":
                data = data.decode("utf-8").replace(
                    "</Relationships>",
                    '<Relationship Id="rIdCust1" Type="http://schemas.openxmlformats.org/package'
                    '/2006/relationships/custom-properties" Target="docProps/custom.xml"'
                    "/></Relationships>").encode("utf-8")
            zout.writestr(name, data)
        if not have_core:
            zout.writestr("docProps/core.xml", _DIRTY_CORE_XML.encode("utf-8"))
        zout.writestr("docProps/custom.xml", _CUSTOM_PROPS_XML)


def e72():
    """统一落盘口清零元数据（compact 代表 revise/comment/image/toc-pages 同一落盘口）：
    改历史标书直接交付，不留原署名、不留 Word 应用与模板痕迹、不带 WPS 身份码与缩略图，
    正文照常处理。"""
    p = WORK / "历史标书-压缩.docx"
    _dirty_docx(p)
    code, out = tk("compact", "--docx", p, "--line-pt", "28")
    z = zipfile.ZipFile(p)
    names = z.namelist()
    core = z.read("docProps/core.xml").decode("utf-8")
    app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else ""
    check("E72a compact 落盘即清零：core 全字段空白（含标题/主题/关键词/类别）",
          code == 0 and "张三" not in core and "李四" not in core and "python-docx" not in core
          and "星灵创科" not in core and "YLCG" not in core and "投标" not in core,
          f"exit={code} core={core[:400]}")
    check("E72b app.xml 应用名/模板名/总编辑时长清零",
          "Microsoft" not in app and "Normal.dotm" not in app and "<TotalTime>" not in app,
          app[:400])
    check("E72c docProps/custom.xml 整体摘除（ICV 身份码不随文件走）",
          "docProps/custom.xml" not in names, str(names))
    check("E72d 包级首页缩略图摘除", "docProps/thumbnail.jpeg" not in names, str(names))
    check("E72e 正文照常压缩落盘", "[版式压缩]" in out and "正文段 1；" in out, out[:300])


def e73():
    """合册交付稿（composer 落盘路径）不带任何元数据——含混进 body/ 的外部脏节文件。"""
    work, _ = make_work("w-md", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n", ["甲节"])
    _dirty_docx(work / "body" / "甲节.docx", fresh=False)
    code, out = tk("assemble", "--work", work)
    book = next((work / "body").glob("整本-*.docx"))
    z = zipfile.ZipFile(book)
    names = z.namelist()
    core = z.read("docProps/core.xml").decode("utf-8")
    app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else ""
    check("E73a 合册整本 core/app/自定义/缩略图全清",
          code == 0 and "python-docx" not in core and "张三" not in core
          and "Microsoft" not in app and "Normal.dotm" not in app
          and "docProps/custom.xml" not in names and "docProps/thumbnail.jpeg" not in names,
          f"exit={code} core={core[:200]} app={app[:200]} names={names}")
    code2, out2 = ins(book)
    check("E73b 检查整本不再报元数据问题（自定义属性/缩略图/工具痕迹全无）",
          code2 == 0 and "自定义属性" not in out2 and "缩略图" not in out2
          and "工具痕迹" not in out2 and "作者或单位信息" not in out2,
          f"exit={code2} {out2[:400]}")


def e74():
    """inspect 点名元数据残留（自定义属性/首页缩略图/单位信息，ICV 与 KSO 记工具痕迹）；
    Application=Microsoft Office Word 是 Word 文件常态不算；清零后转阴性。"""
    p = WORK / "属性-残留.docx"
    _dirty_docx(p)
    jf = WORK / "属性-残留.json"
    code, out = ins(p, "--json", jf)
    check("E74a 自定义属性被点名（ICV/KSOProductBuildVer 记工具痕迹；全名看 --json）",
          code == 0 and "自定义属性" in out and "ICV" in out and "工具痕迹" in out
          and "KSOProductBuildVer" in jf.read_text(encoding="utf-8"), f"exit={code} {out[:500]}")
    check("E74b core 脏值与首页缩略图被点名，Word 应用名不算痕迹",
          "张三" in out and "YLCG-2026-GK-0387" in out and "缩略图" in out
          and "Microsoft" not in out, out[:500])
    code2, out2 = tk("compact", "--docx", p, "--line-pt", "28")
    code3, out3 = ins(p)
    check("E74c 清零后转阴性",
          code2 == 0 and code3 == 0 and "自定义属性" not in out3 and "缩略图" not in out3
          and "工具痕迹" not in out3 and "作者或单位信息" not in out3,
          f"exit={code3} {out3[:400]}")


# ---------------------------------------------------------------- 第 12 批 落盘路径防呆


def e75():
    """落盘路径防呆：--json 在 docx_toolkit 与 check_writing 里是「结果写到哪个文件」，不是
    「按 JSON 输出」的开关。真机上模型当开关用（inspect --json 标书.docx）：文件参数被 --json
    吃掉、检查了 0 份文件，然后把一份真实标书写成了 63 字节的报告。四扇落盘门写之前过
    同一道闸——与被读文件同路径、后缀不是它写出的格式，一律拒绝（退出码 2）且不写任何
    文件；inspect 没给文件同样报参数错（过去给的是「0 份文件：未发现问题」）。
    每条子用例各用一份夹具：闸失灵时有的用例会真把夹具写掉，不能让后面的用例读到它。"""
    def fresh(name):
        d = _doc()
        d.add_paragraph(f"{name} 的正文。")
        p = WORK / f"{name}.docx"
        d.save(str(p))
        return p, p.read_bytes()

    def refused(label, code, out, p, sha):
        check(f"{label}：退出码 2，没写任何文件", code == 2 and "[结果已落盘]" not in out,
              f"exit={code} {out[:300]}")
        check(f"{label}：原文件逐字节未变", p.is_file() and p.read_bytes() == sha,
              "文件被改动" if p.is_file() else "文件丢失")

    # ---- inspect --json ----
    a, a_sha = fresh("防呆-事故形态")
    code, out = ins("--json", a)
    refused("E75a inspect --json <docx>（事故形态）", code, out, a, a_sha)
    check("E75a2 报错点破 --json 是落盘路径、要查的文件跟在 inspect 后面",
          "落盘" in out and "inspect" in out and "未发现问题" not in out, out[:300])
    b_in, _ = fresh("防呆-另一输入")
    b, b_sha = fresh("防呆-落盘目标")
    code, out = ins(b_in, "--json", b)
    refused("E75b inspect 甲.docx --json 乙.docx（被写掉的是另一份）", code, out, b, b_sha)
    c, c_sha = fresh("防呆-同路径")
    code, out = ins(c, "--json", c)
    refused("E75c 显式同路径", code, out, c, c_sha)
    check("E75c2 报错说明写进去会覆盖它", "覆盖" in out, out[:300])
    work, sec = _amount_work("w-e75")
    book = _assemble(work)
    book_sha, sec_sha = book.read_bytes(), sec.read_bytes()
    code, out = ins("--work", work, "--json", book)
    refused("E75d --work 模式落盘到整本", code, out, book, book_sha)
    code, out = ins("--work", work, "--json", sec)
    refused("E75e --work 模式落盘到节文件", code, out, sec, sec_sha)
    f, f_sha = fresh("防呆-后缀")
    for bad in ("e75-结果.txt", "e75-无后缀"):
        target = WORK / bad
        code, out = ins(f, "--json", target)
        check(f"E75f 落盘路径 {bad} 须以 .json 结尾：退出码 2、未创建",
              code == 2 and not target.exists() and ".json" in out, f"exit={code} {out[:300]}")
    check("E75f2 被检查的文件未变", f.read_bytes() == f_sha)
    jf = WORK / "e75-无输入.json"
    code, out = ins("--json", jf)
    check("E75g 没给要检查的文件：退出码 2、不落盘、报错指路",
          code == 2 and not jf.exists() and "没给" in out and "inspect" in out, f"exit={code} {out[:300]}")
    code, out = ins()
    check("E75g2 一个参数都不给：退出码 2", code == 2 and "没给" in out, f"exit={code} {out[:300]}")

    # ---- 合法落盘照常 ----
    h, h_sha = fresh("防呆-合法")
    code, plain_out = ins(h)
    jf = WORK / "e75-嵌套" / "结果.json"
    code2, out = ins(h, "--json", jf)
    data = json.loads(jf.read_text(encoding="utf-8")) if jf.is_file() else {}
    stripped = "\n".join(ln for ln in out.splitlines() if "[结果已落盘]" not in ln).strip()
    check("E75h 合法落盘照常：退出码 0、结果文件生成、控制台只多一行落盘提示、被检查的文件未变",
          code == 0 and code2 == 0 and set(data) >= {"at", "files", "findings"}
          and f"[结果已落盘] {jf}" in out and stripped == plain_out and h.read_bytes() == h_sha,
          f"exit={code}/{code2} keys={sorted(data)}")

    # ---- layout-read：--json 与 --out 两扇门（每条各用一份夹具）----
    for flag, label in (("--json", "E75i layout-read --json 落到被读文件"),
                        ("--out", "E75j layout-read --out 落到被读文件")):
        lr = _direct_doc(f"防呆-版式{flag[2:]}")
        lr_sha = lr.read_bytes()
        code, out = tk("layout-read", lr, flag, lr)
        refused(label, code, out, lr, lr_sha)
    for flag, bad in (("--json", "e75-版式.txt"), ("--out", "e75-规格.json")):
        lr = _direct_doc(f"防呆-版式后缀{flag[2:]}")
        lr_sha = lr.read_bytes()
        target = WORK / bad
        code, out = tk("layout-read", lr, flag, target)
        check(f"E75k layout-read {flag} {bad} 后缀不符：退出码 2、未创建、被读取的文件未变",
              code == 2 and not target.exists() and lr.read_bytes() == lr_sha, f"exit={code} {out[:200]}")

    # ---- check_writing：校验之前就拒（独立工作区：上面的节文件在闸失灵时已被写掉）----
    work_o, _ = _amount_work("w-e75-cw-o")
    res = work_o / "review" / "check-result.json"
    code, out = cw("--work", work_o, "--json", res)
    data = json.loads(res.read_text(encoding="utf-8")) if res.is_file() else {}
    check("E75l check_writing 合法落盘照常（写作进度面板读的那份）",
          code == 0 and "findings" in data and "[结果已落盘]" in out, f"exit={code} {out[:300]}")
    work_m, sec_m = _amount_work("w-e75-cw-m")
    sec_m_sha = sec_m.read_bytes()
    code, out = cw("--work", work_m, "--json", sec_m)
    refused("E75m check_writing --json 落到节文件", code, out, sec_m, sec_m_sha)
    check("E75m2 校验之前就拒了（输出只有一行失败）", out.startswith("[失败]") and "\n" not in out, out[:300])
    work_n, sec_n = _amount_work("w-e75-cw-n")
    sec_n_sha = sec_n.read_bytes()
    code, out = cw("--work", work_n, "--section", sec_n, "--json", sec_n)
    refused("E75n check_writing --section 甲 --json 甲", code, out, sec_n, sec_n_sha)
    work_p, _ = _amount_work("w-e75-cw-p")
    ghost = work_p / "防呆-尚不存在的节.docx"   # 放在工作区根目录：闸失灵时也不会混进节文件的枚举
    code, out = cw("--work", work_p, "--json", ghost)
    check("E75o check_writing --json 指向还不存在的 docx：退出码 2、不创建",
          code == 2 and not ghost.exists(), f"exit={code} {out[:200]}")
    check("E75p 全程节文件未变", sec_m.read_bytes() == sec_m_sha and sec_n.read_bytes() == sec_n_sha)


# ------------------------------------------------ 第 13 批 .doc 通道：没有 LibreOffice 时把转换交给模型


def e76():
    """套件里另外两处 .doc 通道的回退（用户 2026-11-27 拍板「同意」一并改——LibreOffice 体积大、
    不少用户装不了，不能是前提；脱敏那处见 sanitize_smoke S17）：
    · parse_document：没有 LibreOffice 时给回退指引并退出；模型用自选办法转好 .docx 后用
      --converted 交回来——来源登记、解析目录、转换件落点与自动转换完全一样，下游按行号
      拷贝格式件时看到的还是原来的布局；已解析且产物完好时没有转换器也直接沿用；
    · layout-read：只读、无产物依赖，转好后对那份 .docx 重跑同一条命令即等价，只改提示；
    · 三份脚本（脱敏／解析／版式读取）的探测与指引同口径——环节技能各自独立安装不共享
      模块，只能各留一份，这里钉住逐字一致。
    回退部分全程把 LibreOffice 探测摆成「没有」，不依赖本机装没装；只有最后的自动转换
    回归需要它。"""
    import contextlib
    import hashlib
    import types
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    pd_mod = _load("tender-analysis/scripts/parse_document.py")
    sm_mod = _load("tender-sanitize/scripts/strip_metadata.py")

    # 「模型用自选办法转出来的 docx」与它的 .doc 源：接力时解析器只读 docx，.doc 只算指纹
    conv = _direct_doc("回退-转换件")
    fake_doc = conv.with_name("回退-转换件.doc")
    fake_doc.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 64)
    doc_sha, conv_sha = fake_doc.read_bytes(), conv.read_bytes()

    def office():
        return ["Microsoft Word", "WPS Office"]

    def parse(argv):
        """进程内跑 parse_document 的 main()，返回 (退出码, stdout)。"""
        buf = io.StringIO()
        with mock.patch.object(sys, "argv", ["parse_document.py", *[str(a) for a in argv]]), \
                contextlib.redirect_stdout(buf):
            try:
                pd_mod.main()
                code = 0
            except SystemExit as exc:
                code = exc.code if isinstance(exc.code, int) else 1
        return code, buf.getvalue()

    def indented(out):
        lines = [ln for ln in out.splitlines() if ln.strip()]
        return len(lines) >= 3 and all(ln.startswith(" ") for ln in lines[1:]), out[:400]

    def layout_read():
        """进程内跑 layout-read 的命令出口，返回 (退出码, stdout)。"""
        buf = io.StringIO()
        args = types.SimpleNamespace(files=[str(fake_doc)], out=None, json=None)
        with contextlib.redirect_stdout(buf):
            try:
                tk_mod.cmd_layout_read(args)
                code = 0
            except SystemExit as exc:
                code = exc.code
        return code, buf.getvalue()

    # ---- 解析 .doc：没有 LibreOffice，也没人交接力 ----
    out_a = WORK / "e76-parse-a"
    with mock.patch.object(pd_mod, "_find_soffice", lambda: None), \
            mock.patch.object(pd_mod, "_installed_office", office, create=True):
        code, out = parse([fake_doc, "-o", out_a])
    check("E76a 解析 .doc 没有 LibreOffice：退出码 2，点明要先转成 .docx、脚本内置的自动转换只有它",
          code == 2 and "[解析失败]" in out and ".docx" in out and "LibreOffice" in out, f"exit={code} {out[:400]}")
    check("E76a2 回退指引点名检测到的套件，并给出接力的写法（--converted 重跑本命令）",
          "Microsoft Word" in out and "WPS Office" in out and "--converted" in out and "重跑本命令" in out,
          out[:600])
    check("E76a3 不劝人安装 LibreOffice", "请安装" not in out and "装 LibreOffice" not in out, out[:400])
    check("E76a4 原 .doc 未动，没建解析目录", fake_doc.read_bytes() == doc_sha and not out_a.exists())
    check("E76a5 多行指引的续行带缩进", *indented(out))

    # ---- 接力：模型转好 docx，用 --converted 交回来（仍然没有 LibreOffice）----
    out_b, out_c = WORK / "e76-parse-handoff", WORK / "e76-parse-direct"
    dir_b = out_b / fake_doc.name
    stem = conv.stem
    with mock.patch.object(pd_mod, "_find_soffice", lambda: None):
        code, out = parse([fake_doc, "--converted", conv, "-o", out_b])
    mp = dir_b / f"{stem}.meta.json"
    meta = json.loads(mp.read_text(encoding="utf-8")) if mp.is_file() else {}
    check("E76b 解析 .doc --converted <docx> 在没有 LibreOffice 时解析成功",
          code == 0 and "[解析成功]" in out and "Traceback" not in out, f"exit={code} {out[:400]}")
    check("E76b2 来源仍按 .doc 登记（meta.source 与 sha256 是原 .doc 的）",
          meta.get("source") == fake_doc.name
          and meta.get("sha256") == hashlib.sha256(doc_sha).hexdigest(), str(meta)[:300])
    check("E76b3 转换件落到解析目录的 _converted.docx，交来的文件没被动",
          (dir_b / "_converted.docx").is_file() and (dir_b / "_converted.docx").read_bytes() == conv_sha
          and conv.read_bytes() == conv_sha)
    check("E76b4 解析警示写明源为 .doc（交付级内容要核对原件）",
          any("源为 .doc" in w for w in meta.get("warnings", [])), str(meta.get("warnings"))[:300])
    code_c, _ = parse([conv, "-o", out_c])            # 对照：把这份 docx 直接当源解析
    dir_c = out_c / conv.name
    same = code_c == 0 and all(
        (dir_b / n).is_file() and (dir_c / n).is_file()
        and (dir_b / n).read_bytes() == (dir_c / n).read_bytes()
        for n in (f"{stem}.md", f"{stem}.outline.json", f"{stem}.elements.json"))
    check("E76b5 解析产物与「直接解析这份 docx」逐字节一致（md / outline / elements）", same)
    with mock.patch.object(pd_mod, "_find_soffice", lambda: None):
        code, out = parse([fake_doc, "--converted", conv, "-o", out_b])
    check("E76b6 幂等：带 --converted 重跑沿用已有产物",
          code == 0 and "[解析跳过]" in out, f"exit={code} {out[:300]}")
    with mock.patch.object(pd_mod, "_find_soffice", lambda: None):
        code, out = parse([fake_doc, "-o", out_b])    # 不带 --converted，也没有 LibreOffice
    check("E76b7 已解析且产物完好：没有转换器也不带 --converted，照样沿用（不走回退指引）",
          code == 0 and "[解析跳过]" in out and "解析失败" not in out and "交给你选办法" not in out,
          f"exit={code} {out[:300]}")
    md = dir_b / f"{stem}.md"
    if md.is_file():                                 # 前面没走通时不在这里抛异常，让后面的断言照跑
        md.write_text(md.read_text(encoding="utf-8") + "被外部改过的一行\n", encoding="utf-8")
    with mock.patch.object(pd_mod, "_find_soffice", lambda: None):
        code, out = parse([fake_doc, "-o", out_b])
    check("E76b8 产物被外部改过就不沿用，仍给回退指引",
          code == 2 and "--converted" in out and "[解析跳过]" not in out, f"exit={code} {out[:300]}")

    for label, argv in (("源是 .docx 却带 --converted", [conv, "--converted", conv, "-o", out_c]),
                        ("--converted 指向不存在的文件",
                         [fake_doc, "--converted", WORK / "没有这份.docx", "-o", out_c]),
                        ("--converted 指向的不是 .docx", [fake_doc, "--converted", fake_doc, "-o", out_c])):
        code, out = parse(argv)
        check(f"E76c 误用（{label}）：退出码 2，脚本报错点明用法",
              code == 2 and "[解析失败]" in out and "--converted" in out, f"exit={code} {out[:300]}")

    with mock.patch.object(pd_mod, "_find_soffice", lambda: "/x/soffice"), \
            mock.patch.object(pd_mod.subprocess, "run",
                              lambda *a, **k: types.SimpleNamespace(stdout="", stderr="")):
        code, out = parse([fake_doc, "-o", WORK / "e76-parse-fail"])
    check("E76d 有 LibreOffice 但转换没产出文件：退出码非 0，出路指向 --converted 或请用户另存",
          code != 0 and ("--converted" in out or "另存为 .docx" in out) and "请安装" not in out,
          f"exit={code} {out[:300]}")

    # ---- 版式读取 .doc：没有 LibreOffice ----
    with mock.patch.object(tk_mod, "_find_soffice", lambda: None), \
            mock.patch.object(tk_mod, "_installed_office", office, create=True):
        code, out = layout_read()
    check("E76e 版式读取 .doc 没有 LibreOffice：退出码 2，点明要先转成 .docx、脚本内置的自动转换只有它",
          code == 2 and "[失败]" in out and ".docx" in out and "LibreOffice" in out, f"exit={code} {out[:400]}")
    check("E76e2 回退指引点名检测到的套件，说清对那份 .docx 重跑同一条命令",
          "Microsoft Word" in out and "WPS Office" in out and "同一条命令" in out, out[:600])
    check("E76e3 不劝人安装 LibreOffice，原 .doc 未动",
          "请安装" not in out and "装 LibreOffice" not in out and fake_doc.read_bytes() == doc_sha, out[:400])
    check("E76e4 多行指引的续行带缩进", *indented(out))
    with mock.patch.object(tk_mod, "_find_soffice", lambda: "/x/soffice"), \
            mock.patch.object(subprocess, "run",   # docx_toolkit 在函数内局部 import subprocess
                              lambda *a, **k: types.SimpleNamespace(stdout="", stderr="")):
        code, out = layout_read()
    check("E76f 版式读取有 LibreOffice 但转换没成：提示请用户自己另存为 .docx（不请模型另想办法）",
          code != 0 and "另存为 .docx" in out and "请用户" in out, f"exit={code} {out[:300]}")

    # ---- 三份脚本同口径（各自独立安装不共享模块，只能各留一份）----
    hits = {"/Applications/wpsoffice.app", "/Applications/Microsoft Word.app"}
    real_isdir = os.path.isdir
    with mock.patch.object(sys, "platform", "darwin"), \
            mock.patch.object(os.path, "isdir", lambda p: p in hits or real_isdir(p)):
        got = [m._installed_office() for m in (tk_mod, pd_mod, sm_mod)]
    check("E76g 三份脚本的 _installed_office 同口径（同一台机器探出同样的套件）",
          got[0] == got[1] == got[2] == ["Microsoft Word", "WPS Office"], str(got))
    hints = [m._hand_over_hint("X 只认 .docx，得先把它转成 .docx", "转好后做 Y。")
             for m in (tk_mod, pd_mod, sm_mod)]
    check("E76g2 三份脚本的回退指引同口径（同一入参逐字相同，场景只差在入参）",
          hints[0] == hints[1] == hints[2] and "X 只认 .docx" in hints[0] and "转好后做 Y。" in hints[0]
          and "请安装" not in hints[0], str(hints)[:500])

    # ---- 回归：有 LibreOffice 时解析照常自动转换（本机缺它按口径记失败，不跳过）----
    soffice = pd_mod._find_soffice()
    ok_auto, detail = False, "本机没有 LibreOffice"
    if soffice:
        real = _direct_doc("回退-自动源")
        real_doc = WORK / "回退-自动.doc"
        with tk_mod._tempdir() as td:
            subprocess.run([soffice, "--headless", f"-env:UserInstallation={Path(td).resolve().as_uri()}",
                            "--convert-to", "doc", "--outdir", td, str(real)],
                           capture_output=True, text=True, errors="replace", timeout=300)
            if (Path(td) / f"{real.stem}.doc").is_file():
                shutil.copy2(Path(td) / f"{real.stem}.doc", real_doc)
        if real_doc.is_file():
            code, out = parse([real_doc, "-o", WORK / "e76-parse-auto"])
            mp = WORK / "e76-parse-auto" / real_doc.name / f"{real_doc.stem}.meta.json"
            m2 = json.loads(mp.read_text(encoding="utf-8")) if mp.is_file() else {}
            ok_auto = (code == 0 and "[解析成功]" in out and (mp.parent / "_converted.docx").is_file()
                       and m2.get("source") == real_doc.name
                       and any("LibreOffice 转换" in w for w in m2.get("warnings", [])))
            detail = f"exit={code} {out[:300]}"
    check("E76h 有 LibreOffice 时解析照常自动转换（回归）", ok_auto, detail)


# ------------------------------------------------ 第 14 批 版式读取遇到不规整的真实文档
# （2026-11-27 真机：读资格文件、报价文件时整份崩；模型自己转出来的 docx 又崩在页面上。
#   每一处都先在真实文件上复现，再在这里写成合成夹具）

def _shallow_heading_doc(name) -> Path:
    """最浅只用到「标题 2」的小册子（真机上资格文件、报价文件的形态）：章挂在标题 2 上，
    全文没有一级标题。"""
    from docx.enum.text import WD_BREAK
    d = _doc()
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)        # 封面后的分页
    for text, lvl in (("第一章 投标函", 2), ("一、投标函正文", 3)):
        h = d.add_heading(text, level=lvl)
        for r in h.runs:
            _run_fmt(r, "黑体", 16 if lvl == 2 else 12, True)
        d.add_paragraph(_LSENT)
    h = d.add_heading("第二章 开标一览表", 2)                   # 第二个最浅级标题：段前分页
    for r in h.runs:
        _run_fmt(r, "黑体", 16, True)
    h.paragraph_format.page_break_before = True
    d.add_paragraph(_LSENT)
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _no_page_setup_doc(name, drop_sect=False) -> Path:
    """节里不写纸张与页边距（别的转换器转出来的 docx 常这样整节空着）；drop_sect 连节属性
    本身都不留。"""
    d = _doc()
    d.add_paragraph(_LSENT)
    sp = d.sections[0]._sectPr
    if drop_sect:
        sp.getparent().remove(sp)
    else:
        for tag in ("w:pgSz", "w:pgMar"):
            el = sp.find(qn(tag))
            if el is not None:
                sp.remove(el)
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def e77():
    """版式读取遇到不规整的真实文档：没有一级标题、节里没写页面设置、某一组读不出来——
    都不崩；读得出的照读，读不出的照实标出来。"""
    # ---- 没有一级标题：原先拿空列表的首元素，IndexError 整份读不出 ----
    p = _shallow_heading_doc("无一级标题")
    code, out = tk("layout-read", p)
    row = _row(out, "章起新页")
    check("E77a 没有一级标题不崩：退出码 0，章起新页按最浅的一级照读、依据写明并记疑似",
          code == 0 and _val(out, "章起新页") == "是" and "没识别出一级标题" in row
          and "标题二级 2 处" in row and "其余 1 处里 1 处起新页" in row and "疑似" in row,
          f"exit={code} {row or out[:300]}")
    check("E77a2 其余各组照读，表里没有「读取中断」",
          "正文字体" in out and "页边距" in out and "读取中断" not in out and "本组未读" not in out,
          out[:500])

    # ---- 节里没写纸张与页边距：原先取属性取到 None，AttributeError 整份读不出 ----
    for label, path, no_sect in (("节里没写页面设置", _no_page_setup_doc("没写页面设置"), False),
                                 ("一个节属性都没有", _no_page_setup_doc("没有节属性", True), True)):
        code, out = tk("layout-read", path)
        ok = (code == 0 and _val(out, "纸张") == "Letter 纵向"
              and _val(out, "页边距") == "上 2.54、下 2.54、左 3.17、右 3.17 厘米"
              and "取的是 Word 缺省值" in _row(out, "纸张") and "取的是 Word 缺省值" in _row(out, "页边距")
              and "疑似" in _row(out, "纸张") and "疑似" in _row(out, "页边距")
              and "读取中断" not in out and "本组未读" not in out and "Traceback" not in out
              and _val(out, "页眉") == "无" and _val(out, "封面页设置") == "同正文")
        check(f"E77b {label}不崩：退出码 0，纸张 Letter、页边距按 Word 缺省值读出、注明并记疑似",
              ok, f"exit={code} {out[:500]}")

    # ---- 回归：带单位的纸张写法照旧走 python-docx 的换算，没被缺省值顶掉 ----
    d = _doc()
    d.add_paragraph(_LSENT)
    sz = d.sections[0]._sectPr.find(qn("w:pgSz"))
    sz.set(qn("w:w"), "21cm")
    sz.set(qn("w:h"), "29.7cm")
    p = WORK / "带单位纸张.docx"
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E77c 带单位的纸张写法（21cm）照读成 A4、记确定，不落到缺省值",
          code == 0 and _val(out, "纸张") == "A4 纵向" and "确定" in _row(out, "纸张")
          and "缺省值" not in _row(out, "纸张"), f"exit={code} {_row(out, '纸张')}")

    # ---- 某一组读不出来不拖垮整份 ----
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    tk_mod.require_docx()
    with mock.patch.object(tk_mod, "_layout_headings", side_effect=RuntimeError("标题组坏了")):
        res = tk_mod.layout_read_doc(Document(str(_company_doc("读取中断"))))
    txt = tk_mod.format_layout(res, "公司标书")
    seg = txt[txt.index("【标题】"):txt.index("【表格】")]
    broken = [r for r in res["rows"] if r[3] == "读取时出错"]
    check("E77d 标题组读不出来：标题表里只剩一行「本组读取中断」写明异常，其余各组照读",
          broken == [("标题", "本组读取中断", "RuntimeError: 标题组坏了", "读取时出错", "疑似")]
          and len([ln for ln in seg.splitlines() if ln.startswith("| ")]) == 2
          and _val(txt, "正文字体") == "中文仿宋、西文 Times New Roman"
          and _val(txt, "纸张") == "A4 纵向" and _val(txt, "目录层级") == "3 级"
          and _val(txt, "表题") and _val(txt, "封面版式"), txt[:1500])
    check("E77d2 首行先点明有没能读取的项（模型读回显的第一眼）",
          "其中 1 项没能读取" in txt.splitlines()[0] and "只读" in txt.splitlines()[0], txt.splitlines()[0])

    with mock.patch.object(tk_mod, "_layout_page", side_effect=RuntimeError("页面组坏了")):
        res = tk_mod.layout_read_doc(Document(str(_company_doc("上游没读出"))))
    txt = tk_mod.format_layout(res, "公司标书")
    rows = res["rows"]
    skipped = sorted(r[:2] for r in rows if r[1] == "本组未读")
    check("E77d3 上游没读出来的组不白跑：页眉页脚／图／封面各记一行「本组未读」，不依赖它的组照读",
          skipped == sorted([("页面", "本组未读"), ("图", "本组未读"), ("封面", "本组未读")])
          and any(r[:2] == ("页面", "本组读取中断") and "页面组坏了" in r[2] for r in rows)
          and _val(txt, "正文字体") == "中文仿宋、西文 Times New Roman"
          and _val(txt, "标题层级") and _val(txt, "表格边框")
          and "其中 4 项没能读取" in txt.splitlines()[0] and "页面（没读出来）" in txt,
          txt[:1800])

    raised = False
    with mock.patch.object(tk_mod, "_layout_headings", side_effect=RuntimeError("标题组坏了")), \
            mock.patch.dict(os.environ, {"TENDER_TRACEBACK": "1"}):
        try:
            tk_mod.layout_read_doc(Document(str(_company_doc("调试栈"))))
        except RuntimeError:
            raised = True
    check("E77d4 TENDER_TRACEBACK=1 时照旧抛出（调试与真实样本回放要的是完整栈，不被吞掉）", raised)

    # ---- 阴性对照：正常文档没有任何「读取中断」 ----
    code, out = tk("layout-read", _company_doc("读取阴性对照"))
    check("E77e 阴性对照：正常文档没有「读取中断／本组未读／[注意]」，首行不提没能读取",
          code == 0 and out.startswith("[版式读取]") and "读取中断" not in out
          and "本组未读" not in out and "没能读取" not in out and "[注意]" not in out, out[:300])

    # ---- 把文字抽出来重新拼成的 docx（真机上模型自己「转」出来的那份）：包里连样式部件都没有，要点破 ----
    # （夹具必须是「没有部件」而不是「空样式表」：读段落样式时 python-docx 会就地补一份默认样式表，
    #   判断的时机晚一步就数不出来了——真实文件回放抓到的。stylesWithEffects 那份副本也要摘掉）
    d = _doc()
    d.add_paragraph(_LSENT)
    for rid, rel in list(d.part.rels.items()):
        if rel.reltype.rsplit("/", 1)[-1].lower().startswith("styles"):
            d.part.rels.pop(rid, None)
    p = WORK / "拼出来的.docx"
    d.save(str(p))
    check("E77f0 夹具自身：包里确实没有任何样式部件（否则下面的断言测不到东西）",
          not [n for n in zipfile.ZipFile(p).namelist() if n.startswith("word/styles")])
    code, out = tk("layout-read", p)
    check("E77f 包里一个样式都没有的 docx：不崩，紧跟首行有一行 [注意] 点破「不是 Office 另存出来的、不能当版式样本」",
          code == 0 and "[注意]" in out and "一个样式都没有" in out and "不能当版式样本" in out
          and "另存为 .docx" in out and out.index("[注意]") < out.index("【页面】")
          and out.index("[注意]") > out.index("[版式读取]"), out[:700])
    spec = WORK / "拼出来的-规格.md"
    code, out = tk("layout-read", p, "--out", spec)
    check("E77f2 --out 落盘的规格表开头带着同一条注意（不能留下一张看着正常的规格表）",
          code == 0 and "注意：" in spec.read_text(encoding="utf-8").split("【页面】")[0],
          spec.read_text(encoding="utf-8")[:400] if spec.is_file() else f"exit={code} {out[:300]}")


# ------------------------------------------------ 第 15 批 用户模板缺标题样式
# （2026-11-27 真机：模型照 SKILL 写的「缺则回落内置样式并点名」向用户转述；实测只补了个空样式、
#   不回显，合出来的章标题跟着正文走、和正文一个样）

def _ea(font_owner) -> str | None:
    rpr = font_owner.element.rPr if hasattr(font_owner, "element") else font_owner._element.rPr
    return rpr.rFonts.get(qn("w:eastAsia")) if rpr is not None and rpr.rFonts is not None else None


def _drop_heading_styles(d) -> None:
    """把 Heading 1-4 从样式表里摘掉——真实历史标书的常见形态：字体全直写，样式表里没有标题样式。"""
    for name in ("Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        d.styles.element.remove(d.styles.element.find(f"{qn('w:style')}[@{qn('w:styleId')}='{d.styles[name].style_id}']"))


def e78():
    """用户模板没有 Heading 1-4 样式时：按内置样式补建并在回显里点名；模板自带的标题样式不被覆盖。"""
    tpl = WORK / "无标题样式模板.docx"
    d = _doc()
    _drop_heading_styles(d)
    d.styles["Normal"].font.name = "仿宋"
    d.styles["Normal"].font.size = Pt(10.5)          # 模板的正文设置，认得出标题有没有跟着它走
    d.save(str(tpl))
    work, _ = make_work("w-tpl-missing", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n  - 乙节\n",
                        ["甲章", "乙节"], with_body=True)
    code, out = tk("assemble", "--work", work, "--template", tpl)
    note = next((ln for ln in out.splitlines() if ln.startswith("[模板标题样式]")), "")
    check("E78a 合册用缺标题样式的模板：退出码 0，回显 [模板标题样式] 点名缺了哪些、已按内置样式补建",
          code == 0 and "Heading 1" in note and "Heading 2" in note and "补建" in note and tpl.name in note,
          f"exit={code} {note or out[:500]}")
    book = Document(str(next((work / "body").glob("整本-*.docx"))))
    pt = lambda st: st.font.size.pt if st.font.size is not None else None   # noqa: E731  空样式没有字号
    h1_xml = book.styles["Heading 1"].element.xml
    check("E78b 补建的标题样式带字体字号（黑体，一级 16 磅、二级 14 磅，加粗），章标题不跟正文的仿宋 10.5 磅走",
          pt(book.styles["Heading 1"]) == 16 and pt(book.styles["Heading 2"]) == 14
          and book.styles["Heading 1"].font.bold and _ea(book.styles["Heading 1"]) == "黑体"
          and "Theme=" not in h1_xml            # 显式字体，没有能被模板主题顶掉的主题字体属性
          and pt(book.styles["Normal"]) == 10.5,
          f"H1={pt(book.styles['Heading 1'])} {_ea(book.styles['Heading 1'])!r} Normal={pt(book.styles['Normal'])}")
    code, out = tk("create", work / "body" / "新节.docx", "--title", "新节", "--template", tpl)
    check("E78c 建节走同一条路：同样回显点名", code == 0 and "[模板标题样式]" in out and "Heading 1" in out,
          f"exit={code} {out[:300]}")

    # ---- 阴性对照：模板自带标题样式时不回显、不覆盖 ----
    tpl2 = WORK / "自带标题样式模板.docx"
    d = _doc()
    d.styles["Heading 1"].font.size = Pt(18)          # 公司自己的标题样式
    d.save(str(tpl2))
    work2, _ = make_work("w-tpl-own", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲章\n", ["甲章"], with_body=True)
    code, out = tk("assemble", "--work", work2, "--template", tpl2)
    book2 = Document(str(next((work2 / "body").glob("整本-*.docx"))))
    check("E78d 阴性对照：模板自带标题样式时不回显、不覆盖（公司自己的 18 磅照旧）",
          code == 0 and "[模板标题样式]" not in out and pt(book2.styles["Heading 1"]) == 18,
          f"exit={code} {out[:300]}")


# ------------------------------------------------ 第 16 批 超过 5MB 的 .doc 不自动转换
# （用户 2026-11-27 定：超过 5M 的提醒用户自己转换。起因：WorkBuddy 真机上一份 26.7MB 的 .doc
#   让 LibreOffice 转满 300 秒超时，模型随后把文字抽出来自己拼了一份丢光版式的 docx 交回来。
#   5MB 是用户定的值，不是标定值——本机实测 9 份真实 .doc，5.4–23.6MB 的都是 2–3 秒转完，
#   只有那一份 26.7MB 的转不出来，卡住的是内容不是体积）

def _fake_doc(name, size: int) -> Path:
    """指定大小的假 .doc（文件头对、后面是零；稀疏文件，不占盘）。"""
    p = WORK / name
    with open(p, "wb") as fh:
        fh.write(b"\xd0\xcf\x11\xe0")
        fh.truncate(size)
    return p


def e79():
    """超过 5MB 的 .doc：三处通道都不自动转换、不请模型另想办法，直接请用户在 Word/WPS 里
    另存；解析带 --converted 或已有完好的转换件时照常走。"""
    import contextlib
    import types
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    pd_mod = _load("tender-analysis/scripts/parse_document.py")
    sm_mod = _load("tender-sanitize/scripts/strip_metadata.py")
    limit = getattr(tk_mod, "_DOC_AUTO_LIMIT", 5_000_000)
    big = _fake_doc("超大.doc", 26_739_022)            # 真机那份的大小
    edge = _fake_doc("压线.doc", limit)
    over = _fake_doc("刚超线.doc", limit + 1)
    head = big.read_bytes()[:4096]
    ran = []

    def fake_run(*a, **k):
        ran.append(1)
        return types.SimpleNamespace(stdout="", stderr="")

    def parse(argv):
        """进程内跑 parse_document 的 main()，返回 (退出码, stdout)。"""
        buf = io.StringIO()
        with mock.patch.object(sys, "argv", ["parse_document.py", *[str(a) for a in argv]]), \
                contextlib.redirect_stdout(buf):
            try:
                pd_mod.main()
                code = 0
            except SystemExit as exc:
                code = exc.code if isinstance(exc.code, int) else 1
        return code, buf.getvalue()

    def layout_read(path):
        buf = io.StringIO()
        args = types.SimpleNamespace(files=[str(path)], out=None, json=None)
        with contextlib.redirect_stdout(buf):
            try:
                tk_mod.cmd_layout_read(args)
                code = 0
            except SystemExit as exc:
                code = exc.code
        return code, buf.getvalue()

    # ---- 版式读取：有 LibreOffice 也不转 ----
    with mock.patch.object(tk_mod, "_find_soffice", lambda: "/opt/libreoffice/soffice"), \
            mock.patch.object(subprocess, "run", fake_run):
        code, out = layout_read(big)
    check("E79a 版式读取超过 5MB 的 .doc：退出码 2，写明大小与 5MB 这条线，请用户自己另存，不启动转换",
          code == 2 and "26.7MB" in out and "超过 5MB" in out and "请用户" in out and "另存为 .docx" in out
          and "重跑同一条命令" in out and not ran, f"exit={code} 转换调用={len(ran)} {out[:400]}")
    check("E79a2 不请模型另想办法转（那正是真机上拼出坏文件的路），原 .doc 未动",
          "选办法" not in out and big.stat().st_size == 26_739_022 and big.read_bytes()[:4096] == head,
          out[:400])

    # ---- 压线：正好 5MB 不算超（「超过」是严格大于）；多一个字节就算，且排在「没有 LibreOffice」前面 ----
    with mock.patch.object(tk_mod, "_find_soffice", lambda: None):
        code, out = layout_read(edge)
        code2, out2 = layout_read(over)
    check("E79b 正好 5MB 不算超：走的还是「没有 LibreOffice」的回退指引；多一个字节就按超限处理",
          code == 2 and "没有找到 LibreOffice" in out and "超过 5MB" not in out
          and code2 == 2 and "超过 5MB" in out2 and "没有找到 LibreOffice" not in out2,
          f"{out[:200]} ｜ {out2[:200]}")

    # ---- 解析：不建解析目录、指路 --converted；带上用户另存的转换件照常解析 ----
    out_a = WORK / "e79-parse"
    ran.clear()
    with mock.patch.object(pd_mod, "_find_soffice", lambda: "/opt/libreoffice/soffice"), \
            mock.patch.object(subprocess, "run", fake_run):
        code, out = parse([big, "-o", out_a])
    check("E79c 解析超过 5MB 的 .doc：退出码 2，不启动转换、不建解析目录，指路 --converted",
          code == 2 and "超过 5MB" in out and "--converted" in out and "请用户" in out
          and not ran and not out_a.exists(), f"exit={code} 转换调用={len(ran)} {out[:400]}")
    conv = _direct_doc("超限-转换件")                  # 用户自己在 Word/WPS 里另存出来的 docx
    out_b = WORK / "e79-parse-handoff"
    code, out = parse([big, "-o", out_b, "--converted", conv])
    check("E79d 带上用户自己另存的 --converted：照常解析（大小线只拦自动转换）",
          code == 0 and "[解析成功]" in out and (out_b / big.name / "_converted.docx").is_file(),
          f"exit={code} {out[:300]}")

    # ---- 已有完好的转换件：超过 5MB 也直接沿用 ----
    def must_not_convert(*a, **k):
        raise AssertionError("已有完好的转换件时不应再转换")

    try:
        with mock.patch.object(pd_mod, "_doc_to_docx", must_not_convert), \
                mock.patch.object(pd_mod, "_find_soffice", lambda: "/opt/libreoffice/soffice"):
            code, out = parse([big, "-o", out_b])
    except AssertionError as exc:
        code, out = -1, str(exc)
    check("E79e 已有完好的转换件：超过 5MB 也直接沿用，不转也不拦",
          code == 0 and "解析跳过" in out and "沿用已有产物" in out, f"exit={code} {out[:300]}")

    # ---- 三份脚本同口径（各自独立安装不共享模块，只能各留一份）----
    hs = [getattr(m, "_oversize_doc_hint", None) for m in (tk_mod, pd_mod, sm_mod)]
    limits = [getattr(m, "_DOC_AUTO_LIMIT", None) for m in (tk_mod, pd_mod, sm_mod)]
    texts = [h(26_739_022, "转好后做 Y。") if callable(h) else "" for h in hs]
    check("E79f 三份脚本的大文件提示同口径（同一入参逐字相同），大小线一致、都是 5MB",
          texts[0] == texts[1] == texts[2] and "26.7MB" in texts[0] and "超过 5MB" in texts[0]
          and "请用户" in texts[0] and "转好后做 Y。" in texts[0] and limits == [5_000_000] * 3,
          f"{limits} {str(texts)[:300]}")


# ------------------------------------------------ 第 17 批 LibreOffice 转换超时：杀干净，再提醒用户
# （用户 2026-11-27 定：转换超时就 kill 掉，然后提醒用户。起因是真机上一份 .doc 让 LibreOffice 的
#   排版陷进死循环——进程占满一个核、永不结束。改前两个问题：①subprocess.run 超时只杀直接子进程，
#   soffice 启动器在 Linux／Windows 上只是外壳，真正干活的 soffice.bin 是它拉起来的子进程，会留下
#   继续空转；②提示语说的是「换一种办法试试」，等于请模型自己想办法，而它交回来的是把文字抽出来
#   重新拼的、丢光版式的 docx）

def _hanging_soffice(name) -> tuple[Path, Path]:
    """假的 soffice：自己再拉起一个子进程，然后父子俩都不结束；把两个 pid 写进文件。
    用来验证「超时之后整棵进程树都被杀掉」。返回 (假 soffice 路径, pid 文件路径)。"""
    pids, script = WORK / f"{name}.pids", WORK / f"{name}.py"
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
    """等这些 pid 都消失；返回到点还活着的（调用方负责收尾杀掉，别留给后面的用例）。"""
    end = time.time() + seconds
    while True:
        alive = [p for p in pids if _alive(p)]
        if not alive or time.time() > end:
            return alive
        time.sleep(0.2)


def _reap(pids: list) -> None:
    for p in pids:
        try:
            os.kill(p, 9)
        except OSError:
            pass


def e80():
    """LibreOffice 转换超时：整棵进程树杀干净、不留空转的 soffice；提示请用户自己另存，不请模型
    另想办法。三处 .doc 通道与目录页码实测共用同一个「起 LibreOffice 并等它结束」的底座。"""
    import contextlib
    import types
    tk_mod = _load("tender-docx/scripts/docx_toolkit.py")
    pd_mod = _load("tender-analysis/scripts/parse_document.py")
    sm_mod = _load("tender-sanitize/scripts/strip_metadata.py")
    real_run = subprocess.run

    def clamp(*a, **k):   # 改前的代码把 300 秒写死在 subprocess.run 的入参里：夹到 2 秒，红的那一步才跑得完
        return real_run(*a, **{**k, "timeout": min(k.get("timeout", 2), 2)})

    # ---- 底座：三份脚本各自的「起 LibreOffice 并等它结束」，超时后子进程树一个不留 ----
    for tag, mod in (("docx_toolkit", tk_mod), ("parse_document", pd_mod), ("strip_metadata", sm_mod)):
        runner, pids = _hanging_soffice(f"hang-{tag}")
        kind, left, ran = "没抛异常", [], False
        try:
            with mock.patch.object(mod, "_LO_TIMEOUT", 2, create=True):
                mod._run_lo([str(runner)], 2)
        except AttributeError as exc:
            kind = f"AttributeError: {str(exc)[:60]}"
        except Exception as exc:  # noqa: BLE001
            kind = type(exc).__name__
        finally:
            ran = pids.is_file()
            left = _wait_dead([int(x) for x in pids.read_text().split()]) if ran else []
            _reap(left)
        check(f"E80a {tag}：转换超时后抛 _LoTimeout，父子进程都杀掉、不留空转的进程",
              kind == "_LoTimeout" and ran and not left, f"抛出={kind}｜假 soffice 起没起={ran}｜残留 {left}")

    # ---- 三份脚本的提示与时限逐字同口径 ----
    hs = [getattr(m, "_stuck_doc_hint", None) for m in (tk_mod, pd_mod, sm_mod)]
    texts = [h(" 超过 300 秒没转完", "转好后做 Y。") if callable(h) else "" for h in hs]
    check("E80b 三份脚本的「转不动」提示逐字同口径：提醒用户自己另存，不请模型另想办法",
          texts[0] == texts[1] == texts[2] and "请用户" in texts[0] and "另存为 .docx" in texts[0]
          and "不要替用户自动化" in texts[0] and "不要另想办法转" in texts[0] and "已终止" in texts[0]
          and "转好后做 Y。" in texts[0] and "换一种" not in texts[0] and "选办法" not in texts[0],
          str(texts)[:400])
    limits = [getattr(m, "_LO_TIMEOUT", None) for m in (tk_mod, pd_mod, sm_mod)]
    check("E80b2 三份脚本的时限一致、都是 300 秒（经验值：本机 10 份能转的真实 .doc 最慢 38 秒）",
          limits == [300, 300, 300], str(limits))

    fake_doc = _fake_doc("转不动的小文件.doc", 4_000_000)          # 5MB 以内，不会被大小线拦下
    head = fake_doc.read_bytes()[:4096]

    def parse(argv):
        """进程内跑 parse_document 的 main()，返回 (退出码, stdout)。"""
        buf = io.StringIO()
        with mock.patch.object(sys, "argv", ["parse_document.py", *[str(a) for a in argv]]), \
                contextlib.redirect_stdout(buf):
            try:
                pd_mod.main()
                code = 0
            except SystemExit as exc:
                code = exc.code if isinstance(exc.code, int) else 1
        return code, buf.getvalue()

    def layout_read(path):
        buf = io.StringIO()
        args = types.SimpleNamespace(files=[str(path)], out=None, json=None)
        with contextlib.redirect_stdout(buf):
            try:
                tk_mod.cmd_layout_read(args)
                code = 0
            except SystemExit as exc:
                code = exc.code
        return code, buf.getvalue()

    # ---- 版式读取：真起一个卡死的假 soffice ----
    runner, pids = _hanging_soffice("hang-layout")
    t0 = time.time()
    left = []
    try:
        with mock.patch.object(tk_mod, "_find_soffice", lambda: str(runner)), \
                mock.patch.object(tk_mod, "_LO_TIMEOUT", 2, create=True), \
                mock.patch.object(subprocess, "run", clamp):
            code, out = layout_read(fake_doc)
    finally:
        if pids.is_file():                      # 断言之前先收尾：改前的代码会留下孤儿子进程
            left = _wait_dead([int(x) for x in pids.read_text().split()])
            _reap(left)
    took = time.time() - t0
    check("E80c 版式读取遇到卡死的转换：时限一到就返回（退出码 2），提示请用户另存、不再说「换一种办法」",
          code == 2 and "超时" in out and "已终止" in out and "请用户" in out and "另存为 .docx" in out
          and "重跑同一条命令" in out and "换一种" not in out and took < 20, f"exit={code} 用时 {took:.0f}s {out[:400]}")
    check("E80c2 转换进程杀干净（父子都不留），原 .doc 未动",
          pids.is_file() and not left and fake_doc.read_bytes()[:4096] == head, f"残留 {left}")
    check("E80c3 多行提示的续行带缩进", all(ln.startswith("        ") for ln in out.splitlines()[1:] if ln.strip()), out[:400])

    # ---- 解析：同一条路，出口指向 --converted ----
    runner, pids = _hanging_soffice("hang-parse")
    out_a = WORK / "e80-parse"
    left = []
    try:
        with mock.patch.object(pd_mod, "_find_soffice", lambda: str(runner)), \
                mock.patch.object(pd_mod, "_LO_TIMEOUT", 2, create=True), \
                mock.patch.object(subprocess, "run", clamp):
            code, out = parse([fake_doc, "-o", out_a])
    finally:
        if pids.is_file():
            left = _wait_dead([int(x) for x in pids.read_text().split()])
            _reap(left)
    check("E80d 解析遇到卡死的转换：退出码 2，提示请用户另存后带 --converted 重跑，不产出转换件",
          code == 2 and "[解析失败]" in out and "已终止" in out and "请用户" in out and "--converted" in out
          and "换一种" not in out and not (out_a / fake_doc.name / "_converted.docx").exists()
          and pids.is_file() and not left, f"exit={code} 残留 {left} {out[:400]}")

    # ---- 转换失败（进程退了但没产出文件）也走同一个出口 ----
    def no_output(*a, **k):
        return subprocess.CompletedProcess(a[0], 1, "Error: source file could not be loaded", "")
    with mock.patch.object(tk_mod, "_find_soffice", lambda: "/x/soffice"), \
            mock.patch.object(tk_mod, "_run_lo", no_output, create=True), \
            mock.patch.object(subprocess, "run", lambda *a, **k: types.SimpleNamespace(stdout="Error: x", stderr="")):
        code, out = layout_read(fake_doc)
    check("E80e 版式读取转换失败（没产出文件）：同样是请用户另存，并写明失败原因",
          code == 2 and "失败" in out and "source file could not be loaded" in out and "请用户" in out
          and "换一种" not in out, f"exit={code} {out[:400]}")
    with mock.patch.object(pd_mod, "_find_soffice", lambda: "/x/soffice"), \
            mock.patch.object(pd_mod, "_run_lo", no_output, create=True), \
            mock.patch.object(subprocess, "run", lambda *a, **k: None):
        code, out = parse([fake_doc, "-o", WORK / "e80-parse-fail"])
    check("E80e2 解析转换失败（没产出文件）：同样是请用户另存后带 --converted 重跑",
          code == 2 and "失败" in out and "请用户" in out and "--converted" in out and "换一种" not in out,
          f"exit={code} {out[:400]}")

    # ---- 目录页码实测那条路（docx → pdf）超时：文案照旧（没有等价替代），不再夹带 .doc 的提示 ----
    buf = io.StringIO()
    boom = lambda *a, **k: (_ for _ in ()).throw(tk_mod._LoTimeout(300))   # noqa: E731
    try:
        with mock.patch.object(tk_mod, "_run_lo", boom, create=True), contextlib.redirect_stdout(buf):
            tk_mod._lo_convert("/x/soffice", WORK / "整本.docx", WORK, WORK / "p")
        code = 0
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
    except Exception as exc:  # noqa: BLE001
        code = f"异常 {type(exc).__name__}"
    txt = buf.getvalue()
    check("E80f 目录页码实测转换超时：文案照旧（可跳过本步稍后再试），不夹带 .doc 的「请用户另存」",
          code != 0 and "超时" in txt and "可跳过本步稍后再试" in txt and "请用户" not in txt, f"code={code} {txt[:300]}")

    # ---- Windows 分支（只在 mac 上模拟）：按进程树强杀；taskkill 用不了也不抛 ----
    calls, proc, errs = [], types.SimpleNamespace(pid=4321, killed=0), []
    proc.kill = lambda: setattr(proc, "killed", proc.killed + 1)
    for label, fake_run in (("有 taskkill", lambda *a, **k: calls.append(list(a[0])) or types.SimpleNamespace(returncode=0)),
                            ("没有 taskkill", lambda *a, **k: (_ for _ in ()).throw(OSError("没有 taskkill")))):
        try:
            with mock.patch.object(sys, "platform", "win32"), mock.patch.object(subprocess, "run", fake_run):
                tk_mod._kill_tree(proc)
        except Exception as exc:  # noqa: BLE001
            errs.append(f"{label}: {type(exc).__name__}")
    check("E80g Windows 下按进程树强杀（taskkill /PID <pid> /T /F），两种情况都不抛，最后都补一刀杀 proc 自己",
          not errs and calls == [["taskkill", "/PID", "4321", "/T", "/F"]] and proc.killed == 2,
          f"{errs} {calls} killed={proc.killed}")


# ---------------------------------------------------------------- 第 18 批 版式读取复盘（WorkBuddy 真机回放，2026-11-27）
#
# 读数本身对（独立探针对真实文件逐项核过），钉的是脚本自己的说法。19 份真实文档改前回放：页面差异
# 被截断 2 份、页码依据节号写错 2 份、「全是直接格式」命中 2 份（其中一份一段直接格式都没有）。


def _fact(out):
    return next((ln for ln in out.splitlines() if ln.startswith("事实：")), "")


def _basis(out, item):
    row = _row(out, item)
    return row.split("|")[3].strip() if row else ""


def e81():
    """事实行：命名标题样式 0 段时，外观从哪来说哪。「全是直接格式」是写死的断言——真实文件里标题
    一段直接格式都没有（全在正文默认样式里），模型读到这句又查到文字上没有格式，自相矛盾地查了三轮。"""
    p = _plain_doc("纯文字标书")
    code, out = tk("layout-read", p)
    fact = _fact(out)
    check("E81a 标题没挂样式也没写直接格式：事实行说外观取自正文默认样式，不说「全是直接格式」",
          code == 0 and "命名标题样式 0/" in fact and "正文默认样式" in fact and "直接格式" not in fact, fact)
    check("E81b 读数照旧：标题与正文同为宋体小四、不加粗",
          _val(out, "标题一级") == "中文宋体、西文 Times New Roman｜小四（12 磅）｜左"
          and _val(out, "标题二级") == "中文宋体、西文 Times New Roman｜小四（12 磅）｜左",
          _row(out, "标题一级") + _row(out, "标题二级"))
    d = Document(str(p))
    h = next(q for q in d.paragraphs if q.text.startswith("第七章"))
    _run_fmt(h.runs[0], "黑体", 16, True)
    d.save(str(p))
    code, out = tk("layout-read", p)
    fact = _fact(out)
    check("E81c 个别标题写了直接格式：点名几段，不笼统说「全是」",
          "命名标题样式 0/" in fact and "直接格式 1 段" in fact and "全是直接格式" not in fact, fact)
    code, out = tk("layout-read", _direct_doc("纯文字对照"))
    check("E81d 阴性对照——外观真全写在文字上的（E57 那份）：仍说「标题全是直接格式」",
          "全是直接格式" in _fact(out), _fact(out))


def e82():
    """页面设置差异：设置相同的节归成一组；一行装不下时按节收、写明另有几节——不逐节罗列到 400 字
    拦腰截成「第 2…」（真实文件 119 节里 59 个图页节设置完全相同）。"""
    p = _plain_doc("多节标书")
    code, out = tk("layout-read", p)
    v = _val(out, "页面设置差异")
    check("E82a 设置相同的节归成一组：第 2、4、6 节等 12 节，整行不带省略号",
          code == 0 and "第 2、4、6 节等 12 节" in v and "纸张 A4 横向" in v and "…" not in v
          and "第 8 节" not in v, v)
    check("E82b 纸张依据写明是正文段最多的那一节（第 13 节），不被读成全文节数",
          "正文段最多" in _basis(out, "纸张") and "第 13 节" in _basis(out, "纸张"), _row(out, "纸张"))
    spec = WORK / "多节-规格.md"
    code, out = tk("layout-read", p, "--out", spec)
    row = next((ln for ln in spec.read_text(encoding="utf-8").splitlines()
                if ln.startswith("| 页面设置差异 |")), "") if spec.is_file() else ""
    check("E82c 落盘的规格表同一行同样归并、同样不截断", "第 2、4、6 节等 12 节" in row and "…" not in row, row)
    # 各节设置都不相同、一行装不下：按段收，写明另有几段；行里存的是完整取值，--json 结果里拿得到全量
    p2 = _ragged_secs_doc("节节不同")
    jp = WORK / "节节不同.json"
    code, out = tk("layout-read", p2, "--json", jp)
    v = _val(out, "页面设置差异")
    full, basis = "", _basis(out, "页面设置差异")
    if jp.is_file():
        data = json.loads(jp.read_text(encoding="utf-8"))
        full = next((r[2] for f in data.get("files", []) for r in f.get("rows", [])
                     if r[:2] == ["页面", "页面设置差异"]), "")
    check("E82d 一行装不下：按段收、写明另有几段，不出现省略号；依据写明共几节不同；--json 里是完整取值",
          code == 0 and len(v) <= 400 and "…" not in v and re.search(r"另有 \d+ 段未列出", v)
          and "第 1 节" in v and "第 14 节" not in v and "其中 13 节与正文节不同" in basis
          and "第 14 节" in full and "未列出" not in full,
          f"{len(v)} 字，依据：{basis} ｜ 末尾：{v[-60:]} ｜ json 末尾：{full[-50:]}")
    code, out = tk("layout-read", _company_doc("多节对照"))
    v = _val(out, "页面设置差异")
    check("E82e 阴性对照——每节都不同且一行装得下的（E55 那份三节文档）：仍逐节照写",
          "第 1 节" in v and "第 2 节" in v and "3.17" in v and "1.91" in v and "另有" not in v, v)


def e83():
    """页码依据：正文所在节没单独定义页脚、沿用前面某一节的，要写「正文所在第 N 节沿用第 K 节的页脚」——
    改前写成「正文所在第 1 节」，跟同一张表里的「正文所在第 103 节」自相矛盾。"""
    p = _plain_doc("页脚沿用")
    code, out = tk("layout-read", p)
    b = _basis(out, "页码写法")
    check("E83a 正文节沿用前面节的页脚：依据写沿用关系，不把第 1 节说成正文所在节",
          code == 0 and "正文所在第 13 节" in b and "沿用第 1 节" in b and "正文所在第 1 节" not in b,
          _row(out, "页码写法"))
    check("E83b 页码取值照旧（第 X 页、居中）",
          "第 X 页" in _val(out, "页码写法") and "居中" in _val(out, "页码写法"), _row(out, "页码写法"))
    code, out = tk("layout-read", _company_doc("页脚自有"))
    b = _basis(out, "页码写法")
    check("E83c 阴性对照——正文节自己定义页脚的（E55 那份）：依据照旧，不提沿用",
          "正文所在第 3 节" in b and "沿用" not in b, _row(out, "页码写法"))
    # 页码是正文节之后才定义的：正文节没有页码，取值与依据都要照实说
    p = _plain_doc("页脚在后面")
    d = Document(str(p))
    d.sections[0].footer.is_linked_to_previous = True       # 第 1 节不再定义页脚
    f19 = d.sections[18].footer                             # 第 19 节（第 10 章的正文节）才定义
    f19.is_linked_to_previous = False                       # 先给它一份自己的，别沿链加回第 1 节
    _field(f19.paragraphs[0], "PAGE", "10", prefix="第 ", suffix=" 页")
    d.save(str(p))
    code, out = tk("layout-read", p)
    v, b = _val(out, "页码写法"), _basis(out, "页码写法")
    check("E83d 页码在正文节之后才定义：取值说正文所在节没有页码，依据写明取的是哪一节的写法",
          code == 0 and v.startswith("正文所在节没有页码（第 19 节起才有：第 X 页")
          and "正文所在第 13 节没有页码" in b, _row(out, "页码写法"))


def e84():
    """目录缩进：缩进打在条目文字里的要读出来（手工目录的常见写法：二级条目以全角空格起头）——
    改前只看段落缩进，143 条二级条目读成「缩进 0 厘米」，模型据此说这份标书「目录不缩进」。"""
    p = _plain_doc("目录文字缩进")
    code, out = tk("layout-read", p)
    v, r = _val(out, "目录缩进"), _row(out, "目录缩进")
    check("E84a 段落缩进为 0、文字以全角空格起头：两件事都读出来",
          code == 0 and "一级 0 厘米" in v and "二级 0 厘米" in v and "2 个全角空格" in v and "3/3 条" in v, r)
    check("E84b 依据写明读的是段落缩进与条目文字", "文字" in _basis(out, "目录缩进"), r)
    code, out = tk("layout-read", _company_doc("目录段落缩进"))
    v = _val(out, "目录缩进")
    check("E84c 阴性对照——缩进在段落上、文字不带空格的（E58 那份）：取值不变、不提空格",
          "0.74" in v and "空格" not in v, v)


def e85():
    """回显末尾带下一步与取值纪律。真机上模型只看了回显最后 22 行，照样越过了 SKILL.md 里加粗写着的
    边界：手工搭了两份模板、自创标题推荐成默认、写进用户长期记忆。模型每跑一次都会读回显，下一步写在这里。"""
    p = _company_doc("边界回显")
    code, out = tk("layout-read", p)
    tail = out.splitlines()[-2:]
    check("E85a 回显最后两行是下一步：表交给用户确认，点头后登记为公司版式；不替用户手工搭模板",
          code == 0 and len(tail) == 2 and tail[0].startswith("[下一步]") and "用户确认" in tail[0]
          and "layout-adopt" in tail[0] and "不替用户手工搭模板" in tail[0], "\n".join(tail))
    check("E85b 紧跟着是取值纪律：读出什么报什么、由用户确认、不推荐「更好看」的写法",
          tail[-1].startswith("[取值]") and "读出什么报什么" in tail[-1] and "由用户确认" in tail[-1]
          and "不推荐" in tail[-1], tail[-1] if tail else "")
    spec = WORK / "边界-规格.md"
    code, out = tk("layout-read", p, "--out", spec)
    check("E85c 带 --out 时两条仍在回显里（且排在落盘提示之前）；不写进规格表文件",
          code == 0 and "[下一步]" in out and "[取值]" in out
          and out.index("[下一步]") < out.index("[规格表已落盘]")
          and "[下一步]" not in spec.read_text(encoding="utf-8"), out[-260:])
    code, out = tk("layout-read", p, p)
    check("E85d 多份文件：每份后面各带一次", out.count("[下一步]") == 2 and out.count("[取值]") == 2,
          f"下一步 {out.count('[下一步]')} 次")


# ---------------------------------------------------------------- 第 19 批 公司版式：登记与自动采用
#
# 规格表是公司版式的本体：确认后登记到用户目录，之后建节、合册不指定模板就自动按它出。用例各用
# 一个独立的登记目录（环境变量 TENDER_LAYOUT_HOME），tk 缺省把它指向空目录——电池既不读本机
# 真实的登记，也不往里写。


def _LE(home):
    return {"TENDER_LAYOUT_HOME": str(home)}


def _lay_home(name):
    """一个独立的版式登记目录（每个用例自己一份，互不串）。"""
    h = WORK / f"layout-{name}"
    h.mkdir(parents=True, exist_ok=True)
    return h


def _adopt(name, doc, *sets, source=None):
    """读夹具的规格表并登记（可带对话修订、可带 --source）。返回 (登记目录, 规格表路径, 登记回显, 退出码)。"""
    home = _lay_home(name)
    spec = WORK / f"{name}-规格.md"
    code, out = tk("layout-read", doc, "--out", spec, env=_LE(home))
    assert code == 0 and spec.is_file(), out
    args = ["layout-adopt", spec, *[x for s in sets for x in ("--set", s)]]
    if source is not None:
        args += ["--source", str(source)]
    code, out = tk(*args, env=_LE(home))
    return home, spec, out, code


def _lay_work(name, home, outline, leaves):
    """搭一个可合册的工作区，节文件按登记的版式建。leaves：[(节名, body 块列表)]。"""
    work, _ = make_work(name, outline)
    for leaf, blocks in leaves:
        code, out = tk("create", work / "body" / f"{leaf}.docx", "--title", leaf,
                       "--body", body(f"lay-{name}-{leaf}", blocks), env=_LE(home))
        assert code == 0, out
    return work


_LSENT2 = _LSENT * 2      # 够 30 字，进得了正文候选


def _header_doc(name):
    """带封面、正文节页眉（logo 图 + 公司名 + 横线）与页码的形态——登记封面与页眉用的真实形态。
    正文段都放在第二节（正文所在节=有页眉页码的那节），封面节不带页脚。"""
    from docx.enum.section import WD_SECTION_START
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    d = _doc()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    for t in ("测试科技股份有限公司", "智慧园区数据平台项目投标文件"):
        p = d.add_paragraph(t)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic = d.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(_png(name)), width=Cm(5))
    p = d.add_paragraph("日期：2026 年 9 月")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_section(WD_SECTION_START.NEW_PAGE)     # 封面节到此为止
    s2 = d.sections[-1]
    s2.header.is_linked_to_previous = False
    hp = s2.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.add_run().add_picture(str(_png(name + "-logo")), width=Cm(0.8))
    _run_fmt(hp.add_run("　测试科技股份有限公司"), "宋体", 9)
    _p_bottom_border(hp)
    s2.footer.is_linked_to_previous = False
    fp = s2.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _field(fp, "PAGE", "1")
    for _ in range(3):
        d.add_paragraph(_LSENT2)
    p = WORK / f"{name}.docx"
    d.save(str(p))
    return p


def _style_fonts(d, name):
    rpr = d.styles[name].element.rPr
    rf = rpr.rFonts if rpr is not None else None
    return ((rf.get(qn("w:eastAsia")) if rf is not None else None),
            (rf.get(qn("w:ascii")) if rf is not None else None))


def _table_borders(d):
    b = d.tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
    return {e.tag.split("}")[1]: e.get(qn("w:val")) for e in b} if b is not None else {}


def _spec_table(text):
    """规格表（或 layout-read 回显）里的 {项目: 取值}。"""
    out, grp = {}, ""
    for ln in text.splitlines():
        if ln.startswith("【") and ln.endswith("】"):
            grp = ln[1:-1]
        elif ln.startswith("| ") and grp:
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            if len(cells) >= 4 and cells[0] != "项目" and not set(cells[0]) <= set("-: "):
                out.setdefault(cells[0], cells[1])
    return out


# 往返核对的项目：读数写法与出稿口径一一对应的那些（目录缩进、目录点线的读法与写法不同义，单列在 E88）
_RT = ("纸张", "页边距", "文档网格", "页码写法", "正文字体", "正文字号", "行距", "段前段后", "缩进", "对齐",
       "目录标题", "目录一级条目", "目录二级条目", "目录三级条目", "目录行距", "目录层级", "编号形态", "编号间隔",
       "章起新页", "标题层级", "标题一级", "标题一级间距", "标题二级", "标题二级间距", "标题三级", "标题三级间距",
       "表格边框", "表头样式", "表格文字", "表题", "图题")


def e86():
    """layout-adopt：把确认后的规格表登记为公司版式；layout-show 列出；不是规格表的文件拒收。"""
    p = _plain_doc("登记源")
    sha = p.read_bytes()
    home, spec, out, code = _adopt("登记", p)
    check("E86a 登记成功：名称取自规格表来源行，回显写明后面的标书自动按它出、生效与未生效清单、存放位置",
          code == 0 and out.startswith("[版式登记]") and "「登记源」" in out and "自动按它出" in out
          and "生效：" in out and "还没接上" in out and "封面" in out and str(home) in out, out[:600])
    check("E86b 规格表存进登记目录（layout.md 逐字相同），登记记录在 layout.json；被读的文件没动",
          (home / "layout.md").is_file()
          and (home / "layout.md").read_text(encoding="utf-8") == spec.read_text(encoding="utf-8")
          and (home / "layout.json").is_file() and p.read_bytes() == sha)
    code, out = tk("layout-show", env=_LE(home))
    check("E86c layout-show：名称、登记时间、规格表全文、生效与未生效",
          code == 0 and "「登记源」" in out and "【页面】" in out and "| 纸张 | A4 纵向 |" in out
          and "还没接上" in out, out[:300])
    code, out = tk("layout-show", env=_LE(_lay_home("空")))
    check("E86d 没登记时如实说没有（后面的标书按内置版式出）",
          code == 0 and "没有登记" in out and "内置版式" in out, out[:200])
    bad = WORK / "随手写的.md"
    bad.write_text("# 随手写的\n\n这里没有规格表。\n", encoding="utf-8")
    h2 = _lay_home("拒收")
    code, out = tk("layout-adopt", bad, env=_LE(h2))
    check("E86e 不是规格表的文件拒收：退出码 2、一行中文提示、不落登记",
          code == 2 and "版式规格表" in out and "Traceback" not in out and not (h2 / "layout.json").exists(),
          f"exit={code} {out[:200]}")
    code, out = tk("layout-adopt", WORK / "不存在.md", env=_LE(h2))
    check("E86f 规格表不存在一行提示", code == 2 and "不存在" in out, f"exit={code} {out[:200]}")


def e87():
    """建节自动按登记的版式出（不传 --template）：样式、缩进、表格、表题全跟版式走。"""
    home, spec, out, _ = _adopt("建节", _direct_doc("建节源"),
                                "正文字体=中文仿宋、西文 Arial", "正文字号=14 磅", "行距=固定 24 磅", "缩进=无",
                                "段前段后=段前自动、段后 6 磅", "对齐=两端",
                                "标题一级=中文黑体、西文 Times New Roman｜小二（18 磅）｜加粗｜居中",
                                "表格边框=三线表", "表头样式=无底纹", "表格文字=中文仿宋｜小四（12 磅）｜左",
                                "表题=表下方、中文仿宋、西文 Times New Roman｜小四（12 磅）｜居中、编号「表 N-M」")
    path = WORK / "建节-正文.docx"
    code, out = tk("create", path, "--title", "项目理解", "--body", body("建节体", [
        {"text": _LSENT2}, {"text": "小节标题", "style": "h2"},
        {"type": "table", "caption": "表 1-1　人员配置表", "header": ["岗位", "人数"], "rows": [["项目经理", "1"]]}]),
        env=_LE(home))
    check("E87a 建节回显点名按公司版式出（没传 --template）",
          code == 0 and "[版式] 本次按公司版式「建节源」" in out, out[:300])
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    d = Document(str(path))
    normal = d.styles["Normal"]
    pf = normal.paragraph_format
    sp = normal.element.pPr.find(qn("w:spacing")) if normal.element.pPr is not None else None
    check("E87b 正文样式跟版式：仿宋／Arial、14 磅、固定 24 磅、段前自动、段后 6 磅、两端对齐",
          _style_fonts(d, "Normal") == ("仿宋", "Arial") and normal.font.size.pt == 14
          and pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY and round(pf.line_spacing.pt, 1) == 24
          and pf.space_after.pt == 6 and sp is not None and sp.get(qn("w:beforeAutospacing")) == "1"
          and pf.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY,
          f"{_style_fonts(d, 'Normal')} {normal.font.size.pt} {pf.line_spacing} {pf.space_after} {sp.attrib if sp is not None else None}")
    h1 = d.styles["Heading 1"]
    char = d.styles["Heading 1 Char"].element.xml
    check("E87c 一级标题跟版式（黑体 18 磅加粗居中），链接字符样式同步、不留主题字体与主题蓝",
          _style_fonts(d, "Heading 1") == ("黑体", "Times New Roman") and h1.font.size.pt == 18 and h1.font.bold
          and h1.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
          and "<w:b/>" in char and "themeColor" not in char and "asciiTheme" not in char, char[:400])
    ind = [p for p in d.paragraphs if p.text == _LSENT2][0]._p.pPr.find(qn("w:ind"))
    check("E87d 正文段缩进跟版式（无）：显式写零", ind is not None and ind.get(qn("w:firstLine")) == "0")
    check("E87e 节内小标题挂二级标题样式",
          [p.style.name for p in d.paragraphs if p.text == "小节标题"] == ["Heading 2"])
    t = d.tables[0]
    hdr_run, cell_run = t.cell(0, 0).paragraphs[0].runs[0], t.cell(1, 0).paragraphs[0].runs[0]
    tcpr = t.cell(0, 0)._tc.tcPr
    borders = _table_borders(d)
    check("E87f 表格跟版式：仿宋 12 磅、表头不加粗无底纹、三线表",
          hdr_run._element.rPr.rFonts.get(qn("w:eastAsia")) == "仿宋" and hdr_run.font.size.pt == 12
          and not hdr_run.font.bold and cell_run.font.size.pt == 12
          and (tcpr is None or tcpr.find(qn("w:shd")) is None)
          and borders.get("top") == "single" and borders.get("bottom") == "single"
          and all(borders.get(k) in ("nil", "none") for k in ("left", "right", "insideH", "insideV")),
          f"{borders} bold={hdr_run.font.bold}")
    cap = [p for p in d.paragraphs if p.text.startswith("表 1-1")][0]
    kids = list(d.element.body)
    tc = d.styles["Table Caption"]
    check("E87g 表题跟版式：仿宋 12 磅不加粗居中，落在表格下方",
          cap.style.name == "Table Caption" and _style_fonts(d, "Table Caption") == ("仿宋", "Times New Roman")
          and tc.font.size.pt == 12 and not tc.font.bold
          and tc.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
          and kids.index(cap._p) > next(i for i, c in enumerate(kids) if c.tag == qn("w:tbl")),
          f"{cap.style.name} {_style_fonts(d, 'Table Caption')} {tc.font.size.pt} {tc.font.bold}")


def e88():
    """合册自动按登记的版式出：页面与网格、页码、编号与间隔、章起新页、正文缩进、目录标题与缩进点线。"""
    home, spec, out, _ = _adopt("合册", _plain_doc("合册源"),
                                "页边距=上 3、下 2.54、左 2.6、右 2.4 厘米", "对称页边距=是",
                                "文档网格=行网格（每页约 44 行）", "页码写法=— X —、页脚右｜五号（10.5 磅）",
                                "编号形态=decimal（1/1.1）", "编号间隔=紧接", "章起新页=否",
                                "目录点线=制表位、无引导线", "目录缩进=一级 0 厘米、二级 0.74 厘米、三级 1.48 厘米",
                                "目录标题=中文宋体、西文 Times New Roman｜小四（12 磅）｜加粗｜左")
    work = _lay_work("合册", home, "# 响应文件：投标文件\n## 目录\n- 目录\n- 项目理解\n  - 建设背景\n"
                      "- 技术方案\n  - 总体架构\n",
                     [("建设背景", [{"text": _LSENT2}, {"text": "小节标题", "style": "h2"}, {"text": _LSENT2}]),
                      ("总体架构", [{"text": _LSENT2}])])
    code, out = tk("assemble", "--work", work, env=_LE(home))
    check("E88a 合册回显点名按公司版式出（没传 --template、没传 --numbering）",
          code == 0 and "[版式] 本次按公司版式「合册源」" in out, out[:300])
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    vol = next((work / "body").glob("整本-*.docx"))
    d = Document(str(vol))
    s = d.sections[-1]
    grid = s._sectPr.find(qn("w:docGrid"))
    sides = tuple(round(x.cm, 1) for x in (s.top_margin, s.bottom_margin, s.left_margin, s.right_margin))
    check("E88b 页面跟版式：页边距 3／2.54／2.6／2.4、行网格（每页约 44 行）、对称页边距",
          sides == (3.0, 2.5, 2.6, 2.4) and grid is not None and grid.get(qn("w:type")) == "lines"
          and 300 <= int(grid.get(qn("w:linePitch"))) <= 322 and "<w:mirrorMargins" in xml(vol, "word/settings.xml"),
          f"{sides} {grid.attrib if grid is not None else None}")
    fp = s.footer.paragraphs[0]
    check("E88c 页码跟版式：— X —、靠右、五号",
          fp.text.startswith("— ") and fp.text.endswith(" —") and fp.alignment == WD_ALIGN_PARAGRAPH.RIGHT
          and fp.runs[0].font.size.pt == 10.5
          and "PAGE" in "".join(e.text or "" for e in fp._p.iter(qn("w:instrText"))), f"{fp.text!r} {fp.alignment}")
    h1 = [p for p in d.paragraphs if p.style.name == "Heading 1"]
    h2 = [p.text for p in d.paragraphs if p.style.name == "Heading 2"]
    h3 = [p.text for p in d.paragraphs if p.style.name == "Heading 3"]
    check("E88d 编号跟版式：decimal 方案、编号紧接标题；章不起新页",
          [p.text for p in h1] == ["1项目理解", "2技术方案"] and "1.1建设背景" in h2 and "1.1.1小节标题" in h3
          and not h1[1].paragraph_format.page_break_before, f"{[p.text for p in h1]} {h2} {h3}")
    ind = [p for p in d.paragraphs if p.text.endswith(_LSENT2)][0]._p.pPr.find(qn("w:ind"))
    check("E88e 正文段缩进跟版式（首行 2 字）", ind is not None and ind.get(qn("w:firstLineChars")) == "200")
    title = [p for p in d.paragraphs if p.text == "目录"][0]
    check("E88f 目录标题跟版式：宋体小四加粗居左",
          title.alignment == WD_ALIGN_PARAGRAPH.LEFT and title.runs[0].font.bold
          and title.runs[0].font.size.pt == 12 and title.runs[0]._element.rPr.rFonts.get(qn("w:eastAsia")) == "宋体")
    e1 = [p for p in d.paragraphs if p.style.name == "TOC 1"][0]
    tab = e1._p.pPr.find(qn("w:tabs")).find(qn("w:tab"))
    toc_cm = [round(d.styles[f"TOC {n}"].paragraph_format.left_indent.cm, 2) for n in (1, 2, 3)]
    check("E88g 目录条目跟版式：无引导线、逐级缩进 0／0.74／1.48 厘米",
          tab.get(qn("w:leader")) in (None, "none") and toc_cm == [0.0, 0.74, 1.48],
          f"leader={tab.get(qn('w:leader'))} 缩进={toc_cm}")


def e89():
    """显式模板与 --ignore-layout 都排在登记的版式前面；没有登记时一切照旧。"""
    home, spec, out, _ = _adopt("优先级", _plain_doc("优先级源"))
    base = ROOT / "tender-docx" / "templates" / "base-template.docx"
    for tag, extra in (("显式传 --template", ["--template", str(base)]), ("--ignore-layout", ["--ignore-layout"])):
        path = WORK / f"优先级-{tag[-4:-1]}.docx"
        code, out = tk("create", path, "--title", "甲", *extra, env=_LE(home))
        ok = path.is_file() and Document(str(path)).styles["Heading 1"].font.size.pt == 16
        check(f"E89 {tag}：按内置版式出（一级标题黑体 16 磅加粗），回显不提公司版式",
              code == 0 and "[版式]" not in out and ok, f"exit={code} {out[:200]}")
    path = WORK / "无登记-正文.docx"
    code, out = tk("create", path, "--title", "甲")     # tk 缺省把登记目录指到空目录
    check("E89 没有登记时照旧：内置版式，回显不提公司版式",
          code == 0 and "[版式]" not in out and Document(str(path)).styles["Heading 1"].font.size.pt == 16)


def e90():
    """往返一致：按登记的版式出一本样张，再把样张读回规格表——读出来的和登记的对得上（两种真实形态）。"""
    for label, fn, deep in (("纯文字形态", _plain_doc, True), ("直接格式形态", _direct_doc, False)):
        name = "往返纯" if deep else "往返直"
        home, spec, out, code = _adopt(name, fn(f"{name}源"))
        assert code == 0, out
        blocks = [{"text": _LSENT2}] + ([{"text": "小节标题", "style": "h2"}] if deep else [{"text": _LSENT2}]) + [
            {"type": "table", "caption": "表 1-1　人员配置表", "header": ["岗位", "人数"], "rows": [["项目经理", "1"]]},
            {"text": _LSENT2}]
        work = _lay_work(name, home, "# 响应文件：投标文件\n## 目录\n- 目录\n- 项目理解\n  - 建设背景\n"
                                     "- 技术方案\n  - 总体架构\n",
                         [("建设背景", blocks), ("总体架构", [{"text": _LSENT2}])])
        code, out = tk("image", work / "body" / "建设背景.docx", "--img", _png(name), "--caption", "图 1-1　拓扑图",
                       env=_LE(home))
        assert code == 0, out
        code, out = tk("assemble", "--work", work, env=_LE(home))
        assert code == 0, out
        vol = next((work / "body").glob("整本-*.docx"))
        code, out = tk("layout-read", vol, env=_LE(home))
        want, got = _spec_table(spec.read_text(encoding="utf-8")), _spec_table(out)
        # 历史标书里没有的东西（未发现表格、未见表题）按内置写法出，不参与核对
        diff = [f"{k}：登记「{want[k]}」／读回「{got.get(k, '〈缺〉')}」"
                for k in _RT if k in want and not want[k].startswith(("未发现", "未见")) and want[k] != got.get(k)]
        check(f"E90 {label}：按登记的版式出的样张，读回的规格与登记的一致", code == 0 and not diff,
              "；".join(diff)[:700])


def e91():
    """对话修订：--set、直接改登记的规格表、改坏了；layout-forget 取消登记。"""
    home, spec, out, _ = _adopt("修订", _plain_doc("修订源"))
    val = "中文黑体、西文 Times New Roman｜三号（16 磅）｜加粗｜居中"
    code, out = tk("layout-adopt", spec, "--set", f"标题一级={val}", env=_LE(home))
    md = (home / "layout.md").read_text(encoding="utf-8")
    row = next((ln for ln in md.splitlines() if ln.startswith("| 标题一级 |")), "")
    check("E91a 对话修订 --set：取值写进登记的规格表（依据注明对话修订），回显点名改了哪一项",
          code == 0 and f"| 标题一级 | {val} |" in row and "对话修订" in row and "标题一级" in out, row or out[:300])

    def build(tag):
        path = WORK / f"修订-{tag}.docx"
        code, out = tk("create", path, "--title", "甲", env=_LE(home))
        return code, out, (Document(str(path)) if path.is_file() else None)

    code, out, d = build("改后")
    check("E91b 修订后建节按改后的版式出",
          code == 0 and d.styles["Heading 1"].font.size.pt == 16 and d.styles["Heading 1"].font.bold
          and _style_fonts(d, "Heading 1")[0] == "黑体")
    md2 = md.replace("| 页边距 | 上 2、下 2、左 2、右 2 厘米 |", "| 页边距 | 上 3、下 3、左 3、右 3 厘米 |")
    (home / "layout.md").write_text(md2, encoding="utf-8")
    code, out, d = build("直改")
    check("E91c 直接改了登记的规格表：下一次出稿按改后的表读，并回显一句",
          code == 0 and "规格表在登记后改过" in out and round(d.sections[0].top_margin.cm, 1) == 3.0, out[:200])
    (home / "layout.md").write_text(md2.replace("| 行距 | 1.5 倍 |", "| 行距 | 随便写的 |"), encoding="utf-8")
    code, out, d = build("改坏")
    check("E91d 改坏了：仍按上一次读得出来的取值出稿，回显点名哪一行读不出来",
          code == 0 and "读不出来" in out and "行距" in out
          and d.styles["Normal"].paragraph_format.line_spacing == 1.5, out[:300])
    before = (home / "layout.md").read_text(encoding="utf-8")
    code, out = tk("layout-adopt", spec, "--set", "标题一级=黑体三号", env=_LE(home))
    check("E91e --set 取值写法不对：拒收并给出这个项目的写法样例，登记不变",
          code == 2 and "标题一级" in out and "中文" in out and "｜" in out
          and (home / "layout.md").read_text(encoding="utf-8") == before, f"exit={code} {out[:300]}")
    code, out = tk("layout-adopt", spec, "--set", "没有这一项=随便", env=_LE(home))
    check("E91f --set 项目名不对：拒收并列出可改的项目", code == 2 and "可改" in out, f"exit={code} {out[:300]}")
    code, out = tk("layout-forget", env=_LE(home))
    kept = list((home / "历史").glob("*.md")) if (home / "历史").is_dir() else []
    check("E91g layout-forget：取消登记（原规格表留档），之后的标书按内置版式出",
          code == 0 and "[版式取消]" in out and not (home / "layout.json").exists() and kept, out[:300])
    code, out, d = build("取消后")
    check("E91h 取消后建节回到内置版式、不提公司版式",
          code == 0 and "[版式]" not in out and d.styles["Heading 1"].font.size.pt == 16
          and d.styles["Heading 1"].paragraph_format.alignment is None)
    # 用户给版式起的叫法：修订的是同一份表，叫法留着，不被来源文件名顶掉
    code, out = tk("layout-adopt", spec, "--name", "我们公司的版式", env=_LE(home))
    code, out = tk("layout-adopt", home / "layout.md", "--set", "缩进=无", env=_LE(home))
    check("E91i 起过叫法的版式，对话修订后叫法不变",
          code == 0 and "「我们公司的版式」" in out and "对话修订已记入：缩进" in out, out[:300])


def e92():
    """读数补洞（按表落模板要用）：段前段后／行距带上文档默认值；目录条目段前段后单列一行；decimal 编号也读间隔。"""
    d = _doc()     # python-docx 默认模板的文档默认值：段后 10 磅、行距 1.15 倍
    d.styles["Normal"].paragraph_format.space_after = Pt(0)
    for _ in range(3):
        d.add_paragraph(_LSENT2)
    p = WORK / "默认段距.docx"
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E92a 行距只写在文档默认值里的也读出来（1.15 倍，不是「单倍」）",
          code == 0 and _val(out, "行距") == "1.15 倍" and _val(out, "段前段后") == "段后 0 磅",
          _row(out, "行距") + _row(out, "段前段后"))
    d = _doc()
    d.styles["Normal"].paragraph_format.line_spacing = 1.5
    for _ in range(3):
        d.add_paragraph(_LSENT2)
    p = WORK / "默认段距二.docx"
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E92b 样式链上都没设段距、文档默认值里是段后 10 磅：读成「段后 10 磅」",
          code == 0 and _val(out, "段前段后") == "段后 10 磅", _row(out, "段前段后"))
    code, out = tk("layout-read", _company_doc("目录段距"))
    check("E92c 目录条目的段前段后单列一行",
          code == 0 and _row(out, "目录段前段后") != "" and "目录条目" in _row(out, "目录段前段后"),
          _row(out, "目录段前段后"))
    d = _doc()
    for t in ("1.1　建设背景", "1.2　总体方案", "1.3　实施计划"):
        d.add_paragraph().add_run(t).font.bold = True
        d.add_paragraph(_LSENT2)
    p = WORK / "十进制编号.docx"
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E92d decimal 编号（文字自带 1.1）也读编号与文字的间隔",
          code == 0 and _val(out, "编号形态").startswith("decimal") and _val(out, "编号间隔") == "全角空格"
          and "二级标题" in _row(out, "编号间隔"), _row(out, "编号间隔") + _row(out, "编号形态"))
    # 页码是纯数字（域前后没有文字）：字号看域 run 自己写的，不退到文档默认字号（真实回放抓到）
    p = _plain_doc("纯数字页码")
    d = Document(str(p))
    fp = d.sections[0].footer.paragraphs[0]
    for r in list(fp.runs):
        r._element.getparent().remove(r._element)
    _field(fp, "PAGE", "1", prefix="", suffix="")
    for r in fp.runs:
        r.font.size = Pt(10.5)
    d.save(str(p))
    code, out = tk("layout-read", p)
    check("E92e 纯数字页码的字号读域 run 上的（五号），不是文档默认的 11 磅",
          code == 0 and _val(out, "页码写法") == "纯数字、页脚居中｜五号（10.5 磅）", _row(out, "页码写法"))


def e93():
    """layout-adopt --source：提取封面与页眉存进登记目录；不带 --source 时维持「还没接上」并给一句怎么接。"""
    src = _header_doc("提取源")
    home, spec, out, code = _adopt("提取", src, source=src)
    rec = json.loads((home / "layout.json").read_text(encoding="utf-8"))
    assets = rec.get("assets") or {}
    hd = assets.get("header") or {}
    cd = Document(str(home / "封面.docx")) if (home / "封面.docx").is_file() else None
    n_cover_img = len([r for r in cd.part.rels.values() if r.reltype.endswith("/image")]) if cd else 0
    check("E93a 带源登记：封面.docx 落进登记目录，里面有封面的文字与图；页眉 logo 图进 媒体/",
          code == 0 and cd is not None
          and "智慧园区数据平台项目投标文件" in cd.paragraphs[1].text and n_cover_img >= 1
          and hd.get("images") and (home / hd["images"][0]["file"]).is_file(), out[:400])
    check("E93b 页眉规格进了登记：文字、字号、对齐、横线、图尺寸都在",
          hd.get("text") == "测试科技股份有限公司" and hd.get("align") == "居中" and hd.get("border")
          and hd["images"][0].get("cx"), json.dumps(hd, ensure_ascii=False)[:300])
    check("E93c 登记回显：生效清单里有封面与页眉，「还没接上」只剩表格列宽",
          code == 0 and "封面（" in out and "页眉（" in out and "表格列宽" in out
          and "封面整页" not in out and "logo 与文字）" not in out.split("还没接上")[-1], out[-600:])
    home2, spec2, out2, code2 = _adopt("不提取", _header_doc("提取源二"))
    check("E93d 不带 --source：维持「还没接上」，并给一句怎么接（--source）",
          code2 == 0 and "封面整页" in out2 and "logo 与文字" in out2 and "--source" in out2, out2[-500:])


def e94():
    """合册落地：封面节点无节文件时落公司封面并点名待改字段；正文页页眉带 logo／公司名／横线，前置页不带。"""
    home, spec, out, code = _adopt("落地", _header_doc("落地源"), source=_header_doc("落地源二"))
    work = _lay_work("落地", home, "# 响应文件：投标文件\n## 目录\n- 封面\n- 目录\n- 甲节\n",
                     [("甲节", [{"text": _LSENT2}])])
    code, out = tk("assemble", "--work", work, env=_LE(home))
    cover_file = work / "body" / "封面.docx"
    check("E94a 封面落为公司封面 body/封面.docx（内容与图都在），回显点名待改字段",
          code == 0 and cover_file.is_file()
          and "智慧园区数据平台项目投标文件" in Document(str(cover_file)).paragraphs[1].text
          and "公司封面" in out and "字段" in out,
          f"exit={code} {out[:400]}")
    vol = next((work / "body").glob("整本-*.docx"))
    d = Document(str(vol))
    body_sec = d.sections[-1]
    front_sec = d.sections[0]
    hp = body_sec.header
    img_rels = [r for r in hp.part.rels.values() if r.reltype.endswith("/image")]
    pBdr = hp.paragraphs[0]._p.pPr.find(qn("w:pBdr")) if hp.paragraphs and hp.paragraphs[0]._p.pPr is not None else None
    check("E94b 正文节页眉带 logo 图、公司名与横线",
          any(p.text.strip().endswith("测试科技股份有限公司") for p in hp.paragraphs)
          and len(img_rels) >= 1 and pBdr is not None and pBdr.find(qn("w:bottom")) is not None,
          f"headers={[p.text for p in hp.paragraphs]} imgs={len(img_rels)} border={pBdr is not None}")
    check("E94c 前置节（封面/目录页）不带页眉页脚引用；正文节页码口径不变（页脚有页码域）",
          front_sec.header.is_linked_to_previous and front_sec.footer.is_linked_to_previous
          and "PAGE" in "".join(e.text or "" for e in body_sec.footer.paragraphs[0]._p.iter(qn("w:instrText"))),
          f"front linked={front_sec.header.is_linked_to_previous}/{front_sec.footer.is_linked_to_previous}")
    d2 = Document(str(cover_file))
    check("E94d 封面节文件元数据干净（统一落盘口；节文件不记指纹——那是整本的约定）",
          (d2.core_properties.author or "") == "", "")


def e95():
    """用户（或模型）已写 body/封面.docx 时：不覆盖、不点名公司封面。"""
    home, spec, out, code = _adopt("不覆盖", _header_doc("不覆盖源"), source=_header_doc("不覆盖源二"))
    work = _lay_work("不覆盖", home, "# 响应文件：投标文件\n## 目录\n- 封面\n- 目录\n- 甲节\n",
                     [("甲节", [{"text": _LSENT2}])])
    code, out = tk("create", work / "body" / "封面.docx", "--title", "封面",
                   "--body", body("封面体", [{"text": "这是模型按本次项目写的封面。", "style": "cover-title"}]),
                   env=_LE(home))
    assert code == 0, out
    code, out = tk("assemble", "--work", work, env=_LE(home))
    vol = next((work / "body").glob("整本-*.docx"))
    check("E95 已有封面节文件：按用户写的合册，不落公司封面、不点名",
          code == 0 and "公司封面" not in out
          and "这是模型按本次项目写的封面。" in "\n".join(p.text for p in Document(str(vol)).paragraphs),
          f"exit={code} {out[:300]}")


def e96():
    """单元格边距跟版式：读到非缺省值（0.19 厘米）才显式写进表格，缺省值不添乱。"""
    home, spec, out, code = _adopt("边距", _plain_doc("边距源"), "单元格边距=左右 0.35 厘米")
    path = WORK / "边距-正文.docx"
    code, out = tk("create", path, "--title", "甲",
                   "--body", body("边距体", [{"type": "table", "header": ["岗位", "人数"], "rows": [["项目经理", "1"]]}]),
                   env=_LE(home))
    mar = Document(str(path)).tables[0]._tbl.tblPr.find(qn("w:tblCellMar"))
    left = mar.find(qn("w:left")).get(qn("w:w")) if mar is not None else None
    check("E96a 版式读到 0.35 厘米：表格显式写单元格边距（198 缇）",
          code == 0 and mar is not None and left == "198", f"mar={mar.attrib if mar is not None else None}")
    home2, spec2, out2, code2 = _adopt("边距缺省", _plain_doc("边距源二"))
    path2 = WORK / "边距缺省-正文.docx"
    code2, out2 = tk("create", path2, "--title", "甲",
                     "--body", body("边距体二", [{"type": "table", "header": ["岗位", "人数"], "rows": [["项目经理", "1"]]}]),
                     env=_LE(home2))
    mar2 = Document(str(path2)).tables[0]._tbl.tblPr.find(qn("w:tblCellMar"))
    check("E96b 缺省 0.19 厘米：不显式写（Word 缺省就是它）",
          code2 == 0 and mar2 is None, f"exit={code2} mar={mar2 is not None}")


def e97():
    """首次登记：登记目录还不存在（真机第一次跑就是这个形态）时 --source 提取照常成功。
    改前 _extract_layout_assets 在 _write_record 之前跑，home 不存在直接 FileNotFoundError，
    整条登记失败，报错还谎称「规格表本身已登记生效」。"""
    src = _header_doc("全新登记源")
    home = WORK / "layout-全新"          # 故意不建目录
    spec = WORK / "全新-规格.md"
    code, out = tk("layout-read", src, "--out", spec, env=_LE(home))
    assert code == 0 and spec.is_file(), out
    code, out = tk("layout-adopt", spec, "--source", str(src), env=_LE(home))
    check("E97a 首次登记（登记目录不存在）：登记成功、封面与页眉落进登记目录",
          code == 0 and (home / "封面.docx").is_file() and (home / "layout.json").is_file()
          and (home / "媒体" / "页眉-1.png").is_file(), f"exit={code} {out[:300]}")
    rec = json.loads((home / "layout.json").read_text(encoding="utf-8"))
    check("E97b 登记记录带资产清单（封面+页眉）",
          (rec.get("assets") or {}).get("cover") == "封面.docx"
          and (rec["assets"].get("header") or {}).get("images"), json.dumps(rec.get("assets"), ensure_ascii=False)[:200])
    code, out = tk("layout-show", env=_LE(home))
    check("E97c 登记后查看：生效清单含封面与页眉，还没接上只剩表格列宽",
          code == 0 and "封面（" in out and "页眉（" in out and "表格列宽" in out, out[:420])


def e98():
    """节文件小标题跳级（直接用 h3、没经过 h2）：合册不崩、编号按缺位续——真机 2026-11-28
    KeyError（sub_counters 缺中间层级的键），模型就地打的补丁落成官方修复。"""
    work, _ = make_work("跳级", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n- 乙节\n", [])
    for leaf, blocks in (("甲节", [{"text": "跳过二级的小标题", "style": "h3"}, {"text": _LSENT2}]),
                         ("乙节", [{"text": _LSENT2}])):
        code, out = tk("create", work / "body" / f"{leaf}.docx", "--title", leaf,
                       "--body", body(f"跳级-{leaf}", blocks))
        assert code == 0, out
    code, out = tk("assemble", "--work", work)
    vol = next((work / "body").glob("整本-*.docx"), None)
    d = Document(str(vol)) if vol else None
    sub = next((p.text for p in d.paragraphs if p.style.name == "Heading 3"), "") if d else ""
    check("E98 小标题跳级（h3 没有 h2）：合册不崩、编号缺位按 1 续（1.1.1）",
          code == 0 and sub.startswith("1.1.1"), f"exit={code} {out[:280]} 小标题={sub!r}")


def e99():
    """封面提取保留空段——空段就是封面的竖向排版。真机 2026-11-28：提取时把空段滤掉，
    公司封面合进整本后内容全挤到页顶。"""
    d = _doc()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    for _ in range(3):
        d.add_paragraph()                      # 封面顶部的空段（竖向定位全靠它们）
    d.add_paragraph("测试科技股份有限公司")
    d.add_paragraph("智慧园区数据平台项目投标文件")
    from docx.enum.text import WD_BREAK
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph(_LSENT2)
    src = WORK / "封面空段源.docx"
    d.save(str(src))
    home, spec, out, code = _adopt("封面空段", src, source=src)
    cd = Document(str(home / "封面.docx")) if (home / "封面.docx").is_file() else None
    lead_empty = len(list(itertools.takewhile(lambda p_: not p_.text.strip(),
                                              cd.paragraphs))) if cd else -1
    check("E99 封面提取保留空段：顶部的 3 个空段原样带走，首行文字还在第 4 段",
          code == 0 and cd is not None and lead_empty == 3
          and cd.paragraphs[3].text.strip() == "测试科技股份有限公司",
          f"exit={code} 首个文字段位次={lead_empty + 1}")


def _parse_for_smoke(src, parse_root):
    """真跑 parse_document 生成带源指纹的解析产物（素材新鲜度闸的登记来源）。"""
    r = subprocess.run([sys.executable, "-B", str(ROOT / "tender-analysis" / "scripts" / "parse_document.py"),
                        str(src), "-o", str(parse_root)],
                       capture_output=True, text=True, errors="replace")
    assert r.returncode == 0, (r.stdout + r.stderr)
    return parse_root / src.name


def e100():
    src = WORK / "素材新鲜.docx"
    d = Document()
    d.add_paragraph("售后服务方案总则。")
    d.add_paragraph("故障响应时间为 2 小时。")
    d.add_paragraph("质保期三年。")
    d.save(str(src))
    pdir = _parse_for_smoke(src, WORK / "parse-素材")
    tgt = WORK / "新鲜节.docx"
    code, out = tk("create", tgt, "--title", "售后服务")
    assert code == 0, out
    code, out = tk("copy-from", src, tgt, "--parse", pdir)     # 阴性：未改源，整份照常拷
    ok0 = code == 0 and "已拷贝" in out
    d2 = Document(str(src))
    d2.add_paragraph("中途插入的一段，把后面的行整体推下去。")
    d2.save(str(src))                                          # 模拟用户在解析后改了素材，未重新解析
    code2, out2 = tk("copy-from", src, tgt, "--lines", "L1-L2", "--parse", pdir)
    check("E100 copy-from 源文件在解析后被改：拒绝（退出码 2）并给重解析+重核指引区间的处置",
          ok0 and code2 == 2 and "改过" in out2 and "重新解析" in out2 and "区间" in out2,
          f"阴性exit={code} 改后exit={code2} out={out2[:180]}")


def e101():
    src = WORK / "素材新鲜.docx"
    pdir = _parse_for_smoke(src, WORK / "parse-素材")          # 重新解析：登记记下新指纹
    tgt = WORK / "新鲜节.docx"
    code, out = tk("copy-from", src, tgt, "--parse", pdir)
    check("E101 重新解析后（指纹更新）照常拷贝", code == 0 and "已拷贝" in out, f"exit={code} out={out[:160]}")


def e102():
    pdir = WORK / "parse-素材" / "素材新鲜.docx"
    (pdir / "素材新鲜.meta.json").unlink()                     # 旧产物形态：elements 在、meta 不在
    tgt = WORK / "新鲜节.docx"
    code, out = tk("copy-from", WORK / "素材新鲜.docx", tgt, "--parse", pdir)
    check("E102 旧解析产物无源指纹：不拦、提醒建议重新解析、照常拷",
          code == 0 and "已拷贝" in out and "重新解析" in out, f"exit={code} out={out[:200]}")


def e103():
    """旧名名单.txt 隐形约定退场＋空名单不装查过。

    改前的两条缺陷：①残留扫描的名单除了 --old-names 参数还偷偷读 CWD 的 旧名名单.txt——
    全仓没有任何代码生成它、SKILL 与 --help 都不提它、换个目录跑校验就静默失效（名单
    唯一入口应回到显式的 --old-names）；②名单空着（模型没传名字，这是无人负责名单时的
    常态）时残留扫描压根没跑，输出照样打「[校验通过] 全部合规」——没查装查过。"""
    work, _ = make_work("w-oldname", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n")
    code, out = tk("create", work / "body" / "甲节.docx", "--title", "甲节", "--body",
                   body("b-oldname", [{"text": "我公司服务过旧甲方建设集团及其多个关联项目。"}]))
    assert code == 0, out
    names = WORK / "e103-cwd"                                   # 模拟「跑校验的目录」里放着名单
    names.mkdir(exist_ok=True)
    (names / "旧名名单.txt").write_text("旧甲方建设集团\n", encoding="utf-8")
    code, out = cw("--work", work, cwd=names)
    check("E103a 旧名名单.txt 不被读：名单在场也不报残留（隐形约定退场）",
          code == 0 and "旧名残留" not in out, f"exit={code} {out[:240]}")
    check("E103b 空名单明说残留扫描未执行——不许没查装查过",
          "残留扫描未执行" in out and "[校验通过]" in out, f"exit={code} {out[:240]}")
    code, out = cw("--work", work, "--old-names", "旧甲方建设集团", cwd=names)
    check("E103c --old-names 给了照扫：残留点名出处，且不再提示未执行",
          code == 1 and "旧名残留" in out and "旧甲方建设集团" in out and "残留扫描未执行" not in out,
          f"exit={code} {out[:240]}")


def e104():
    """登记资料库旧名自动载入（2026-11-28 设计讨论拍板「拷贝不改的废标面」）：本次主体
    从承诺清单「投标主体」行读；其他主体的登记名与素材旧名候选=疑似残留（不拦、点破
    「引用性质还是漏改」），自己主体不扫；显式 --old-names 仍是红线级；登记库自动载入
    时不打「残留扫描未执行」，登记库在而主体悬空时点名去确认主体。"""
    work, _ = make_work("w-lib", "# 响应文件：投标文件\n## 目录\n- 目录\n- 甲节\n")
    (work / "body" / "关键事实与承诺.md").write_text(
        "# 关键事实与承诺\n\n事项｜值｜说明\n投标主体｜奥哲｜G2 对话确认\n", encoding="utf-8")
    code, out = tk("create", work / "body" / "甲节.docx", "--title", "甲节", "--body",
                   body("b-lib", [{"text": "本方案沿用星灵旧项目乙的服务体系（编号 XL-2025-BD-0001），"
                                           "并结合奥哲历史项目甲的经验。显式旧名甲的内容也应被点到。"}]))
    assert code == 0, out
    lib = WORK / "资料库-lib"
    lib.mkdir(exist_ok=True)
    (lib / "索引.json").write_text(json.dumps({"materials": [
        {"entity": "奥哲", "entity_display": "奥哲", "file": "a.docx", "kind": "docx",
         "old_names": ["奥哲历史项目甲"]},
        {"entity": "星灵", "entity_display": "星灵创科", "file": "b.docx", "kind": "docx",
         "old_names": ["星灵旧项目乙", "XL-2025-BD-0001"]},
    ]}, ensure_ascii=False), encoding="utf-8")
    libenv = {"TENDER_MATERIAL_HOME": str(lib)}
    code, out = cw("--work", work, "--old-names", "显式旧名甲", env=libenv)
    check("E104a 登记库自动载入：他主体痕迹=疑似残留（不拦）、自己主体不被点名、显式清单仍是红线",
          code == 1 and "疑似" in out and "星灵旧项目乙" in out and "XL-2025-BD-0001" in out
          and "痕迹「奥哲历史项目甲」" not in out and "旧名残留" in out and "显式旧名甲" in out
          and "残留扫描未执行" not in out,
          f"exit={code} {out[:300]}")
    code, out = cw("--work", work, env=libenv)
    check("E104b 只靠登记库也能扫：疑似残留照报、不打「未执行」",
          code == 0 and "疑似" in out and "星灵旧项目乙" in out and "残留扫描未执行" not in out,
          f"exit={code} {out[:300]}")
    code, out = cw("--work", work, env={"TENDER_MATERIAL_HOME": str(WORK / "没有资料库")})
    check("E104c 没有登记库：不载入、按空名单口径明说未执行",
          code == 0 and "星灵旧项目乙" not in out and "残留扫描未执行" in out,
          f"exit={code} {out[:300]}")
    (work / "body" / "关键事实与承诺.md").write_text(
        "# 关键事实与承诺\n\n事项｜值｜说明\n项目名称｜云澜市智慧园区建设工程｜招标文件\n", encoding="utf-8")
    code, out = cw("--work", work, env=libenv)
    check("E104d 主体悬空（承诺清单没有投标主体行）不猜：库在也不载入，明说去确认主体",
          code == 0 and "星灵旧项目乙" not in out and "投标主体" in out,
          f"exit={code} {out[:300]}")


def e105():
    """多册合册 + 登记版式封面：封面节点没有节文件时按册落公司封面。真机 2026-11-28：
    _materialize_cover 在多册形态把「册名/封面.docx」写成字符串相除（str / str），
    首次多册合册就 TypeError 崩——单册形态 E94 盖不住这条路径。"""
    home, spec, out, code = _adopt("多册封面", _header_doc("多册封面源"), source=_header_doc("多册封面源二"))
    outline = ("# 响应文件：资格文件\n## 目录\n- 封面\n- 目录\n- 甲节\n\n"
               "# 响应文件：商务技术文件\n## 目录\n- 封面\n- 目录\n- 乙节\n")
    work, _ = make_work("多册封面", outline, [])
    for vol, leaf in (("资格文件", "甲节"), ("商务技术文件", "乙节")):
        (work / "body" / vol).mkdir(parents=True, exist_ok=True)
        code, out = tk("create", work / "body" / vol / f"{leaf}.docx", "--title", leaf,
                       "--body", body(f"多册封面-{leaf}", [{"text": _LSENT2}]), env=_LE(home))
        assert code == 0, out
    code, out = tk("assemble", "--work", work, env=_LE(home))
    covers = [work / "body" / v / "封面.docx" for v in ("资格文件", "商务技术文件")]
    vols = sorted((work / "body").glob("整本-*.docx"))
    check("E105a 多册合册不崩：两册封面各按册落位（body/<册名>/封面.docx），两本整本都在",
          code == 0 and all(c.is_file() for c in covers) and len(vols) == 2,
          f"exit={code} out={out[:280]}")
    check("E105b 落的是公司封面（内容对得上），且两册都点名待改字段",
          all("智慧园区数据平台项目投标文件" in Document(str(c)).paragraphs[1].text for c in covers)
          and out.count("公司封面") >= 2,
          f"out={out[:280]}")


def e106():
    """删掉整本再合册：指纹记录还在、整本不在——正当处置不许崩。真机 2026-11-28：
    删掉两册带旧名残留的整本后重合册，_book_edited_note 对不存在的文件算 sha256
    抛 FileNotFoundError。文件不在就谈不上「被手工改过」，返回空提醒照常重建。"""
    work, _ = make_work("指纹缺本", "# 响应文件：投标文件\n## 目录\n- 封面\n- 目录\n- 甲节\n",
                        leaves=("甲节",), with_body=True)
    code, out = tk("assemble", "--work", work)
    vol = next((work / "body").glob("整本-*.docx"), None)
    fp = Path(str(vol) + ".fingerprint.json") if vol else None
    check("E106a 首次合册记指纹：整本与 .fingerprint.json 都在",
          code == 0 and vol is not None and fp is not None and fp.is_file(),
          f"exit={code} out={out[:200]}")
    vol.unlink()                                # 删整本、指纹记录留着（真机删的就是带旧名的整本）
    code, out = tk("assemble", "--work", work)
    import hashlib
    vol2 = work / "body" / vol.name
    fp2 = Path(str(vol2) + ".fingerprint.json")
    rec = json.loads(fp2.read_text(encoding="utf-8")) if fp2.is_file() else {}
    h = hashlib.sha256(vol2.read_bytes()).hexdigest() if vol2.is_file() else ""
    check("E106b 整本已删、指纹还在：重合册不崩、整本重建、指纹刷新成新文件的",
          code == 0 and vol2.is_file() and rec.get("sha256") == h,
          f"exit={code} out={out[:240]}")


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
                     ("E50", e50), ("E51", e51), ("E52", e52), ("E53", e53), ("E54", e54),
                     ("E55", e55), ("E56", e56), ("E57", e57), ("E58", e58), ("E59", e59),
                     ("E60", e60), ("E61", e61), ("E62", e62), ("E63", e63), ("E64", e64),
                     ("E65", e65), ("E66", e66),
                     ("E67", e67), ("E68", e68), ("E69", e69),
                     ("E70", e70), ("E71", e71),
                     ("E72", e72), ("E73", e73), ("E74", e74), ("E75", e75), ("E76", e76), ("E77", e77), ("E78", e78),
                     ("E79", e79), ("E80", e80), ("E81", e81), ("E82", e82), ("E83", e83), ("E84", e84),
                     ("E85", e85), ("E86", e86), ("E87", e87), ("E88", e88), ("E89", e89),
                     ("E90", e90), ("E91", e91), ("E92", e92),
                     ("E93", e93), ("E94", e94), ("E95", e95), ("E96", e96),
                     ("E97", e97), ("E98", e98), ("E99", e99),
                     ("E100", e100), ("E101", e101), ("E102", e102), ("E103", e103), ("E104", e104),
                     ("E105", e105), ("E106", e106)):
        run(name, fn)
    print(f"\n{len(PASS)} 通过 / {len(FAIL)} 失败")
    for f_ in FAIL:
        print("  FAIL:", f_)
    if FAIL:
        print(f"  现场保留：{WORK}")
    else:
        shutil.rmtree(WORK, ignore_errors=True)
    sys.exit(1 if FAIL else 0)
