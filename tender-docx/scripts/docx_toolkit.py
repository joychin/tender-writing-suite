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
    inspect    文档检查（只读）：金额验算、表格版式、图片、编号与占位、修订痕迹等，
               出文字结论 + 一张表；不改被检查的文件、不拦流程（也可 --work 查整套的整本）
               （--json 后面跟的是结果落盘路径，不是开关；不给就只打印）
    layout-read 版式规格读取（只读）：把 Word 版历史标书读成版式规格表
               （项目｜取值｜依据｜确定/疑似）——「按这个样式来做」场景的读取端
    layout-adopt 登记公司版式：确认后的规格表登记为本机默认版式，之后建节、合册
               不指定模板就自动按它出（--set 做对话修订）
    layout-show 查看登记的公司版式      layout-forget 取消登记（原规格表留档）

依赖：python-docx（assemble 另需 docxcompose；render-page 需 pypdfium2+Pillow）。
模型不直接改 docx 二进制——内容修改一律走本工具。
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import os
import random
import re
import sys
import time
import unicodedata
from collections import Counter
from datetime import datetime, timezone
from decimal import ROUND_HALF_UP, Decimal
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
    first, *rest = str(msg).split("\n")
    print(f"[失败] {first}")
    for ln in rest:  # 多行消息（如 .doc 的回退指引）：续行缩进对齐，不散落成正文
        print(f"        {ln}")
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


# 结果落盘路径的后缀与它写出的格式对不上就拒：最常见的误用是把要检查的 Word 文件
# 跟在 --json 后面——--json 是「结果写到哪个文件」，不是「按 JSON 输出」的开关。
_DOC_SUFFIXES = {".docx", ".docm", ".doc", ".dotx", ".dotm", ".pdf", ".xlsx", ".xlsm", ".pptx", ".ppsx"}


def _work_docx(work) -> list[Path]:
    """工作区 body/ 下的全部 Word 文件（整本与各册节文件），跳过 Word/WPS 的占用锁文件。
    落盘防呆拿它当「本次会读到的文件」——比 inspect 自己枚举的宽，宁可多挡。"""
    body = Path(work) / "body"
    if not body.is_dir():
        return []
    # 只认文件：工作区里有按源文件名命名的目录，名字本身就以 .docx 结尾
    return [p for p in sorted(body.rglob("*.docx"))
            if p.is_file() and not p.name.startswith(("~$", ".~"))]


def guard_dump_path(out, inputs=(), suffix: str = ".json", flag: str = "--json", hint: str = "") -> None:
    """结果落盘前的防呆（真机事故：模型把 `inspect --json 标书.docx` 当成「按 JSON 输出
    检查结果」——文件参数被 --json 吃掉，检查了 0 份文件，然后把一份真实标书写成了 63
    字节的报告）。写之前拦，两条都拒（退出码 2）：
    ① 落盘路径与本次要读的文件是同一个——写进去就是覆盖它；
    ② 后缀不是这个参数写出的格式（--json 须 .json、--out 须 .md）——要读的 Word 文件
    直接跟在命令后面，不是跟在 --json 后面。
    调用方一律在读文件之前调：拦下时什么都还没做。hint 是正确写法的示例，进报错。"""
    o = Path(out)
    tail = f"，如 {hint}" if hint else ""
    for p in inputs:
        try:
            same = o.resolve() == Path(p).resolve() or (o.exists() and os.path.samefile(o, p))
        except OSError:
            same = False
        if same:
            die(f"{flag} 要写入的文件和要读取的文件是同一个，写进去就把它覆盖了：{p}"
                f"——换一个 {suffix} 路径{tail}", 2)
    if o.suffix.lower() == suffix:
        return
    if o.suffix.lower() in _DOC_SUFFIXES:
        die(f"{flag} 后面跟的是结果落盘的路径，不是要读取的文件：{o.name} 是 Word/PDF 文档，"
            f"写进去就把它覆盖了。要读取的文件直接跟在命令后面{tail}", 2)
    die(f"{flag} 的落盘路径须以 {suffix} 结尾（它不是开关）：{o}{tail.replace('，如', '。如')}", 2)


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
        if any(k in sname for k in _HEADING_FONT_KEYS):
            face = _lay("headings", "1", "ea", default=None) or "黑体"   # 登记了公司版式：标题的中文跟版式
        else:
            face = _lay("body", "ea", default=None) or "宋体"
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
    """正文段首行缩进（firstLineChars 而非固定磅，随字号缩放）。登记了公司版式按版式的缩进来；
    「无」也显式写零——合册收尾只给没有 w:ind 的段补缩进。"""
    ind = p._p.get_or_add_pPr().get_or_add_ind()
    spec = _lay("body", "ind")
    if spec is None:
        ind.set(qn("w:firstLineChars"), "200")
        if ind.get(qn("w:firstLine")):
            del ind.attrib[qn("w:firstLine")]
        return
    for a in ("w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]
    if "firstChars" in spec:
        ind.set(qn("w:firstLineChars"), str(spec["firstChars"]))
    else:
        ind.set(qn("w:firstLine"), str(spec.get("first", 0)))
        ind.set(qn("w:firstLineChars"), "0")


def _clear_first_line_indent(p) -> None:
    """清掉段落继承的首行缩进（表格单元格/标题/题注用）。"""
    p.paragraph_format.first_line_indent = Pt(0)
    ind = p._p.get_or_add_pPr().find(qn("w:ind"))
    if ind is not None and ind.get(qn("w:firstLineChars")):
        del ind.attrib[qn("w:firstLineChars")]


# 内置标题族：一级 16 磅、二级与三级 14 磅、四级 12 磅，黑体加粗。基准模板的 setup_styles 与
# 用户模板缺样式时的补建（_ensure_builtin_styles）共用这一份定义
_HEADING_SPECS = (("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 14), ("Heading 4", 12))


def _apply_heading_spec(st, size: int) -> None:
    st.font.name = "黑体"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = None
    set_east_asia(st.element, "黑体")


def setup_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"  # 西文/数字按标书惯例用 Times，汉字走 eastAsia 宋体
    normal.font.size = Pt(12)
    set_east_asia(normal.element, "宋体")
    # 注意：首行缩进不放 Normal——表格单元格/题注/目录条目都会继承，
    # 窄列里会退化成一字一行。缩进由 _set_body_indent 按正文段逐段设置。
    normal.paragraph_format.line_spacing = 1.5

    for name, size in _HEADING_SPECS:
        _apply_heading_spec(doc.styles[name], size)

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


def _ensure_builtin_styles(doc: Document) -> list[str]:
    """文档缺标准样式时补建（编号/目录依赖它们），返回补建了哪些，由调用方回显。
    标题样式补的是内置定义（黑体加粗、按级 16/14/14/12 磅）——只建一个空样式的话章标题
    跟着正文走、和正文一个样。真实历史标书多是没有标题样式的，拿它当模板时 SKILL 里写的
    「缺则回落内置样式并点名」靠这里兑现（2026-11-27 真机：实测只补了空壳、也没回显）。
    Normal 缺了只建空壳：它是别的样式的基样式，字体字号由模板的文档默认值决定。"""
    added: list[str] = []
    sizes = dict(_HEADING_SPECS)
    for name in ("Normal",) + tuple(sizes):
        try:
            doc.styles[name]
        except KeyError:
            try:
                st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
            except Exception:  # noqa: BLE001
                continue
            if name in sizes:
                _apply_heading_spec(st, sizes[name])
            added.append(name)
    return added


def _template_styles_note(label: str, added: list[str]) -> str:
    """用户模板补建了样式时的回显（给模型读，由它转达用户）。"""
    heads = [n for n in added if n != "Normal"]
    if not heads:
        return f"[模板标题样式] 模板 {label} 里没有 Normal 样式：已补建（正文段落都挂在它上面）"
    return (f"[模板标题样式] 模板 {label} 里没有 {'、'.join(heads)}：已按内置样式补建"
            "（标题为黑体加粗，一级 16 磅、二级与三级 14 磅、四级 12 磅）——这几级标题不跟模板的"
            "正文设置走。要按公司自己的样式，请先在模板里定义这些标题样式"
            + ("；另补建了 Normal" if "Normal" in added else ""))


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


_CORE_PART = "/docProps/core.xml"
_CORE_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
_DC_NS = "{http://purl.org/dc/elements/1.1/}"
_DCTERMS_NS = "{http://purl.org/dc/terms/}"
_XSI_TYPE = "{http://www.w3.org/2001/XMLSchema-instance}type"
_APP_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
_APP_RESIDUE_TAGS = ("Application", "AppVersion", "Company", "Manager", "Template", "TotalTime")


def _props_parts(doc: Document) -> dict:
    """docProps 各部件（core.xml / app.xml / custom.xml），没有就没有。python-docx 只给
    API 管 core，app.xml 与 custom.xml 一直以裸 blob 随包走——Word 应用名、公司、模板名
    与 WPS 的 ICV（云账号身份码）就藏在里面。"""
    want = {_CORE_PART, "/docProps/app.xml", "/docProps/custom.xml"}
    return {str(p.partname): p for p in doc.part.package.iter_parts() if str(p.partname) in want}


def _scrub_metadata(doc: Document) -> None:
    """落盘前把元数据清零（用户拍板：保存的时候不存入任何 metadata）——幂等，挂
    统一落盘口 _save_doc 与 base_document 两道：
    ①core 全字段清空（含标题/主题/关键词/类别——交付稿的属性面板全空，时间写当前），
      contentStatus/language/identifier/version/lastPrinted 整元素摘除；
    ②app.xml 摘掉应用名/版本/公司/管理者/模板名/总编辑时长（基准模板自带
    「Microsoft Macintosh Word」，改历史标书还会带进原单位与原模板名）；
    ③docProps/custom.xml 整体摘除（WPS 在这里写 ICV 云账号身份码）；
    ④包级首页缩略图摘除（基准模板自带一张 Word 截图，随文件交出去）。
    摘部件 = 摘包级关系：部件不可达即不落盘，[Content_Types].xml 落盘时按实存部件重建。
    ⚠ core 一律直改 core.xml 元素树、不走 python-docx 属性 API：docxcompose 在 import
    时就把全局 nsmap["cp"] 改成错误命名空间（docxcompose/utils.py 的 nsmap.update），
    同进程里属性 setter 会建出命名空间错误的重复元素，脏值藏在里面、属性 API 读不到
    也清不掉；顺带把历史文件里这类异命名空间同名的残留按本地名摘掉。"""
    try:
        from lxml import etree

        core = _props_parts(doc).get(_CORE_PART)
        if core is not None:
            root = (core._element if getattr(core, "_element", None) is not None
                    else etree.fromstring(core.blob))
            now = datetime.now(timezone.utc).replace(microsecond=0,
                                                     tzinfo=None).strftime("%Y-%m-%dT%H:%M:%SZ")
            for ns, tag, text in (
                    (_DC_NS, "title", ""), (_DC_NS, "subject", ""), (_DC_NS, "creator", ""),
                    (_CORE_NS, "keywords", ""), (_DC_NS, "description", ""),
                    (_CORE_NS, "lastModifiedBy", ""), (_CORE_NS, "revision", "1"),
                    (_DCTERMS_NS, "created", now), (_DCTERMS_NS, "modified", now),
                    (_CORE_NS, "category", "")):
                for el in root.findall(ns + tag):
                    el.text = text
                    if ns == _DCTERMS_NS:
                        el.set(_XSI_TYPE, "dcterms:W3CDTF")
            for tag in ("contentStatus", "language", "version", "lastPrinted"):
                for el in root.findall(_CORE_NS + tag) + root.findall(_DC_NS + tag):
                    root.remove(el)
            ok = {_CORE_NS, _DC_NS, _DCTERMS_NS}
            names = {"title", "subject", "creator", "keywords", "description",
                     "lastModifiedBy", "revision", "created", "modified", "category",
                     "contentStatus", "language", "version", "lastPrinted"}
            for el in list(root):
                q = etree.QName(el)
                if q.localname in names and q.namespace not in ok:
                    root.remove(el)
            if getattr(core, "_element", None) is None:  # 裸 blob 部件手动回写；XmlPart 由序列化器落盘
                core._blob = etree.tostring(root, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)
    except Exception:  # noqa: BLE001
        pass
    try:
        app = _props_parts(doc).get("/docProps/app.xml")
        if app is not None:
            from lxml import etree

            root = (app._element if getattr(app, "_element", None) is not None
                    else etree.fromstring(app.blob))
            for tag in _APP_RESIDUE_TAGS:
                for el in root.findall(_APP_NS + tag):
                    root.remove(el)
            if getattr(app, "_element", None) is None:  # 裸 blob 部件手动回写；XmlPart 由序列化器落盘
                app._blob = etree.tostring(root, xml_declaration=True,
                                           encoding="UTF-8", standalone=True)
    except Exception:  # noqa: BLE001
        pass
    try:
        rels = doc.part.package.rels
        for rid in [rid for rid, rel in list(rels.items())
                    if rel.reltype.rsplit("/", 1)[-1] in ("thumbnail", "thumbnail-picture",
                                                          "custom-properties")]:
            rels.pop(rid, None)
    except Exception:  # noqa: BLE001
        pass


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
        added = _ensure_builtin_styles(doc)
        if _LAYOUT is not None and template is None:
            apply_layout(doc, _LAYOUT)     # 登记了公司版式：页面与样式从内置基准模板现编译
        notes = ensure_cjk_theme(doc)
        _scrub_metadata(doc)
        if notes and user_template:
            print("[模板中文字体] " + "；".join(notes)
                  + f"（模板 {path.name} 的主题字体语言会顶掉显式中文字体，已归一）")
        if added and user_template:
            print(_template_styles_note(path.name, added))
        return doc, user_template
    if user_template:
        die(f"模板文件不存在：{path}")
    doc = Document()
    _scrub_metadata(doc)
    setup_styles(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.18)
    return doc, False


_PAGE_W = {"plain": ("", ""), "x页": ("第 ", " 页"), "dash": ("— ", " —")}


def _page_field(paragraph, pn: dict | None = None) -> None:
    """页码域。pn 是版式里的页码写法（写法／是否带共 Y 页／字号），None 是内置口径（纯数字、9 磅）。"""
    pn = pn or {}
    kind = pn.get("style") or "plain"
    pieces: list = []
    field = lambda instr: [("w:fldChar", {"w:fldCharType": "begin"}, None),   # noqa: E731
                           ("w:instrText", {"xml:space": "preserve"}, f" {instr} "),
                           ("w:fldChar", {"w:fldCharType": "end"}, None)]
    if pn.get("total") and kind == "x页":      # 「第 X 页、共 Y 页」
        pieces += [("w:t", {}, _PAGE_W[kind][0]), *field("PAGE"), ("w:t", {}, "、共 "), *field("NUMPAGES"),
                   ("w:t", {}, _PAGE_W[kind][1])]
    else:
        pieces += [("w:t", {}, _PAGE_W[kind][0]), *field("PAGE"), ("w:t", {}, _PAGE_W[kind][1])]
    for tag, attrs, text in (pc for pc in pieces if pc[0] != "w:t" or pc[2]):
        run = paragraph.add_run()
        if tag == "w:t":
            run.text = text
            continue
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(qn(k), v)
        if text is not None:
            el.text = text
        run._element.append(el)
    for run in paragraph.runs:
        style_run(run, "Times New Roman", pn.get("pt") or 9)


def _apply_registered_header(master: Document) -> None:
    """登记的公司页眉落到整本：logo 图（按登记的原尺寸）＋文字＋横线，前置节不带（合册
    收尾剥引用）。页码走页码写法那条线，与页眉各占一段。"""
    hd = (_LAYOUT or {}).get("assets", {}).get("header") if _LAYOUT is not None else None
    if not hd:
        return
    from docx.shared import Emu

    sec = master.sections[0]
    h = sec.header
    h.is_linked_to_previous = False
    for p in list(h.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p = h.paragraphs[0] if h.paragraphs else h.add_paragraph()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    for im in hd.get("images") or []:
        path = layout_home() / im["file"]
        if not path.is_file():
            continue
        run = p.add_run()
        try:
            run.add_picture(str(path), height=Emu(im["cy"]) if im.get("cy") else None,
                            width=Emu(im["cx"]) if im.get("cx") else Cm(1.5))
        except Exception:  # noqa: BLE001
            continue
    if hd.get("text"):
        _run_fmt(p.add_run(hd["text"]), hd)
    p.alignment = {"两端": WD_ALIGN_PARAGRAPH.JUSTIFY, "居中": WD_ALIGN_PARAGRAPH.CENTER,
                   "右": WD_ALIGN_PARAGRAPH.RIGHT, "分散": WD_ALIGN_PARAGRAPH.DISTRIBUTE}.get(
                       hd.get("align"), WD_ALIGN_PARAGRAPH.CENTER)
    if hd.get("border"):
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        pBdr.append(bottom)
        p._p.get_or_add_pPr().append(pBdr)


def _materialize_cover(body_dir: Path, vol_name: str, multi: bool) -> Path:
    """登记的公司封面落成封面节文件（统一落盘口：清元数据＋记指纹），之后它就是普通
    节文件——模型按批注把项目字段改掉，重新合册不再重复落。"""
    src = layout_home() / ((_LAYOUT.get("assets") or {}).get("cover") or "")
    sf = body_dir / (sanitize_name(vol_name) / "封面.docx" if multi else "封面.docx")
    sf.parent.mkdir(parents=True, exist_ok=True)
    _save_doc(Document(str(src)), sf)
    return sf


def enable_footer_pagenum(doc: Document, different_first: bool) -> None:
    sec = doc.sections[0]
    sec.different_first_page_header_footer = different_first
    if _LAYOUT is not None and "pagenum" in _LAYOUT["values"] and _LAYOUT["values"]["pagenum"] is None:
        # 版式不编页码：基准模板页脚里的页码域一并摘掉
        for ref in sec._sectPr.findall(qn("w:footerReference")):
            sec._sectPr.remove(ref)
        return
    pn = _lay("pagenum", default={}) if _LAYOUT is not None else {}
    hf = sec.header if pn.get("where") == "页眉" else sec.footer
    if hf is sec.header and any(pp.runs for pp in hf.paragraphs):
        p = hf.add_paragraph()   # 公司页眉已占第一段：页码另起一段，不清掉公司内容
    else:
        p = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
    p.alignment = {"居中": WD_ALIGN_PARAGRAPH.CENTER, "左": WD_ALIGN_PARAGRAPH.LEFT}.get(
        pn.get("pos"), WD_ALIGN_PARAGRAPH.RIGHT)
    # 幂等：先清空段内既有内容再加域（基准模板已带页码域，重复添加会渲染成 11/22）
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    _page_field(p, pn)
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
    _save_doc(doc, path)
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


def _lay_table() -> dict:
    """出表口径：登记了公司版式按版式（没读到的项回落内置），否则内置——全线 0.5 磅、表头灰底
    加粗居中、宋体五号。各键：ea／ascii／pt／align（数据行）、hb／ha／shade（表头行）、border。"""
    return _lay_table_of(_lay("table", default={}) or {})


def _tbl_borders_el(spec: dict):
    """表格框线（tblBorders 子元素按 schema 定序：top→left→bottom→right→insideH→insideV）。"""
    kind, sz = spec.get("kind", "all"), int(spec.get("sz") or 4)
    on = {"all": {"top", "left", "bottom", "right", "insideH", "insideV"}, "none": set(),
          "three": {"top", "bottom"}, "partial": set(spec.get("sides") or [])}.get(kind, set())
    if kind == "three" and spec.get("inner"):
        on = on | {"insideH"}
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        if side in on:
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), str(sz))
            b.set(qn("w:color"), "000000")
        else:
            b.set(qn("w:val"), "nil")
        borders.append(b)
    return borders


def _make_tr(specs: list[tuple[str, int, str | None]], col_w: list[int], *,
             bold: bool, size: float, header: bool, shade: str | None) -> "OxmlElement":
    lt = _lay_table()
    jc_val = _ALIGN_W.get(lt["ha"] if header else lt["align"]) or "left"
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
        jc = OxmlElement("w:jc")   # 每格各建一个：同一个元素对象往第二格里追加会被搬走，前一格就丢了
        jc.set(qn("w:val"), jc_val)
        ppr.append(jc)
        p.append(ppr)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), lt["ascii"])
        rfonts.set(qn("w:eastAsia"), lt["ea"])
        rpr.append(rfonts)
        if bold or (header and lt["hb"]):
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
    cap_below, cp = False, None
    if cap:
        cap_below = (_lay("caption", "table", default={}) or {}).get("above") is False   # 版式：表题在表下方
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
    base_pt = _lay_table()["pt"]
    size = base_pt if grid_w <= 6 else (max(base_pt - 1, 7.5) if grid_w <= 8 else max(base_pt - 1.5, 7))
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
    tblPr.append(_tbl_borders_el(_lay_table()["border"]))
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")  # 固定布局，杜绝 Word 自适应挤压
    tblPr.append(layout)
    if compact or (_lay_table()["mar"] or 108) != 108:
        mar = OxmlElement("w:tblCellMar")
        wide = _lay_table()["mar"] not in (None, 108)   # 版式读到非缺省边距：所有表都按它（优先于紧凑口径）
        for side, v in (("top", "15"), ("left", str(_lay_table()["mar"]) if wide else "60"),
                        ("bottom", "15"), ("right", str(_lay_table()["mar"]) if wide else "60")):
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
        tbl.append(_make_tr(hspecs, col_w, bold=False, size=size,
                            header=True, shade=_lay_table()["shade"]))   # 表头加不加粗由版式的「表头样式」定
    for specs in rspecs:
        tbl.append(_make_tr(specs, col_w, bold=False, size=size,
                            header=False, shade=None))
    _body_insert_before_sectpr(doc, tbl)
    if cp is not None and cap_below:
        tbl.addnext(cp._p)   # 版式说表题在表下方：题注挪到表后（表后空段之前）
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
            _save_doc(doc, args.path)
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
        _save_doc(doc, args.path)
        print(f"[已修订] {args.path}（{args.cell} ← {args.text[:30]}）")
        return
    para = _find_para(doc, args.para)
    if args.mode == "delete":
        para._p.getparent().remove(para._p)
        n_orphan = _purge_orphan_bodies(doc)
        _save_doc(doc, args.path)
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
        _save_doc(doc, args.path)
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
        _save_doc(doc, args.path)
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
    _save_doc(doc, args.path)
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
    _save_doc(doc, args.path)
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
    _save_doc(doc, args.path)
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
        # 「页宽 − 左右边距」算出来是普通整数（Length 的减法不返回 Length），没有 .cm——
        # 以前写 .cm 那一步总是抛异常被下面的 except 吞掉，图宽一直是兜底的 15 厘米：
        # A4 页左右边距 3.18 厘米时正文区只有 14.64 厘米，每张图都比正文和表格宽出 0.36 厘米
        content_cm = int(sec.page_width - sec.left_margin - sec.right_margin) / 360000
    except Exception:  # noqa: BLE001
        pass
    para = doc.add_paragraph()
    center = _lay("image", "center", default=True)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.first_line_indent = Pt(0)
    if anchor is not None:
        anchor._p.addnext(para._p)
    run = para.add_run()
    stream = _normalized_image(img)
    # 图宽：登记了公司版式且读到占版心的比例就按比例（多数历史标书的图不满版心），否则满版心
    content_cm *= (_lay("image", "ratio", default=1.0) or 1.0)
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
        # 图题优先用图题样式（登记了公司版式时表题与图题可以不一样：表题居左、图题居中）
        cap_style = _get_style(doc, "Figure Caption") or _get_style(doc, "Table Caption")
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
    _save_doc(doc, args.path)
    where = "全宽居中" if center else "全宽居左"
    if _lay("image", "ratio", default=1.0) not in (None, 1.0):
        where = f"{'居中' if center else '居左'}、宽 {content_cm:.1f} 厘米（按公司版式占版心 {_lay('image', 'ratio'):.0%}）"
    print(f"[已插图] {args.path}（{img.name}，{where}{note}）")


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
    _save_doc(doc, args.docx)
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
    _save_doc(doc, docx_path)
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
    # 素材新鲜度闸：指引素材列记的「文件名+行号区间」是解析时登记的位置，写作逐节进行、
    # 素材文件是活的——解析之后被改过，还没拷的节照旧区间拷错且无声进标书（拷过的节不受影响，
    # 同一素材不同区间可分派多节）。指纹对不上就停：重解析 + 重核指引区间，宁停勿错。
    meta_path = parse_dir / (src_path.stem + ".meta.json")
    if meta_path.is_file():
        try:
            rec = json.loads(meta_path.read_text(encoding="utf-8"))
            rec_sha = rec.get("sha256") if isinstance(rec, dict) else None
        except Exception:  # noqa: BLE001
            rec_sha = None
        if rec_sha and rec_sha != _sha256(src_path):
            die("源文件在解析之后被改过（内容指纹与解析登记不一致）——素材列记的行号区间已不可信，本次不拷。\n"
                "处置：先重新解析这份素材（parse_document.py <源文件> -o <解析目录>，解析是幂等的），\n"
                "再核对写作指引素材列的行号区间是否仍指着要的内容，然后重跑 copy-from", code=2)
    else:
        print("  ⚠️ 解析产物没记源文件指纹（旧版产物），无法核对素材是否在解析后被改——建议重新解析后再拷")
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
    _save_doc(tgt_doc, tgt_path)
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

def _num_gap(depth: int) -> str:
    """编号与标题之间的空白：登记了公司版式且读到「编号间隔」就按它（全角空格／半角空格／紧接）；
    否则内置口径——章（第X章）后全角空格，其余层级半角空格。"""
    sep = _lay("numbering", "sep", default=None)
    return sep if sep is not None else ("\u3000" if depth == 1 else " ")


def _chap_new_page() -> bool:
    """每章是否另起一页：登记了公司版式按版式的「章起新页」，否则起新页。"""
    return bool(_lay("chapter_new_page", default=True))


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
        return f"第{_cn(counters[1])}章{_num_gap(1)}" if depth == 1 else \
            ".".join(str(counters[k]) for k in range(1, depth + 1)) + _num_gap(depth)
    if scheme == "gov":
        return _gov_label(depth, counters)
    # decimal
    return ".".join(str(counters[k]) for k in range(1, depth + 1)) + _num_gap(depth)


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
    # 编号方案与目录层级：用户（或招标文件的要求）明说了按明说的；没说就跟登记的公司版式；都没有用内置
    numbering = args.numbering or _lay("numbering", "scheme", default=None) or "chapter"
    toc_depth = args.toc_depth or _lay("toc", "depth", default=None) or 3
    body_hdr_pn = body_ftr_pn = None    # 主文档自己的页眉/页脚部件（合并后把正文节引用拨回来）

    for vol in volumes:
        vol["tree"] = build_tree(vol["nodes"])

    master = None
    composer = None
    toc_entries: list[tuple[str, int, bool]] = []  # (编号+标题, 缩进层级, 是否另附)

    for vol in volumes:
        used_files: set[Path] = set()  # 每册重置——防上一册节文件混进本册批注迁移/孤儿判定
        out_name = f"整本-{sanitize_name(vol['name'])}.docx"
        edited = _book_edited_note(body_dir / out_name)
        if edited:
            print(f"  ⚠️ {edited}")
        master, user_tpl = base_document(args.template)
        if not user_tpl:
            _apply_registered_header(master)
            enable_footer_pagenum(master, different_first=True)
            hdr = master.sections[0].header
            ftr = master.sections[0].footer
            body_hdr_pn = str(hdr.part.partname) if not hdr.is_linked_to_previous else None
            body_ftr_pn = str(ftr.part.partname) if not ftr.is_linked_to_previous else None
        composer = Composer(master)
        counters: dict[int, int] = {}
        chapter_plan: list[tuple[str, int, str, bool]] = []  # (title, depth, num, in_front)
        landscape_titles: set[str] = set()  # 本册横向节文件的章标题（合册后补分节重排）
        container_files: set[Path] = set()  # 容器章同名节文件（容器分支已点名，孤儿扫描不再重复报）
        taken_ids: set[str] = set()  # 本册已占用的批注号（跨节撞号就地重排）
        seen_chapter = False  # 首章前一律分页（版式说「章不起新页」也不让第一章贴着目录）
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
            nonlocal in_front, toc_marker, merged, taken_ids, seen_chapter
            title = node["text"]
            if title == "封面":
                sf = _section_file(body_dir, vol["name"], title, multi)
                if sf and sf.is_file():
                    composer.append(Document(str(sf)))
                    used_files.add(sf)
                    _strip_first_heading(master, title)
                else:
                    cover_tpl = (_LAYOUT or {}).get("assets", {}).get("cover") if _LAYOUT is not None else None
                    if cover_tpl:
                        sf = _materialize_cover(body_dir, vol["name"], multi)
                        composer.append(Document(str(sf)))
                        used_files.add(sf)
                        _strip_first_heading(master, title)
                        fields = _lay("cover_fields", default=[]) or []
                        warnings.append(
                            f"封面用了登记的公司封面（已落为 {sf.as_posix()}）——里面的项目名、投标人、日期"
                            f"还是历史标书上的；开写时先把这些字段改成本次项目的"
                            f"{'：' + '、'.join(fields) if fields else '（规格表封面字段行没读到，逐行核一遍）'}")
                    else:
                        _append_simple_cover(master, vol["name"])
                merged += 1
                return
            if title == "目录":
                in_front = False
                toc_marker = _append_toc_marker(master)
                merged += 1
                return
            num = "" if in_front else _number_scheme_leaf(depth, counters, numbering)
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
                                          page_break=(depth == 1 and bool(num)
                                                      and (_chap_new_page() or not seen_chapter)))
                if depth == 1 and num:
                    seen_chapter = True
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
                    _save_doc(sub_doc, sf)
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
                                                     or (depth == 1 and not in_front and bool(num)
                                                         and (_chap_new_page() or not seen_chapter))):
                if exists:
                    warnings.append(f"节文件内标题与目录不一致（目录：{title}）——编号未应用，"
                                    f"改节文件标题与目录逐字一致后重新合册")
            if depth == 1 and not in_front and num:
                seen_chapter = True
            chapter_plan.append((title, depth, num, in_front,
                                 str(counters[1]) if counters.get(1) else ""))
            toc_entries.append(((num + title).strip(), depth, not exists))
            merged += 1

        for top in vol["tree"]:
            _emit_node(top, 1)
        if not user_tpl:
            _repoint_body_refs(master, body_hdr_pn, body_ftr_pn)
        toc_end = None
        if toc_marker is not None:
            toc_end = _fill_toc(master, toc_marker, toc_entries, toc_depth)
        _restyle_all(master, chapter_plan, numbering, args.flat_outline)
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
        _save_doc(master, out_path, "assemble")
        print(f"[已合册] {out_path}（合并 {merged} 节，编号 {numbering}，"
              f"批注 {n_comments} 条待处理）")
        if font_notes:
            print("  [字体固化] " + "；".join(font_notes))
        print("  对用户口径：整本 Word 文档已生成（「合册」是内部动作，"
              "不进对话、不作提问——reply-style 硬规则 9）。")
        for w in warnings:
            print(f"  ⚠️ {w}")
        warnings = []
        toc_entries = []


def _repoint_body_refs(master: Document, hdr_pn: str | None, ftr_pn: str | None) -> None:
    """docxcompose 合并会把节文件自带的页眉页脚部件一并导入并重排引用，正文节（终节）的
    默认引用可能落到节文件带进来的空部件上（真实回放：公司封面节文件合册后，正文页码指向
    封件带进来的空页脚，页码整本消失）。收尾把默认引用拨回主文档自己的页眉/页脚部件。"""
    final = master.element.body.find(qn("w:sectPr"))
    if final is None:
        return

    def rid_of(partname: str | None):
        if not partname:
            return None
        for rid, rel in master.part.rels.items():
            if rel.is_external:
                continue
            try:
                if str(rel.target_part.partname) == partname:
                    return rid
            except Exception:  # noqa: BLE001
                continue
        return None

    for tag, want in ((qn("w:headerReference"), rid_of(hdr_pn)),
                      (qn("w:footerReference"), rid_of(ftr_pn))):
        if want is None:
            continue
        for ref in final.findall(tag):
            if ref.get(qn("w:type")) == "default":
                ref.set(qn("r:id"), want)


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
    """目录占位标题段（条目在全部节合并后回填）；段前分页——目录不与封面同页。
    字体字号加粗对齐跟登记的公司版式（目录标题那一行），没登记是内置的黑体三号加粗居中。"""
    from docx.text.paragraph import Paragraph

    h = OxmlElement("w:p")
    _body_insert_before_sectpr(master, h)
    hp = Paragraph(h, master)
    hp.paragraph_format.first_line_indent = Pt(0)
    hp.paragraph_format.page_break_before = True
    r = hp.add_run("目录")
    tt = _lay("toc", "title", default=None)
    if tt:
        _run_fmt(r, tt)
        hp.alignment = {"两端": WD_ALIGN_PARAGRAPH.JUSTIFY, "居中": WD_ALIGN_PARAGRAPH.CENTER,
                        "右": WD_ALIGN_PARAGRAPH.RIGHT, "分散": WD_ALIGN_PARAGRAPH.DISTRIBUTE}.get(
                            tt.get("align"), WD_ALIGN_PARAGRAPH.LEFT)
    else:
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
    # 引导线跟登记的公司版式（点线／短线／下划线／无）；制表位落在版心右缘——登记了版式按版心的页边距算，
    # 没登记是内置的 14.4 厘米
    from docx.shared import Emu
    leader = _lay("toc", "leader", default="dot") or "dot"
    leader_enum = {"dot": WD_TAB_LEADER.DOTS, "hyphen": WD_TAB_LEADER.DASHES,
                   "underscore": WD_TAB_LEADER.LINES, "heavy": WD_TAB_LEADER.HEAVY}.get(leader, WD_TAB_LEADER.SPACES)
    if _LAYOUT is not None:
        sec = master.sections[0]
        tab_pos = Emu(int(sec.page_width - sec.left_margin - sec.right_margin) - 90000)
    else:
        tab_pos = Cm(14.4)

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
        entry.paragraph_format.tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT, leader_enum)
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
        parts = []
        for k in range(depth + 1, target + 1):
            sub_counters.setdefault(k, 1)   # 跳级（节文件直接用 h3、没经过 h2）：缺的中间层级按 1 续
            parts.append(str(sub_counters[k]))
        sub = ".".join(parts)
        prefix = arab if (depth == 1 and scheme == "chapter" and arab) else num.strip()
        _rewrite_para_text(p, f"{prefix}.{sub}{_num_gap(depth)}{p.text}" if sub else f"{prefix}{_num_gap(depth)}{p.text}")


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


def _lo_convert(soffice: str, docx_path: Path, outdir: Path, profile: Path,
                target: str = "pdf", stuck=None) -> Path:
    """LibreOffice 无界面转换；超时与起不来都在这里收口（超时后由 _run_lo 把进程树杀干净）。
    stuck 是 .doc 通道给的出口——一个函数，收到「这次怎么回事」的一句话（如「 超过 300 秒
    没转完」），返回给模型的完整提示（请用户自己另存）；没给的调用方（目录页码实测）没有
    等价替代，用缺省文案。"""
    # 配置目录地址必须由 as_uri 生成：手工拼 file://+路径 在 Windows 上是坏地址（盘符
    # 冒号、反斜杠），路径含空格/# 也不转义——地址坏了 LibreOffice 就退回默认配置，
    # 界面版开着时转换被转交给它、本进程拿不到产物
    cmd = [soffice, "--headless", f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
           "--convert-to", target, "--outdir", str(outdir), str(docx_path)]

    def give_up(what: str, detail: str = "", timed_out: bool = False) -> None:
        if stuck is not None:
            die(stuck(what), 2)
        die(f"LibreOffice 转 {target} 超时（{_LO_TIMEOUT}s）——可跳过本步稍后再试" if timed_out
            else f"LibreOffice 转 {target} 失败：{detail[:200]}")

    try:
        r = _run_lo(cmd, _LO_TIMEOUT)
    except _LoTimeout:
        give_up(f" 超时（超过 {_LO_TIMEOUT} 秒没转完）", timed_out=True)
    except OSError as exc:
        give_up(f" 失败（LibreOffice 起不来：{exc}）", f"起不来（{exc}）")
    out = outdir / (docx_path.stem + "." + target)
    if not out.is_file():
        detail = (r.stdout or r.stderr).strip()[:200] or "没产出文件，也没有任何输出"
        give_up(f" 失败（{detail}）", detail)
    return out


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
        _save_doc(doc, docx_path)
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


# ---------------------------------------------------------------- 文档检查（inspect，只读）

# 检查只负责发现：不改被检查的文件、不拦流程（--json 是显式的结果落盘，写之前过防呆）。
# 结果 = 一行结论 + 一张表（位置／问题／依据／
# 确定还是疑似）+ 定位行 + 已检查范围，用户要求修订时才动文件。「确定」= 判法是算术或
# 结构事实（合价对不上、图片字节损坏）；「疑似」= 靠估算，要人看一眼（列宽折行、分辨率）。
# 最危险的失败不是误报，是不声不响地没查：读不出的金额格、认出是报价表却一行都没验算成，
# 都要点名——两张真实报价表回放时，一个写错的量词正则曾让整行验算静默跳过且不报。
# 坐标口径与 read 视图一致：P# 只数正文层段落；T#R#C# 的 C 是网格列（合并格取起始列）。

FP_SUFFIX = ".fingerprint.json"   # 整本指纹记录：<整本名>.docx.fingerprint.json
_AMT_TOL = Decimal("0.01")
_AMT_MULT = {"万": 10 ** 4, "千": 10 ** 3, "百": 10 ** 2, "亿": 10 ** 8}
_CN_DIGIT = {"零": 0, "〇": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9,
             "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
             "兩": 2, "參": 3, "陸": 6, "柒": 7, "捌": 8, "玖": 9}
_CN_NUM_CHARS = "零〇一二三四五六七八九十百千万亿壹贰叁參肆伍陆陸柒捌玖拾佰仟萬億"  # 含「叁」：
    # 真实一览表「人民币叁佰柒拾玖万伍仟元整」回放时因缺这个字被判成读不出来
_CN_DIG = "零〇一二三四五六七八九壹贰參肆伍陆陸柒捌玖"
_TOTAL_KW = "投标总价|报价总额|总报价|响应报价|投标报价|初次报价|最后报价|最终报价"
# 小写／大写金额：冒号可有可无（「小写）人民币…」「大写 660000 元」都认），值取到同行末或另一个词前
_LOWER_RE = re.compile(r"小\s*写\s*[)）]?\s*[：:]?\s*([^\n]{1,40}?)(?=\s*[(（]?\s*大\s*写|\s*$)", re.M)
_UPPER_RE = re.compile(r"大\s*写\s*[)）]?\s*[：:]?\s*([^\n]{1,40}?)(?=\s*[(（]?\s*小\s*写|\s*$)", re.M)
# 「人民币壹拾万元整（¥100,000.00）」：大写在前、括号里是小写
_UP_THEN_NUM = re.compile("([" + _CN_NUM_CHARS + "]{2,}(?:元|圆)(?:[" + _CN_DIG + "]?[角分])?(?:整|正)?)"
                          r"\s*[(（]\s*[¥￥]?\s*([\d,，]+(?:\.\d+)?)\s*([万千百亿])?\s*元?\s*[)）]")
_KWD_NUM = re.compile("(?:" + _TOTAL_KW + r")\s*[(（]?[^)）\n：:]{0,10}[)）]?\s*(?:金额)?\s*(?:为|是|合计|[:：])"
                      r"\s*(?:人民币)?\s*[¥￥]?\s*([\d,，]+(?:\.\d+)?)\s*(万元|元)")
_KWD_UP = re.compile("(?:" + _TOTAL_KW + r")[^\n]{0,14}?(?:为|是|[:：])\s*(?:人民币)?\s*([" + _CN_NUM_CHARS
                     + "]{3,}(?:元|圆)(?:整|正)?)")
_FINAL_KW = re.compile(r"最后报价|最终报价|二次报价|成交报价")
_NOT_TOTAL_KW = re.compile(r"保证金|保函|押金|罚款|罚金|违约金|赔偿|预算|控制价|最高限价")
_QUOTE_CTX = re.compile(r"报价|开标|投标总价|响应报价|一览表")
_INTERNAL_ID_RE = re.compile(r"(?<![A-Za-z0-9])(?:MAND|TPL|REQ|SCORE)-\d{1,3}(?![A-Za-z0-9])", re.I)
PLACEHOLDER_CORE = re.compile(r"【\s*(待补|待澄清|待确认|待补充|待定|待核验)[^】]*】|TODO|FIXME")
_PASTE_RE = re.compile(r"[(（]\s*此处贴入")
_BLANK_AMOUNT_RE = re.compile(r"[_＿\s×Xxｘ/／—\-–]*")   # 「小写：_______」：该填总价的地方还是空白
_PLACEHOLDER_RE = re.compile(r"×{2,}|[xXｘ]{3,}|此处填写|按实际填写|请填写|待填写|待补充|待后补|填写此处")
_FIELD_ERR_RE = re.compile(r"未找到引用源|未定义书签|Reference\s+source\s+not\s+found|Bookmark\s+not\s+defined",
                           re.I)
_TOP_TOTAL_RE = re.compile(r"^\s*(合计|总计|总额|总报价|投标总价|报价总额|累计|总金额|响应报价|开标报价|项目总价)")
_SUB_TOTAL_RE = re.compile(r"^\s*(小计)")
_NOTE_ROW_RE = re.compile(r"^\s*(?:注|说明|备注)\s*[：:]")
_TOOL_TRACE_RE = re.compile(r"python-docx|libreoffice|openai|anthropic|claude|chatgpt|generated\s+by"
                            r"|kingsoft|kso|icv", re.I)


def _sha256(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def write_fingerprint(path, by: str = "") -> None:
    """引擎写完一份整本后记下指纹（sha256）。下次检查时与现文件比：不一致 = 整本在引擎
    之外被改过（用户在 Word 里手工调了版式——之江实验室项目手工修好的表，从工作区重新
    生成会冲回）。记不上也不许影响主流程。"""
    try:
        Path(str(path) + FP_SUFFIX).write_text(
            json.dumps({"sha256": _sha256(path), "at": time.strftime("%Y-%m-%d %H:%M"), "by": by},
                       ensure_ascii=False), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def _save_doc(doc, path, by: str = "") -> None:
    """引擎的统一落盘口：落盘前清零元数据（保存即无痕，写整本与写节都要走这里）；
    写的是整本（整本-*.docx）就顺手刷新指纹。"""
    _scrub_metadata(doc)
    doc.save(str(path))
    if Path(path).name.startswith("整本-"):
        write_fingerprint(path, by)


def _book_edited_note(book) -> str:
    """整本在上次生成之后被手工改过 → 返回一句提醒；没有指纹记录（旧版生成的）无从判断，返回空串。"""
    try:
        rec = json.loads(Path(str(book) + FP_SUFFIX).read_text(encoding="utf-8"))
        want = rec.get("sha256") if isinstance(rec, dict) else None
    except Exception:  # noqa: BLE001
        return ""
    if not want or want == _sha256(book):
        return ""
    return (f"现有整本 {Path(book).name} 在上次生成（{rec.get('at', '时间不详')}）之后被手工改过"
            f"——本次生成会覆盖这些改动；要保留，先问清用户改了什么、回到节文件里改")


def _cn_amount(raw: str):
    """中文大写金额 → Decimal（人民币叁佰柒拾玖万伍仟元整 → 3795000）；读不出返回 None。
    认「拾万」省「壹」、「零」占位、角分、整／正，也认照小写习惯写的一二三六十六万。"""
    s = re.sub(r"[(（][^()（）]*[)）]", "", raw or "")
    s = re.sub(r"人民币|RMB|[，,、\s]", "", s, flags=re.I)
    m = re.match("[" + _CN_NUM_CHARS + "角分元圆整正]+", s)
    if not m:
        return None
    run = re.sub(r"[整正]+$", "", m.group(0))
    frac = Decimal(0)
    for unit, denom in (("角", 10), ("分", 100)):
        jm = re.search("([" + _CN_DIG + "])" + unit, run)
        if jm:
            frac += Decimal(_CN_DIGIT[jm.group(1)]) / denom
    ip = re.split(r"[元圆角分]", run)[0]
    total = section = number = 0
    seen = False
    for ch in ip:
        if ch in _CN_DIGIT:
            number, seen = _CN_DIGIT[ch], True
        elif ch in "拾十":
            if not seen:
                number = 1
            section += number * 10
            number, seen = 0, False
        elif ch in "佰百":
            section += number * 100
            number, seen = 0, False
        elif ch in "仟千":
            section += number * 1000
            number, seen = 0, False
        elif ch in "萬万":
            section += number
            total += section * 10 ** 4 if section else 10 ** 4
            section = number = 0
            seen = False
        elif ch in "億亿":
            section += number
            total += section * 10 ** 8 if section else 10 ** 8
            section = number = 0
            seen = False
        else:
            return None
    if not ip and not frac:
        return None
    return Decimal(total + section + number) + frac


def _num(raw: str, qty: bool = False, price: bool = False):
    """格内文本 → Decimal。认全角数字、千分位、¥ 与「人民币」前缀、「万/千/百/亿」量词、「元」
    后缀与括注（含税）；qty 再认尾部单位（30万条、1 批），price 再认「元/条」类计价单位。
    读不出返回 None——调用方要把 None 当一条发现，不许跳过。"""
    s = unicodedata.normalize("NFKC", raw or "")
    s = re.sub(r"¥|\$|人民币|RMB", "", s, flags=re.I)
    s = re.sub(r"[(][^()]*[)]", "", s)  # NFKC 后全角括号已是半角
    s = re.sub(r"(?<=\d),(?=\d)", "", s)
    s = re.sub(r"\s+", "", s).rstrip("；;，,。")
    core = r"(-?\d+(?:\.\d+)?)"
    if price:
        pat = core + r"([万千百亿])?元?(?:[/每][^\d\-/]{1,8})?"
    elif qty:
        pat = core + r"([万千百亿])?[^\d.\-]{0,4}"
    else:
        pat = core + r"([万千百亿])?元?"
    m = re.fullmatch(pat, s)
    if not m:
        return None
    try:
        return Decimal(m.group(1)) * _AMT_MULT.get(m.group(2), 1)
    except Exception:  # noqa: BLE001
        return None


def _rate(raw: str):
    """税率 → 小数（13%、13％、0.13 都认）；读不出返回 None。"""
    s = unicodedata.normalize("NFKC", raw or "").strip()
    if s.endswith("%"):
        v = _num(s[:-1])
        return v / 100 if v is not None else None
    v = _num(s)
    if v is None:
        return None
    return v if v < 1 else v / 100


def _is_na(text: str) -> bool:
    t = (text or "").strip()
    return (not t) or re.fullmatch(r"[—\-–_/／\\]+", t) is not None or t.upper() in ("N/A", "NA", "无", "不适用")


def _places(src: str) -> int:
    m = re.search(r"\d+\.(\d+)", src or "")
    return min(4, len(m.group(1))) if m else 0


def _fm(v, places=None) -> str:
    """数值的人话写法：千分位；places=None 时整数不带小数、小数按原位数（最多 4 位）。"""
    if v is None:
        return "—"
    if places is None:
        if v == v.to_integral_value():
            return f"{v:,.0f}"
        return f"{v.quantize(Decimal('0.0001')).normalize():,f}"
    return f"{v.quantize(Decimal(1).scaleb(-places)):,.{places}f}"


def _rows_grid(tbl_el) -> list[list[dict]]:
    """表 → 逐行格清单 [{g 起始网格列, span, tc, text, cont 竖向续行}]（按真实格，合并格不重复）。"""
    out = []
    for tr in tbl_el.findall(qn("w:tr")):
        cells, g = [], 0
        for tc in tr.findall(qn("w:tc")):
            pr = tc.find(qn("w:tcPr"))
            span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
            vm = pr.find(qn("w:vMerge")) if pr is not None else None
            span = int(span_el.get(qn("w:val"))) if (span_el is not None and span_el.get(qn("w:val"))) else 1
            cells.append({"g": g, "span": max(1, span), "tc": tc,
                          "text": "\n".join(para_text_full(p) for p in tc.findall(qn("w:p"))).strip(),
                          "cont": vm is not None and vm.get(qn("w:val")) is None})
            g += max(1, span)
        out.append(cells)
    return out


def _cell_at(row, g):
    return next((c for c in row if c["g"] <= g < c["g"] + c["span"]), None)


def _grid_widths(tbl_el) -> list[int]:
    grid = tbl_el.find(qn("w:tblGrid"))
    out = []
    for col in (grid.findall(qn("w:gridCol")) if grid is not None else []):
        try:
            out.append(int(col.get(qn("w:w")) or 0))
        except (TypeError, ValueError):
            out.append(0)
    return out


def _hdr_index(rows) -> int:
    """表头行号：首行不足两个非空格（整行合并的标题行）时表头在第二行。"""
    return 1 if len(rows) > 1 and len({c["tc"] for c in rows[0] if c["text"]}) < 2 else 0


def _hdr_cols(hdr) -> dict:
    """表头 → 各类列的网格列号清单。税率／税额／价税合计先认，剩下的才是普通金额列。"""
    out = {"qty": [], "price": [], "amt": [], "rate": [], "tax": [], "gross": []}
    for c in hdr:
        h = re.sub(r"\s+", "", c["text"])
        if not h or c["cont"]:
            continue
        if "数量" in h:
            out["qty"].append(c["g"])
        if "单价" in h:
            out["price"].append(c["g"])
        if "税率" in h:
            out["rate"].append(c["g"])
        elif "税额" in h:
            out["tax"].append(c["g"])
        elif re.search(r"价税合计|含税合计|含税金额", h):
            out["gross"].append(c["g"])
        elif re.search(r"合价|小计|金额|总价|合计", h):
            out["amt"].append(c["g"])
    return out


def _text_em(text: str) -> float:
    """文字宽度估算（em）：汉字与全角算 1，其余算 0.55。"""
    return sum(1.0 if ord(ch) > 0x2E7F else 0.55 for ch in text)


def _style_chain(p_el, doc: Document) -> list:
    """段落的样式链（自身样式 → 基样式 → …），用来取继承来的行距与字号。"""
    from docx.text.paragraph import Paragraph

    chain = []
    try:
        st = Paragraph(p_el, doc).style
    except Exception:  # noqa: BLE001
        return chain
    while st is not None and len(chain) < 8 and all(st.style_id != s.style_id for s in chain):
        chain.append(st)
        st = st.base_style
    return chain


def _doc_default_pt(doc: Document):
    try:
        sz = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")) \
            .find(qn("w:rPr")).find(qn("w:sz"))
        return int(sz.get(qn("w:val"))) / 2 if sz is not None and sz.get(qn("w:val")) else None
    except Exception:  # noqa: BLE001
        return None


def _run_size_pt(r_el, chain, default):
    rpr = r_el.find(qn("w:rPr"))
    sz = rpr.find(qn("w:sz")) if rpr is not None else None
    if sz is not None and sz.get(qn("w:val")):
        return int(sz.get(qn("w:val"))) / 2
    for st in chain:
        rpr = st.element.find(qn("w:rPr"))
        sz = rpr.find(qn("w:sz")) if rpr is not None else None
        if sz is not None and sz.get(qn("w:val")):
            return int(sz.get(qn("w:val"))) / 2
    return default


def _para_max_pt(p_el, doc: Document) -> float:
    """段内最大字号（磅）：直写优先，其次样式链，再次文档默认。"""
    chain = _style_chain(p_el, doc)
    default = _doc_default_pt(doc) or 10.5
    best = 0.0
    for r in p_el.iter(qn("w:r")):
        if not "".join(t.text or "" for t in r.iter(qn("w:t"))).strip():
            continue
        best = max(best, _run_size_pt(r, chain, default))
    return best or default


def _exact_line_pt(p_el, doc: Document):
    """段落生效的「固定值」行距（磅）；没设固定行距返回 None。直写优先，其次样式链。"""
    holders = [p_el.find(qn("w:pPr"))] + [st.element.find(qn("w:pPr")) for st in _style_chain(p_el, doc)]
    for ppr in holders:
        sp = ppr.find(qn("w:spacing")) if ppr is not None else None
        if sp is not None and sp.get(qn("w:line")):
            return int(sp.get(qn("w:line"))) / 20 if sp.get(qn("w:lineRule")) == "exact" else None
    return None


def _sect_width_emu(sectpr) -> int:
    if sectpr is not None:
        sz, mar = sectpr.find(qn("w:pgSz")), sectpr.find(qn("w:pgMar"))
        try:
            w, left, right = int(sz.get(qn("w:w"))), int(mar.get(qn("w:left"))), int(mar.get(qn("w:right")))
            gutter = int(mar.get(qn("w:gutter")) or 0)
            if w > 2000:
                return (w - left - right - gutter) * 635
        except Exception:  # noqa: BLE001
            pass
    return 8505 * 635  # 15cm 兜底（与 _content_width_twips 同口径）


def _block_widths_emu(doc: Document, blocks: list) -> dict:
    """正文顶层块 → 所在节的可用宽度（EMU，页宽减左右页边距与装订线）。节的分界是段落里的
    分节符（那一节到该段为止）与文末的分节属性，所以从后往前扫。键是块的 id——blocks
    必须由调用方拿住不放，lxml 的代理对象被回收后 id 会对不上。"""
    body = doc.element.body
    out: dict = {}
    cur = body.find(qn("w:sectPr"))
    for child in reversed(blocks):
        out[id(child)] = _sect_width_emu(cur)
        if child.tag == qn("w:p"):
            ppr = child.find(qn("w:pPr"))
            sp = ppr.find(qn("w:sectPr")) if ppr is not None else None
            if sp is not None:
                cur = sp
    return out


def _new_finding(kind, level, file, where, what, basis, anchor="", order=0) -> dict:
    return {"kind": kind, "level": level, "file": file, "where": where, "what": what,
            "basis": basis, "anchor": anchor, "order": order}


def inspect_docx(path, final=False, core_placeholders=True, props=True, label=None, totals=None,
                 scope=None) -> list[dict]:
    """单份文档的只读检查，返回发现清单（见本段开头说明）。

    final=True 按递交前的稿子查（未处理的批注算问题）；core_placeholders=False 时不重复查
    【待补】类占位（流水线里 check_writing 已当红线拦）。totals 收集各处写的总价供跨文件
    对账；scope 记已检查的范围。每类检查各自兜异常——检查自身出错只产出一条「没能执行」
    的疑似发现，不许拖垮调用方（与格式件保真核对同一条纪律）。props=False 不看文档属性。"""
    path = Path(path)
    label = path.stem if label is None else label
    totals = [] if totals is None else totals
    scope = {} if scope is None else scope
    out: list[dict] = []

    def add(kind, level, where, what, basis, anchor=""):
        out.append(_new_finding(kind, level, label, where, what, basis, anchor, len(out)))

    def guard(name, fn):
        try:
            fn()
        except Exception as exc:  # noqa: BLE001
            add("check_error", "疑似", "整份文件", f"「{name}」这一项检查没能执行",
                f"{type(exc).__name__}: {exc}——这一类请人工核对")

    try:
        doc = Document(str(path))
    except Exception as exc:  # noqa: BLE001
        add("open_error", "确定", "整份文件", "文件打不开", f"{type(exc).__name__}: {exc}")
        return out
    body = doc.element.body
    blocks = list(body.iterchildren())
    widths = _block_widths_emu(doc, blocks)
    try:
        tables = list(doc.tables)
    except Exception:  # noqa: BLE001
        tables = []
    grid = [_rows_grid(t._tbl) for t in tables]
    cell_pos = {id(c["tc"]): (ti, ri, c["g"])
                for ti, rows in enumerate(grid, 1) for ri, row in enumerate(rows, 1) for c in row}
    para_no = {id(el): n for n, el in enumerate((b for b in blocks if b.tag == qn("w:p")), 1)}

    def hdr_name(ti, g):
        rows = grid[ti - 1]
        cell = _cell_at(rows[_hdr_index(rows)], g) if rows else None
        name = re.sub(r"\s+", "", (cell["text"] if cell else "").split("\n")[0])[:10]
        return name or f"第 {g + 1} 列"

    def cell_where(ti, ri, g):
        row = grid[ti - 1][ri - 1] if ri - 1 < len(grid[ti - 1]) else []
        first = next((re.sub(r"\s+", "", c["text"])[:12] for c in row if c["text"]), "")
        return f"第 {ti} 张表第 {ri} 行" + (f"（{first}）" if first else "") + f"「{hdr_name(ti, g)}」那格"

    def para_where(p_el):
        txt = re.sub(r"\s+", " ", para_text_full(p_el)).strip()
        head = txt[:16] + ("…" if len(txt) > 16 else "")
        tc = next(p_el.iterancestors(qn("w:tc")), None)
        if tc is not None:
            pos = cell_pos.get(id(tc))
            return (cell_where(*pos), f"T{pos[0]}R{pos[1]}C{pos[2] + 1}") if pos else (f"表格里「{head}」", "")
        n = para_no.get(id(p_el))
        return (f"正文「{head}」那段", f"P{n}") if n else (f"「{head}」那段", "")

    # ---- 文字类：内部对账编号、占位写法、域报错、大小写金额、正文里的总价
    def text_unit(text, where, anchor, short):
        ids = sorted({m.group(0).upper() for m in _INTERNAL_ID_RE.finditer(text)})
        if ids:
            add("internal_id", "确定", where, "正文里出现内部对账编号",
                "、".join(f"「{i}」" for i in ids) + "是写作时用的对账编号，评标人看不懂——删掉，或换成招标文件里真实的章节号",
                anchor)
        m = PLACEHOLDER_CORE.search(text) if core_placeholders else None
        if m:
            add("placeholder_core", "确定", where, "正文里有占位文字", f"「{m.group(0)}」会随交付稿出去——改成正式内容，或用 Word 批注记待办", anchor)
        m = _PASTE_RE.search(text)
        if m:
            add("placeholder_paste", "确定", where, "贴入位还在", f"「{text[:30]}」——附件没贴，或贴了之后这句占位没删", anchor)
        m = _PLACEHOLDER_RE.search(text)
        if m:
            add("placeholder_other", "确定", where, "正文里有没替换的占位写法", f"「{m.group(0)}」是模板里留的占位，要换成实际内容", anchor)
        m = _FIELD_ERR_RE.search(text)
        if m:
            add("field_error", "确定", where, "正文里有域报错文字", f"「{m.group(0)}」——引用的对象不存在或没更新，删掉重做这个引用", anchor)
        ctx = label + " " + text
        if _NOT_TOTAL_KW.search(text[:24]):
            cat = "other"  # 保证金、预算价不是投标总价，不参与总价对账
        else:
            cat = "final" if _FINAL_KW.search(ctx) else "first"
        lo, up = _LOWER_RE.search(text), _UPPER_RE.search(text)
        lo_txt = lo.group(1).strip() if lo else ""
        if lo and lo_txt and _BLANK_AMOUNT_RE.fullmatch(lo_txt):
            add("amount_blank", "确定", where, "总价还没填",
                f"「{text[:24]}」——投标总价要填在这里", anchor)
            lo, lo_txt = None, ""
        up_txt = up.group(1).strip() if up else ""
        if not re.search(r"[\d¥￥]", lo_txt):
            lo, lo_txt = None, ""
        if len(re.findall(f"[{_CN_NUM_CHARS}]", up_txt)) < 2:
            up, up_txt = None, ""
        lv = uv = None
        if lo:
            lv = _num(lo_txt)
            if lv is None:
                add("amount_unread", "确定", where, "小写金额读不出来", f"「{lo_txt[:24]}」不是能核对的数字，请人工核对", anchor)
        if up:
            uv = _cn_amount(up_txt)
            if uv is None:
                add("amount_unread", "确定", where, "大写金额读不出来", f"「{up_txt[:24]}」按大写金额的写法读不出来，请人工核对", anchor)
        if lv is not None and uv is not None and abs(lv - uv) > Decimal("0.004"):
            add("amount_case", "确定", where, "大写与小写不一致", f"大写是 {_fm(uv)}，小写是 {_fm(lv, _places(lo_txt))}", anchor)
        if lv is not None or uv is not None:
            totals.append({"cat": cat, "value": lv if lv is not None else uv, "where": short, "file": label})
            return
        for pat, conv in ((_KWD_NUM, lambda m: _num(m.group(1) + m.group(2))),
                          (_KWD_UP, lambda m: _cn_amount(m.group(1))),
                          (_UP_THEN_NUM, lambda m: _num(m.group(2) + (m.group(3) or "") + "元"))):
            for m in pat.finditer(text):
                v = conv(m)
                if v is not None:
                    totals.append({"cat": cat, "value": v, "where": short, "file": label})

    def text_pass():
        ti = 0
        for child in blocks:
            if child.tag == qn("w:p"):
                txt = para_text_full(child).strip()
                if txt:
                    where, anchor = para_where(child)
                    text_unit(txt, where, anchor, f"正文「{txt[:10]}」")
            elif child.tag == qn("w:tbl"):
                ti += 1
                for ri, row in enumerate(grid[ti - 1], 1):
                    for c in row:
                        if c["cont"] or not c["text"]:
                            continue
                        text_unit(c["text"], cell_where(ti, ri, c["g"]), f"T{ti}R{ri}C{c['g'] + 1}",
                                  f"第 {ti} 张表第 {ri} 行")

    def line_pass():
        for p in body.iter(qn("w:p")):
            txt = para_text_full(p).strip()
            if not txt:
                continue  # 空段的定行高是刻意的间隔（册尾空段就压成 2 磅），不是压字
            line = _exact_line_pt(p, doc)
            size = _para_max_pt(p, doc)
            if line is None or line <= 0:
                continue
            if line < size:
                level = "确定"
            elif line < size * 1.1:
                level = "疑似"
            else:
                continue
            where, anchor = para_where(p)
            add("line_clip", level, where, "行距比字小，文字会被压掉", f"行距固定 {line:g} 磅，字有 {size:g} 磅", anchor)

    # ---- 金额验算：行（单价×数量）、列（合计=明细之和，多级小计逐级）、0 与负数、税额
    AMT_NAMES = {"qty": "数量", "price": "单价", "amt": "金额"}

    def amount_table(tbl_el, ti, rows, hi, cols, ctx):
        if not (cols["amt"] or cols["price"]):
            return  # 没有金额列也没有单价列（人员、设备台账里的「数量」）——不是报价表，不在这里验算
        scope["amount_tables"] = scope.get("amount_tables", 0) + 1
        top_last: dict = {}
        detail_sum: dict = {}
        sub_sum: dict = {}
        sub_totals: dict = {}
        n_detail = n_verified = n_numeric = n_named_blank = n_blank_rows = 0
        names_blank: list[str] = []
        for ri, row in enumerate(rows[hi + 1:], hi + 2):
            first = next((c["text"] for c in row if c["text"]), "")
            mt = _TOP_TOTAL_RE.match(first) or _SUB_TOTAL_RE.match(first)
            if mt:
                top = mt.re is _TOP_TOTAL_RE
                for g in cols["amt"]:
                    cell = _cell_at(row, g)
                    src = cell["text"] if cell is not None and not cell["cont"] else ""
                    val = _num(src)
                    if val is None:  # 合计行的数不在金额列（整行合并、只写一个数）
                        nums = [v for v in (_num(c["text"]) for c in row[1:]) if v is not None]
                        val = nums[0] if len(nums) == 1 else None
                    if val is None:
                        continue
                    refs = []
                    if top:
                        refs += [x for x in (detail_sum.get(g), sub_totals.get(g)) if x is not None]
                    elif g in sub_sum:
                        refs.append(sub_sum[g])
                    if refs and all(abs(val - r) > _AMT_TOL for r in refs):
                        ref = refs[0]
                        pl = _places(src) or 2
                        add("amount_col", "确定", cell_where(ti, ri, g), f"{mt.group(1)}行 ≠ 明细之和",
                            f"{mt.group(1)}行写的是 {_fm(val, pl)}，上方明细相加是 {_fm(ref, pl)}，差 {_fm(abs(val - ref), pl)}",
                            f"T{ti}R{ri}C{g + 1}")
                        scope["amount_checks"] = scope.get("amount_checks", 0) + 1
                    if top:
                        top_last[g] = (val, first)
                    else:
                        sub_totals[g] = sub_totals.get(g, Decimal(0)) + val
                if top:
                    detail_sum, sub_totals = {}, {}
                sub_sum = {}
                continue
            if _NOTE_ROW_RE.match(first):
                continue
            got = {}
            for kind in ("qty", "price", "amt", "rate", "tax", "gross"):
                for g in cols[kind]:
                    cell = _cell_at(row, g)
                    txt = cell["text"] if cell is not None and not cell["cont"] else ""
                    # 「小写：¥ … 大写：人民币…」写在同一格是常见形态，归文字类的大小写核对处理；
                    # 这里按空格看待，不然会误报成「金额读不出来」（真实报价表回放时踩到）
                    blank = _is_na(txt) or bool(_LOWER_RE.search(txt) or _UPPER_RE.search(txt))
                    if blank:
                        val = None
                    elif kind == "rate":
                        val = _rate(txt)
                    else:
                        val = _num(txt, qty=(kind == "qty"), price=(kind == "price"))
                    got[(kind, g)] = (txt, val, blank, cell["span"] if cell is not None else 1)
            if any(got[(k, g_)][1] is not None for k in ("qty", "price", "amt") for g_ in cols[k]):
                n_numeric += 1
            elif first and not _NOTE_ROW_RE.match(first) and all(
                    got[(k, g_)][2] for k in ("qty", "price", "amt") for g_ in cols[k]):
                # 有名称、金额格全空（真空白，不含「填了但读不出」）的行：记下来——
                # 整张表一个数都没有时点名「金额格都空着」
                n_named_blank += 1
                names_blank.append(next((c["text"][:10] for c in row[:3]
                                          if c["text"] and not re.fullmatch(r"[\d.、．()（）]+", c["text"])), first[:10]))
            if all(v[2] for v in got.values()):
                continue  # 只有名称的分组行、整行空格
            n_detail += 1

            def pick(kind):
                for g in cols[kind]:
                    txt, val, _b, _s = got[(kind, g)]
                    if val is not None:
                        return g, txt, val
                return None

            def all_blank(kind):
                return all(got[(kind, g)][2] for g in cols[kind])

            a, p, q = pick("amt"), pick("price"), pick("qty")
            bad = [(kind, g, txt, val) for (kind, g), (txt, val, _b, _s) in got.items()
                   if kind in AMT_NAMES and val is not None and val <= 0]
            if bad:
                neg = any(v < 0 for _k, _g, _t, v in bad)
                k0, g0 = bad[0][0], bad[0][1]
                add("amount_sign", "确定" if neg else "疑似", cell_where(ti, ri, g0), f"{AMT_NAMES[k0]}为 0 或负数",
                    "、".join(f"{AMT_NAMES[k]} {t}" for k, _g, t, _v in bad)
                    + ("——负数要改" if neg else "——0 元若是免费项可忽略"), f"T{ti}R{ri}C{g0 + 1}")
            for (kind, g), (txt, val, blank, span) in got.items():
                if kind in AMT_NAMES and not blank and val is None and span == 1:
                    # 跨列的格是整行合并的标签或一句话，不当成数字格去读（「项目总价」占四列那种）
                    add("amount_unread", "确定", cell_where(ti, ri, g), f"{AMT_NAMES[kind]}读不出来",
                        f"「{txt[:24]}」不是能核对的数字，这一格请人工核对", f"T{ti}R{ri}C{g + 1}")

            def blank_row(kind, what, basis, level="确定"):
                """该填的格空着：点在那一列的第一格上。n_blank_rows 用来避免表级再重复点一次。"""
                nonlocal n_blank_rows
                n_blank_rows += 1
                g = cols[kind][0]
                add("amount_blank", level, cell_where(ti, ri, g), what, basis, f"T{ti}R{ri}C{g + 1}")

            if a and q and cols["price"] and all_blank("price"):
                blank_row("price", "单价空着", f"数量 {q[1]}、金额 {a[1]} 都填了，单价格是空的")
            if a and p and cols["qty"] and all_blank("qty"):
                blank_row("qty", "数量空着", f"单价 {p[1]}、金额 {a[1]} 都填了，数量格是空的")
            if p and q and cols["amt"] and all_blank("amt"):
                blank_row("amt", "合价空着", f"单价 {p[1]}、数量 {q[1]} 都填了，合价格是空的")
            # 真实工作区回放：一览表只填了数量、单价与合计空着，小写处还是下划线——报价根本没填
            if q and not p and not a and cols["price"] and cols["amt"] and all_blank("price") and all_blank("amt"):
                blank_row("price", "单价和合价都空着", f"数量 {q[1]} 填了，单价、合价格是空的——填上才能验算")
            if p and not q and not a and cols["qty"] and cols["amt"] and all_blank("qty") and all_blank("amt"):
                blank_row("qty", "数量和合价都空着", f"单价 {p[1]} 填了，数量、合价格是空的——填上才能验算")
            if a and cols["price"] and cols["qty"] and all_blank("price") and all_blank("qty"):
                blank_row("price", "数量和单价都空着", f"这一行只有金额 {a[1]}，没法验算——补上数量和单价，或人工核对",
                          "疑似")
            if p and q and a:
                n_verified += 1
                scope["amount_checks"] = scope.get("amount_checks", 0) + 1
                diff = p[2] * q[2] - a[2]
                if abs(diff) > _AMT_TOL:
                    pl = _places(a[1]) or max(_places(p[1]), _places(q[1]))
                    add("amount_row", "确定", cell_where(ti, ri, a[0]), "单价 × 数量 ≠ 合价",
                        f"{_fm(p[2], _places(p[1]))} × {_fm(q[2], _places(q[1]))} = {_fm(p[2] * q[2], pl)}"
                        f"，表里写的是 {_fm(a[2], _places(a[1]))}，差 {_fm(abs(diff), pl)}",
                        f"T{ti}R{ri}C{a[0] + 1}")
            for g in cols["amt"]:
                v = got[("amt", g)][1]
                if v is not None:
                    detail_sum[g] = detail_sum.get(g, Decimal(0)) + v
                    sub_sum[g] = sub_sum.get(g, Decimal(0)) + v
            r_, t_, g_ = pick("rate"), pick("tax"), pick("gross")
            if a and r_ and t_:
                exp = (a[2] * r_[2]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if abs(t_[2] - exp) > Decimal("0.02"):
                    add("amount_tax", "确定", cell_where(ti, ri, t_[0]), "税额与金额 × 税率对不上",
                        f"{_fm(a[2], _places(a[1]))} × {r_[1]} = {_fm(exp, 2)}，表里税额写的是 {_fm(t_[2], _places(t_[1]))}",
                        f"T{ti}R{ri}C{t_[0] + 1}")
            if a and t_ and g_ and abs(g_[2] - (a[2] + t_[2])) > _AMT_TOL:
                pl = _places(g_[1]) or 2
                add("amount_gross", "确定", cell_where(ti, ri, g_[0]), "价税合计 ≠ 金额 + 税额",
                    f"{_fm(a[2], _places(a[1]))} + {_fm(t_[2], _places(t_[1]))} = {_fm(a[2] + t_[2], pl)}"
                    f"，表里写的是 {_fm(g_[2], pl)}", f"T{ti}R{ri}C{g_[0] + 1}")
        if cols["price"] and cols["amt"] and n_named_blank and not n_numeric:
            # 真实工作区回放：有项目的分项报价明细表数量、单价、合计金额整列是空的——
            # 「没有一行能验算」说不到点上，这里直接点名一个数都没填
            add("amount_blank", "确定", f"第 {ti} 张表", "报价表的金额格都空着",
                f"「{'」「'.join(names_blank[:3])}」等 {n_named_blank} 行的数量、单价、金额一个数都没填")
        elif (cols["qty"] and cols["price"] and cols["amt"] and n_detail and not n_verified
              and not n_blank_rows):
            scope.setdefault("amount_unverified", 0)
            scope["amount_unverified"] += 1
            add("amount_none", "确定", f"第 {ti} 张表", "认出是报价表，但没有一行能验算",
                f"表头有「{hdr_name(ti, cols['qty'][0])}」「{hdr_name(ti, cols['price'][0])}」这些列，"
                f"可 {n_detail} 行明细的数字读不出来或没填全——逐格人工核对")
        g = cols["amt"][0] if cols["amt"] else None
        if g is not None and g in top_last and _QUOTE_CTX.search(ctx + " " + top_last[g][1]):
            cat = "final" if _FINAL_KW.search(ctx) else "first"
            totals.append({"cat": cat, "value": top_last[g][0], "where": f"第 {ti} 张表{top_last[g][1][:4]}行", "file": label})

    # ---- 表格版式：超宽、列太窄、固定行高裁字
    def geometry(tbl_el, ti, rows, usable, hi):
        widths_tw = _grid_widths(tbl_el)
        if not widths_tw:
            return
        pr = tbl_el.find(qn("w:tblPr"))
        layout = pr.find(qn("w:tblLayout")) if pr is not None else None
        tblw = pr.find(qn("w:tblW")) if pr is not None else None
        ind = pr.find(qn("w:tblInd")) if pr is not None else None
        fixed = layout is not None and layout.get(qn("w:type")) == "fixed"
        mar = {"left": 108, "right": 108}
        cm = pr.find(qn("w:tblCellMar")) if pr is not None else None
        if cm is not None:
            for side in mar:
                e = cm.find(qn(f"w:{side}"))
                if e is not None and e.get(qn("w:w")) and e.get(qn("w:type"), "dxa") == "dxa":
                    mar[side] = int(e.get(qn("w:w")))
        total_tw = sum(widths_tw)
        if ind is not None and ind.get(qn("w:w")) and ind.get(qn("w:type"), "dxa") == "dxa":
            total_tw += int(ind.get(qn("w:w")))
        # 容差 0.5 厘米：拷入的表常比正文区宽 0.36 厘米左右（落进页边距里，打印无碍），再宽才点名
        if usable and total_tw - usable / 635 > 283:
            dxa = tblw is not None and tblw.get(qn("w:type")) == "dxa"
            add("table_wide", "确定" if (fixed or dxa) else "疑似", f"第 {ti} 张表（{hdr_name(ti, 0)}）", "表格超出可用宽度",
                f"表格宽 {total_tw / 567:.1f} 厘米，页面可用 {usable / 635 / 567:.1f} 厘米，"
                f"多出 {(total_tw - usable / 635) / 567:.1f} 厘米")
        for g, w in enumerate(widths_tw):
            if w <= 0:
                continue
            inner_cm = w / 567 - (mar["left"] + mar["right"]) / 567
            worst = None
            for row in rows[hi:]:
                cell = _cell_at(row, g)
                if cell is None or cell["cont"] or cell["g"] != g or cell["span"] != 1:
                    continue
                txt = re.sub(r"\s+", "", cell["text"])
                if _text_em(txt) < 4:
                    continue  # 只有表头那两个字、或格内是「√」「12」这类短内容：窄是刻意的
                size = max((_para_max_pt(p, doc) for p in cell["tc"].findall(qn("w:p"))), default=10.5)
                em_line = inner_cm / (size * 0.03528) if inner_cm > 0 else 0.0
                if worst is None or em_line < worst[0]:
                    worst = (em_line, txt, size)
            if worst and worst[0] <= 2:
                add("table_narrow", "疑似", f"第 {ti} 张表「{hdr_name(ti, g)}」列", "列太窄，文字会一字一行",
                    f"列宽 {w / 567:.1f} 厘米，按 {worst[2]:g} 磅字每行只放得下约 {worst[0]:.1f} 个字，格内有「{worst[1][:12]}」")
        for ri, tr in enumerate(tbl_el.findall(qn("w:tr")), 1):
            trpr = tr.find(qn("w:trPr"))
            h = trpr.find(qn("w:trHeight")) if trpr is not None else None
            if h is None or h.get(qn("w:hRule")) != "exact" or not h.get(qn("w:val")):
                continue
            hpt = int(h.get(qn("w:val"))) / 20
            need = size_max = 0.0
            g = 0
            for tc in tr.findall(qn("w:tc")):
                pr_ = tc.find(qn("w:tcPr"))
                se = pr_.find(qn("w:gridSpan")) if pr_ is not None else None
                span = int(se.get(qn("w:val"))) if (se is not None and se.get(qn("w:val"))) else 1
                txt = re.sub(r"\s+", "", "".join(t.text or "" for t in tc.iter(qn("w:t"))))
                if txt:
                    size = max((_para_max_pt(p, doc) for p in tc.findall(qn("w:p"))), default=10.5)
                    inner = (sum(widths_tw[g:g + max(1, span)]) - mar["left"] - mar["right"]) / 567
                    em_line = max(inner, 0.01) / (size * 0.03528)
                    need = max(need, math.ceil(_text_em(txt) / em_line) * size * 1.3)
                    size_max = max(size_max, size)
                g += max(1, span)
            if not size_max or hpt >= size_max * 1.2:
                continue
            level = "确定" if hpt < size_max else "疑似"
            basis = f"行高固定 {hpt:g} 磅，格内文字一行就要约 {size_max * 1.3:.0f} 磅"
            if need > size_max * 1.3:
                basis += f"，全部约 {need:.0f} 磅"
            add("row_clip", level, f"第 {ti} 张表第 {ri} 行", "固定行高会裁掉文字", basis)

    def tables_pass():
        head = ""
        ti = 0
        for child in blocks:
            if child.tag == qn("w:p"):
                st = _style_chain(child, doc)
                if st and re.sub(r"\s+", "", st[0].name).lower().startswith(("heading", "标题")):
                    head = para_text_full(child).strip()
            elif child.tag == qn("w:tbl"):
                ti += 1
                rows = grid[ti - 1]
                if not rows:
                    continue
                scope["tables"] = scope.get("tables", 0) + 1
                hi = _hdr_index(rows)
                ctx = f"{label} {head}"
                guard(f"第 {ti} 张表金额验算",
                      lambda t=tables[ti - 1], ti_=ti, r=rows, h=hi, c=ctx: amount_table(t._tbl, ti_, r, h, _hdr_cols(r[h]), c))
                guard(f"第 {ti} 张表版式",
                      lambda t=tables[ti - 1], ti_=ti, r=rows, u=widths.get(id(child)), h=hi:
                      geometry(t._tbl, ti_, r, u, h))

    # ---- 图片：拉伸变形、超宽、打不开、分辨率偏低
    def images_pass():
        try:
            from PIL import Image
        except ImportError:
            Image = None
            scope.setdefault("skipped", []).append(("图片像素核对", "本机缺 Pillow"))
        k = 0
        for blip in body.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            holder = blip
            while holder is not None and holder.tag not in (qn("wp:inline"), qn("wp:anchor")):
                holder = holder.getparent()
            ext = holder.find(qn("wp:extent")) if holder is not None else None
            if not rid or ext is None:
                continue
            k += 1
            scope["images"] = scope.get("images", 0) + 1
            cx, cy = int(ext.get("cx") or 0), int(ext.get("cy") or 0)
            inline = holder.tag == qn("wp:inline")
            p_el = next(blip.iterancestors(qn("w:p")), None)
            where = f"第 {k} 张图"
            if p_el is not None:
                tc = next(p_el.iterancestors(qn("w:tc")), None)
                if tc is not None and id(tc) in cell_pos:
                    where = f"第 {k} 张图（{cell_where(*cell_pos[id(tc)])}）"
                else:
                    ctx = ""
                    for sib in p_el.itersiblings(preceding=True):
                        if sib.tag == qn("w:p"):
                            ctx = para_text_full(sib).strip()
                            if ctx:
                                break
                        elif sib.tag == qn("w:tbl"):
                            ctx = "上一张表"
                            break
                    where = f"第 {k} 张图（「{ctx[:12]}{'…' if len(ctx) > 12 else ''}」之后）" if ctx else where
            try:
                part = doc.part.rels[rid].target_part
            except Exception:  # noqa: BLE001
                part = None
            if part is None:
                add("image_broken", "确定", where, "图片打不开", "图片的引用指向一个不存在的文件")
                continue
            w = h = 0
            if Image is not None and not re.search(r"emf|wmf", part.content_type or "", re.I):
                try:
                    with Image.open(io.BytesIO(part.blob)) as im:
                        w, h = im.size
                        im.load()
                    if w <= 0 or h <= 0:
                        raise ValueError("空图")
                except Exception:  # noqa: BLE001
                    add("image_broken", "确定", where, "图片打不开", "文件里的图片数据不是一张能解码的图——重新插入这张图")
                    continue
            blk = p_el
            while blk is not None and blk.getparent() is not None and blk.getparent() is not body:
                blk = blk.getparent()
            usable = widths.get(id(blk), 0) if blk is not None else 0
            # 容差 0.5 厘米：真实图片普遍比可用宽度多 0.4 厘米左右（落进页边距里，打印无碍），
            # 超过 1 厘米才确定是问题
            if usable and cx - usable > 180000:
                over = (cx - usable) / 360000
                add("image_wide", "确定" if over > 1 else "疑似", where, "图片超出可用宽度",
                    f"图宽 {cx / 360000:.1f} 厘米，页面可用 {usable / 360000:.1f} 厘米，多出 {over:.1f} 厘米")
            if w and h and cx and cy:
                w2, h2 = w, h
                src = blip.getparent().find(qn("a:srcRect")) if blip.getparent() is not None else None
                if src is not None:
                    l_, t_, r_, b_ = (int(src.get(x) or 0) / 100000 for x in ("l", "t", "r", "b"))
                    w2, h2 = w * (1 - l_ - r_), h * (1 - t_ - b_)
                if w2 > 0 and h2 > 0 and abs((cx / cy) / (w2 / h2) - 1) > 0.05:
                    add("image_stretch", "确定", where, "图片被拉伸变形",
                        f"原图 {w}×{h}，页面上按 {cx / 360000:.1f}×{cy / 360000:.1f} 厘米显示")
                elif cx:
                    dpi = w / (cx / 914400)
                    if dpi < 72:
                        add("image_lowres", "疑似", where, "图片分辨率偏低",
                            f"原图 {w}×{h} 像素，按 {cx / 360000:.1f} 厘米宽显示，约 {dpi:.0f} DPI——打印会发虚")

    # ---- 修订痕迹、文档属性、批注
    def revision_pass():
        ins = body.findall(".//" + qn("w:ins"))
        dels = body.findall(".//" + qn("w:del"))
        fmt = body.findall(".//" + qn("w:rPrChange")) + body.findall(".//" + qn("w:pPrChange"))
        if not (ins or dels or fmt):
            return
        first = next((el for el in body.iter(qn("w:p"))
                      if any(el.findall(".//" + t) for t in (qn("w:ins"), qn("w:del")))), None)
        where, anchor = para_where(first) if first is not None else ("正文", "")
        add("tracked_changes", "确定", where, "有没接受的修订",
            f"{len(ins) + len(dels) + len(fmt)} 处（插入 {len(ins)}、删除 {len(dels)}、格式 {len(fmt)}）"
            "——定稿前逐条接受或拒绝", anchor)

    def props_pass():
        cp = doc.core_properties
        filled = [(name, (getattr(cp, attr) or "").strip()) for name, attr in
                  (("作者", "author"), ("最后修改者", "last_modified_by"), ("备注", "comments"),
                   ("主题", "subject"), ("关键词", "keywords"), ("类别", "category"))]
        app_name = ""
        parts = _props_parts(doc)
        app = parts.get("/docProps/app.xml")
        if app is not None:
            try:
                from lxml import etree

                aroot = (app._element if getattr(app, "_element", None) is not None
                         else etree.fromstring(app.blob))
                app_name = (aroot.findtext(_APP_NS + "Application") or "").strip()
                for tag, name in (("Company", "单位"), ("Manager", "管理者")):
                    v = (aroot.findtext(_APP_NS + tag) or "").strip()
                    if v:
                        filled.append((name, v))
            except Exception:  # noqa: BLE001
                pass
        cust = parts.get("/docProps/custom.xml")
        cust_names = []
        if cust is not None:
            try:
                from lxml import etree

                cust_names = [el.get("name") or "" for el in etree.fromstring(cust.blob)]
                cust_names = [n for n in cust_names if n]
            except Exception:  # noqa: BLE001
                pass
        filled = [(n, v) for n, v in filled if v]
        thumb = any(rel.reltype.rsplit("/", 1)[-1] in ("thumbnail", "thumbnail-picture")
                    for rel in doc.part.package.rels.values())
        if filled or cust_names:
            tool = bool(_TOOL_TRACE_RE.search(
                "；".join(v for _n, v in filled) + " " + app_name + " " + " ".join(cust_names)))
            where_bits = [f"{n}={v[:30]}" for n, v in filled]
            if cust_names:
                where_bits.append("自定义属性 " + "、".join(cust_names))
            add("doc_props", "确定" if tool else "疑似", "文档属性（文件 › 信息）",
                "文档属性里留着工具痕迹" if tool else "文档属性里有作者或单位信息",
                "；".join(where_bits)
                + ("——随文件一起交出去，生成文档时要清掉" if tool else "——匿名评审时会暴露身份，递交前确认要不要清"))
        if thumb:
            add("thumbnail", "疑似", "文档属性（首页缩略图）", "文件里带着首页缩略图",
                "缩略图是生成时模板带的预览图，随文件一起交出去——生成文档时要清掉")

    def comments_pass():
        anchored = len(body.findall(".//" + qn("w:commentReference")))
        n_bodies = 0
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/comments"):
                n_bodies += len(_comments_root(rel.target_part).findall(qn("w:comment")))
        if n_bodies > anchored:
            add("comments_orphan", "确定", "批注", "有看不见的批注残留",
                f"{n_bodies - anchored} 条批注的正文锚点已经没了——Word 里看不到，却留在文件里，交付前清掉")
        if final and anchored:
            add("comments_open", "确定", "批注", f"还有 {anchored} 条批注没处理", "递交前逐条处理并删除——批注会随 Word 文件一起交出去")
        scope["comments"] = scope.get("comments", 0) + anchored + n_bodies

    scope["files"] = scope.get("files", 0) + 1
    guard("文字（编号、占位、大小写金额）", text_pass)
    guard("行距", line_pass)
    guard("表格", tables_pass)
    guard("图片", images_pass)
    guard("修订痕迹", revision_pass)
    if props:
        guard("文档属性", props_pass)
    guard("批注", comments_pass)
    return out


def _cross_totals(totals, multi=False) -> list[dict]:
    """各处写的总价互相对账：同一报价口径下出现两个以上的数就点名。初次报价与泛指的
    「投标总价」放一起比；最后报价（下浮之后）另算，不与初次比；保证金、预算价不参与。"""
    out = []
    for cat, name in (("first", "投标总价"), ("final", "最后报价")):
        groups: dict = {}
        for t in totals:
            if t["cat"] == cat:
                groups.setdefault(t["value"].normalize(), []).append(t)
        if len(groups) < 2:
            continue
        items = []
        for v in sorted(groups):
            where = "、".join(dict.fromkeys((f"{t['file']} " if multi else "") + t["where"] for t in groups[v]))
            items.append(f"{_fm(v, 2)}（{where[:60]}）")
        out.append(_new_finding("amount_cross", "确定", "", name, "各处写的总价不一致", "；".join(items), "", -1))
    return out


def _sort_findings(findings: list[dict]) -> list[dict]:
    """确定的排前面（跨文件总价最先），其余按文件与出现先后。"""
    findings.sort(key=lambda f: (f["level"] != "确定", f["order"] > 0, f["order"] if f["order"] < 0 else 0,
                                 f["file"], abs(f["order"])))
    return findings


def inspect_files(paths, final=False, core_placeholders=True, props=True, scope=None) -> list[dict]:
    """多份文档的只读检查：各文件逐份查，再把各处写的总价放在一起对账。"""
    scope = {} if scope is None else scope
    findings, totals = [], []
    multi = len(paths) > 1
    for p in paths:
        findings += inspect_docx(p, final=final, core_placeholders=core_placeholders, props=props,
                                 totals=totals, scope=scope)
    scope["totals"] = len(totals)
    findings += _cross_totals(totals, multi)
    if multi:
        for f in findings:
            if f["kind"] != "amount_cross":
                f["where"] = f"{f['file']} › {f['where']}"
    return _sort_findings(findings)


def _book_recon(book: Path, secs: list) -> list[dict]:
    """整本与节文件的三项对账：节文件比整本新、整本被手工改过、批注数对不上。"""
    out, name = [], book.stem
    try:
        btime = book.stat().st_mtime
        newer = [p for p in secs if p.stat().st_mtime > btime + 2]
        if newer:
            shown = "、".join(p.stem for p in newer[:3]) + ("…" if len(newer) > 3 else "")
            out.append(_new_finding("book_stale", "确定", name, "整份文件", "整本比节文件旧",
                                    f"「{shown}」在整本生成之后改过——整本没有重新生成，现在交付会漏掉这些改动", "", -4))
        note = _book_edited_note(book)
        if note:
            out.append(_new_finding("book_edited", "确定", name, "整份文件", "整本被手工改过",
                                    note.split("——", 1)[1] if "——" in note else note, "", -3))
        n_book = len(Document(str(book)).element.body.findall(".//" + qn("w:commentReference")))
        n_sec = sum(len(Document(str(p)).element.body.findall(".//" + qn("w:commentReference"))) for p in secs)
        if n_book != n_sec:
            out.append(_new_finding("comments_mismatch", "确定", name, "整份文件", "整本里的批注数与各节之和对不上",
                                    f"各节相加 {n_sec} 条，整本里 {n_book} 条——批注在生成整本时丢了，或整本里的批注被手工删了；"
                                    "核对后决定是重新生成整本，还是把节文件里的批注销掉", "", -2))
    except Exception as exc:  # noqa: BLE001
        out.append(_new_finding("check_error", "疑似", name, name, "「整本对账」这一项检查没能执行",
                                f"{type(exc).__name__}: {exc}——请人工核对", "", 999))
    return out


def inspect_work(work, scope=None) -> list[dict]:
    """整套工作区的只读检查：每册整本的文档检查 + 整本与节文件的对账。只读。"""
    work = Path(work)
    scope = {} if scope is None else scope
    body = work / "body"
    books = sorted(p for p in body.glob("整本-*.docx") if not p.name.startswith(("~$", ".~"))) \
        if body.is_dir() else []
    if not books:
        scope["no_book"] = True
        return []
    scope["books"] = [b.name for b in books]
    findings, totals = [], []
    for book in books:
        if not Path(str(book) + FP_SUFFIX).is_file():
            # 「未发现」只对已检查范围成立：没有指纹记录（别的版本生成的整本）就明说这一项查不了
            scope.setdefault("skipped", []).append(
                ("整本指纹对账", f"{book.name} 没有指纹记录（不是当前版本的引擎生成的），「整本被手工改过」这一项查不了"))
        vol = book.stem[len("整本-"):]
        sec_dir = body / vol if (body / vol).is_dir() else body
        secs = sorted(p for p in sec_dir.glob("*.docx") if not p.name.startswith(("整本-", "~$", ".~")))
        fs = inspect_docx(book, totals=totals, scope=scope)
        fs += _book_recon(book, secs)
        findings += fs
    scope["totals"] = len(totals)
    findings += _cross_totals(totals, len(books) > 1)
    if len(books) > 1:
        for f in findings:
            if f["kind"] != "amount_cross":
                f["where"] = f"{f['file']} › {f['where']}"
    return _sort_findings(findings)


def _md_cell(s, limit: int) -> str:
    s = re.sub(r"\s+", " ", str(s if s is not None else "")).replace("|", "\\|").strip()
    return s if len(s) <= limit else s[:limit - 1] + "…"


def _md_val(s, limit: int = 400) -> str:
    """规格表的取值格。一格里放不下时按「；」整段地让位、写明还有几段——不在半截处截断
    （「第 2…」这种截法读不出后面还有哪些节）。行里存的是完整取值，--json 结果里拿得到全量。
    没有分段的超长取值照旧截断加省略号。"""
    s = re.sub(r"\s+", " ", str(s if s is not None else "")).strip()
    if len(s) <= limit:
        return s.replace("|", "\\|")
    parts = s.split("；")
    keep = len(parts) - 1
    while keep >= 1:
        note = f"（另有 {len(parts) - keep} 段未列出，完整取值在 --json 结果里）"
        head = "；".join(parts[:keep])
        if len(head) + len(note) <= limit:
            return (head + note).replace("|", "\\|")
        keep -= 1
    return s[:limit - 1] + "…"


def _scope_text(scope: dict) -> str:
    """已检查范围——「未发现」只对这个范围成立，所以必须写出来。"""
    parts = []
    if scope.get("books"):
        parts.append(f"整本对账 {len(scope['books'])} 本（节文件新旧、指纹、批注数）")
    if scope.get("tables"):
        parts.append(f"表格 {scope['tables']} 张——其中 {scope.get('amount_tables', 0)} 张带金额列，"
                     f"核对了 {scope.get('amount_checks', 0)} 处（单价×数量、合计、大小写、税额）")
    else:
        parts.append("表格 0 张（没有可验算的金额表）")
    if scope.get("totals"):
        parts.append(f"总价前后一致（{scope['totals']} 处）")
    parts.append(f"图片 {scope.get('images', 0)} 张（拉伸、超宽、损坏、分辨率）")
    parts.append("文字（内部对账编号、占位写法、域报错、行距、修订痕迹）")
    parts.append("文档属性")
    if scope.get("comments"):
        parts.append(f"批注 {scope['comments']} 条")
    parts += [f"{k}未做（{v}）" for k, v in scope.get("skipped", [])]
    return "；".join(parts)


_BULK_KINDS = ("image_wide", "image_lowres", "image_stretch", "table_narrow", "table_wide")


def _drop_file(f: dict, multi: bool) -> str:
    """去掉位置里的文件名前缀（多文件时已加在前面，合并行的「位置」列自己会带）。"""
    prefix = f"{f['file']} › "
    return f["where"][len(prefix):] if multi and f["where"].startswith(prefix) else f["where"]


def _table_rows(findings: list[dict], multi: bool) -> list[dict]:
    """表里一行一条。同类、同一份文件、同一结论的版式问题满 3 处就并成一行——真实工作区回放时
    一本整本里有几百张图超出可用宽度，逐张列出来没法看；逐处的位置仍在 JSON 里。"""
    counts = Counter((f["kind"], f["file"], f["level"]) for f in findings if f["kind"] in _BULK_KINDS)
    out, seen = [], set()
    for f in findings:
        key = (f["kind"], f["file"], f["level"])
        if f["kind"] in _BULK_KINDS and counts[key] >= 3:
            if key in seen:
                continue
            seen.add(key)
            out.append({**f,
                        "where": (f"{f['file']} › " if multi else "") + f"共 {counts[key]} 处",
                        "basis": f"例：{_drop_file(f, multi)}——{f['basis']}；同类共 {counts[key]} 处，"
                               "逐处位置见 --json 结果"})
        else:
            out.append(f)
    return out


def format_findings(findings: list[dict], n_files: int, scope: dict) -> str:
    """检查结果的呈现：一行结论 + 一张表 + 定位行 + 已检查范围。回显给模型，模型照表转述。"""
    sure = sum(1 for f in findings if f["level"] == "确定")
    lines = []
    if findings:
        lines.append(f"[文档检查] {n_files} 份文件：查出 {len(findings)} 处（确定 {sure}、疑似 {len(findings) - sure}）；"
                     "本次只读，未改被检查的文件")
        lines += ["", "| 位置 | 问题 | 依据 | 确定/疑似 |", "|---|---|---|---|"]
        for f in _table_rows(findings, n_files > 1):
            lines.append(f"| {_md_cell(f['where'], 48)} | {_md_cell(f['what'], 30)} | {_md_cell(f['basis'], 90)} | {f['level']} |")
        # 定位行给修订用的坐标（T#R#C#／P#）。版式与图片类没有格坐标，位置写在表的「位置」列里
        anchors = [f["anchor"] for f in findings if f.get("anchor")]
        if anchors:
            lines += ["", "定位（要修订时用）：" + "；".join(anchors[:60])
                      + (f"；另有 {len(anchors) - 60} 处坐标未列出（--json 里有全部）" if len(anchors) > 60 else "")]
    else:
        lines.append(f"[文档检查] {n_files} 份文件：未发现问题；本次只读，未改被检查的文件")
    lines += ["", f"已检查：{_scope_text(scope)}"]
    return "\n".join(lines)


def cmd_inspect(args) -> None:
    """只读检查的命令行出口：单份/多份 docx，或 --work 整套工作区的整本。退出码恒为 0
    （只有参数错才非 0）——检查只负责发现，拦不拦是流程的事。--json 是结果落盘路径
    （不是开关），写之前过防呆；不给就不写任何文件。没给要检查的文件是参数错——
    检查 0 份文件再报「未发现问题」正是最危险的那种没查。"""
    scope: dict = {}
    if args.json:
        inputs = _work_docx(args.work) if args.work else [Path(x) for x in args.docx]
        guard_dump_path(args.json, inputs, ".json", "--json",
                        hint='inspect "标书.docx" --json work/review/inspect-result.json')
    if args.work:
        findings = inspect_work(args.work, scope)
        if scope.get("no_book"):
            print("[整本检查] 还没有整本 Word 文档——先完成第七步生成整本，再检查")
            return
        print(format_findings(findings, len(scope.get("books", [])), scope).replace("[文档检查]", "[整本检查]", 1))
        files = scope.get("books", [])
    elif not args.docx:
        die("没给要检查的文件——docx 路径直接跟在 inspect 后面（可多份），或用 --work 指工作目录；"
            "--json 后面跟的是结果落盘路径（.json），不是开关", 2)
    else:
        paths = [Path(x) for x in args.docx]
        for p in paths:
            if not p.is_file():
                die(f"文件不存在：{p}")
        findings = inspect_files(paths, final=not args.draft, scope=scope)
        print(format_findings(findings, len(paths), scope))
        files = [p.name for p in paths]
    if args.json:
        out = Path(args.json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({"at": time.strftime("%Y-%m-%d %H:%M"), "files": files, "findings": findings},
                                  ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[结果已落盘] {out}")


# ---------------------------------------------------------------- 公司版式（layout-read：版式规格读取，只读）
#
# 「按这个样式来做」场景的读取端：把用户给的 Word 版历史标书读成一张版式规格表
# （项目｜取值｜依据｜确定/疑似）。规格表是本体——用户确认后交给后续生成批落成模板。
# 只读：不改被读取的文件；原件不复制（历史标书里有价格与人名），只记路径与指纹。
# --out / --json 是显式的结果落盘（写之前过防呆，guard_dump_path）。
#
# 读法三种：机器读（页面设置这类整份文档只有一种取法，直接采信）；看一眼（正文、标题
# 这类可能混着当年招标要求的格式——逐段算生效格式后取多数，请用户过目）；对话定
# （封面字段这类意图项，机器只给候选）。
#
# 读数一律按生效格式算：直写 → 样式链 → docDefaults → 主题。真实历史标书常把字体
# 100% 直写在文字上、样式表里定义上百个样式却一个标题样式都不用，只看样式表会看错。
#
# 确定线（经验值，手头只有两份真实文档，样本多了要复核）：只有一种取值记确定；有
# 两种以上并存时最高占比过了 75% 才记确定，否则记疑似并把前三种并列进取值。只看
# 占比不看样本量——1/1 也记确定，样本数写在依据列里，由读表的人判断。
#
# 只收 .docx/.doc——PDF 读不准字体与磅值，不收。

_PT_CN = ((42.0, "初号"), (36.0, "小初"), (26.0, "一号"), (24.0, "小一"), (22.0, "二号"),
          (18.0, "小二"), (16.0, "三号"), (15.0, "小三"), (14.0, "四号"), (12.0, "小四"),
          (10.5, "五号"), (9.0, "小五"), (7.5, "六号"), (6.5, "小六"), (5.5, "七号"), (5.0, "八号"))
_LAYOUT_COEXIST = 0.75     # 两种以上取值并存时，最高占比过了这条线才记确定
_BROKEN = "读取时出错"      # 规格表「依据」列的这个值＝这一行不是读数，是没能读取的记录
_ANY = object()
_H_CH = re.compile(r"^第\s*[0-9一二三四五六七八九十百]+\s*[章部篇]")
_H_GOV1 = re.compile(r"^[一二三四五六七八九十百]+\s*[、]")
_H_GOV2 = re.compile(r"^[（(]\s*[一二三四五六七八九十百]+\s*[)）]")
_H_DEC2 = re.compile(r"^[1-9]\d*(?:\.\d+){1,3}(?![\d.])")
_H_DEC1 = re.compile(r"^[1-9]\d*\s*[、.](?!\d)")
_H_CIRC = re.compile(r"^[①-⑳]")
_TOC_TITLE = re.compile(r"^目\s*录$")
_TOC_DOTS = re.compile(r"\.{3,}\s*\d{1,4}\s*$")
_CAP = re.compile(r"^(图|表)\s*[0-9一二三四五六七八九十]+(\s*[-–－.．]\s*[0-9]+)?")
_F_PAGE = re.compile(r"(?<![A-Z])PAGE(?![A-Z])")
_F_NUMPAGES = re.compile(r"(?<![A-Z])NUMPAGES(?![A-Z])")
_F_PAGEREF = re.compile(r"(?<![A-Z])PAGEREF(?![A-Z])")
_VAR_HINT = ("日期", "时间", "编号", "项目名称", "采购人", "招标人", "投标人", "供应商",
             "代理机构", "正本", "副本")
_VAR_NUM = re.compile(r"\d{4}\s*年|\d{4}\s*[-/.]\s*\d{1,2}|\d{5,}")
_DEPTH_CN = ("一", "二", "三", "四", "五", "六", "七", "八", "九")
_SCHEME_VAL = {"chapter": "chapter（第X章/1.1）", "gov": "gov（一、/（一）/1.）",
               "decimal": "decimal（1/1.1）", "none": "none（不编号）"}
_TAG_NAME = {"chapter": "第X章", "gov": "一、/（一）", "decimal": "1.1"}
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_V_IMGDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_CJK_SW = (("w:kinsoku", "避头尾", True), ("w:overflowPunct", "标点溢出边界", True),
           ("w:topLinePunct", "行首标点压缩", False), ("w:autoSpaceDE", "中西文间距", True),
           ("w:autoSpaceDN", "数字间距", True), ("w:wordWrap", "西文单词中间换行", False))
_AL = {"center": "居中", "right": "右", "end": "右", "left": "左", "start": "左", "both": "两端",
       "distribute": "分散"}
_TBL_SIDES = (("top", "上"), ("left", "左"), ("bottom", "下"), ("right", "右"),
              ("insideH", "内横"), ("insideV", "内竖"))


def _cm_s(v) -> str:
    """厘米值显成短串：2.0→2、2.5→2.5。"""
    s = f"{round(float(v), 2):.2f}".rstrip("0").rstrip(".")
    return s or "0"


def _tw_cm(tw) -> str:
    """缇→厘米短串。"""
    return _cm_s(int(tw) / 567)


def _pt_label(pt) -> str:
    if pt is None:
        return "（未设）"
    for v, name in _PT_CN:
        if abs(pt - v) < 0.26:
            return f"{name}（{pt:g} 磅）"
    return f"{pt:g} 磅"


def _line_label(line, rule) -> str:
    if not line:
        return "单倍"
    v = int(line)
    if rule == "exact":
        return f"固定 {v / 20:g} 磅"
    if rule == "atLeast":
        return f"最小 {v / 20:g} 磅"
    m = v / 240.0
    if abs(m - 1.0) < 0.05:
        return "单倍"
    return f"{m:g} 倍"


def _pct(r: float) -> str:
    return f"{round(r * 100, 1):g}%"


# ---------------------------------------------------------------- 文字与结构的取法

def _text_lines(p_el) -> list[str]:
    """段落里的文字行。按换行符与文本框内段落切行；文本框在 mc:Choice 与 mc:Fallback
    里各有一份，只读 Choice——不去重会把同一行读成两遍（真实标书封面即此形态）。"""
    out: list[str] = []
    buf: list[str] = []

    def flush():
        s = "".join(buf).strip()
        buf.clear()
        if s:
            out.append(s)

    def walk(el):
        for ch in el:
            tag = ch.tag
            if tag == _MC_FALLBACK:
                continue
            if tag == qn("w:t"):
                buf.append(ch.text or "")
            elif tag == qn("w:br") and ch.get(qn("w:type")) != "page":
                flush()
            elif tag in (qn("w:p"), qn("w:txbxContent")):
                flush()
                walk(ch)
                flush()
            else:
                walk(ch)

    walk(p_el)
    flush()
    return out


def _el_text(el) -> str:
    if el.tag == qn("w:p"):
        return "".join(_text_lines(el))
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _layout_blocks(doc) -> list:
    """顶层块（段落/表格），外加一层内容控件（w:sdt）里的块——Word 生成的目录常包在
    内容控件里，按普通顶层子元素枚举会整段看不见。"""
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag in (qn("w:p"), qn("w:tbl")):
            out.append(ch)
        elif ch.tag == qn("w:sdt"):
            cont = ch.find(qn("w:sdtContent"))
            if cont is not None:
                out.extend(c for c in cont.iterchildren() if c.tag in (qn("w:p"), qn("w:tbl")))
    return out


def _style_name(p_el, doc) -> str:
    try:
        from docx.text.paragraph import Paragraph
        st = Paragraph(p_el, doc).style
        return (st.name or "").strip() if st is not None else ""
    except Exception:  # noqa: BLE001
        return ""


def _named_depth(name: str) -> int | None:
    m = re.match(r"^(?:heading|标题)\s*(\d)$", name, re.I)
    return min(int(m.group(1)), 9) if m else None


def _toc_level(name: str) -> int | None:
    m = re.match(r"^(?:toc|目录)\s*(\d)$", name, re.I)
    return int(m.group(1)) if m else None


def _has_field(el, pat) -> bool:
    if any(pat.search(t.text or "") for t in el.iter(qn("w:instrText"))):
        return True
    return any(pat.search(f.get(qn("w:instr")) or "") for f in el.iter(qn("w:fldSimple")))


def _has_tab_char(p_el) -> bool:
    return any(r.find(qn("w:tab")) is not None for r in p_el.iter(qn("w:r")))


def _breaks_after(el) -> bool:
    """块自身带着结束性的分页：段内分页符或分节符。"""
    if el.tag != qn("w:p"):
        return False
    if el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
        return True
    return any(br.get(qn("w:type")) == "page" for br in el.iter(qn("w:br")))


def _breaks_before(el) -> bool:
    """块自身设了段前分页。"""
    if el.tag != qn("w:p"):
        return False
    pb = el.find(f"{qn('w:pPr')}/{qn('w:pageBreakBefore')}")
    return pb is not None and pb.get(qn("w:val")) not in ("0", "false", "off")


def _block_sections(blocks) -> list[int]:
    """每个顶层块属于第几节（1 起）：带分节符的段落结束一节。"""
    out, cur = [], 1
    for el in blocks:
        out.append(cur)
        if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            cur += 1
    return out


def _is_toc_entry(el, doc) -> bool:
    """目录条目的四种形态：toc 样式、页码引用域、点线页码、制表符加页码。"""
    if el.tag != qn("w:p"):
        return False
    if _toc_level(_style_name(el, doc)) or _has_field(el, _F_PAGEREF):
        return True
    t = _el_text(el)
    if not t or len(t) > 60:
        return False
    if _TOC_DOTS.search(t):
        return True
    return _has_tab_char(el) and bool(re.search(r"\d\s*$", t))


def _is_list(p_el, doc) -> bool:
    for holder in [p_el] + [st.element for st in _style_chain(p_el, doc)]:
        ppr = holder.find(qn("w:pPr"))
        npr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if npr is not None:
            nid = npr.find(qn("w:numId"))
            return nid is None or nid.get(qn("w:val")) != "0"   # numId=0 是取消编号
    return False


def _head_tag(text: str) -> str | None:
    for pat, tag in ((_H_CH, "chapter"), (_H_GOV1, "gov"), (_H_GOV2, "gov"),
                     (_H_DEC2, "decimal"), (_H_DEC1, "decimal"), (_H_CIRC, "gov")):
        if pat.match(text):
            return tag
    return None


def _dot_depth(text: str) -> int:
    head = text.split()[0] if text.split() else ""
    return min(head.count(".") + 1, 9)


def _depth_of(tag: str, text: str, scheme: str) -> int:
    if tag == "chapter":
        return 1
    if tag == "decimal":
        return _dot_depth(text) if _H_DEC2.match(text) else (3 if scheme == "gov" else 2)
    if _H_GOV1.match(text):
        return 2 if scheme == "chapter" else 1
    if _H_GOV2.match(text):
        return 3 if scheme == "chapter" else 2
    return 4


def _count_pictures(el) -> int:
    """图片数：a:blip（drawing）与 v:imagedata（VML），兼容分支里的那份不计。"""
    n = 0
    for ch in el.iter():
        if ch.tag in (_A_NS + "blip", _V_IMGDATA) \
                and not any(a.tag == _MC_FALLBACK for a in ch.iterancestors()):
            n += 1
    return n


# ---------------------------------------------------------------- 生效格式

def _theme_faces(doc) -> dict:
    """主题字体：major/minor 的西文与东亚字形。东亚先取简体脚本字体（Hans），没有再取
    <a:ea>。"""
    out: dict = {}
    for part in _theme_parts(doc):
        try:
            from lxml import etree
            root = etree.fromstring(part.blob)
        except Exception:  # noqa: BLE001
            continue
        for tag, key in ((_A_NS + "majorFont", "major"), (_A_NS + "minorFont", "minor")):
            for scheme in root.iter(tag):
                if key + "Latin" not in out:
                    el = scheme.find(_A_NS + "latin")
                    t = (el.get("typeface") or "").strip() if el is not None else ""
                    if t:
                        out[key + "Latin"] = t
                if key + "Ea" not in out:
                    hans = [f.get("typeface") for f in scheme.findall(_A_NS + "font")
                            if f.get("script") == "Hans" and (f.get("typeface") or "").strip()]
                    el = scheme.find(_A_NS + "ea")
                    ea = (el.get("typeface") or "").strip() if el is not None else ""
                    if hans:
                        out[key + "Ea"] = hans[0]
                    elif ea:
                        out[key + "Ea"] = ea
    return out


def _doc_default_fonts(doc) -> dict:
    try:
        dd = doc.styles.element.find(qn("w:docDefaults"))
        rpr = dd.find(qn("w:rPrDefault")).find(qn("w:rPr")) if dd is not None else None
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        return {"ea": rf.get(qn("w:eastAsia")) if rf is not None else None,
                "ascii": rf.get(qn("w:ascii")) if rf is not None else None}
    except Exception:  # noqa: BLE001
        return {"ea": None, "ascii": None}


def _fonts_of(holders, theme, defaults) -> tuple[str | None, str | None]:
    """按 holders 顺序解析 (东亚字体, 西文字体)。同一个 rFonts 里主题属性优先于显式
    字体（OOXML 的规定；2.1.6 中文字体链路批坐实的正是「显式黑体被主题顶掉」）。"""
    ea = ascii_ = None
    for rpr in holders:
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if rf is None:
            continue
        if ea is None:
            t = rf.get(qn("w:eastAsiaTheme"))
            ea = ((theme.get(("major" if "major" in t else "minor") + "Ea")) if t else None) \
                or rf.get(qn("w:eastAsia")) or None
        if ascii_ is None:
            t = rf.get(qn("w:asciiTheme"))
            ascii_ = ((theme.get(("major" if "major" in t else "minor") + "Latin")) if t else None) \
                or rf.get(qn("w:ascii")) or None
        if ea and ascii_:
            break
    return (ea or defaults.get("ea"), ascii_ or defaults.get("ascii"))


def _spacing_side(pprs, pt_attr: str, lines_attr: str, auto_attr: str) -> str | None:
    """段前或段后的生效值：自动 > 行 > 磅（同时出现时前者顶掉后者），样式链取第一个
    有设置的。没设返回 None——None 不进格式化（真实标书回放抓到的崩点）。"""
    for ppr in pprs:
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            continue
        if sp.get(qn(auto_attr)) in ("1", "true", "on"):
            return "自动"
        if sp.get(qn(lines_attr)) is not None:
            return f"{int(sp.get(qn(lines_attr))) / 100:g} 行"
        if sp.get(qn(pt_attr)) is not None:
            return f"{int(sp.get(qn(pt_attr))) / 20:g} 磅"
    return None


def _cjk_label(pprs) -> str:
    """中文版式：逐个开关按样式链取生效值，与 Word 缺省不同的点名。这些开关写在
    段落和样式里（几乎不写进 docDefaults），所以必须逐段走样式链。"""
    off, on = [], []
    for tag, label, dflt in _CJK_SW:
        v = None
        for ppr in pprs:
            el = ppr.find(qn(tag))
            if el is not None:
                v = el.get(qn("w:val")) not in ("0", "false", "off")
                break
        if v is None:
            v = dflt
        if v and not dflt:
            on.append(label)
        elif dflt and not v:
            off.append(label)
    if not off and not on:
        return "标准"
    parts = []
    if off:
        parts.append("关闭了" + "、".join(off))
    if on:
        parts.append("开启了" + "、".join(on))
    return "；".join(parts)


def _dominant_run(p_el):
    """段落里字数最多的那个 run——外观以它为准（上标、空格这类零碎 run 不代表性）。"""
    best, best_n = None, 0
    for r in p_el.iter(qn("w:r")):
        n = len("".join(t.text or "" for t in r.iter(qn("w:t"))).strip())
        if n > best_n:
            best, best_n = r, n
    return best


def _default_para_style_id(doc) -> str | None:
    for st in doc.styles.element.findall(qn("w:style")):
        if st.get(qn("w:type")) in (None, "paragraph") and st.get(qn("w:default")) in ("1", "true"):
            return st.get(qn("w:styleId"))
    return None


def _in_default_style(el, doc) -> bool:
    """段落挂在正文默认样式上（没写段落样式，或写的就是文档的默认段落样式）。"""
    ppr = el.find(qn("w:pPr"))
    ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
    return ps is None or ps.get(qn("w:val")) == _default_para_style_id(doc)


def _unstyled_head_note(doc, heads) -> str:
    """标题一段都没挂命名样式时，外观从哪来说哪。「全是直接格式」不能写死——真实回放里有一份
    技术标全篇不挂样式、文字上只写了颜色，标题的外观取自正文默认样式；写死的那句让读表的人
    查到文字上没有格式后自相矛盾。"""
    direct = sum(1 for h in heads if _head_direct(h["el"]))
    if direct == len(heads):
        return "标题全是直接格式，读数按生效格式算"
    if direct == 0 and all(_in_default_style(h["el"], doc) for h in heads):
        return "标题都没挂样式，外观取自正文默认样式，读数按生效格式算"
    return f"标题都没挂标题样式，直接格式 {direct} 段、其余取自段落样式，读数按生效格式算"


def _head_direct(el) -> bool:
    """外观是不是直接写在文字上（字体、字号、加粗任一项写在主导 run 上）。"""
    r = _dominant_run(el)
    rpr = r.find(qn("w:rPr")) if r is not None else None
    return rpr is not None and any(rpr.find(qn(t)) is not None for t in ("w:rFonts", "w:sz", "w:b"))


def _para_fmt(p_el, doc, theme, defaults) -> dict:
    """一个段落的生效格式。字体字号加粗取主导 run（按字数）；段落属性走样式链，
    直写优先。没设对齐就是左对齐（Word 缺省），不替文档往「两端」上猜。"""
    chain = _style_chain(p_el, doc)
    best = _dominant_run(p_el)
    rprs = [st.element.find(qn("w:rPr")) for st in chain]
    if best is not None and best.find(qn("w:rPr")) is not None:
        rprs.insert(0, best.find(qn("w:rPr")))
    ea, ascii_ = _fonts_of(rprs, theme, defaults)
    default_pt = _doc_default_pt(doc) or 10.5
    pt = _run_size_pt(best, chain, default_pt) if best is not None else default_pt
    bold = None
    for rpr in rprs:
        b = rpr.find(qn("w:b")) if rpr is not None else None
        if b is not None:
            bold = b.get(qn("w:val")) not in ("0", "false", "none")
            break
    # 段落属性走样式链，链尾是文档默认值——段后 10 磅、行距 1.15 倍常写在 docDefaults 里（Word 英文
    # 模板的常态），样式链上都没设时它就是生效值；只走样式链会把这两种读成「无」「单倍」。
    # 文档默认值里的 0 等于没设（Word 的内建缺省就是 0）：摘掉，免得和样式里显式写的 0 混为一谈
    dd = doc.styles.element.find(qn("w:docDefaults"))
    dd_ppr = dd.find(qn("w:pPrDefault")).find(qn("w:pPr")) if dd is not None else None
    if dd_ppr is not None:
        sp0 = dd_ppr.find(qn("w:spacing"))
        if sp0 is not None and ("0" in (sp0.get(qn("w:before")), sp0.get(qn("w:after")))):
            from copy import deepcopy
            dd_ppr = deepcopy(dd_ppr)
            sp0 = dd_ppr.find(qn("w:spacing"))
            for a in ("w:before", "w:after", "w:beforeLines", "w:afterLines"):
                if sp0.get(qn(a)) == "0":
                    del sp0.attrib[qn(a)]
    pprs = [p for p in [p_el.find(qn("w:pPr"))] + [st.element.find(qn("w:pPr")) for st in chain] + [dd_ppr]
            if p is not None]
    line = rule = None
    for ppr in pprs:
        sp = ppr.find(qn("w:spacing"))
        if sp is not None and sp.get(qn("w:line")):
            line, rule = sp.get(qn("w:line")), sp.get(qn("w:lineRule"))
            break
    ind_label, left_label = "无", "0"
    for ppr in pprs:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            continue
        flc, fl = ind.get(qn("w:firstLineChars")), ind.get(qn("w:firstLine"))
        if flc:
            ind_label = f"首行 {int(flc) / 100:g} 字"
        elif fl:
            # 绝对值缩进按生效字号折成字（420–480 缇 @12 磅 ≈ 2 字，与 firstLineChars 同义）
            tw = int(fl)
            chars = round(tw / (pt * 20)) if pt else None
            ind_label = (f"首行 {chars} 字" if chars and 1 <= chars <= 4
                         else f"首行 {_tw_cm(tw)} 厘米")
        if ind_label == "无" and (ind.get(qn("w:hangingChars")) or ind.get(qn("w:hanging"))):
            ind_label = "悬挂缩进"
        lc, lv = ind.get(qn("w:leftChars")), ind.get(qn("w:left"))
        if lc:
            left_label = f"{int(lc) / 100:g} 字"
        elif lv:
            left_label = f"{_tw_cm(lv)} 厘米"
        break
    al = "左"
    for ppr in pprs:
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            al = _AL.get(jc.get(qn("w:val")), "左")
            break
    return {"ea": ea, "ascii": ascii_, "pt": pt, "bold": bool(bold),
            "line": _line_label(line, rule),
            "before": _spacing_side(pprs, "w:before", "w:beforeLines", "w:beforeAutospacing"),
            "after": _spacing_side(pprs, "w:after", "w:afterLines", "w:afterAutospacing"),
            "ind": ind_label, "left": left_label, "al": al, "cjk": _cjk_label(pprs)}


def _ba_label(k) -> str:
    """段前段后标签；单边未设就只写设了的那一边。"""
    b, a = k
    if b is None and a is None:
        return "无"
    out = []
    if b is not None:
        out.append("段前自动" if b == "自动" else f"段前 {b}")
    if a is not None:
        out.append("段后自动" if a == "自动" else f"段后 {a}")
    return "、".join(out)


def _sp_label(k) -> str:
    return f"行距 {k[0]}、{_ba_label((k[1], k[2]))}"


def _fmt_str(f) -> str:
    s = f"中文{f['ea'] or '（默认）'}"
    if f.get("ascii"):
        s += f"、西文 {f['ascii']}"
    if f.get("pt"):
        s += f"｜{_pt_label(f['pt'])}"
    if f.get("bold"):
        s += "｜加粗"
    return s + f"｜{f['al']}"


def _font_key(f) -> tuple:
    return (f["ea"], f["ascii"], f["pt"], f["bold"], f["al"])


def _fk_label(k) -> str:
    """字体键元组 → 取值标签（取多数时标签函数拿到的是键，不是格式字典）。"""
    ea, asc, pt, bold, al = k
    s = f"中文{ea or '（默认）'}"
    if asc:
        s += f"、西文 {asc}"
    if pt:
        s += f"｜{_pt_label(pt)}"
    if bold:
        s += "｜加粗"
    return s + f"｜{al}"


# ---------------------------------------------------------------- 取多数

def _vote(keys, label_of, basis_name):
    """逐段取值后取多数，返回 (取值, 依据, 是否确定)。只有一种取值记确定；两种以上
    并存时最高占比过了 _LAYOUT_COEXIST 才记确定，否则记疑似并把前三种并列进取值。"""
    cnt = Counter(keys)
    if not cnt:
        return None
    common = cnt.most_common()
    total = sum(cnt.values())
    ratio = common[0][1] / total
    basis = f"{basis_name} {common[0][1]}/{total}（{_pct(ratio)}）"
    if len(cnt) == 1 or ratio >= _LAYOUT_COEXIST:
        return label_of(common[0][0]), basis, True
    alts = " ／ ".join(label_of(k) for k, _ in common[:3])
    return alts, basis + f"，{len(cnt)} 种并存", False


class _Rows:
    """规格表的行：(组, 项目, 取值, 依据, 确定/疑似)。"""

    def __init__(self):
        self.items: list[tuple[str, str, str, str, str]] = []

    def add(self, grp, item, val, basis, sure):
        self.items.append((grp, item, val, basis, "确定" if sure else "疑似"))

    def vote(self, grp, item, keys, label_of, basis_name):
        v = _vote(keys, label_of, basis_name)
        if v is not None:
            self.add(grp, item, *v)
        return v


# ---------------------------------------------------------------- 区域切分与编号

def _split_regions(blocks, doc) -> tuple[list, list, list, list]:
    """封面区 / 前置页 / 目录区 / 正文区。封面=首页（到第一个分页或分节符为止，没有
    分页就到目录为止）；目录区=「目录」标题（或首条目录条目）到末条目录条目。真实
    标书的封面和目录之间夹着一页评分索引表：它既不是封面也不是目录，落在前置页，
    三个区都不进正文统计。"""
    paras = [i for i, el in enumerate(blocks) if el.tag == qn("w:p")]
    first_head = next((i for i in paras if _named_depth(_style_name(blocks[i], doc))), None)
    limit = first_head if first_head is not None else len(blocks)
    toc_title = None
    for i in paras:
        if i >= limit:
            break
        name = _style_name(blocks[i], doc).replace(" ", "").lower()
        if _TOC_TITLE.match(_el_text(blocks[i])) or name in ("tocheading", "目录标题"):
            toc_title = i
            break
    toc_idx = [i for i in paras if i < limit and _is_toc_entry(blocks[i], doc)]
    if toc_title is not None:
        # 还有一种形态：标题之后是纯文字条目（不带页码、不带样式），到第一个分页符为止
        last, j = toc_title, toc_title + 1
        while j < len(blocks):
            el = blocks[j]
            if el.tag != qn("w:p") or _breaks_after(el) or _breaks_before(el) \
                    or _named_depth(_style_name(el, doc)):
                break
            txt = _el_text(el)
            if txt:
                if len(txt) > 40 or not _head_tag(_text_lines(el)[0]):
                    break
                last = j
            j += 1
        toc_idx = sorted(set(toc_idx + list(range(toc_title + 1, last + 1))))
    if toc_title is None and not toc_idx:
        cut = 0
        for i, el in enumerate(blocks[:limit]):
            if _breaks_before(el):
                cut = i
                break
            if _breaks_after(el):
                cut = i + 1
                break
        return blocks[:cut], [], [], blocks[cut:]
    toc_start = min(x for x in (toc_title, toc_idx[0] if toc_idx else None) if x is not None)
    toc_end = max(x for x in (toc_title, toc_idx[-1] if toc_idx else None) if x is not None)
    cover_end = toc_start
    for i, el in enumerate(blocks[:toc_start]):
        if _breaks_before(el):
            cover_end = i
            break
        if _breaks_after(el):
            cover_end = i + 1
            break
    return blocks[:cover_end], blocks[cover_end:toc_start], blocks[toc_start:toc_end + 1], blocks[toc_end + 1:]


def _numbering_root(doc):
    for part in doc.part.package.iter_parts():
        if str(getattr(part, "partname", "")).startswith("/word/numbering"):
            try:
                from lxml import etree
                return etree.fromstring(part.blob)
            except Exception:  # noqa: BLE001
                return None
    return None


def _n_styles(doc) -> int:
    """包里样式表定义了多少个样式；没有样式部件、读不出来都算 0。直接读包，不走 doc.styles——
    缺样式部件时它会就地补一份默认的，那就看不出「这个文件本来没有样式」了。只认
    /word/styles.xml 本身：同目录的 stylesWithEffects.xml 是它的副本，不能拿来充数。"""
    for part in doc.part.package.iter_parts():
        if re.fullmatch(r"/word/styles\d*\.xml", str(getattr(part, "partname", ""))):
            try:
                from lxml import etree
                return len(etree.fromstring(part.blob).findall(qn("w:style")))
            except Exception:  # noqa: BLE001
                return 0
    return 0


def _num_lvl(root, num_id, ilvl: int):
    num = root.find(f"{qn('w:num')}[@{qn('w:numId')}='{num_id}']")
    aid = num.find(qn("w:abstractNumId")) if num is not None else None
    absn = None
    if aid is not None:
        absn = root.find(f"{qn('w:abstractNum')}[@{qn('w:abstractNumId')}='{aid.get(qn('w:val'))}']")
    return absn.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='{ilvl}']") if absn is not None else None


def _num_info(p_el, doc, root) -> tuple[str, str] | None:
    """段落生效的编号定义 (numFmt, lvlText)：直写 numPr → 样式链上的 numPr。"""
    if root is None:
        return None
    for holder in [p_el] + [st.element for st in _style_chain(p_el, doc)]:
        ppr = holder.find(qn("w:pPr"))
        npr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if npr is None:
            continue
        nid = npr.find(qn("w:numId"))
        if nid is None or nid.get(qn("w:val")) == "0":
            return None
        il = npr.find(qn("w:ilvl"))
        lvl = _num_lvl(root, nid.get(qn("w:val")), int(il.get(qn("w:val"))) if il is not None else 0)
        if lvl is None:
            return ("", "")
        fmt = lvl.find(qn("w:numFmt"))
        txt = lvl.find(qn("w:lvlText"))
        return ((fmt.get(qn("w:val")) or "decimal") if fmt is not None else "decimal",
                (txt.get(qn("w:val")) or "") if txt is not None else "")
    return None


def _num_label(fmt: str, txt: str) -> str:
    if fmt == "bullet":
        return "项目符号"
    shown = re.sub(r"%\d", "1", txt).strip() or "1"
    if fmt.startswith("chinese") and shown.startswith("1"):
        shown = "一" + shown[1:]
    return f"编号 {shown}"


def _classify_num(fmt: str, txt: str) -> str:
    if "第" in txt and re.search(r"[章部篇节]", txt):
        return "chapter"
    if fmt.startswith("chinese") or txt.startswith("（") or txt.startswith("("):
        return "gov"
    return "decimal"


def _heading_auto_scheme(doc, style_names) -> tuple[str | None, str]:
    """标题样式绑定的自动编号（多级列表按 w:pStyle 挂到标题样式上）。取该列表第 0 级
    的写法定形态。"""
    root = _numbering_root(doc)
    if root is None or not style_names:
        return None, ""
    ids = set()
    for nm in style_names:
        try:
            ids.add(doc.styles[nm].style_id)
        except KeyError:
            pass
    for absn in root.findall(qn("w:abstractNum")):
        for lvl in absn.findall(qn("w:lvl")):
            ps = lvl.find(qn("w:pStyle"))
            if ps is None or ps.get(qn("w:val")) not in ids:
                continue
            top = absn.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='0']")
            if top is None:
                top = lvl
            tf = top.find(qn("w:lvlText"))
            fm = top.find(qn("w:numFmt"))
            txt = (tf.get(qn("w:val")) or "") if tf is not None else ""
            fmt = (fm.get(qn("w:val")) or "decimal") if fm is not None else "decimal"
            return _classify_num(fmt, txt), re.sub(r"%\d", "1", txt)
    return None, ""


# ---------------------------------------------------------------- 页面

# 节里没写纸张或页边距时 Word 按这组缺省值排版（ECMA-376 §17.6.26 pgSz、§17.6.22 pgMar）。
# 别的转换器转出来的 docx 会整节空着（真机上一份 26MB 的 .doc 手动转出来的就是这样），
# 原先在这里取属性取到 None 直接崩，整份读不出来。
_PG_DEF = {"w": 12240, "h": 15840, "top": 1440, "bottom": 1440, "left": 1800, "right": 1800, "gutter": 0}


def _sect_geom(sect) -> dict:
    """一节的纸张与页边距（缇）。直写优先；带单位的写法（如 21cm）走 python-docx 的换算；
    没写的取 Word 缺省值，并在 no_sz / no_mar 里记下「不是从文件里读出来的」。"""
    sp = sect._sectPr
    sz, mar = sp.find(qn("w:pgSz")), sp.find(qn("w:pgMar"))

    def tw(el, attr, pd, dflt):
        v = el.get(qn(attr)) if el is not None else None
        if v is not None and v.isdigit():
            return int(v)
        return int(pd.twips) if pd is not None else dflt

    g = {"w": tw(sz, "w:w", sect.page_width, _PG_DEF["w"]),
         "h": tw(sz, "w:h", sect.page_height, _PG_DEF["h"]),
         "top": tw(mar, "w:top", sect.top_margin, _PG_DEF["top"]),
         "bottom": tw(mar, "w:bottom", sect.bottom_margin, _PG_DEF["bottom"]),
         "left": tw(mar, "w:left", sect.left_margin, _PG_DEF["left"]),
         "right": tw(mar, "w:right", sect.right_margin, _PG_DEF["right"]),
         "gutter": tw(mar, "w:gutter", sect.gutter, 0)}
    g["no_sz"] = sz is None or sz.get(qn("w:w")) is None
    g["no_mar"] = mar is None or mar.get(qn("w:left")) is None
    return g


def _paper_label(g: dict) -> str:
    landscape = g["w"] > g["h"]
    w, h = (g["h"], g["w"]) if landscape else (g["w"], g["h"])
    name = next((n for pw, ph, n in ((11906, 16838, "A4"), (16838, 23811, "A3"),
                                     (8391, 11906, "A5"), (12240, 15840, "Letter"))
                  if abs(w - pw) <= 60 and abs(h - ph) <= 60), "")
    size = f"{name} " if name else f"{_tw_cm(w)}×{_tw_cm(h)} 厘米 "
    return size + ("横向" if landscape else "纵向")


def _grid_label(sect) -> str:
    g = sect._sectPr.find(qn("w:docGrid"))
    if g is None:
        return "未设置（无网格）"
    kind = g.get(qn("w:type")) or "default"
    lp = g.get(qn("w:linePitch"))
    if kind == "default":
        # Word 每份文档都带 linePitch，没写 type 就是没有网格——它只是行距基准
        return "无网格" + (f"（行距基准 {int(lp) / 20:g} 磅）" if lp and lp.isdigit() else "")
    if kind == "snapToChars":
        return "字符网格对齐"
    lines = None
    if lp and lp.isdigit():
        gm = _sect_geom(sect)
        lines = max(1, round((gm["h"] - gm["top"] - gm["bottom"]) / int(lp)))
    if kind == "lines":
        return f"行网格（每页约 {lines} 行）" if lines else "行网格"
    if not lines:
        return "行和字符网格"
    cs = g.get(qn("w:charSpace"))
    return f"行和字符网格（每页约 {lines} 行、字符间距 {cs}）" if cs else f"行和字符网格（每页约 {lines} 行）"


def _sec_group_label(nums: list) -> str:
    """一组节号的说法：第 2、4、6 节等 12 节（超过三个只列前三个）；第 1、2 节；第 1 节。"""
    shown = "、".join(str(n) for n in nums[:3])
    return f"第 {shown} 节等 {len(nums)} 节" if len(nums) > 3 else f"第 {shown} 节"


def _page_diffs(geoms, body_no, g) -> tuple[str, int]:
    """页面设置差异：同一种差异的节归成一组，返回 (取值, 与正文节不同的节数)。取值给全——
    排表时一行放不下由 _md_val 按段收（真实回放：119 节里 59 个图页节设置完全相同，逐节罗列
    被拦腰截成「第 2…」，读表的人既看不到后面有哪些节、也不知道被截掉了多少）。"""
    groups: dict[str, list] = {}
    for i, gi in enumerate(geoms, 1):
        if i == body_no:
            continue
        sides = [f"{n} {_tw_cm(gi[k])}" for n, k in (("上", "top"), ("下", "bottom"),
                                                      ("左", "left"), ("右", "right")) if gi[k] != g[k]]
        if (gi["w"], gi["h"]) != (g["w"], g["h"]):
            sides.append(f"纸张 {_paper_label(gi)}")
        if sides:
            desc = "、".join(sides) + ("（这一节没写页面设置，取的是缺省值）" if gi["no_sz"] or gi["no_mar"] else "")
            groups.setdefault(desc, []).append(i)
    total = sum(len(v) for v in groups.values())
    if not total:
        return "全篇各节相同", 0
    return "；".join(f"{_sec_group_label(nums)}：{desc}" for desc, nums in groups.items()) + "（与正文节不同）", total


def _layout_page(ctx, rows, secs, info) -> dict:
    # 文档一个节属性都没有（个别转换器的产物）：整份算一节、页面设置全是缺省值，不在这里崩
    geoms = [_sect_geom(s) for s in secs] or [dict(_PG_DEF, no_sz=True, no_mar=True)]
    per = info["body_per_sec"]
    body_no = max(per, key=per.get) if per else len(secs)   # 正文候选最多的那节=正文节
    body_no = min(max(body_no, 1), len(geoms))
    g = geoms[body_no - 1]
    # 多节文档里点明「正文节」是怎么挑的——节多的文档里「正文所在第 103 节」被读成全文有 103 节过
    where = f"正文所在第 {body_no} 节" + ("（正文段最多的一节）" if len(geoms) > 1 else "")
    rows.add("页面", "纸张", _paper_label(g),
             f"{where}，{_tw_cm(g['w'])}×{_tw_cm(g['h'])} 厘米"
             + ("；这一节没写纸张，取的是 Word 缺省值" if g["no_sz"] else ""), not g["no_sz"])
    rows.add("页面", "页边距",
             f"上 {_tw_cm(g['top'])}、下 {_tw_cm(g['bottom'])}、左 {_tw_cm(g['left'])}、右 {_tw_cm(g['right'])} 厘米"
             + (f"（装订线 {_tw_cm(g['gutter'])} 厘米）" if g["gutter"] else ""),
             f"正文所在第 {body_no} 节" + ("；这一节没写页边距，取的是 Word 缺省值" if g["no_mar"] else ""),
             not g["no_mar"])
    diffs, n_diff = _page_diffs(geoms, body_no, g)
    rows.add("页面", "页面设置差异", diffs,
             f"共 {len(geoms)} 节" + (f"，其中 {n_diff} 节与正文节不同" if n_diff else "")
             + "，逐节比对纸张与页边距", True)
    rows.add("页面", "对称页边距", "是（双面打印时装订线落在内侧）"
             if ctx["doc"].settings.element.find(qn("w:mirrorMargins")) is not None else "否", "文档设置", True)
    rows.add("页面", "文档网格", _grid_label(secs[body_no - 1]) if secs else "无网格",
             f"正文所在第 {body_no} 节", True)
    return {"body_no": body_no, "geoms": geoms,
            "text_tw": [max(1, x["w"] - x["left"] - x["right"] - x["gutter"]) for x in geoms]}


def _p_bottom_border(p_el) -> bool:
    ppr = p_el.find(qn("w:pPr"))
    pbdr = ppr.find(qn("w:pBdr")) if ppr is not None else None
    return pbdr is not None and pbdr.find(qn("w:bottom")) is not None


def _header_desc(sec, ctx) -> str | None:
    texts, sizes, align, border, imgs = [], [], None, False, 0
    for p in sec.header.paragraphs:
        lines = _text_lines(p._p)
        if lines:
            f = _para_fmt(p._p, ctx["doc"], ctx["theme"], ctx["defaults"])
            texts.append(lines[0])
            sizes.append(f["pt"])
            align = align or f["al"]
        border = border or _p_bottom_border(p._p)
        imgs += _count_pictures(p._p)
    if not texts and not imgs:
        return None
    val = ""
    if texts:
        val = f"文字「{texts[0][:24]}」｜{_pt_label(Counter(sizes).most_common(1)[0][0])}｜{align or '左'}"
    if imgs:
        val += ("、含图（logo）" if texts else "含图（logo）")
    if border:
        val += "、带横线"
    return val


def _odd_even(doc, secs) -> bool:
    el = doc.settings.element.find(qn("w:evenAndOddHeaders"))
    if el is None or el.get(qn("w:val")) in ("0", "false", "off"):
        return False
    return any(not hf.is_linked_to_previous for sec in secs
               for hf in (sec.even_page_header, sec.even_page_footer))


def _para_hpos(p_el, doc, text_tw: int) -> str:
    """页码在页面上的横向位置：对齐方式；没设对齐就看制表位（「制表符 + 页码」的常见
    写法，制表位可能定义在页脚样式里）。"""
    pprs = [p for p in [p_el.find(qn("w:pPr"))] + [st.element.find(qn("w:pPr")) for st in _style_chain(p_el, doc)]
            if p is not None]
    for ppr in pprs:
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            v = jc.get(qn("w:val"))
            return "居中" if v == "center" else ("右" if v in ("right", "end") else "左")
    if _has_tab_char(p_el):
        stops = []
        for ppr in pprs:
            tabs = ppr.find(qn("w:tabs"))
            if tabs is not None:
                stops += [(int(tb.get(qn("w:pos"), "0")), tb.get(qn("w:val")) or "left")
                          for tb in tabs.findall(qn("w:tab"))]
        stops = sorted(s for s in stops if s[0] <= text_tw * 1.05)
        if stops:
            pos, kind = stops[-1]
            if kind == "right" or pos >= text_tw * 0.8:
                return "右"
            if kind == "center" or text_tw * 0.2 < pos < text_tw * 0.8:
                return "居中"
    return "左"


def _layout_header_footer(ctx, rows, secs, page) -> None:
    doc = ctx["doc"]
    tail = "；奇偶页不同" if _odd_even(doc, secs) else ""
    heads = [(i, d) for i, sec in enumerate(secs, 1)
             if not sec.header.is_linked_to_previous and (d := _header_desc(sec, ctx))]
    if not heads:
        rows.add("页面", "页眉", "无" + tail, f"全部 {len(secs)} 节页眉", True)
    else:
        first = heads[0][0]
        spans: list = []
        for i, d in heads:
            if spans and spans[-1][2] == d and spans[-1][1] == i - 1:
                spans[-1][1] = i
            else:
                spans.append([i, i, d])
        val = "；".join(f"第 {a}{'–' + str(b) if b > a else ''} 节：{d}" for a, b, d in spans)
        if first > 1:
            val += f"（第 1–{first - 1} 节无）"
        rows.add("页面", "页眉", val + tail, "逐节读页眉（继承前一节的不重复计）", True)
    pns = []   # (节号, 取值, 位置)——逐节读，以正文所在节为准
    for i, sec in enumerate(secs, 1):
        for hf, where in ((sec.footer, "页脚"), (sec.header, "页眉"), (sec.first_page_footer, "首页页脚")):
            if hf.is_linked_to_previous:
                continue
            for p in hf.paragraphs:
                # 域一律按后代级检测——run 可能嵌在内容控件或图形容器里，fldSimple 的
                # instr 还是属性（真实标书两种形态都有，按段落直接子级扫会读成「无页码」）
                if not _has_field(p._p, _F_PAGE):
                    continue
                literal = "".join(x.text or "" for x in p._p.iter(qn("w:t")))
                if "第" in literal and "页" in literal:
                    style = "第 X 页"
                elif "—" in literal or "–" in literal:
                    style = "— X —"
                else:
                    style = "纯数字"
                if _has_field(p._p, _F_NUMPAGES):
                    style += "、共 Y 页"
                fmt = _para_fmt(p._p, doc, ctx["theme"], ctx["defaults"])
                # 页码是纯数字（域前后没有文字）时段落里没有文字 run，_para_fmt 取不到主导 run、
                # 会退到文档默认字号——字号要看域 run 自己写的
                for r in p._p.iter(qn("w:r")):
                    szr = r.find(qn("w:rPr")).find(qn("w:sz")) if r.find(qn("w:rPr")) is not None else None
                    if szr is not None and szr.get(qn("w:val")):
                        fmt = dict(fmt, pt=int(szr.get(qn("w:val"))) / 2)
                        break
                pos = _para_hpos(p._p, doc, page["text_tw"][i - 1])
                pns.append((i, f"{style}、{where}{pos}｜{_pt_label(fmt['pt'])}", f"{where}{pos}"))
                break
            if pns and pns[-1][0] == i:
                break
    if not pns:
        rows.add("页面", "页码写法", "无页码", f"全部 {len(secs)} 节页眉页脚", True)
    else:
        body_no = page["body_no"]
        main = next((x for x in pns if x[0] == body_no), None)
        if main is None:
            # 正文节没单独定义页眉页脚：沿用前面最近一个定义（Word 的「链接到前一节」）；前面一个
            # 都没有，就是正文节没有页码。改前这里落到第一个有页码的节、依据却照写「正文所在第 N 节」
            # ——真实回放：正文在第 103 节，依据写成「正文所在第 1 节」，同一张表里自相矛盾
            before = [x for x in pns if x[0] < body_no]
            main = before[-1] if before else pns[0]
        first = pns[0][0]
        if main[0] == body_no:
            head_txt = f"页码域（正文所在第 {body_no} 节"
            val = main[1]
        elif main[0] < body_no:
            head_txt = f"页码域（正文所在第 {body_no} 节沿用第 {main[0]} 节的页脚"
            val = main[1]
        else:
            head_txt = f"页码域（正文所在第 {body_no} 节没有页码；下面是第 {main[0]} 节的写法"
            val = f"正文所在节没有页码（第 {main[0]} 节起才有：{main[1]}）"
        others = [x for x in pns if x[0] != main[0] and x[2] != main[2]]
        if others and main[0] <= body_no:
            val += "（" + "、".join(f"第 {x[0]} 节{x[2]}" for x in others[:3]) + "）"
        basis = head_txt
        basis += (f"；第 {first} 节起才有，第 1 节无）" if first == 2
                  else f"；第 {first} 节起才有，第 1–{first - 1} 节无）") if first > 1 else "）"
        rows.add("页面", "页码写法", val + tail, basis, True)
    other = None
    for i, sec in enumerate(secs, 1):
        if sec.footer.is_linked_to_previous:
            continue
        for p in sec.footer.paragraphs:
            if _el_text(p._p) and not _has_field(p._p, _F_PAGE):
                other = (i, _el_text(p._p))
                break
        if other:
            break
    rows.add("页面", "页脚其他文字", f"文字「{other[1][:24]}」" if other else "无",
             f"第 {other[0]} 节页脚" if other else "全部节页脚", True)


# ---------------------------------------------------------------- 正文

def _collect_body(ctx, body, sec_of_el) -> dict:
    """正文区分类：标题候选（命名标题样式，或自带编号的短段）、正文候选（30 字以上的
    非列表段）、列表段。目录条目不在这里——区域切分时已经分出去。"""
    doc = ctx["doc"]
    num_root = _numbering_root(doc)
    info = {"body_fmts": [], "body_per_sec": Counter(), "heads": [], "list_infos": [],
            "n_list": 0, "gaps": Counter(), "fonts": Counter()}
    empties, seen = 0, False
    for el in body:
        if el.tag != qn("w:p"):
            continue
        if not _el_text(el):
            if seen:
                empties += 1
            continue
        if seen:
            info["gaps"][min(empties, 2)] += 1
        seen, empties = True, 0
        lines = _text_lines(el)
        text = "".join(lines)
        depth = _named_depth(_style_name(el, doc))
        # 标题样式常绑着自动编号（多级列表）——那是标题不是列表，不算列表段
        is_list = depth is None and _is_list(el, doc)
        if is_list:
            info["n_list"] += 1
            ni = _num_info(el, doc, num_root)
            if ni:
                info["list_infos"].append(ni)
        tag = None if is_list else _head_tag(lines[0])
        if tag and len(text) > 40:
            tag = None   # 带编号的长段是正文条款，不是标题
        fmt = None
        if depth is None and tag:
            fmt = _para_fmt(el, doc, ctx["theme"], ctx["defaults"])
            # 手打编号的短段里混着真标题和条款：标题不以句末标点收尾；不加粗又超过
            # 20 字的是条款（真实招标文件里「N.」开头的 75 个短段近半是条款）
            if text[-1] in "。；;！？" or (not fmt["bold"] and len(text) > 20):
                tag = None
        if depth is not None or tag:
            info["heads"].append({"lvl": depth, "named": depth is not None, "tag": tag,
                                  "text": lines[0], "el": el, "style": _style_name(el, doc),
                                  "fmt": fmt or _para_fmt(el, doc, ctx["theme"], ctx["defaults"])})
        elif len(text) >= 30 and not is_list:
            fmt = _para_fmt(el, doc, ctx["theme"], ctx["defaults"])
            info["body_fmts"].append(fmt)
            info["body_per_sec"][sec_of_el.get(el, 1)] += 1
    for f in info["body_fmts"] + [h["fmt"] for h in info["heads"]]:
        info["fonts"][f["ea"] or ctx["defaults"].get("ea") or "未设"] += 1
    return info


def _layout_body(ctx, rows, info) -> None:
    bf = info["body_fmts"]
    if not bf:
        for item in ("正文字体", "正文字号", "行距", "段前段后", "缩进", "对齐", "中文版式"):
            rows.add("正文", item, "未发现正文候选段（30 字以上的非列表段）", "正文区", False)
    else:
        rows.vote("正文", "正文字体", [(f["ea"], f["ascii"]) for f in bf],
                  lambda k: f"中文{k[0] or '（默认）'}、西文 {k[1] or '（同中文）'}", "正文候选")
        rows.vote("正文", "正文字号", [f["pt"] for f in bf], _pt_label, "正文候选")
        rows.vote("正文", "行距", [f["line"] for f in bf], lambda k: k, "正文候选")
        rows.vote("正文", "段前段后", [(f["before"], f["after"]) for f in bf], _ba_label, "正文候选")
        rows.vote("正文", "缩进", [f["ind"] for f in bf], lambda k: k, "正文候选")
        rows.vote("正文", "对齐", [f["al"] for f in bf], lambda k: k, "正文候选")
        rows.vote("正文", "中文版式", [f["cjk"] for f in bf], lambda k: k, "正文候选")
    n_list, lis = info["n_list"], info["list_infos"]
    if n_list:
        top = Counter(_num_label(f, t) for f, t in lis).most_common(1)
        val = f"自动编号 {n_list} 段" + (f"，以 {top[0][0]} 为主" if top else "")
    else:
        val = "未见自动编号段落"
    rows.add("正文", "列表习惯", val + "（列表样式随拷入的原件走，不落模板）", "正文区", True)
    gaps = info["gaps"]
    habit = "段间常空 1 行以上" if (gaps[1] + gaps[2]) > gaps[0] else "段间不空行"
    rows.add("正文", "空行习惯", habit + "（记录，不落模板）", f"正文区段间 {sum(gaps.values())} 处", True)


# ---------------------------------------------------------------- 目录

def _toc_guess_level(text: str) -> int:
    if _H_CH.match(text):
        return 1
    if _H_DEC2.match(text):
        return _dot_depth(text)
    if _H_GOV1.match(text):
        return 2
    if _H_GOV2.match(text):
        return 3
    return 1


def _toc_leader(el, doc, text: str) -> str:
    """条目与页码之间的引导线：制表位前导符（样式里的也算）、点线字符、或没有。"""
    pprs = [el.find(qn("w:pPr"))] + [st.element.find(qn("w:pPr")) for st in _style_chain(el, doc)]
    names = {"dot": "点线", "middleDot": "点线", "hyphen": "短线", "underscore": "下划线", "heavy": "粗线"}
    for ppr in pprs:
        tabs = ppr.find(qn("w:tabs")) if ppr is not None else None
        if tabs is None:
            continue
        for t in tabs.findall(qn("w:tab")):
            ld = t.get(qn("w:leader"))
            if ld and ld != "none":
                return "制表位引导线（" + names.get(ld, ld) + "）"
    if _TOC_DOTS.search(text):
        return "点线字符"
    return "制表位、无引导线" if _has_tab_char(el) else "无"


def _toc_info(toc_blocks, ctx) -> dict:
    doc = ctx["doc"]
    info = {"title": None, "entries": []}
    for el in toc_blocks:
        if el.tag != qn("w:p"):
            continue
        lines = _text_lines(el)
        text = "".join(lines)
        if not text:
            continue
        name = _style_name(el, doc).replace(" ", "").lower()
        if info["title"] is None and (_TOC_TITLE.match(text) or name in ("tocheading", "目录标题")):
            info["title"] = _para_fmt(el, doc, ctx["theme"], ctx["defaults"])
            continue
        info["entries"].append({"lvl": _toc_level(_style_name(el, doc)) or _toc_guess_level(lines[0]),
                                "text": lines[0], "fmt": _para_fmt(el, doc, ctx["theme"], ctx["defaults"]),
                                "lead": _lead_spaces(el), "leader": _toc_leader(el, doc, text)})
    return info


def _lead_spaces(p_el) -> str:
    """条目文字开头的空格（全角、半角、制表符）。_text_lines 会把行首行尾的空白去掉，
    缩进打在文字里的条目要从原始文字上取。"""
    raw = "".join(t.text or "" for t in p_el.iter(qn("w:t")))[:40]
    m = re.match(r"^[\u3000\u00a0 \t]+", raw)
    return m.group(0) if m else ""


def _layout_toc(rows, toc_info) -> None:
    rows.add("目录", "目录标题", _fmt_str(toc_info["title"]) if toc_info["title"] else "未发现目录标题",
             "「目录」标题段" if toc_info["title"] else "—", toc_info["title"] is not None)
    ents = toc_info["entries"]
    if not ents:
        rows.add("目录", "目录层级", "未发现目录条目", "—", False)
        return
    by: dict[int, list] = {}
    for e in ents:
        by.setdefault(e["lvl"], []).append(e)
    for lvl in sorted(by):
        name = f"目录{_DEPTH_CN[lvl - 1]}级条目" if lvl <= 4 else "目录五级及以上条目"
        rows.vote("目录", name, [_font_key(e["fmt"]) for e in by[lvl]], _fk_label, name)
    shown = sorted(by)[:4]
    rows.add("目录", "目录层级", f"{max(by)} 级",
             f"目录条目 {len(ents)} 条（" + "、".join(f"{_DEPTH_CN[l - 1]}级 {len(by[l])}" for l in shown) + "）", True)
    rows.vote("目录", "目录行距", [e["fmt"]["line"] for e in ents], lambda k: k, "目录条目")
    rows.vote("目录", "目录段前段后", [(e["fmt"]["before"], e["fmt"]["after"]) for e in ents], _ba_label, "目录条目")
    def _indent(l):
        top = Counter(e["fmt"]["left"] for e in by[l]).most_common(1)[0][0]
        s = f"{_DEPTH_CN[l - 1]}级 {top}" + ("" if top.endswith(("字", "米")) else " 厘米")
        # 缩进打在条目文字里的也要读出来（手工目录的常见写法：二级条目以全角空格起头）——只看段落
        # 缩进会读成「缩进 0 厘米」（真实回放：143 条二级条目全靠 2 个全角空格起头）
        lead, n = Counter(e["lead"] for e in by[l]).most_common(1)[0]
        if lead:
            unit = "全角空格" if "\u3000" in lead else ("制表符" if "\t" in lead else "空格")
            s += f"（{n}/{len(by[l])} 条文字以 {len(lead)} 个{unit}起头）"
        return s

    rows.add("目录", "目录缩进", "、".join(_indent(l) for l in shown), "各级条目左缩进与条目文字开头的空格", True)
    rows.vote("目录", "目录点线", [e["leader"] for e in ents], lambda k: k, "目录条目")


# ---------------------------------------------------------------- 标题

def _level_name(depth: int) -> str:
    return f"标题{_DEPTH_CN[depth - 1]}级" if depth <= 4 else "标题五级及以上"


def _short_level(name: str) -> str:
    return name[2:] if name.startswith("标题") else name


def _layout_headings(ctx, rows, info, blocks, sec_of_el, toc_info) -> None:
    heads = info["heads"]
    if not heads:
        for item in ("编号形态", "标题层级", "章起新页"):
            rows.add("标题", item, "未发现标题", "—", False)
        return
    sch = Counter(h["tag"] for h in heads if h["tag"])
    scheme = ("chapter" if sch.get("chapter")
              else ("gov" if sch.get("gov", 0) >= sch.get("decimal", 0) else "decimal")) if sch else "none"
    for h in heads:
        if h["lvl"] is None:
            h["lvl"] = _depth_of(h["tag"], h["text"], scheme) if h["tag"] else 2
    tagged = [h for h in heads if h["tag"]]
    named = [h for h in heads if h["named"]]
    if len(tagged) >= max(1, 0.3 * len(heads)):
        basis = "标题文字自带编号：" + "、".join(f"{_TAG_NAME[k]} {v}" for k, v in sch.most_common())
        sure = scheme == "chapter" or sch[scheme] / len(tagged) >= _LAYOUT_COEXIST
        rows.add("标题", "编号形态", _SCHEME_VAL[scheme], basis, sure)
    else:
        auto, txt = _heading_auto_scheme(ctx["doc"], {h["style"] for h in named})
        if auto:
            rows.add("标题", "编号形态", _SCHEME_VAL[auto], f"标题文字不带编号，标题样式绑定自动编号「{txt}」", True)
        elif toc_info["entries"]:
            shown = []
            for e in toc_info["entries"]:
                if e["lvl"] not in [x for x, _ in shown]:
                    shown.append((e["lvl"], e["text"][:12]))
            hint = _head_tag(toc_info["entries"][0]["text"]) or "decimal"
            rows.add("标题", "编号形态", _SCHEME_VAL[hint],
                     "标题文字不带编号，标题样式未绑定编号；目录条目显示 "
                     + " / ".join(t for _, t in shown[:3]), False)
        else:
            rows.add("标题", "编号形态", _SCHEME_VAL["none"], "标题文字不带编号，标题样式也未绑定编号", True)
    groups: dict[str, list] = {}
    for h in heads:
        groups.setdefault(_level_name(h["lvl"]), []).append(h["fmt"])
    for name in ("标题一级", "标题二级", "标题三级", "标题四级", "标题五级及以上"):
        fmts = groups.get(name)
        if not fmts:
            continue
        rows.vote("标题", name, [_font_key(f) for f in fmts], _fk_label, name)
        rows.vote("标题", name + "间距", [(f["line"], f["before"], f["after"]) for f in fmts], _sp_label, name)
    if scheme == "chapter":
        seps = []
        for h in tagged:
            if h["tag"] == "chapter":
                mm = re.match(r"^第\s*[^章部篇]+\s*[章部篇](\s*)", h["text"])
                if mm:
                    s = mm.group(1)
                    seps.append("全角空格" if "　" in s else ("半角空格" if s else "紧接"))
        if seps:
            rows.vote("标题", "编号间隔", seps, lambda k: k, "一级标题")
    elif scheme == "decimal":
        # 十进制编号（文字自带 1.1）：间隔同样读文字里编号后面的空白，按这一级的标题报
        seps, lvls = [], []
        for h in tagged:
            if h["tag"] != "decimal":
                continue
            mm = re.match(r"^\s*[1-9]\d*(?:\.\d+)+(\s*)", h["text"])
            if mm:
                s = mm.group(1)
                seps.append("全角空格" if "\u3000" in s else ("半角空格" if s else "紧接"))
                lvls.append(h["lvl"])
        if seps:
            rows.vote("标题", "编号间隔", seps, lambda k: k,
                      _short_level(_level_name(Counter(lvls).most_common(1)[0][0])) + "标题")
    idx_of = {el: i for i, el in enumerate(blocks)}

    def starts_page(h) -> bool:
        el = h["el"]
        if _breaks_before(el):
            return True
        i = idx_of.get(el)
        if not i:
            return False
        prev = blocks[i - 1]   # 分页符常落在上一段末尾（或上一段的空行里）
        return _breaks_after(prev) or sec_of_el.get(prev) != sec_of_el.get(el)

    # 「章」取这份文档里最浅的那一级标题。常规文档就是一级；通篇没有一级的（章用「1、」
    # 编号、或只用了标题 2 样式——真机上的资格文件、报价文件就是）没有可取的首元素，原先在
    # 这里 IndexError 整份读不出。退到最浅一级照实读，「没识别出一级」写进依据并记疑似。
    top = min(h["lvl"] for h in heads)
    ones = [h for h in heads if h["lvl"] == top]
    first_np = starts_page(ones[0])
    rest = ones[1:]
    k = sum(1 for h in rest if starts_page(h))
    if not rest or k == len(rest) or (rest and k / len(rest) >= _LAYOUT_COEXIST):
        val, sure = "是", True
    elif k == 0:
        val, sure = "否", True
    else:
        val, sure = f"部分（{k}/{len(rest)}）", False
    stat = f"首章前{'有' if first_np else '无'}分页，其余 {len(rest)} 处里 {k} 处起新页"
    if top == 1:
        basis = f"一级标题 {len(ones)} 处：{stat}"
    else:
        basis, sure = f"没识别出一级标题，按最浅的标题{_DEPTH_CN[max(top, 1) - 1]}级 {len(ones)} 处看：{stat}", False
    rows.add("标题", "章起新页", val, basis, sure)
    rows.add("标题", "标题层级", f"用到 {max(h['lvl'] for h in heads)} 级", f"标题候选 {len(heads)} 处", True)


# ---------------------------------------------------------------- 表格

def _table_style_defs(doc, style_id) -> dict:
    """表格样式链（basedOn 逐级）上第一个有定义的边框、首行底纹、首行加粗、单元格边距。"""
    out = {"borders": None, "shd": None, "bold": False, "mar": None}
    styles_el = doc.styles.element
    sid, guard = style_id, 0
    while sid and guard < 8:
        guard += 1
        st = styles_el.find(f"{qn('w:style')}[@{qn('w:styleId')}='{sid}']")
        if st is None:
            break
        tp = st.find(qn("w:tblPr"))
        if tp is not None:
            if out["borders"] is None and tp.find(qn("w:tblBorders")) is not None:
                out["borders"] = tp.find(qn("w:tblBorders"))
            if out["mar"] is None and tp.find(qn("w:tblCellMar")) is not None:
                out["mar"] = tp.find(qn("w:tblCellMar"))
        for cf in st.findall(qn("w:tblStylePr")):
            if cf.get(qn("w:type")) != "firstRow":
                continue
            tcpr = cf.find(qn("w:tcPr"))
            shd = tcpr.find(qn("w:shd")) if tcpr is not None else None
            fill = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
            if fill and fill not in ("AUTO", "FFFFFF"):
                out["shd"] = fill
            rpr = cf.find(qn("w:rPr"))
            b = rpr.find(qn("w:b")) if rpr is not None else None
            if b is not None and b.get(qn("w:val")) not in ("0", "false"):
                out["bold"] = True
        base = st.find(qn("w:basedOn"))
        sid = base.get(qn("w:val")) if base is not None else None
    return out


def _cell_mar(mar_el, side: str) -> int | None:
    el = mar_el.find(qn("w:" + side)) if mar_el is not None else None
    v = el.get(qn("w:w")) if el is not None else None
    return int(v) if v is not None and v.isdigit() else None


def _layout_tables(ctx, rows, body) -> None:
    """只统计正文区的表：封面信息栏、评分索引表这类招标规定的格式件不是公司的表样。"""
    from docx.table import Table
    doc = ctx["doc"]
    kinds, hstyles, dfmts, widths, mars = [], [], [], [], []
    fixed = cant_n = hdr_n = n = 0
    for el in body:
        if el.tag != qn("w:tbl"):
            continue
        n += 1
        tblpr = el.tblPr
        st_el = tblpr.find(qn("w:tblStyle")) if tblpr is not None else None
        sd = _table_style_defs(doc, st_el.get(qn("w:val")) if st_el is not None else None)
        sides = {}
        for side, _cn in _TBL_SIDES:
            d = tblpr.find(f"{qn('w:tblBorders')}/{qn('w:' + side)}") if tblpr is not None else None
            src = d if d is not None else (sd["borders"].find(qn("w:" + side)) if sd["borders"] is not None else None)
            if src is not None:
                sides[side] = (src.get(qn("w:val")) or "single", src.get(qn("w:sz")))
        vis = [s for s, _cn in _TBL_SIDES if sides.get(s, ("none",))[0] not in ("none", "nil")]
        if not sides or not vis:
            kind = "无框线"
        elif len(vis) == 6:
            kind = f"全线 {max(int(v[1] or 4) for v in sides.values()) / 8:g} 磅"
        elif set(vis) <= {"top", "bottom", "insideH"}:
            kind = "三线表" + ("（带内横线）" if "insideH" in vis else "")
        else:
            kind = "部分框线（" + "、".join(cn for s, cn in _TBL_SIDES if s in vis) + "）"
        trs = el.findall(qn("w:tr"))
        fill_t, hb, ha = None, [], []
        if trs:
            for tc in trs[0].findall(qn("w:tc")):
                shd = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
                fill = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
                if fill and fill not in ("AUTO", "NONE", "FFFFFF"):   # 纯白底纹就是没有底纹
                    fill_t = fill
                for p in tc.findall(qn("w:p")):
                    if _el_text(p):
                        f = _para_fmt(p, doc, ctx["theme"], ctx["defaults"])
                        hb.append(f["bold"])
                        ha.append(f["al"])
        look = tblpr.find(qn("w:tblLook")) if tblpr is not None else None
        first_on = look is not None and look.get(qn("w:firstRow")) not in (None, "0", "false")
        if fill_t is None and first_on:
            fill_t = sd["shd"]
        kinds.append(kind)
        hstyles.append((fill_t, (all(hb) if hb else False) or (first_on and sd["bold"]),
                        Counter(ha).most_common(1)[0][0] if ha else "左"))
        for tr in trs[1:4]:
            for tc in tr.findall(qn("w:tc"))[:8]:
                for p in tc.findall(qn("w:p")):
                    if _el_text(p):
                        dfmts.append(_para_fmt(p, doc, ctx["theme"], ctx["defaults"]))
        tw_el = tblpr.find(qn("w:tblW")) if tblpr is not None else None
        w_type = tw_el.get(qn("w:type")) if tw_el is not None else None
        w_val = tw_el.get(qn("w:w")) if tw_el is not None else None
        if w_type == "pct" and w_val and w_val.isdigit():
            widths.append(f"占版心 {int(w_val) / 50:g}%")
        elif w_type == "dxa" and w_val and w_val.isdigit():
            widths.append(f"固定 {_tw_cm(w_val)} 厘米")
        else:
            widths.append("自动")
        mar_el = tblpr.find(qn("w:tblCellMar")) if tblpr is not None else None
        if mar_el is None:
            mar_el = sd["mar"]
        side_v = _cell_mar(mar_el, "left")
        if side_v is None:
            side_v = 108   # Word 缺省 0.19 厘米
        mars.append(f"左右 {_tw_cm(side_v)} 厘米")
        lay = tblpr.find(qn("w:tblLayout")) if tblpr is not None else None
        fixed += lay is not None and lay.get(qn("w:type")) == "fixed"
        if trs and sum(1 for tr in trs if tr.find(f"{qn('w:trPr')}/{qn('w:cantSplit')}") is not None) * 2 >= len(trs):
            cant_n += 1
        hdr_n += bool(trs) and trs[0].find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None
    if not n:
        rows.add("表格", "表格边框", "未发现表格", "—", True)
        return
    rows.vote("表格", "表格边框", kinds, lambda k: k, f"表格 {n} 张")
    def _hs_label(k):
        s = f"底纹 {k[0]}" if k[0] else "无底纹"
        if k[1]:
            s += "、表头加粗"
        if k[2] != "左":
            s += f"、表头{k[2]}"
        return s

    rows.vote("表格", "表头样式", hstyles, _hs_label, "表头行（直接设置优先，其次表格样式）")
    if dfmts:
        rows.vote("表格", "表格文字", [(f["ea"], f["pt"], f["al"]) for f in dfmts],
                  lambda k: f"中文{k[0] or '（默认）'}｜{_pt_label(k[1])}｜{k[2]}", f"数据行单元格 {len(dfmts)} 格")
    rows.vote("表格", "表格宽度", widths, lambda k: k, f"表格 {n} 张")
    rows.vote("表格", "单元格边距", mars, lambda k: k, f"表格 {n} 张（缺省 0.19 厘米）")
    lay_val = "固定宽度" if fixed * 2 > n else "自动宽度"
    if cant_n * 2 > n:
        lay_val += "、行不跨页"
    if hdr_n * 2 > n:
        lay_val += "、表头跨页重复"
    rows.add("表格", "表格布局", lay_val, f"表格 {n} 张", True)


# ---------------------------------------------------------------- 表题、图题、图宽

def _layout_figures(ctx, rows, blocks, body, sec_of_el, page) -> dict:
    """表题、图题、图宽。只统计正文区的图与题注；题注位置看它挨着的是表还是图。"""
    doc = ctx["doc"]
    in_body = set(body)
    caps = {"表": [], "图": []}
    nums = {"表": Counter(), "图": Counter()}
    poss = {"表": Counter(), "图": Counter()}
    widths, centered, n_pic = [], 0, 0
    n_tbl = 0
    for i, el in enumerate(blocks):
        if el not in in_body:
            continue
        if el.tag == qn("w:tbl"):
            n_tbl += 1
            continue
        pics = _count_pictures(el)
        if pics:
            n_pic += pics
            for e in el.iter(qn("wp:extent")):
                if any(a.tag == _MC_FALLBACK for a in e.iterancestors()):
                    continue
                sec_no = sec_of_el.get(el, page["body_no"])
                text_cm = page["text_tw"][min(sec_no, len(page["text_tw"])) - 1] / 567
                widths.append((round(int(e.get("cx")) / 360000, 1), int(e.get("cx")) / 360000 / text_cm))
                centered += _para_fmt(el, doc, ctx["theme"], ctx["defaults"])["al"] == "居中"
        lines = _text_lines(el)
        if not lines:
            continue
        m = _CAP.match(lines[0])
        if not m:
            continue
        kind = m.group(1)
        nums[kind]["N-M" if m.group(2) else "N"] += 1
        prev, nxt = (blocks[i - 1] if i > 0 else None), (blocks[i + 1] if i + 1 < len(blocks) else None)
        if kind == "表":
            above = nxt is not None and nxt.tag == qn("w:tbl")
            below = prev is not None and prev.tag == qn("w:tbl")
        else:
            above = nxt is not None and nxt.tag == qn("w:p") and _count_pictures(nxt) > 0
            below = prev is not None and prev.tag == qn("w:p") and _count_pictures(prev) > 0
        if above != below:   # 位置是看题注挨着表还是图读出来的
            poss[kind]["表上方" if kind == "表" else "图上方"] += above
            poss[kind]["表下方" if kind == "表" else "图下方"] += below
            caps[kind].append(_para_fmt(el, doc, ctx["theme"], ctx["defaults"]))
    for kind, grp in (("表", "表格"), ("图", "图")):
        item = kind + "题"
        if caps[kind]:
            pos = poss[kind].most_common(1)[0][0]
            pat = nums[kind].most_common(1)[0][0]
            v = _vote([_font_key(f) for f in caps[kind]], _fk_label, kind + "题")
            rows.add(grp, item, f"{pos}、{v[0]}、编号「{kind} {pat}」", v[1], v[2])
        else:
            where = f"表格 {n_tbl} 张" if kind == "表" else f"插图 {n_pic} 张"
            rows.add(grp, item, f"未见{item}（记录，不落模板）", where, True)
    if widths:
        ratio = sum(r for _cm, r in widths) / len(widths)
        mc = Counter(cm for cm, _r in widths).most_common(1)[0][0]
        lead = "居中、" if centered * 2 >= len(widths) else ""
        rows.add("图", "图宽", f"{lead}多数 {mc:g} 厘米（约为版心宽的 {ratio * 100:.0f}%）",
                 f"插图 {len(widths)} 张（版心宽按所在节的页宽减边距）", True)
    else:
        rows.add("图", "图宽", "未发现插图", "—", True)
    return {"n_tbl": n_tbl, "n_pic": n_pic}


# ---------------------------------------------------------------- 封面

def _cover_flags(text: str) -> list[str]:
    flags = [k for k in _VAR_HINT if k in text]
    if _VAR_NUM.search(text):
        flags.append("数字串")
    return sorted(set(flags))


def _layout_cover(ctx, rows, cover, page) -> list:
    lines, n_tbl, n_pic, n_box = [], 0, 0, 0
    for el in cover:
        n_pic += _count_pictures(el)
        n_box += sum(1 for x in el.iter(qn("w:txbxContent"))
                     if not any(a.tag == _MC_FALLBACK for a in x.iterancestors()))
        if el.tag == qn("w:p"):
            lines += [(t, _cover_flags(t)) for t in _text_lines(el)]
        elif el.tag == qn("w:tbl"):
            n_tbl += 1
            for tr in el.findall(qn("w:tr")):
                cells = [" ".join(" ".join(_text_lines(p)) for p in tc.findall(qn("w:p"))).strip()
                         for tc in tr.findall(qn("w:tc"))]
                text = " ｜ ".join(c for c in cells if c)
                if text:
                    lines.append(("[表格] " + text, _cover_flags(text)))
    if not cover:
        rows.add("封面", "封面版式", "未识别出封面区（首页之后没有分页或分节）", "—", False)
    else:
        rows.add("封面", "封面版式",
                 f"整页拷贝：文字 {len(lines)} 行、表格 {n_tbl} 张、图 {n_pic} 张、文本框 {n_box} 个",
                 "首页（到第一个分页或分节符为止）", True)
    cats = sorted({c for _t, fl in lines for c in fl})
    rows.add("封面", "封面字段", "、".join(cats) if cats else "未见像会变的字", "封面区逐行标记（只给候选，哪些做成字段由用户定）", True)
    g0, gb = page["geoms"][0], page["geoms"][page["body_no"] - 1]
    sides = [f"{n} {_tw_cm(g0[k])}" for n, k in (("上", "top"), ("下", "bottom"), ("左", "left"), ("右", "right"))
             if g0[k] != gb[k]]
    rows.add("封面", "封面页设置", "同正文" if not sides else "与正文不同（" + "、".join(sides) + "）",
             "封面所在第 1 节 对比 正文所在节", True)
    return lines


# ---------------------------------------------------------------- 汇总与命令行

def layout_read_doc(doc) -> dict:
    """只读读取一份已打开文档的版式规格。返回 rows/facts/cover_lines/scope/warnings。"""
    # 样式数必须最先数：后面一碰段落样式，python-docx 就会给没有样式部件的文档就地补一份
    # 默认的，到最后再数就数不出「这个文件本来没有样式」了（真实文件回放抓到的）
    bare = _n_styles(doc) == 0
    ctx = {"doc": doc, "theme": _theme_faces(doc), "defaults": _doc_default_fonts(doc)}
    blocks = _layout_blocks(doc)
    secs = list(doc.sections)
    sec_of_el = dict(zip(blocks, _block_sections(blocks)))
    cover, front, toc_blocks, body = _split_regions(blocks, doc)
    rows = _Rows()
    info = _collect_body(ctx, body, sec_of_el)

    def step(grp, fn, *a, dep=_ANY):
        """一组读不出来不拖垮整份：记一行疑似、写明原因，后面的组照读。dep 是这一组依赖的
        上游结果，上游没读出来就不白跑；TENDER_TRACEBACK=1 时照旧抛出——调试和真实样本
        回放要的是完整栈，不能被这里吞掉。"""
        if dep is None:
            rows.add(grp, "本组未读", "它依赖的那一组没读出来，这一组没法读", _BROKEN, False)
            return None
        try:
            return fn(*a)
        except Exception as exc:  # noqa: BLE001
            if os.environ.get("TENDER_TRACEBACK"):
                raise
            rows.add(grp, "本组读取中断", f"{type(exc).__name__}: {str(exc)[:120]}", _BROKEN, False)
            return None

    page = step("页面", _layout_page, ctx, rows, secs, info)
    step("页面", _layout_header_footer, ctx, rows, secs, page, dep=page)
    step("正文", _layout_body, ctx, rows, info)
    toc_info = _toc_info(toc_blocks, ctx)
    step("目录", _layout_toc, rows, toc_info)
    step("标题", _layout_headings, ctx, rows, info, blocks, sec_of_el, toc_info)
    step("表格", _layout_tables, ctx, rows, body)
    fig = step("图", _layout_figures, ctx, rows, blocks, body, sec_of_el, page, dep=page) \
        or {"n_tbl": 0, "n_pic": 0}
    cover_lines = step("封面", _layout_cover, ctx, rows, cover, page, dep=page) or []

    heads = info["heads"]
    named_ct = sum(1 for h in heads if h["named"])
    if heads:
        if named_ct / len(heads) >= 0.8:
            note = "标题走命名样式"
        elif named_ct:
            note = "标题的命名样式与直接格式混用"
        else:
            note = _unstyled_head_note(doc, heads)
        fact1 = f"命名标题样式 {named_ct}/{len(heads)} 段（{note}）"
    else:
        fact1 = "未发现标题"
    facts = [fact1, "字体清单：" + ("、".join(k for k, _v in info["fonts"].most_common(8)) or "未读出")]
    dist = Counter(_short_level(_level_name(h["lvl"])) for h in heads)
    scope = [f"正文候选 {len(info['body_fmts'])} 段",
             (f"标题候选 {len(heads)} 段（" + "、".join(f"{k} {v}" for k, v in dist.most_common()) + "）")
             if heads else "标题候选 0 段",
             f"目录条目 {len(toc_info['entries'])} 条",
             f"表格 {fig['n_tbl']} 张（边框、表头、单元格、宽度、边距）",
             f"插图 {fig['n_pic']} 张（宽度、对齐）",
             (f"页面（正文所在第 {page['body_no']} 节的纸张、边距、网格；各节差异；页眉页脚）"
              if page else "页面（没读出来）"),
             f"封面区 {len(cover_lines)} 行、前置页 {len(front)} 块（不计入统计）",
             "未读：标题颜色、脚注尾注、水印与页面边框、分栏、列表缩进细节"]
    warnings = []
    if bare:
        # 任何 Office 另存出来的 docx 都带样式表。一个样式都没有的，多半是把文字抽出来重新拼成
        # 的 docx 冒充的转换件（真机上模型转一份 26MB 的 .doc 超了时，自己拼了一份：17.9 万段
        # 纯文本、8.8 万个分页符），原稿版式一点没跟过来，读到的全是缺省值——要点破，不能
        # 让它当版式样本交出去
        warnings.append("这份文件里一个样式都没有，不像 Office 另存出来的文档——多半是把文字抽出来"
                        "重新拼成的 docx，原稿的版式没有跟过来；下面读到的纸张、字体都是缺省值，"
                        "不能当版式样本。请在 Word/WPS 里打开原文件另存为 .docx 后再读。")
    return {"rows": rows.items, "facts": facts, "cover_lines": cover_lines, "scope": scope,
            "warnings": warnings}


def layout_read_file(path: Path) -> dict:
    """layout-read 的文件入口：.doc 先转 .docx，再读。只读，不落任何产物。
    转换这一步内置的只有 LibreOffice；没有它就把「怎么转」交给模型（LibreOffice 不是前提）。"""
    if path.suffix.lower() == ".doc":
        soffice = _find_soffice()
        if not soffice:
            die(_hand_over_hint(
                "版式读取只认 .docx，得先把它转成 .docx",
                "转好后对那份 .docx 重跑同一条命令即可，读取仍由脚本做；原 .doc 不用动。"
                "本机确实没有可用的转换方式时，再请用户在 Word/WPS 里另存为 .docx。"), 2)
        with _tempdir() as td:
            out = _lo_convert(soffice, path, Path(td), Path(td) / "lo-profile", target="docx",
                              stuck=lambda what: _stuck_doc_hint(
                                  what, "转好后对那份 .docx 重跑同一条命令即可，读取仍由脚本做；原 .doc 不用动。"))
            print(f"[.doc] 已转 docx 读取（LibreOffice）：{path.name}")
            return layout_read_doc(Document(str(out)))
    return layout_read_doc(Document(str(path)))


def _layout_group_lines(res: dict) -> tuple[list[str], int]:
    """规格表正文：每组一张四列表。取值不在半截处截断——「并列的另一种」被截掉就读不出所以然，
    落盘的规格表带着省略号后面也没法解析；一格里放不下的按段收（见 _md_val）。"""
    lines: list[str] = []
    cur = None
    sure = 0
    for grp, item, val, basis, lv in res["rows"]:
        if grp != cur:
            if lines:
                lines.append("")
            cur = grp
            lines += [f"【{grp}】", "| 项目 | 取值 | 依据 | 确定/疑似 |", "|---|---|---|---|"]
        sure += lv == "确定"
        lines.append(f"| {item} | {_md_val(val)} | {_md_cell(basis, 80)} | {lv} |")
    return lines, sure


def format_layout(res: dict, label: str) -> str:
    lines, sure = _layout_group_lines(res)
    n = len(res["rows"])
    # 首行是模型读回显的第一眼：有没能读取的组必须在这里点出来，不能只躺在表里
    head = f"[版式读取] {label}：读出 {n} 项（确定 {sure}、疑似 {n - sure}）"
    broken = sum(1 for r in res["rows"] if r[3] == _BROKEN)
    if broken:
        head += f"，其中 {broken} 项没能读取（表里标「读取中断／本组未读」，请原样反馈）"
    out = [head + "；本次只读，未改被读取的文件", ""]
    out += [f"[注意] {w}" for w in res.get("warnings", ())]
    if res.get("warnings"):
        out.append("")
    out += lines
    out += ["", "事实：" + "；".join(res["facts"])]
    if res["cover_lines"]:
        out += ["", "封面文字（→ = 像会变的字，确认哪些做成字段；· = 看着是固定文字）："]
        for t, flags in res["cover_lines"][:40]:
            tail = f"（{'、'.join(flags)}）" if flags else ""
            out.append(f"  {'→' if flags else '·'} {_md_cell(t, 120)}{tail}")
        more = len(res["cover_lines"]) - 40
        if more > 0:
            out.append(f"  （另有 {more} 行未列出）")
    out += ["", "已检查：" + "；".join(res["scope"])]
    # 末尾两行是给读回显的模型看的（真机上它只看了回显最后 22 行）：下一步与取值纪律写在这里
    out += ["[下一步] 把这张表交给用户确认（读得准的照报，需要看一眼的点名）；用户点头后跑 layout-adopt "
            "登记为公司版式，后面的标书自动按它出——不替用户手工搭模板，不把参数甩给用户。",
            "[取值] 读出什么报什么，由用户确认：不替用户改良版式，不推荐「更好看」的写法；"
            "用户提出要改，用 layout-adopt --set 记进规格表。"]
    return "\n".join(out)


def write_layout_spec(res: dict, path: Path, src: Path) -> None:
    lines = ["# 版式规格", "",
             f"> 来源：{src.name}（sha256 {_sha256(src)[:12]}…）；读取时间 {time.strftime('%Y-%m-%d %H:%M')}。",
             "> 本表是公司版式的本体：取值可直接改，「确定/疑似」列是机器读数的把握程度，"
             "「依据」列写着样本数。标注「只记不落」的项不进模板。"]
    lines += [f"> 注意：{w}" for w in res.get("warnings", ())]   # 这份文件当不了版式样本时，规格表开头就要说
    lines.append("")
    lines += _layout_group_lines(res)[0]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def cmd_layout_read(args) -> None:
    files = [Path(x) for x in args.files]
    for p in files:
        if not p.is_file():
            die(f"文件不存在：{p}")
        if p.suffix.lower() not in (".docx", ".doc"):
            die(f"只收 .docx / .doc——PDF 等读不准字体与磅值，请提供 Word 版：{p.name}", 2)
    if len(files) > 1 and args.out:
        die("--out 只支持单份文件")
    # 两个落盘参数都是「结果写到哪个文件」：读之前先过防呆（落到被读文件上、后缀不符即拒）
    hints = {"--out": 'layout-read "历史标书.docx" --out work/版式规格.md',
             "--json": 'layout-read "历史标书.docx" --json work/review/layout-read.json'}
    for flag, val, suffix in (("--out", getattr(args, "out", None), ".md"),
                              ("--json", getattr(args, "json", None), ".json")):
        if val:
            guard_dump_path(val, files, suffix, flag, hint=hints[flag])
    # 超过 5MB 的 .doc 不自动转换：读之前就拦下，多份文件时不留半截输出
    big = [p for p in files if p.suffix.lower() == ".doc" and _doc_too_big(p)]
    if big:
        die("\n".join(_oversize_doc_hint(
            os.stat(p).st_size,
            "转好后对那份 .docx 重跑同一条命令即可，读取仍由脚本做；原 .doc 不用动。") for p in big), 2)
    results = []
    for p in files:
        res = layout_read_file(p)
        print(format_layout(res, p.name))
        results.append({"file": p.name, **res})
        if getattr(args, "out", None):
            write_layout_spec(res, Path(args.out), p)
            print(f"[规格表已落盘] {args.out}（确认后的取值就是公司版式的本体）")
    if getattr(args, "json", None):
        jp = Path(args.json)   # 多份文件逐份保留，不以后一份覆盖前一份
        jp.parent.mkdir(parents=True, exist_ok=True)
        jp.write_text(json.dumps({"at": time.strftime("%Y-%m-%d %H:%M"), "files": results},
                                 ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[结果已落盘] {jp}")


# ---------------------------------------------------------------- 公司版式：登记与自动采用
#
# 规格表是公司版式的本体（layout-read 读出、用户确认、对话修订都落在它上面）。layout-adopt 把它
# 登记到用户目录；之后建节、合册不指定模板就自动按它出。登记目录只存规格表与解析结果，模板每次
# 从当前的内置基准模板现编译——引擎升级后登记的版式不会卡在旧模板上。
# 2026-11-27 WorkBuddy 真机：套件到读出规格表就停了，模型只能自己搭模板，还让用户「建节和合册时
# 传入这个参数」——用户只通过对话表达意图，参数不该甩给用户。

LAYOUT_HOME_ENV = "TENDER_LAYOUT_HOME"
LAYOUT_MD = "layout.md"
LAYOUT_JSON = "layout.json"
LAYOUT_COVER = "封面.docx"
_LAYOUT: dict | None = None     # 本进程生效的公司版式（登记记录）；None＝内置口径

_MAC_FONTS = ("Songti SC", "Heiti SC", "STSong", "STHeiti", "PingFang SC", "Hiragino Sans GB", "Kaiti SC",
              "STKaiti", "STFangsong", "Baoli SC", "Yuanti SC", "Lantinghei SC", "Libian SC", "Xingkai SC")
_SPEC_HEAD = ["项目", "取值", "依据", "确定/疑似"]

# 规格表里各个项目属于哪一组——对话修订补一行时要知道往哪张表里补
_ITEM_GROUP = {
    **dict.fromkeys(("纸张", "页边距", "页面设置差异", "对称页边距", "文档网格", "页眉", "页码写法", "页脚其他文字"), "页面"),
    **dict.fromkeys(("正文字体", "正文字号", "行距", "段前段后", "缩进", "对齐", "中文版式", "列表习惯", "空行习惯"), "正文"),
    **dict.fromkeys(("目录标题", "目录一级条目", "目录二级条目", "目录三级条目", "目录四级条目", "目录层级",
                     "目录行距", "目录段前段后", "目录缩进", "目录点线"), "目录"),
    **dict.fromkeys(("编号形态", "编号间隔", "章起新页", "标题层级", "标题一级", "标题一级间距", "标题二级",
                     "标题二级间距", "标题三级", "标题三级间距", "标题四级", "标题四级间距"), "标题"),
    **dict.fromkeys(("表格边框", "表头样式", "表格文字", "表格宽度", "单元格边距", "表格布局", "表题"), "表格"),
    "图题": "图", "图宽": "图",
}


def layout_home() -> Path:
    """版式登记目录：环境变量 TENDER_LAYOUT_HOME 优先（测试与多套登记用），缺省在用户主目录下
    （ZCode 与 WorkBuddy 两个宿主共用同一份登记——说一次就够）。"""
    return Path(os.environ.get(LAYOUT_HOME_ENV) or (Path.home() / ".lingyan" / "tender-layout"))


def _lay(*path, default=None):
    """登记的版式里某一项取值。没登记、显式给了模板、或 --ignore-layout 时是 default——
    引擎各处据此回落内置口径，行为与登记前逐项相同。"""
    v = (_LAYOUT or {}).get("values")
    for k in path:
        if not isinstance(v, dict) or k not in v:
            return default
        v = v[k]
    return default if v is None else v


# ---- 规格表的读与认


def _md_cells(line: str) -> list[str]:
    """md 表格行拆格（反斜杠转义的竖线不算分隔符）。"""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    return [c.replace("\x00", "|").strip() for c in s.replace("\\|", "\x00").split("|")]


def _spec_rows(text: str) -> list[tuple]:
    """规格表正文 → [(组, 项目, 取值, 依据, 确定/疑似, 行号)]。「读取时出错」的行不是读数，跳过。"""
    rows, grp, head = [], "", False
    for no, ln in enumerate(text.splitlines(), 1):
        if ln.startswith("【") and ln.endswith("】"):
            grp = ln[1:-1]
            continue
        if not ln.startswith("|"):
            continue
        cells = _md_cells(ln)
        if cells[:4] == _SPEC_HEAD:
            head = True
        elif cells and not all(set(c) <= set("-: ") for c in cells):
            if len(cells) >= 4 and grp:
                rows.append((grp, cells[0], cells[1], cells[2], cells[3], no))
    if not head or not rows:
        raise ValueError("这不是版式规格表（没有「项目｜取值｜依据｜确定/疑似」四列表）"
                         "——先用 layout-read 读出规格表，取值经用户确认后再登记")
    return [r for r in rows if r[4] != _BROKEN]


def _spec_name(text: str) -> str:
    m = re.search(r"^> 来源：(.+?)（sha256", text, re.M)
    return Path(m.group(1).strip()).stem if m and m.group(1).strip() else "公司标书"


def _p_font(s: str) -> tuple[str, str]:
    """「中文宋体、西文 Times New Roman」→ (中文, 西文)；「（默认）」「（同中文）」算没设。"""
    m = re.match(r"^中文(.+?)(?:、西文 (.+))?$", s.strip())
    if not m:
        raise ValueError(f"字体要写成「中文宋体、西文 Times New Roman」这种样子，「{s}」认不出")
    ea, asc = m.group(1).strip(), (m.group(2) or "").strip()
    return ("" if ea.startswith("（") else ea, "" if asc.startswith("（") else asc)


_PT_NAMES = {name: v for v, name in _PT_CN}


def _p_pt(s: str) -> float:
    """「小四（12 磅）」「12 磅」「三号」→ 磅值。"""
    m = re.search(r"([0-9]+(?:\.[0-9]+)?) 磅", s)
    if m:
        return float(m.group(1))
    if s.strip() in _PT_NAMES:
        return _PT_NAMES[s.strip()]
    raise ValueError(f"字号要带磅值（「小四（12 磅）」「12 磅」）或用字号名（「三号」），「{s}」认不出")


# 对齐的写法 → OOXML w:jc 的值（左＝不写）
_ALIGN_W = {"左": None, "两端": "both", "居中": "center", "右": "right", "分散": "distribute"}


def _p_fmt(s: str) -> dict:
    """「中文宋体、西文 Times New Roman｜小四（12 磅）｜加粗｜左」→ 字体、字号、加粗、对齐。各段都可省。"""
    f = {"ea": "", "ascii": "", "pt": None, "bold": False, "align": "左"}
    for part in [x.strip() for x in s.split("｜")]:
        if part.startswith("中文"):
            f["ea"], f["ascii"] = _p_font(part)
        elif "磅" in part or part in _PT_NAMES:
            f["pt"] = _p_pt(part)
        elif part in ("加粗", "不加粗"):
            f["bold"] = part == "加粗"
        elif part in _ALIGN_W:
            f["align"] = part
        elif part in ("倾斜", "不倾斜"):
            continue
        else:
            raise ValueError(f"「{part}」认不出——写成「中文宋体、西文 Times New Roman｜小四（12 磅）｜加粗｜左」，"
                             "字体、字号、加粗、对齐四段都可以省")
    return f


def _p_line(s: str) -> tuple[str, str]:
    """行距 → (w:line, w:lineRule)。"""
    s = s.strip()
    for pat, rule, k in ((r"^固定 ([0-9.]+) 磅$", "exact", 20), (r"^最小 ([0-9.]+) 磅$", "atLeast", 20),
                         (r"^([0-9.]+) 倍$", "auto", 240)):
        m = re.match(pat, s)
        if m:
            return str(int(round(float(m.group(1)) * k))), rule
    if s == "单倍":
        return "240", "auto"
    raise ValueError(f"行距写成「1.5 倍」「单倍」「固定 22 磅」「最小 22 磅」其中一种，「{s}」认不出")


def _p_side(s: str) -> dict:
    """段前或段后的一侧 → {auto}／{lines}／{tw}。"""
    s = s.strip()
    if s == "自动":
        return {"auto": 1}
    m = re.match(r"^([0-9.]+) 行$", s)
    if m:
        return {"lines": int(round(float(m.group(1)) * 100))}
    m = re.match(r"^([0-9.]+) 磅$", s)
    if m:
        return {"tw": int(round(float(m.group(1)) * 20))}
    raise ValueError(f"段前段后写成「段前 6 磅、段后 0.5 行」「段前自动」或「无」，「{s}」认不出")


def _p_ba(s: str) -> dict:
    """「段前 0 磅、段后 0 磅」「段后 10 磅」「无」→ 两侧段距；没写的那一侧是 None（不落）。"""
    out = {"before": None, "after": None}
    for part in [x.strip() for x in s.split("、")]:
        if part == "无":
            continue
        if part.startswith("段前"):
            out["before"] = _p_side(part[2:])
        elif part.startswith("段后"):
            out["after"] = _p_side(part[2:])
        else:
            raise ValueError(f"段前段后写成「段前 6 磅、段后 0.5 行」「段前自动」或「无」，「{part}」认不出")
    return out


def _p_spacing(s: str) -> dict:
    """「行距 1.5 倍、段前自动」「行距 单倍、无」→ 行距与两侧段距。"""
    out, rest = {"line": None}, []
    for part in [x.strip() for x in s.split("、")]:
        if part.startswith("行距 "):
            out["line"] = _p_line(part[3:])
        else:
            rest.append(part)
    out.update(_p_ba("、".join(rest)) if rest else {"before": None, "after": None})
    return out


def _p_ind(s: str) -> dict:
    s = s.strip()
    if s in ("无", "悬挂缩进"):
        return {"first": 0}      # 显式零：合册收尾只给没有 w:ind 的段补缩进
    m = re.match(r"^首行 ([0-9.]+) 字$", s)
    if m:
        return {"firstChars": int(round(float(m.group(1)) * 100))}
    m = re.match(r"^首行 ([0-9.]+) 厘米$", s)
    if m:
        return {"first": int(round(float(m.group(1)) * 567))}
    raise ValueError(f"缩进写成「首行 2 字」「首行 0.85 厘米」或「无」，「{s}」认不出")


def _p_paper(s: str) -> dict:
    s = s.strip()
    m = re.match(r"^(A4|A3|A5|Letter) (纵向|横向)$", s)
    if m:
        w, h = next((pw, ph) for pw, ph, n in ((11906, 16838, "A4"), (16838, 23811, "A3"),
                                               (8391, 11906, "A5"), (12240, 15840, "Letter")) if n == m.group(1))
        portrait = m.group(2) == "纵向"
    else:
        m = re.match(r"^([0-9.]+)×([0-9.]+) 厘米 (纵向|横向)$", s)
        if not m:
            raise ValueError(f"纸张写成「A4 纵向」「A4 横向」或「21×29.7 厘米 纵向」，「{s}」认不出")
        w, h, portrait = int(round(float(m.group(1)) * 567)), int(round(float(m.group(2)) * 567)), m.group(3) == "纵向"
    return {"w": w, "h": h} if portrait else {"w": h, "h": w}


def _p_margins(s: str) -> dict:
    out = {}
    for cn, key in (("上", "top"), ("下", "bottom"), ("左", "left"), ("右", "right")):
        m = re.search(cn + r"\s*([0-9.]+)", s)
        if not m:
            raise ValueError(f"页边距写成「上 2、下 2、左 2、右 2 厘米」，这里没读到「{cn}」")
        out[key] = int(round(float(m.group(1)) * 567))
    m = re.search(r"装订线 ([0-9.]+) 厘米", s)
    out["gutter"] = int(round(float(m.group(1)) * 567)) if m else 0
    return out


def _p_grid(s: str) -> dict:
    s = s.strip()
    lp = lambda mm: int(round(float(mm.group(1)) * 20)) if mm else 360   # noqa: E731
    if s.startswith(("无网格", "未设置")):
        return {"type": "default", "linePitch": lp(re.search(r"行距基准 ([0-9.]+) 磅", s))}
    if s.startswith("行网格"):
        m = re.search(r"每页约 (\d+) 行", s)
        return {"type": "lines", "lines": int(m.group(1)) if m else 0, "linePitch": 312}
    if s.startswith("行和字符网格"):
        m = re.search(r"每页约 (\d+) 行", s)
        return {"type": "linesAndChars", "lines": int(m.group(1)) if m else 0, "linePitch": 312}
    if s.startswith("字符网格对齐"):
        return {"type": "snapToChars", "linePitch": 312}
    raise ValueError(f"文档网格这一行是机器读的，写成「无网格（行距基准 18 磅）」「行网格（每页约 44 行）」其中一种，「{s}」认不出")


def _p_pagenum(s: str):
    """「第 X 页、页脚居中｜小四（12 磅）」→ 页码写法；「无页码」是 None。"""
    s = s.strip()
    if s.startswith(("无页码", "正文所在节没有页码")):
        return None
    parts = [x.strip() for x in s.split("｜")]
    kinds = [x.strip() for x in parts[0].split("、")]
    where = next((k for k in kinds if k.startswith(("页脚", "页眉"))), None)
    if not where or len(parts) < 2:
        raise ValueError(f"页码写成「纯数字、页脚居中｜五号（10.5 磅）」这种样子，「{s}」认不出")
    pos = where[2:]
    if pos not in ("居中", "左", "右"):
        raise ValueError(f"页码位置写成页脚居中／页脚左／页脚右，「{s}」认不出")
    return {"style": "x页" if "第 X 页" in kinds else ("dash" if "— X —" in kinds else "plain"),
            "total": "共 Y 页" in kinds, "where": where[:2], "pos": pos, "pt": _p_pt(parts[1])}


_SIDE_CN = {"上": "top", "下": "bottom", "左": "left", "右": "right", "内横": "insideH", "内竖": "insideV"}


def _p_border(s: str):
    s = s.strip()
    m = re.match(r"^全线 ([0-9.]+) 磅$", s)
    if m:
        return {"kind": "all", "sz": max(2, int(round(float(m.group(1)) * 8)))}
    if s == "无框线":
        return {"kind": "none", "sz": 4}
    if s.startswith("三线表"):
        return {"kind": "three", "sz": 12, "inner": "带内横线" in s}
    m = re.match(r"^部分框线（(.+)）$", s)
    if m:
        return {"kind": "partial", "sz": 4,
                "sides": [_SIDE_CN[x.strip()] for x in m.group(1).split("、") if x.strip() in _SIDE_CN]}
    if s.startswith("未发现表格"):
        return None
    raise ValueError(f"表格边框写成「全线 0.5 磅」「无框线」「三线表」，「{s}」认不出")


def _p_cellmar(s: str) -> int:
    m = re.search(r"左右 ([0-9.]+) 厘米", s)
    if not m:
        raise ValueError(f"单元格边距写成「左右 0.19 厘米」，「{s}」认不出")
    return int(round(float(m.group(1)) * 567))


def _p_header(s: str) -> dict:
    h = {"shade": None, "bold": False, "align": "左"}
    for part in [x.strip() for x in s.split("、")]:
        m = re.match(r"^底纹 ([0-9A-Fa-f]{6})$", part)
        if m:
            h["shade"] = m.group(1).upper()
        elif part == "表头加粗":
            h["bold"] = True
        elif part == "无底纹":
            continue
        elif part.startswith("表头") and part[2:] in _ALIGN_W:
            h["align"] = part[2:]
        else:
            raise ValueError(f"表头写成「无底纹」「底纹 D9D9D9、表头加粗、表头居中」这种样子，「{part}」认不出")
    return h


def _p_caption(s: str):
    """「表上方、中文宋体、西文 Times New Roman｜小四（12 磅）｜左、编号「表 N-M」」；「未见表题」是 None。"""
    s = s.strip()
    if s.startswith("未见"):
        return None
    m = re.match(r"^[表图](上|下)方、", s)
    body = s[m.end():] if m else s
    tail = re.search(r"、编号「[^」]*」$", body)      # 「、编号「表 N-M」」是给写题注的模型看的，不是格式
    if tail:
        body = body[:tail.start()]
    f = _p_fmt(body)
    f["above"] = (m.group(1) == "上") if m else None
    return f


def _p_toc_indent(s: str, pt: float) -> dict:
    """「一级 0 厘米、二级 2 字、三级 4 字」→ {"1": {"left": 缇}, …}。缩进打在条目文字里的全角空格
    （手工目录的常见写法）折成同量的段落缩进——生成的目录是域，不能靠打空格缩进。"""
    out = {}
    for part in [x.strip() for x in s.split("、")]:
        m = re.match(r"^([一二三四五])级 (.+)$", part)
        if not m:
            raise ValueError(f"目录缩进写成「一级 0 厘米、二级 2 字、三级 4 字」，「{part}」认不出")
        val, tw = m.group(2), 0.0
        lead = re.search(r"文字以 (\d+) 个(全角空格|空格)起头", val)
        if lead:
            tw += int(lead.group(1)) * pt * 20 * (1 if lead.group(2) == "全角空格" else 0.5)
            val = re.sub(r"（.*?）", "", val).strip()
        if val and not re.match(r"^[0-9.]+ (字|厘米)$", val):
            raise ValueError(f"目录缩进写成「一级 0 厘米、二级 2 字、三级 4 字」，「{val}」认不出")
        if val:
            num, unit = val.split(" ")
            tw += float(num) * (pt * 20 if unit == "字" else 567)
        out[str(_DEPTH_CN.index(m.group(1)) + 1)] = {"left": int(round(tw))}
    return out


def _p_leader(s: str) -> str:
    s = s.strip()
    if s.startswith("未发现"):
        return "dot"
    for word, kind in (("点线", "dot"), ("短线", "hyphen"), ("下划线", "underscore"), ("粗线", "heavy")):
        if word in s:
            return kind
    if "无引导线" in s or s == "无":
        return "none"
    raise ValueError(f"目录点线写成「点线字符」「制表位引导线（点线）」「制表位、无引导线」其中一种，「{s}」认不出")


_SEP_W = {"全角空格": "\u3000", "半角空格": " ", "紧接": ""}


def parse_layout_spec(text: str) -> dict:
    """规格表正文 → 版式记录的 values 与「需要用户看一眼」清单。取值有几种并存（疑似行）时取排在
    最前的那种，其余带回清单。认不出的取值抛 ValueError，消息点名行号、项目与写法。"""
    by: dict[str, tuple] = {}
    for grp, item, val, basis, lv, no in _spec_rows(text):
        by.setdefault(item, (val, lv, no))
    suspect, errors = [], []

    def first(item, required=True):
        if item not in by:
            if required:
                errors.append(f"规格表里没有「{item}」这一行")
            return None
        val, lv, no = by[item]
        alts = [v.strip() for v in val.split(" ／ ")]
        if lv == "疑似" and len(alts) > 1:
            suspect.append(f"{item}：{' ／ '.join(alts)}（按排在前面的出了）")
        return alts[0], no

    def one(item, fn, required=True):
        got = first(item, required)
        if got is None:
            return None
        if got[0].startswith(("未发现", "未见")) and item != "表格边框":
            return None     # 历史标书里没有的东西：这一项按内置写法
        try:
            return fn(got[0])
        except ValueError as exc:
            errors.append(f"第 {got[1]} 行「{item}」：{exc}")
            return None

    v: dict = {"page": {}, "body": {}, "headings": {}, "toc": {}, "numbering": {}, "table": {},
               "caption": {}, "pagenum": None, "image": None}
    pg = v["page"]
    pg.update(one("纸张", _p_paper, required=False) or {"w": 11906, "h": 16838})
    pg.update(one("页边距", _p_margins, required=False)
              or {"top": 1440, "bottom": 1440, "left": 1803, "right": 1803, "gutter": 0})
    pg["mirror"] = bool(by.get("对称页边距")) and by["对称页边距"][0].strip() == "是"
    hd_row = first("页眉", required=False)
    pg["header_row"] = (hd_row[0] if hd_row else "无").strip()
    grid = one("文档网格", _p_grid, required=False) or {"type": "default", "linePitch": 360}
    if grid.get("lines"):
        grid["linePitch"] = max(120, round((pg["h"] - pg["top"] - pg["bottom"]) / grid["lines"]))
    pg["grid"] = grid

    b = v["body"]
    b["ea"], b["ascii"] = one("正文字体", _p_font, required=False) or ("宋体", "Times New Roman")
    b["pt"] = one("正文字号", _p_pt, required=False) or 12
    b["line"] = one("行距", _p_line, required=False) or ["360", "auto"]
    b.update(one("段前段后", _p_ba, required=False) or {"before": None, "after": None})
    b["ind"] = one("缩进", _p_ind, required=False)
    b["align"] = first("对齐", required=False) or ["左"]
    b["align"] = b["align"][0]
    if b["align"] not in _ALIGN_W:
        errors.append(f"第 {by['对齐'][2]} 行「对齐」：写成左／两端／居中／右／分散其中一种，「{b['align']}」认不出")

    heads = v["headings"]
    for n in (1, 2, 3, 4):
        name = f"标题{_DEPTH_CN[n - 1]}级"
        f = one(name, _p_fmt, required=False)
        if f is not None:
            heads[str(n)] = {**f, **(one(name + "间距", _p_spacing, required=False)
                                     or {"line": None, "before": None, "after": None})}
    for n in ("3", "2"):     # 读缺的层级沿用上一级的外观
        if heads and str(n) not in heads and str(int(n) - 1) in heads:
            heads[n] = dict(heads[str(int(n) - 1)])
    if heads and "4" not in heads and "3" in heads:
        heads["4"] = dict(heads["3"])

    t = v["toc"]
    t["title"] = one("目录标题", _p_fmt, required=False)
    levels = {}
    for n in (1, 2, 3, 4):
        f = one(f"目录{_DEPTH_CN[n - 1]}级条目", _p_fmt, required=False)
        if f is not None:
            levels[str(n)] = f
    t["levels"] = levels
    t["line"] = one("目录行距", _p_line, required=False)
    t.update(one("目录段前段后", _p_ba, required=False) or {"before": None, "after": None})
    toc_pt = max((f["pt"] or 12 for f in levels.values()), default=12)
    t["indent"] = one("目录缩进", lambda s: _p_toc_indent(s, toc_pt), required=False) or {}
    m = re.match(r"(\d) 级", (first("目录层级", required=False) or ["3"])[0])
    t["depth"] = int(m.group(1)) if m else 3
    t["leader"] = one("目录点线", _p_leader, required=False)
    if not levels and t["title"] is None and not t["indent"] and t["leader"] is None:
        v["toc"] = t = {}      # 历史标书里没有目录：目录整组按内置写法

    sep = first("编号间隔", required=False)
    v["numbering"] = {"scheme": one("编号形态", lambda s: _p_scheme(s), required=False) or "chapter",
                      "sep": _SEP_W[sep[0]] if sep and sep[0] in _SEP_W else None}
    if sep and sep[0] not in _SEP_W:
        errors.append(f"第 {sep[1]} 行「编号间隔」：写成全角空格／半角空格／紧接，「{sep[0]}」认不出")
    cnp = first("章起新页", required=False)
    if cnp is None:
        v["chapter_new_page"] = True
    elif cnp[0] == "是" or cnp[0].startswith("未发现"):
        v["chapter_new_page"] = True
    elif cnp[0] == "否":
        v["chapter_new_page"] = False
    else:
        m = re.match(r"^部分（(\d+)/(\d+)）$", cnp[0])
        if not m:
            errors.append(f"第 {cnp[1]} 行「章起新页」：写成是／否，「{cnp[0]}」认不出")
        v["chapter_new_page"] = int(m.group(1)) * 2 > int(m.group(2)) if m else True

    tb = v["table"]
    tb["border"] = one("表格边框", _p_border, required=False)
    tb["header"] = one("表头样式", _p_header, required=False)
    tb["text"] = one("表格文字", _p_fmt, required=False)
    tb["mar"] = one("单元格边距", _p_cellmar, required=False)
    v["caption"] = {"table": one("表题", _p_caption, required=False), "figure": one("图题", _p_caption, required=False)}
    fields = first("封面字段", required=False)
    v["cover_fields"] = ([x.strip() for x in fields[0].split("、") if x.strip()] if fields else [])
    v["pagenum"] = one("页码写法", _p_pagenum, required=False)
    v["image"] = one("图宽", _p_image, required=False)
    if errors:
        raise ValueError("；".join(errors))
    return {"values": v, "suspect": suspect}


def _p_scheme(s: str) -> str:
    m = re.match(r"^(chapter|decimal|gov|none)", s.strip())
    if not m:
        raise ValueError(f"编号形态写成 chapter／decimal／gov／none 其中一种，「{s}」认不出")
    return m.group(1)


def _p_image(s: str):
    if s.strip().startswith("未发现"):
        return None
    m = re.search(r"版心宽的 ([0-9]+)%", s)
    return {"ratio": int(m.group(1)) / 100 if m else 1.0, "center": s.strip().startswith("居中")}


# ---- 登记目录的读写


def _sha_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _read_record() -> dict:
    try:
        return json.loads((layout_home() / LAYOUT_JSON).read_text(encoding="utf-8")) or {}
    except (OSError, ValueError):
        return {}


def _write_record(home: Path, rec: dict, text: str) -> None:
    home.mkdir(parents=True, exist_ok=True)
    md = home / LAYOUT_MD
    if md.is_file() and md.read_text(encoding="utf-8") != text:     # 上一份留档（最多 20 份）
        hist = home / "历史"
        hist.mkdir(exist_ok=True)
        (hist / f"layout-{time.strftime('%Y%m%d-%H%M%S')}.md").write_text(md.read_text(encoding="utf-8"), encoding="utf-8")
        for old in sorted(hist.glob("layout-*.md"))[:-20]:
            old.unlink()
    md.write_text(text, encoding="utf-8")
    (home / LAYOUT_JSON).write_text(json.dumps(rec, ensure_ascii=False, indent=2), encoding="utf-8")


def _layout_pending(rec: dict) -> list[str]:
    """这一版还没接上的项：登记时明说，模型据此转告用户（不许含糊成「都按这个样式」）。
    封面与页眉看登记时有没有提取到资产；页眉公司本来就没有（规格表页眉行=无）也不算没接上。"""
    assets = rec.get("assets") or {}
    out = []
    if not assets.get("cover"):
        out.append("封面整页（公司封面上的图与文字——重新登记时带 --source 即可接上）")
    hd_row = ((rec.get("values") or {}).get("page") or {}).get("header_row") or ""
    if not assets.get("header") and hd_row not in ("无", ""):
        out.append("页眉（logo 与文字——重新登记时带 --source 即可接上）")
    out.append("表格列宽（列宽仍由引擎按内容分配）")
    return out


def _header_images(p_el, part) -> list[dict]:
    """页眉段里的图片：二进制与显示尺寸（wp:extent）一起取——落地时按原尺寸回贴。"""
    out = []
    for blip in p_el.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        try:
            img = part.related_parts[rid]
        except KeyError:
            continue
        cx = cy = None
        for anc in blip.iterancestors():
            if anc.tag in (qn("wp:inline"), qn("wp:anchor")):
                ext = anc.find(qn("wp:extent"))
                if ext is not None:
                    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
                break
        out.append({"blob": img.blob, "ext": Path(str(img.partname)).suffix.lower() or ".png",
                    "cx": cx, "cy": cy})
    return out


def _header_pick(doc) -> dict | None:
    """挑节页眉里最有内容的一份（优先带 logo 的；都不带图取第一份有文字的）。"""
    theme, defaults = _theme_faces(doc), _doc_default_fonts(doc)
    best = None
    for sec in doc.sections:
        h = sec.header
        if h.is_linked_to_previous:
            continue
        lines, pts, eas, asiis, bolds, aligns, imgs = [], [], [], [], [], [], []
        for p in h.paragraphs:
            fmt = _para_fmt(p._p, doc, theme, defaults)
            ls = _text_lines(p._p)
            ni = _count_pictures(p._p)
            if ls or ni:
                lines += ls
                pts.append(fmt["pt"])
                eas.append(fmt["ea"])
                asiis.append(fmt["ascii"])
                bolds.append(fmt["bold"])
                aligns.append(fmt["al"])
            imgs += _header_images(p._p, h.part)
        if not lines and not imgs:
            continue
        cand = {"text": "　".join(lines)[:60], "ea": Counter(eas).most_common(1)[0][0] if eas else "宋体",
                "ascii": Counter(asiis).most_common(1)[0][0] if asiis else "Times New Roman",
                "pt": Counter(x for x in pts if x).most_common(1)[0][0] if pts else 9,
                "bold": any(bolds), "align": Counter(aligns).most_common(1)[0][0] if aligns else "居中",
                "border": any(_p_bottom_border(p._p) for p in h.paragraphs), "images": imgs}
        if imgs and (best is None or not best["images"]):
            best = cand          # 带 logo 的优先，取第一份
        elif best is None:
            best = cand          # 都不带图：取第一份有文字的
    return best


def _trim_cover_tail(clone, last: bool) -> bool:
    """封面末件清尾巴：自带的分节符与分页符都摘掉（整本里封面后面本来就另起页，带着会顶出
    空白页）。清完只剩空段时连段一起删（返回 True 让调用方跳过）。只在最后一个元素上生效。"""
    if not last:
        return False
    for s_el in clone.findall(".//" + qn("w:sectPr")):
        s_el.getparent().remove(s_el)
    for br in clone.findall(".//" + qn("w:br")):
        if br.get(qn("w:type")) == "page":
            br.getparent().remove(br)
    if clone.tag == qn("w:p") and not _el_text(clone) and not _count_pictures(clone):
        return True
    return False


def _build_cover_doc(src_doc: Document, cover_els: list) -> tuple[Document, int, list[str]]:
    """把封面区元素级拷进一份干净文档（图片双通道重挂、痕迹清理、用到的样式随迁；
    末件削掉分节符与分页符）。返回 (封面文档, 图片数, 跳过清单)。"""
    cd = Document(str(DEFAULT_TEMPLATE))
    _strip_body_content(cd)
    src_map = {s.style_id: s for s in src_doc.styles if getattr(s, "style_id", None)}
    tgt_ids = {s.style_id for s in cd.styles if getattr(s, "style_id", None)}
    body = cd.element.body
    sectpr = body.find(qn("w:sectPr"))
    img_n, skipped = 0, []
    for i, blk in enumerate(cover_els):
        el = blk   # _layout_blocks 给的是 lxml 元素本身（copy-from 那边才是带 ._p 的包装对象）
        clone, n, sk, _st = _clone_with_images(el, src_doc, cd)
        img_n += n
        skipped += sk
        for pstyle in clone.findall(".//" + qn("w:pStyle")):
            sid = pstyle.get(qn("w:val"))
            if sid and sid not in tgt_ids and sid in src_map:
                cd.styles.element.append(_deepcopy(src_map[sid].element))
                tgt_ids.add(sid)
        if _trim_cover_tail(clone, i == len(cover_els) - 1):
            continue
        if sectpr is not None:
            sectpr.addprevious(clone)
        else:
            body.append(clone)
    return cd, img_n, skipped


def _extract_layout_assets(src: Path, home: Path) -> dict:
    """从登记的历史标书里提取封面与页眉，存进登记目录（封面=封面.docx 元素级拷贝；
    页眉 logo 图=媒体/页眉-N.<ext>，带原显示尺寸）。返回 assets（空 dict＝没有可提取的）。"""
    home.mkdir(parents=True, exist_ok=True)  # 首次登记时登记目录还不存在（真机 2026-11-28：直接 FileNotFoundError）
    doc = Document(str(src))
    assets: dict = {}
    hd = _header_pick(doc)
    if hd:
        media = home / "媒体"
        media.mkdir(exist_ok=True)
        files = []
        for i, im in enumerate(hd["images"], 1):
            f = media / f"页眉-{i}{im['ext']}"
            f.write_bytes(im.pop("blob"))
            files.append({"file": f.relative_to(home).as_posix(), "cx": im["cx"], "cy": im["cy"]})
        hd["images"] = files
        assets["header"] = hd
    blocks = _layout_blocks(doc)
    cover_els, _f, _t, _b = _split_regions(blocks, doc)
    # 不过滤空段——空段就是封面的竖向排版，滤掉后封面内容全挤到页顶（真机 2026-11-28 抓到）；
    # 末件自带的分节符与分页符由 _trim_cover_tail 削
    if cover_els:
        cd, _n, _sk = _build_cover_doc(doc, cover_els)
        _save_doc(cd, home / LAYOUT_COVER)
        assets["cover"] = LAYOUT_COVER
    return assets


def load_layout() -> dict | None:
    """出文档前读本机登记的公司版式。规格表在登记后改过就按改后的表重读（对话修订不必非重跑
    登记）；改坏了就沿用上一次读得出来的取值并点名。没有登记返回 None。"""
    home = layout_home()
    md = home / LAYOUT_MD
    rec = _read_record()
    if not md.is_file():
        return rec if rec.get("values") else None
    text = md.read_text(encoding="utf-8")
    if not rec.get("values"):
        # 只有规格表、没有解析结果（手工拷过来的）：现读一份
        try:
            parsed = parse_layout_spec(text)
        except ValueError:
            return None
        rec = {"name": _spec_name(text), "registered_at": time.strftime("%Y-%m-%d %H:%M"),
               "values": parsed["values"], "suspect": parsed["suspect"], "md_sha256": _sha_text(text)}
        _write_record(home, rec, text)
        return rec
    if _sha_text(text) != rec.get("md_sha256"):
        try:
            parsed = parse_layout_spec(text)
        except ValueError as exc:
            print(f"  ⚠️ 登记的版式规格表在登记后改过，有一行读不出来（{exc}）——本次仍按上一次读得出来的"
                  f"取值出稿；修正后重跑 layout-adopt")
        else:
            rec.update(values=parsed["values"], suspect=parsed["suspect"], md_sha256=_sha_text(text))
            (home / LAYOUT_JSON).write_text(json.dumps(rec, ensure_ascii=False, indent=2), encoding="utf-8")
            print("[版式] 规格表在登记后改过，已按改后的表读取")
    return rec


def activate_layout(args) -> None:
    """出文档的命令（建节／合册／插图）开工前决定本次按不按登记的公司版式出。显式给了模板、或
    --ignore-layout（用户说这次不用公司版式），一律回到内置口径。"""
    global _LAYOUT
    _LAYOUT = None
    if args.cmd not in ("create", "assemble", "image"):
        return
    if getattr(args, "template", None) or getattr(args, "ignore_layout", False):
        return
    rec = load_layout()
    if not rec or not rec.get("values"):
        return
    _LAYOUT = rec
    if args.cmd in ("create", "assemble"):
        print(f"[版式] 本次按公司版式「{rec.get('name', '')}」出稿（{rec.get('registered_at', '')} 登记；"
              f"招标文件另有格式要求的以招标的为准）")


# ---- 把版式写进文档（样式与页面）


_RPR_ORDER = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof "
              "snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd "
              "fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath").split()
_PPR_ORDER = ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr "
              "shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN "
              "bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc "
              "textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange").split()


def _rebuild(style_el, tag: str, items: list) -> None:
    """把样式的 rPr／pPr 整体换成给定子元素（按 OOXML 定序）。留着旧属性会顶掉新设的——基准模板的
    标题样式带着主题字体与加粗，链接的字符样式还带着主题蓝，只改其中几项等于没改。"""
    for old in style_el.findall(qn(f"w:{tag}")):
        style_el.remove(old)
    if not items:
        return
    order = _RPR_ORDER if tag == "rPr" else _PPR_ORDER
    parent = OxmlElement(f"w:{tag}")
    for name, at in sorted(items, key=lambda it: order.index(it[0])):
        el = OxmlElement(f"w:{name}")
        for k, val in at.items():
            el.set(qn(k), str(val))
        parent.append(el)
    style_el.append(parent)     # 先 pPr 后 rPr 依次调用，落在 w:style 的末尾——正是定序里的位置


def _rpr_items(f: dict) -> list:
    at = {}
    if f.get("ascii"):
        at.update({"w:ascii": f["ascii"], "w:hAnsi": f["ascii"], "w:cs": f["ascii"]})
    if f.get("ea"):
        at["w:eastAsia"] = f["ea"]
    items = [("rFonts", at)] if at else []
    if f.get("bold"):
        items += [("b", {}), ("bCs", {})]
    if f.get("pt"):
        items += [("sz", {"w:val": int(round(f["pt"] * 2))}), ("szCs", {"w:val": int(round(f["pt"] * 2))})]
    return items


def _spacing_at(sp: dict) -> dict:
    at = {}
    if sp.get("line"):
        at["w:line"], at["w:lineRule"] = sp["line"][0], sp["line"][1]
    for side, pre in (("before", "w:before"), ("after", "w:after")):
        v = sp.get(side)
        if not v:
            continue
        if v.get("auto"):
            at[pre + "Autospacing"] = "1"
        elif "lines" in v:
            at[pre + "Lines"], at[pre] = v["lines"], "0"
        else:
            at[pre] = v.get("tw", 0)     # 不写「N 行＝0」：读数先看行数，写了会把「0 磅」读成「0 行」
    return at


def _jc_at(align: str) -> list:
    return [("jc", {"w:val": _ALIGN_W[align]})] if _ALIGN_W.get(align) else []


def _run_fmt(run, f: dict) -> None:
    """按版式里的字体字号加粗给一个 run 写直写格式（目录标题这类不挂样式的段用）。"""
    asc = f.get("ascii") or f.get("ea") or "Times New Roman"
    run.font.name = asc
    set_east_asia(run._element, f.get("ea") or asc)
    if f.get("pt"):
        run.font.size = Pt(f["pt"])
    run.font.bold = bool(f.get("bold"))


def _caption_style(style, f: dict, line) -> None:
    """题注样式：字体字号加粗对齐跟版式，行距随正文，首行不缩进。"""
    _rebuild(style.element, "pPr", [("ind", {"w:firstLine": "0"}),
                                    ("spacing", _spacing_at({"line": line})), *_jc_at(f.get("align", "左"))])
    _rebuild(style.element, "rPr", _rpr_items(f))


def apply_layout(doc: Document, rec: dict) -> None:
    """把登记的公司版式写进一份从内置基准模板起的文档：页面设置与版式管得着的样式。"""
    from docx.shared import Emu

    v = rec["values"]
    pg = v.get("page") or {}
    sec = doc.sections[0]
    if pg.get("w"):
        sec.page_width, sec.page_height = Emu(pg["w"] * 635), Emu(pg["h"] * 635)
    for key, attr in (("top", "top_margin"), ("bottom", "bottom_margin"), ("left", "left_margin"),
                      ("right", "right_margin"), ("gutter", "gutter")):
        if pg.get(key) is not None:
            setattr(sec, attr, Emu(int(pg[key]) * 635))
    grid = pg.get("grid") or {}
    dg = sec._sectPr.find(qn("w:docGrid"))
    if dg is None:
        dg = OxmlElement("w:docGrid")
        succ = next((s for s in sec._sectPr if s.tag in (qn("w:printerSettings"), qn("w:sectPrChange"))), None)
        succ.addprevious(dg) if succ is not None else sec._sectPr.append(dg)
    for a in list(dg.attrib):
        del dg.attrib[a]
    dg.set(qn("w:type"), grid.get("type", "default"))
    dg.set(qn("w:linePitch"), str(grid.get("linePitch", 360)))
    settings = doc.settings.element
    if pg.get("mirror") and settings.find(qn("w:mirrorMargins")) is None:
        mm = OxmlElement("w:mirrorMargins")
        mm.set(qn("w:val"), "true")
        settings.insert(0, mm)

    # 文档默认段距清零：基准模板的 docDefaults 里是段后 10 磅、行距 1.15 倍（Word 英文模板的常态），
    # 版式里「没设」的那一侧会被它顶成 10 磅。靠它撑着间距的封面两样式按原有效值写死，观感不变
    dd = doc.styles.element.find(qn("w:docDefaults"))
    dppr = dd.find(qn("w:pPrDefault")).find(qn("w:pPr")) if dd is not None else None
    if dppr is not None:
        dsp = dppr.find(qn("w:spacing"))
        if dsp is not None:
            dsp.set(qn("w:before"), "0")
            dsp.set(qn("w:after"), "0")
    for name in ("Cover Title", "Cover Sub"):
        st = _get_style(doc, name)
        if st is not None:
            _rebuild(st.element, "pPr", [("spacing", {"w:before": "0", "w:after": "200", "w:line": "276", "w:lineRule": "auto"}),
                                         ("ind", {"w:firstLine": "0"}), ("jc", {"w:val": "center"})])

    b = v.get("body") or {}
    normal = doc.styles["Normal"]
    _rebuild(normal.element, "pPr", [("spacing", _spacing_at(b)), *_jc_at(b.get("align", "左"))])
    _rebuild(normal.element, "rPr", _rpr_items(b))

    for n in (1, 2, 3, 4):
        h = (v.get("headings") or {}).get(str(n))
        st = _get_style(doc, f"Heading {n}")
        if h is None or st is None:
            continue
        spc = {"line": h.get("line"), "before": h.get("before"), "after": h.get("after")}
        _rebuild(st.element, "pPr", [("keepNext", {}), ("keepLines", {}), ("spacing", _spacing_at(spc)),
                                     *_jc_at(h.get("align", "左")), ("outlineLvl", {"w:val": n - 1})])
        _rebuild(st.element, "rPr", _rpr_items(h))
        ch = _get_style(doc, f"Heading {n} Char")
        if ch is not None:
            _rebuild(ch.element, "rPr", _rpr_items(h))

    t = v.get("toc") or {}
    levels = t.get("levels") or {}
    ind = t.get("indent") or {}

    def lvl_font(n: int) -> dict:
        have = sorted(levels, key=int)
        if not have:
            return {"ea": "宋体", "ascii": "宋体", "pt": 12, "bold": n == 1, "align": "左"}
        f = dict(levels.get(str(n)) or levels[min(have, key=lambda k: abs(int(k) - n))])
        f.setdefault("pt", 12)
        f.setdefault("bold", False)
        return f

    for n in (1, 2, 3, 4):
        st = _get_style(doc, f"TOC {n}")
        if st is None:
            continue
        f = lvl_font(n)
        if t:
            spc = {"line": t.get("line") or ["240", "auto"], "before": t.get("before"), "after": t.get("after")}
            if ind:
                have = sorted(ind, key=int)
                left = (ind.get(str(n)) or {}).get("left")
                if left is None:      # 表里没列的层级：往上的归零，往下的比最深一级多缩 2 字
                    left = 0 if n < int(have[0]) else ind[have[-1]]["left"] + round(f["pt"] * 40)
            else:
                left = round(0.74 * 567 * (n - 1))
        else:       # 历史标书里没有目录：目录整组按内置写法（把文档默认值原先给的有效值写死）
            spc, left = {"line": ["240", "auto"], "before": {"tw": 0}, "after": {"tw": 200}}, round(0.74 * 567 * (n - 1))
        _rebuild(st.element, "pPr", [("spacing", _spacing_at(spc)), ("ind", {"w:left": left, "w:firstLine": "0"})])
        _rebuild(st.element, "rPr", _rpr_items(f))

    # 题注：表题、图题各有各的样式（版式里两种题注常不一样：表题居左、图题居中）。版式没读到的那一种
    # 按内置写法——内置写法原先靠文档默认值撑着段后 10 磅，这里连同它一起写死，观感不变
    cap = v.get("caption") or {}
    builtin_cap = {"ea": "宋体", "ascii": "宋体", "pt": 9, "bold": True, "align": "居中", "above": None}
    for name, spec in (("Table Caption", cap.get("table")), ("Figure Caption", cap.get("figure"))):
        st = _get_style(doc, name)
        if st is None:
            st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if spec is None:
            _rebuild(st.element, "pPr", [("ind", {"w:firstLine": "0"}),
                                         ("spacing", {"w:before": "0", "w:after": "200", "w:line": "276", "w:lineRule": "auto"}),
                                         ("jc", {"w:val": "center"})])
            _rebuild(st.element, "rPr", _rpr_items(builtin_cap))
        else:
            _caption_style(st, spec, (v.get("body") or {}).get("line"))


# ---- 三个命令


def _fmt_desc(f: dict) -> str:
    s = f.get("ea") or "正文字体"
    if f.get("ascii"):
        s += f"／{f['ascii']}"
    if f.get("pt"):
        s += f"、{_pt_label(f['pt'])}"
    if f.get("bold"):
        s += "、加粗"
    return s


def _line_desc(line) -> str:
    return _line_label(*line) if line else "随正文"


def _ind_desc(ind) -> str:
    if not ind:
        return "首行缩进按内置 2 字"
    if ind.get("firstChars"):
        return f"首行缩进 {ind['firstChars'] / 100:g} 字"
    return f"首行缩进 {_tw_cm(ind['first'])} 厘米" if ind.get("first") else "首行不缩进"


def _border_desc(bd) -> str:
    return {"all": f"全线 {(bd or {}).get('sz', 4) / 8:g} 磅", "none": "无框线", "three": "三线表",
            "partial": "部分框线"}.get((bd or {}).get("kind"), "全线 0.5 磅")


def _layout_lines(rec: dict) -> tuple[list[str], list[str]]:
    """（生效清单，还没接上的清单）——登记、查看、模型转告用户用的是同一份话。"""
    v = rec["values"]
    pg, b = v.get("page") or {}, v.get("body") or {}
    on = [f"纸张与页边距（{_paper_label(pg).strip()}，上 {_tw_cm(pg.get('top', 1440))}、下 {_tw_cm(pg.get('bottom', 1440))}、"
          f"左 {_tw_cm(pg.get('left', 1803))}、右 {_tw_cm(pg.get('right', 1803))} 厘米）"]
    on.append(f"正文（{_fmt_desc({**b, 'bold': False})}、行距 {_line_desc(b.get('line'))}、{_ind_desc(b.get('ind'))}、"
              f"{b.get('align', '左')}对齐）")
    if v.get("headings"):
        on.append("一至四级标题（按登记表里标题各行）")
    t = v.get("toc") or {}
    if t:
        on.append(f"目录（{t.get('depth', 3)} 级，标题与各级条目的字体字号按登记表）")
    tb = v.get("table") or {}
    on.append(f"表格（{_border_desc(tb.get('border'))}，"
              f"{'表头加粗' if _lay_table_of(tb)['hb'] else '表头不加粗'}，"
              f"{'底纹 ' + _lay_table_of(tb)['shade'] if _lay_table_of(tb)['shade'] else '无底纹'}，"
              f"{_fmt_desc({**_lay_table_of(tb), 'bold': False})}）")
    cap = v.get("caption") or {}
    if cap.get("table"):
        on.append(f"表题（{_fmt_desc(cap['table'])}，表{'上' if cap['table'].get('above', True) else '下'}方）")
    if cap.get("figure"):
        on.append(f"图题（{_fmt_desc(cap['figure'])}）")
    pn = v.get("pagenum")
    on.append("页码（" + (f"{'第 X 页' if pn['style'] == 'x页' else ('— X —' if pn['style'] == 'dash' else '纯数字')}"
                 f"{'、共 Y 页' if pn.get('total') else ''}、{pn['where']}{pn['pos']}、{_pt_label(pn['pt'])}）"
                 if pn else "不编页码") + "；封面、前置页与目录页不编号，正文从第 1 页起")
    nm = v.get("numbering") or {}
    on.append(f"编号（{nm.get('scheme', 'chapter')}"
              f"{'、编号与标题之间' + {'　': '全角空格', ' ': '半角空格', '': '紧接'}[nm['sep']] if nm.get('sep') is not None else ''}"
              f"、{'章起新页' if v.get('chapter_new_page', True) else '章不另起页'}）")
    grid = (pg.get("grid") or {}).get("type", "default")
    if grid != "default":
        on.append(f"文档网格（{'行网格' if grid == 'lines' else ('行和字符网格' if grid == 'linesAndChars' else '字符网格对齐')}）")
    assets = rec.get("assets") or {}
    if assets.get("cover"):
        on.append(f"封面（登记的公司封面 {assets['cover']}；项目名、投标人、日期这些字段每本标书开写时再改）")
    hd_row = pg.get("header_row") or ""
    if assets.get("header"):
        hd = assets["header"]
        parts = []
        if hd.get("images"):
            parts.append("logo 图")
        if hd.get("text"):
            parts.append("文字「%s」" % hd["text"])
        if hd.get("pt"):
            parts.append(_pt_label(hd["pt"]))
        if hd.get("border"):
            parts.append("带横线")
        on.append("页眉（正文页显示：" + "、".join(parts) + "）")
    elif hd_row == "无":
        on.append("页眉（公司标书没有页眉，不出）")
    if (v.get("table") or {}).get("mar") not in (None, 108):
        on.append(f"表格单元格边距（左右 {_tw_cm(v['table']['mar'])} 厘米）")
    return on, _layout_pending(rec)


def _lay_table_of(tb: dict) -> dict:
    """表格口径（没读到的项回落内置）——与引擎出表时用的是同一份算法。"""
    text, head, bd = tb.get("text") or {}, tb.get("header") or {}, tb.get("border") or {}
    return {"ea": text.get("ea") or "宋体", "ascii": text.get("ascii") or "Times New Roman",
            "pt": text.get("pt") or 10.5, "align": text.get("align") or "左",
            "hb": head.get("bold", True) if tb.get("header") else True,
            "ha": (head.get("align") or "左") if tb.get("header") else "居中",
            "shade": head.get("shade") if tb.get("header") else "D9D9D9",
            "mar": tb.get("mar"),
            "border": bd if tb.get("border") and bd.get("kind") else {"kind": "all", "sz": 4}}


def _mac_font_note(rec: dict) -> str | None:
    v = rec["values"]
    names = set()
    for f in [v.get("body") or {}, *(v.get("headings") or {}).values(), *((v.get("toc") or {}).get("levels") or {}).values()]:
        if f.get("ea") in _MAC_FONTS:
            names.add(f["ea"])
    return (f"字体提醒：中文字体「{'、'.join(sorted(names))}」是 macOS 自带字体，出稿的机器是 Windows 时会回退成"
            f"别的字体——把登记表里的字体名改成那台机器上有的（如宋体），再登记一次") if names else None


def cmd_layout_adopt(args) -> None:
    src = Path(args.spec)
    if not src.is_file():
        die(f"规格表不存在：{src}（先用 layout-read 读出规格表，取值经用户确认后再登记）", 2)
    text = src.read_text(encoding="utf-8-sig")
    old = _read_record()
    try:
        if args.set:
            text = _apply_sets(text, args.set)
        parsed = parse_layout_spec(text)
    except ValueError as exc:
        die(str(exc), 2)
    home = layout_home()
    src_line = re.search(r"^> 来源：(.+)$", text, re.M)
    source = (src_line.group(1).strip() if src_line else "") or old.get("source", "")
    # 叫法：用户起了的优先；修订的是同一份表（来源行没变）就沿用上一次的叫法，别被来源文件名顶掉
    keep = bool(old.get("name")) and old.get("source") == source
    rec = {"name": args.name or (old.get("name") if keep else (_spec_name(text) if src_line else old.get("name")))
           or "公司标书",
           "source": source,
           "registered_at": time.strftime("%Y-%m-%d %H:%M"), "md_sha256": _sha_text(text),
           "values": parsed["values"], "suspect": parsed["suspect"]}
    # 封面与页眉要从历史标书里取件：带 --source 就提取存进登记目录；不带就沿用上一次的资产，
    # 从来没提取过则如实列进「还没接上」并给一句怎么接。取值先登记生效（登记目录在这里创建），
    # 资产提取失败只警告不回滚——登记的主意图是取值，别让它陪葬；话术不许谎称没发生的事
    rec["assets"] = old.get("assets") or {}
    _write_record(home, rec, text)
    if getattr(args, "source", None):
        spath = Path(args.source)
        if not spath.is_file():
            die(f"要提取封面与页眉的历史标书不存在：{spath}（版式取值已登记生效；"
                f"确认路径后重跑 layout-adopt <规格表> --source <历史标书>）", 2)
        try:
            rec["assets"] = _extract_layout_assets(spath, home)
        except Exception as exc:  # noqa: BLE001
            if os.environ.get("TENDER_TRACEBACK"):
                raise
            print(f"  ⚠️ 封面与页眉没能提取（{type(exc).__name__}: {str(exc)[:120]}）——"
                  f"版式取值已登记生效；要接上请重跑 layout-adopt <规格表> --source <历史标书>")
        else:
            _write_record(home, rec, text)   # 资产写回登记记录（同一份规格表，不进历史归档）
    on, off = _layout_lines(rec)
    print(f"[版式登记] 已登记为公司版式「{rec['name']}」，后面的标书自动按它出——建节、合册不用再指定模板；"
          f"招标文件另有格式要求的以招标的为准。")
    if args.set:
        print(f"  对话修订已记入：{'、'.join(s.split('=', 1)[0] for s in args.set)}")
    print("  生效：" + "；".join(on))
    if rec["suspect"]:
        print("  需要用户看一眼（历史标书里几种写法并存，先按排在前面的出了；用户要换就说）："
              + "；".join(s.replace("（按排在前面的出了）", "") for s in rec["suspect"]))
    if off:
        print("  这一版还没接上（标书里暂时按内置写法）：" + "、".join(off))
        if not getattr(args, "source", None) and not rec["assets"]:
            print("  要连封面与页眉一起接上：重新登记时把那份历史标书一起给——"
                  "layout-adopt <规格表> --source 历史标书.docx")
    note = _mac_font_note(rec)
    if note:
        print("  " + note)
    print(f"  存放：{home / LAYOUT_MD}（规格表就是版式的本体：要改版式，改这张表或用 --set，再登记一次；"
          f"用户说「别用公司版式了」跑 layout-forget）")
    print("  对用户怎么说：「已经记下了，以后新写的标书会自动按这个版式出；招标文件另有格式要求的，"
          "以招标的为准。」")


def _apply_sets(text: str, sets: list[str]) -> str:
    """对话修订：--set 项目=取值，把规格表里这一行的取值换掉（没有这一行就在所属组里补一行）。"""
    lines = text.splitlines()
    for kv in sets:
        item, sep, val = kv.partition("=")
        item, val = item.strip(), val.strip()
        if item not in _ITEM_GROUP:
            raise ValueError(f"没有「{item}」这一项——可改的项目：{'、'.join(_ITEM_GROUP)}")
        if not sep or not val:
            raise ValueError(f"--set 要写成「项目=取值」，「{kv}」里没有取值")
        new = f"| {item} | {val.replace('|', chr(92) + '|')} | 对话修订（用户 {time.strftime('%Y-%m-%d')} 说改的） | 确定 |"
        hit = next((i for i, ln in enumerate(lines) if ln.startswith("|") and _md_cells(ln)[:1] == [item]), None)
        if hit is not None:
            lines[hit] = new
            continue
        grp, head = _ITEM_GROUP[item], [f"【{_ITEM_GROUP[item]}】", "| 项目 | 取值 | 依据 | 确定/疑似 |", "|---|---|---|---|"]
        try:
            at = next(i for i, ln in enumerate(lines) if ln.strip() == head[0])
        except StopIteration:
            lines += ["", *head, new]
            continue
        end = at + 1
        while end < len(lines) and lines[end].startswith("|"):
            end += 1
        lines[end:end] = ([head[1], head[2]] if end == at + 1 else []) + [new]
    return "\n".join(lines) + "\n"


def cmd_layout_show(args) -> None:
    home = layout_home()
    rec = _read_record()
    if not (home / LAYOUT_MD).is_file() or not rec.get("values"):
        print("[版式] 没有登记公司版式——后面的标书按内置版式出。用户要登记：给一份 Word 版历史标书，"
              "layout-read 读出规格表、确认后 layout-adopt。")
        return
    on, off = _layout_lines(rec)
    print(f"[版式登记] 公司版式「{rec.get('name', '')}」（{rec.get('registered_at', '')} 登记"
          + (f"，来源 {rec['source']}" if rec.get("source") else "") + f"），存放：{home / LAYOUT_MD}")
    print((home / LAYOUT_MD).read_text(encoding="utf-8").rstrip())
    print("生效：" + "；".join(on))
    print("这一版还没接上：" + "、".join(off))


def cmd_layout_forget(args) -> None:
    home = layout_home()
    md = home / LAYOUT_MD
    if not md.is_file():
        print("[版式] 没有登记公司版式——后面的标书本来就按内置版式出，不用取消。")
        return
    hist = home / "历史"
    hist.mkdir(exist_ok=True)
    keep = hist / f"layout-{time.strftime('%Y%m%d-%H%M%S')}-取消登记.md"
    keep.write_text(md.read_text(encoding="utf-8"), encoding="utf-8")
    md.unlink()
    (home / LAYOUT_JSON).unlink(missing_ok=True)
    print(f"[版式取消] 已取消登记，后面的标书按内置版式出。原规格表留在 {keep}"
          f"（要恢复：对它重跑 layout-adopt）。")


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
    sp.add_argument("--template", help="版式模板 docx（缺省：登记了公司版式就按公司版式出，没登记用包内基准模板；"
                                       "显式给了模板就不按公司版式）")
    sp.add_argument("--ignore-layout", action="store_true",
                    help="本次不按登记的公司版式出，用内置版式（用户明说「这次不用公司版式」时才加）")
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
    sp.add_argument("--numbering", choices=["chapter", "decimal", "gov", "none"], default=None,
                    help="章节编号形态：chapter=第X章/1.1；decimal=1/1.1；gov=一、/（一）/1.；none=不编。"
                         "不给就跟登记的公司版式，没登记用 chapter（招标文件对编号有要求时明给）")
    sp.add_argument("--toc-depth", type=int, choices=[1, 2, 3, 4], default=None,
                    help="目录层级（静态条目与 TOC 域刷新口径一致）。不给就跟登记的公司版式，没登记用 3 级")
    sp.add_argument("--flat-outline", action="store_true",
                    help="节内小标题摘出导航大纲（outlineLvl=9）：导航窗格降噪，代价是"
                         "导航全扁平、F9 刷目录丢二级以下条目——默认关闭")
    sp.add_argument("--template", help="版式模板 docx（缺省：登记了公司版式就按公司版式出，没登记用包内基准模板；"
                                       "显式给了模板就不按公司版式，模板的页面设置与页眉页脚原样保留）")
    sp.add_argument("--ignore-layout", action="store_true",
                    help="本次不按登记的公司版式出，用内置版式（用户明说「这次不用公司版式」时才加）")
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

    sp = sub.add_parser("inspect", help="文档检查（只读）：金额验算、表格版式、图片、内部编号与占位、修订痕迹等，"
                                         "出文字结论 + 一张表；不改被检查的文件、不拦流程")
    sp.add_argument("docx", nargs="*", help="要检查的 docx，可多份（各处写的总价会放在一起对账）")
    sp.add_argument("--work", default=None, help="检查整套工作区的整本 Word 文档（含整本与节文件的对账）")
    sp.add_argument("--draft", action="store_true", help="按写作中的稿子查：不把还没处理的批注当问题")
    sp.add_argument("--json", default=None, metavar="结果.json",
                    help="把检查结果写成 JSON 文件——后面跟落盘路径（须 .json，如 work/review/inspect-result.json），"
                         "不是开关；不给就只打印、不写任何文件")
    sp.set_defaults(func=cmd_inspect)

    sp = sub.add_parser("layout-read", help="版式规格读取（只读）：把 Word 版历史标书读成版式规格表"
                                         "（项目｜取值｜依据｜确定/疑似）+封面文字清单；不改被读取的文件")
    sp.add_argument("files", nargs="+", help="历史标书 .docx/.doc（PDF 等读不准字体与磅值，不收）")
    sp.add_argument("--out", default=None, metavar="版式规格.md",
                    help="把版式规格表写成文件——后面跟落盘路径（须 .md）；缺省只打印。取值经用户确认后"
                         "交给 layout-adopt 登记为公司版式")
    sp.add_argument("--json", default=None, metavar="读取结果.json",
                    help="把读取结果写成 JSON 文件——后面跟落盘路径（须 .json，如 work/review/layout-read.json）；"
                         "不给就只打印")
    sp.set_defaults(func=cmd_layout_read)

    sp = sub.add_parser("layout-adopt", help="登记公司版式：把确认后的版式规格表登记为本机的默认版式，"
                                          "之后建节、合册不指定模板就自动按它出")
    sp.add_argument("spec", help="版式规格表 .md（layout-read --out 落盘的那份，或登记目录里的 layout.md）")
    sp.add_argument("--source", default=None,
                    help="历史标书 .docx——连封面整页与页眉（logo／文字／横线）一起提取存进登记目录；"
                         "与 --out 规格表读自同一份文件。不给则封面与页眉暂不接上")
    sp.add_argument("--name", default=None, help="版式的叫法（缺省取规格表来源行里的文件名）")
    sp.add_argument("--set", action="append", metavar="项目=取值", default=[],
                    help="对话修订：把某一项目的取值换成用户说的（如 '标题一级=中文黑体、西文 Times New Roman｜"
                         "三号（16 磅）｜加粗｜居中'），写法要跟规格表一致；可重复")
    sp.set_defaults(func=cmd_layout_adopt)

    sp = sub.add_parser("layout-show", help="查看登记的公司版式（用户问「现在用的什么版式」）")
    sp.set_defaults(func=cmd_layout_show)

    sp = sub.add_parser("layout-forget", help="取消公司版式登记（用户说「别用公司版式了」）；原规格表留档")
    sp.set_defaults(func=cmd_layout_forget)

    args = ap.parse_args()
    require_docx()  # --help 已在 parse_args 内退出，这里才真正需要 python-docx
    activate_layout(args)  # 出文档的命令先定本次按不按登记的公司版式出
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
