#!/usr/bin/env python3
"""check_writing.py — 正文机器校验（提示不是门禁，收尾前必须全绿）。

五种模式：
    python3 check_writing.py --guide                 校验写作指引（格式契约）
    python3 check_writing.py --analysis              校验分析登记表（coverage + 出处引用三层）
    python3 check_writing.py --outline               目录来源标注对账（悬空/未引用来源ID）
    python3 check_writing.py --section <节.docx>     单节自查（占位/残字/素材使用率/残留）
    python3 check_writing.py [--old-names "a,b"]     全局：承诺比对 + 节间查重 + 占位清点 + 残留

本模块同时是共用纯函数的落点（build_review / apply_review / build_dispatch 从这里
import，依赖方向恒为 上游→本模块，禁止反向）：table_rows / _read / _resolve_source /
_split_tables / _registry_payload / _parse_annotations / ID_RE / LNUM_RE。

依赖：python-docx。
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx_toolkit import build_tree, iter_blocks, parse_outline, sanitize_name  # noqa: E402

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

PLACEHOLDER = re.compile(r"【\s*(待补|待澄清|待确认|待补充|待定|待核验)[^】]*】|TODO|FIXME")
MD_RESIDUE = re.compile(r"^(#{1,6}\s|\||\*\*|```)")
WEAK_MD = re.compile(r"^-\s")
MATERIAL_TOKEN = re.compile(r"^(.+\.docx)\s+L(\d+)(?:-L?(\d+))?$")
# 指引缺口/备注列的待办标记——列里声明了缺口或批注待办，节 docx 就必须带
# Word 批注（散文参数化回避不替代批注：待拍板事项要以批注台账为载体点给用户）
GAP_MARK = re.compile(r"【缺|批注")
# 小标题自带编号形态（一、/（一）/1. 式）——合册按树深度接续节号，会双重编号
HEADING_NUMBERING = re.compile(r"^[一二三四五六七八九十]+、|^[（(][一二三四五六七八九十]+[)）]|^\d+[.、]")
LEGAL_MODES = {"素材修订", "格式跟随", "推理撰写", "—"}


def _base_name(tok_path: str) -> str:
    """素材 token 的文件名：剥目录前缀，/ 与 \\ 都认（Windows 环境的 agent 可能
    写反斜杠路径——Path().name 在 posix 下不认 \\，归一失效会重新绕开核对与查重）。"""
    return re.split(r"[/\\]", tok_path)[-1]
LEGAL_FIGURES = {"表", "分层", "甘特"}
# 语言纪律七条的可字面匹配子集（提示不是门禁，warning 级）
AI_FLAVOR = re.compile(
    r"标志着|至关重要|彰显了?|奠定[了]?坚实基础|保驾护航|行业专家认为|"
    r"大量实践证明|业界普遍(?:认为|认同)|有信心、有能力")
PROMISE_SNAPSHOT = "关键事实与承诺.snapshot.json"

# ---- 登记表与目录标注共用常量（原 build_review / apply_review 定义，1.8.0 下沉）----
ID_RE = re.compile(r"^(MAND|TPL|REQ|SCORE)-\d{1,3}$", re.I)
LNUM_RE = re.compile(r"L(\d+)")
ANALYSIS_FILES = ("structure.md", "requirements-format.md", "requirements-business.md",
                  "evaluation.md", "disqualification.md")
# 出处单元格里的独立文件名形态（「招标文件.docx L…」「xx.docx（L…）」）
_FILENAME_LIKE = re.compile(r"^(?P<fn>[^\s（();；/]+?\.(?:docx|pdf|txt|md))(?=[\s（(]|$)")
# coverage 段头（兼容「## coverage」「## coverage 声明」「四、coverage 声明」各形态）
_COVERAGE_HEAD = re.compile(
    r"^#{2,6}\s*(?:[一二三四五六七八九十]+[、.．]?\s*)?coverage\s*(?:声明)?\s*[:：]?\s*$", re.I)
# 缺口/备注列读者分离（warning 级）：该列会被摘出独立展示给用户——实现名词与
# 代词回指摘出后读不通（真机用户反馈「真看不懂」换来的两条）
NOTE_TOOL_RE = re.compile(r"parse_document|docx_toolkit|check_writing|build_review"
                          r"|serve_review|apply_review|build_dispatch|python3|--\w+")
NOTE_ANAPHORA_RE = re.compile(r"上述|以上|前述|上表|前面|前文")


def die(msg: str, code: int = 1):
    print(f"[校验失败] {msg}")
    sys.exit(code)


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def digit_seq_hit(value: str, texts: list[str]) -> bool:
    """承诺比对宽松复核：数字序列按原顺序出现在任一正文里，且值中紧随数字的
    量词在正文同一处近旁出现（「30分钟」不放过正文「30天」——数字同、量词异
    是真差异；量词含 %，「100%」同样不放过正文「100本」）。单数字值（1名/5人）
    信号太弱不做宽松放行，纯文字值不算命中——真漏写不能被洗成黄色提醒
    （真机教训：8 条写法差异告警里混着真实漏写）。"""
    runs = re.findall(r"\d+(?:\.\d+)?", value)
    if not runs or sum(len(d) for d in runs) < 2:
        return False
    pairs = [(m.group(1), m.group(2)) for m in
             re.finditer(r"(\d+(?:\.\d+)?)([\u4e00-\u9fff%]{1,2})", value)]
    for t in texts:
        tseq = re.findall(r"\d+(?:\.\d+)?", t)
        it = iter(tseq)
        if not all(any(d == x for x in it) for d in runs):
            continue
        if all(re.search(re.escape(d) + r"[^\d]{0,4}" + re.escape(u), t)
               for d, u in pairs):
            return True
    return False


def shingles(text: str, k: int = 10) -> set[str]:
    t = norm(text)
    return {t[i:i + k] for i in range(0, max(0, len(t) - k + 1))} if len(t) >= k else ({t} if t else set())


def docx_text(path: Path) -> tuple[str, list[tuple[str, str]]]:
    """全文 + 定位片段列表（P#/T#R#C → 文本）。"""
    doc = Document(str(path))
    labeled: list[tuple[str, str]] = []
    pi = ti = 0
    for block in iter_blocks(doc):
        if hasattr(block, "rows"):
            ti += 1
            seen: list = []
            for ri, row in enumerate(block.rows, 1):
                for ci, cell in enumerate(row.cells, 1):
                    if any(cell._tc is t for t in seen):
                        continue
                    seen.append(cell._tc)
                    labeled.append((f"T{ti}R{ri}C{ci}", cell.text))
            continue
        pi += 1
        labeled.append((f"P{pi}", block.text))
    return "\n".join(t for _, t in labeled), labeled


def table_rows(md_text: str) -> list[list[str]]:
    rows = []
    for line in md_text.splitlines():
        s = line.strip()
        if not (s.startswith("|") and s.endswith("|")):
            continue
        cells = [c.strip() for c in s[1:-1].split("|")]
        if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
            continue
        rows.append(cells)
    return rows


def scan_ai_flavor(name: str, labeled: list[tuple[str, str]]) -> list[str]:
    """AI 腔词表扫描（warning 级，每份文件最多列 3 处，指向语言纪律七条）。"""
    hits: list[str] = []
    for loc, frag in labeled:
        for m in AI_FLAVOR.finditer(frag):
            ctx = frag[max(0, m.start() - 12):m.end() + 12].replace("\n", " ").strip()
            hits.append(f"{name}:{loc} 疑似 AI 腔「{m.group(0)}」：…{ctx}…（改写见 section-writing 语言纪律）")
            if len(hits) >= 3:
                return hits
    return hits


# ---------------------------------------------------------------- 共用纯函数（下沉）

def _read(path: Path) -> str:
    """读文件文本；不存在返回空串（登记表缺失按各校验自己点名）。"""
    return path.read_text(encoding="utf-8") if path.is_file() else ""


def _resolve_source(cell: str, last_file: str) -> tuple[str, str, int | None]:
    """出处线索 → (展示串, 文件名, 首行号)。「同文件」回填最近出现的文件名。"""
    cell = (cell or "").strip()
    m = re.match(r"^(\S+\.(?:docx|pdf|txt|md))\s+(.*)$", cell)
    if m:
        fname, rest = m.group(1), m.group(2)
    elif cell.startswith("同文件") and last_file:
        fname, rest = last_file, cell[len("同文件"):].strip()
    else:
        fname, rest = (last_file or ""), cell
    lm = LNUM_RE.search(rest)
    return f"{fname} {rest}".strip(), fname, (int(lm.group(1)) if lm else None)


def _split_tables(rows: list[list[str]], header_kw: list[str]) -> list[dict]:
    """按表头关键词拆分同文件的多张登记表；其它「序号」表头行会关闭当前表
    （requirements-format.md 里 MAND 表与 TPL 表相邻，不关表会把对方的行吞进来）。"""
    tables, cur = [], None
    for r in rows:
        joined = "".join(r)
        is_sep = all(re.fullmatch(r":?-{2,}:?", c.strip()) for c in r if c.strip())
        has_src = any("出处" in c or "线索" in c for c in r)
        if any(k in c for c in r for k in header_kw) and has_src:
            cur = {"header": [c.strip() for c in r], "rows": []}
            tables.append(cur)
        elif is_sep:
            continue
        elif cur is not None and has_src and (
                "序号" in joined or len(r) != len(cur["header"])):
            # 另一张登记表的表头（「序号」式，或如「▲条款|内容要点|出处」无序号
            # 式——列数与当前表不符即关表，防其表头与数据被吞进当前表错列校验）
            cur = None
        elif cur is not None:
            cur["rows"].append([c.strip() for c in r])
    return tables


def _registry_payload(work: Path) -> tuple[dict, list[dict]]:
    """登记表：编号=清单行序（outline-annotation 约定）。"""
    registry: dict[str, dict] = {}
    scores: list[dict] = []
    specs = [  # (文件, 表头须含, ID 前缀, 内容列名, 出处列名候选, 分值列名)
        ("analysis/requirements-format.md", ["章节名称"], "MAND", "章节名称", ["原文出处线索"], None),
        ("analysis/requirements-format.md", ["模板名称"], "TPL", "模板名称", ["原文出处线索"], None),
        ("analysis/requirements-business.md", ["需求"], "REQ", "需求", ["出处"], None),
        ("analysis/evaluation.md", ["评分项"], "SCORE", "评分项", ["出处"], "分值"),
    ]
    for relpath, header_kw, prefix, text_col, src_cols, score_col in specs:
        path = work / relpath
        if not path.is_file():
            continue
        tables = _split_tables(table_rows(_read(path)), header_kw)
        for t in tables:
            header = t["header"]
            last_file = ""
            for n, r in enumerate(t["rows"], 1):
                r = (r + [""] * len(header))[:len(header)]
                cid = f"{prefix}-{n:02d}"
                ti = next((i for i, c in enumerate(header) if text_col in c), 0)
                si = next((i for i, c in enumerate(header) for sc in src_cols if sc in c), len(header) - 1)
                text = r[ti] if ti < len(r) else ""
                source_raw = r[si] if si < len(r) else ""
                shown, fname, line = _resolve_source(source_raw, last_file)
                if fname:
                    last_file = fname
                entry = {"type": prefix, "text": text, "source": shown,
                         "file": fname or None, "line": line}
                if score_col:
                    sci = next((i for i, c in enumerate(header) if score_col in c), None)
                    if sci is not None and sci < len(r):
                        entry["score"] = r[sci]
                    scores.append({"id": cid, "item": text,
                                   "score": entry.get("score", ""),
                                   "points": next((r[i] for i, c in enumerate(header)
                                                   if "评分要点" in c and i < len(r)), "")})
                registry[cid] = entry
    return registry, scores


def _parse_annotations(lines: list[str]) -> dict[int, dict]:
    """每卷的 来源标注/目录说明 `名称 :: 内容` 映射（按文件行序）。"""
    result: dict[int, dict] = {}
    vi, sec = -1, None
    for ln in lines:
        if ln.startswith("# 响应文件："):
            vi += 1
            sec = None
            result[vi] = {"name": ln[len("# 响应文件："):].strip(), "src:": {}, "desc:": {}}
        elif ln.startswith("## ") and vi >= 0:
            sec = ln[3:].strip()
        elif vi >= 0 and sec in ("来源标注", "目录说明") and ln.lstrip().startswith("- "):
            body = ln.strip()[2:]
            if " :: " in body:
                name, rest = body.split(" :: ", 1)
                result[vi]["src:" if sec == "来源标注" else "desc:"].setdefault(
                    name.strip(), rest.strip())
    return result


def _norm_id(s: str) -> str:
    """来源 ID 归一：MAND-1 / mand-01 → MAND-01（对账两侧同规）。"""
    m = re.match(r"^(MAND|TPL|REQ|SCORE)-0*(\d{1,3})$", s.strip(), re.I)
    return f"{m.group(1).upper()}-{int(m.group(2)):02d}" if m else s.strip().upper()


# ---------------------------------------------------------------- --analysis

# 登记表头部声明行（本包约定：> 来源文件：`xxx.docx`…。下表「同文件」即指该文件。
# 尾巴兼容「`（注释…）」「`。下表…」与裸行尾三种形态）
_HDR_SOURCE_RE = re.compile(r"^>\s*来源文件[：:]\s*`?([^`（(\n]+?)`?\s*(?:[（(。；]|$)")


def _declared_source(text: str) -> str:
    """登记表头部的来源文件声明名（不看集合——是否合法由调用方对账）。"""
    for ln in text.splitlines()[:12]:
        m = _HDR_SOURCE_RE.match(ln.strip())
        if m:
            return m.group(1).strip()
    return ""


def _split_top(cell: str, seps: str) -> list[str]:
    """只在括号深度 0 处按 seps 分割——括号内的「；」是行文注释不是多出处分隔。"""
    out: list[str] = []
    buf: list[str] = []
    depth = 0
    for ch in cell:
        if ch in "（(":
            depth += 1
        elif ch in "）)":
            depth = max(0, depth - 1)
        if ch in seps and depth == 0:
            out.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    out.append("".join(buf))
    return [s.strip() for s in out if s.strip()]


def _known_sources(work: Path) -> tuple[list[str], dict[str, int], str]:
    """sources.json → (确认来源文件名集合, 各来源解析 md 总行数, 降级提示)。"""
    path = work / "parse" / "sources.json"
    known: list[str] = []
    src_lines: dict[str, int] = {}
    note = ""
    if not path.is_file():
        return known, src_lines, "sources.json 不存在——文件名集合与行号上限校验跳过"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return known, src_lines, "sources.json 无法解析——文件名集合与行号上限校验跳过"

    def fname_of(entry) -> str:
        if isinstance(entry, str):
            return entry.strip()
        if isinstance(entry, dict):
            v = entry.get("path") or entry.get("file") or ""
            return Path(str(v)).name if v else ""
        return ""

    main_e = data.get("main")
    known += [fname_of(x) for x in main_e] if isinstance(main_e, list) else [fname_of(main_e)]
    known += [fname_of(s) for s in (data.get("supplements") or [])]
    known = [k for k in known if k]
    for fn in known:
        md = work / "parse" / fn / (Path(fn).stem + ".md")
        if md.is_file():
            src_lines[fn] = len(md.read_text(encoding="utf-8", errors="replace").splitlines())
    return known, src_lines, note


def _check_citation(cell: str, known: list[str], src_lines: dict[str, int],
                    table_anchor: dict) -> list[str]:
    """出处单元格三层校验：①锚定（「；」分段独立锚、「/」继承、「同文件」回填
    表级锚）②文件名 ∈ 来源集合 ③L 行号 ≤ 解析 md 总行数。返回问题列表。"""
    msgs: list[str] = []
    for part in _split_top(cell, "；"):
        anchor: str | None = None
        for seg in (s.strip() for s in part.split("/") if s.strip()):
            fn_hit = next((fn for fn in known
                           if seg == fn or seg.startswith(fn + " ")
                           or seg.startswith(fn + "（") or seg.startswith(fn + "(")), "")
            if fn_hit:
                anchor = fn_hit
                table_anchor["file"] = fn_hit
            elif seg.startswith("同文件"):
                if not table_anchor["file"]:
                    msgs.append("用了「同文件」但本表此前没有出现任何文件名锚——首个出处须写全文件名")
                    anchor = ""
                else:
                    anchor = table_anchor["file"]
            elif _FILENAME_LIKE.match(seg):
                msgs.append(f"文件名「{_FILENAME_LIKE.match(seg).group('fn')}」不在已确认来源"
                            f"集合（{'、'.join(known) if known else '空'}）——核对拼写，"
                            f"不得引用未确认或编造的文件")
                anchor = ""
            elif anchor is None:
                msgs.append("缺文件名或「同文件」前缀——每段出处须能独立定位到文件")
            if anchor:
                nums = [int(x) for x in LNUM_RE.findall(seg)]
                if not nums:
                    msgs.append("缺 L 行号引用（如 L97 / L18-L22）")
                elif anchor in src_lines:
                    over = [n for n in nums if n > src_lines[anchor]]
                    if over:
                        msgs.append(f"行号 L{min(over)} 超出「{anchor}」解析结果总行数 "
                                    f"{src_lines[anchor]}——行号只能来自实际读取的区段")
                else:
                    msgs.append(f"引用的「{anchor}」没有解析产物（解析失败降级的文件不可引用）")
    return msgs


def _coverage_section(text: str) -> str | None:
    """coverage 段文本（段头行到下一个 # 行之前）；没有返回 None。"""
    lines = text.splitlines()
    for i, ln in enumerate(lines):
        if _COVERAGE_HEAD.match(ln.strip()):
            seg = [ln]
            for ln2 in lines[i + 1:]:
                if ln2.lstrip().startswith("#"):
                    break
                seg.append(ln2)
            return "\n".join(seg)
    return None


# --analysis 的表规格：(文件名, 表头关键词, 出处列名候选)
_ANALYSIS_TABLE_SPECS = [
    ("structure.md", ["字段"], ["出处"]),
    ("requirements-format.md", ["章节名称"], ["出处", "线索"]),
    ("requirements-format.md", ["模板名称"], ["出处", "线索"]),
    ("requirements-business.md", ["需求"], ["出处"]),
    ("evaluation.md", ["评分项"], ["出处"]),
    ("disqualification.md", ["条款摘要"], ["出处"]),
]


def check_analysis(work: Path) -> tuple[list[str], list[str]]:
    """分析登记表机器校验：五表存在性 + coverage 声明 + 出处引用三层。
    移植自客户端 validate_analysis（列形按本包六表契约适配）。"""
    issues: list[str] = []
    warnings: list[str] = []
    adir = work / "analysis"
    if not adir.is_dir():
        return ["work/analysis/ 不存在——先走第二步要点提取"], warnings
    known, src_lines, src_note = _known_sources(work)
    if src_note:
        warnings.append(src_note)

    present = sorted(p.name for p in adir.glob("*.md"))
    for f in ANALYSIS_FILES:
        if f not in present:
            issues.append(f"analysis/{f} 缺失（五张登记表必查——对应环节未跑或被删）")
    if "clarifications.md" not in present:
        warnings.append("analysis/clarifications.md 不存在——本项按无澄清处理（确有澄清请登记）")
    for extra in present:
        if extra not in ANALYSIS_FILES and extra != "clarifications.md":
            warnings.append(f"analysis/ 未识别的文件（不在登记表清单）：{extra}")

    n_cited = 0
    for fname, header_kw, src_cols in _ANALYSIS_TABLE_SPECS:
        path = adir / fname
        if not path.is_file():
            continue
        text = _read(path)
        # 头部「> 来源文件：」声明 = 文件级锚种子（本包约定，比客户端多一层；
        # 声明了但不在来源集合 → 照跑但降级点名）
        declared = _declared_source(text)
        if declared and known and declared not in known:
            warnings.append(f"analysis/{fname} 头部声明的来源文件不在 sources.json 集合："
                            f"{declared}——核对是改名还是登记表引用了未确认来源")
        seed = declared if declared in known else ""
        for ti, t in enumerate(_split_tables(table_rows(text), header_kw), 1):
            header = t["header"]
            si = next((i for i, c in enumerate(header)
                       for sc in src_cols if sc in c), -1)
            if si < 0:
                issues.append(f"analysis/{fname} 第{ti}张表缺出处列（列头须含 {'/'.join(src_cols)}）")
                continue
            table_anchor = {"file": seed}  # 表级文件名锚（种子=头部声明），跨行继承
            for ri, r in enumerate(t["rows"], 1):
                r = (r + [""] * len(header))[:len(header)]
                cell = (r[si] or "").strip()
                if not cell:
                    # 空出处只在「实质数据行」上点名——表内分节小标题行（如
                    # 「|技术分|…|」短行）不拦，防真实评分表的嵌行误报
                    if len(norm("".join(r))) >= 8:
                        issues.append(f"analysis/{fname} 表{ti}行{ri} 出处列为空——"
                                      f"每行证据都必须能独立定位到文件")
                    continue
                for msg in _check_citation(cell, known, src_lines, table_anchor):
                    issues.append(f"analysis/{fname} 表{ti}行{ri} {msg}")
                n_cited += 1

    for f in ANALYSIS_FILES:
        path = adir / f
        if not path.is_file():
            continue
        seg = _coverage_section(_read(path))
        if seg is None:
            issues.append(f"analysis/{f} 缺「## coverage」段——「未发现」只能对已检查范围"
                          f"成立，范围必须声明（已检查：…；未检查：…）")
            continue
        missing = [m for m in ("已检查", "未检查") if m not in seg]
        if missing:
            issues.append(f"analysis/{f} coverage 段缺标记 {'、'.join(missing)}"
                          f"（未检查为无也要写「未检查：无」）")
    if not issues:
        print(f"[分析登记表] 出处引用共 {n_cited} 条完成三层校验"
              + ("" if known else "（来源集合降级，仅锚定层生效）"))
    return issues, warnings


# ---------------------------------------------------------------- --outline

def check_outline(work: Path) -> tuple[list[str], list[str]]:
    """目录来源标注对账：悬空 ID=标注引用了登记表没有的编号（issue，多为行序
    漂移或拼写错）；未引用 ID=登记表有条目但目录从未承接（warning，须逐条处置
    点名——移植自客户端 assemble_tender lineage_check）。"""
    issues: list[str] = []
    warnings: list[str] = []
    outline_path = work / "outline" / "tender-response-docs.md"
    if not outline_path.is_file():
        return ["目录文件不存在（work/outline/tender-response-docs.md）"], warnings
    registry, _scores = _registry_payload(work)
    if not registry:
        return ["登记表为空（analysis/ 四表未产出）——来源标注无从对账"], warnings
    referenced: set[str] = set()
    for ann in _parse_annotations(outline_path.read_text(encoding="utf-8").split("\n")).values():
        for rest in ann.get("src:", {}).values():
            for tok in re.split(r"[,，、]", rest):
                tok = tok.strip()
                if tok and ID_RE.match(tok):
                    referenced.add(_norm_id(tok))
    reg_ids = {cid.upper() for cid in registry}
    dangling = sorted(referenced - reg_ids)
    unused = sorted(reg_ids - referenced)
    if dangling:
        issues.append(f"悬空来源ID（标注引用但登记表不存在，多为清单行序漂移或编号拼写"
                      f"错误）：{'、'.join(dangling)}——修正标注或登记表后重跑")
    if unused:
        warnings.append(f"未被任何目录节点引用的来源ID（{len(unused)} 个）："
                        f"{'、'.join(unused[:12])}" + ("…" if len(unused) > 12 else "")
                        + "——逐条处置：能归位的修订目录；确属无需对应章节的（扣分项/"
                          "后续轮次报价等）收尾汇报必须点名列出，不得只字不提")
    return issues, warnings


# ---------------------------------------------------------------- --guide

def check_guide(work: Path) -> tuple[list[str], list[str]]:
    issues: list[str] = []
    warnings: list[str] = []
    outline_path = work / "outline" / "tender-response-docs.md"
    guide_path = work / "body" / "写作指引.md"
    if not outline_path.is_file():
        return ["目录文件不存在（work/outline/tender-response-docs.md）"], warnings
    if not guide_path.is_file():
        return ["写作指引不存在（work/body/写作指引.md）"], warnings

    volumes = parse_outline(outline_path.read_text(encoding="utf-8"))
    multi = len(volumes) > 1
    expected: dict[str, str] = {}      # 全部节点（含容器章）——指引行写容器章也合法
    leaf_keys: set[str] = set()        # 仅叶子——缺行警示只对需要产文件的节
    for vol in volumes:
        # parse_outline 的 nodes 是扁平列表（children 恒空），须先建树再遍历层级
        def walk(ns: list[dict]) -> None:
            for node in ns:
                key = f"{vol['name']}/{node['text']}" if multi else node["text"]
                expected[key] = node["text"]
                if node["children"]:
                    walk(node["children"])
                else:
                    leaf_keys.add(key)
        walk(build_tree(vol["nodes"]))

    rows = table_rows(guide_path.read_text(encoding="utf-8"))
    if not rows or len(rows[0]) < 6:
        return ["指引表缺失或表头列数不足（须为 6 列：节|模式|依据|素材|图示|缺口/备注）"], warnings
    rows = rows[1:]

    seen_keys: set[str] = set()
    material_use: dict[str, list[str]] = {}
    for r in rows:
        r = (r + [""] * 6)[:6]
        key, mode, material, figure = r[0], r[1], r[3], r[4]
        if key not in expected:
            issues.append(f"指引行「{key}」不在目录树节点中（节名须与目录逐字一致"
                          + ("，多册写「册名/标题」" if multi else "") + "）")
            continue
        seen_keys.add(key)
        if mode not in LEGAL_MODES and not all(
                m in {"素材修订", "格式跟随", "推理撰写"} for m in mode.split("+")):
            issues.append(f"「{key}」模式非法：{mode}（素材修订/格式跟随/推理撰写，可 + 组合；非正文节点写 —）")
        if material and material != "—":
            if re.search(r"\.docx\s+L\d+(?:-L?\d+)?\s*、", material):
                # 完整 token 后跟顿号 = 拿顿号当分隔符；文件名内含顿号（顿号在 .docx 前）不受影响
                issues.append(f"「{key}」素材列多个区间用中文逗号分隔（顿号是图示列的分隔符）：{material}")
            else:
                for tok in [t.strip() for t in material.split("，") if t.strip()]:
                    m = MATERIAL_TOKEN.match(tok)
                    if tok == "【缺】":
                        continue
                    if not m:
                        issues.append(f"「{key}」素材列记法非法：{tok}（应为 素材文件.docx L120-L260 / 【缺】 / —）")
                        continue
                    use_key = f"{_base_name(m.group(1))}:L{m.group(2)}-L{m.group(3) or m.group(2)}"
                    material_use.setdefault(use_key, []).append(key)
        if figure and figure != "—":
            for tok in [t.strip() for t in figure.split("、") if t.strip()]:
                kind = tok.split(":", 1)[0]
                if kind not in LEGAL_FIGURES:
                    issues.append(f"「{key}」图示记法非法：{tok}（类型词表：表/分层/甘特）")
        note = r[5] if len(r) > 5 else ""
        if NOTE_TOOL_RE.search(note):
            warnings.append(f"「{key}」缺口/备注列含脚本名或调用参数——该列会摘出独立展示给"
                            f"用户，只写缺口本身，不写实现词")
        if NOTE_ANAPHORA_RE.search(note):
            warnings.append(f"「{key}」缺口/备注列代词回指（上述/以上/前述）——摘出独立展示"
                            f"后读不通，写具体字段名")
    for use_key, where in material_use.items():
        if len(where) > 1:
            issues.append(f"同一素材区间进多节（{use_key}）：{'、'.join(where)}——同一内容区间不得进两节")
    missing = sorted(leaf_keys - seen_keys)
    if missing:
        warnings.append(f"目录叶子无指引行（{len(missing)} 个）：{'、'.join(missing[:8])}"
                        + ("…" if len(missing) > 8 else ""))
    return issues, warnings


# ---------------------------------------------------------------- 图示对账

def figure_gaps(stem: str, row: list[str] | None, doc: Document,
                fallback: bool = False) -> list[str]:
    """指引图示列计划 vs 节内实收（表格 + DrawingML 图片引用）对账（warning 级，
    移植自客户端 validate_body 图示对账）：计划了没产出→点名；没计划却产出且是
    推理撰写节→点名（计划先行）；素材修订/格式跟随节的表图来自注入拷贝，不背
    计划的账（与客户端豁免一致）。"""
    if not row or len(row) < 5:
        return []
    plan = [t.strip() for t in (row[4] or "").split("、")
            if t.strip() and t.strip() != "—"]
    actual = len(doc.tables) + len(doc.element.body.findall(".//" + qn("a:blip")))
    mode = row[1] if len(row) > 1 else ""
    tail = "（单节自查漏跑，收尾兜底）" if fallback else ""
    out: list[str] = []
    if plan and actual < len(plan):
        out.append(f"{stem} 〔图示〕指引计划 {len(plan)} 项、节内表格/图片实收 {actual} 项"
                   f"——未产出项补齐或同步改指引图示列：{'；'.join(plan[:5])}{tail}")
    elif not plan and actual and "推理撰写" in mode:
        out.append(f"{stem} 〔图示〕指引未计划图示、节内有 {actual} 项表格/图片——推理撰写"
                   f"节请补进指引图示列（计划先行），或收尾说明理由{tail}")
    return out


# ---------------------------------------------------------------- --section

def check_section(work: Path, section: Path) -> tuple[list[str], list[str]]:
    issues: list[str] = []
    warnings: list[str] = []
    text, labeled = docx_text(section)

    for loc, frag in labeled:
        for line in frag.splitlines():
            if PLACEHOLDER.search(line):
                issues.append(f"{loc} 占位文字（会随交付稿出去，改用 Word 批注）：{line.strip()[:40]}")
            if MD_RESIDUE.match(line.strip()):
                issues.append(f"{loc} Markdown 残字：{line.strip()[:40]}")
            elif WEAK_MD.match(line.strip()):
                warnings.append(f"{loc} 疑似列表残字（正文应为完整句）：{line.strip()[:40]}")
    warnings += scan_ai_flavor(section.stem, labeled)

    sec_doc = Document(str(section))
    h1 = [pp.text.strip() for pp in sec_doc.paragraphs
          if (pp.style.name or "").startswith(("Heading 1", "标题 1")) and pp.text.strip()]
    if len(h1) > 1:
        issues.append(f"节内有 {len(h1)} 个一级标题（{'、'.join(h1[:3])}）——"
                      f"素材拷贝带入的源章标题未删，合册后会出现双章标题，删到只剩本节标题")
    numbered = [pp.text.strip() for pp in sec_doc.paragraphs
                if (pp.style.name or "").startswith(("Heading 2", "Heading 3", "Heading 4",
                                                     "标题 2", "标题 3", "标题 4"))
                and HEADING_NUMBERING.match(pp.text.strip())]
    if numbered:
        issues.append(f"小标题自带编号（{'、'.join(t[:24] for t in numbered[:3])}）——"
                      f"合册按树深度接续节号，会双重编号；改为陈述性短语，层级用 Heading 2/3 表达")

    # 尾部空段（表后规范空段豁免 1 个；建节不补尾空段，多余的是合册空白页来源）
    kids = [c for c in sec_doc.element.body if c.tag != qn("w:sectPr")]
    tail = 0
    for c in reversed(kids):
        if c.tag != qn("w:p"):
            break
        if any((t.text or "").strip() for t in c.findall(".//" + qn("w:t"))):
            break
        if c.findall(".//" + qn("w:drawing")) or c.findall(".//" + qn("w:commentReference")):
            break
        tail += 1
    if tail:
        prev = kids[len(kids) - tail - 1] if tail < len(kids) else None
        allowed = 1 if prev is not None and prev.tag == qn("w:tbl") else 0
        if tail > allowed:
            warnings.append(f"节尾有 {tail} 个空段落（表后规范空段已豁免）——多余空段在合册"
                            f"分页后会顶出空白页，删掉")

    # 素材使用率与残留提示（从指引行取该节的素材区间）
    guide_path = work / "body" / "写作指引.md"
    leaf_title = section.stem
    if guide_path.is_file():
        rows = table_rows(guide_path.read_text(encoding="utf-8"))
        row = next((r for r in rows if r and norm(r[0].rsplit("/", 1)[-1]) == norm(leaf_title)), None)
        # 缺口 × 批注交叉：指引声明了待补/批注待办，节内就必须有 Word 批注
        if row and len(row) >= 6 and GAP_MARK.search(row[5] or ""):
            n_ref = len(sec_doc.element.body.findall(".//" + qn("w:commentReference")))
            if n_ref == 0:
                issues.append("指引缺口/备注列声明了待补或批注待办，节内 0 条 Word 批注——"
                              "缺料处逐点补批注（参数化措辞回避不替代批注，待拍板项要进"
                              "批注台账）；缺口已解决的先同步清理指引缺口列")
        if row:
            warnings += figure_gaps(section.stem, row, sec_doc)
        if row and len(row) >= 4 and row[3] not in ("", "—", "【缺】"):
            sec_shingles = shingles(text)
            for tok in [t.strip() for t in row[3].split("，") if t.strip()]:
                m = MATERIAL_TOKEN.match(tok)
                if not m:
                    continue
                fname = _base_name(m.group(1))
                lo, hi = int(m.group(2)), int(m.group(3) or m.group(2))
                src_md = work / "parse" / fname / (Path(fname).stem + ".md")
                if not src_md.is_file():
                    warnings.append(f"素材文件未解析，使用率无法核对：{tok}")
                    continue
                lines = src_md.read_text(encoding="utf-8", errors="replace").splitlines()
                seg = "\n".join(lines[lo - 1:hi])
                seg = re.sub(r"^\s*\|?[\s:|-]*-{2,}[\s:|-]*$", "", seg, flags=re.M).replace("|", " ")
                src_shingles = shingles(seg)
                if not sec_shingles or not src_shingles:
                    continue
                overlap = len(sec_shingles & src_shingles) / min(len(sec_shingles), len(src_shingles))
                if overlap < 0.10:
                    warnings.append(f"素材使用率过低（{overlap:.0%}）：{tok}——疑似未实质使用（拷贝后未改写或凭空自写）")
                stem_hint = Path(fname).stem
                if len(stem_hint) >= 4 and norm(stem_hint) not in norm("招标文件") and stem_hint in text:
                    warnings.append(f"疑似残留（素材文件名主干出现在正文）：{stem_hint}——逐条核对，通用术语在收尾说明")
    return issues, warnings


# ---------------------------------------------------------------- 全局

def check_global(work: Path, old_names: list[str]) -> tuple[list[str], list[str]]:
    issues: list[str] = []
    warnings: list[str] = []
    body_dir = work / "body"
    if not body_dir.is_dir():
        return ["work/body/ 不存在"], warnings
    sections = sorted(
        p for p in list(body_dir.glob("*.docx")) + list(body_dir.rglob("*/*.docx"))
        if not p.name.startswith("整本-")
    )
    if not sections:
        warnings.append("work/body/ 下没有节文件")

    texts: dict[Path, tuple[str, list[tuple[str, str]]]] = {}
    for p in sections:
        texts[p] = docx_text(p)

    # 占位清点 + AI 腔扫描 + 双章标题兜底（单节自查漏跑时收尾仍拦得住）
    n_ph = 0
    gdocs: dict[Path, Document] = {}
    for p, (_, labeled) in texts.items():
        for loc, frag in labeled:
            for line in frag.splitlines():
                if PLACEHOLDER.search(line):
                    issues.append(f"{p.name}:{loc} 占位文字：{line.strip()[:40]}")
                    n_ph += 1
        warnings += scan_ai_flavor(p.stem, labeled)
        gdoc = Document(str(p))
        gdocs[p] = gdoc
        h1 = [pp.text.strip() for pp in gdoc.paragraphs
              if (pp.style.name or "").startswith(("Heading 1", "标题 1")) and pp.text.strip()]
        if len(h1) > 1:
            issues.append(f"{p.name} 内有 {len(h1)} 个一级标题（{'、'.join(h1[:3])}）——"
                          f"素材拷贝带入的源章标题未删，合册后会出现双章标题，删到只剩本节标题")
        numbered = [pp.text.strip() for pp in gdoc.paragraphs
                    if (pp.style.name or "").startswith(("Heading 2", "Heading 3", "Heading 4",
                                                         "标题 2", "标题 3", "标题 4"))
                    and HEADING_NUMBERING.match(pp.text.strip())]
        if numbered:
            issues.append(f"{p.name} 小标题自带编号（{'、'.join(t[:24] for t in numbered[:3])}）——"
                          f"合册按树深度接续节号，会双重编号；改为陈述性短语")

    # 缺口 × 批注交叉兜底：指引声明了待补/批注待办的节，docx 里 0 条批注即点名
    # （单节自查漏跑时收尾仍拦得住）
    guide_md = body_dir / "写作指引.md"
    if guide_md.is_file():
        gap_rows = table_rows(guide_md.read_text(encoding="utf-8"))
        for p, gdoc in gdocs.items():
            row = next((r for r in gap_rows if r and norm(r[0].rsplit("/", 1)[-1]) == norm(p.stem)), None)
            if row and len(row) >= 6 and GAP_MARK.search(row[5] or "") \
                    and not gdoc.element.body.findall(".//" + qn("w:commentReference")):
                issues.append(f"{p.name} 指引缺口/备注列声明了待补或批注待办，节内 0 条 Word 批注"
                              f"——缺料处逐点补批注；缺口已解决的先同步清理指引缺口列")
            warnings += figure_gaps(p.name, row, gdoc, fallback=True)

    # 承诺比对 + 旧值快照（清单值变更时旧值自动纳入残留扫描）
    promise_path = body_dir / "关键事实与承诺.md"
    if promise_path.is_file():
        rows = table_rows(promise_path.read_text(encoding="utf-8"))
        current_values: set[str] = set()
        for r in rows:
            r = (r + [""] * 3)[:3]
            item, value = r[0], r[1]
            if not value or value in ("—", "值") or re.search(r"待补|待确认|待定", value):
                continue
            v = norm(value)
            if len(v) < 2:
                continue
            current_values.add(v)
            hit = any(v in norm(full) for full, _ in texts.values())
            if not hit and digit_seq_hit(v, [norm(full) for full, _ in texts.values()]):
                # 数字吻合仅措辞差异：与真漏写分开（拦截级误报会让报警失去分辨力）
                warnings.append(f"承诺值写法差异（数字吻合）：{value}（事项：{item}）"
                                f"——正文未按清单原文落值，人工归一（改正文措辞或更新清单）")
            elif not hit:
                issues.append(f"承诺值未落进任何正文：{value}（事项：{item}）——该节未写或值被改写")
        snap_path = body_dir / PROMISE_SNAPSHOT
        prev: set[str] | None = None
        if snap_path.is_file():
            try:
                prev = set(json.loads(snap_path.read_text(encoding="utf-8")))
            except Exception:
                prev = None
        if prev is not None:
            for old in sorted(prev - current_values):
                for p, (_, labeled) in texts.items():
                    for loc, frag in labeled:
                        if old in norm(frag):
                            issues.append(f"{p.name}:{loc} 承诺旧值残留「{old}」：{frag.strip()[:50]}"
                                          f"——承诺清单已更新，正文该处未同步")
        snap_path.write_text(json.dumps(sorted(current_values), ensure_ascii=False),
                             encoding="utf-8")
    else:
        warnings.append("关键事实与承诺.md 不存在，承诺比对跳过")

    # 节间查重（先剔样板：出现在 ≥max(3, ⌈节数/3⌉) 的 10 字 shingle 视为落款/
    # 格式套语。真实项目实测：落款/封面类样板常只覆盖约半数节（如 17/33），
    # 75% 档阈值剔不动；⌈n/3⌉ 档实测误报趋零且真重叠保留——真拷贝对几乎总是
    # 恰好 2 节共享，与样板阈值之间留有判别空间。剔除后有效特征 <30 的节对
    # 〔短格式件〕不参与——防「短节被长节包含」误报。
    # 多册按册分组比对（不跨册——跨册同名材料是目录阶段已裁量的边界问题），
    # 封面/目录节不参与节间比对（封面按册仅一份、模板性强，互比必高重叠）。）
    counts: Counter[str] = Counter()
    for _, (full, _) in texts.items():
        counts.update(shingles(full))
    boiler_floor = max(3, math.ceil(len(texts) / 3))
    boiler = {s for s, c in counts.items() if c >= boiler_floor}
    dedup_targets = [p for p in sections if p.stem not in ("封面", "目录")]
    sig = {p: shingles(texts[p][0]) - boiler for p in dedup_targets}
    groups: dict[str, list[Path]] = {}
    for p in dedup_targets:
        rel = p.parent.relative_to(body_dir)
        groups.setdefault(rel.parts[0] if rel.parts else ".", []).append(p)
    for members in groups.values():
        items = sorted((p, sig[p]) for p in members)
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                (pa, sa), (pb, sb) = items[i], items[j]
                if min(len(sa), len(sb)) < 30:
                    continue
                ov = len(sa & sb) / min(len(sa), len(sb))
                if ov >= 0.5:
                    issues.append(f"节间大面积重叠（{ov:.0%}）：{pa.name} × {pb.name}——同一素材进多节或原文照搬，拆改其一")

    # 残留扫描
    if old_names:
        for name in old_names:
            name = name.strip()
            if len(name) < 2:
                continue
            for p, (_, labeled) in texts.items():
                for loc, frag in labeled:
                    if name in frag:
                        issues.append(f"{p.name}:{loc} 旧名残留「{name}」：{frag.strip()[:50]}——必须清零")
    return issues, warnings


def main() -> None:
    ap = argparse.ArgumentParser(description="正文机器校验")
    ap.add_argument("--work", default="work")
    ap.add_argument("--guide", action="store_true", help="校验写作指引")
    ap.add_argument("--analysis", action="store_true",
                    help="校验分析登记表（coverage 声明 + 出处引用三层）")
    ap.add_argument("--outline", action="store_true",
                    help="目录来源标注对账（悬空/未引用来源ID）")
    ap.add_argument("--section", help="单节自查（.docx 路径）")
    ap.add_argument("--old-names", default="", help="残留扫描旧名清单（逗号分隔）")
    ap.add_argument("--json", default="", help="结果落盘 JSON（写作进度面板读它展示；"
                    "推荐 work/review/check-result.json，控制台输出与退出码不变）")
    args = ap.parse_args()
    work = Path(args.work)

    if args.guide:
        issues, warnings = check_guide(work)
    elif args.section:
        sec = Path(args.section)
        if not sec.is_file():
            die(f"节文件不存在：{sec}")
        issues, warnings = check_section(work, sec)
    elif args.analysis:
        issues, warnings = check_analysis(work)
    elif args.outline:
        issues, warnings = check_outline(work)
    else:
        old_names = [x for x in args.old_names.split(",") if x.strip()]
        names_file = Path("旧名名单.txt")
        if names_file.is_file():
            old_names += [l.strip() for l in names_file.read_text(encoding="utf-8").splitlines()
                          if l.strip()]
        issues, warnings = check_global(work, old_names)

    if args.json:
        out = Path(args.json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({
            "at": time.strftime("%Y-%m-%d %H:%M"),
            "mode": ("guide" if args.guide else "section" if args.section
                     else "analysis" if args.analysis else "outline" if args.outline
                     else "global"),
            "issues": issues, "warnings": warnings,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[结果已落盘] {out}")

    for w in warnings:
        print(f"⚠️ {w}")
    if not issues:
        print("[校验通过] 全部合规")
        sys.exit(0)
    print(f"[校验未通过] 共 {len(issues)} 处问题：")
    for i in issues:
        print(f"- {i}")
    sys.exit(1)


if __name__ == "__main__":
    main()
