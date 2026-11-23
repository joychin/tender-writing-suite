#!/usr/bin/env python3
"""build_review.py — 组装评审台数据并注入 HTML 模板，生成静态评审台页面。

输入（work/ 既有产物，缺什么少什么就少渲染什么，不报错）：
- outline/tender-response-docs.md   目录树 + 来源标注 + 目录说明 + 项目信息
- body/写作指引.md、body/关键事实与承诺.md
- analysis/clarifications.md（含 `## 处理记录` 段，旧「裁决记录」兼容）、analysis/*.md 登记表（编号=行序）
- analysis/structure.md / requirements-format.md / requirements-business.md（总览只读展示区）
- parse/sources.json + parse/*/<stem>.meta.json 解析概况
- body/**/*.docx 判断「已写」状态；decisions.json 确认台账

输出：work/review/评审台.html（file:// 可直接打开的静态版；离线导出修改包用）
与 work/review/招标文件分析报告.html（只读报告页——解析/体检完成时刻的展示产物，
零交互零按钮，自包含单文件可转发）。
G3 起（--pages all 且目录已定）另产 work/review/写作进度.html（写作阶段监视面——
进度/批注台账/待拍板承诺/校验结果，服务模式 30s 自刷新，见 progress_payload）。
serve_review.py 复用本模块的 collect_payload / render_html / render_report /
render_progress 做实时注入。

安全声明：纯本地文件读取，无网络、无外部进程。
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
# 环节间共享模块解析（三形态）：单技能扁平包（全部环节脚本同目录）→ 五技能
# 平级 ../<环节> → stages/<环节> 收纳布局；都缺才报错。
# 本段是发货源码：打包器不再装配时替换——本地电池测到的就是发货版。
_SCRIPTS_DIR = Path(__file__).resolve().parent
_STAGE_ROOT = _SCRIPTS_DIR.parent
_SENTINEL = {"tender-docx": "docx_toolkit.py", "tender-writing": "check_writing.py"}


def _sibling_scripts(name: str) -> Path:
    if (_SCRIPTS_DIR / _SENTINEL[name]).is_file():
        return _SCRIPTS_DIR
    for cand in (_STAGE_ROOT.parent / name, _STAGE_ROOT / "stages" / name):
        scripts = cand / "scripts"
        if scripts.is_dir():
            return scripts
    raise SystemExit(
        f"缺少兄弟环节 {name}（单包 / 平级 / stages 三形态均未找到）：{_STAGE_ROOT.parent}"
    )


for _sib in ("tender-writing", "tender-docx"):
    sys.path.append(str(_sibling_scripts(_sib)))
from apply_review import GUIDE_RELPATH, OUTLINE_RELPATH, PROMISE_RELPATH, _sha256, _volumes
from check_writing import (ID_RE, _parse_annotations, _read, _registry_payload,
                           _resolve_source, _split_tables, table_rows)
from docx_toolkit import para_text_full, require_docx, sanitize_name

TEMPLATE = Path(__file__).parent.parent / "templates" / "review-workbench.html"
REPORT_TEMPLATE = Path(__file__).parent.parent / "templates" / "analysis-report.html"
PROGRESS_TEMPLATE = Path(__file__).parent.parent / "templates" / "progress.html"
# 行首锚定：只认真实记录行，避免把说明文字/示例里的 `- CLAR-01 :: 处理=…` 当成记录；
# 值一律止于行尾（[^|\n]*），否则「时间=」会跨行吞掉下一条记录
VERDICT_RE = re.compile(r"^\s*-\s*(CLAR-\d{1,3})\s*::\s*(?:处理|裁决)=([^|\n]*)\s*(?:\|\s*备注=([^|\n]*))?(?:\|\s*时间=([^|\n]*))?", re.M)
# 占位值（…、...、···）：是格式示例不是用户答复
VERDICT_PLACEHOLDER_RE = re.compile(r"^[…·．.\s]+$")
# （ID_RE / _read / _resolve_source / _split_tables / _registry_payload 已下沉
# check_writing——1.8.0，check_writing 为纯函数落点，依赖方向恒为上游→它）


def _outline_payload(work: Path) -> dict:
    path = work / OUTLINE_RELPATH
    text = _read(path)
    if not text:
        return {"project": {}, "volumes": []}
    ann_all = _parse_annotations(text.split("\n"))
    vols = _volumes(text)
    # scope：卷头行与第一个 ## 段之间的正文
    lines = text.split("\n")
    scopes: dict[int, str] = {}
    vi, buf, closed = -1, [], False
    for ln in lines:
        if ln.startswith("# 响应文件："):
            if vi >= 0:
                scopes[vi] = " ".join(s.strip() for s in buf if s.strip())
            vi += 1
            buf, closed = [], False
        elif vi >= 0 and not closed:
            if ln.startswith("## "):
                closed = True  # 卷头与第一个 ## 段之间才是编制说明
            elif not ln.startswith("# "):
                buf.append(ln)
    if vi >= 0:
        scopes[vi] = " ".join(s.strip() for s in buf if s.strip())
    # 项目信息（首个 ## 段之前的 - key：value 行）
    project: dict[str, str] = {}
    in_proj = False
    for ln in lines:
        if ln.startswith("## "):
            in_proj = ln[3:].strip() == "项目信息"
            continue
        if in_proj and ln.lstrip().startswith("- ") and "：" in ln:
            k, v = ln.strip()[2:].split("：", 1)
            project.setdefault(k.strip(), v.strip())

    def desc_dict(rest: str) -> dict:
        out = {}
        for part in rest.split("|"):
            if "=" in part:
                k, v = part.split("=", 1)
                out[k.strip()] = v.strip()
        return out

    def src_ids(rest: str) -> list[str]:
        return [x.strip() for x in re.split(r"[,，、]", rest) if ID_RE.match(x.strip())]

    volumes = []
    for i, v in enumerate(vols):
        ann = ann_all.get(i, {"src:": {}, "desc:": {}})

        def wrap(ns):
            return [{"text": n["text"],
                     "src": src_ids(ann["src:"].get(n["text"], "")),
                     "desc": desc_dict(ann["desc:"].get(n["text"], "")) if n["text"] in ann["desc:"] else None,
                     "children": wrap(n["children"])} for n in ns]

        volumes.append({"name": v["name"], "scope": scopes.get(i, ""), "nodes": wrap(v["roots"])})
    return {"project": project, "volumes": volumes}


def _table_rows_of(work: Path, relpath: str) -> list[list[str]]:
    text = _read(work / relpath)
    return table_rows(text) if text else []


# （_resolve_source / _split_tables / _registry_payload 已下沉 check_writing——1.8.0）


def _clar_payload(work: Path, decisions: dict) -> dict:
    items: list[dict] = []
    rows = _table_rows_of(work, "analysis/clarifications.md")
    if rows:
        header = [c.strip() for c in rows[0]]
        for r in rows[1:]:
            r = (r + [""] * len(header))[:len(header)]
            if not r[0].upper().startswith("CLAR"):
                continue
            item = {"id": r[0].strip().upper()}
            for j, h in enumerate(header):
                if "类型" in h:
                    item["type"] = r[j]
                elif "问题" in h:
                    item["question"] = r[j]
                elif "影响" in h:
                    item["impact"] = r[j]
                elif "建议" in h or "动作" in h:
                    item["action"] = r[j]
            items.append(item)
    verdicts: dict[str, dict] = {}
    text = _read(work / "analysis/clarifications.md")
    sec = re.search(r"^##\s*(?:处理|裁决)记录\s*$", text, re.M)
    if sec:
        for m in VERDICT_RE.finditer(text[sec.end():]):
            verdict = m.group(2).strip()
            if not verdict or VERDICT_PLACEHOLDER_RE.match(verdict):
                continue  # 格式示例行（处理=…）不是用户答复，不生成「已处理」
            verdicts[m.group(1).upper()] = {
                "verdict": verdict, "note": (m.group(3) or "").strip(),
                "at": (m.group(4) or "").strip()}
    for cid, v in (decisions.get("clarifications") or {}).items():
        cid = cid.upper()
        cur = verdicts.get(cid)
        cand = {"verdict": v.get("verdict", ""), "note": v.get("note", ""), "at": v.get("at", "")}
        # 两份台账（处理记录 md / decisions.json）正常情况下同步写入；若分叉，取时间较新的一条
        if cur and cur.get("at") and cand["at"] and cand["at"] < cur["at"]:
            continue
        verdicts[cid] = cand
    return {"items": items, "verdicts": verdicts}


def _parse_payload(work: Path) -> list[dict]:
    out = []
    pdir = work / "parse"
    roles: dict[str, str] = {}
    src = pdir / "sources.json"
    if src.is_file():
        try:
            sj = json.loads(src.read_text(encoding="utf-8"))
            roles[sj.get("main", "")] = "main"
            for s in sj.get("supplements", []):
                roles[s.get("file", "")] = s.get("role", "补充")
        except Exception:
            pass
    if pdir.is_dir():
        for meta_p in sorted(pdir.glob("*/*.meta.json")):  # parse_document 写 <stem>.meta.json 于 <源文件名>/ 子目录
            mf = meta_p.parent
            try:
                m = json.loads(meta_p.read_text(encoding="utf-8"))
            except Exception:
                continue
            out.append({"file": m.get("source", mf.name),
                        "role": roles.get(m.get("source", ""), ""),
                        "chars": m.get("chars"), "pages": m.get("pages"),
                        "conversion": m.get("conversion"),
                        "warnings": [w for w in (m.get("warnings") or []) if w]})
    return out


def _disq_payload(work: Path) -> tuple[int | None, list[dict]]:
    """废标登记全量解析：表格形态（条款摘要｜触发条件｜后果｜出处）与顶层 bullet 双支持。

    analysis-lite 约定废标清单写成表格；bullet 兜底容纳手写形态。返回 (条目数, 全量条目)。
    """
    text = _read(work / "analysis/disqualification.md")
    if not text:
        return None, []
    items: list[dict] = []

    def mk(t: str, trigger: str = "", consequence: str = "", source: str = "") -> dict:
        if not source:
            lm = re.search(r"(（L[^）]*）|L\d+[-L\d]*)", t)
            source = lm.group(1) if lm else ""
        return {"text": t[:160], "trigger": (trigger or "")[:100],
                "consequence": (consequence or "")[:100], "source": source}

    rows = _table_rows_of(work, "analysis/disqualification.md")
    # 头行判定：整行都是已知列名（存量项目表头是「编号」/「#」，非「条款摘要」）。
    # 文件允许多段表（## A. 资格审查…），table_rows 会拼起全部行——头行须全程过滤，不只 rows[0]。
    _DQ_HEADER = re.compile(r"编号|序号|#|条款摘要|条款|摘要|触发条件|后果|出处")
    for r in rows:
        cells = [c.strip() for c in (list(r) + [""] * 4)[:4]]
        if not cells[0]:
            continue
        if all(_DQ_HEADER.fullmatch(c) for c in cells if c):
            continue
        items.append(mk(cells[0], cells[1], cells[2], cells[3]))
    if not items:  # 无表格 → 顶层 bullet 兜底
        for ln in text.split("\n"):
            s = ln.strip()
            if s.startswith("- ") and len(s) > 4 and not s.startswith("- 覆盖"):
                items.append(mk(s[2:].strip()))
    return len(items), items


def _coverage_of(text: str) -> str:
    m = re.search(r"coverage\s*声明[:：]\s*(.+)", text)
    return m.group(1).strip() if m else ""


def _section(text: str, heading: str) -> str:
    """`## <heading>…` 段正文（到下一个 `## ` 为止）；无该段返回空串。"""
    m = re.search(rf"^##\s*{re.escape(heading)}.*$", text, re.M)
    if not m:
        return ""
    nxt = re.search(r"^##\s", text[m.end():], re.M)
    return text[m.end():m.end() + nxt.start()] if nxt else text[m.end():]


def _structure_payload(work: Path) -> dict | None:
    """structure.md（字段｜值｜出处）→ 项目基本情况 kv 行；文件缺失返回 None（页面空态）。"""
    text = _read(work / "analysis/structure.md")
    if not text:
        return None
    rows: list[dict] = []
    for r in table_rows(text):
        cells = [c.strip() for c in (list(r) + ["", "", ""])[:3]]
        if not cells[0] or (cells[0] in {"字段", "项目", "条目"} and cells[1] in {"值", "内容", "说明"}):
            continue  # 表头行
        rows.append({"k": cells[0], "v": cells[1], "src": cells[2]})
    return {"rows": rows, "coverage": _coverage_of(text)}


def _fmt_payload(work: Path) -> dict | None:
    """requirements-format.md 三节全量：一、结构要求（散文+嵌套大纲）；二、必须章节表；三、模板表。"""
    text = _read(work / "analysis/requirements-format.md")
    if not text:
        return None
    prose: list[str] = []
    outline: list[dict] = []
    outline_note = ""
    for ln in _section(text, "一、").split("\n"):
        st = ln.strip()
        if not st:
            continue
        if st.startswith("```") or st.startswith("~~~"):
            continue  # 代码围栏行（缩进大纲的书写变体）不进散文，围栏内 - 行仍按大纲收
        if st.startswith("- "):
            depth = (len(ln) - len(ln.lstrip(" "))) // 2
            outline.append({"depth": depth, "text": st[2:].strip()})
        elif st.startswith("（大纲出处"):
            outline_note = st
        elif not st.startswith("结构大纲"):
            prose.append(st)
    mand = tpl = None
    for t in _split_tables(table_rows(text), ["章节名称"]):
        mand = {"header": t["header"], "rows": t["rows"]}
        break
    for t in _split_tables(table_rows(text), ["模板名称"]):
        tpl = {"header": t["header"], "rows": t["rows"]}
        break
    return {"prose": prose, "outline": outline, "outline_note": outline_note,
            "mand": mand, "tpl": tpl, "coverage": _coverage_of(text)}


def _biz_payload(work: Path) -> dict | None:
    """requirements-business.md（需求｜出处）→ 商务技术要求全量；REQ 编号=行序（与登记表同规）。"""
    text = _read(work / "analysis/requirements-business.md")
    if not text:
        return None
    rows: list[dict] = []
    n = 0
    for r in table_rows(text):
        cells = [c.strip() for c in (list(r) + ["", ""])[:2]]
        if not cells[0] or (cells[0].startswith("需求") and "出处" in cells[1]):
            continue  # 表头行
        n += 1
        rows.append({"id": f"REQ-{n:02d}", "text": cells[0], "src": cells[1]})
    return {"rows": rows, "coverage": _coverage_of(text)}


def _written_keys(work: Path, rows: list[dict]) -> list[str]:
    stems = {p.stem for p in (work / "body").rglob("*.docx")} if (work / "body").is_dir() else set()
    return [r["key"] for r in rows if sanitize_name(r["key"].split("/")[-1]) in stems]


def collect_payload(work: Path) -> dict:
    work = Path(work)
    outline = _outline_payload(work)
    guide_rows = []
    rows = _table_rows_of(work, GUIDE_RELPATH)
    if rows:
        header = [c.strip() for c in rows[0]]
        for r in rows[1:]:
            r = (r + [""] * 6)[:6]
            if not r[0].strip():
                continue
            guide_rows.append({"key": r[0], "mode": r[1], "basis": r[2],
                               "material": r[3], "figure": r[4], "note": r[5]})
    promise_rows = []
    rows = _table_rows_of(work, PROMISE_RELPATH)
    if rows:
        for r in rows[1:]:
            r = (r + [""] * 3)[:3]
            if not r[0].strip():
                continue
            promise_rows.append({"item": r[0], "value": r[1], "note": r[2]})
    decisions = {}
    dp = work / "decisions.json"
    if dp.is_file():
        try:
            decisions = json.loads(dp.read_text(encoding="utf-8"))
        except Exception:
            decisions = {}
    registry, scores = _registry_payload(work)
    disq_n, disq_top = _disq_payload(work)

    def sha_or(relpath: str) -> str:
        p = work / relpath
        return _sha256(_read(p)) if p.is_file() else ""

    return {
        "generated_at": _now(),
        "project": outline["project"],
        "volumes": outline["volumes"],
        "base": {"outline": sha_or(OUTLINE_RELPATH), "guide": sha_or(GUIDE_RELPATH),
                 "promise": sha_or(PROMISE_RELPATH)},
        "guide": {"rows": guide_rows},
        "promise": {"rows": promise_rows},
        "clar": _clar_payload(work, decisions),
        "structure": _structure_payload(work),
        "fmt": _fmt_payload(work),
        "biz": _biz_payload(work),
        "registry": registry,
        "scores": scores,
        "parse": _parse_payload(work),
        "disq_count": disq_n,
        "disq_items": disq_top,
        "written_keys": _written_keys(work, guide_rows),
        "decisions": {"gates": decisions.get("gates", []),
                      "clarifications": decisions.get("clarifications", {})},
        "draft": None,
    }


def _now() -> str:
    import time
    return time.strftime("%Y-%m-%d %H:%M")


def _serve_base(work: Path) -> str:
    """serve.json（{port,token,pid}）→ 评审台服务基址；服务未启动返回空串。"""
    try:
        sj = json.loads((work / "review" / "serve.json").read_text(encoding="utf-8"))
        return f"http://127.0.0.1:{sj['port']}/{sj['token']}/"
    except Exception:
        return ""


# ---------------------------------------------------------------- 产物指纹

_FP_EXCLUDE = ("review/serve.json", "review/draft.json", "review/.progress-opened")
_FP_EXCLUDE_PREFIX = ("review/backup/",)


def fingerprint(work: Path) -> str:
    """工作目录产物指纹：全部产物文件的 relpath+mtime_ns+size 聚合 sha256（前 16 位）。

    serve_review 的 healthz 心跳与写作进度面板自刷新共用一个来源；评审台 GET
    渲染时把它注入 DATA 作基线，apply 回包带应用后的新指纹，页面保存后基线
    自然同步（不把自己的保存误报成外部更新）。
    排除 serve.json（重启即变）与 draft.json（页面草稿暂存高频写）；
    review/backup/ 只在 apply 时新增、apply 回包会同步基线，计入无碍。
    """
    import hashlib
    work = Path(work)
    h = hashlib.sha256()
    for p in sorted(work.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(work).as_posix()
        if rel in _FP_EXCLUDE or rel.startswith(_FP_EXCLUDE_PREFIX):
            continue
        st = p.stat()
        h.update(f"{rel}:{st.st_mtime_ns}:{st.st_size}\n".encode())
    return h.hexdigest()[:16]


# ---------------------------------------------------------------- 写作进度面板

PEND_VALUE_RE = re.compile(r"待补|待确认|待定")


def _comments_scan(work: Path) -> list[dict]:
    """扫 body/ 各节 docx 的 Word 批注（含合册产物排除）。

    取数路径照 docx_toolkit.read_comments/doc_comments（rels → comments part），
    另取锚定段落文本（段落内 commentReference 对 id，照 cmd_read 的扫描写法）。
    逐文件 try/except：坏文件跳过不拖垮整页。

    python-docx 惰性加载：没有节文件时无需扫也无需依赖（纯 md 也能出评审台页）；
    真有节文件却缺依赖时走 require_docx 的中文可执行提示，不静默漏报批注。"""
    out = []
    body = work / "body"
    if not body.is_dir():
        return out
    cands = [p for p in sorted(body.rglob("*.docx"))
             if not (p.stem.startswith("整本-") or p.name.startswith("~$"))]
    if not cands:
        return out
    try:
        import docx
        from docx.oxml.ns import qn
    except ImportError:
        require_docx()
        return out
    for path in cands:
        try:
            doc = docx.Document(str(path))
            cmap = {}
            for rel in doc.part.rels.values():
                if rel.reltype.endswith("/comments"):
                    from docx.oxml import parse_xml
                    root = parse_xml(rel.target_part.blob)
                    for c in root.findall(qn("w:comment")):
                        cid = c.get(qn("w:id"))
                        cmap[cid] = "".join(t.text or "" for t in c.iter(qn("w:t")))
            if not cmap:
                continue
            anchors = {}
            for para in doc.paragraphs:
                for ref in para._p.findall(".//" + qn("w:commentReference")):
                    cid = ref.get(qn("w:id"))
                    if cid in cmap:
                        text = para_text_full(para._p).strip()
                        anchors[cid] = (text[:40] + "…") if len(text) > 40 else text
            for cid, ctext in cmap.items():
                out.append({"file": path.stem, "text": ctext, "anchor": anchors.get(cid, "")})
        except Exception:
            continue
    return out


def progress_payload(work: Path) -> dict:
    """写作进度面板的独立采集（不进 collect_payload——批注扫描别拖慢评审台 GET）。"""
    work = Path(work)
    outline = _outline_payload(work)
    stems = {p.stem for p in (work / "body").rglob("*.docx")} if (work / "body").is_dir() else set()

    def leaves(nodes):
        for n in nodes:
            if n.get("children"):
                yield from leaves(n["children"])
            elif n["text"] not in ("封面", "目录"):
                yield n["text"]

    vols = []
    for vol in outline["volumes"]:
        names = list(leaves(vol["nodes"]))
        written = [t for t in names if sanitize_name(t) in stems]
        vols.append({"vol": vol.get("vol", ""), "total": len(names), "written": len(written),
                     "pending": [t for t in names if sanitize_name(t) not in stems]})
    total = sum(v["total"] for v in vols)
    written_n = sum(v["written"] for v in vols)

    promise_rows = []
    rows = _table_rows_of(work, PROMISE_RELPATH)
    if rows:
        for r in rows[1:]:
            r = (r + [""] * 3)[:3]
            if r[0].strip():
                promise_rows.append({"item": r[0], "value": r[1], "note": r[2]})
    pending_proms = [p for p in promise_rows
                     if PEND_VALUE_RE.search(p["value"] or "") or not (p["value"] or "").strip() or p["value"] == "—"]

    comments = _comments_scan(work)

    check = None
    cp = work / "review" / "check-result.json"
    if cp.is_file():
        try:
            check = json.loads(cp.read_text(encoding="utf-8"))
        except Exception:
            check = None

    st = _structure_payload(work) or {}
    strow = {r["k"]: r.get("v", "") for r in (st.get("rows") or [])}
    docno = strow.get("项目编号") or strow.get("招标编号") or ""

    return {
        "generated_at": _now(),
        "docno": docno,
        "project": outline["project"],
        "live": False,
        "vols": vols,
        "totals": {"written": written_n, "total": total,
                   "pending_promises": len(pending_proms),
                   "comments": len(comments),
                   "issues": len((check or {}).get("issues", []))},
        "comments": comments,
        "promises": pending_proms,
        "check": check,
    }


def render_progress(payload: dict, live: bool = False) -> str:
    """写作进度面板：监视面渲染出口。live=True（服务模式）开指纹门控自刷新。

    自刷新由模板内 JS 实现：每 30s 探一次 healthz 产物指纹，变了才整页重载
    （保滚动位置）；指纹缺失（旧服务/静态版）时 DATA.live 之下退化为 30s 盲刷。
    """
    tpl = PROGRESS_TEMPLATE.read_text(encoding="utf-8")
    data = dict(payload)
    data["live"] = live
    data_json = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    return tpl.replace("__REVIEW_DATA__", data_json)


def render_html(payload: dict, api_base: str = "") -> str:
    tpl = TEMPLATE.read_text(encoding="utf-8")
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    return tpl.replace("__REVIEW_DATA__", data_json).replace("__REVIEW_API__", api_base)


WORKBENCH_FILENAME = "评审台.html"


def render_report(payload: dict, workbench: bool = False) -> str:
    """分析报告页：同一 payload 的只读渲染出口（方向A 纸感公文体，零交互）。

    workbench=True 时页脚渲染指向评审台的相对链接（两页同目录，file:// 与服务模式
    行为一致）；False 时维持纯文字指引（G1/G2 阶段评审台尚未生成，不放死链）。
    """
    tpl = REPORT_TEMPLATE.read_text(encoding="utf-8")
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    link = (f'进入写作流程后确认动作在 <a href="{WORKBENCH_FILENAME}">评审页面</a>。'
            if workbench else "进入写作流程后确认动作在评审页面。")
    return tpl.replace("__REVIEW_DATA__", data_json).replace("__WORKBENCH_LINK__", link)


def main() -> None:
    import argparse
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(errors="replace")  # 中文 Windows 管道（GBK）下 ⚠️ 会编码崩，降级输出
    ap = argparse.ArgumentParser(description="生成评审台/分析报告静态页面（file:// 可开，离线导出修改包用）")
    ap.add_argument("--work", default="work", help="工作目录（默认 work）")
    ap.add_argument("--out", default="", help="输出路径（默认 <work>/review/评审台.html）")
    ap.add_argument("--pages", choices=["all", "report"], default="all",
                    help="report=只产分析报告（分析阶段 G1/G2）；all=连评审台一起（G3 目录确认起）")
    ap.add_argument("--open-progress", action="store_true",
                    help="首开写作进度面板（第五步首节开写前的第一动作）：按 --pages all "
                         "构建后落首开标记 work/review/.progress-opened 并打印面板链接回显——"
                         "未首开时 docx_toolkit 与 check_writing --section 会硬拦 body 节写入")
    args = ap.parse_args()
    work = Path(args.work)
    if args.open_progress:
        args.pages = "all"
    payload = collect_payload(work)
    report = work / "review" / "招标文件分析报告.html"
    report.parent.mkdir(parents=True, exist_ok=True)
    # 本轮会产评审台（--pages all）或工作目录里已有（此前 G3+ 生成过），报告页脚才带回链
    with_workbench = args.pages == "all" or (work / "review" / WORKBENCH_FILENAME).exists()
    report.write_text(render_report(payload, workbench=with_workbench), encoding="utf-8")
    tail = "，页脚含评审台回链" if with_workbench else ""
    print(f"[分析报告已生成] {report}（只读展示产物——分析阶段的唯一交付面，可转发{tail}）")
    if args.pages == "all":
        out = Path(args.out) if args.out else work / "review" / WORKBENCH_FILENAME
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(render_html(payload), encoding="utf-8")
        print(f"[评审台已生成] {out}（G3 目录确认起的确认面；总览有分析报告入口）")
        print("离线版：浏览器直接打开，编辑后点「导出改动文件」把文件发回给助手应用。")
        if payload.get("volumes"):
            prog = work / "review" / "写作进度.html"
            pp = progress_payload(work)
            prog.write_text(render_progress(pp), encoding="utf-8")
            t = pp["totals"]
            print(f"[写作进度已生成] {prog}（写作阶段监视面——第五步首节开写前须首开，"
                  f"未首开写节被硬拦；已写 {t['written']}/{t['total']} 节 · "
                  f"批注 {t['comments']} 条 · 待拍板 {t['pending_promises']} 项；"
                  f"服务模式 /progress.html 自动刷新（数据变化才重载））")
            if args.open_progress:
                marker = work / "review" / ".progress-opened"
                marker.write_text(json.dumps({"at": _now()}), encoding="utf-8")
                print(f"[进度面板已首开] 标记已落盘 {marker}——写节首开闸解除"
                      "（docx_toolkit / check_writing --section 放行）。")
                base = _serve_base(work)
                fp8 = fingerprint(work)[:8]
                link = f"{base}progress.html?v={fp8}" if base else str(prog)
                print(f"★ 把写作进度面板以 Markdown 链接发给用户再写第一节"
                      f"{'（链接带版本参数，宿主对相同 URL 不重载，带参点击即强制刷新）' if base else '（静态模式发文件路径）'}："
                      f"[打开写作进度面板]({link})。链接后带一句预期管理："
                      "「正文开始连续生成，进度随时在这个面板看；写完或遇到要你定的事我再停」。")
    base = _serve_base(work)
    if base:
        print(f"[交付链接] {base}?v={fingerprint(work)[:8]}"
              "（带版本参数——宿主预览面板对相同 URL 不重新加载，带参链接点击即强制重载出最新数据；"
              "按门追加锚点 #tab=tree / #tab=guide / #tab=promise / #tab=overview；"
              "每道门重跑本命令后复制新的回显链接，勿复用旧链接）")
    print("[停轮提醒] 以 Markdown 链接或文件指引让用户打开页面，本轮停轮等答复（指针带「需要你补充」表格；固定告知：不补充也可以继续——缺失部分以 Word 批注标注）"
          "（用户已明确授权连续模式除外）。"
          "用户回「继续/先这样」只越过本门等待、推进到下一门产物落盘为止，"
          "下一门交付后照常停轮——门不得同轮连过。")


if __name__ == "__main__":
    main()
