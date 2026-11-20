#!/usr/bin/env python3
"""build_review.py — 组装评审台数据并注入 HTML 模板，生成静态评审台页面。

输入（work/ 既有产物，缺什么少什么就少渲染什么，不报错）：
- outline/tender-response-docs.md   目录树 + 来源标注 + 目录说明 + 项目信息
- body/写作指引.md、body/关键事实与承诺.md
- analysis/clarifications.md（含 `## 裁决记录` 段）、analysis/*.md 登记表（编号=行序）
- parse/sources.json + parse/*/meta.json 解析概况
- body/**/*.docx 判断「已写」状态；decisions.json 确认台账

输出：work/review/评审台.html（file:// 可直接打开的静态版；离线导出修改包用）。
serve_review.py 复用本模块的 collect_payload / render_html 做实时注入。

安全声明：纯本地文件读取，无网络、无外部进程。
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from apply_review import (GUIDE_RELPATH, OUTLINE_RELPATH, PROMISE_RELPATH,
                          _parse_annotations, _sha256, _volumes)
from check_writing import table_rows
from docx_toolkit import sanitize_name

TEMPLATE = Path(__file__).parent.parent / "templates" / "review-workbench.html"
ID_RE = re.compile(r"^(MAND|TPL|REQ|SCORE)-\d{1,3}$", re.I)
LNUM_RE = re.compile(r"L(\d+)")
VERDICT_RE = re.compile(r"-\s*(CLAR-\d{1,3})\s*::\s*裁决=([^|]*)\s*(?:\|\s*备注=([^|]*))?(?:\|\s*时间=([^|]*))?")


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8") if path.is_file() else ""


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
        rows = table_rows(_read(path))
        if not rows:
            continue
        # 按表头关键词拆分同文件的多张表；其它登记表的表头行会关闭当前表
        # （requirements-format.md 里 MAND 表与 TPL 表相邻，不关表会把对方的行吞进来）
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
            elif "序号" in joined and has_src:
                cur = None  # 另一张登记表的表头——关闭当前表
            elif cur is not None:
                cur["rows"].append([c.strip() for c in r])
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
    sec = re.search(r"^##\s*裁决记录\s*$", text, re.M)
    if sec:
        for m in VERDICT_RE.finditer(text[sec.end():]):
            verdicts[m.group(1).upper()] = {
                "verdict": m.group(2).strip(), "note": (m.group(3) or "").strip(),
                "at": (m.group(4) or "").strip()}
    for cid, v in (decisions.get("clarifications") or {}).items():
        verdicts[cid.upper()] = {"verdict": v.get("verdict", ""), "note": v.get("note", ""),
                                 "at": v.get("at", "")}
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
        for mf in sorted(pdir.glob("*/")):
            meta_p = mf / "meta.json"
            if not meta_p.is_file():
                continue
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
    text = _read(work / "analysis/disqualification.md")
    if not text:
        return None, []
    items = []
    for ln in text.split("\n"):
        s = ln.strip()
        if s.startswith("- ") and len(s) > 4 and not s.startswith("- 覆盖"):
            t = s[2:].strip()
            src = ""
            lm = re.search(r"(（L[^）]*）|L\d+[-L\d]*)", t)
            if lm:
                src = lm.group(1)
            items.append({"text": t[:80], "source": src})
    return len(items), items[:3]


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
        "registry": registry,
        "scores": scores,
        "parse": _parse_payload(work),
        "disq_count": disq_n,
        "disq_top": disq_top,
        "written_keys": _written_keys(work, guide_rows),
        "decisions": {"gates": decisions.get("gates", []),
                      "clarifications": decisions.get("clarifications", {})},
        "draft": None,
    }


def _now() -> str:
    import time
    return time.strftime("%Y-%m-%d %H:%M")


def render_html(payload: dict, api_base: str = "") -> str:
    tpl = TEMPLATE.read_text(encoding="utf-8")
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    return tpl.replace("__REVIEW_DATA__", data_json).replace("__REVIEW_API__", api_base)


def main() -> None:
    import argparse
    ap = argparse.ArgumentParser(description="生成评审台静态页面（file:// 可开，离线导出修改包用）")
    ap.add_argument("--work", default="work", help="工作目录（默认 work）")
    ap.add_argument("--out", default="", help="输出路径（默认 <work>/review/评审台.html）")
    args = ap.parse_args()
    work = Path(args.work)
    payload = collect_payload(work)
    out = Path(args.out) if args.out else work / "review" / "评审台.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render_html(payload), encoding="utf-8")
    print(f"[评审台已生成] {out}")
    print("离线版：浏览器直接打开，编辑后点「导出修改包」把文件发回给助手应用。")


if __name__ == "__main__":
    main()
