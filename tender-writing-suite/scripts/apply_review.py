#!/usr/bin/env python3
"""apply_review.py — 评审台修改包应用引擎（落盘 md + 备份 + 结构 diff 报告）。

职责边界：
- 破坏性写入（目录树 / 写作指引 / 承诺清单）先校验基底 sha，写前留备份（每文件留 3 份）；
- 追加性写入（澄清裁决、确认门记录）只追加/定点替换，不整写文件；
- 引擎只落盘与报 diff，不跑 check_writing（调用方在应用后自行重跑校验）。

修改包 schema（op 制，完整示例见 templates/example-review-patch.json）：
{
  "base": {"outline": "<sha256|null>", "guide": "...", "promise": "..."},
  "tree":         [{"vol": "册名", "ops": [{"op": "move|rename|add|remove", ...}]}],
  "guide":        [{"key": "册名/标题", "mode": "...", "basis": "...", "material": "...",
                    "figure": "...", "note": "...", "op": "upsert|remove"}],
  "promises":     [{"item": "工期", "value": "12 个月", "note": "...", "op": "upsert"}],
  "clarifications": [{"id": "CLAR-01", "verdict": "按建议", "note": ""}],
  "gates":        [{"gate": "G3", "verdict": "pass", "note": ""}]
}

树 op 字段：move{path,to_parent,index} / rename{path,text} / add{parent,index,text} /
remove{path}。path 为册内兄弟索引数组（[2,0]=第 3 个顶层节点的第 1 个子节点），
[] = 顶层；索引按 op 顺序在前序 op 之后的树状态上解释；节点深度 = len(path)。

安全声明：纯本地文件操作，无网络、无外部进程。
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from check_writing import LEGAL_FIGURES, LEGAL_MODES, MATERIAL_TOKEN, table_rows
from docx_toolkit import build_tree, parse_outline, sanitize_name

OUTLINE_RELPATH = "outline/tender-response-docs.md"
GUIDE_RELPATH = "body/写作指引.md"
PROMISE_RELPATH = "body/关键事实与承诺.md"
CLAR_RELPATH = "analysis/clarifications.md"
DECISIONS_RELPATH = "decisions.json"
BACKUP_KEEP = 3
MAX_LEVEL = 5
CLAR_ID = re.compile(r"^CLAR-\d{1,3}$")
GUIDE_HEADER = ["节", "模式", "依据", "素材", "图示", "缺口/备注"]
PROMISE_HEADER = ["事项", "值", "说明"]
GUIDE_FIELDMAP = {"节": "key", "模式": "mode", "依据": "basis",
                  "素材": "material", "图示": "figure", "缺口/备注": "note"}
PROMISE_FIELDMAP = {"事项": "item", "值": "value", "说明": "note"}
NEW_NOTE = "（评审台新增，待补归位理由与来源标注）"


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _cell(value) -> str:
    """写进 md 表格单元格的值：竖线与换行会破坏表格结构，替换掉。"""
    return str(value if value is not None else "").replace("|", "／").replace("\n", " ").strip()


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M")


# ---------- 基底校验与备份 ----------

def _check_base(name: str, relpath: str, work: Path, patch_base: dict, errors: list) -> None:
    path = work / relpath
    if not path.is_file():
        errors.append(f"{name}文件不存在：work/{relpath}")
        return
    given = (patch_base or {}).get(name)
    if not given:
        errors.append(f"修改包缺少 {name} 基底校验值（base.{name}）——请重新生成评审台页面后再提交")
    elif given != _sha256(path.read_text(encoding="utf-8")):
        errors.append(f"{name}文件在评审台打开后已被外部修改（基底不符）——"
                      f"请让助手重新生成评审台；本次修改未应用（旧基底的整写会覆盖他人改动）")


def _backup(work: Path, relpath: str, counter: list) -> str | None:
    path = work / relpath
    if not path.is_file():
        return None
    bdir = work / "review" / "backup"
    bdir.mkdir(parents=True, exist_ok=True)
    counter[0] += 1
    stem = relpath.replace("/", "_")
    dest = bdir / f"{counter[0]:04d}.{stem}.bak"
    shutil.copy2(path, dest)
    for old in sorted(bdir.glob(f"*.{stem}.bak"))[:-BACKUP_KEEP]:
        old.unlink(missing_ok=True)
    return f"review/backup/{dest.name}"


# ---------- 目录树：解析 / op / 回写 ----------

def _volumes(text: str) -> list[dict]:
    """parse_outline 薄封装：卷名 + 嵌套树（roots，children 就绪）。"""
    vols = parse_outline(text)
    for v in vols:
        v["roots"] = build_tree(
            [{"text": n["text"], "depth": n["depth"], "children": []} for n in v["nodes"]])
        for r in v["roots"]:
            _deep_children(r)
    return vols


def _deep_children(node: dict) -> None:
    node.setdefault("children", [])
    for c in node["children"]:
        _deep_children(c)


def _node_at(roots: list[dict], path: list) -> dict:
    cur = roots
    node = None
    for i in path:
        node = cur[int(i)]
        cur = node["children"]
    return node


def _siblings(roots: list[dict], path: list) -> list[dict]:
    return roots if len(path) <= 1 else _node_at(roots, path[:-1])["children"]


def _is_prefix(a: list, b: list) -> bool:
    return len(a) < len(b) and list(b[:len(a)]) == list(a)


def _flatten(roots: list[dict]) -> list[tuple[str, int]]:
    out: list[tuple[str, int]] = []

    def walk(nodes: list[dict], depth: int) -> None:
        for n in nodes:
            out.append((n["text"], depth))
            walk(n["children"], depth + 1)

    walk(roots, 1)
    return out


def apply_tree_ops(roots: list[dict], ops: list[dict], errors: list, vol_label: str,
                   touched: dict) -> None:
    """在内存树上顺序执行 op；出错只记 errors 不中断后续校验。
    touched 收集 {"renames": [(旧,新)], "added": [名], "removed": [名]}。"""
    for k, op in enumerate(ops):
        kind = op.get("op")
        tag = f"树操作[{k + 1}] {kind}"
        try:
            if kind == "rename":
                node = _node_at(roots, op["path"])
                old, new = node["text"], str(op.get("text", "")).strip()
                if not new:
                    errors.append(f"{tag}（{vol_label}）：新名称为空")
                elif new != old:
                    node["text"] = new
                    touched["renames"].append((old, new))
            elif kind == "add":
                parent_path = [int(i) for i in (op.get("parent") or [])]
                # 新节点的兄弟集 = 父节点的 children（[] = 顶层）
                sibs = _node_at(roots, parent_path)["children"] if parent_path else roots
                text = str(op.get("text", "")).strip()
                if not text:
                    errors.append(f"{tag}（{vol_label}）：新节点名称为空")
                    continue
                if len(parent_path) + 1 > MAX_LEVEL:
                    errors.append(f"{tag}（{vol_label}）：层级封顶 {MAX_LEVEL} 级，不能再加子级")
                    continue
                idx = max(0, min(int(op.get("index", len(sibs))), len(sibs)))
                sibs.insert(idx, {"text": text, "children": []})
                touched["added"].append(text)
            elif kind == "remove":
                path = [int(i) for i in op["path"]]
                touched["removed"].append(_node_at(roots, path)["text"])
                _siblings(roots, path).pop(path[-1])
            elif kind == "move":
                path = [int(i) for i in op["path"]]
                to_parent = [int(i) for i in (op.get("to_parent") or [])]
                if path == to_parent or _is_prefix(path, to_parent):
                    errors.append(f"{tag}（{vol_label}）：不能移动到自身或其子孙节点内")
                    continue
                if len(to_parent) + 1 > MAX_LEVEL:
                    errors.append(f"{tag}（{vol_label}）：层级封顶 {MAX_LEVEL} 级，不能移入该位置")
                    continue
                node = _node_at(roots, path)  # 先取节点（取出后 path 失效）
                _siblings(roots, path).pop(path[-1])
                new_sibs = _node_at(roots, to_parent)["children"] if to_parent else roots
                idx = max(0, min(int(op.get("index", 0)), len(new_sibs)))
                new_sibs.insert(idx, node)
                touched["moved"] += 1
            else:
                errors.append(f"{tag}（{vol_label}）：未知 op（支持 move/rename/add/remove）")
        except (IndexError, KeyError, TypeError, ValueError) as exc:
            errors.append(f"{tag}（{vol_label}）：路径或参数无法定位（{exc}）——"
                          f"修改包与当前目录不一致，请刷新评审台后重试")


# ---------- outline md 回写（三段外科替换） ----------

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


def _section_spans(lines: list[str]) -> dict[int, dict[str, tuple[int, int]]]:
    """每卷每段的行区间 [段头行, 段终点)（终点=下一段头/卷头/EOF，排他）。"""
    spans: dict[int, dict[str, tuple[int, int]]] = {}
    vi, sec = -1, None
    for i, ln in enumerate(lines):
        if ln.startswith("# 响应文件："):
            vi += 1
            sec = None
        elif ln.startswith("## ") and vi >= 0:
            sec = ln[3:].strip()
            spans.setdefault(vi, {})[sec] = (i, len(lines))
    for vmap in spans.values():
        ordered = sorted(vmap.items(), key=lambda kv: kv[1][0])
        for j, (sec, (start, _)) in enumerate(ordered):
            end = ordered[j + 1][1][0] if j + 1 < len(ordered) else len(lines)
            vmap[sec] = (start, end)
    return spans


def _migrate_ann_lines(mapping: dict, renames: dict, removed: set, which: str) -> list[str]:
    out = []
    for name, rest in mapping.items():
        if name in removed:
            continue
        out.append(f"- {renames.get(name, name)} :: {rest}")
    return out


def regen_outline(text: str, volumes: list[dict], flags: list) -> str:
    """按演算后的卷树重写每卷 `## 目录` 段；来源标注/目录说明定点迁移；
    新增节点在目录说明补占位行（来源标注按纪律不补，旗标提醒）。"""
    lines = text.split("\n")
    spans = _section_spans(lines)
    repls: list[tuple[int, int, list[str]]] = []  # (start, end, 新行) start 降序应用
    for vi, vol in enumerate(volumes):
        flat = _flatten(vol["roots"])
        repls.append((spans[vi]["目录"][0], spans[vi]["目录"][1],
                      ["  " * (d - 1) + "- " + t for t, d in flat]))
        ann = vol.get("_annotations", {"src:": {}, "desc:": {}})
        renames = dict(vol.get("_renames", []))
        removed = set(vol.get("_removed", []))
        if "来源标注" in spans.get(vi, {}):
            s, e = spans[vi]["来源标注"]
            repls.append((s, e, _migrate_ann_lines(ann.get("src:", {}), renames, removed, "src")))
        if "目录说明" in spans.get(vi, {}):
            s, e = spans[vi]["目录说明"]
            desc = _migrate_ann_lines(ann.get("desc:", {}), renames, removed, "desc")
            existing = {ln.split(" :: ")[0][2:] for ln in desc if " :: " in ln}
            for name in vol.get("_added", []):
                if name not in existing:
                    desc.append(f"- {name} :: 交付形态=正文编写 | 归位理由={NEW_NOTE} | 理由来源= | 概述=")
                    flags.append(f"{vol['name']}：新增节点「{name}」目录说明为占位行、"
                                 f"来源标注待补（无来源过渡章可不补）")
            repls.append((s, e, desc))
    repls.sort(key=lambda r: r[0], reverse=True)  # 行号从后往前替换，避免位移
    for s, e, new in repls:
        lines[s:e] = [lines[s]] + new  # 段头行保留
    return "\n".join(lines)


def structure_flags(volumes: list[dict], flags: list, report: dict) -> None:
    """结构 diff 汇总 + 封面/目录位置、重名叶子旗标。"""
    for vol in volumes:
        vn, flat = vol["name"], _flatten(vol["roots"])
        names = [t for t, _ in flat]
        if vol.get("_renames"):
            report["renames"] += [f"{vn}：{a} → {b}" for a, b in vol["_renames"]]
        if vol.get("_removed"):
            report["removed"] += [f"{vn}：{a}" for a in vol["_removed"]]
        if vol.get("_added"):
            report["added"] += [f"{vn}：{a}" for a in vol["_added"]]
        if vol.get("_renames") or vol.get("_added"):
            report["moved"] += 1  # 记一次结构事件（exact move 数不单独展示）
        if "封面" in names and names[0] != "封面":
            flags.append(f"{vn}：「封面」不在树首——合册按名识别封面页，建议保持首位")
        if "目录" in names and "封面" in names and names.index("目录") != 1:
            flags.append(f"{vn}：「目录」不在封面之后——合册按名识别目录页，建议保持第二位")
        seen: dict[str, int] = {}
        for t, _d in flat:
            seen[sanitize_name(t)] = seen.get(sanitize_name(t), 0) + 1
        for k, n in seen.items():
            if n > 1:
                flags.append(f"{vn}：{n} 个节点同名「{k}」——同册同名叶子会撞节文件名，请改名其一")


# ---------- 表格合并（写作指引 / 承诺清单） ----------

def merge_table(work: Path, relpath: str, header: list[str], fieldmap: dict[str, str],
                entries: list[dict], key_json: str, label: str,
                errors: list, report: dict, counter: list) -> None:
    path = work / relpath
    text = path.read_text(encoding="utf-8")
    rows = table_rows(text)
    if not rows or [c.strip() for c in rows[0][:len(header)]] != header:
        errors.append(f"{label}表头不符或表格缺失（期望 {'|'.join(header)}）：work/{relpath}")
        return
    data = [(r + [""] * len(header))[:len(header)] for r in rows[1:]]
    index: dict[str, int] = {}
    for i, r in enumerate(data):
        index.setdefault(r[0].strip(), i)

    for e in entries:
        key = _cell(e.get(key_json, ""))
        if not key:
            errors.append(f"{label}条目缺少 {key_json} 字段")
            continue
        if e.get("op", "upsert") == "remove":
            if key in index:
                report[f"{label}_removed"].append(key)
                data[index[key]] = None
            else:
                errors.append(f"{label}「{key}」：行不存在，无法删除")
            continue
        if key in index:
            row = data[index[key]]
            for j, h in enumerate(header):
                val = e.get(fieldmap[h])
                if val is not None:
                    row[j] = _cell(val)
            report[f"{label}_updated"].append(key)
        else:
            row = [_cell(e.get(fieldmap[h], "—")) for h in header]
            row[0] = key
            data.append(row)
            report[f"{label}_added"].append(key)
    data = [r for r in data if r is not None]
    report["written"].append(_save_table(work, relpath, text, header, data, counter))


def _save_table(work: Path, relpath: str, text: str, header: list[str],
                data: list[list[str]], counter: list) -> dict:
    """非表格行原样保留，表格块（首行 | 到末行 |）整体重写。"""
    lines = text.split("\n")
    tbl = [i for i, ln in enumerate(lines) if ln.strip().startswith("|")]
    start, end = (tbl[0], tbl[-1] + 1) if tbl else (len(lines), len(lines))
    body = ["| " + " | ".join(header) + " |",
            "| " + " | ".join(["---"] * len(header)) + " |"]
    body += ["| " + " | ".join(r) + " |" for r in data]
    bak = _backup(work, relpath, counter)
    (work / relpath).write_text("\n".join(lines[:start] + body + lines[end:]), encoding="utf-8")
    return {"file": relpath, "backup": bak}


def migrate_guide_keys(work: Path, key_renames: dict[str, str], report: dict,
                       counter: list, flags: list) -> None:
    """树节点改名后，写作指引表同名行键自动跟随（与 CLIENT 目录改名行为一致）。"""
    if not key_renames:
        return
    path = work / GUIDE_RELPATH
    if not path.is_file():
        return
    text = path.read_text(encoding="utf-8")
    rows = table_rows(text)
    if not rows:
        return
    data = [(r + [""] * len(GUIDE_HEADER))[:len(GUIDE_HEADER)] for r in rows[1:]]
    migrated: list[str] = []
    for r in data:
        k = r[0].strip()
        if k in key_renames:
            migrated.append(f"{k} → {key_renames[k]}")
            r[0] = key_renames[k]
    keys = [r[0].strip() for r in data]
    for new in key_renames.values():
        if keys.count(new) > 1:
            flags.append(f"写作指引出现重复行「{new}」——改名迁移与新增撞行，请删其一")
    if migrated:
        report["written"].append(_save_table(work, GUIDE_RELPATH, text, GUIDE_HEADER,
                                             data, counter))
        report["guide_key_migrated"] = migrated


def validate_guide_entry(e: dict, expected_keys: set[str]) -> str | None:
    key = str(e.get("key", "")).strip()
    if not key:
        return "缺少节名"
    if key not in expected_keys:
        near = [k for k in sorted(expected_keys) if key[-6:] in k][:3]
        return ("节名不在目录树中（须与目录逐字一致，多册写「册名/标题」）"
                + (f"；相近：{'、'.join(near)}" if near else ""))
    mode = str(e.get("mode", "")).strip()
    if mode and mode != "—" and mode not in LEGAL_MODES and not all(
            m in {"素材修订", "格式跟随", "推理撰写"} for m in mode.split("+")):
        return f"模式非法：{mode}（合法：素材修订/格式跟随/推理撰写，可 + 组合，或 —）"
    material = str(e.get("material", "")).strip()
    if material and material not in ("—", "【缺】"):
        for tok in [t.strip() for t in material.split("，") if t.strip()]:
            if not MATERIAL_TOKEN.match(tok):
                return f"素材记法非法：{tok}（应为「文件.docx L120-L260」，多个用中文逗号分隔，或 —/【缺】）"
    figure = str(e.get("figure", "")).strip()
    if figure and figure != "—":
        for tok in [t.strip() for t in figure.split("、") if t.strip()]:
            if tok.split(":", 1)[0] not in LEGAL_FIGURES:
                return f"图示类型非法：{tok}（类型词限：{'、'.join(sorted(LEGAL_FIGURES))}，记法「类型:主题」）"
    return None


# ---------- 澄清裁决 / 确认门（追加性写入） ----------

def merge_decisions(work: Path, patch: dict, report: dict) -> None:
    dpath = work / DECISIONS_RELPATH
    data: dict = {}
    if dpath.is_file():
        try:
            data = json.loads(dpath.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    data.setdefault("gates", [])
    data.setdefault("clarifications", {})
    for g in patch.get("gates", []):
        gate = str(g.get("gate", "")).strip()
        if gate:
            data["gates"].append({"gate": gate, "verdict": str(g.get("verdict", "pass")),
                                  "note": str(g.get("note", "")), "at": _now()})
            report["gates"].append(f"{gate} · {g.get('verdict', 'pass')}")
    for c in patch.get("clarifications", []):
        cid = str(c.get("id", "")).strip().upper()
        if not CLAR_ID.match(cid):
            report["flags"].append(f"澄清编号不合规（期望 CLAR-xx）：{cid!r}，已跳过")
            continue
        data["clarifications"][cid] = {"verdict": str(c.get("verdict", "")),
                                       "note": str(c.get("note", "")), "at": _now()}
        report["clarifications"].append(f"{cid} · {c.get('verdict', '')}")
    dpath.parent.mkdir(parents=True, exist_ok=True)
    dpath.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def append_clar_verdicts(work: Path, patch: dict, errors: list, report: dict,
                         counter: list) -> None:
    """裁决写进 clarifications.md 的 `## 裁决记录` 段（同 id 定点替换，无则追加段）。"""
    items = patch.get("clarifications", [])
    if not items:
        return
    relpath = CLAR_RELPATH
    path = work / relpath
    if not path.is_file():
        errors.append(f"澄清清单不存在：work/{relpath}（先完成要点提取再裁决）")
        return
    lines = path.read_text(encoding="utf-8").split("\n")
    sec_start = next((i for i, ln in enumerate(lines) if ln.strip() == "## 裁决记录"), None)
    if sec_start is None:
        lines += ["", "## 裁决记录"]
        sec_start = len(lines) - 1
    sec_end = len(lines)
    for j in range(sec_start + 1, len(lines)):
        if lines[j].startswith("#"):
            sec_end = j
            break
    existing: dict[str, int] = {}
    for j in range(sec_start + 1, sec_end):
        m = re.match(r"-\s*(CLAR-\d{1,3})\s*::", lines[j], re.I)
        if m:
            existing.setdefault(m.group(1).upper(), j)
    for c in items:
        cid = str(c.get("id", "")).strip().upper()
        if not CLAR_ID.match(cid):
            continue
        line = (f"- {cid} :: 裁决={_cell(c.get('verdict', '按建议'))}"
                f" | 备注={_cell(c.get('note', ''))} | 时间={_now()}")
        if cid in existing:
            lines[existing[cid]] = line
        else:
            lines.insert(sec_end, line)
            sec_end += 1
    bak = _backup(work, relpath, counter)
    path.write_text("\n".join(lines), encoding="utf-8")
    report["written"].append({"file": relpath, "backup": bak})


# ---------- 主入口 ----------

def apply_patch(work: Path, patch: dict) -> dict:
    """应用修改包；树/指引/承诺三路全量校验通过才落盘（原子），追加性写入随后。"""
    errors: list[str] = []
    report: dict = {"written": [], "added": [], "removed": [], "renames": [], "moved": 0,
                    "flags": [], "gates": [], "clarifications": [],
                    "写作指引_updated": [], "写作指引_added": [], "写作指引_removed": [],
                    "承诺清单_updated": [], "承诺清单_added": []}
    counter = [0]
    work = Path(work)
    tree_patch = patch.get("tree", [])
    guide_patch = patch.get("guide", [])
    promise_patch = patch.get("promises", [])

    # 1) 基底校验（破坏性写入前）
    for name, relpath, wanted in (("outline", OUTLINE_RELPATH, bool(tree_patch)),
                                  ("guide", GUIDE_RELPATH, bool(guide_patch)),
                                  ("promise", PROMISE_RELPATH, bool(promise_patch))):
        if wanted:
            _check_base(name, relpath, work, patch.get("base", {}), errors)
    if errors:
        return {"ok": False, "stage": "base", "errors": errors, **report}

    # 2) 树操作内存演算（不改盘）
    volumes: list[dict] = []
    staged_text = None
    if tree_patch:
        text = (work / OUTLINE_RELPATH).read_text(encoding="utf-8")
        volumes = _volumes(text)
        vol_index = {v["name"]: i for i, v in enumerate(volumes)}
        ann_all = _parse_annotations(text.split("\n"))
        for entry in tree_patch:
            vn = str(entry.get("vol", ""))
            vi = vol_index.get(vn)
            if vi is None:
                errors.append(f"树操作指向不存在的响应文件「{vn}」（现有：{'、'.join(vol_index) or '—'}）")
                continue
            vol = volumes[vi]
            vol["_annotations"] = {"src:": ann_all[vi]["src:"], "desc:": ann_all[vi]["desc:"]}
            touched = {"renames": [], "added": [], "removed": [], "moved": 0}
            apply_tree_ops(vol["roots"], entry.get("ops", []), errors, vn, touched)
            vol.update({"_renames": touched["renames"], "_added": touched["added"],
                        "_removed": touched["removed"], "_moved": touched["moved"]})
        if errors:
            return {"ok": False, "stage": "tree", "errors": errors, **report}
        staged_text = regen_outline(text, volumes, report["flags"])
        structure_flags(volumes, report["flags"], report)

    # 3) 指引校验（对照演算后的树；无树操作时读当前目录）
    if guide_patch:
        if not volumes:
            cur = (work / OUTLINE_RELPATH).read_text(encoding="utf-8")
            volumes = _volumes(cur)
        expected: set[str] = set()
        for vol in volumes:
            pre = f"{vol['name']}/" if len(volumes) > 1 else ""
            expected |= {pre + t for t, _d in _flatten(vol["roots"])}
        for e in guide_patch:
            bad = validate_guide_entry(e, expected)
            if bad:
                errors.append(f"写作指引「{e.get('key', '')}」：{bad}")
        if errors:
            return {"ok": False, "stage": "guide", "errors": errors, **report}

    # 4) 落盘：树 → 指引键迁移 → 指引 → 承诺（原子段）
    try:
        if tree_patch:
            bak = _backup(work, OUTLINE_RELPATH, counter)
            (work / OUTLINE_RELPATH).write_text(staged_text, encoding="utf-8")
            report["written"].append({"file": OUTLINE_RELPATH, "backup": bak})
            multi = len(volumes) > 1
            key_renames = {}
            for vol in volumes:
                for old, new in vol.get("_renames", []):
                    key_renames[f"{vol['name']}/{old}" if multi else old] = \
                        f"{vol['name']}/{new}" if multi else new
            migrate_guide_keys(work, key_renames, report, counter, report["flags"])
        if guide_patch:
            merge_table(work, GUIDE_RELPATH, GUIDE_HEADER, GUIDE_FIELDMAP, guide_patch,
                        "key", "写作指引", errors, report, counter)
        if promise_patch:
            merge_table(work, PROMISE_RELPATH, PROMISE_HEADER, PROMISE_FIELDMAP, promise_patch,
                        "item", "承诺清单", errors, report, counter)
        if errors:  # 表格合并中途失败则到此为止（已写部分有备份可回退）
            return {"ok": False, "stage": "write", "errors": errors, **report}
        # 5) 追加性写入：澄清裁决 + 确认门
        append_clar_verdicts(work, patch, errors, report, counter)
        merge_decisions(work, patch, report)
    except OSError as exc:
        errors.append(f"落盘失败：{exc}")
        return {"ok": False, "stage": "write", "errors": errors, **report}

    # 6) 下游影响旗标
    if report["removed"]:
        report["flags"].append("有节点被删除——其写作指引行与节文件不会自动清理；"
                               "目录新叶子也还没有指引行，建议让助手增量补跑写作指引")
    if report["added"] or report["renames"]:
        report["flags"].append("目录结构有实质变化（新增/改名）——正文阶段以新目录为准；"
                               "已建节的旧名文件会在合册时按「目录外文件」点名")
    if any(g.startswith("G3") for g in report["gates"]) and (tree_patch or guide_patch):
        report["flags"].append("本次改动了目录或指引——G3/G4 已确认过的结论建议重新过一遍门")
    return {"ok": True, "errors": [], **report}


def main() -> None:
    import argparse
    ap = argparse.ArgumentParser(description="应用评审台修改包（落盘 md + 备份 + diff 报告）")
    ap.add_argument("--patch", required=True, help="修改包 JSON 路径")
    ap.add_argument("--work", default="work", help="工作目录（默认 work）")
    args = ap.parse_args()
    patch = json.loads(Path(args.patch).read_text(encoding="utf-8"))
    result = apply_patch(Path(args.work), patch)
    if not result["ok"]:
        print(f"[应用拒绝] {len(result['errors'])} 处：")
        for e in result["errors"]:
            print(f"- {e}")
        sys.exit(2)
    print("[应用成功]")
    for w in result["written"]:
        print(f"- 已写入 work/{w['file']}" + (f"（备份 {w['backup']}）" if w["backup"] else ""))
    for group, label in (("added", "新增节点"), ("removed", "删除节点"), ("renames", "改名"),
                         ("gates", "确认门"), ("clarifications", "澄清裁决"),
                         ("写作指引_updated", "指引更新"), ("写作指引_added", "指引新增"),
                         ("写作指引_removed", "指引删行"), ("承诺清单_updated", "承诺更新"),
                         ("guide_key_migrated", "指引改键"),
                         ("承诺清单_added", "承诺新增")):
        for item in result.get(group, []):
            print(f"- {label}：{item}")
    if result["flags"]:
        print("下游注意：")
        for f in result["flags"]:
            print(f"- {f}")
    print("应用后请重跑：python3 scripts/check_writing.py --guide")


if __name__ == "__main__":
    main()
