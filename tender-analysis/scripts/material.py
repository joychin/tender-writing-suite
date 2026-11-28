#!/usr/bin/env python3
r"""material.py — 公司资料库登记（素材复用的确定性通道）。

技能无状态，状态放用户目录：默认 ~/.lingyan/tender-materials/（环境变量
TENDER_MATERIAL_HOME 可改，各宿主共用同一份）。资料按「主体」分目录登记——不同主体
（母公司/子公司/联合体成员）的标书各用各的资料；检索缺省限定本次主体，命中其他主体的
素材单独列出点名确认（引用性质还是误命中，归用户拍板），不静默丢弃也不静默使用。

入库（add）：原件一个字节不动——strip_metadata 出脱敏副本存库（库内文件永远不带作者
等元数据），parse_document 解析库内副本生成带行号的章节树与解析产物；行号区间与库内
文件永远一致，copy-from 直接以库内文件为源（--parse 指向库内解析目录，素材新鲜度闸
照常生效）。索引（索引.json / 索引.md）由本脚本确定性生成：章节树、旧名候选（项目
名称/招标编号/采购人等标签行的值与编号模式，供 check_writing 残留检查自动载入）、
体检摘要。检索是机械活，语义匹配（哪节配哪段素材）归模型。

命令：
  add <文件或目录> --entity <主体>     入库（docx/doc：脱敏+解析+索引；pdf：脱敏记图源）
  list [--entity <主体>]               登记清单（空库给入库指路）
  show <文件名> [--entity <主体>]       一份素材的章节树（带行号）与旧名候选
  search <关键词...> [--entity <主体>] [--limit N]   检索章节，输出素材列记法
  forget <文件名> [--entity <主体>]     出库（删库内副本与解析产物，原件不动）

检索优先级与素材改写纪律见写作环节的 guide-format.md 与 section-writing.md（references/）。
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

_SCRIPTS_DIR = Path(__file__).resolve().parent
_STAGE_ROOT = _SCRIPTS_DIR.parent
_PARSE = _SCRIPTS_DIR / "parse_document.py"
_SENTINEL = {"tender-sanitize": "strip_metadata.py"}
_INDEX_NAME = "索引.json"
_HOME_ENV = "TENDER_MATERIAL_HOME"
_DEFAULT_HOME = Path.home() / ".lingyan" / "tender-materials"
_SUPPORTED = (".docx", ".doc", ".pdf")
_STORE_EXT = {".docx": ".docx", ".doc": ".docx", ".pdf": ".pdf"}
_FIG_NOTE = "图源：附件节贴图用（render-page 转 PNG 后 image 插入），不进素材列"

# 旧名候选：标签行取值（高精度）+ 编号模式（字母开头带连字符的编号，如 YLCG-2026-GK-0387）
_LABELS = ("项目名称", "项目编号", "招标编号", "采购编号", "标段编号",
           "采购人", "采购人名称", "招标人", "招标方", "招标代理", "招标代理机构")
_LABEL_RE = re.compile(r"^\s*[*#>\-\s|]*(" + "|".join(_LABELS) + r")\s*[：:]\s*(.+?)\s*$")
_CODE_RE = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]{2,10}(?:[-–－][A-Za-z0-9]{1,10}){1,4})(?![A-Za-z0-9])")
_VALUE_MAX = 40
_OLDNAMES_CAP = 20


def die(msg: str, code: int = 2):
    first, *rest = str(msg).split("\n")
    print(f"[失败] {first}")
    for ln in rest:  # 多行消息续行缩进对齐，不散落成正文
        print(f"        {ln}")
    sys.exit(code)


def _home() -> Path:
    v = os.environ.get(_HOME_ENV, "").strip()
    return Path(v) if v else _DEFAULT_HOME


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _entity_key(entity: str) -> str:
    key = re.sub(r'[\\/:*?"<>|]+', "", entity).strip()
    key = re.sub(r"\s+", " ", key)
    if not key:
        die("主体名清不出可用目录名——换一个主体名（去掉 / \\ : * ? \" < > | 等符号）")
    return key[:40].strip()


def _load_index(home: Path) -> dict:
    p = home / _INDEX_NAME
    if not p.is_file():
        return {"materials": [], "updated_at": ""}
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(d, dict) and isinstance(d.get("materials"), list):
            return d
    except Exception:  # noqa: BLE001
        pass
    print(f"  ⚠️ 索引损坏或不可读（{p}）——按空库重建，原有登记条目丢失，可用 add 重新入库")
    return {"materials": [], "updated_at": ""}


def _save_index(home: Path, idx: dict) -> None:
    home.mkdir(parents=True, exist_ok=True)
    idx["updated_at"] = time.strftime("%Y-%m-%d %H:%M")
    (home / _INDEX_NAME).write_text(json.dumps(idx, ensure_ascii=False, indent=1), encoding="utf-8")
    _write_index_md(home, idx)


def _write_index_md(home: Path, idx: dict) -> None:
    lines = [f"# 公司资料库索引", f"", f"更新：{idx.get('updated_at', '')}", ""]
    by_entity: dict[str, list] = {}
    for m in idx["materials"]:
        by_entity.setdefault(m.get("entity", "?"), []).append(m)
    for ek in sorted(by_entity):
        disp = by_entity[ek][0].get("entity_display") or ek
        lines.append(f"## 主体：{disp}（{len(by_entity[ek])} 份）")
        lines.append("")
        for m in by_entity[ek]:
            extra = f"，{_FIG_NOTE}" if m.get("kind") == "pdf" else f"，{m.get('chapters', 0)} 章，旧名候选 {len(m.get('old_names') or [])} 条"
            lines.append(f"- {m['file']} — {m.get('kind', 'docx')}，入库 {m.get('added_at', '')}{extra}")
        lines.append("")
    (home / "索引.md").write_text("\n".join(lines), encoding="utf-8")


def _collect(paths: list[str]) -> tuple[list[Path], list[str]]:
    """收集支持的文件；目录一层枚举，锁文件与不支持后缀跳过并点名。"""
    out: list[Path] = []
    skipped: list[str] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            for f in sorted(p.iterdir()):
                if not f.is_file():
                    continue
                if f.name.startswith(("~$", ".~")):
                    skipped.append(f"{f.name} —— Word/WPS 占用锁文件")
                    continue
                if f.suffix.lower() not in _SUPPORTED:
                    skipped.append(f"{f.name} —— 暂只支持 .docx / .doc / .pdf")
                    continue
                out.append(f)
        elif p.is_file():
            if p.suffix.lower() not in _SUPPORTED:
                die(f"暂只支持 .docx / .doc / .pdf：{p.name}")
            out.append(p)
        else:
            die(f"路径不存在：{p}")
    return out, skipped


def _clean_copy(src: Path, edir: Path) -> tuple[Path, str]:
    """strip_metadata 出脱敏副本（原件不动），落到库内标准名。返回 (库内文件, 体检摘要)。"""
    ext = _STORE_EXT[src.suffix.lower()]
    edir.mkdir(parents=True, exist_ok=True)
    findings = 0
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run(
            [sys.executable, "-B", str(_sibling_scripts("tender-sanitize") / "strip_metadata.py"),
             str(src), "-o", td, "--json"],
            capture_output=True, text=True, errors="replace",
            env=dict(os.environ, PYTHONDONTWRITEBYTECODE="1"))
        if r.returncode != 0:
            die("脱敏入库失败（原件未动，资料库未写入）：\n" + (r.stdout + r.stderr).strip()[:1600])
        try:
            rep = json.loads(r.stdout)
            findings = int((rep.get("summary") or {}).get("findings") or 0)
        except Exception:  # noqa: BLE001
            pass
        produced = Path(td) / f"{src.stem}_cleaned{ext}"
        if not produced.is_file():
            die(f"脱敏副本没有生成（strip_metadata 退出码 0 但产物缺失）：期望 {produced.name}")
        stored = edir / f"{src.stem}{ext}"
        if stored.exists():
            stored.unlink()
        shutil.move(str(produced), str(stored))
    return stored, f"清理 {findings} 项"


def _sibling_scripts(name: str) -> Path:
    if (_SCRIPTS_DIR / _SENTINEL[name]).is_file():
        return _SCRIPTS_DIR
    for cand in (_STAGE_ROOT.parent / name, _STAGE_ROOT / "stages" / name):
        scripts = cand / "scripts"
        if scripts.is_dir():
            return scripts
    raise SystemExit(f"缺少兄弟环节 {name}（单包 / 平级 / stages 三形态均未找到）：{_STAGE_ROOT.parent}")


def _parse_stored(stored: Path, edir: Path) -> Path:
    """解析库内副本：行号区间与库内文件永远一致，copy-from 直接用。返回解析目录。"""
    r = subprocess.run(
        [sys.executable, "-B", str(_PARSE), str(stored), "-o", str(edir / "解析")],
        capture_output=True, text=True, errors="replace",
        env=dict(os.environ, PYTHONDONTWRITEBYTECODE="1"))
    if r.returncode != 0:
        die("解析失败（脱敏副本已存库，解析产物未生成）：\n" + (r.stdout + r.stderr).strip()[:1600])
    return edir / "解析" / stored.name


def _old_names(md_path: Path) -> list[str]:
    try:
        lines = md_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:  # noqa: BLE001
        return []
    names: list[str] = []
    for ln in lines:
        t = ln.strip().strip("*").strip()
        m = _LABEL_RE.match(ln) or _LABEL_RE.match(t)
        if m:
            v = m.group(2).strip().strip("*").strip().strip("|").strip()
        elif t.startswith("|"):
            cells = [c.strip().strip("*").strip() for c in t.strip("|").split("|")]
            v = ""
            for i, c in enumerate(cells):
                if c in _LABELS and i + 1 < len(cells) and cells[i + 1]:
                    v = cells[i + 1]
                    break
        else:
            v = ""
        if v and len(v) <= _VALUE_MAX and not re.fullmatch(r"[\d\s年月日.:/-]+", v) and v not in names:
            names.append(v)
        for c in _CODE_RE.findall(ln):
            if len(c) >= 8 and c not in names:
                names.append(c)
    return names[:_OLDNAMES_CAP]


def _flatten(nodes: list, out: list, depth: int = 0) -> None:
    for n in nodes:
        try:
            out.append((depth, str(n.get("标题", "")), int(n.get("start_line", 0)), int(n.get("end_line", 0))))
        except Exception:  # noqa: BLE001
            continue
        _flatten(n.get("children") or [], out, depth + 1)


def cmd_add(args) -> None:
    if not args.entity.strip():
        die("缺 --entity：资料按主体分目录登记（不同主体的标书各用各的资料，混用是废标面）——\n"
            "先向用户确认这是哪家主体的资料，再带 --entity 主体名 重跑")
    home = _home()
    ekey = _entity_key(args.entity)
    edir = home / ekey
    files, skipped = _collect(args.paths)
    if not files:
        die("没有可入库的文件（只收 .docx / .doc / .pdf）" + ("；跳过：" + "；".join(s for s in skipped) if skipped else ""))
    idx = _load_index(home)
    added = 0
    for src in files:
        ext = _STORE_EXT[src.suffix.lower()]
        fname = src.stem + ext
        existing = [m for m in idx["materials"] if m.get("entity") == ekey and m.get("file") == fname]
        if existing:
            m = existing[0]
            if m.get("source_sha256") == _sha256(src):
                print(f"[已登记] {fname}（主体 {ekey}）——内容未变，沿用已有登记（重新入库请先 forget）")
                continue
            die(f"同名不同物：主体「{ekey}」已有 {fname}（入库于 {m.get('added_at', '')}，内容指纹不同）——\n"
                f"先 material.py forget {fname} --entity {ekey} 再重新入库，或换个文件名")
        stored, clean_note = _clean_copy(src, edir)
        entry = {
            "entity": ekey, "entity_display": args.entity.strip(),
            "file": fname, "kind": "pdf" if ext == ".pdf" else "docx",
            "sha256": _sha256(stored), "source_sha256": _sha256(src),
            "cleaned": True, "clean_summary": clean_note,
            "parse_dir": "", "chapters": 0, "old_names": [], "note": "", "added_at": time.strftime("%Y-%m-%d"),
        }
        if ext == ".pdf":
            entry["note"] = _FIG_NOTE
            print(f"[已入库] {fname} → 主体「{ekey}」（{clean_note}，原件未动；{_FIG_NOTE}）")
        else:
            parse_dir = _parse_stored(stored, edir)
            entry["parse_dir"] = str(parse_dir.relative_to(home))
            ol = json.loads((parse_dir / (stored.stem + ".outline.json")).read_text(encoding="utf-8"))
            flat: list = []
            _flatten(ol, flat)
            entry["chapters"] = len(flat)
            entry["old_names"] = _old_names(parse_dir / (stored.stem + ".md"))
            print(f"[已入库] {fname} → 主体「{ekey}」（{clean_note}，原件未动）")
            print(f"  章节 {len(flat)} 章 / 旧名候选 {len(entry['old_names'])} 条")
            print(f"  解析目录：{parse_dir}")
            print(f"  保真拷贝：docx_toolkit copy-from {stored} <节.docx> --lines L起-L止 --parse {parse_dir}")
        idx["materials"].append(entry)
        _save_index(home, idx)
        added += 1
    for s in skipped:
        print(f"[跳过] {s}")
    print(f"[资料库] {home}（本次入库 {added} 份，共 {len(idx['materials'])} 份登记）")


def _match(entries: list, name: str, entity: str) -> list:
    out = [m for m in entries if m.get("file") == name]
    if entity:
        out = [m for m in out if m.get("entity") == _entity_key(entity)]
    return out


def cmd_list(args) -> None:
    home = _home()
    idx = _load_index(home)
    mats = idx["materials"]
    if args.entity:
        mats = [m for m in mats if m.get("entity") == _entity_key(args.entity)]
    if not mats:
        print(f"[资料库空] {home} 还没有登记"
              + ("（该主体名下暂无）" if args.entity else "")
              + "——用户丢来文件/目录即可入库：material.py add <路径> --entity <主体名>")
        return
    print(f"[资料库] {home}（{len(mats)} 份登记）")
    by_entity: dict[str, list] = {}
    for m in mats:
        by_entity.setdefault(m.get("entity", "?"), []).append(m)
    for ek in sorted(by_entity):
        disp = by_entity[ek][0].get("entity_display") or ek
        print(f"{disp}（{len(by_entity[ek])} 份）")
        for m in by_entity[ek]:
            extra = "图源" if m.get("kind") == "pdf" else f"{m.get('chapters', 0)} 章/旧名 {len(m.get('old_names') or [])}"
            print(f"  {m['file']}  {m.get('kind', 'docx')}  {extra}  入库 {m.get('added_at', '')}")


def cmd_show(args) -> None:
    home = _home()
    idx = _load_index(home)
    hits = _match(idx["materials"], args.name, args.entity)
    if not hits:
        die(f"没有「{args.name}」的登记——material.py list 先看库里有什么")
    for m in hits:
        disp = m.get("entity_display") or m.get("entity")
        print(f"[素材] {m['file']}（主体 {disp}，{m.get('kind', 'docx')}，入库 {m.get('added_at', '')}，{m.get('clean_summary', '')}）")
        stored = home / m.get("entity", "") / m["file"]
        print(f"  库内文件：{stored}")
        if m.get("kind") == "pdf":
            print(f"  {_FIG_NOTE}")
            continue
        pdir = home / (m.get("parse_dir") or "")
        if not pdir.is_dir():
            print("  [注意] 解析目录缺失——重新入库（先 forget）后才能保真拷贝")
            continue
        print(f"  解析目录：{pdir}")
        try:
            ol = json.loads((pdir / (Path(m['file']).stem + ".outline.json")).read_text(encoding="utf-8"))
            flat: list = []
            _flatten(ol, flat)
        except Exception:  # noqa: BLE001
            flat = []
        if flat:
            print("  章节（行号=解析 md 行号，素材列与 copy-from 同口径）：")
            for depth, title, s, e in flat:
                print(f"    {'  ' * depth}{title}  L{s}-L{e}")
        else:
            print("  章节树为空（全文无标题结构）——素材列可记整份文件区间，或按 md 行号手动定位")
        olds = m.get("old_names") or []
        if olds:
            print("  旧名候选（check_writing 残留检查自动载入）：" + "、".join(olds))
        print(f"  保真拷贝：docx_toolkit copy-from {stored} <节.docx> --lines L起-L止 --parse {pdir}")


def cmd_search(args) -> None:
    home = _home()
    idx = _load_index(home)
    mats = idx["materials"]
    kws = [k for k in args.keywords if k.strip()]
    if not kws:
        die("缺关键词：material.py search <关键词...>（可多个，任一命中即算）")
    entity = args.entity.strip()
    mine = [m for m in mats if (not entity or m.get("entity") == _entity_key(entity))]
    others = [m for m in mats if entity and m.get("entity") != _entity_key(entity)]
    scope = f"主体：{entity}" if entity else "主体：未指定（全部主体）"
    print(f"[检索] 关键词：{'、'.join(kws)}  {scope}（库内 {len(mats)} 份登记，检索 {len([m for m in mine if m.get('kind') == 'docx'])} 份 docx）")
    if not entity:
        print("[注意] 未指定主体——先与用户确认本次投标主体再选用登记库素材（不同主体资料混用是废标面）；下面按主体分组列出")
    any_hit = False
    for group, is_other in ((mine, False), (others, True)):
        hits: list[tuple] = []
        for m in group:
            if m.get("kind") != "docx":
                continue
            pdir = home / (m.get("parse_dir") or "")
            ol_p = pdir / (Path(m["file"]).stem + ".outline.json")
            md_p = pdir / (Path(m["file"]).stem + ".md")
            if not (ol_p.is_file() and md_p.is_file()):
                continue
            try:
                lines = md_p.read_text(encoding="utf-8", errors="replace").splitlines()
                ol = json.loads(ol_p.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                continue
            flat: list = []
            _flatten(ol, flat)
            low = [ln.lower() for ln in lines]
            for depth, title, s, e in flat:
                n = sum(low[i - 1].count(kw.lower()) for i in range(max(1, s), min(e, len(low)) + 1) for kw in kws)
                if n > 0:
                    hits.append((n, m, depth, title, s, e))
        hits.sort(key=lambda x: -x[0])
        kept: list[tuple] = []
        for h in hits:
            _, _, _, _, s, e = h
            if any(ks <= s and ke >= e and (ks, ke) != (s, e) for _, _, _, _, ks, ke in kept):
                continue  # 父章整个区间已被子节覆盖时报子节，不重复报父章
            kept.append(h)
            if len(kept) >= args.limit:
                break
        if not kept:
            continue
        any_hit = True
        if is_other:
            print("[他主体命中] —— 命中了其他主体的资料：先向用户确认是引用性质（集团背书/第三方业绩）还是误命中，再决定用不用")
        for n, m, depth, title, s, e in kept:
            disp = m.get("entity_display") or m.get("entity")
            print(f"  [{disp}] {m['file']} L{s}-L{e}（{'  ' * depth}{title}；命中 {n} 处）")
        if not is_other:
            print("[用法] 素材列照抄「文件名 L起-L止」；保真拷贝用 show 回显里的 copy-from 用法行（--parse 指库内解析目录）。")
    if not any_hit:
        print(f"[未命中] 关键词 {'、'.join(kws)} 在登记资料里没有命中——检索留痕照记；可向用户索取素材文件，或查宿主知识库/资料库通道")


def cmd_forget(args) -> None:
    home = _home()
    idx = _load_index(home)
    hits = _match(idx["materials"], args.name, args.entity)
    if not hits:
        die(f"没有「{args.name}」的登记——material.py list 先看库里有什么")
    if len(hits) > 1:
        lines = "、".join(f"主体 {m.get('entity')}" for m in hits)
        die(f"「{args.name}」登记在多个主体名下（{lines}）——带 --entity 指定要出库哪个")
    m = hits[0]
    idx["materials"].remove(m)
    _save_index(home, idx)
    stored = home / m.get("entity", "") / m["file"]
    if stored.exists():
        stored.unlink()
    pdir = home / (m.get("parse_dir") or "")
    if m.get("parse_dir") and pdir.is_dir():
        shutil.rmtree(pdir, ignore_errors=True)
    print(f"[已出库] {m['file']}（主体 {m.get('entity_display') or m.get('entity')}）——库内副本与解析产物已删，原件未动")


def main() -> None:
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(errors="replace")
        except Exception:  # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(
        prog="material.py",
        description="公司资料库登记（素材复用的确定性通道）：按主体分目录，脱敏副本+解析+索引；检索输出素材列记法")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("add", help="入库：脱敏副本 + 解析 + 索引（docx/doc/pdf）")
    p.add_argument("paths", nargs="+", help="文件或目录（目录一层枚举）")
    p.add_argument("--entity", default="", help="主体名（必填：资料按主体分目录登记，先向用户确认是哪家主体）")
    p.set_defaults(func=cmd_add)

    p = sub.add_parser("list", help="登记清单（空库给入库指路）")
    p.add_argument("--entity", default="", help="只看这个主体")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("show", help="一份素材的章节树（带行号）与旧名候选")
    p.add_argument("name", help="素材文件名")
    p.add_argument("--entity", default="", help="限定主体（同名素材登记在多个主体名下时用来指定）")
    p.set_defaults(func=cmd_show)

    p = sub.add_parser("search", help="检索章节，输出素材列记法（其他主体命中单独点名）")
    p.add_argument("keywords", nargs="+", help="关键词（可多个，任一命中即算）")
    p.add_argument("--entity", default="", help="本次投标主体（缺省列出全部主体并提醒先确认主体）")
    p.add_argument("--limit", type=int, default=12, help="每个分组最多几条（缺省 12）")
    p.set_defaults(func=cmd_search)

    p = sub.add_parser("forget", help="出库：删库内副本与解析产物（原件不动）")
    p.add_argument("name", help="素材文件名")
    p.add_argument("--entity", default="", help="限定主体（同名素材在多个主体名下时必填）")
    p.set_defaults(func=cmd_forget)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
