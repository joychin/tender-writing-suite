#!/usr/bin/env python3
"""check_writing.py — 正文机器校验（提示不是门禁，收尾前必须全绿）。

三种模式：
    python3 check_writing.py --guide                 校验写作指引（格式契约）
    python3 check_writing.py --section <节.docx>     单节自查（占位/残字/素材使用率/残留）
    python3 check_writing.py [--old-names "a,b"]     全局：承诺比对 + 节间查重 + 占位清点 + 残留

依赖：python-docx。
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
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


def die(msg: str, code: int = 1):
    print(f"[校验失败] {msg}")
    sys.exit(code)


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def digit_seq_hit(value: str, texts: list[str]) -> bool:
    """承诺比对宽松复核：值里的数字序列按原顺序出现在任一正文里（空格/标点/
    同义前后缀差异放行，如「每天10万对」↔「日均 10 万对」、「7×24」↔「7*24」）。
    值无数字（纯文字承诺）不算命中——无信号可核。"""
    seq = re.findall(r"\d+(?:\.\d+)?", value)
    if not seq:
        return False
    for t in texts:
        tseq = re.findall(r"\d+(?:\.\d+)?", t)
        it = iter(tseq)
        if all(any(d == x for x in it) for d in seq):
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
    for use_key, where in material_use.items():
        if len(where) > 1:
            issues.append(f"同一素材区间进多节（{use_key}）：{'、'.join(where)}——同一内容区间不得进两节")
    missing = sorted(leaf_keys - seen_keys)
    if missing:
        warnings.append(f"目录叶子无指引行（{len(missing)} 个）：{'、'.join(missing[:8])}"
                        + ("…" if len(missing) > 8 else ""))
    return issues, warnings


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
    for p, (_, labeled) in texts.items():
        for loc, frag in labeled:
            for line in frag.splitlines():
                if PLACEHOLDER.search(line):
                    issues.append(f"{p.name}:{loc} 占位文字：{line.strip()[:40]}")
                    n_ph += 1
        warnings += scan_ai_flavor(p.stem, labeled)
        gdoc = Document(str(p))
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
    ap.add_argument("--section", help="单节自查（.docx 路径）")
    ap.add_argument("--old-names", default="", help="残留扫描旧名清单（逗号分隔）")
    args = ap.parse_args()
    work = Path(args.work)

    if args.guide:
        issues, warnings = check_guide(work)
    elif args.section:
        sec = Path(args.section)
        if not sec.is_file():
            die(f"节文件不存在：{sec}")
        issues, warnings = check_section(work, sec)
    else:
        old_names = [x for x in args.old_names.split(",") if x.strip()]
        names_file = Path("旧名名单.txt")
        if names_file.is_file():
            old_names += [l.strip() for l in names_file.read_text(encoding="utf-8").splitlines()
                          if l.strip()]
        issues, warnings = check_global(work, old_names)

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
