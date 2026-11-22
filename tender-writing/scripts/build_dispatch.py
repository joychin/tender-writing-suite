#!/usr/bin/env python3
"""build_dispatch.py — 第五步并行分派的批次提示词拼装器。

把主会话选好的一批节拼成一份自包含批次文件（work/dispatch/批次-NN-<缩写>.md），
子任务以该文件为输入开工——主会话不再手写自包含提示词（模型罗列的清单实测会
塌成几字节缩写，子代理只能 ls/grep 自救重建上下文，占整本输入 token 六成）。
拼装内容全部来自落盘产物（目录树/写作指引/承诺清单/已解析素材/已写节），同批
字节确定、可复查、改产物后重跑即得新批次。

用法：
    python3 build_dispatch.py --work work --sections "项目理解,进度规划"
    python3 build_dispatch.py --work work --sections "商务标/服务方案,售后服务方案"

节名匹配：与写作指引「节」列同规（多册写「册名/标题」，单册裸名；也接受裸标题
在全部卷内唯一命中的写法）。节名自带「/」（如「…身份证明/授权委托书」）按整名
优先匹配目录叶子，「册名/」前缀只在消歧时切。任一节名对不上指引行 → 整体报错
退出，不产半批（拼装节与未拼装节信息不对称比失败更糟）。

批次文件结构（块序固定）：任务行 / 开工纪律 / 脚本绝对路径 / 今天日期 /
逐节块【第 i 节】（输出路径·模式·图示·要求原文·素材名片·缺口原文）/
承诺清单全文 / 兄弟节开头摘要 / 写作方法论全文（section-writing.md 内联）。

安全声明：纯本地文件读取，写一个批次文件；无网络、无外部进程。
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
# 环节间共享模块解析（双形态）：先按五技能平级找 ../<环节>，再按上架单包找
# ../stages/<环节>（环节收在 stages/ 下的单技能包），都缺才报错
_STAGE_ROOT = Path(__file__).resolve().parent.parent


def _sibling_scripts(name: str) -> Path:
    for cand in (_STAGE_ROOT.parent / name, _STAGE_ROOT.parent / "stages" / name):
        scripts = cand / "scripts"
        if scripts.is_dir():
            return scripts
    raise SystemExit(f"缺少兄弟环节 {name}（平级与 stages/ 单包形态均未找到）：{_STAGE_ROOT.parent}")


_DOCX_SCRIPTS = _sibling_scripts("tender-docx")
sys.path.append(str(_DOCX_SCRIPTS))
from check_writing import _registry_payload, table_rows  # noqa: E402
from docx_toolkit import build_tree, parse_outline, sanitize_name  # noqa: E402

SCRIPTS_DIR = Path(__file__).parent
METHOD_PATH = SCRIPTS_DIR.parent / "references" / "section-writing.md"
INTERNAL_ID_RE = re.compile(r"(?:MAND|TPL|REQ|SCORE)-\d{1,3}[、，]?\s*")
BARE_ID_RE = re.compile(r"^(MAND|TPL|REQ|SCORE)-\d{1,3}$", re.I)


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8") if path.is_file() else ""


def _outline_leaves(work: Path) -> tuple[list[str], dict[str, str]]:
    """(卷名列表, 叶子标题→卷名)。叶子键=目录树原文名。"""
    outline = work / "outline" / "tender-response-docs.md"
    if not outline.is_file():
        sys.exit("[派发中止] 目录文件不存在（work/outline/tender-response-docs.md）"
                 "——先走第三步目录树")
    vols = parse_outline(outline.read_text(encoding="utf-8"))
    leaf_vol: dict[str, str] = {}

    def walk(ns: list[dict], vol: str) -> None:
        for n in ns:
            if n["children"]:
                walk(n["children"], vol)
            else:
                leaf_vol.setdefault(n["text"], vol)

    names = []
    for v in vols:
        names.append(v["name"])
        walk(build_tree(v["nodes"]), v["name"])
    return names, leaf_vol


def _load_guide(work: Path) -> list[list[str]]:
    guide = work / "body" / "写作指引.md"
    rows = table_rows(_read(guide))
    if not rows or len(rows[0]) < 6:
        sys.exit("[派发中止] 写作指引缺失或表头不足 6 列——先走第四步")
    return [(r + [""] * 6)[:6] for r in rows[1:]]


def _resolve_sections(want: list[str], guide_rows: list[list[str]],
                      leaf_vol: dict[str, str]) -> list[tuple[str, list[str], str]]:
    """节名 → (叶子标题, 指引行, 卷名)。错名/歧义全部收集后一次报出（不半拼）。

    标题解析整名优先：指引键与目录叶子逐字相等直接用（单册裸名含「/」的节，
    如「…身份证明/授权委托书」）；不等才退回「册名/」前缀切分找叶子。切分只
    用于定位，不改写标题——节文件名与目录逐字一致全靠这一条。
    """
    by_key: dict[str, list[str]] = {}
    for r in guide_rows:
        by_key.setdefault(r[0].strip(), []).append(r)
    by_bare: dict[str, list[str]] = {}
    for key, rows in by_key.items():
        bare = key.rsplit("/", 1)[-1]
        by_bare.setdefault(bare, []).extend(rows)

    errors: list[str] = []
    picked: list[tuple[str, list[str], str]] = []
    seen: set[str] = set()
    for name in want:
        name = name.strip()
        cand = by_key.get(name) or by_bare.get(name) or by_key.get(sanitize_name(name)) \
            or by_bare.get(sanitize_name(name))
        if not cand:
            errors.append(f"节名「{name}」不在写作指引——指引节列写什么就传什么"
                          f"（多册带「册名/」前缀）")
            continue
        if len(cand) > 1:
            errors.append(f"节名「{name}」命中 {len(cand)} 条指引行（跨册重名）——"
                          f"请写「册名/{name}」消歧")
            continue
        row = cand[0]
        key_full = row[0].strip()
        if key_full in leaf_vol:
            title, vol = key_full, leaf_vol[key_full]
        else:
            title = key_full.rsplit("/", 1)[-1]
            vol = leaf_vol.get(title)
            if vol is None:
                errors.append(f"「{key_full}」不在目录树叶子中（裸名「{title}」也试过）"
                              f"——先对齐目录与指引")
                continue
        if key_full in seen:
            continue  # 同一节名写两遍 → 静默去重
        seen.add(key_full)
        picked.append((title, row, vol))
    if errors:
        sys.exit("[派发中止] 节名对不上（整体退出，不产半批——改对名字后重跑）：\n"
                 + "\n".join(f"- {e}" for e in errors))
    return picked


def _out_path(work: Path, title: str, vol: str, multi: bool) -> str:
    """节文件输出路径——带 work 前缀的工作区根相对路径（裸 body/ 会被按工作
    区根解析成 <根>/body/，真机实测出过错）。"""
    base = sanitize_name(title) + ".docx"
    rel = f"{sanitize_name(vol)}/{base}" if multi else base
    return f"{work.as_posix()}/body/{rel}"


def _req_lines(basis: str, registry: dict[str, dict]) -> tuple[list[str], list[str]]:
    """依据列 ID → 「- 要求原文（出处：…）」行；(解析出的行, 未知 ID 列表)。"""
    out: list[str] = []
    unknown: list[str] = []
    for tok in [t.strip() for t in re.split(r"[,，、]", basis) if t.strip()]:
        if not BARE_ID_RE.match(tok):
            if tok not in ("—", ""):
                out.append(f"- {tok}")
            continue
        cid = tok.upper()
        m = re.match(r"^(MAND|TPL|REQ|SCORE)-0*(\d{1,3})$", cid)
        cid = f"{m.group(1)}-{int(m.group(2)):02d}"
        entry = registry.get(cid)
        if not entry:
            unknown.append(tok)
            continue
        out.append(f"- {entry['text']}（出处：{entry['source']}）")
    return out, unknown


def _material_cards(material: str, work: Path) -> list[str]:
    """素材列 token → 名片行（区间字数+含图数）。素材未解析时降级点名不中断。"""
    out: list[str] = []
    for tok in [t.strip() for t in material.split("，") if t.strip()]:
        if tok == "【缺】":
            out.append("- 【缺】——资料通道与用户文件未命中，正文相应处加批注点名，禁止编造")
            continue
        m = re.match(r"^(.+\.docx)\s+L(\d+)(?:-L?(\d+))?$", tok)
        if not m:
            out.append(f"- {tok}（记法异常——应为 素材.docx L120-L260）")
            continue
        fname = re.split(r"[/\\]", m.group(1))[-1]
        lo, hi = int(m.group(2)), int(m.group(3) or m.group(2))
        md = work / "parse" / fname / (Path(fname).stem + ".md")
        if not md.is_file():
            out.append(f"- {tok}（素材未解析——主会话先跑 parse_document 后重新生成批次）")
            continue
        seg = "\n".join(md.read_text(encoding="utf-8", errors="replace").splitlines()[lo - 1:hi])
        chars = len(re.sub(r"\s+", "", seg))
        figs = seg.count("![")
        tag = f"，含图 {figs} 处" if figs else ""
        out.append(f"- 《{fname}》L{lo}-L{hi}（约 {chars:,} 字{tag}）")
    return out


def _promise_lines(work: Path) -> list[str]:
    rows = table_rows(_read(work / "body" / "关键事实与承诺.md"))
    out: list[str] = []
    for r in rows[1:] if rows and rows[0] and "事项" in rows[0][0] else rows:
        r = (r + [""] * 3)[:3]
        if not r[0].strip() or r[0].strip() == "事项":
            continue
        note = f"（{r[2].strip()}）" if r[2].strip() else ""
        out.append(f"- {r[0].strip()}：{r[1].strip() or '待补'}{note}")
    return out


def _sibling_summaries(work: Path, vol_dir: str, exclude: set[str]) -> list[str]:
    body_dir = work / "body" / vol_dir
    if not body_dir.is_dir():
        return []
    from docx import Document  # 局部 import：只有真要摘要时才载入 python-docx
    cands = []
    for p in sorted(body_dir.glob("*.docx")):
        if p.stem in exclude or p.name.startswith("整本-"):
            continue
        cands.append(p)
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    out: list[str] = []
    for p in cands[:3]:
        try:
            head: list[str] = []
            total = 0
            for para in Document(str(p)).paragraphs:
                t = para.text.strip()
                if t:
                    head.append(t)
                    total += len(t)
                    if total >= 200:
                        break
            out.append(f"- {p.stem}：{''.join(head)[:200]}")
        except Exception:
            continue
    return out


def build(work: Path, sections_arg: str, out_dir: Path) -> Path:
    want = [s for s in re.split(r"[,，]", sections_arg) if s.strip()]
    if not want:
        sys.exit("[派发中止] --sections 为空——传节名清单（逗号分隔）")
    vols, leaf_vol = _outline_leaves(work)
    guide_rows = _load_guide(work)
    registry, _scores = _registry_payload(work)
    picked = _resolve_sections(want, guide_rows, leaf_vol)
    multi = len(vols) > 1

    # 依据列 ID 先全部对账（任一未知 → 整体报错，不半拼）
    unknown_all: list[str] = []
    for _key, row, _vol in picked:
        _lines, unknown = _req_lines(row[2], registry)
        unknown_all += unknown
    if unknown_all:
        sys.exit("[派发中止] 依据列引用了登记表不存在的编号：" + "、".join(sorted(set(unknown_all)))
                 + "——登记表与指引漂移，先重跑 check_writing.py --outline 对账")

    blocks: list[str] = []
    titles = "、".join(t for t, _r, _v in picked)
    blocks.append(f"【写节任务】{titles}（共 {len(picked)} 节，逐节完成：建节→注入/撰写→修订"
                  f"→跑 check_writing.py --section 自查，全节自查通过再收尾）")
    blocks.append(
        "〔系统附：本批派发上下文（程序生成，直接使用）。开工纪律——输出路径（工作区根"
        "相对路径，按所写路径直接用，不自拼 body/）、要求原文、承诺值、素材名片全部在下方："
        "开局禁止 ls/glob 探测目录；禁止检索或使用 "
        "MAND/TPL/REQ/SCORE 内部编号（要求已按原文解析给出，编号不在你的语境）；"
        "无需再读写作指引/关键事实与承诺；直接从第一节开工。〕")
    blocks.append(f"脚本绝对路径：check_writing.py 在 {SCRIPTS_DIR}；"
                  f"docx_toolkit.py 在 {_DOCX_SCRIPTS}"
                  f"（所有命令用绝对路径调用，不假设工作目录）")
    blocks.append(f"今天日期：{date.today().isoformat()}（日期只取此行，不自编）")

    for i, (title, row, vol) in enumerate(picked, 1):
        lines = [f"【第 {i} 节：{title}】", f"输出路径：{_out_path(work, title, vol, multi)}",
                 f"写作模式：{row[1].strip() or '—'}"]
        figs = [t.strip() for t in row[4].split("、") if t.strip() and t.strip() != "—"]
        lines.append("图示清单（逐项产出，收尾对账；表=建节 body 的 table 块、分层/甘特="
                     "docx_toolkit diagram 子命令）：" + ("；".join(figs) if figs else "—（不为丰富而凑图）"))
        reqs, _u = _req_lines(row[2], registry)
        lines.append("要求清单（正文呼应要求本身或原文真实条款号，编号一个不写进正文）：")
        lines += [f"  {r}" for r in reqs] if reqs else ["  —（无登记依据）"]
        lines.append("素材名片（拷贝授权范围；docx 原件用 copy-from 按行号区间保真注入，"
                     "禁止凭文本自写素材）：" if row[3].strip() not in ("", "—") else "素材：—（无素材节）")
        if row[3].strip() not in ("", "—"):
            lines += [f"  {c}" for c in _material_cards(row[3], work)]
        note = INTERNAL_ID_RE.sub("", row[5]).strip()
        lines.append("缺口/备注原文（【缺】=没有的资料，批注点名待办；【资料】=已命中，"
                     "数字照抄不改写）：" + (note if note and note != "—" else "—"))
        blocks.append("\n".join(lines))

    promises = _promise_lines(work)
    blocks.append("承诺清单全部值（承诺类数字只能用这里；「待补」=未拍板——相应位置"
                  "批注点名，禁止编造）：\n" + ("\n".join(promises) if promises
                                               else "  （清单缺失或为空——缺项一律批注点名，禁止编造）"))

    exclude = {sanitize_name(t) for t, _r, _v in picked}
    sibs = _sibling_summaries(work, sanitize_name(picked[0][2]) if multi else "",
                              exclude)
    blocks.append("兄弟节开头摘要（避免跨节重复展开；已排除本批节）："
                  + ("\n" + "\n".join(sibs) if sibs else "\n  （同册暂无已写节）"))

    method = _read(METHOD_PATH)
    if method:
        blocks.append("写作纪律（完整方法论已内联——不要再读 section-writing.md）：\n"
                      + method)

    out_dir.mkdir(parents=True, exist_ok=True)
    seq = len(list(out_dir.glob("批次-*.md"))) + 1
    slug = sanitize_name(picked[0][0])[:12]
    dest = out_dir / f"批次-{seq:02d}-{slug}.md"
    dest.write_text("\n\n".join(blocks) + "\n", encoding="utf-8")
    return dest


def main() -> None:
    ap = argparse.ArgumentParser(description="第五步并行分派批次提示词拼装器")
    ap.add_argument("--work", default="work")
    ap.add_argument("--sections", required=True, help="节名清单（逗号分隔；多册写「册名/标题」）")
    ap.add_argument("--out", default="", help="批次文件目录（默认 <work>/dispatch/）")
    args = ap.parse_args()
    work = Path(args.work)
    out_dir = Path(args.out) if args.out else work / "dispatch"
    dest = build(work, args.sections, out_dir)
    size = dest.stat().st_size
    print(f"[批次已生成] {dest}（{size / 1024:.0f} KB）")
    if not (work / "review" / ".progress-opened").is_file():
        print("⚠️ 写作进度面板未首开（work/review/.progress-opened 不存在）——派发前先跑 "
              "build_review.py 加 --open-progress 首开面板并把链接发给用户，"
              "否则子任务首次建节即被 docx_toolkit 硬拦。")
    print("派发提示词模板：把该文件路径交给子任务——「读取并严格按 <批次文件> 执行，"
          "文件自包含（纪律/输出路径/素材名片/承诺值都在里面），完成后逐节回传自查结果」。")


if __name__ == "__main__":
    main()
