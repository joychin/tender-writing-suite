#!/usr/bin/env python3
"""build_package.py — 上架单包构建器（平台一 zip＝一技能、只认 zip 根 SKILL.md）。

装配：入口技能 tender-writing-suite 在包根；四个环节技能整体搬入 stages/<环节>/；
入口细则三件在 tender-writing-suite/references/ 留构建期镜像——包内
`@tender-writing-suite/references/…` 限定引用按字面即能解析，无需改源文。

改写：包内 .md 的环节路径机械加 stages/ 前缀——
`(tender-analysis|tender-writing|tender-review|tender-docx)/` → `stages/<名>/`；
已含 stages/ 的行跳过（幂等）；`tender-writing-suite/` 天然不匹配（后接 `-` 非 `/`）。
仓库源目录不动（五技能平级形态仍是开发/逐技能安装的源形态）。

校验（任一不过即退出非零、不产 zip）：包根 SKILL.md 存在；全部 .py 可编译；
md 内 @引用与脚本路径在包内真实存在；改写后无遗漏的平级环节路径。

产物：out-zips/tender-writing-suite-<version>.zip（SKILL.md 在 zip 根），
装配树留在 out-zips/build/ 供人工检查。
安全声明：纯本地文件操作，无网络、无外部进程。
"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENTRY = "tender-writing-suite"
STAGES = ("tender-analysis", "tender-writing", "tender-review", "tender-docx")
VERSION_RE = re.compile(r"^version:\s*(\S+)", re.M)
STAGE_PATH = re.compile(rf"\b({'|'.join(STAGES)})/")
# 残留扫描：只抓未加 stages/ 前缀的环节路径（已改写的 stages/<环节>/ 不算）
STAGE_LEFTOVER = re.compile(rf"(?<!stages/)\b({'|'.join(STAGES)})/")
REF_ANY = re.compile(r"@([\w./-]+\.md)")            # 含裸 @references/ 与限定 @路径/ 两形
SCRIPT_REF = re.compile(r"([\w./-]*scripts/[\w.-]+\.py)")
SCAN_SUFFIXES = {".md", ".html", ".json", ".py"}


def copy_clean(src: Path, dst: Path) -> None:
    shutil.copytree(src, dst,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"))


def transform_md(md: Path) -> bool:
    """环节路径加 stages/ 前缀；已含 stages/ 的行原样保留。返回是否发生改写。"""
    lines = md.read_text(encoding="utf-8").splitlines(keepends=True)
    out, changed = [], False
    for line in lines:
        if "stages/" in line or not STAGE_PATH.search(line):
            out.append(line)
            continue
        out.append(STAGE_PATH.sub(r"stages/\1/", line))
        changed = True
    if changed:
        md.write_text("".join(out), encoding="utf-8")
    return changed


def main() -> int:
    entry_dir = ROOT / ENTRY
    if not (entry_dir / "SKILL.md").is_file():
        print(f"[错误] 入口技能缺失：{entry_dir}", file=sys.stderr)
        return 1
    for s in STAGES:
        if not (ROOT / s / "SKILL.md").is_file():
            print(f"[错误] 环节技能缺失：{ROOT / s}", file=sys.stderr)
            return 1
    m = VERSION_RE.search((entry_dir / "SKILL.md").read_text(encoding="utf-8"))
    if not m:
        print("[错误] 入口 SKILL.md 缺 version 字段", file=sys.stderr)
        return 1
    version = m.group(1)

    build_root = ROOT / "out-zips" / "build"
    pkg = build_root / ENTRY
    shutil.rmtree(build_root, ignore_errors=True)
    build_root.mkdir(parents=True)

    # ① 装配：入口在包根 + 四环节搬入 stages/ + 入口细则镜像
    copy_clean(entry_dir, pkg)
    (pkg / "stages").mkdir()
    for s in STAGES:
        copy_clean(ROOT / s, pkg / "stages" / s)
    mirror = pkg / ENTRY / "references"
    mirror.mkdir(parents=True)
    for ref in sorted((pkg / "references").glob("*.md")):
        shutil.copy2(ref, mirror / ref.name)

    # ② 改写
    n_changed = sum(transform_md(md) for md in pkg.rglob("*.md"))

    # ③ 校验
    problems: list[str] = []
    if not (pkg / "SKILL.md").is_file():
        problems.append("包根缺 SKILL.md")
    for py in sorted(pkg.rglob("*.py")):
        try:
            # builtin compile 只做语法编译不落 bytecode；py_compile 会在装配树里写出
            # __pycache__/*.pyc（曾随包出去，违背「明文可审」声明）
            compile(py.read_text(encoding="utf-8"), str(py), "exec")
        except SyntaxError as e:
            problems.append(f"编译失败 {py.relative_to(pkg)}: {e}")
    for md in sorted(pkg.rglob("*.md")):
        text = md.read_text(encoding="utf-8")
        # 裸简写（@references/…、scripts/…）按该 md 所属技能根解析；
        # 限定路径（stages/…、tender-writing-suite/…）按包根解析
        rel = md.relative_to(pkg)
        md_skill_root = pkg / "stages" / rel.parts[1] if rel.parts[0] == "stages" else pkg
        for ref in REF_ANY.findall(text):
            target = (md_skill_root / ref) if ref.startswith("references/") else (pkg / ref)
            if not target.is_file():
                problems.append(f"{md.relative_to(pkg)}: @{ref} 不存在")
        for script in SCRIPT_REF.findall(text):
            target = (md_skill_root / script) if script.startswith("scripts/") else (pkg / script)
            if not target.is_file():
                problems.append(f"{md.relative_to(pkg)}: {script} 不存在")
    for f in sorted(pkg.rglob("*")):
        if f.is_file() and f.suffix in SCAN_SUFFIXES:
            for i, line in enumerate(f.read_text(encoding="utf-8", errors="ignore").splitlines(), 1):
                for hit in STAGE_LEFTOVER.finditer(line):
                    problems.append(f"{f.relative_to(pkg)}:{i} 环节路径未加 stages/ 前缀: {hit.group(0)}")
    if problems:
        print("[构建失败] 校验未过：", file=sys.stderr)
        for p in problems:
            print(f"- {p}", file=sys.stderr)
        return 1

    out = ROOT / "out-zips" / f"{ENTRY}-{version}.zip"
    out.unlink(missing_ok=True)
    for cache in pkg.rglob("__pycache__"):
        shutil.rmtree(cache, ignore_errors=True)
    files = [f for f in sorted(pkg.rglob("*"))
             if f.is_file() and "__pycache__" not in f.parts]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.relative_to(pkg).as_posix())
    print(f"[构建完成] {version}：改写 {n_changed} 个 md，{len(files)} 个文件，"
          f"{out.stat().st_size // 1024} KB")
    print(f"zip：{out}（SKILL.md 在 zip 根）")
    print(f"装配树（人工检查用）：{pkg}")
    print("[验收提示] 真机验证：解压到技能目录只留 tender-writing-suite/ 一个目录，"
          "用真实招标文件跑完整流程")
    return 0


if __name__ == "__main__":
    sys.exit(main())
