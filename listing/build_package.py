#!/usr/bin/env python3
"""build_package.py — 上架单包构建器（扁平形态）。

平台硬限制（2026-11-21 实测解析报错：「目录层级超限 — Skill 仅支持两级目录结构
（根目录/二级目录/文件）」）：包内只允许 `SKILL.md` + `references/` + `scripts/` +
`templates/` 四个位置、全部文件平铺。任何三级目录（原 stages/<环节>/scripts/x.py）
都会解析失败。

因此本构建器把六技能源目录装配成**单技能扁平包**：

    SKILL.md                     ← 入口（tender-writing-suite/SKILL.md）
    references/
        stage-<环节>.md          ← 五个环节的 SKILL.md（去 frontmatter）
        <方法论>.md              ← references/*.md 拍平（已校验无重名）
    scripts/                     ← 九个脚本拍平（无重名）
    templates/                   ← 模板与页面拍平（无重名）

改写（只改装配副本，仓库源目录不动）：
1. `tender-<环节>/(references|scripts|templates)/x` → `\\2/x`
2. `tender-writing-suite/(references|scripts|templates)/x` → `\\1/x`
3. `tender-<环节>/SKILL.md` → `references/stage-<环节>.md`
4. 裸环节名 `tender-<环节>` → `stage-<环节>`（路由表用）
5. 「跨技能路径约定」整节替换为扁平常量；开头两段说明同样替换
6. 四个脚本的 `_sibling_scripts` 换成「扁平优先 → 平级 → stages」三形态解析

校验（任一不过即退出非零、不产 zip）：包根 SKILL.md 唯一存在；**所有文件层级 ≤2**；
各目录无重名；全部 .py 可编译；md 内 @引用与脚本路径在包内真实存在；无残留
`tender-<环节>/`、`tender-writing-suite`、`stages/` 字样。

产物：out-zips/tender-writing-suite-<version>.zip（无外层目录），
装配树留在 out-zips/build/tender-writing-suite/ 供人工检查。
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
STAGES = ("tender-analysis", "tender-writing", "tender-review", "tender-docx", "tender-sanitize")
SHORT = {s: s.replace("tender-", "") for s in STAGES}
VERSION_RE = re.compile(r"^version:\s*(\S+)", re.M)
FRONTMATTER_RE = re.compile(r"\A---\n.*?\n---\n", re.S)

# ---- 路径改写规则（顺序敏感） ----
STAGE_PATH = re.compile(rf"\b({'|'.join(STAGES)})/(references|scripts|templates)/")
SUITE_PATH = re.compile(rf"\b{ENTRY}/(references|scripts|templates)/")
STAGE_SKILLMD = re.compile(rf"\b({'|'.join(STAGES)})/SKILL\.md")
STAGE_BARE = re.compile(rf"\btender-({'|'.join(SHORT.values())})\b(?![-/])")
SUITE_LOAD = re.compile(rf"载入入口技能\s*{ENTRY}")
STAGE_LOAD = re.compile(r"载入\s+(stage-[\w]+)")
SUITE_BARE = re.compile(rf"(?<![-/\w]){ENTRY}(?![-/\w])")

# ---- 残留扫描 ----
LEFTOVER = (
    ("stages/ 布局残留", re.compile(r"stages/")),
    ("环节目录残留", re.compile(rf"\b({'|'.join(STAGES)})/")),
    ("裸入口名残留", re.compile(rf"(?<![-/\w]){ENTRY}(?![-/\w])")),
)
REF_ANY = re.compile(r"@([\w./-]+\.md)")
SCRIPT_REF = re.compile(r"([\w./-]*scripts/[\w.-]+\.py)")
SCAN_SUFFIXES = {".md", ".html", ".json", ".py"}

# ---- 整段替换 ----
INTRO_OLD = """从招标文件到可递交的标书 Word。本技能是六技能套件的**入口与总控**：红线、
确认门、环节路由与工作目录约定都在本文件；具体环节由平级技能执行——进入哪步
先载入该步点名的技能，再按那个技能的 SKILL.md 执行。

套件构成：本技能（入口）、tender-analysis（解析分析）、tender-writing（写作）、
tender-review（评审台）、tender-docx（合册交付）、tender-sanitize（文档脱敏）
——发布为单包（环节在包内 `stages/` 下）或六技能平级两种形态安装。"""
INTRO_NEW = """从招标文件到可递交的标书 Word。本技能是**唯一入口与总控**：红线、确认门、环节
路由与工作目录约定都在本文件；具体环节细则收在同包的 `references/stage-*.md`，
进入哪步就读那份细则并按它执行。

包内环节：stage-analysis（解析分析）、stage-writing（写作）、stage-review
（评审台）、stage-docx（合册交付）、stage-sanitize（文档脱敏）——脚本全在
`scripts/`、模板与页面全在 `templates/`，一套包一次安装即可跑完全流程。"""

# 「环节技能」是五技能布局的说法，扁平包里没有独立技能，统一改口
PACKAGE_FIXES = (
    ("回到发起环节技能（不强制下一环）", "回到发起环节（不强制下一环）"),
    ("每个环节技能只装自己的流程与细则指针", "每个环节细则只含自己的流程与指针"),
)

CONV_HEADING = "## 跨技能路径约定"

# 入口件在扁平包里的措辞修正（源文按五技能平级写的表述，单包内会误导执行）
ENTRY_FIXES = (
    ("本技能是六技能套件的入口与总控", "本技能是标书全流程套件的入口与总控"),
    ("过门：载入下一环节技能不构成连跑授权。", "过门：读下一环节细则不构成连跑授权。"),
    ("存在（见各环节技能「上游检查」）", "存在（见各环节细则「上游检查」）"),
    ("- 环节技能被单独触发时自带上游检查：", "- 各环节自带上游检查："),
)
CONV_NEW = """## 包内路径约定

本技能是**单技能扁平包**（平台限制：只允许「根目录 / 二级目录 / 文件」两级，
不得再嵌套子目录）：

- 入口：包根 `SKILL.md`（本文件）；
- 环节细则：`@references/stage-<环节>.md`——路由表里的 `stage-analysis` 等简写
  即指这份文件，进入该环节就读它并按其执行；
- 方法论与红线细则：`@references/<文件>.md`（如 `@references/stage-gates.md`）；
- 脚本：`scripts/<脚本>.py`（如 `scripts/build_review.py`）——四个环节的脚本同处
  一个目录，脚本自行解析彼此依赖，无需任何配置；
- 模板与页面：`templates/<文件>`（如 `templates/review-workbench.html`）；
- 产物只写工作目录 `work/`，绝不写技能目录。

**脚本不在工作目录**：命令示例里的相对路径以**本技能目录**为基准，执行前解析成
绝对路径（宿主技能目录变量如 `CODEBUDDY_SKILL_DIR` / `CLAUDE_PLUGIN_ROOT` 可直接
定位本技能）；工作目录只放 `work/` 产物。

"""

SIB_OLD = re.compile(
    r"# 环节间共享模块解析.*?"
    r"raise SystemExit\(f\"缺少兄弟环节 \{name\}.*?\)\n",
    re.S,
)
SIB_NEW = '''# 环节间共享模块解析（三形态）：单技能扁平包（全部环节脚本同目录）→ 五技能
# 平级 ../<环节> → 旧 stages/<环节> 布局；都缺才报错
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
'''


def copy_files(src: Path, dst: Path, pattern: str = "*") -> int:
    """把 src 下匹配的文件平铺拷进 dst（不保留子目录）。"""
    if not src.is_dir():
        return 0
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for f in sorted(src.glob(pattern)):
        if f.is_file() and f.name != ".DS_Store":
            shutil.copy2(f, dst / f.name)
            n += 1
    return n


def rewrite_paths(text: str) -> str:
    # frontmatter 是平台元数据（name/version 等），绝不参与路径改写
    m = FRONTMATTER_RE.match(text)
    head, body = (m.group(0), text[m.end():]) if m else ("", text)
    body = STAGE_PATH.sub(r"\2/", body)
    body = SUITE_PATH.sub(r"\1/", body)
    body = STAGE_SKILLMD.sub(r"references/stage-\1.md", body)
    body = SUITE_LOAD.sub("读本包根 SKILL.md", body)
    body = STAGE_BARE.sub(r"stage-\1", body)
    body = SUITE_BARE.sub("本包根 SKILL.md", body)
    # 「载入 stage-analysis」在扁平包里等于「读 references/stage-analysis.md」
    body = STAGE_LOAD.sub(r"读 @references/\1.md", body)
    body = body.replace("一个环节技能", "一份环节细则")
    return head + body


def replace_section(text: str, heading: str, new_block: str) -> str:
    """把 heading 起、至下一个同级/更高级标题前的内容整体替换。"""
    start = text.find(heading)
    if start < 0:
        raise SystemExit(f"[错误] 找不到章节：{heading}")
    m = re.search(r"\n## ", text[start + len(heading):])
    end = start + len(heading) + (m.start() + 1 if m else len(text) - start - len(heading))
    return text[:start] + new_block + text[end:]


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
    (pkg / "references").mkdir(parents=True)

    # ① 装配：入口 SKILL.md + 入口细则 + 五环节（细则/脚本/模板全拍平）
    shutil.copy2(entry_dir / "SKILL.md", pkg / "SKILL.md")
    n_ref = copy_files(entry_dir / "references", pkg / "references", "*.md")
    for s in STAGES:
        copy_files(ROOT / s / "references", pkg / "references", "*.md")
        copy_files(ROOT / s / "scripts", pkg / "scripts", "*.py")
        copy_files(ROOT / s / "templates", pkg / "templates")
        # 环节 SKILL.md → references/stage-<环节>.md（去 frontmatter）
        body = FRONTMATTER_RE.sub("", (ROOT / s / "SKILL.md").read_text(encoding="utf-8"))
        (pkg / "references" / f"stage-{SHORT[s]}.md").write_text(
            body.lstrip("\n"), encoding="utf-8")

    # ② 改写 md
    n_md = 0
    for md in pkg.rglob("*.md"):
        text = md.read_text(encoding="utf-8")
        if md.name == "SKILL.md" and md.parent == pkg:
            # 整段替换先做（锚点是源文原样），再做机械路径改写
            if INTRO_OLD not in text:
                print("[错误] 入口开头两段与预期不符，无法替换", file=sys.stderr)
                return 1
            text = text.replace(INTRO_OLD, INTRO_NEW)
            text = text.replace("| 阶段 | 载入技能 | 做什么 | 门 |",
                                "| 阶段 | 载入环节 | 做什么 | 门 |")
            text = replace_section(text, CONV_HEADING, CONV_NEW)
            for old, new in ENTRY_FIXES:
                if old not in text:
                    print(f"[警告] 入口措辞未命中（源文可能已改）：{old[:24]}", file=sys.stderr)
                text = text.replace(old, new)
        new = rewrite_paths(text)
        for old, repl in PACKAGE_FIXES:
            new = new.replace(old, repl)
        if new != text:
            md.write_text(new, encoding="utf-8")
            n_md += 1

    # ③ 脚本闸：依赖兄弟环节的脚本必须已是三形态源码（发货版代码存在于仓库源，
    # 本地电池测到的才是发货版——--help 在走到 _sibling_scripts 前就退出，装配
    # 替换曾让「单包形态复跑 --help」从没执行过发货版解析）
    problems: list[str] = []
    n_py = 0
    for py in pkg.rglob("*.py"):
        text = py.read_text(encoding="utf-8")
        if "_SENTINEL" in text:
            n_py += 1
            continue
        if "_sibling_scripts" not in text:
            n_py += 1  # 不做跨环节解析的脚本（docx_toolkit 本体、同目录 import 的 make_template 等）
            continue
        new, k = SIB_OLD.subn(SIB_NEW, text)
        if k:
            py.write_text(new, encoding="utf-8")
            n_py += 1
            problems.append(f"{py.relative_to(pkg)}: 兄弟解析仍是双形态源码（包内已临时替换）"
                            f"——把三形态版本回写仓库源后再打包")
        else:
            problems.append(f"{py.relative_to(pkg)}: 依赖兄弟环节但未找到解析段（_sibling_scripts）")
    for cache in pkg.rglob("__pycache__"):
        shutil.rmtree(cache, ignore_errors=True)

    # ④ 校验
    files = [f for f in sorted(pkg.rglob("*")) if f.is_file()]
    for f in files:
        rel = f.relative_to(pkg)
        if len(rel.parts) > 2:
            problems.append(f"层级超限（平台仅两级）：{rel}")
    roots = {f.relative_to(pkg).parts[0] for f in files}
    for r in roots - {"SKILL.md", "references", "scripts", "templates"}:
        problems.append(f"包根出现非法条目：{r}")
    for d in ("references", "scripts", "templates"):
        names = [f.name for f in (pkg / d).glob("*") if f.is_file()]
        dup = {n for n in names if names.count(n) > 1}
        for n in dup:
            problems.append(f"{d}/ 重名：{n}")
    for s in STAGES:
        if not (pkg / "references" / f"stage-{SHORT[s]}.md").is_file():
            problems.append(f"缺环节细则：references/stage-{SHORT[s]}.md")

    # 平台元数据闸：五项必填 + version 三段式 + name 与目录同名
    fm = FRONTMATTER_RE.match((pkg / "SKILL.md").read_text(encoding="utf-8"))
    if not fm:
        problems.append("SKILL.md 缺 YAML frontmatter")
    else:
        fm_text = fm.group(0)
        for key in ("name", "display_name", "display_name_en",
                    "description_zh", "description_en", "version"):
            if not re.search(rf"^{key}:\s*\S", fm_text, re.M):
                problems.append(f"frontmatter 缺 {key}")
        nm = re.search(r"^name:\s*(\S+)", fm_text, re.M)
        if nm and nm.group(1) != ENTRY:
            problems.append(f"frontmatter name={nm.group(1)} 与技能目录名 {ENTRY} 不一致")
        if not re.fullmatch(r"\d+\.\d+\.\d+", version):
            problems.append(f"version={version} 非三段式（平台要求 x.y.z）")

    for py in sorted(pkg.rglob("*.py")):
        try:
            compile(py.read_text(encoding="utf-8"), str(py), "exec")
        except SyntaxError as e:
            problems.append(f"编译失败 {py.relative_to(pkg)}: {e}")

    for md in sorted(pkg.rglob("*.md")):
        text = md.read_text(encoding="utf-8")
        for ref in REF_ANY.findall(text):
            cand = pkg / ref if ref.startswith(("references/", "scripts/", "templates/")) \
                else md.parent / ref
            if not cand.is_file():
                problems.append(f"{md.relative_to(pkg)}: @{ref} 不存在")
        for script in SCRIPT_REF.findall(text):
            cand = pkg / script if script.startswith("scripts/") else md.parent / script
            if not cand.is_file():
                problems.append(f"{md.relative_to(pkg)}: {script} 不存在")

    for f in sorted(pkg.rglob("*")):
        if f.is_file() and f.suffix in SCAN_SUFFIXES:
            is_py = f.suffix == ".py"
            text = f.read_text(encoding="utf-8", errors="ignore")
            offset = 0
            if f.name == "SKILL.md" and f.parent == pkg:
                # frontmatter 的 name/version 本就是技能标识，不算残留
                fm_entry = FRONTMATTER_RE.match(text)
                if fm_entry:
                    offset = text[:fm_entry.end()].count("\n")
                    text = text[fm_entry.end():]
            for i, line in enumerate(text.splitlines(), 1 + offset):
                for label, rx in LEFTOVER:
                    # 脚本保留 stages/ 兜底解析（三形态兼容），只查文档里的残留
                    if is_py and label.startswith("stages/"):
                        continue
                    for hit in rx.finditer(line):
                        problems.append(f"{f.relative_to(pkg)}:{i} {label}: {hit.group(0)}")

    if problems:
        print("[构建失败] 校验未过：", file=sys.stderr)
        for p in problems[:40]:
            print(f"- {p}", file=sys.stderr)
        if len(problems) > 40:
            print(f"… 另有 {len(problems) - 40} 条", file=sys.stderr)
        return 1

    # ⑤ 打包（无外层目录）
    out = ROOT / "out-zips" / f"{ENTRY}-{version}.zip"
    out.unlink(missing_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.relative_to(pkg).as_posix())

    size_kb = out.stat().st_size // 1024
    print(f"[构建完成] {version}：{len(files)} 个文件，{size_kb} KB"
          f"（改写 {n_md} 个 md / {n_py} 个 py，references 含 {n_ref} 份入口细则）")
    print(f"zip：{out}")
    print("顶层：" + " / ".join(sorted({f.relative_to(pkg).parts[0] for f in files})))
    if size_kb >= 3072:
        print("[警告] 超过平台 3MB 上限，需削减参考资料", file=sys.stderr)
        return 1
    print("[验收提示] 解压到技能目录后只应存在一个 tender-writing-suite/ 目录，"
          "用真实招标文件跑完整流程后再上传")
    return 0


if __name__ == "__main__":
    sys.exit(main())
