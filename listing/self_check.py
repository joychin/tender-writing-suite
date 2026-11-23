#!/usr/bin/env python3
"""self_check.py — 技能包上架前结构自检（重发新版时重跑）。

检查：frontmatter 必填字段、name/description 规格上限、SKILL.md 正文 token 规模
（逐技能）、产品专属词残留、沿革从句、同技能与跨技能引用指针的文件存在性。
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent

REQUIRED_FIELDS = ["name", "display_name", "display_name_en", "description",
                   "description_zh", "description_en", "version", "author"]
FORBIDDEN = [
    "check_pipeline_state", "publish_artifact", "read_artifact", "ask_human",
    "validate_analysis", "validate_body", "check_name_residue", "assemble_tender",
    "docx_section_", "docx_comment_", "docx_image_", "docx_material_",
    "docx_source_", "docx_assemble_", "docx_diagram_", "docx_html_",
    "update_task_progress", "write_todos", "search_references",
    "search_company_assets", "tender-outline-writer", "tender-body-writer",
    "恢复点", "产物面板", "t_ab12", "tender.directory", "聊天产物卡",
    "task 工具", "deepagents",
]
CHANGELOG_PHRASES = re.compile(r"此前|不再|原先|旧版|自\s*20\d\d\s*年起|历史上曾")
# 套件跨技能引用：@<技能名>/references/<文件>.md 与 <技能名>/scripts/<脚本>.py
CROSS_REF = re.compile(r"@([a-z0-9-]+)/references/([\w.-]+\.md)")
CROSS_SCRIPT = re.compile(r"([a-z0-9-]+)/scripts/([\w.-]+\.py)")

# SKILL.md 正文 token 闸（官方建议 body < 5000 tokens，见 agentskills.io 规范）。
# 汉字按 1 token/字数词远低报真实成本（实测约 1.6 token/字），闸门必须按 token 算：
# 装了 tiktoken 用 cl100k 实测；没有则用无依赖估算——16 个本仓文件拟合，
# cjk×1.55 + latin词×1.0，偏差 ±4% 以内且偏保守（宁可高估放行少）。
BODY_TOKEN_LIMIT = 5000


def parse_frontmatter(text: str) -> dict:
    assert text.startswith("---"), "缺少 frontmatter"
    end = text.index("---", 3)
    block = text[3:end]
    meta = {}
    for line in block.splitlines():
        if ":" in line and not line.startswith((" ", "-")):
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip().strip('"')
    return meta, text[end + 3:]


def yaml_safety(text: str) -> str | None:
    """frontmatter 的 YAML 语法门：装包平台用真 YAML 解析器，值里未加引号的
    半角冒号+空格（如 `writing: parse`）会被判嵌套映射而拒装。优先 PyYAML
    权威校验；环境无 PyYAML 时退化为逐行启发式（零依赖兜底）。"""
    block = text[3:text.index("---", 3)]
    try:
        import yaml
    except ImportError:
        for ln, line in enumerate(block.splitlines(), 1):
            s = line.strip()
            if not s or s.startswith("-") or ":" not in s:
                continue
            v = s.split(":", 1)[1].lstrip()
            if v.startswith(('"', "'")):
                continue
            if ": " in v or v.endswith(":"):
                return (f"frontmatter 第 {ln} 行的值含未引号包裹的半角冒号"
                        f"（YAML 会误判嵌套映射，装包被拒）——值加双引号：{s[:60]}")
        return None
    try:
        yaml.safe_load(block)
        return None
    except Exception as exc:
        return f"frontmatter YAML 解析失败（装包平台同样会拒）：{exc}"


def token_count(body: str) -> tuple[int, str]:
    """返回 (token 数, 口径标记)。tiktoken 实测优先，无依赖时估算。"""
    try:
        import tiktoken
    except ImportError:
        cjk = len(re.findall(r"[\u4e00-\u9fff]", body))
        latin = len(re.findall(r"[A-Za-z0-9]+", body))
        return round(cjk * 1.55 + latin), "估算"
    return len(tiktoken.get_encoding("cl100k_base").encode(body)), "实测"


DEP_HINT_MARK = "pip install python-docx"
# 模板主题语言的中文口径：python-docx 自带模板写 ja-JP，标题样式里的主题字体属性
# 会顶掉显式「黑体」，汉字落到日文字体并逐字回退。开发机常没装日文字体、
# 渲染时静默替换，肉眼验收看不见——只能量化检查模板本身。
CJK_THEME_LANG = "zh-CN"


def check_template_cjk_fonts() -> list[str]:
    """基准模板的中文字体链路：主题语言必须中文、theme 东亚字形非空（空 a:ea 时
    西文主题仍兜住东亚解析），且不得有主题属性顶掉显式东亚字体。"""
    import zipfile

    problems: list[str] = []
    tpl = ROOT / "tender-docx" / "templates" / "base-template.docx"
    if not tpl.is_file():
        return ["tender-docx/templates/base-template.docx 缺失"]
    try:
        with zipfile.ZipFile(tpl) as z:
            settings = z.read("word/settings.xml").decode("utf-8")
            styles = z.read("word/styles.xml").decode("utf-8")
            theme = z.read("word/theme/theme1.xml").decode("utf-8")
    except (zipfile.BadZipFile, KeyError) as exc:
        return [f"base-template.docx 读不出 settings/styles/theme：{exc}"]
    lang = re.search(r'<w:themeFontLang[^/]*w:eastAsia="([^"]*)"', settings)
    if not lang:
        problems.append("base-template.docx: settings.xml 缺 themeFontLang/@eastAsia"
                        "（主题东亚字体无从判定，Word 按界面语言猜）")
    elif lang.group(1) != CJK_THEME_LANG:
        problems.append(f"base-template.docx: 主题东亚语言是 {lang.group(1)}"
                        f"（应为 {CJK_THEME_LANG}）——日文等分支字体顶掉中文字体，"
                        "标题会落到 ＭＳ 明朝/ゴシック 并逐字回退；改 setup_styles 后重跑 make_template.py")
    shadowed = []
    for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>',
                         styles, re.S):
        sid, body = m.group(1), m.group(2)
        rf = re.search(r"<w:rFonts[^/]*/>", body)
        if not rf:
            continue
        tag = rf.group(0)
        for theme_attr, explicit_attr in (("eastAsiaTheme", "w:eastAsia"),
                                          ("asciiTheme", "w:ascii"),
                                          ("hAnsiTheme", "w:hAnsi")):
            if theme_attr + "=" in tag and explicit_attr + "=" in tag:
                shadowed.append(f"{sid}({theme_attr})")
                break
    if shadowed:
        problems.append("base-template.docx: 样式同时带主题字体属性与显式字体，"
                        f"主题属性优先会顶掉后者——{'、'.join(shadowed[:6])}"
                        f"{' 等共 ' + str(len(shadowed)) + ' 处' if len(shadowed) > 6 else ''}")
    empty_ea = len(re.findall(r'<a:ea typeface="\s*"', theme))
    if empty_ea:
        problems.append(f"base-template.docx: theme1.xml 有 {empty_ea} 处空 a:ea——"
                        "东亚分支留空时西文主题兜住解析，光归一 themeFontLang 压不住；"
                        "改 setup_styles（含 _set_theme_ea）后重跑 make_template.py")
    print(f"  模板中文字体：themeFontLang eastAsia={lang.group(1) if lang else '(缺)'}，"
          f"主题属性顶显式字体 {len(shadowed)} 处，空 a:ea {empty_ea} 处")
    return problems


def check_changelog_head() -> list[str]:
    """沿革从句闸补 CHANGELOG 维度：只查顶行（最新版本行）——AGENTS.md「版本史只
    陈述当前版本做什么」对新增行生效；历史行的叙事词不追溯（全表会命中 8 个
    历史版本行，属存量清理范畴）。"""
    p = ROOT / "listing" / "CHANGELOG.md"
    if not p.is_file():
        return ["listing/CHANGELOG.md 缺失"]
    first = p.read_text(encoding="utf-8").splitlines()[0]
    m = CHANGELOG_PHRASES.search(first)
    if m:
        s = max(0, m.start() - 15)
        return [f"listing/CHANGELOG.md 顶行: 沿革从句「{m.group(0)}」…{first[s:m.end() + 15]!r}"]
    return []


def no_dep_smoke() -> list[str]:
    """缺依赖冒烟：用不加载 site-packages 的解释器（-S）跑每个脚本的 --help。

    开发机都装了 python-docx，本地冒烟永远看不见「没装依赖」这种环境——
    沙箱评测/纯净解释器/别人的电脑里，顶层 import 会让脚本启动第一步就崩，
    平台据此判「文件质量检测不通过」。所以这条必须在零第三方库条件下过：
    ① 每个脚本 --help 必须成功；② 真需要 Word 能力的脚本须给人话提示而非 traceback。
    """
    problems: list[str] = []
    scripts = sorted(ROOT.glob("*/scripts/*.py"))
    for script in scripts:
        r = subprocess.run([sys.executable, "-S", str(script), "--help"],
                           capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            tail = (r.stderr or r.stdout).strip().splitlines()
            problems.append(f"{script.relative_to(ROOT)}: 缺依赖时 --help 失败——"
                            f"{tail[-1] if tail else '(无输出)'}")
    toolkit = ROOT / "tender-docx" / "scripts" / "docx_toolkit.py"
    if toolkit.is_file():
        r = subprocess.run([sys.executable, "-S", str(toolkit), "read", "缺依赖探测.docx"],
                           capture_output=True, text=True, timeout=120)
        out = r.stdout + r.stderr
        if DEP_HINT_MARK not in out:
            problems.append("docx_toolkit.py: 缺依赖时未给出安装提示（应打印 "
                            f"「{DEP_HINT_MARK} …」而非 traceback）")
        elif "Traceback" in out:
            problems.append("docx_toolkit.py: 缺依赖时仍抛 traceback（应只给可执行提示）")
    print(f"  缺依赖冒烟：{len(scripts)} 个脚本 --help"
          f"（{sys.executable.split('/')[-1]} -S）")
    return problems


def check(pkg: Path) -> list[str]:
    problems: list[str] = []
    skill = pkg / "SKILL.md"
    if not skill.is_file():
        return [f"{pkg.name}: SKILL.md 缺失"]
    meta, body = parse_frontmatter(skill.read_text(encoding="utf-8"))
    yerr = yaml_safety(skill.read_text(encoding="utf-8"))
    if yerr:
        problems.append(f"{pkg.name}: {yerr}")
    for f in REQUIRED_FIELDS:
        if not meta.get(f):
            problems.append(f"{pkg.name}: frontmatter 缺 {f}")
    name = meta.get("name", "")
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]{0,63}", name):
        problems.append(f"{pkg.name}: name 不合规（小写字母/数字/连字符 ≤64）：{name}")
    if len(meta.get("description", "")) > 1024:
        problems.append(f"{pkg.name}: description 超 1024 字符")
    tc, basis = token_count(body)
    if tc > BODY_TOKEN_LIMIT:
        problems.append(f"{pkg.name}: SKILL.md 正文约 {tc} tokens（{basis}），"
                        f"超出 {BODY_TOKEN_LIMIT} 上限——细节下沉 references/")
    for d in ("references", "scripts", "templates"):
        if (pkg / d).is_dir() and not any((pkg / d).iterdir()):
            problems.append(f"{pkg.name}: {d}/ 为空目录")
    for md in list(pkg.rglob("*.md")):
        text = md.read_text(encoding="utf-8")
        for ref in re.findall(r"@references/([\w.-]+\.md)", text):
            if not (pkg / "references" / ref).is_file():
                problems.append(f"{md.relative_to(ROOT)}: 同技能 @references/{ref} 不存在")
        for skill_name, ref in CROSS_REF.findall(text):
            if not (ROOT / skill_name / "references" / ref).is_file():
                problems.append(f"{md.relative_to(ROOT)}: 跨技能 @{skill_name}/references/{ref} 不存在")
        for skill_name, script in CROSS_SCRIPT.findall(text):
            if not (ROOT / skill_name / "scripts" / script).is_file():
                problems.append(f"{md.relative_to(ROOT)}: 跨技能 {skill_name}/scripts/{script} 不存在")
        for w in FORBIDDEN:
            if w in text:
                first = text.index(w)
                problems.append(f"{md.relative_to(ROOT)}: 产品专属词「{w}」…{text[max(0,first-20):first+30]!r}")
        m = CHANGELOG_PHRASES.search(text)
        if m:
            s = max(0, m.start() - 20)
            problems.append(f"{md.relative_to(ROOT)}: 沿革从句「{m.group(0)}」…{text[s:m.end()+20]!r}")
    print(f"  {pkg.name}: 正文约 {tc} tokens（{basis}，上限 {BODY_TOKEN_LIMIT}），"
          f"frontmatter {len(meta)} 字段")
    return problems


def main() -> None:
    problems: list[str] = []
    versions: dict[str, str] = {}
    for pkg in sorted(p for p in ROOT.iterdir() if p.is_dir() and (p / "SKILL.md").exists()):
        problems += check(pkg)
        m = re.search(r"^version:\s*(\S+)",
                      (pkg / "SKILL.md").read_text(encoding="utf-8"), re.M)
        versions[pkg.name] = m.group(1) if m else "(缺)"
    # 版本一致性：五个 SKILL.md 必须同版（README 维护约定），漂移会让打包产物版本语义失真
    if len(set(versions.values())) > 1:
        detail = "、".join(f"{k}={v}" for k, v in sorted(versions.items()))
        problems.append(f"SKILL.md 版本不一致：{detail}——同步递增后再打包")
    problems += no_dep_smoke()
    problems += check_template_cjk_fonts()
    problems += check_changelog_head()
    if problems:
        print(f"\n[自检未通过] {len(problems)} 处：")
        for p in problems:
            print(f"- {p}")
        sys.exit(1)
    print("\n[自检通过] 技能包结构合规")


if __name__ == "__main__":
    main()
