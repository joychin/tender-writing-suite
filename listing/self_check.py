#!/usr/bin/env python3
"""self_check.py — 技能包上架前结构自检（重发新版时重跑）。

检查：frontmatter 必填字段、name/description 规格上限、SKILL.md 正文 token 规模、
产品专属词残留、沿革从句、@references 指向的文件存在。
"""

from __future__ import annotations

import re
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
    for ref in re.findall(r"@references/([\w.-]+\.md)", body):
        if not (pkg / "references" / ref).is_file():
            problems.append(f"{pkg.name}: @references/{ref} 文件不存在")
    for md in list(pkg.rglob("*.md")):
        text = md.read_text(encoding="utf-8")
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
    for pkg in sorted(p for p in ROOT.iterdir() if p.is_dir() and (p / "SKILL.md").exists()):
        problems += check(pkg)
    if problems:
        print(f"\n[自检未通过] {len(problems)} 处：")
        for p in problems:
            print(f"- {p}")
        sys.exit(1)
    print("\n[自检通过] 技能包结构合规")


if __name__ == "__main__":
    main()
