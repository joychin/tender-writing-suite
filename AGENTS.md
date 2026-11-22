# AGENTS.md — 本仓库维护须知

「灵燕标书写作套件」开源仓库：一个把招标文件写成可递交投标文件 Word 的 Agent Skill
（五技能套件），Python 3.10+，纯本地运行（所有脚本零网络请求）。
依赖：`python-docx pypdfium2 docxcompose Pillow`（可能还需 `tiktoken`/`PyYAML`，缺省有零依赖兜底）。

## 目录结构

- `tender-writing-suite/` — 入口技能（总控）：红线 7 条、确认门、技能路由、跨技能路径约定
- `tender-analysis/` `tender-writing/` `tender-review/` `tender-docx/` — 四个环节技能，平级安装；
  各自 `references/`（方法论 md）、`scripts/`（确定性 Python 引擎）、`templates/`
- `listing/` — 发布工具链：`self_check.py`（结构自检）、`build_package.py`（上架单包构建器）、
  `CHANGELOG.md`（版本史）；`上架指南.md`/`商品文案.md` 为内部运营材料（gitignored）
- `testruns/` — 测试夹具与验收电池（gitignored，含真实项目工作区样本）
- `out-zips/` — 打包产物（gitignored）；`docs/` — README 图片

仓库源是**五目录平级开发形态**；上架产物是**单技能扁平包**（平台硬限制：Skill 仅允许两级目录，
三级嵌套如 `stages/<环节>/scripts/x.py` 会被「目录层级超限」硬拒）。

## 常用命令

```bash
python3 listing/self_check.py        # 改动技能包后必跑：frontmatter/YAML/逐技能 token 闸/跨技能引用存在性/沿革从句门/缺依赖冒烟
python3 listing/build_package.py     # 发版打包：产出 out-zips/tender-writing-suite-<version>.zip（唯一单包）
python3 testruns/fp_smoke.py         # 评审台产物指纹链路冒烟（A1–A8）
python3 testruns/fp_e2e_cli.py       # 评审台 CLI 级全面断言（B1–B8，含 fp_smoke 12 例复跑）
# 脚本冒烟：八个脚本各自 `--help`（全部 argparse 化，--help 只看说明不落盘）
```

## 硬约束（发布闸，self_check 强制）

- **SKILL.md 正文 < 5000 tokens（逐技能）**：tiktoken cl100k 实测优先，无依赖时按
  `cjk×1.55 + latin词×1.0` 估算兜底。汉字按词数会低报 ~50%，不能按字数省。
- **FORBIDDEN 词表**（self_check.py 内）：旧内部工具名（`check_pipeline_state`、`ask_human`、
  `docx_section_` 系列等）不得出现在任何 SKILL.md / md 中。
- **沿革从句门**：CHANGELOG 与 SKILL.md 禁用「此前/不再/原先/旧版/历史上曾」等词——版本史
  只陈述当前版本做什么，不写对比叙事。
- **frontmatter 纪律**：值必须单行双引号（值内半角冒号+空格会被平台 YAML 解析器判嵌套映射而拒装）；
  必填字段见 self_check 的 `REQUIRED_FIELDS`。
- **版本联动**：五个 `SKILL.md` frontmatter `version` 必须同步递增，并在 `listing/CHANGELOG.md`
  顶部追加一行版本史。
- **结构纪律**：红线必须置顶（长会话压缩只保留文件开头）；路由表指向的环节技能与
  `@<技能>/references/*.md`、`<技能>/scripts/*.py` 引用必须真实存在（self_check 校验）。
- **第三方依赖惰性加载**：脚本顶层不得 `import python-docx`（及 pypdfium2/PIL 同族）——
  沙箱评测、纯净解释器、别人的电脑里启动第一步就 `ModuleNotFoundError`，平台据此判
  「文件质量检测不通过」。经 `docx_toolkit.require_docx()` 在 `parse_args()` 之后加载，
  缺失时打印中文安装提示并以退出码 2 退出（不抛 traceback）。开发机装了依赖，本地
  `--help` 永远看不出这个问题——所以 self_check 的 `no_dep_smoke` 用 `python3 -S`
  （不加载 site-packages）守着这条维度。

## 架构与编码约定

- **语义归模型，机械归程序**：解析、表格引擎、章节编号、目录页码、对账校验、评审台页面
  全部是确定性脚本；SKILL.md 只装流程纪律，不装可脚本化的逻辑。
- **环节间共享模块三形态解析**：脚本用 `_sibling_scripts()` 按
  扁平(上架包) → 平级 `../<环节>` → `../stages/<环节>` 顺序找兄弟环节 `scripts/`。
  新增跨环节引用必须走这个约定，不要写死路径。
- **产物 HTML 统一「方向 A 纸感公文体」**（评审台/报告页/进度页三模板），评审台是唯一确认面；
  改产物页面后跑 fp 两电池验收（认 DOM 断言，不认视觉模型）。
- **build_package.py 只改装配副本**：仓库源目录不动；md 路径前缀改写、环节 SKILL.md 转
  `references/stage-*.md`、`__pycache__` 不进包、改写后残留 `tender-<环节>/`/`stages/` 字样即失败。
- 产物/交付里的占位一律走 Word 批注，正文不许有占位文字（机器扫描兜底）。

## 发版流程（每次发版全跑）

1. 改动五个技能目录 → `python3 listing/self_check.py` 全绿
2. 八脚本 `--help` 冒烟 OK（平级形态）；缺依赖维度已由 self_check 的 `python3 -S`
   冒烟覆盖（装机有 deps 的本机 `--help` 看不见这条）
3. 涉及评审台/产物页面 → fp 两电池全绿
4. `python3 listing/build_package.py` 重打包（版本号已递增），
   可在 `out-zips/build/` 人工检查装配树、在单包形态下复跑脚本 `--help`
5. 发版后同步 workbuddy 侧副本（平台货架用，见 memory `tender-suite-hard-constraints`）

## 事实与演示口径

README 与示例中的「云澜市 / 星灵创科 / YLCG-2026-GK-0387」均为虚构演示主体；
`testruns/` 内含真实项目样本，属本机数据，不外发、不进包。
