# AGENTS.md — 本仓库维护须知

「灵燕标书写作套件」开源仓库：一个把招标文件写成可递交投标文件 Word 的 Agent Skill
（五技能套件），Python 3.10+，纯本地运行（所有脚本零网络请求）。
运行依赖见 `requirements.txt`（`python-docx pypdfium2 docxcompose Pillow`）；`tiktoken`/`PyYAML`
只给 self_check 用（缺省有零依赖兜底），`matplotlib` 只给甘特/架构真图用（无则降级出图）。
本机验收另需 LibreOffice（`soffice`）：目录实测页码与 `.doc` 转换靠它。

## 目录结构

- `tender-writing-suite/` — 入口技能（总控）：红线 8 条、确认门、技能路由、跨技能路径约定
- `tender-analysis/` `tender-writing/` `tender-review/` `tender-docx/` — 四个环节技能，平级安装；
  各自 `references/`（方法论 md）、`scripts/`（确定性 Python 引擎）、`templates/`
- `listing/` — 发布工具链：`self_check.py`（结构自检）、`build_package.py`（上架单包构建器）、
  `CHANGELOG.md`（版本史）；`上架指南.md`/`上架指引.docx`/`商品文案.md`/`提审说明.md` 为
  内部运营材料（gitignored，不随开源仓库分发）
- `tests/` — `engine_smoke.py`（docx_toolkit 引擎回归电池，入库版，合成夹具运行时生成；
  每条缺陷先写成失败用例再修）、`review_smoke.py`（评审台服务生命周期电池 R01–R09，
  跨平台零第三方依赖，守 Windows 探活/关停语义）
- `testruns/` — 测试夹具与验收电池（gitignored，含真实项目工作区样本）；`fp_smoke.py`/
  `fp_e2e_cli.py` 两电池及其夹具 `qszbh260186/work` 只存在于本机，新克隆的仓库跑不了
- `原型/` — 评审台页面四方向视觉原型（「方向 A 纸感公文体」已采纳实施，留存备查）
- `out-zips/` — 打包产物（gitignored）；`docs/` — README 图片

仓库源是**五目录平级开发形态**；上架产物是**单技能扁平包**（平台硬限制：Skill 仅允许两级目录，
三级嵌套如 `stages/<环节>/scripts/x.py` 会被「目录层级超限」硬拒）。

## 常用命令

```bash
python3 listing/self_check.py        # 改动技能包后必跑：frontmatter 必填与规格/YAML/逐技能 token 闸/空目录/
                                     #   同技能与跨技能引用存在性/产品专属词/沿革从句/五包版本一致/
                                     #   缺依赖冒烟/模板中文字体/CHANGELOG 顶行
python3 tests/engine_smoke.py        # 引擎回归电池 E01–E54（内容保真/批注与编号/跨平台与发版闸/真机复盘/Windows 兼容/
                                     #   格式件保真红线/文档检查）；
                                     #   改 docx_toolkit 前先补失败用例；E24/E29/E34 靠 LibreOffice，缺了记失败不跳过
python3 tests/review_smoke.py        # 评审台服务生命周期电池 R01–R09（零第三方依赖，跨平台）；
                                     #   动 serve_review 的探活/关停前先补失败用例
python3 tender-docx/scripts/make_template.py   # 改 docx_toolkit 的 setup_styles 后重生成 base-template.docx
python3 tender-docx/scripts/docx_toolkit.py inspect --work work   # 文档检查（只读）：整本自身 + 与节文件的对账；
                                     #   也可直接给一份或几份 docx；改检查逻辑前先在本机真实工作区回放找漏报误报
python3 testruns/fp_smoke.py         # 评审台产物指纹链路冒烟（A1–A8；本机专属，见 testruns 一条）
python3 testruns/fp_e2e_cli.py       # 评审台 CLI 级全面断言（B1–B8，含 fp_smoke 12 例复跑；本机专属）
python3 listing/build_package.py     # 发版打包：产出 out-zips/tender-writing-suite-<version>.zip（唯一单包）
# 脚本冒烟：八个脚本各自 `--help`（全部 argparse 化，--help 只看说明不落盘）
# 调试：docx_toolkit 缺省只回一行式报错，加环境变量 TENDER_TRACEBACK=1 才出完整栈
```

## 硬约束（发布闸，self_check 强制）

- **SKILL.md 正文 < 5000 tokens（逐技能）**：tiktoken cl100k 实测优先，无依赖时按
  `cjk×1.55 + latin词×1.0` 估算兜底。汉字按词数会低报 ~50%，不能按字数省。
- **产品专属词**（self_check 的 `FORBIDDEN`，子串匹配）：旧内部工具名与内部词
  （`check_pipeline_state`、`ask_human`、`docx_section_` 系列等）不得出现在五个技能目录下的
  任何 md（含 `references/` 与 `templates/*.md`）。
- **沿革从句门**（`CHANGELOG_PHRASES`）：同一扫描范围禁用「此前/不再/原先/旧版/自 20xx 年起/
  历史上曾」——版本史只陈述当前版本做什么，不写对比叙事。CHANGELOG 只查顶行（新版本行）；
  闸是子串匹配，在版本行里引用禁词同样命中，改用「改前」一类说法。
- **frontmatter 纪律**：字符串值必须单行双引号（值内半角冒号+空格会被平台 YAML 解析器判嵌套映射而拒装）；
  `name` 限小写字母/数字/连字符且 ≤64，`description` ≤1024 字符；必填字段见 self_check 的 `REQUIRED_FIELDS`。
- **版本联动**：五个 `SKILL.md` frontmatter `version` 必须同步递增，并在 `listing/CHANGELOG.md`
  第 1 行插入新版本行（一行一版本、无表头，最新在最上）。README 的版本徽章是手写值，
  没有任何闸校验，发版时手动同步。
- **结构纪律**：红线必须置顶（长会话压缩只保留文件开头）；路由表指向的环节技能与
  `@<技能>/references/*.md`、`<技能>/scripts/*.py` 引用必须真实存在（self_check 校验）。
- **第三方依赖惰性加载**：脚本顶层不得 `import python-docx`（及 pypdfium2/PIL 同族）——
  沙箱评测、纯净解释器、别人的电脑里启动第一步就 `ModuleNotFoundError`，平台据此判
  「文件质量检测不通过」。经 `docx_toolkit.require_docx()` 在 `parse_args()` 之后加载，
  缺失时打印中文安装提示并以退出码 2 退出（不抛 traceback）。开发机装了依赖，本地
  `--help` 永远看不出这个问题——所以 self_check 的 `no_dep_smoke` 用 `python3 -S`
  （不加载 site-packages）守着这条维度；新增脚本放进 `<技能>/scripts/` 即自动纳入。

## 架构与编码约定

- **语义归模型，机械归程序**：解析、表格引擎、章节编号、目录页码、对账校验、评审台页面
  全部是确定性脚本；SKILL.md 只装流程纪律，不装可脚本化的逻辑。增补纪律（节奏/停轮/门类
  规则）前先问能否做成产物前置检查或脚本闸——能就做机制，不能才写进 SKILL.md，且只写在门粒度。
- **招标格式件只拷不改是红线**（入口红线 8；用户 2026-11-24 拍板：擅自改模板=废标级，
  版面难看不加调宽工具、加批注请用户手动处理）。机制是 `docx_toolkit.format_fidelity`：
  按目录「来源标注」的 TPL 编号回分析登记表取出处区间（编号=模板表行序，与
  `_registry_payload` 同口径；「同文件」依次退到登记表头部声明、sources.json 主文件、
  唯一的那份解析产物），在解析 md 里找原件表，与节内表格比网格列数、标签保留、整表
  存在；`check_writing --section`／全局校验与 `assemble` 三道共用，`--allow-format-change`
  只在用户对话明确同意后使用。**触发只认目录「来源标注」**——指引依据列是模型写的，
  错挂 TPL 会把无关的节拖进来误报。动它之前先看 `tests/engine_smoke.py`（E40）。
- **文档检查只读不拦**（`docx_toolkit inspect`；`check_writing` 的单节自查与全局校验
  后面跟同一段，不改变退出码）：结果 = 一行结论 + 一张表（位置／问题／依据／确定还是
  疑似）+ 定位行 + 已检查范围，用户要求修订才动文件。检查自身出错只产出一条「没能
  执行」的疑似发现；读不出的金额格、认出是报价表却没有一行能验算成的必须点名——最
  危险的失败不是误报，是不声不响地没查。三条写盘约定：①引擎写整本一律走
  `_save_doc()`（内部调 `write_fingerprint()`，指纹记在 `<整本名>.docx.fingerprint.json`），
  新增写整本的代码不许绕开，否则「整本被手工改过」会误报；②`base_document()` 起建
  即清文档属性（作者／最后修改者／备注），别在别处再写作者；③`inspect` 任何情况下
  不落盘。阈值是拿本机真实工作区回放标定的（图片 72 DPI、拉伸 5%、图与表超宽容差
  0.5 厘米、窄列须有 4 字以上实际内容、`image` 按页宽减左右边距插图），改之前先回放。
  动它之前先看 `tests/engine_smoke.py`（E41–E54）。
- **环节间共享模块三形态解析**：脚本用 `_sibling_scripts()` 按
  扁平(上架包) → 平级 `../<环节>` → `stages/<环节>` 顺序找兄弟环节 `scripts/`
  （现由 `build_dispatch`/`check_writing`/`build_review`/`apply_review` 四脚本使用）。
  新增跨环节引用必须走这个约定，不要写死路径。**这段是发货源码**：打包器不再
  装配时替换（源码缺三形态会打包失败）——改解析逻辑只改源码一处，本地电池测到
  的就是发货版。
- **发版闸闸表**：self_check 新增维度时同步更新本文件与常用命令注释；沿革从句
  对 CHANGELOG 只查顶行（历史行不追溯）。
- **产物 HTML 统一「方向 A 纸感公文体」**（评审台/报告页/进度页三模板），评审台是唯一确认面；
  改产物页面后跑 fp 两电池验收（认 DOM 断言，不认视觉模型）。进度页五区标题被
  `fp_e2e_cli` 的 B2h 逐字断言，改标题须同步断言。
- **用户可见措辞有定稿口径**：交付物叫「整本 Word 文档」，「合册」是内部动作词不进对话；
  页面自称「评审页面」；说「批注清单」「待你确认」。细则见入口
  `references/reply-style.md`，改 SKILL.md、页面模板、脚本回显里的措辞时对照。
- **build_package.py 只改装配副本**：仓库源目录不动；md 路径前缀改写、环节 SKILL.md 转
  `references/stage-*.md`、`__pycache__` 不进包、改写后残留 `tender-<环节>/`/`stages/` 字样即失败。
  **入口 `SKILL.md` 有三处逐字锚点**：开头两段（`INTRO_OLD`，内含看似过时的 `stages/` 表述，
  别单方面改）、路由表表头 `| 阶段 | 载入技能 | 做什么 | 门 |`、`## 跨技能路径约定` 标题
  ——改源文必须同步改打包器，否则构建直接失败（`ENTRY_FIXES` 的措辞未命中只告警）。
- 产物/交付里的占位一律走 Word 批注，正文不许有占位文字（机器扫描兜底）。
- **docx_toolkit 不变量**（每条背后都有一次静默丢内容或跑版的事故，本地未必看得出来；
  动它之前先看 engine_smoke 对应用例）：
  批注读写只走 `_comments_root`/`_comments_flush`，禁止直摸 `part._blob`（python-docx ≥1.2 下
  是死存储，跨进程第二次加的批注正文会丢）；取段落全文一律 `para_text_full()`
  （`paragraph.text` 漏超链接与修订插入里的文字）；清格重写走 `_rewrite_cell_text()`
  （`cell.text = ""` 连批注锚一起抹）；`base-template.docx` 是 `make_template.py` 从
  `setup_styles` 生成的，改样式定义后必须重跑。中文字体跑偏在没装日文字体的开发机上渲染
  会静默回退，肉眼验收看不见，只认 self_check 的模板字体闸。LibreOffice 查找有两份同口径的
  `_find_soffice`（本文件与 `parse_document.py`，环节技能各自独立安装故不共享模块），
  改一处同步另一处（E35 对两份断言同一结果）。
- **serve_review 的进程探活必须跨平台**：Windows 下 `os.kill(pid, 0)` 不是探活——信号 0
  就是 `CTRL_C_EVENT`，走 `GenerateConsoleCtrlEvent`：目标不是进程组组长就抛 WinError 87
  （`OSError`，2.2.0 真机 `--status` 的 traceback 即此），是组长（宿主以新进程组起后台
  任务时）就等于给服务发 Ctrl+C。`_pid_alive` 在 Windows 走 `OpenProcess` 只读查询，任何
  探测异常按「不存活」处理；`--stop` 在 Windows 是强制结束（没有 SIGTERM 投递），`SIGKILL`
  用 `getattr` 兜底（Windows 无此信号）。**结束进程前先验身份**：记录的 pid 必须在
  serve.json 记的地址上应答 healthz 才结束——进程号会被别的程序复用（Windows 回收很快），
  对复用者发结束信号就是误杀。环回探活用不带代理的 opener（`urlopen` 缺省读系统代理，
  豁免名单没含 127.0.0.1 时活着的服务会被判「未运行」）。动它之前先看
  `tests/review_smoke.py`（R01–R09）。
- **脚本回显是给模型读的**：docx_toolkit 缺省一行式报错（`TENDER_TRACEBACK=1` 才出完整栈）；
  八个脚本起手都对 stdout/stderr `reconfigure(errors="replace")`（中文 Windows 的 GBK
  管道下遇 emoji 会崩），新增脚本照做。
- **Windows 兼容清单**（开发机是 macOS，这些差异本地跑不出来；真机是 workbuddy on
  Windows——bash 风格 shell、宿主自带后台运行、模型用 `python` 起脚本。新增子进程、
  临时目录、外部文件读取、`body/` 枚举时逐条对照；对应用例靠 mock 在 macOS 上复现）：
  外部程序不假设在 PATH 里（LibreOffice 安装器不写 PATH，查约定安装位置与注册表）；
  文件地址用 `Path.as_uri()` 生成（手工拼 `file://`+路径 在 Windows 是坏地址）；子进程
  输出解码带 `errors="replace"`；临时目录里的句柄用完即关，清理失败不得盖掉真正的报错
  （`docx_toolkit._tempdir()`）；枚举 `body/` 下 `*.docx` 跳过 `~$`/`.~` 占用锁文件
  （Word/WPS 打开节文件时落的，不是 zip）；读外部文本先去 BOM；环回请求不走系统代理。
- **脚本回显是给模型读的**：docx_toolkit 缺省一行式报错（`TENDER_TRACEBACK=1` 才出完整栈）；
  八个脚本起手都对 stdout/stderr `reconfigure(errors="replace")`（中文 Windows 的 GBK
  管道下遇 emoji 会崩），新增脚本照做。

## 发版流程（每次发版全跑）

1. 改动五个技能目录，五包 `version` 同步递增 + CHANGELOG 第 1 行加版本行 + README 徽章同步
   → `python3 listing/self_check.py` 全绿
2. 八脚本 `--help` 冒烟 OK（平级形态）；缺依赖维度已由 self_check 的 `python3 -S`
   冒烟覆盖（装机有 deps 的本机 `--help` 看不见这条）
3. 涉及评审台/产物页面 → fp 两电池全绿；涉及引擎脚本 → `python3 tests/engine_smoke.py` 全绿
4. `python3 listing/build_package.py` 重打包（版本号已递增），
   可在 `out-zips/build/` 人工检查装配树、在单包形态下复跑脚本 `--help`
5. 发版后同步 workbuddy 侧副本：`~/.workbuddy/skills/` 下五个技能目录
   （同步方式见 memory `tender-suite-hard-constraints`）

## 事实与演示口径

README 与示例中的「云澜市 / 星灵创科 / YLCG-2026-GK-0387」均为虚构演示主体；
`testruns/` 内含真实项目样本，属本机数据，不外发、不进包。
