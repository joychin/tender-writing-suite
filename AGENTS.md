# AGENTS.md — 本仓库维护须知

「灵燕标书写作套件」开源仓库：一个把招标文件写成可递交投标文件 Word 的 Agent Skill
（六技能套件），Python 3.10+，纯本地运行（所有脚本零网络请求）。
运行依赖见 `requirements.txt`（`python-docx pypdfium2 docxcompose Pillow pypdf`）；`tiktoken`/`PyYAML`
只给 self_check 用（缺省有零依赖兜底），`matplotlib` 只给甘特/架构真图用（无则降级出图），
`pypdf` 只给文档脱敏清 PDF 属性用（docx 清理纯标准库）。
本机验收另需 LibreOffice（`soffice`）：目录实测页码靠它，电池里自动转换的回归用例也要它
（E24/E29/E34/E63/E76h，缺了记失败不跳过）；`.doc` 的转换本身不是硬依赖（见「`.doc` 通道」）。

## 目录结构

- `tender-writing-suite/` — 入口技能（总控）：红线 8 条、确认门、技能路由、跨技能路径约定
- `tender-analysis/` `tender-writing/` `tender-review/` `tender-docx/` — 四个环节技能，平级安装；
  各自 `references/`（方法论 md）、`scripts/`（确定性 Python 引擎）、`templates/`
- `tender-sanitize/` — 文档脱敏技能（Word/PDF 属性体检与清理；单独请求可用，不占流水线
  步序；docx 清理纯标准库，pdf 用 pypdf；.doc 先转 .docx 后处理——脚本内置的自动转换只有
  LibreOffice，没有时探出已装 Office 并把「怎么转」交给模型选办法，LibreOffice 不是前提）
- `listing/` — 发布工具链：`self_check.py`（结构自检）、`build_package.py`（上架单包构建器）、
  `CHANGELOG.md`（版本史）；`上架指南.md`/`上架指引.docx`/`商品文案.md`/`提审说明.md` 为
  内部运营材料（gitignored，不随开源仓库分发）
- `tests/` — `engine_smoke.py`（docx_toolkit 引擎回归电池，入库版，合成夹具运行时生成；
  每条缺陷先写成失败用例再修）、`review_smoke.py`（评审台服务生命周期电池 R01–R09，
  跨平台零第三方依赖，守 Windows 探活/关停语义）、`sanitize_smoke.py`（文档脱敏引擎电池
  S01–S19，零第三方依赖合成夹具；PDF 用例缺 pypdf 记失败不跳过；S11 断言 strip_metadata
  的 `_find_soffice` 与 docx_toolkit 同口径）、`material_smoke.py`（公司资料库电池
  M01–M12；登记目录一律指临时目录，真机登记目录不进电池）
- `testruns/` — 测试夹具与验收电池（gitignored，含真实项目工作区样本）；`fp_smoke.py`/
  `fp_e2e_cli.py` 两电池及其夹具 `qszbh260186/work` 只存在于本机，新克隆的仓库跑不了
- `原型/` — 评审台页面四方向视觉原型（「方向 A 纸感公文体」已采纳实施，留存备查）
- `out-zips/` — 打包产物（gitignored）；`docs/` — README 图片

仓库源是**六目录平级开发形态**；上架产物是**单技能扁平包**（平台硬限制：Skill 仅允许两级目录，
三级嵌套如 `stages/<环节>/scripts/x.py` 会被「目录层级超限」硬拒）。

## 常用命令

```bash
python3 listing/self_check.py        # 改动技能包后必跑：frontmatter 必填与规格/examples 上限与中英一致/
                                     #   YAML/逐技能 token 闸/空目录/同技能与跨技能引用存在性/产品专属词/
                                     #   沿革从句/六包版本一致/缺依赖冒烟/模板中文字体/CHANGELOG 顶行
python3 tests/engine_smoke.py        # 引擎回归电池 E01–E106（内容保真/批注与编号/跨平台与发版闸/真机复盘/Windows 兼容/
                                     #   格式件保真红线/文档检查/公司版式读取/落盘元数据清零/落盘路径防呆/.doc 通道回退/
                                     #   版式读取遇到不规整文档/用户模板缺标题样式/超过 5MB 的 .doc 不自动转/
                                     #   LibreOffice 转换超时杀干净再提醒用户/版式读取的说法与回显/
                                     #   公司版式登记与自动采用（对话修订、往返一致、读数补洞、封面页眉落地、
                                     #   首次登记与小标题跳级与封面空段三处真机缺陷/
                                     #   E105 多册合册落公司封面/E106 删整本重合册指纹不崩））；
                                     #   改 docx_toolkit 前先补失败用例；E24/E29/E34/E63/E76h 靠 LibreOffice，缺了记失败不跳过
python3 tests/review_smoke.py        # 评审台服务生命周期电池 R01–R09（零第三方依赖，跨平台）；
                                     #   动 serve_review 的探活/关停前先补失败用例
python3 tests/sanitize_smoke.py      # 文档脱敏引擎电池 S01–S19（零第三方依赖合成夹具；PDF 用例缺
                                     #   pypdf 记失败不跳过）；动 strip_metadata 前先补失败用例
python3 tests/material_smoke.py      # 公司资料库电池 M01–M12（登记目录一律指临时目录）；
                                     #   动 material.py 前先补失败用例
python3 tender-analysis/scripts/material.py add <文件或目录> --entity <主体>   # 公司资料库入库：
                                     #   脱敏副本+解析+索引，按主体分目录（TENDER_MATERIAL_HOME 可改）
python3 tender-analysis/scripts/material.py search "关键词" --entity <主体>   # 检索章节（输出素材列记法；
                                     #   他主体命中单独点名）；list/show/forget 同族
python3 tender-docx/scripts/make_template.py   # 改 docx_toolkit 的 setup_styles 后重生成 base-template.docx
python3 tender-docx/scripts/docx_toolkit.py inspect --work work   # 文档检查（只读）：整本自身 + 与节文件的对账；
                                     #   也可直接给一份或几份 docx；改检查逻辑前先在本机真实工作区回放找漏报误报
python3 tender-docx/scripts/docx_toolkit.py layout-read "历史标书.docx"   # 版式规格读取（只读）：
                                     #   「按这个样式」场景的读取端，细则见 tender-docx/references/company-layout.md
python3 tender-docx/scripts/docx_toolkit.py layout-adopt work/版式规格.md   # 登记公司版式（用户确认后）：
                                     #   之后建节/合册不指定模板自动按它出；--set "项目=取值" 做对话修订；
                                     #   layout-show 查看、layout-forget 取消；登记目录 ~/.lingyan/tender-layout/
                                     #   （环境变量 TENDER_LAYOUT_HOME 可改——试验与用例一律指到临时目录或空目录）
python3 testruns/fp_smoke.py         # 评审台产物指纹链路冒烟（A1–A8；本机专属，见 testruns 一条）
python3 testruns/fp_e2e_cli.py       # 评审台 CLI 级全面断言（B1–B8，含 fp_smoke 12 例复跑；本机专属）
python3 listing/build_package.py     # 发版打包：产出 out-zips/tender-writing-suite-<version>.zip（唯一单包）
# 脚本冒烟：九个脚本各自 `--help`（全部 argparse 化，--help 只看说明不落盘）
# 调试：docx_toolkit 缺省只回一行式报错，加环境变量 TENDER_TRACEBACK=1 才出完整栈
```

## 硬约束（发布闸，self_check 强制）

- **SKILL.md 正文 < 5000 tokens（逐技能）**：tiktoken cl100k 实测优先，无依赖时按
  `cjk×1.55 + latin词×1.0` 估算兜底。汉字按词数会低报 ~50%，不能按字数省。
- **产品专属词**（self_check 的 `FORBIDDEN`，子串匹配）：旧内部工具名与内部词
  （`check_pipeline_state`、`ask_human`、`docx_section_` 系列等）不得出现在六个技能目录下的
  任何 md（含 `references/` 与 `templates/*.md`）。
- **沿革从句门**（`CHANGELOG_PHRASES`）：同一扫描范围禁用「此前/不再/原先/旧版/自 20xx 年起/
  历史上曾」——版本史只陈述当前版本做什么，不写对比叙事。CHANGELOG 只查顶行（新版本行）；
  闸是子串匹配，在版本行里引用禁词同样命中，改用「改前」一类说法。
- **frontmatter 纪律**：字符串值必须单行双引号（值内半角冒号+空格会被平台 YAML 解析器判嵌套映射而拒装）；
  `name` 限小写字母/数字/连字符且 ≤64，`description` ≤1024 字符；
  `examples_zh`/`examples_en` 各最多 3 条且两侧数量必须一致（平台上架解析硬拒——2.4.2 曾以
  中文 6 条/英文 4 条上传被解析失败拒装，self_check 已补该维度）；被裁掉的示例其触发词由
  `description` 承接；必填字段见 self_check 的 `REQUIRED_FIELDS`。
- **版本联动**：六个 `SKILL.md` frontmatter `version` 必须同步递增，并在 `listing/CHANGELOG.md`
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
  新增写盘代码不许绕开（create／copy-from／assemble 节回写／composer 落盘已收编），
  否则「整本被手工改过」会误报；指纹比对面前的整本不存在（删了要重新生成）谈不上
  被手工改过，跳过比对照常重建（E106，2026-11-28 真机：删带旧名的整本重合册必崩）；②保存即无痕（用户 2026-11-26 拍板「保存时不存入任何
  metadata」）：`_save_doc()` 落盘前 `_scrub_metadata()` 清零元数据——core 全字段
  （含标题/主题/关键词/类别，contentStatus 等整元素摘除，时间写当前）、app.xml 应用名/
  公司/管理者/模板名/时长、`docProps/custom.xml` 整体摘除（WPS 的 ICV 是云账号身份码）、
  包级首页缩略图；`base_document()` 与 `make_template.py` 同挂。core 直改元素树、不走
  python-docx 属性 API——docxcompose 在 import 时把全局 `nsmap["cp"]` 改成错误命名空间
  （docxcompose/utils.py 的 nsmap.update），同进程里属性 setter 会建出命名空间错误的
  重复元素、脏值藏在里面属性 API 读不到也清不掉；inspect 点名自定义属性/缩略图/单位
  信息残留（Application=Microsoft Office Word 是 Word 文件常态不算）。与 tender-sanitize
  分工：引擎管自家产物保存即无痕，sanitize 管外部文件的事后体检与清理（rsid、批注/
  修订作者名录、PDF、.doc 只有它管）；③`inspect` / `layout-read` 不改被检查（读取）的文件，
  **结果落盘参数是路径不是开关**：`inspect --json`、`layout-read --json` 与 `--out`、
  `check_writing --json` 后面跟的是结果要写入的文件（`.json` / `.md`），写之前一律过
  `guard_dump_path`（docx_toolkit）——落到本次要读的文件上、或后缀不是它写出的格式即拒
  （退出码 2，报错点破语义并给正确写法），且必须在读取与校验**之前**调，拦下时什么都
  还没做；新增带落盘参数的命令照此过闸。`inspect` 没给要检查的文件同样报参数错（检查
  0 份文件再报「未发现问题」是最危险的那种没查）。真机事故（2026-11-26）：模型把
  `inspect --json 标书.docx` 当「按 JSON 输出」的开关，文件参数被吃掉，一份真实标书被写成
  63 字节的报告。`strip_metadata` 的 `--json` 是「把 JSON 报告打印出来」的开关，与上面三个
  不是一回事。阈值是拿本机真实工作区回放标定的（图片 72 DPI、拉伸 5%、图与表超宽容差
  0.5 厘米、窄列须有 4 字以上实际内容、`image` 按页宽减左右边距插图），改之前先回放。
  动它之前先看 `tests/engine_smoke.py`（E41–E54、E72–E75）。
- **`.doc` 通道：LibreOffice 不是前提**（用户 2026-11-27 拍板：体积大、不少用户装不了；
  脱敏、解析、版式读取三处通道同日一并改）：三份脚本（`strip_metadata`／`parse_document`／
  `docx_toolkit` 的 layout-read）内置的自动转换都只有 LibreOffice（本机唯一实测过的通道）；
  没有它时 `_installed_office()` 探出已装的 Office 套件，`_hand_over_hint(why, next_step)`
  把「转换交给你选办法＋转好之后怎么接」回显后退出（退出码 2），提示里不劝人安装。
  **选哪个工具转换归模型，后续仍归脚本**；不内置没在真机验证过的 Word/WPS 自动化
  （WPS for Mac 实测没有脚本接口；Windows 分支只在 mac 上模拟）。接力写法各不同：脱敏与
  版式读取对转出的 `.docx` 重跑同一条命令；解析多一个 `--converted <那份 .docx>`——解析
  目录按源文件全名命名、来源登记记的是原 `.doc` 的名字与 sha256、后续按行号拷贝格式件要
  回解析目录取 `_converted.docx`，直接去解析那份 `.docx` 会让这些全对不上，所以由解析器
  把交来的文件原样落到同一位置（真实回放：两条路径的 md／outline／elements 逐字节一致）。
  解析另有一条预判：已解析且产物完好、`_converted.docx` 在，没有转换器也不带 `--converted`
  照样沿用。**超过 5MB 的 `.doc` 三处都不自动转换**（用户 2026-11-27 定「超过 5M，提醒
  用户自己转换」；`_DOC_AUTO_LIMIT`／`_oversize_doc_hint`／`_doc_too_big` 同样三份逐字
  同口径，E79f 钉住）：有 LibreOffice 也直接回显并退出（退出码 2），不请模型另想办法，
  请用户在 Word/WPS 里另存；解析带 `--converted` 或已有完好转换件时照常走；版式读取
  多份文件时读之前整体拦下、不留半截输出。5MB 是用户定的值、不是标定值——本机 9 份
  真实 `.doc` 实测（LibreOffice 26.2）：5.4／8.5／12.4／23.6MB 的都是 2–3 秒转完，
  4.4MB 的 38 秒，只有一份 26.7MB 的 28 分钟没转出来，卡住的是内容不是体积。
  两份转换件若来自 LibreOffice 的两次独立转换，字节不同是正常的（部件内容全同，
  只差 zip 条目时间戳）——别拿它当回归。**转换超时就杀干净、再提醒用户**（用户 2026-11-27
  定；`_LO_TIMEOUT`／`_LoTimeout`／`_kill_tree`／`_run_lo`／`_stuck_doc_hint` 同样三份逐字
  同口径，E80、S19 钉住）：起 LibreOffice 一律走 `_run_lo`——soffice 自成进程组，超过 300
  秒（经验值）整组 SIGKILL，Windows 走 `taskkill /T /F`；不要再用 `subprocess.run(timeout=)`
  （它只杀直接子进程，Linux／Windows 上 soffice 外壳拉起的 soffice.bin 会留下继续空转——
  E80 改前红的三条就是「残留进程」）；超时、起不来、没产出文件走同一个出口，只请用户
  自己另存、不请模型另想办法。那份 26.7MB 的 `.doc` 转不出的真因即在此：9 个图片／绘图
  对象的尺寸数据不规范，LibreOffice 把 8 个的宽度读成 0.8–19.4 米、1 个单张就卡死，排版
  陷进死循环（`sample` 抓栈全在 `CalcLayout → FormatFlyContent`）。目录页码实测
  （`toc-pages`，含预览 PDF）也靠 LibreOffice，没有等价替代：同走 `_run_lo` 杀得干净，
  超时文案不变（可跳过本步稍后再试）。动这几处前先看 `tests/engine_smoke.py`
  （E76、E79、E80）与 `tests/sanitize_smoke.py`（S17–S19）。
- **文档脱敏只动元数据**（`strip_metadata`；细则 tender-sanitize/references/metadata-fields.md）：
  体检（`--check`，退出码 1=有发现可当验收闸）→ 清理出副本（缺省 `<原名>_cleaned<扩展名>`，
  原件一个字节不动；`--in-place` 须用户对话明说）→ 两道自证（正文指纹 + 复扫）。
  docx 走 zip 内 XML 定点改写（python-docx 只暴露 core.xml，碰不到 app/custom/缩略图/
  修订批注作者/rsid）；pdf 用 pypdf 克隆重写（引 PyMuPDF 之前先想许可证——AGPL）；
  .doc 先转 .docx 再处理（输出一律 .docx，不支持就地覆盖；转换办法见上一条「.doc 通道」，
  接力路径与自动转换的结果正文指纹一致）。
  副本输出天然避开整本指纹机制；对整本 `--in-place` 会让 `.fingerprint.json` 失配
  （references 有条目）。动它之前先看 `tests/sanitize_smoke.py`（S01–S19）。
- **公司资料库与素材复用**（`tender-analysis/scripts/material.py`；2026-11-28 设计讨论拍板，
  用户真机教训：skill 会检索到不属于本次投标主体的资料）：技能无状态，库在用户目录
  `~/.lingyan/tender-materials/`（`TENDER_MATERIAL_HOME` 可改——电池与试验一律指临时
  目录）；**按主体分目录**，检索缺省限定本次主体（承诺清单「投标主体」行，G2 收尾与
  资格门槛同框必答，免问三条件见 analysis-lite），命中其他主体的素材单独点名确认
  （引用性质还是误命中），主体悬空不检索。入库=strip_metadata 出脱敏副本存库（原件
  一个字节不动；**经子进程调 tender-sanitize，不复制其函数**，`.doc` 通道/超 5MB/超时
  杀树全走 strip 的现成出口）→ parse_document 解析**库内副本**（行号与库内文件永远
  一致，copy-from 直接以库内文件为源并带 `--parse` 指库内解析目录）。索引由脚本确定性
  生成（章节树/旧名候选/体检摘要），语义匹配归模型；旧名候选只做项目名称/招标编号/
  机构名标签行与字母开头编号模式（人名同名率高不做，用户拍板）。三道闸与检查：
  **素材新鲜度闸**（copy-from，E100–E102）——拷前比对源文件与解析登记的指纹，不一致
  拒绝（退出码 2）并指引重解析+重核指引素材列的行号区间（素材文件是活的，旧区间拷错
  段无声进标书），无指纹登记的旧产物放行但提醒；**登记库旧名自动载入**（check_writing
  全局校验，E104）——其他主体的登记名与旧名候选进残留扫描当疑似级线索（不拦；显式
  `--old-names` 仍是红线级，E103），短名是长名子串时丢短名防双重点名，主体悬空不载入
  并明说；**素材指标符合性**（section-writing 第 4 步纪律）——素材数字承诺 vs 招标硬
  指标核对，对不上批注点名提醒用户是否修订，不擅自改（机械扫描待回放标定，第二批）。
  动它之前先看 `tests/material_smoke.py`（M01–M12）与 `tests/engine_smoke.py`
  （E100–E104）。
- **公司版式读取只读**（`docx_toolkit layout-read`；细则 tender-docx/references/company-layout.md）：
  「按这个样式」场景的读取端，本版只到读出规格表（按表出稿、登记默认、对话修订在
  后续批次）。读数一律按生效格式（直写→样式链→docDefaults→主题）；页面与页码取
  **正文所在节**（真实标书五节边距各不相同，读第 1 节会把封面节的数当成正文的）；
  封面止于第一个分页或分节符、目录条目认五种形态、两区之间的前置页单列，三区都不
  进正文统计，表格与图也只统计正文区；标题候选=命名标题样式或自带编号的短段（不超
  40 字、不以句末标点收尾、加粗或不超 20 字——手打编号的条款行不算）；文本框文字
  只读 mc:Choice（兼容分支里那份是重复）；确定线只有一条——两种以上并存时最高占比
  过 75% 才记确定（经验值不是标定值，只看占比不看样本量）。改读取逻辑前先回放
  `testruns/company-samples/` 与真实招标文件，动它之前先看 `tests/engine_smoke.py`
  （E55–E66、E77、E81–E85）。**2.3.0 的教训**：首版凭合成夹具报了「读全」，真实标书回放坐实 8 处
  读数错误——读取类功能没有真实样本回放不算做完。**2026-11-27 的教训**：WorkBuddy 真机
  读三份真实标书两份整份崩（本机 12 份真实标书里 4 份没有一级标题——章挂在「标题 2」
  上，`ones[0]` 取空列表；别的转换器转出的 docx 节里不写页面设置，取属性取到 None），
  模型就地给安装目录里的脚本打补丁绕过。现在：没有一级标题退到最浅一级读并记疑似；
  没写页面设置取 Word 缺省值并记疑似；某一组读不出来记一行「本组读取中断」不拖垮
  整份（TENDER_TRACEBACK=1 照旧抛）；包里没有样式表的「拼出来」docx 首行紧跟 `[注意]`
  点破。两个只有真实回放才抓得到的坑：①`_n_styles` 必须在读段落样式**之前**数——
  python-docx 一碰 `doc.styles` 就给没有样式部件的文档就地补一份默认的，合成夹具
  （空样式表）测不出「没有部件」这个形态；②只认 `/word/styles.xml`，同目录的
  `stylesWithEffects.xml` 是副本。**2026-11-27 晚的教训（读数对、说法错、模型越界）**：
  WorkBuddy 真机读一份全篇不挂样式的技术标，读数逐项对（独立探针核过），但脚本自己
  说错、说漏的四句话把模型带偏——「标题全是直接格式」是写死的断言（事实是外观取自
  正文默认样式，那份标书的编制说明明写「统一字体与统一字号」）；119 节逐节罗列到 400
  字拦腰截断（第 118 节页边距 2.5 厘米被截掉）；页码依据写成「正文所在第 1 节」（实为
  第 103 节沿用第 1 节）；目录缩进只看段落缩进（143 条二级条目靠文字里的 2 个全角
  空格起头）。模型据此把用户的版式说成「做得差」，自创黑体标题推荐成默认、写进用户
  长期记忆，并越过「本版只到读出规格表」手工搭了两份模板。现在：事实行逐段量
  （`_unstyled_head_note`）；页面差异同种归组、行里存完整取值、排表按「；」整段让位
  并写明另有几段（`_page_diffs`／`_md_val`，完整取值在 --json 结果里）；页码依据写
  沿用关系；目录缩进连条目文字开头的空格一起读（`_lead_spaces`）；**回显末尾固定
  `[边界]`、`[取值]` 两行**——SKILL.md 里加粗的边界拦不住只看回显最后几行的模型，
  纪律要落在模型每次都会读的地方；规格表里的说法必须是量出来的，写死的断言比不写
  更糟。验收口径：19 份真实文档副本改前改后回放逐份 diff（数字一个没变，新冒出来的
  取值用独立探针逐项核）。**已知缺口**：章标题是「正文样式 + 自动编号
  （第%1章，编号不在文字里）」时进不了标题候选（套件自己合册的整本就是这形态），
  标题组缺「标题一级」行——补这条候选规则要走真实样本标定。
- **公司版式登记与自动采用**（`docx_toolkit layout-adopt`／`layout-show`／`layout-forget`；
  细则 tender-docx/references/company-layout.md；用户 2026-11-27 拍板：「正常情况下，
  用户只会通过对话和 agent 表达使用模板的意图」——真机上套件到读出规格表就停了，模型
  让用户「建节和合册时传入这个参数」，当晚把原排在下一批的登记与采用提前做完）。
  **规格表是公司版式的本体**：确认后登记到用户目录 `~/.lingyan/tender-layout/`
  （`layout.md`＋`layout.json`＋`历史/`；环境变量 `TENDER_LAYOUT_HOME` 可改，ZCode 与
  WorkBuddy 两个宿主共用同一份）。登记后 `create`／`assemble`／`image` 不指定模板就
  自动按它出（回显首行 `[版式]`）；模板每次出稿时由 `apply_layout` 从当前的内置基准
  模板现编译，引擎升级后登记的版式不会卡在旧模板上。机制与口径：
  - `activate_layout(args)` 在 `parse_args()`／`require_docx()` 之后、写节闸之前跑，只对
    出文档的命令生效（新增出文档的命令要进它的名单）；显式 `--template` 或
    `--ignore-layout` 一律回到内置口径。进程内全局 `_LAYOUT`，引擎各处用
    `_lay("组", "项", default=内置值)` 取值——**没有登记时行为必须与登记前逐项相同**，
    老用例就是阴性对照（电池里 `tk()` 缺省把登记目录指到空目录，既不读本机真实的
    登记也不往里写）。
  - 跟版式走的项（没登记全是内置口径）：正文首行缩进（`_set_body_indent`，「无」也显式
    写零——合册收尾只给没有 `w:ind` 的段补缩进）；表格口径 `_lay_table()`（字体字号、
    表头加粗底纹对齐、框线 `_tbl_borders_el`、表题在表上方还是下方）；图题走新样式
    `Figure Caption`（表题与图题可以不一样）；页码 `_page_field`／`enable_footer_pagenum`
    （纯数字／第 X 页／— X —／共 Y 页、页脚或页眉、位置、字号；封面/前置/目录不编号、
    正文自 1 起的口径不变）；目录标题 `_append_toc_marker`、条目 `_fill_toc`（引导线、
    制表位落在版心右缘）；编号方案与间隔 `_num_gap`、章起新页 `_chap_new_page`（首章前
    一律分页）、目录层级、文档网格、对称页边距、图宽比例；`--numbering`／`--toc-depth`
    缺省 None，跟登记的版式，用户或招标文件明说了的优先。
  - **规格表的解析是读数写法的逆**（`parse_layout_spec` 与 `_p_*`）：读数新增或改一种
    写法，必须同步解析器与 `_ITEM_GROUP`（对话修订补行要知道项目属于哪一组）。
    几种写法并存（疑似行）取排在最前的那种；历史标书里没有的项（未发现表格、未见
    表题）按内置写法出。改了读数或解析器，**E90 往返一致是闸**：按登记的版式出样张 →
    `layout-read` 读回 → 与登记表逐项相同。
  - 对话修订：`layout-adopt … --set "项目=取值"`（写法与规格表一致，写错拒收并给这个
    项目的写法样例）；直接改 `layout.md` 也行——`load_layout` 按指纹发现改过就重读，
    改坏了沿用上一次读得出来的取值并点名。
  - 三条踩过的坑：①往样式里写属性前把 rPr／pPr 整体重建（`_rebuild`）——基准模板的
    标题样式带着主题字体与加粗、链接的字符样式带着主题蓝，只改几项等于没改；②同一个
    lxml 元素对象只能挂在一处——一行的对齐元素建了一个往每个单元格里追加，被逐格搬走，
    最后只剩最后一格有；③段距写「N 磅」时别顺手写「N 行＝0」——读数先看行数，会把
    「0 磅」读成「0 行」。另：基准模板的文档默认值是段后 10 磅、行距 1.15 倍（Word 英文
    模板的常态），版式里「没设」的那一侧会被它顶成 10 磅，`apply_layout` 先把它清零。
    **2026-11-28 真机（WorkBuddy 首次真实登记）补三条，全是电池盖不住的「从零开始」形态**：
    ④登记目录不存在——`_extract_layout_assets` 先 `mkdir(parents=True)`（E97）；登记命令
    先写记录再提取，提取失败只警告不回滚（退出码 0），报错话术不许谎称没发生的事
    （改前 die 的话术声称「规格表本身已登记生效」而实际还没写）；⑤节文件小标题跳级
    （直接 h3、没经过 h2）——`_restyle_all` 对缺的中间层级 `setdefault(k, 1)` 续上，
    改前 KeyError（E98）；⑥封面提取不过滤空段——空段就是封面的竖向排版，滤掉后
    封面内容全挤到页顶（E99）。这条也是流程教训：**电池的登记目录总是先建好，真机是
    从零开始的——「首次」形态要专门造（登记目录不存在、目录树跳级这类），不然真机
    第一跑就崩，模型又会就地打补丁。**
  - **封面整页与页眉（2.4.1 接上）**：`layout-adopt --source <历史标书>` 提取两类资产存进
    登记目录——封面＝`封面.docx`（`_split_regions` 取封面区元素，`_clone_with_images`
    元素级拷贝＋图片双通道重挂＋痕迹清理＋样式随迁，末件削掉分节符与分页符——整本里
    封面后面本来就另起页）；页眉＝`媒体/页眉-N.<ext>`（`_header_pick` 挑最有内容的节页眉，
    优先带 logo 的；文字字号对齐横线与图片的 wp:extent 原尺寸一起存进 layout.json 的
    assets）。不带 `--source` 沿用旧资产、从没有过则列进「还没接上」（动态清单
    `_layout_pending`，页眉公司本来就没有——规格表页眉行=无——不算没接上）。合册落地：
    封面节点无节文件时 `_materialize_cover` 落公司封面为 body/封面.docx（走统一落盘口）
    并点名待改字段（`values.cover_fields`，来自规格表封面字段行），已有封面节文件则不碰；
    `_apply_registered_header` 把 logo＋文字＋横线落到正文节页眉（页码与页眉共存时页码
    另起一段）；**`_repoint_body_refs`**——docxcompose 会把正文节默认页眉/页脚引用挪到节
    文件带入的部件上（真实回放：公司封面合册后正文页码指向封面带进来的空页脚，页码整本
    消失），收尾把默认引用拨回主文档自己的部件（主文档部件名在 compose 前捕获）。单元格
    边距跟版式（`_p_cellmar`，读到非缺省 0.19 厘米才显式写，宽表紧凑口径让位）。
    已知小缺口：封面页沿用正文页边距，公司封面自带绝对定位文本框且页边距不同时会有挪位
    （真实样张已见），下一小批把封面页边距接上。E93–E99 钉住封面/页眉/单元格边距与
    首次登记、编号跳级、封面空段三处真机缺陷；E105 钉住多册合册给没有封面节的册落
    公司封面（2026-11-28 真机：多册形态把「册名/封面.docx」写成字符串相除，首次多册
    合册必 TypeError 崩）。
  - **验收口径**：两份真实公司标书（`testruns/company-samples/`，结构化形态与纯文字形态
    各一）→ 登记（带 `--source`）→ 出样张（含封面节点）→ LibreOffice 渲染看图（封面、
    页眉、页码）→ 读回规格表与登记表逐项比（疑似行取排在前面的、层级类项目除外；
    2.4.1 实测 26/26 与 29/29）。缩略图上「看着加粗」不可信（WorkBuddy 那个模型就被
    它带偏了 12 轮），加粗与否看 PDF 里的字体名。本机真实的登记目录别拿来当试验场——
    验收用 `TENDER_LAYOUT_HOME` 指到临时目录（⚠ 命令行上每个命令的内联环境变量只管
    自己那一条，2026-11-27 深夜曾被这个坑把测试登记写进真实目录）。
  - **这一版还没接上**（登记回显明说，不许含糊成「都按这个样式」）：表格列宽——要转
    历史标书的列宽习惯，另批做；招标给了列宽的格式件照旧只拷不改。
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
  `setup_styles` 生成的，改样式定义后必须重跑；内置标题族的定义只有 `_HEADING_SPECS`
  一份（基准模板与用户模板缺样式时的补建共用）。**用户模板缺 Normal／Heading 1-4 时按
  内置定义补建并回显 `[模板标题样式]`**（`_ensure_builtin_styles`，E78）——只补空壳的话
  章标题跟着正文走、还不吭声，拿没有标题样式的历史标书当模板正是这个情形。中文字体
  跑偏在没装日文字体的开发机上渲染会静默回退，肉眼验收看不见，只认 self_check 的模板字体闸。LibreOffice 查找、`.doc` 回退
  与超时处置各有多组三份同口径的函数——`_find_soffice`；`_installed_office`／`_hand_over_hint`；
  `_DOC_AUTO_LIMIT`／`_oversize_doc_hint`／`_doc_too_big`；`_LO_TIMEOUT`／`_LoTimeout`／`_kill_tree`／
  `_run_lo`／`_stuck_doc_hint`（本文件、
  `parse_document.py` 与 `tender-sanitize/scripts/strip_metadata.py`，环节技能各自独立安装
  故不共享模块），改一处同步三处（E35 对 docx_toolkit 与 parse_document 的 `_find_soffice`
  两份断言同一结果，sanitize_smoke S11 对 strip_metadata 这份断言；E76g 对回退那两个函数的
  三份断言逐字一致，E79f 对大文件提示与大小线、E80 对超时处置同样三份断言一致）。
  **新写调用 LibreOffice 的代码一律走 `_run_lo`**，不要再用 `subprocess.run(timeout=)` 起
  soffice；模拟 LibreOffice 的用例顶替 `_run_lo`，别顶替 `subprocess.run`（E36、S18 改过）。
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
  九个脚本起手都对 stdout/stderr `reconfigure(errors="replace")`（中文 Windows 的 GBK
  管道下遇 emoji 会崩），新增脚本照做。
- **Windows 兼容清单**（开发机是 macOS，这些差异本地跑不出来；真机是 workbuddy on
  Windows——bash 风格 shell、宿主自带后台运行、模型用 `python` 起脚本。新增子进程、
  临时目录、外部文件读取、`body/` 枚举时逐条对照；对应用例靠 mock 在 macOS 上复现）：
  外部程序不假设在 PATH 里（LibreOffice 安装器不写 PATH，查约定安装位置与注册表）；
  文件地址用 `Path.as_uri()` 生成（手工拼 `file://`+路径 在 Windows 是坏地址）；子进程
  输出解码带 `errors="replace"`；临时目录里的句柄用完即关，清理失败不得盖掉真正的报错
  （`docx_toolkit._tempdir()`）；枚举 `body/` 下 `*.docx` 跳过 `~$`/`.~` 占用锁文件
  （Word/WPS 打开节文件时落的，不是 zip）；读外部文本先去 BOM；环回请求不走系统代理。

## 发版流程（每次发版全跑）

1. 改动六个技能目录，六包 `version` 同步递增 + CHANGELOG 第 1 行加版本行 + README 徽章同步
   → `python3 listing/self_check.py` 全绿
2. 十脚本 `--help` 冒烟 OK（平级形态）；缺依赖维度已由 self_check 的 `python3 -S`
   冒烟覆盖（装机有 deps 的本机 `--help` 看不见这条）
3. 涉及评审台/产物页面 → fp 两电池全绿；涉及引擎脚本 → `python3 tests/engine_smoke.py` /
   `python3 tests/sanitize_smoke.py` 全绿
4. `python3 listing/build_package.py` 重打包（版本号已递增），
   可在 `out-zips/build/` 人工检查装配树、在单包形态下复跑脚本 `--help`
5. 发版后同步两侧技能副本：`~/.workbuddy/skills/` 与 `~/.zcode/skills/` 下六个技能目录，
   逐目录 `rsync -a --delete --exclude __pycache__ --exclude .DS_Store`。`--delete` 会删掉
   目标里有、仓库没有的文件——先 `--dry-run -i` 看清删除项再真跑；同步后校验两侧 version、
   文件清单与逐字节内容，并在安装位置复跑十脚本 `--help` 与一道防呆。本机 rsync 是
   openrsync：`-i` 对新文件标 `>f.......`（点，不是经典 rsync 的加号），新目录看父目录的
   `cd+++++++`。两侧的其它技能目录（如 `remove-metadata`）不属于套件，不要碰。

## 事实与演示口径

README 与示例中的「云澜市 / 星灵创科 / YLCG-2026-GK-0387」均为虚构演示主体；
`testruns/` 内含真实项目样本，属本机数据，不外发、不进包。
