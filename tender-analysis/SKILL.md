---
name: tender-analysis
display_name: 灵燕标书·解析分析
display_name_en: Bid Analysis (Parse & Outline)
description: "分析招标文件：解析全文、提取要点与登记表、全扫废标条款、梳理评分办法与资格门槛、登记澄清清单、规划响应文件与投标目录。适用于：分析标书、解读招标文件、查废标风险、梳理评分办法、资格门槛、投标目录、要点提取、投标要准备哪些材料。用户要写整本标书时由入口技能路由到本技能做前置分析。"
description_zh: "标书分析环节：把招标文件读成可执行的作战地图——废标条款全扫、评分办法与资格门槛成表、澄清清单逐条带出处、响应文件与目录树成形。只分析不写正文，分析结论可被写作环节直接承接。"
description_en: "Tender analysis stage: parse the full tender document, extract key points into registry tables, scan all disqualification clauses, map scoring criteria and eligibility thresholds, register clarifications with provenance, and plan response volumes and the bid outline. Analysis only — no body writing."
examples_zh:
  - "帮我分析这份招标文件，废标风险有哪些"
  - "把评分办法和资格门槛整理成表"
  - "规划这个项目的投标目录"
examples_en:
  - "Analyze this tender document and list the disqualification risks"
  - "Turn the scoring criteria and eligibility thresholds into tables"
  - "Plan the bid response outline for this project"
version: 2.1.5
author: 灵燕智能
---

# 解析、分析与目录（第零步–第三步）

开始前先载入入口技能 tender-writing-suite：红线、确认门、工作目录与回复口径
以它为准；本文件只讲本环节做法。

## 上游检查

产物落在当前工作目录 `work/`：

- `work/parse/sources.json` 缺失 → 从第一步开始；
- 已存在 → 解析已完成，从要点提取续接；
- `work/outline/tender-response-docs.md` 已存在 → 本环节已完成，载入
  tender-writing 进入写作准备。

## 第零步：单轮问答与体检（不进建书流程）

用户只要答案、不要标书（废标风险/资格门槛/评分办法/要备哪些材料）时走本步：
解析（同第一步，含来源确认）→ 读 @tender-writing-suite/references/evidence-rules.md
→ 按 @references/analysis-lite.md 出本题要用的登记表 → 同一命令产出只读分析报告
（分析阶段唯一交付面，见 @tender-review/references/review-workbench.md）→ 逐条带
出处直接回答，末尾点明已检查范围，提示「要生成完整投标文件，回复开始写标书」。
不做目录、不写正文、不合册；已出登记表全流程沿用，缺的第二步补齐。涉公司家底
按家底确认纪律先确认再答；判断不了再问一句。

## 第一步：解析招标文件

1. **来源确认**：多份文件先问清哪份是**主文件**（结构骨架基准）、其余各是什么
   角色（补遗/澄清/附录/无关）；不明确按最可能者先行并声明假设；同名不同扩展
   以 docx 为准。结果写 `work/parse/sources.json`（main / supplements /
   excluded）。
2. **并行解析**：同一条消息里对每份来源文件并行发
   `python3 scripts/parse_document.py "<文件路径>"`（幂等）。扫描件返回
   `[解析转写待办]` → 按 @references/transcribe-scans.md 誊录后重跑（超页数上限
   先停轮确认）。主文件失败→说清原因停止；补充文件失败→声明降级不可引用。
   素材文件（历史标书/公司资料）为 docx 的也在此一并解析。
3. **G1 门（异常门——仅异常时停轮）**：解析无警示且来源角色已决 → **不产报告
   不停轮**，直接进第二步（解析概况并入 G2 的报告与指针交付）；有解析警示
   （识别降级/可信度低/扫描件待誊录）或来源角色未决 → 待确认事项登记
   `work/analysis/clarifications.md`（编号与第二步续排），生成分析报告
   （`python3 tender-review/scripts/build_review.py --work work --pages report`）
   交付即停——概况与警示全量在报告，对话只留指针。

## 第二步：要点提取

读 @tender-writing-suite/references/evidence-rules.md（执行前必读），按
@references/analysis-lite.md 产出四张登记表 + 废标全扫 + 澄清清单；收尾跑
`python3 tender-writing/scripts/check_writing.py --analysis` 校验
（coverage 与出处引用；`--work` 默认 `work`）。定位纪律：行号区间或关键词检索，**行号区间精读取证，
禁止整读解析 md**。过 G2 门：刷新分析报告（正常路径的首次报告交付——第一节
概况表即解析结论）；对话指针开头带解析概况一句（主文件/规模/识别方式），关键
要点（废标最严几条、评分框架、
资格硬门槛）在对话给结论级摘要（只述事实与出处，不作商业评价），六张分析表与
澄清清单全量只在报告；澄清处理用 `python3 tender-review/scripts/apply_review.py`
记录；未处理项默认先行（批注标注）。追问本公司得分/资料缺口按家底确认纪律
先确认再答。

## 第三步：目录树

1. **响应文件分解**（@references/response-files.md）：划分结论（拟分 N 册及册
   名）必报——方案唯一也报；多方案或证据不明**停下来问一次**；未明确
   选择按最合理方案先行并声明理由。
2. **目录初稿**（@references/outline-generate.md + @references/outline-annotation.md）：
   每响应文件补全 `## 目录`、`## 来源标注`、`## 目录说明` 三段；来源 ID 只引用
   analysis 已有编号；树首固定 `- 封面`，其后 `- 目录`（合册机械生成）。
3. **三道清理**（@references/outline-cleanup.md，①查漏补缺 → ②评分对齐 →
   ③走查定稿；无对应输入跳过并声明）；清理后跑
   `python3 tender-writing/scripts/check_writing.py --outline`
   对账（悬空/未引用来源 ID）。
4. **G3 门（评审台首建）**：`python3 tender-review/scripts/build_review.py --work work`
   （不带 `--pages report`），打开目录 tab 停轮等确认；用户改目录落盘后按增量
   走查回流（页面规程见 @tender-review/references/review-workbench.md）。

**树格式红线**：每行 `- ` 开头、无编号前缀、每级缩进 2 空格、独立附件号条目
平级、不用代码块包裹——违反的行不会被合册解析。

## 链尾 HARD-GATE

G3 交付并停轮收指针后，进入第四步写作准备：只载入 tender-writing 一个环节
技能。用户明确只要分析/体检（不要标书）时，第零步交付即止，不载入下一环。
