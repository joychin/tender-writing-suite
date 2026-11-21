---
name: tender-writing
display_name: 灵燕标书·写作与校验
display_name_en: Bid Writing & Validation
description: "标书写作环节：按写作指引逐节生成正文（技术标/商务标/资格文件），维护写作指引与关键事实承诺清单，全局校验承诺数字、章节查重与旧名残留。适用于：开始写正文、写技术标、写商务标、接着写、继续写、逐节写、写作指引、承诺清单、全文校验。用户要求整本标书时由入口技能路由到本技能。"
description_zh: "标书写作环节：把目录树写成一节节可交付的 Word——史料段表格按原件拷贝、推理节按模式重写，承诺数字只取拍板值、缺的挂批注；写完机器校验承诺对账、查重与旧名残留。"
description_en: "Bid writing stage: generate section docx files per the writing guide (technical/commercial/qualification volumes), maintain the writing guide and approved-commitments list, and run global validation on commitment figures, duplication and residual old names."
examples_zh:
  - "目录确认了，开始逐节写正文"
  - "接着写技术标的下一节"
  - "跑一遍全局校验，看承诺数字对不对"
examples_en:
  - "The outline is approved — start writing the sections"
  - "Continue with the next technical section"
  - "Run the global validation on commitments and duplication"
version: 2.0.2
author: 灵燕智能
---

# 写作与校验（第四步–第六步）

开始前先载入入口技能 tender-writing-suite：红线、确认门、工作目录与回复口径
以它为准；本文件只讲本环节做法。写节前置条件：`work/review/评审台.html` 已存在
且 G5 拍板卡已呈现过——缺任一即门被跳，回对应门补「建/刷新确认面 → 发指针 →
停轮」再开写（见 @tender-writing-suite/references/stage-gates.md 写节前置自检条）。

## 上游检查

- `work/outline/tender-response-docs.md` 缺失 → 载入 tender-analysis 先做目录树；
- `work/body/写作指引.md` 缺失 → 从第四步开始；
- 指引已存在 → 按指引续写，只补未完成节。

## 第四步：写作准备（先指引、后承诺；每份停轮汇报一次）

1. 生成 `work/body/写作指引.md`（格式契约 @references/guide-format.md）：逐节判
   写作模式（素材修订/格式跟随/推理撰写），可用素材与公司事实查清填表——检索
   通道以环境实际的知识库/资料库工具为准，没有则向用户索取（**查到不
   写进指引=白查**）。跑 `python3 scripts/check_writing.py --guide` 过 G4 门：
   刷新评审台（`python3 tender-review/scripts/build_review.py --work work`）打开
   指引 tab；承诺事项一次列全（主体/工期/人员/服务承诺/报价口径；清单已存在
   则并入，只列缺项），问答阶段已确认的直接沿用，不重复问。
2. 用户答复后落 `work/body/关键事实与承诺.md`（事项｜值｜说明；值只取用户答复
   与分析产物里的事实，缺的标待补）。**值列写可独立检索的裸值**（人名/数字/
   期限，如「12 个月」），解释放说明列——校验按值列做全文子串比对。过 G5 门：
   刷新评审台打开承诺 tab（逐项拍板卡在页面），对话指针=承诺值一行汇总；未拍板
   的值落「待补」。

## 第五步：逐节生成（按指引模式路由）

首节开写前**首开写作进度面板**（`work/review/写作进度.html`——监视面：进度/
批注台账/待拍板承诺；服务模式自刷新，静态模式每节重跑
`tender-review/scripts/build_review.py`）——**未首开禁写第一节**（与入口红线
「评审台未首建禁建节」同构）。**进入本步必读
@references/section-writing.md**：三模式细则、工具链速查（docx_toolkit 在
`tender-docx/scripts/`）、素材三级落地、语言纪律都在其中。核心路由：素材修订 =
copy-from 注入素材授权区间（保真）→ revise 适配；格式跟随 = copy-from 拷格式件
区间 → revise 填空（未定值批注）；推理撰写 = create 的 body 块一次成形（表格用
table 块，不流水句铺）。每节写完跑 `check_writing.py --section "<节.docx>"` 自查；
多节分派跑 `python3 scripts/build_dispatch.py --work work --sections "节A,节B"`
生成批次提示词（规则见细则）。

## 第六步：全局校验

全部节完成后过 G6 门：刷新评审台与进度面板（待办卡反映未写节；批注台账与校验
结果全量在进度面板——对话只给结论级汇报），**校验跑完直接进第七步合册**——
合册代价小（秒级重跑）、用户调整都在整本层做，**禁止问「现在合册还是先补齐」
类岔口**；补齐清单进进度页与收尾汇报，由用户拿到整本后处理（issue 未清零也
合册，收尾逐条点名）：

```bash
python3 scripts/check_writing.py            # 承诺比对 + 节间查重 + 残留扫描
python3 scripts/check_writing.py --old-names "旧公司名,旧项目名"   # 素材改写后必查残留
python3 scripts/check_writing.py --json work/review/check-result.json   # 落盘供进度面板读，收尾必带
```

- **旧名残留必须清零**（行业第一大事故）：`--old-names` 把用户提过的旧机构名
  都传进去（素材来源项目的招标方、我方曾用名等），别只靠素材文件名兜底。

## 链尾 HARD-GATE

G6 交付并停轮收指针后，进入第七步合册：只载入 tender-docx 一个环节技能。
