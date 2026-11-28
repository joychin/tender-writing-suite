---
name: tender-writing
display_name: 灵燕标书·写作与校验
display_name_en: Bid Writing & Validation
description: "标书写作环节：按写作指引逐节生成正文（技术标/商务标/资格文件），维护写作指引与关键事实承诺清单，全局校验承诺数字、章节查重与旧名残留。适用于：开始写正文、写技术标、写商务标、接着写、继续写、逐节写、写作指引、承诺清单、全文校验。用户要求整本标书时由入口技能路由到本技能。"
description_zh: "标书写作环节：把目录树写成一节节可交付的 Word——史料段表格按原件拷贝、推理节按模式重写，承诺数字只取拍板值、缺的挂批注；写完自动校验承诺对账、查重与旧名残留。"
description_en: "Bid writing stage: generate section docx files per the writing guide (technical/commercial/qualification volumes), maintain the writing guide and approved-commitments list, and run global validation on commitment figures, duplication and residual old names."
examples_zh:
  - "目录确认了，开始逐节写正文"
  - "接着写技术标的下一节"
  - "检查一下承诺数字和重复内容对不对"
examples_en:
  - "The outline is approved — start writing the sections"
  - "Continue with the next technical section"
  - "Check the commitment figures and duplicated content"
version: 2.5.0
author: 灵燕智能
---

# 写作与校验（第四步–第六步）

开始前先载入入口技能 tender-writing-suite：红线、确认门、工作目录与回复口径
以它为准；本文件只讲本环节做法。写节前置条件：`work/review/评审台.html` 已存在
且承诺卡已呈现过（G4/G5 合并门已停轮交付）——缺任一即门被跳，回对应门补
「建/刷新确认面 → 发指针 →
停轮」再开写（见 @tender-writing-suite/references/stage-gates.md 写节前置自检条）。

## 上游检查

- `work/outline/tender-response-docs.md` 缺失 → 载入 tender-analysis 先做目录树；
- `work/body/写作指引.md` 缺失 → 从第四步开始；
- 指引已存在 → 按指引续写，只补未完成节。

## 第四步：写作准备（指引与承诺一次成形，G4/G5 合并门一次停轮）

1. **一次生成两份产物**：`work/body/写作指引.md`（细则
   @references/guide-format.md：逐节判写作模式，可用素材与公司事实查清填表
   ——检索通道以环境实际的知识库/资料库工具为准，没有则向用户索取，
   **查到的素材与事实必须落进指引，否则写作用不上**）与 `work/body/关键事实与承诺.md`（事项｜值｜说明，
   一次列全：主体/工期/人员/服务承诺/报价口径；问答阶段已确认的直接沿用，
   值只取用户答复与分析产物里的事实，缺的标待补，不重复问）。
   **值列写可直接取用的值**（人名/数字/期限，如「12 个月」），解释放说明列
   ——说明列写用户能直接回答的问句或填空指引（带选项/示例，如「投标主体
   是哪种：企业/事业单位/个体工商户/自然人？」），**不写「须用户确认××」
   这类写给流程看的话**——校验按值列做全文子串比对。
2. **过 G4/G5 合并门（一次停轮）**：跑 `python3 scripts/check_writing.py --guide`
   → 刷新评审台（`python3 tender-review/scripts/build_review.py --work work`），
   指引与承诺两 tab 全量呈现（两 tab 只读，改动与给值走对话）；对话指针=
   写作构成一行（平实说法）+ 承诺值一行汇总（已知直列、待补点名）+「素材与
   值直接在对话里给我，或看完页面回『继续』」（规范见
   @tender-writing-suite/references/stage-gates.md G4/G5 行）。
3. **答复后直接开写**：用户答复（素材/值/「继续」）→ 更新指引与承诺清单
   （值落列）→ 首开进度面板，开写消息带「承诺值一行汇总 + 面板链接」与预期
   管理句 → 进第五步——本门只停这一次轮（未给值照默认：待补+批注）。

## 第五步：逐节生成（按指引模式路由）

首节开写前第一动作：`python3 tender-review/scripts/build_review.py --work work
--open-progress` **首开写作进度面板**（`work/review/写作进度.html`——
进度/批注清单/待确认承诺；服务模式自刷新，静态模式每节重跑 build_review），
把回显的面板链接以 Markdown 发给用户，并带一句预期管理：**「正文开始连续
生成，进度随时在这个面板看；写完或遇到要你定的事我再停」**（这一步之后是
长静默期，先说清预期）——**未首开禁写第一节**
（docx_toolkit 拒写 body 节、`check_writing --section` 拒查；首开标记
`work/review/.progress-opened` 落盘后放行，与入口红线「评审台未首建禁建节」
同构）。**进入本步必读
@references/section-writing.md**：三模式细则、工具链速查（docx_toolkit 在
`tender-docx/scripts/`）、素材三级落地、语言纪律都在其中。核心路由：素材修订 =
copy-from 注入素材授权区间（保真）→ revise 适配；格式跟随 = copy-from 拷格式件
区间 → revise 填空（表格只填不改，嫌版面难看加批注请用户手动处理；未定值批注）；推理撰写 = create 的 body 块一次成形（表格用
table 块，不流水句铺）。每节写完跑 `check_writing.py --section "<节.docx>"` 自查；
多节分派跑 `python3 scripts/build_dispatch.py --work work --sections "节A,节B"`
生成批次提示词（规则见细则）。

## 第六步：全局校验

全部节完成后过 G6 门：刷新评审台与进度面板（待办卡反映未写节；批注台账与校验
结果全量在进度面板——对话只给结论级汇报），**校验跑完直接进第七步合册**——
合册代价小（秒级重跑）、用户调整都在整本层做，**禁止问「现在合册还是先补齐」
类岔口**；补齐清单进进度页与收尾汇报，由用户拿到整本后处理（校验项没全清也
合册，收尾逐条点名）——招标格式件表格与原件不一致除外，那是废标级红线（入口
红线 8）：先修复再合册，合册也会拒绝。**收尾对话口径：交付物=整本 Word 文档——「合册」是
内部动作，不进对话、不作提问**（reply-style 硬规则 9）：

```bash
python3 scripts/check_writing.py            # 承诺比对 + 节间查重 + 残留扫描 + 招标格式件表格比对
                                            # 后面跟文档检查（金额验算、表格版式等）：只列表，不改变校验结论
python3 scripts/check_writing.py --old-names "旧公司名,旧项目名"   # 素材改写后必查残留
python3 scripts/check_writing.py --json work/review/check-result.json   # 落盘供进度面板读，收尾必带
```

- **旧名残留必查必清**（旧机构名留在交付稿里是高频事故）：`--old-names` 把用户提过的旧机构名
  都传进去（素材来源项目的招标方、我方曾用名等），别只靠素材文件名兜底；登记资料库里
  其他主体的登记名与旧名候选自动带上（疑似级，不拦）。
- **文档检查表随收尾交给用户**：单节自查与全局校验后面跟一段文档检查（单价 × 数量
  是否等于合价、合计对不对、大小写是否一致、各处总价是否一致、表格与图片版式、
  内部对账编号外泄等），不改变校验结论；那张表随收尾汇报原样贴给用户，用户说修
  才动文件。

## 链尾 HARD-GATE

第六步校验通过后**同轮**进入第七步合册：只载入 tender-docx 一个环节技能。
G6 的停轮交付（校验结论 + 整本 Word 文档路径 + 收尾汇报）由第七步收尾时
**一并完成、一次停轮**——校验后不单独停轮、不问岔口（与第六步「直接
合册」及 stage-gates G6 行同一口径——不是校验完先停下等你、下一轮再
合册，而是一路做到收尾汇报才停这一次）。
