---
name: tender-review
display_name: 灵燕标书·评审台
display_name_en: Bid Review Workbench
description: "标书确认面：构建并运行分析报告页、评审台（目录/指引/承诺）与写作进度面板，应用页面上的修改。适用于：打开评审台、评审页面、看分析报告、在页面上改目录、确认目录、应用修改包、刷新确认面。由分析/写作环节在确认门调用，不单独承担标书写作。"
description_zh: "标书确认面：G3 起评审台是唯一确认面——目录树、写作指引、承诺拍板全在页面里改，保存直写产物并出 diff 报告；分析阶段的报告页与写作进度面板同源构建。"
description_en: "The bid confirmation surface: builds and serves the analysis report, the review workbench (outline/guide/commitments) and the writing progress panel; page edits are saved straight back to artifacts with a diff report."
examples_zh:
  - "打开评审台，我要改目录"
  - "刷新一下承诺拍板页面"
  - "把我刚改的目录应用落盘"
examples_en:
  - "Open the review workbench so I can edit the outline"
  - "Refresh the commitments review page"
  - "Apply my outline edits to the artifacts"
version: 2.2.4
author: 灵燕智能
---

# 评审台（唯一确认面）

开始前先载入入口技能 tender-writing-suite：确认门纪律、停轮与指针口径以
@tender-writing-suite/references/stage-gates.md 为准；本文件只讲页面怎么建、
怎么服务、改动怎么应用。

## 职责边界

本技能只做确认面的构建、服务与改动应用，不做解析、写作与合册。被
tender-analysis（G1/G2 报告、G3 首建）与 tender-writing（G4/G5/G6 刷新、
进度面板）在门点调用。

## 上游检查

`work/` 下已有任一产物（parse/analysis/outline/body）即可构建；空工作区 →
引导先跑 tender-analysis。

## 三套页面

| 页面 | 模板 | 构建方式 |
|---|---|---|
| 分析报告（G1/G2） | `templates/analysis-report.html` | `build_review.py --pages report` |
| 评审台（G3 起，目录/指引/承诺 tab） | `templates/review-workbench.html` | `build_review.py` 全量 |
| 写作进度（第五步首开） | `templates/progress.html` | `build_review.py` 全量产出；服务模式 `/<token>/progress.html` 自刷新 |

```bash
python3 tender-review/scripts/build_review.py --work work --pages report   # 分析阶段（G1/G2）：只产分析报告
python3 tender-review/scripts/serve_review.py --work work            # 先起服务（打印带 token 的裸 URL）
python3 tender-review/scripts/build_review.py --work work            # G3 起全量生成；服务已起时回显 [交付链接]（带 ?v=），复制它交付
python3 tender-review/scripts/serve_review.py --work work --status   # 后续引用前探活（刚起服务不要立刻跑，后台启动有竞态）
python3 tender-review/scripts/serve_review.py --work work --stop     # 收尾关停（必做）
python3 tender-review/scripts/apply_review.py --patch 改动文件.json --work work   # 助手代为应用
```

服务未起时 `build_review.py` 不回显交付链接，只提示「先起服务再重跑」——G3 的
正确顺序是**先起服务、再重跑全量 build**，不是 build 完再起服务（初建那次没有
`?v=` 链接，容易误用起台时的裸 URL）。

模板即唯一视觉来源：令牌集中在 `templates/review-workbench.html` 的 `:root`，
改视觉先改模板；服务模式 GET 实时读盘读模板，刷新即生效（静态离线版是生成时点
快照，每道门须重跑 `build_review.py`）。

## 页面规程

@references/review-workbench.md：锚点 tab（overview/tree/guide/promise）、
通道矩阵、保存直写产物与 diff 报告、应用后的必做回流（重跑
`tender-writing/scripts/check_writing.py --guide` 等）。

## 停轮与链接

页面交付后当轮以指针收束停轮：一行结论 + 待办去向一句话 + 「需要你补充」表格
（`| 需补充 | 影响 | 不补充的处理 |`，处理列一律「Word 批注标注，可后补」）+
固定告知「不补充也可以继续往下——缺失部分会以 Word 批注标在正文对应位置」+
至多 3 条标题级提醒；URL 一律 Markdown 链接发出（裸代码串不算交付）。
`serve_review.py` / `build_review.py` 启动回显里的停轮提醒与本条同源，改动时
两处同改。

## 链尾

页面构建或改动应用完成、发指针停轮后，回到发起环节技能（不强制下一环）；
门是否 pass 由 agent 依用户答复写入 `work/decisions.json`。
