---
name: tender-docx
display_name: 灵燕标书·整本交付
display_name_en: Bid Assembly & Delivery
description: "标书整本交付：把节 docx 自动编号合成整本 Word，实测目录页码，处理格式件前置分页与导出检查，输出收尾汇报。适用于：合册、合成整本、生成最终 Word、编章节页码、目录页码、导出 PDF、交付前检查。"
description_zh: "标书交付环节：把逐节写成的正文合成整本 Word——编号与目录页由整本生成机械完成，待办批注原样迁移，导出「标题→页码」映射供回填；整本是派生产物，改内容回节文件层改再生成。收尾逐条点名待办批注、素材异议与导出提醒，交付物口径=整本 Word 文档（「合册」是内部动作，不进对话）。"
description_en: "Bid assembly stage: assemble section docx files into the final numbered volume with measured TOC page numbers, migrate pending comments, export a heading-to-page map, and deliver the closing report."
examples_zh:
  - "把各节合成整本并生成目录页码"
  - "合册后导出 PDF 预览"
  - "交付前检查一遍待办批注"
examples_en:
  - "Assemble all sections into the final volume with TOC page numbers"
  - "Export a PDF preview after assembly"
  - "Check pending comments before delivery"
version: 2.2.3
author: 灵燕智能
---

# 合册与交付（第七步）

开始前先载入入口技能 tender-writing-suite：红线与门行为以它为准；本文件只讲
合册与交付。环境依赖与工作目录约定同在入口技能。

## 上游检查

`work/body/<节>.docx`、`work/body/写作指引.md`、`work/body/关键事实与承诺.md`
是合册输入；节未全部完成 → 载入 tender-writing 回到第五步。`work/review/`
评审台与 G6 校验结果应在合册前刷新过（见 @tender-writing-suite/references/stage-gates.md
G6 条）。

## 第七步：合册整本

```bash
python3 tender-docx/scripts/docx_toolkit.py assemble --work work --numbering chapter
python3 tender-docx/scripts/docx_toolkit.py toc-pages "work/body/整本-投标文件.docx" --pdf work/导出/预览.pdf --map work/pagemap.json
```

`--numbering` 方案见 `--help`；`--toc-depth` 控制目录层级（默认 3 级，静态条目
与域刷新口径一致）；`--flat-outline` 把节内小标题摘出导航窗格（导航降噪用，
默认关——开了 F9 刷目录会丢二级以下条目）。`--map` 导出「标题→页码」JSON 供
评分索引表等机械回填。**页码口径=行业惯例**：封面/前置/目录不编页码，正文分节
自第 1 页起编。**整本是派生产物**：编号与目录页由合册机械生成、待办批注原样
迁移；改内容回节文件层改再合册（直接手改整本会被覆盖）；返回的警告逐条点名。

**合册前先核对招标格式件**：节内表格与原件格式表对不上（列数变了、原件标签丢了、
整张表没了）即拒绝合册——擅自改格式件是废标级红线（入口红线 8）；用户在对话里
明确同意改结构的节，加 `--allow-format-change "<节名>"`（多个用逗号分隔）放行。

`toc-pages` 是**实测→回填→复测的收敛闭环**（默认至多 5 轮，`--max-rounds` 调）：
回填会改目录占行、正文跟着再漂，两轮实测一致才算收敛；未收敛回滚并报警，不
静默输出错页码。收敛后顺带核对表格里「第X章」交叉引用（只报事实，章号写错或
章节增删未同步会被点名）。

**终稿字体固化（合册无条件执行）**：theme 东亚字形补显式 + 样式层归一 + run 级
剥主题属性并固化 eastAsia——模板会被 skill 更新覆盖、拷入的历史标书 run 会带
越权主题属性，终稿后处理让字体问题在交付面上不可见；回显 `[字体固化]` 一行。

版式模板：默认 `templates/base-template.docx`；固定模板传 `--template`（需含
Heading 1-3，缺则回落内置样式并点名）。需重建模板时跑
`python3 tender-docx/scripts/make_template.py`。中文字体链路起建时自动归一；
用户模板带日文主题语言时建节回显 `[模板中文字体]` 一行，用户报「标题字体不对」
先看它。

## 版式与图片工具

- **页数压缩** `compact --docx <file> [--line-pt 22]`：正文段固定行距（exact）。
  大字号段（字号×1.2 超行距）、含图段、目录段保护常开不可关——exact 行距会把
  36pt 封面标题压扁、图片裁成一条，靠保护跳过而非事后补救。表格与页眉页脚不动。
- **图片方向核对** `image-scan --docx <file>`：导出全部内嵌图片的缩略图网格
  （带序号），人眼扫一遍核方向——证件方向自动判定不可靠；发现错向记序号，
  `image-rotate --docx <file> --index N --deg 90|180|270` 就地旋转（自动按新
  比例重算引用处 extent，不会拉伸变形）。

## 人工修改与回流

节文件、写作指引、关键事实与承诺、目录树都是**普通文件**，可任意编辑器直改；
**评审页面是推荐的修改入口**（`tender-review/scripts/serve_review.py` 起服务后
页面直改，保存落盘出 diff 报告，规程见
@tender-review/references/review-workbench.md）。回流规则：改了节 .docx →
重新 `assemble`；改了目录树 → 重新 `assemble`（缺失节会被点名）；改指引/承诺 →
重跑 `tender-writing/scripts/check_writing.py` 对账。

**修改入口口径与停服务联动**：本地服务收尾即 `--stop` 关停，停后评审页面链接
失效——交付/收尾消息里修改与批注销号的入口只写 Word 批注（批注在文档里，不
依赖服务）；用户想用页面改，给离线文件路径（`work/review/评审台.html` 浏览器
直接打开）或重起服务命令，不把本地链接当活入口引用。

## 收尾汇报

- 各节完成情况、未写节及原因、整本 Word 文档路径（响应文件构成有几册就几个
  文件）——**对话口径=整本 Word 文档，不说「合册」**（reply-style 硬规则 9）；
- **待办批注逐条点名**（缺什么/要用户定什么）；素材异议点名（节名 + 不符点 +
  建议）；图示计划 vs 实际产出对账；
- **沉淀建议**（口径见 @tender-writing-suite/references/reply-style.md，只列
  清单不擅动库）；
- **导出提醒（必报）**：交付 PDF 用 Word/WPS 导出（目录域自动刷新）；
  LibreOffice 导出的 PDF 宋体被换黑体，只作内部预览。
- **不自动生成 PDF 交付稿**：AI 生成稿必然还要人工调整，PDF 由用户在终稿后
  自行从 Word/WPS 导出；`toc-pages --pdf` 的预览 PDF 仅内部测页码用，不作
  交付物主动提供。

## 链尾

交付完成，无下一环。用户提出改正文 → 回 tender-writing 写作环节改节文件后
重新合册。
