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
version: 2.1.3
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

`--numbering` 方案见 `--help`；`--map` 导出「标题→页码」JSON 供评分索引表等
机械回填。**页码口径=行业惯例**：封面/前置/目录不编页码，正文分节自第 1 页起编。
**整本是派生产物**：编号与目录页由合册机械生成、待办批注原样迁移；改内容回节
文件层改再合册（直接手改整本会被覆盖）；返回的警告逐条点名。

版式模板：默认 `templates/base-template.docx`；固定模板传 `--template`（需含
Heading 1-3，缺则回落内置样式并点名）。需重建模板时跑
`python3 tender-docx/scripts/make_template.py`。

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
