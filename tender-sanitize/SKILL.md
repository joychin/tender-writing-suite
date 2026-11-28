---
name: tender-sanitize
display_name: 灵燕标书·文档脱敏
display_name_en: Bid Document Metadata Sanitizer
description: "递交前清除 Word 与 PDF 文档里的个人信息与工具痕迹：作者、最后修改人、公司名、管理者、模板路径（常含系统用户名）、WPS 与企业模板的自定义属性、批注与修订作者名录、缩略图、编辑会话标记（rsid），PDF 的 /Info、XMP、内嵌附件。先体检出清单再清理出副本，原文件不动；只动文档属性，正文一个字不动，清理后自动复核正文一致。适用于：清除元数据、删除文档属性、去掉作者信息、文档脱敏、匿名化、暗标递交前清痕迹、查这份文件还能不能看出是谁做的。.doc 先转 .docx 再清。"
description_zh: "投标文件的文档属性可能暴露身份：作者写着真名、扩展属性里存着公司名、模板路径带着系统用户名、WPS 保存写入自定义标识——匿名评审前值得清一遍。本技能先体检出一份清单（高=直接指向人或单位，中=工具指纹与缩略图，低=能跨文件关联的编辑会话标记），确认后清理出副本，原文件一个字节不动；只动文档属性，正文、格式、图片原样保留，清理后自动复核正文一致并复查属性。套件产出的整本也能在交付前加一道脱敏。"
description_en: "Strips identifying metadata from bid documents before submission: author, last-modified-by, company, manager, template paths leaking OS usernames, WPS custom properties, revision/comment author names, thumbnails, rsid session marks, and PDF /Info, XMP and embedded attachments. Always checks first, then cleans into a copy while the original stays untouched; body text, formatting and images are preserved and re-verified after cleaning."
examples_zh:
  - "把这份标书的作者信息清掉，暗标要交"
  - "查一下这几个文件里还能不能看出是谁做的"
  - "递交前把投标文件里所有文档属性清一遍"
examples_en:
  - "Strip author info from this bid document for anonymous submission"
  - "Check whether these files still reveal who made them"
  - "Clean all document properties before submission"
version: 2.4.4
author: 灵燕智能
---

# 文档脱敏（不占流水线步序，随时可用）

开始前先载入入口技能 tender-writing-suite：红线与回复口径以它为准；本文件只讲
文档脱敏。环境依赖：docx/docm 清理只用标准库，PDF 需要 pypdf；.doc 先转成 .docx 再清——
5MB 以内本机有 LibreOffice 脚本自动转，没有就由你选本机可用的办法转换，超过 5MB 的请
用户自己另存（见「纪律」第一条）；工作目录约定同在入口技能。

## 它解决什么

暗标、匿名评审场景：文档属性可能直接暴露身份。可泄露的点不止「作者」一栏——
扩展属性里可能存公司名、模板路径常带系统用户名、WPS 保存会写入自定义标识、修订和批注
各自带作者名录。字段级依据：@references/metadata-fields.md（用户追问时读）。

## 输入

一份或几份 .docx / .docm / .pdf（.doc 先转成 .docx 处理，输出 .docx）；也可以给
目录（`-r` 递归）。没有流水线上游，单独请求直接可用；招标文件本体不做脱敏——
泄露的是投标方身份，招标文件没有可泄的，而且它是格式件来源，不碰。

## 流程（体检 → 清理 → 复核回执）

脚本：`scripts/strip_metadata.py`（命令示例以安装根为基准解析成绝对路径）。

1. **用户要清除**：体检 + 清理 + 复核一次做完（清理是副本输出、原件不动，
   不需要中间停轮确认）：

   ```bash
   python3 scripts/strip_metadata.py <文件…> --json
   ```

   默认输出与源文件同目录的 `<文件名>_cleaned<扩展名>`；用户要求集中输出再加
   `-o <目录>`。回执按「结果 → 影响 → 下一步」说：清了哪些类目、输出在哪、
   正文校验一致、复查干净；`[注意]` 行（签名失效、文本指纹差异、.doc 已转换）
   要逐条转述给用户。
2. **用户只问「还有没有痕迹」**：只跑 `--check`（只读不改；有发现退出码 1），
   把清单摊给用户，等用户决定清不清。
3. **清什么由参数定，用户有特殊口径就调参**：要保留 PDF 内嵌附件 →
   `--keep-attachments`；要保留创建/修改时间 → `--keep-dates`；只清文档属性、
   修订作者与 rsid 都留 → `--basic`。用户没说就按默认（附件与时间都清）。

## 纪律

- **.doc 先转成 .docx 再清，转换办法按顺序找（LibreOffice 不是必装前提）**：①**超过 5MB 的
  .doc**：脚本不自动转（LibreOffice 转大文件可能很久都转不完），直接退出（退出码 2）——
  请用户在 Word/WPS 里另存为 .docx，这一步不要替用户自动化；②脚本自己找本机 LibreOffice，
  找到就直接转（超时或失败时脚本会终止转换并退出，同样请用户另存）；③没找到：脚本列出检测到的 Office 套件后退出（退出码 2）——**由你选本机
  最可行的办法把 .doc 转成 .docx**（Word、WPS、Pages 的自动化另存或导出，宿主自带的
  文档转换能力，本机其它转换工具）；④都没有：请用户在 Word/WPS 里另存为 .docx 再发来。
  哪条路转好的，都对那份 .docx 运行同一条命令，清理仍走脚本。LibreOffice 体积大、不少
  用户装不了——它是办法之一，不要一上来就让用户安装。原件全程不用动。
- **副本是缺省，就地覆盖须用户明说**：`--in-place` 只在用户对话里明确说了
  「覆盖原文件」这类话才用，建议同时 `--backup`。对套件生成的整本 Word 用
  `--in-place` 会让整本指纹失配，后续合册校验按被改过处理——整本一律副本输出
  （见 @references/metadata-fields.md 最后一节）。
- **只动元数据不改正文**：引擎写完复核正文指纹（docx 比全部文字节点的 sha256，
  pdf 比页数/页面尺寸/提取文本），不一致即中止且不产出文件——不要绕过校验
  手工改文件。
- **回执用用户的词**：对用户说「文档属性」「个人信息」「脱敏」，[高]/[中]/[低] 分级
  落成人话（高=直接指向人或单位；中=工具指纹与缩略图；低=能跨文件关联的编辑
  会话标记，不算「可以不管」）。
- **没查的要说没查**：`.xlsx`/`.pptx` 不在范围、图片 EXIF 不动、正文人名属于
  内容要另查——这些边界主动说，别让用户以为「全干净」。

## 边界（主动讲清）

- **.doc 的输出是 .docx**：要回存 .doc 在 Word/WPS 里另存（转换办法见「纪律」第一条）。
- **加密 PDF**：口令为空也打不开时报错，请用户先解密。
- **数字签名 PDF**：重写后签名失效，需重新签（引擎探测到会提醒）。
- **修订与批注的内容不删**：作者抹成匿名、正文保留——「接受全部修订／删除全部
  批注」是编辑动作，走 tender-docx 的修订与批注命令，别混在脱敏里做。
- **元数据干净 ≠ 内容干净**：正文里的公司名、项目名、人名要另外查
  （tender-docx 的文档检查覆盖占位与残留扫描）。
