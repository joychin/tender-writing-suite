# 元数据字段清单（doc / docx / pdf）

给「到底删了什么、为什么算泄露」这类追问用的细节。SKILL.md 里只列结论，这里是依据。

---

## 一、DOCX / DOCM 是 zip 包

```
word/document.xml         正文
word/header1.xml …        页眉 / 页脚 / 脚注 / 尾注
word/comments.xml         批注正文与作者
word/people.xml           修订作者名录
word/settings.xml         文档级设置（模板路径、rsid）
word/_rels/*.rels         部件关系
docProps/core.xml         核心属性
docProps/app.xml          扩展属性（Word 版本、公司、模板）
docProps/custom.xml       自定义属性
docProps/thumbnail.*      缩略图（首屏内容预览）
[Content_Types].xml       部件类型声明
```

**为什么不用 python-docx**：它只暴露 `core_properties`（即 core.xml 的十几个字段），
对 `app.xml` 的 Company / Manager / Template、`custom.xml`、缩略图、修订作者、rsid
一概不可见也改不了。而这些恰恰是最主要的泄露点。所以脚本直接读写 zip 内的 XML 部件，
未涉及的部件按原字节复制（图片这类必须逐字节保持不变）。

**为什么清空部件而不是删除部件**：删部件要同时改 `_rels/*.rels` 与
`[Content_Types].xml`，任何一处不一致 Word 就报「文档已损坏」。只有缩略图是例外
（它的内容无法「清空」，只能整个删掉并同步摘掉关系）。

---

## 二、`.doc` 要先转换

`.doc` 是二进制复合文档（OLE），元数据埋在内部流里，跟 zip+XML 两码事，不能在本
层面直接改。所以先转出 `.docx`，对它做完整个 docx 流程再落盘，原 `.doc` 一个字节
不动；输出一律是 `.docx`。要回存 `.doc` 得在 Word/WPS 里另存（转换回去由 Office
自己写，版式最保险）。

**转这一步谁来做，按顺序找（LibreOffice 是办法之一，不是前提）**：

1. **超过 5MB 的 `.doc`**：脚本不自动转——LibreOffice 转大文件可能很久都转不完，让模型
   另想办法转又容易丢表格和版式。脚本回显后退出，**请用户在 Word/WPS 里另存为 `.docx`，
   这一步不要替用户自动化**。
2. 5MB 以内，脚本内置的自动转换只有 LibreOffice——本机找得到就直接转。超过 300 秒没转完，
   脚本把转换进程终止掉再退出（整棵进程树一起杀，不留空转的 soffice）；转换失败也是同一个
   出口——**同样请用户在 Word/WPS 里另存为 `.docx`，不要替用户自动化，也不要另想办法转**。
   转不出来的多半是个别图片／绘图对象的数据不规范，与文件大小无关。
3. 找不到：脚本探出已装的 Office 套件（macOS 看应用包，Windows 看自动化接口的注册），
   连同「转好后对那份 `.docx` 重跑同一条命令」一起回显后退出，**由模型选本机的办法**。
   脚本不替它调用 Word/WPS 的自动化：各平台各版本差异大，没在真机验证过的不内置
   （WPS for Mac 实测没有脚本接口，只能请用户在里面手动另存）。
4. 都不行：请用户在 Word/WPS 里另存为 `.docx`。

无论谁转的，清理都走脚本——只有它带正文指纹自证与复扫；回放确认过，接力路径的结果
与自动转换的结果正文指纹一致。

---

## 三、`docProps/core.xml`

| 元素 | 含义 | 泄露什么 |
|---|---|---|
| `dc:creator` | 首次保存作者 | 建文档的人（Word 里默认是 Office 用户名，常为真实姓名） |
| `cp:lastModifiedBy` | 最后保存者 | 最后一个改稿的人 |
| `dc:title` / `dc:subject` | 标题 / 主题 | 内部命名，可能是「XX项目投标-内部讨论稿」 |
| `cp:keywords` | 关键词 | 常出现「机密」「内部」「草稿」 |
| `dc:description` | 备注 / 批注 | 自由文本，任何内容都有可能 |
| `cp:category` | 类别 | 内部归档分类 |
| `cp:revision` | 修订号 | 说明改了多少轮 |
| `dcterms:created` | 创建时间 | 起草时间（能反推项目节奏） |
| `dcterms:modified` | 修改时间 | 最后改动时间 |
| `cp:lastPrinted` | 最后打印时间 | 打印过、什么时候打的 |
| `cp:contentStatus` / `dc:language` / `dc:identifier` / `cp:version` | 状态 / 语言 / 标识 / 版本 | 内部流程标记 |

`w:author` 是必填属性，所以脚本把值置成空字符串 `w:author=""` 而不是删除属性
（删了 Word 会判非法 XML 并要求修复）。其余属性都是可选的，直接删。

---

## 四、`docProps/app.xml`

| 元素 | 泄露什么 |
|---|---|
| `Company` | **公司名**。投标场景最致命的一项——封面可以不写，属性里却写着 |
| `Manager` | 管理者姓名（组织架构） |
| `Template` | 模板路径，常见 `C:\Users\<用户名>\AppData\Roaming\Microsoft\Templates\...`，**直接泄露操作系统用户名** |
| `TotalTime` | 累计编辑分钟数（投入量、是否临时拼的） |
| `HyperlinkBase` | 内部共享盘路径，如 `\\fileserver\投标\...` |
| `Application` / `AppVersion` | 生成工具与版本（Word / WPS 具体版本号） |
| `Pages` / `Words` / `Characters` / `Lines` / `Paragraphs` | 统计值，可推算内容规模 |

脚本把整个 app.xml 重写为空 `<Properties/>`：这些值全部由文档派生，Word/WPS
下次保存会自行重算，删掉不影响打开。

---

## 五、`docProps/custom.xml`

企业模板和 WPS 会往里塞私有属性，例如：

| name | 来源 | 说明 |
|---|---|---|
| `KSOProductBuildVer` / `KSOTemplateDocerSaveRecord` | WPS | WPS 版本与模板痕迹 |
| `ICV` | WPS | 内部文档标识/校验值，**可用于跨文件关联** |
| `项目负责人` / `部门` / `密级` | 企业模板 | 直接就是人名与组织信息 |
| `_MarkAsFinal` / `ContentTypeId` | Office | 内部状态标记 |

自定义属性没有固定字段名，**任何名字都可能出现**，所以整份清空。

---

## 六、缩略图 `docProps/thumbnail.*`

保存时生成的首页缩略图（jpeg / emf / wmf）。除体积泄露外，缩略图不是正文的一部分，
看不到修订痕迹，等于把「首屏内容」单独留了一份副本。脚本删除该部件，并同步删除
`_rels/.rels` 中 `Type=".../thumbnail"` 的关系（否则关系指向不存在的部件，Word 会报错）。

---

## 七、正文部件里的修订与批注痕迹

出现在 `word/document.xml`、`word/header*.xml`、`word/footer*.xml`、
`word/footnotes.xml`、`word/endnotes.xml`、`word/comments*.xml`。

带 `w:author` 的元素：

| 元素 | 场景 |
|---|---|
| `w:ins` / `w:del` | 插入 / 删除修订 |
| `w:moveFrom` / `w:moveTo` | 移动修订 |
| `w:rPrChange` / `w:pPrChange` / `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange` / `w:tblGridChange` / `w:sectPrChange` / `w:numberingChange` | 格式修订 |
| `w:cellIns` / `w:cellDel` / `w:cellMerge` | 表格单元格修订 |
| `w:comment` | 批注（`word/comments.xml`） |

同一元素上还有 `w:initials`（缩写）与 `w:date`（操作时间）。三者组合起来能还原
「谁在什么时候改了什么」，甚至排出多人协作文档的编辑顺序。

脚本的处理：`w:author=""`、`w:initials=""`、删 `w:date`。**修订内容和批注正文保留**，
因为那属于正文；要不要接受修订是编辑决策，不是清洗决策。

---

## 八、`word/people.xml`

协作期间出现过的作者名录（`<w:person w:author="张三"/>`）。即使某人的修订已经被接受、
正文里再也搜不到他的名字，这份名录仍会列出来。脚本把它重写为空 `<w:people/>`。

---

## 九、`word/settings.xml` 与 rsid

| 项 | 泄露什么 |
|---|---|
| `w:attachedTemplate r:id` | 指向模板文件；真实路径在 `word/_rels/settings.xml.rels` 的 `Target`，常含用户名与共享盘 |
| `w:rsids` 块 | 本次文档所有修订会话 ID 的清单 |
| `w:rsidRoot` | 文档根 rsid |
| `w:rsidR` / `w:rsidRPr` / `w:rsidP` / `w:rsidDel` / `w:rsidSect` | 段落/字符级修订会话标记，散布在正文各处 |

rsid 本身不含人名，但它是**稳定的编辑会话指纹**：把两文件里的 rsid 求交集，就能判断
它们是否出自同一次编辑、是否互相复制粘贴过。所以按「低」级处理但也清掉——
`--check` 里出现即退出码 1，理由是它可以用来做跨文件关联，不是「可以不管」。

脚本删 `w:attachedTemplate` 元素 + `settings.xml.rels` 里的对应关系、
删 `w:rsids` 与 `w:rsidRoot`、删所有 `w:rsid*` 属性。

---

## 十、PDF `/Info` 字典

| 键 | 泄露什么 |
|---|---|
| `/Author` | 作者（导出的 PDF 里常沿用 Office 用户名；程序生成的 PDF 会出现生成库名） |
| `/Creator` | 原始创作程序，如 `WPS 文字`、`Microsoft Word` |
| `/Producer` | 生成链，如 `macOS Quartz PDFContext`、`Acrobat Distiller 21` |
| `/Title` / `/Subject` / `/Keywords` | 内部标题、主题、密级词 |
| `/CreationDate` / `/ModDate` | 生成与修改时间（带时区，可定位到人所在时区） |
| `/Trapped` | 印刷流程标记 |
| `/SourceModified` | 源文件修改时间（与 ModDate 配合可推改稿过程） |

脚本把整个 `/Info` 删除（`--keep-dates` 时只保留两个日期键）。

注意：不要靠 `grep Author 文件.pdf` 判断有没有作者——PDF 1.5+ 的对象流会压缩
`/Info`，明文搜不到但它是存在的。判断要用 `pdfinfo` 或 pypdf 读 `reader.metadata`。

---

## 十一、XMP `/Metadata`

XML 形式的元数据流，内容与 `/Info` 大幅重叠但**独立存在**：删了 `/Info` 不清 XMP，
仍能从 XMP 读到 `dc:creator`、`xmp:CreatorTool`、`pdf:Producer`、`xmp:CreateDate`
等。反过来也一样。所以两边都必须清。

脚本删除 catalog 里的 `/Metadata` 引用与流对象。

---

## 十二、页面级与批注

| 位置 | 内容 |
|---|---|
| 页面 `/Thumb` | 页面缩略图流 |
| 页面 `/PieceInfo` | 生成程序的私有数据（Illustrator 之类会写应用与版本） |
| 页面 `/LastModified` | 单页修改时间 |
| 批注 `/T` | 批注作者名 |
| 批注 `/NM` | 批注唯一标识（可用于跨文件追踪同一条批注） |
| 批注 `/M` | 批注修改时间 |

脚本对每页与每个批注逐项删除。

---

## 十三、内嵌附件与文件标识

| 位置 | 说明 |
|---|---|
| `/Names /EmbeddedFiles` | PDF 里嵌的源文件（Word 稿、Excel 表、图片原图）。附件本身带完整元数据，是最容易被忽略的泄露路径。脚本默认删除并列出文件名，`--keep-attachments` 可保留 |
| `/AF`（Associated Files） | PDF 2.0 的关联文件，同上 |
| trailer `/ID` | 文件标识数组，常由创建时间/路径派生。脚本置空 |

`/ID` 无法在所有情况下真正删除（取决于写入器），脚本尽力而为；即使保留也只是文件
指纹，不含个人信息。

---

## 十四、不在清理范围（别误以为清了）

| 项 | 为什么不动 |
|---|---|
| 正文文字里的姓名、公司、项目名 | 属内容；要匿名得逐处改写 |
| 修订标记本身、批注正文 | 属内容；要处理是「接受修订 / 删除批注」的编辑动作 |
| PDF JavaScript、表单域、隐藏图层、链接 | 属行为/内容，动了会改文档功能 |
| 邮件元数据、图片 EXIF | 范围外；docx 内图片按原字节复制，EXIF 原样保留 |
| 文件系统时间戳 | 不在文件内部，需 `touch -t` 这类系统层面处理 |
| `.xlsx` / `.pptx` | 结构类似但批注作者位置不同（`xl/comments*.xml`、`ppt/commentAuthors.xml`），未实现，遇到要明说 |

## 十五、与套件整本 Word 的指纹机制的关系

套件的合册环节给整本记 sha256 指纹（`<整本名>.docx.fingerprint.json`），用来发现
「整本被手工改过」。对整本做 `--in-place` 清理后指纹文件还留着旧值，两个检查点会
按被改过处理（真机回放验证过）：`docx_toolkit inspect --work` 的整本对账（确定级
「整本被手工改过」）与合册前的覆盖提醒（「本次生成会覆盖这些改动」）。所以**整本
优先用默认的副本输出**（`<原名>_cleaned.docx`，真机回放确认对原整本与指纹零干扰），
或在清理后重新走一遍合册。对用户手头独立文件的清理不涉及这条。
