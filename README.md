# docgo

简洁的 Go 库，用于读写 .docx Word 文档。设计参考 python-docx，专为智能体和开发者设计。

## 安装

```bash
go get github.com/samaidev/docgo
```

## 快速开始

```go
package main

import "github.com/samaidev/docgo"

func main() {
    doc := docgo.New()
    doc.SetTitle("报告").SetAuthor("SamAI")
    doc.AddHeading("季度报告", 1)
    doc.AddParagraph("本季度进展顺利。")
    doc.AddList([]string{"要点一", "要点二"})
    doc.AddTable([][]string{{"模块", "状态"}, {"文档", "✅"}})
    doc.Save("report.docx")
}
```

## 高级功能

- 页眉页脚：`doc.SetHeader("公司名")`
- 页码：`doc.SetPageNumber(true)`
- 自动目录：`doc.AddTableOfContents(3)`
- 表格样式：`table.SetBorderStyle("dashed")`
- 图片嵌入：`doc.AddImage("chart.png", 400, 300)`
- 文本高亮：`run.Highlight("yellow")`

## 保真编辑（重要）

打开一个已存在的 `.docx` 时，`docgo` 采用**保真编辑**策略：

- `OpenFile` 保留包内的完整 XML，未被修改的部件（图片、样式、页眉页脚、主题等）**原样写回**，不重新序列化；
- 仅对目标节点（如文本 run）做精准替换，因此原始排版、样式、媒体资源全部保留；
- 跨多个 `<w:t>` run 拆分的文本也能正确匹配替换（`ReplaceTextAcrossRuns`）。

```go
doc, _ := docgo.OpenFile("report.docx")
n := doc.ReplaceText("报告", "总结") // 替换 n 处
doc.SaveAs("report-new.docx")
```

## 旧版 .doc 支持（只读）

旧版二进制 `.doc` 为 OLE2/CFB 复合文档，支持**只读**读取与转换为 `.docx`：

```go
d, _ := docgo.OpenLegacy("old.doc")
text, _ := d.Text()          // 提取全部文本
docxBytes, _ := d.ToDocx()   // 转换为现代 .docx
```

> 设计约束：旧版 Office 文件仅支持读取，编辑请使用现代 `.docx` 格式。

## 许可证

MIT
