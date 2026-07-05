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

## 许可证

MIT
