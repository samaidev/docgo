package docgo

import (
	"os"
	"strings"
	"testing"
)

func TestCreateAndSave(t *testing.T) {
	doc := New()
	doc.SetTitle("测试文档")
	doc.SetAuthor("Tester")
	doc.AddHeading("标题", 1)
	doc.AddParagraph("正文内容")
	doc.AddList([]string{"项1", "项2"})
	doc.AddTable([][]string{{"A", "B"}, {"1", "2"}})
	doc.AddCodeBlock("go", `fmt.Println("hi")`)

	tmpFile := "/tmp/docgo-test.docx"
	if err := doc.Save(tmpFile); err != nil {
		t.Fatal(err)
	}
	defer os.Remove(tmpFile)

	if _, err := os.Stat(tmpFile); err != nil {
		t.Fatal("file not created")
	}
}

func TestOpenAndRead(t *testing.T) {
	doc := New()
	doc.SetTitle("读取测试")
	doc.SetAuthor("Reader")
	doc.AddParagraph("段落一")
	doc.AddParagraph("段落二")

	tmpFile := "/tmp/docgo-read-test.docx"
	doc.Save(tmpFile)
	defer os.Remove(tmpFile)

	doc2, err := Open(tmpFile)
	if err != nil {
		t.Fatal(err)
	}
	if doc2.Title() != "读取测试" {
		t.Errorf("title = %q, want '读取测试'", doc2.Title())
	}
	if doc2.Author() != "Reader" {
		t.Errorf("author = %q, want 'Reader'", doc2.Author())
	}
	if len(doc2.Paragraphs()) < 2 {
		t.Errorf("paragraphs = %d, want >= 2", len(doc2.Paragraphs()))
	}
}

func TestRunStyling(t *testing.T) {
	doc := New()
	p := doc.AddParagraph("")
	r := p.AddRun("styled")
	r.Bold(true).Italic(true).Color("FF0000").Size(14)

	if !r.bold { t.Error("bold not set") }
	if !r.italic { t.Error("italic not set") }
	if r.color != "FF0000" { t.Error("color not set") }
	if r.size != 14 { t.Error("size not set") }
}

// TestExtractParagraphsWithTable 验证含表格与自闭合空段落的文档
// 能被正确解析（借鉴 LibreOffice 的 w:p / w:r / w:t 分组范式）。
func TestExtractParagraphsWithTable(t *testing.T) {
	// 表格位于 <w:body> 最前、且首段为自闭合空段落 <w:p .../> 的情形。
	xml := `<w:document><w:body>` +
		`<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="4428"/></w:tcPr>` +
		`<w:p w:rsidR="1"><w:r><w:t>CellA</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p w:rsidR="2"><w:r><w:t>CellB</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p w:rsidR="3"><w:r><w:t>Para</w:t><w:tab/></w:r><w:r><w:t>End</w:t></w:r></w:p>` +
		`<w:p w:rsidR="4"/>` + // 自闭合空段落
		`</w:body></w:document>`

	paras := extractParagraphs(xml)
	// 表格内的单元格段落应被排除（见 extractTables），自闭合空段落产生空字符串，
	// 因此过滤后只应有 1 条非空正文段落。
	var nonEmpty []string
	for _, p := range paras {
		if p != "" {
			nonEmpty = append(nonEmpty, p)
		}
	}
	if len(nonEmpty) != 1 {
		t.Fatalf("non-empty paragraphs = %d, want 1 (got %q)", len(nonEmpty), paras)
	}
	if nonEmpty[0] != "Para\tEnd" {
		t.Errorf("P0 = %q, want 'Para\\tEnd'", nonEmpty[0])
	}

	tables := extractTables(xml)
	if len(tables) != 1 {
		t.Fatalf("tables = %d, want 1", len(tables))
	}
	if tables[0] != "CellA\tCellB" {
		t.Errorf("table = %q, want 'CellA\\tCellB'", tables[0])
	}
}

func TestTables(t *testing.T) {
	// 构造含表格的 document.xml 并验证 Tables() 公开 API。
	xml := `<w:document><w:body>` +
		`<w:p><w:r><w:t>前言</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>` +
		`<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`</w:body></w:document>`

	doc := &Document{body: xml}
	tables := doc.Tables()
	if len(tables) != 1 {
		t.Fatalf("tables = %d, want 1", len(tables))
	}
	if tables[0] != "A1\tB1\nA2\tB2" {
		t.Errorf("table = %q, want 'A1\\tB1\\nA2\\tB2'", tables[0])
	}
}

func TestBytes(t *testing.T) {
	doc := New()
	doc.AddHeading("H1", 1)
	doc.AddParagraph("P1")
	data, err := doc.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if len(data) < 500 {
		t.Errorf("bytes too small: %d", len(data))
	}
	if !strings.HasPrefix(string(data), "PK") {
		t.Error("not ZIP format")
	}
}
