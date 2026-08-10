package docgo

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// buildDocxWithExtras 构造一个包含图片、样式、页眉等附加部件的 .docx。
//
// 这些部件正是本库元素模型未建模的内容，因此最能检验保真性：
// 编辑正文后它们必须原封不动。
func buildDocxWithExtras(t *testing.T, bodyXML string) []byte {
	t.Helper()

	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` + bodyXML + `</w:body></w:document>`,
		"word/styles.xml": `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="CustomStyle"><w:name w:val="Custom Style"/></w:style></w:styles>`,
		"word/header1.xml": `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:r><w:t>页眉内容</w:t></w:r></w:p></w:hdr>`,
		"docProps/core.xml": `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>原始标题</dc:title><dc:creator>原作者</dc:creator></cp:coreProperties>`,
	}

	binaryParts := map[string][]byte{
		"word/media/image1.png": {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0xDE, 0xAD, 0xBE, 0xEF},
	}

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, content := range parts {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("创建部件 %s: %v", name, err)
		}
		w.Write([]byte(content))
	}
	for name, data := range binaryParts {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("创建部件 %s: %v", name, err)
		}
		w.Write(data)
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("关闭 zip: %v", err)
	}
	return buf.Bytes()
}

// readParts 读取 .docx 字节流中的全部部件。
func readParts(t *testing.T, data []byte) map[string][]byte {
	t.Helper()

	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("读取 zip: %v", err)
	}

	out := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("打开部件 %s: %v", f.Name, err)
		}
		b, _ := io.ReadAll(rc)
		rc.Close()
		out[f.Name] = b
	}
	return out
}

const simpleBody = `<w:p><w:r><w:t>第一段</w:t></w:r></w:p><w:p><w:r><w:t>第二段包含 2023 年</w:t></w:r></w:p>`

// TestFidelityUntouchedPartsUnchanged 是保真编辑的核心保证：
// 修改正文之后，其余部件必须逐字节保持不变。
func TestFidelityUntouchedPartsUnchanged(t *testing.T) {
	original := buildDocxWithExtras(t, simpleBody)
	before := readParts(t, original)

	doc, err := OpenBytes(original)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	if n := doc.ReplaceText("2023", "2024"); n != 1 {
		t.Fatalf("期望替换 1 处，实际 %d 处", n)
	}

	saved, err := doc.SaveBytes()
	if err != nil {
		t.Fatalf("SaveBytes: %v", err)
	}
	after := readParts(t, saved)

	// 除正文外的所有部件都必须逐字节一致。
	for name, want := range before {
		if name == "word/document.xml" {
			continue
		}
		got, ok := after[name]
		if !ok {
			t.Errorf("部件在保存后丢失: %s", name)
			continue
		}
		if !bytes.Equal(want, got) {
			t.Errorf("部件 %s 在编辑无关内容后发生了变化", name)
		}
	}

	// 图片是最容易被破坏的部件，单独确认。
	if !bytes.Equal(before["word/media/image1.png"], after["word/media/image1.png"]) {
		t.Error("图片部件被破坏")
	}
}

func TestReplaceTextAppliesEdit(t *testing.T) {
	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	doc.ReplaceText("2023", "2024")

	text, err := doc.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if !strings.Contains(text, "2024") {
		t.Errorf("替换未生效，当前文本: %q", text)
	}
	if strings.Contains(text, "2023") {
		t.Errorf("原文本仍然存在: %q", text)
	}
}

// TestTextReassemblesSplitRuns 覆盖此前的缺陷：
// 早期实现把每个 <w:t> 当作一个段落，导致被拆分的句子返回时支离破碎。
func TestTextReassemblesSplitRuns(t *testing.T) {
	// Word 常因格式变化把一句话拆到多个 run 中。
	body := `<w:p><w:r><w:t>项目</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>进度</w:t></w:r><w:r><w:t>报告</w:t></w:r></w:p>`

	doc, err := OpenBytes(buildDocxWithExtras(t, body))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	paras := doc.ParagraphTexts()
	if len(paras) != 1 {
		t.Fatalf("期望 1 个段落，实际 %d 个: %q", len(paras), paras)
	}
	if paras[0] != "项目进度报告" {
		t.Errorf("段落文本应为 %q，实际为 %q", "项目进度报告", paras[0])
	}
}

// TestReplaceTextAcrossRuns 验证跨 run 边界的替换。
// 逐 run 替换无法命中这类文本，需要先在段落层面合并。
func TestReplaceTextAcrossRuns(t *testing.T) {
	body := `<w:p><w:r><w:t>项目</w:t></w:r><w:r><w:t>进度</w:t></w:r><w:r><w:t>报告</w:t></w:r></w:p>`

	doc, err := OpenBytes(buildDocxWithExtras(t, body))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// 逐 run 替换无法匹配跨边界的文本。
	if n := doc.ReplaceText("项目进度", "年度总结"); n != 0 {
		t.Errorf("ReplaceText 不应匹配跨 run 的文本，却替换了 %d 处", n)
	}

	if n := doc.ReplaceTextAcrossRuns("项目进度", "年度总结"); n != 1 {
		t.Fatalf("期望改写 1 个段落，实际 %d 个", n)
	}

	paras := doc.ParagraphTexts()
	if len(paras) != 1 || paras[0] != "年度总结报告" {
		t.Errorf("期望 %q，实际 %q", "年度总结报告", paras)
	}
}

func TestSetParagraphText(t *testing.T) {
	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	if err := doc.SetParagraphText(0, "全新的第一段"); err != nil {
		t.Fatalf("SetParagraphText: %v", err)
	}

	paras := doc.ParagraphTexts()
	if paras[0] != "全新的第一段" {
		t.Errorf("第一段应为 %q，实际为 %q", "全新的第一段", paras[0])
	}
	// 其余段落不应受影响。
	if !strings.Contains(paras[1], "第二段") {
		t.Errorf("第二段被意外修改: %q", paras[1])
	}
}

func TestSetParagraphTextRejectsOutOfRange(t *testing.T) {
	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	if err := doc.SetParagraphText(99, "x"); err == nil {
		t.Error("下标越界时应返回错误")
	}
	if err := doc.SetParagraphText(-1, "x"); err == nil {
		t.Error("负数下标应返回错误")
	}
}

// TestXMLSpecialCharsAreEscaped 确保插入的文本不会破坏 XML 结构。
func TestXMLSpecialCharsAreEscaped(t *testing.T) {
	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	if err := doc.SetParagraphText(0, `a < b & c > d`); err != nil {
		t.Fatalf("SetParagraphText: %v", err)
	}

	saved, err := doc.SaveBytes()
	if err != nil {
		t.Fatalf("SaveBytes: %v", err)
	}

	// 保存后的文档必须能被重新解析，说明 XML 未被破坏。
	reopened, err := OpenBytes(saved)
	if err != nil {
		t.Fatalf("特殊字符导致文档损坏: %v", err)
	}

	if got := reopened.ParagraphTexts()[0]; got != `a < b & c > d` {
		t.Errorf("特殊字符往返后不一致: %q", got)
	}
}

// TestRoundTripStability 验证反复读取保存不会导致内容漂移。
func TestRoundTripStability(t *testing.T) {
	data := buildDocxWithExtras(t, simpleBody)

	for i := 0; i < 3; i++ {
		doc, err := OpenBytes(data)
		if err != nil {
			t.Fatalf("第 %d 轮 OpenBytes: %v", i, err)
		}
		next, err := doc.SaveBytes()
		if err != nil {
			t.Fatalf("第 %d 轮 SaveBytes: %v", i, err)
		}
		data = next
	}

	doc, err := OpenBytes(data)
	if err != nil {
		t.Fatalf("最终 OpenBytes: %v", err)
	}
	paras := doc.ParagraphTexts()
	if len(paras) != 2 {
		t.Errorf("多轮往返后段落数发生变化: %d，内容 %q", len(paras), paras)
	}
}

func TestSaveAsWritesFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "out.docx")

	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	doc.ReplaceText("2023", "2024")

	if err := doc.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	saved, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("读取保存的文件: %v", err)
	}
	reopened, err := OpenBytes(saved)
	if err != nil {
		t.Fatalf("重新打开保存的文件: %v", err)
	}
	if !strings.Contains(strings.Join(reopened.ParagraphTexts(), "\n"), "2024") {
		t.Error("编辑内容未写入文件")
	}
}

// TestNewModeStillWorks 确认保真编辑的加入没有破坏原有的新建流程。
func TestNewModeStillWorks(t *testing.T) {
	doc := New()
	doc.AddHeading("标题", 1)
	doc.AddParagraph("正文")

	if doc.IsOpened() {
		t.Error("New 创建的文档不应处于打开模式")
	}

	text, err := doc.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if !strings.Contains(text, "标题") || !strings.Contains(text, "正文") {
		t.Errorf("新建模式下文本提取有误: %q", text)
	}

	// 新建模式下 SaveBytes 应回退到原有的生成逻辑。
	data, err := doc.SaveBytes()
	if err != nil {
		t.Fatalf("SaveBytes: %v", err)
	}
	if len(data) == 0 {
		t.Error("新建模式生成的文档为空")
	}
}

// TestEditingRequiresOpenMode 验证编辑方法在新建模式下不会误操作。
func TestEditingRequiresOpenMode(t *testing.T) {
	doc := New()
	doc.AddParagraph("内容")

	if n := doc.ReplaceText("内容", "其他"); n != 0 {
		t.Errorf("新建模式下 ReplaceText 应为空操作，却替换了 %d 处", n)
	}
	if err := doc.SetParagraphText(0, "x"); err == nil {
		t.Error("新建模式下 SetParagraphText 应返回错误")
	}
}

func TestOpenRejectsNonDocx(t *testing.T) {
	if _, err := OpenBytes([]byte("这不是一个 docx 文件")); err == nil {
		t.Error("非 ZIP 数据应返回错误")
	}
}

func TestMetadataIsRead(t *testing.T) {
	doc, err := OpenBytes(buildDocxWithExtras(t, simpleBody))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	if doc.title != "原始标题" {
		t.Errorf("标题读取错误: %q", doc.title)
	}
	if doc.author != "原作者" {
		t.Errorf("作者读取错误: %q", doc.author)
	}
}
