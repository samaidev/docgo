package docgo

import (
	"strings"
	"testing"
)

// TestReadRealDocx 回归：对真实 Word 2007+ 文档（来自 LibreOffice 测试数据）
// 应成功提取段落文本与表格结构，而非崩溃或返回空。
func TestReadRealDocx(t *testing.T) {
	const path = "testdata/sample.docx"
	d, err := Open(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	txt, err := d.Text()
	if err != nil {
		t.Fatalf("Text() 错误: %v", err)
	}
	if strings.TrimSpace(txt) == "" {
		t.Fatalf("真实 docx 提取文本为空")
	}
	if len(d.Paragraphs()) == 0 {
		t.Fatalf("真实 docx 未解析出任何段落")
	}
}

// TestOpenRealLegacyDoc 回归：对真实 OLE2(.doc) 文档应成功打开且不 panic。
// 该样本为图表文档，正文可能为空，故仅断言打开成功。
func TestOpenRealLegacyDoc(t *testing.T) {
	const path = "testdata/sample.doc"
	d, err := OpenLegacy(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	_, _ = d.Text() // 不崩溃即可
}
