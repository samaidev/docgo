package docgo

import (
	"encoding/binary"
	"errors"
	"strings"
	"testing"

	"github.com/samaidev/officecore"
)

// buildFIB 构造一个最小的文件信息块。
func buildFIB(useTable1 bool, fcMin, ccpText, fcClx, lcbClx int) []byte {
	wd := make([]byte, 0x800)

	binary.LittleEndian.PutUint16(wd[0:], 0xA5EC) // 魔数
	var flags uint16
	if useTable1 {
		flags |= 0x0200
	}
	binary.LittleEndian.PutUint16(wd[0x0A:], flags)
	binary.LittleEndian.PutUint32(wd[0x18:], uint32(fcMin))
	binary.LittleEndian.PutUint32(wd[0x4C:], uint32(ccpText))
	binary.LittleEndian.PutUint32(wd[0x01A2:], uint32(fcClx))
	binary.LittleEndian.PutUint32(wd[0x01A6:], uint32(lcbClx))

	return wd
}

// buildPieceTable 构造 CLX 中的 piece table 部分。
//
// 布局为 n+1 个字符位置，随后是 n 个 8 字节描述符。
func buildPieceTable(cps []int, fcs []uint32) []byte {
	var out []byte
	for _, cp := range cps {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, uint32(cp))
		out = append(out, b...)
	}
	for _, fc := range fcs {
		pcd := make([]byte, 8)
		binary.LittleEndian.PutUint32(pcd[2:], fc)
		out = append(out, pcd...)
	}
	return out
}

// wrapPcdt 将 piece table 包装为一条 Pcdt 条目。
func wrapPcdt(pt []byte) []byte {
	out := []byte{2}
	length := make([]byte, 4)
	binary.LittleEndian.PutUint32(length, uint32(len(pt)))
	out = append(out, length...)
	return append(out, pt...)
}

func TestParseFIBReadsFields(t *testing.T) {
	wd := buildFIB(true, 0x800, 100, 50, 60)

	f, err := parseFIB(wd)
	if err != nil {
		t.Fatalf("parseFIB: %v", err)
	}

	if !f.whichTableStream {
		t.Error("应识别为使用 1Table")
	}
	if f.fcMin != 0x800 {
		t.Errorf("fcMin = %d，期望 0x800", f.fcMin)
	}
	if f.ccpText != 100 {
		t.Errorf("ccpText = %d，期望 100", f.ccpText)
	}
	if f.fcClx != 50 || f.lcbClx != 60 {
		t.Errorf("CLX 位置解析错误: fc=%d lcb=%d", f.fcClx, f.lcbClx)
	}
}

func TestParseFIBRejectsBadMagic(t *testing.T) {
	wd := make([]byte, 0x800)
	binary.LittleEndian.PutUint16(wd[0:], 0x1234)

	if _, err := parseFIB(wd); err == nil {
		t.Error("魔数错误时应返回错误")
	}
}

func TestParseFIBRejectsTruncated(t *testing.T) {
	if _, err := parseFIB(make([]byte, 100)); err == nil {
		t.Error("FIB 被截断时应返回错误")
	}
}

// TestDecodePieceTableCompressedFlag 覆盖 piece table 中最易错的一点：
// 压缩标志位不仅指明编码方式，还意味着偏移量需要除以 2。
func TestDecodePieceTableCompressedFlag(t *testing.T) {
	// 两个片段：第一个为压缩形式，第二个为 UTF-16。
	pt := buildPieceTable(
		[]int{0, 5, 10},
		[]uint32{0x40000000 | 0x1000, 0x2000},
	)

	pieces, err := decodePieceTable(pt)
	if err != nil {
		t.Fatalf("decodePieceTable: %v", err)
	}

	if len(pieces) != 2 {
		t.Fatalf("期望 2 个片段，实际 %d 个", len(pieces))
	}

	if !pieces[0].compressed {
		t.Error("第一个片段应为压缩形式")
	}
	if pieces[0].fc != 0x1000/2 {
		t.Errorf("压缩片段的偏移应为 %d，实际 %d", 0x1000/2, pieces[0].fc)
	}

	if pieces[1].compressed {
		t.Error("第二个片段不应为压缩形式")
	}
	if pieces[1].fc != 0x2000 {
		t.Errorf("非压缩片段的偏移应为 %d，实际 %d", 0x2000, pieces[1].fc)
	}

	if pieces[0].cpStart != 0 || pieces[0].cpEnd != 5 {
		t.Errorf("第一个片段的字符区间错误: [%d,%d)", pieces[0].cpStart, pieces[0].cpEnd)
	}
}

func TestDecodePieceTableRejectsShortInput(t *testing.T) {
	if _, err := decodePieceTable([]byte{1, 2, 3}); err == nil {
		t.Error("输入过短时应返回错误")
	}
}

// TestParsePieceTableSkipsPrc 验证类型 1 的 Prc 条目被正确跳过。
func TestParsePieceTableSkipsPrc(t *testing.T) {
	pt := buildPieceTable([]int{0, 3}, []uint32{0x40000000 | 0x200})

	// 在 Pcdt 之前放置一条 Prc 条目。
	clx := []byte{1, 4, 0, 0xAA, 0xBB, 0xCC, 0xDD}
	clx = append(clx, wrapPcdt(pt)...)

	pieces, err := parsePieceTable(clx)
	if err != nil {
		t.Fatalf("parsePieceTable: %v", err)
	}
	if len(pieces) != 1 {
		t.Fatalf("期望 1 个片段，实际 %d 个", len(pieces))
	}
	if pieces[0].fc != 0x100 {
		t.Errorf("偏移解析错误: %d", pieces[0].fc)
	}
}

func TestParsePieceTableRejectsUnknownEntry(t *testing.T) {
	if _, err := parsePieceTable([]byte{99, 0, 0}); err == nil {
		t.Error("未知条目类型应返回错误")
	}
}

// TestExtractTextViaPieceTableCompressed 端到端验证压缩片段的解码。
func TestExtractTextViaPieceTableCompressed(t *testing.T) {
	const textOffset = 0x400
	content := "Hello World"

	wd := buildFIB(true, 0x800, len(content), 0, 0)
	copy(wd[textOffset:], content)

	// 压缩片段的偏移需要以 2 倍形式写入。
	pt := buildPieceTable(
		[]int{0, len(content)},
		[]uint32{0x40000000 | uint32(textOffset*2)},
	)
	clx := wrapPcdt(pt)

	f := &fib{ccpText: len(content), fcClx: 0, lcbClx: len(clx)}

	got, err := extractTextViaPieceTable(wd, clx, f)
	if err != nil {
		t.Fatalf("extractTextViaPieceTable: %v", err)
	}
	if got != content {
		t.Errorf("期望 %q，实际 %q", content, got)
	}
}

// TestExtractTextViaPieceTableUTF16 验证 UTF-16 片段的解码，
// 中文内容必须正确还原。
func TestExtractTextViaPieceTableUTF16(t *testing.T) {
	const textOffset = 0x400
	content := "季度报告"

	wd := buildFIB(true, 0x800, len([]rune(content)), 0, 0)
	for i, r := range []rune(content) {
		binary.LittleEndian.PutUint16(wd[textOffset+i*2:], uint16(r))
	}

	pt := buildPieceTable(
		[]int{0, len([]rune(content))},
		[]uint32{uint32(textOffset)},
	)
	clx := wrapPcdt(pt)

	f := &fib{ccpText: len([]rune(content)), fcClx: 0, lcbClx: len(clx)}

	got, err := extractTextViaPieceTable(wd, clx, f)
	if err != nil {
		t.Fatalf("extractTextViaPieceTable: %v", err)
	}
	if got != content {
		t.Errorf("期望 %q，实际 %q", content, got)
	}
}

// TestExtractTextJoinsMultiplePieces 验证多个片段被按序拼接。
// 这正是 piece table 存在的意义：正文在流中并非连续存放。
func TestExtractTextJoinsMultiplePieces(t *testing.T) {
	wd := buildFIB(true, 0x800, 10, 0, 0)
	copy(wd[0x400:], "AAA")
	copy(wd[0x600:], "BBB")

	pt := buildPieceTable(
		[]int{0, 3, 6},
		[]uint32{
			0x40000000 | uint32(0x400*2),
			0x40000000 | uint32(0x600*2),
		},
	)
	clx := wrapPcdt(pt)

	f := &fib{ccpText: 6, fcClx: 0, lcbClx: len(clx)}

	got, err := extractTextViaPieceTable(wd, clx, f)
	if err != nil {
		t.Fatalf("extractTextViaPieceTable: %v", err)
	}
	if got != "AAABBB" {
		t.Errorf("片段拼接错误，期望 AAABBB，实际 %q", got)
	}
}

func TestCleanDocText(t *testing.T) {
	cases := map[string]string{
		"段落一\r段落二":      "段落一\n段落二",
		"单元格\a分隔":       "单元格\t分隔",
		"软\v换行":         "软\n换行",
		"含\x01图片锚点":     "含图片锚点",
		"含\x05控制字符":     "含控制字符",
		"  首尾空白  ":      "首尾空白",
	}
	for in, want := range cases {
		if got := cleanDocText(in); got != want {
			t.Errorf("cleanDocText(%q) = %q，期望 %q", in, got, want)
		}
	}
}

// TestCleanDocTextCollapsesBlankLines 验证剔除控制字符后产生的空行被合并。
func TestCleanDocTextCollapsesBlankLines(t *testing.T) {
	got := cleanDocText("A\r\r\r\r\rB")
	if strings.Contains(got, "\n\n\n") {
		t.Errorf("多余空行未被合并: %q", got)
	}
}

// TestDecodeCP1252 验证 Windows-1252 特有区间的解码。
// 该区间包含中英文文档中常见的弯引号和破折号。
func TestDecodeCP1252(t *testing.T) {
	got := decodeCP1252([]byte{0x93, 'h', 'i', 0x94, 0x97})
	want := "“hi”—"

	if got != want {
		t.Errorf("decodeCP1252 = %q，期望 %q", got, want)
	}
}

func TestDecodeCP1252PassesThroughASCII(t *testing.T) {
	if got := decodeCP1252([]byte("plain text")); got != "plain text" {
		t.Errorf("ASCII 解码错误: %q", got)
	}
}

// TestLegacySaveReturnsReadOnlyError 验证尝试写入 .doc 会明确失败。
func TestLegacySaveReturnsReadOnlyError(t *testing.T) {
	d := &LegacyDocument{text: "内容"}

	err := d.Save("out.doc")
	if err == nil {
		t.Fatal("写入 .doc 应当返回错误")
	}
	if !errors.Is(err, officecore.ErrReadOnlyFormat) {
		t.Errorf("错误应可通过 errors.Is 匹配 ErrReadOnlyFormat，实际为: %v", err)
	}
}

func TestLegacyParagraphs(t *testing.T) {
	d := &LegacyDocument{text: "第一段\n第二段\n第三段"}

	paras := d.Paragraphs()
	if len(paras) != 3 {
		t.Fatalf("期望 3 个段落，实际 %d 个: %v", len(paras), paras)
	}
	if paras[1] != "第二段" {
		t.Errorf("第二段解析错误: %q", paras[1])
	}
}

// TestToDocxConvertsContent 验证转换为可编辑格式的路径可用，
// 这是修改 .doc 内容的推荐做法。
func TestToDocxConvertsContent(t *testing.T) {
	legacy := &LegacyDocument{text: "第一段\n\n第二段"}

	doc := legacy.ToDocx()

	text, err := doc.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if !strings.Contains(text, "第一段") || !strings.Contains(text, "第二段") {
		t.Errorf("内容迁移不完整: %q", text)
	}

	// 生成的文档应当可以保存。
	if _, err := doc.SaveBytes(); err != nil {
		t.Errorf("转换后的文档无法保存: %v", err)
	}
}

func TestOpenLegacyRejectsNonCFB(t *testing.T) {
	if _, err := OpenLegacyBytes([]byte("PK\x03\x04 这是一个 zip")); err == nil {
		t.Error("将 docx 作为 doc 打开应返回错误")
	}
}

// TestFallbackTextWhenPieceTableMissing 验证退化路径：
// 缺少 piece table 时仍应尽可能返回内容，而不是彻底失败。
func TestFallbackTextWhenPieceTableMissing(t *testing.T) {
	wd := buildFIB(true, 0x400, 5, 0, 0)
	copy(wd[0x400:], "Hello")

	f := &fib{fcMin: 0x400, ccpText: 5}

	got := fallbackText(wd, f)
	if !strings.Contains(got, "Hello") {
		t.Errorf("退化路径未提取到文本: %q", got)
	}
}
