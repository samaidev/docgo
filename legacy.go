package docgo

// 本文件实现 .doc（Word 97-2003）的只读解析。
//
// .doc 是三种旧版 Office 格式中最复杂的一种。正文并非连续存放，而是分散在
// WordDocument 流的各处，由一张名为 piece table 的索引负责拼接。读取步骤为：
//
//  1. 从 WordDocument 流开头读取 FIB（文件信息块），其中记录了表流的名称
//     以及 CLX 结构的位置。
//  2. 表流为 "1Table" 或 "0Table"，具体由 FIB 中的一个标志位决定。
//  3. CLX 中包含 piece table：一组字符位置，以及各片段在流中的偏移。
//  4. 每个片段既可能是 UTF-16，也可能是单字节压缩形式，由偏移量中的一个
//     标志位区分。逐片段解码后拼接，才得到完整正文。
//
// 之所以只读，是因为回写需要重建 piece table、字符与段落的格式映射，
// 以及真实文件中大量未公开的结构。需要修改时，请读取文本后写入 .docx。

import (
	"encoding/binary"
	"fmt"
	"strings"
	"unicode/utf16"

	"github.com/samaidev/officecore"
	"github.com/samaidev/officecore/cfb"
)

// LegacyDocument 表示一个只读的 .doc 文档。
type LegacyDocument struct {
	text string
}

// OpenLegacy 读取一个 .doc 文件。
//
// 返回的文档是只读的。若需要修改，请读出文本后用 New 创建 .docx。
func OpenLegacy(path string) (*LegacyDocument, error) {
	c, err := cfb.Open(path)
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	return parseLegacyDoc(c)
}

// OpenLegacyBytes 从内存中的字节流读取 .doc。
func OpenLegacyBytes(data []byte) (*LegacyDocument, error) {
	c, err := cfb.OpenReader(newByteReaderAt(data))
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	return parseLegacyDoc(c)
}

// parseLegacyDoc 从复合文档中提取正文。
func parseLegacyDoc(c *cfb.Container) (*LegacyDocument, error) {
	wd, err := c.MustStream("WordDocument")
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	if len(wd) < 512 {
		return nil, fmt.Errorf("docgo: WordDocument 流过短，文件可能已损坏")
	}

	fib, err := parseFIB(wd)
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}

	// 表流的选择由 FIB 中的 fWhichTblStm 标志位决定。
	tableName := "0Table"
	if fib.whichTableStream {
		tableName = "1Table"
	}

	table, ok := c.Stream(tableName)
	if !ok {
		// 部分早期文件不含表流，此时退化为直接读取正文区间。
		return &LegacyDocument{text: fallbackText(wd, fib)}, nil
	}

	text, err := extractTextViaPieceTable(wd, table, fib)
	if err != nil {
		// piece table 解析失败时退化处理，返回部分内容优于完全失败。
		return &LegacyDocument{text: fallbackText(wd, fib)}, nil
	}

	return &LegacyDocument{text: text}, nil
}

// fib 保存本解析器需要用到的 FIB 字段。
type fib struct {
	// whichTableStream 指明使用 1Table 还是 0Table。
	whichTableStream bool

	// ccpText 为正文的字符数。
	ccpText int

	// fcClx、lcbClx 为 CLX 结构在表流中的偏移与长度。
	fcClx  int
	lcbClx int

	// fcMin 为正文在 WordDocument 流中的起始偏移（退化路径使用）。
	fcMin int
}

// parseFIB 解析文件信息块。
//
// FIB 位于 WordDocument 流的开头，是一个字段偏移固定的结构。
// 这里只读取定位正文所必需的少数字段。
func parseFIB(wd []byte) (*fib, error) {
	if len(wd) < 0x200 {
		return nil, fmt.Errorf("FIB 被截断")
	}

	// 魔数 0xA5EC 标识 Word 97 及更高版本的文档。
	magic := binary.LittleEndian.Uint16(wd[0:])
	if magic != 0xA5EC {
		return nil, fmt.Errorf("不是 Word 97-2003 文档（魔数为 %#04x）", magic)
	}

	f := &fib{}

	// 偏移 0x000A 处的标志位字段，第 0 位为 fEncrypted，第 9 位为 fWhichTblStm。
	flags := binary.LittleEndian.Uint16(wd[0x0A:])
	if flags&0x0001 != 0 {
		return nil, fmt.Errorf("docgo: 不支持读取加密（受密码保护）的 .doc 文件")
	}
	f.whichTableStream = flags&0x0200 != 0

	// 偏移 0x0018 为 fcMin，正文起始位置。
	f.fcMin = int(binary.LittleEndian.Uint32(wd[0x18:]))

	// 偏移 0x004C 为 ccpText，正文字符数。
	f.ccpText = int(binary.LittleEndian.Uint32(wd[0x4C:]))

	// 偏移 0x01A2 为 fcClx，0x01A6 为 lcbClx。
	f.fcClx = int(binary.LittleEndian.Uint32(wd[0x01A2:]))
	f.lcbClx = int(binary.LittleEndian.Uint32(wd[0x01A6:]))

	return f, nil
}

// piece 描述正文的一个片段。
type piece struct {
	// cpStart、cpEnd 为该片段在逻辑正文中的字符区间。
	cpStart, cpEnd int

	// fc 为该片段在 WordDocument 流中的字节偏移。
	fc int

	// compressed 表示该片段采用单字节存储而非 UTF-16。
	compressed bool
}

// extractTextViaPieceTable 依据 piece table 重建正文。
func extractTextViaPieceTable(wd, table []byte, f *fib) (string, error) {
	if f.lcbClx <= 0 || f.fcClx < 0 || f.fcClx+f.lcbClx > len(table) {
		return "", fmt.Errorf("CLX 位置无效")
	}

	clx := table[f.fcClx : f.fcClx+f.lcbClx]

	pieces, err := parsePieceTable(clx)
	if err != nil {
		return "", err
	}
	if len(pieces) == 0 {
		return "", fmt.Errorf("piece table 为空")
	}

	var sb strings.Builder
	for _, p := range pieces {
		n := p.cpEnd - p.cpStart
		if n <= 0 {
			continue
		}

		if p.compressed {
			// 单字节：每个字符占 1 字节，采用 Windows-1252 编码。
			if p.fc+n > len(wd) {
				n = len(wd) - p.fc
			}
			if n <= 0 {
				continue
			}
			sb.WriteString(decodeCP1252(wd[p.fc : p.fc+n]))
		} else {
			// UTF-16LE：每个字符占 2 字节。
			byteLen := n * 2
			if p.fc+byteLen > len(wd) {
				byteLen = ((len(wd) - p.fc) / 2) * 2
			}
			if byteLen <= 0 {
				continue
			}
			u16 := make([]uint16, byteLen/2)
			for i := range u16 {
				u16[i] = binary.LittleEndian.Uint16(wd[p.fc+i*2:])
			}
			sb.WriteString(string(utf16.Decode(u16)))
		}
	}

	return cleanDocText(sb.String()), nil
}

// parsePieceTable 从 CLX 结构中解析出各片段。
//
// CLX 由若干条目组成：类型 1 为 Prc（属性修饰，需跳过），
// 类型 2 为 Pcdt，其中才是 piece table 本身。
func parsePieceTable(clx []byte) ([]piece, error) {
	i := 0
	for i < len(clx) {
		switch clx[i] {
		case 1:
			// Prc：1 字节类型 + 2 字节长度 + 负载，跳过。
			if i+3 > len(clx) {
				return nil, fmt.Errorf("Prc 条目被截断")
			}
			n := int(binary.LittleEndian.Uint16(clx[i+1:]))
			i += 3 + n

		case 2:
			// Pcdt：1 字节类型 + 4 字节长度 + piece table 数据。
			if i+5 > len(clx) {
				return nil, fmt.Errorf("Pcdt 条目被截断")
			}
			n := int(binary.LittleEndian.Uint32(clx[i+1:]))
			i += 5
			if i+n > len(clx) {
				n = len(clx) - i
			}
			return decodePieceTable(clx[i : i+n])

		default:
			return nil, fmt.Errorf("CLX 中存在未知条目类型 %d", clx[i])
		}
	}

	return nil, fmt.Errorf("CLX 中未找到 piece table")
}

// decodePieceTable 解析 PLCF 结构。
//
// 其布局为：n+1 个 4 字节字符位置，随后是 n 个 8 字节的片段描述符。
// 由此可反推出片段数量 n。
func decodePieceTable(data []byte) ([]piece, error) {
	const cpSize = 4
	const pcdSize = 8

	if len(data) < cpSize+pcdSize {
		return nil, fmt.Errorf("piece table 过短")
	}

	// len = (n+1)*4 + n*8  =>  n = (len - 4) / 12
	n := (len(data) - cpSize) / (cpSize + pcdSize)
	if n <= 0 {
		return nil, fmt.Errorf("piece table 中没有片段")
	}

	cps := make([]int, n+1)
	for i := 0; i <= n; i++ {
		cps[i] = int(binary.LittleEndian.Uint32(data[i*cpSize:]))
	}

	pcdBase := (n + 1) * cpSize
	pieces := make([]piece, 0, n)

	for i := 0; i < n; i++ {
		off := pcdBase + i*pcdSize
		if off+pcdSize > len(data) {
			break
		}

		// 描述符的第 2..5 字节为 fc 字段。
		fcRaw := binary.LittleEndian.Uint32(data[off+2:])

		// 第 30 位为压缩标志：置位表示单字节存储，
		// 且真实偏移需要除以 2。
		compressed := fcRaw&0x40000000 != 0
		fc := int(fcRaw & 0x3FFFFFFF)
		if compressed {
			fc /= 2
		}

		pieces = append(pieces, piece{
			cpStart:    cps[i],
			cpEnd:      cps[i+1],
			fc:         fc,
			compressed: compressed,
		})
	}

	return pieces, nil
}

// fallbackText 在 piece table 不可用时直接读取正文区间。
//
// 这样得到的结果不如 piece table 路径准确——可能包含已删除的文本，
// 或漏掉部分内容——但总比完全失败要好。
func fallbackText(wd []byte, f *fib) string {
	start := f.fcMin
	if start <= 0 || start >= len(wd) {
		start = 0x800 // 正文常见的起始位置
	}
	if start >= len(wd) {
		return ""
	}

	end := start + f.ccpText
	if end > len(wd) || f.ccpText <= 0 {
		end = len(wd)
	}

	return cleanDocText(decodeCP1252(wd[start:end]))
}

// cleanDocText 将 Word 的内部字符转换为常规文本。
//
// Word 用 \r 表示段落结束，并以若干控制字符标记单元格边界、
// 图片锚点、域代码等内容，这些都需要转换或剔除。
// cleanDocText 清理 Word 97-2003 的二进制正文字符流。
//
// 借鉴 LibreOffice ww8 对控制字符的处理：
//   - 0x0D 段落结束 → 换行
//   - 0x07 单元格分隔 → 制表符
//   - 0x0B 软换行 → 换行
//   - 0x09 制表符 → 原样保留
//   - 0xA0 不间断空格 → 普通空格（Word 用 0xA0 表示 Non-breaking space）
//   - 0x01/0x02 图片锚点与脚注引用 → 忽略
//   - 字段 0x13(begin)/0x14(separate)/0x15(end)：跳过 begin..separate 之间
//     的指令文本，仅保留 separate..end 之间的结果文本（如目录、页码域的结果）。
//   - 其余控制字符与符号字体私用区码点一律剔除。
func cleanDocText(s string) string {
	var sb strings.Builder
	sb.Grow(len(s))

	// 字段状态：0=普通文本；1=在字段指令内（0x13 后、0x14 前）；2=在字段结果内。
	fieldState := 0
	for _, r := range s {
		switch {
		case r == 0x13: // field begin
			fieldState = 1
		case r == 0x14: // field separate（结果开始）
			fieldState = 2
		case r == 0x15: // field end
			fieldState = 0
		case fieldState == 1:
			// 字段指令文本（如 "PAGE"、"TOC"），无输出价值。
		case r == '\r':
			sb.WriteRune('\n') // 段落结束
		case r == '\a':
			sb.WriteRune('\t') // 单元格分隔
		case r == '\v':
			sb.WriteRune('\n') // 软换行
		case r == 0xA0:
			sb.WriteRune(' ') // 不间断空格
		case r == '\t' || r == '\n':
			sb.WriteRune(r)
		case r == 0x01 || r == 0x02:
			// 图片锚点与脚注引用，不含文本内容。
		case r < 0x20:
			// 其余控制字符一律剔除。
		case r >= 0xF000 && r <= 0xF0FF:
			// 符号字体的私用区码点，无法还原为有意义的文本。
		default:
			sb.WriteRune(r)
		}
	}

	// 合并因剔除控制字符而产生的多余空行。
	text := sb.String()
	for strings.Contains(text, "\n\n\n") {
		text = strings.ReplaceAll(text, "\n\n\n", "\n\n")
	}

	return strings.TrimSpace(text)
}

// cp1252Extra 保存 Windows-1252 中 0x80-0x9F 区间与 Latin-1 不同的映射。
// 该区间在 Latin-1 中为控制字符，Windows-1252 则用于排版符号，
// 引号、破折号等常见字符都落在其中。
var cp1252Extra = map[byte]rune{
	0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…',
	0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8A: 'Š',
	0x8B: '‹', 0x8C: 'Œ', 0x8E: 'Ž', 0x91: '\u2018', 0x92: '\u2019',
	0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—',
	0x98: '˜', 0x99: '™', 0x9A: 'š', 0x9B: '›', 0x9C: 'œ',
	0x9E: 'ž', 0x9F: 'Ÿ',
}

// decodeCP1252 将 Windows-1252 字节序列解码为字符串。
func decodeCP1252(b []byte) string {
	var sb strings.Builder
	sb.Grow(len(b))

	for _, c := range b {
		if r, ok := cp1252Extra[c]; ok {
			sb.WriteRune(r)
			continue
		}
		sb.WriteRune(rune(c))
	}
	return sb.String()
}

// === 公开 API ===

// Text 返回文档的正文内容。
func (d *LegacyDocument) Text() (string, error) {
	return d.text, nil
}

// Paragraphs 返回按段落切分的正文。
func (d *LegacyDocument) Paragraphs() []string {
	if d.text == "" {
		return nil
	}
	return strings.Split(d.text, "\n")
}

// Save 始终返回 ErrReadOnlyFormat。
//
// 该方法的存在是为了让"尝试保存 .doc"这一行为明确失败，
// 而不是让调用方误以为修改已经生效。
func (d *LegacyDocument) Save(path string) error {
	return fmt.Errorf("docgo: 无法写入 .doc: %w", officecore.ErrReadOnlyFormat)
}

// ToDocx 将只读文档的内容复制到一个可编辑的 .docx 文档中。
//
// 这是修改 .doc 内容的推荐途径：先转为现代格式，再进行编辑。
// 需要注意的是，仅正文文本会被迁移，格式、图片和表格结构不会保留。
func (d *LegacyDocument) ToDocx() *Document {
	doc := New()
	for _, para := range d.Paragraphs() {
		if strings.TrimSpace(para) == "" {
			continue
		}
		doc.AddParagraph(para)
	}
	return doc
}

// byteReaderAt 让字节切片满足 io.ReaderAt 接口。
type byteReaderAt struct{ b []byte }

func newByteReaderAt(b []byte) *byteReaderAt { return &byteReaderAt{b} }

func (r *byteReaderAt) ReadAt(p []byte, off int64) (int, error) {
	if off < 0 || off >= int64(len(r.b)) {
		return 0, fmt.Errorf("offset %d out of range", off)
	}
	n := copy(p, r.b[off:])
	if n < len(p) {
		return n, fmt.Errorf("EOF")
	}
	return n, nil
}
