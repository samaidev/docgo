package docgo

// 本文件实现 .docx 的保真编辑。
//
// 背景：Open 早期的实现会把文档解析成本库自己的元素模型，保存时再重新
// 生成。这意味着任何未被建模的内容——图片、页眉页脚、样式、批注、
// 修订记录、嵌入字体——都会在一次读取-保存往返中丢失。对"编辑已有文档"
// 这一诉求而言，这是不可接受的。
//
// 现在的做法是：Open 会完整保留原始 ZIP 包，并把 word/document.xml 当作
// 可原地修改的文本处理。保存时仅重写实际改动过的部件，其余部件按原始
// 字节复制。用户没有触碰的内容，就一定不会发生变化。
//
// 文档因此有两种模式：
//   - 新建模式（New）：内容来自 elements 切片，保存时生成全新的包。
//   - 打开模式（Open）：内容来自 src，保存时改写原包。
// 二者行为不同，且都由 Save 正确处理。

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/samaidev/officecore/ooxml"
)

// documentPart 是 .docx 中承载正文的部件路径。
const documentPart = "word/document.xml"

// OpenFile 以保真模式打开一个 .docx 文件。
//
// 与早期的 Open 不同，它会保留文档的全部内容。修改文本后调用 Save，
// 未被改动的部分——格式、图片、页眉页脚、样式、批注——都原样保留。
//
//	doc, err := docgo.OpenFile("report.docx")
//	doc.ReplaceText("2023", "2024")
//	doc.Save("report-updated.docx")
func OpenFile(path string) (*Document, error) {
	pkg, err := ooxml.Open(path)
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	return fromPackage(pkg)
}

// OpenBytes 以保真模式从内存中的字节流打开 .docx。
func OpenBytes(data []byte) (*Document, error) {
	pkg, err := ooxml.OpenBytes(data)
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	return fromPackage(pkg)
}

// fromPackage 由已加载的包构造处于打开模式的文档。
func fromPackage(pkg *ooxml.Package) (*Document, error) {
	if !pkg.Has(documentPart) {
		return nil, fmt.Errorf("docgo: 缺少 %s，这不是一个 Word 文档", documentPart)
	}

	d := New()
	d.src = pkg
	d.body = string(pkg.Part(documentPart))

	// 元数据仍解析到结构体字段中，便于读取；
	// 但在打开模式下它们不参与保存，除非被显式修改。
	if core, ok := pkg.PartString("docProps/core.xml"); ok {
		d.title = extractTagText(core, "dc:title")
		d.author = extractTagText(core, "dc:creator")
		d.subject = extractTagText(core, "dc:subject")
	}

	return d, nil
}

// IsOpened 报告文档是否由已有文件打开（而非新建）。
//
// 打开模式的文档保存时保留原始内容；新建模式的文档则从元素模型生成。
func (d *Document) IsOpened() bool {
	return d.src != nil
}

// Text 返回文档的纯文本内容，段落之间以换行分隔。
//
// 一个逻辑段落经常被 Word 拆分到多个 run 中——修订标记、拼写检查状态、
// 局部格式变化都会造成拆分。这里按段落聚合，因此返回的是完整句子，
// 而不是支离破碎的片段。
func (d *Document) Text() (string, error) {
	if d.src == nil {
		// 新建模式：从元素模型收集文本。
		return d.textFromElements(), nil
	}

	blocks := ooxml.ExtractTextBlocks([]byte(d.body), "p", "t")
	return strings.Join(blocks, "\n"), nil
}

// ParagraphTexts 返回文档中每个段落的文本，顺序与文档一致。
// 空段落会保留为空字符串，以保证下标位置稳定。
//
// 该方法与 Paragraphs 不同：后者返回新建模式下的段落对象，
// 而本方法返回文本，且在两种模式下都可用。
func (d *Document) ParagraphTexts() []string {
	if d.src == nil {
		return strings.Split(d.textFromElements(), "\n")
	}
	return ooxml.ExtractTextBlocks([]byte(d.body), "p", "t")
}

// textFromElements 在新建模式下从元素模型提取文本。
func (d *Document) textFromElements() string {
	var lines []string
	for _, e := range d.elements {
		switch v := e.(type) {
		case *Heading:
			lines = append(lines, v.Text)
		case *Paragraph:
			var sb strings.Builder
			for _, r := range v.runs {
				sb.WriteString(r.text)
			}
			lines = append(lines, sb.String())
		case *List:
			lines = append(lines, v.items...)
		case *CodeBlock:
			lines = append(lines, v.code)
		case *Table:
			for _, row := range v.rows {
				lines = append(lines, strings.Join(row, "\t"))
			}
		}
	}
	return strings.Join(lines, "\n")
}

// wtRe 匹配 <w:t> 元素及其文本内容，同时捕获可能存在的属性
// （例如 xml:space="preserve"）。
var wtRe = regexp.MustCompile(`(?s)<w:t(\s[^>]*)?>(.*?)</w:t>`)

// ReplaceText 将文档中所有 old 替换为 new，返回替换次数。
//
// 该操作仅对打开模式的文档有效；新建模式的文档应直接修改元素模型。
//
// 替换是逐 run 进行的，因此不会破坏格式。需要注意的是：如果被查找的文本
// 恰好跨越了多个 run 的边界，本方法不会匹配到它。这类情况请使用
// ReplaceTextAcrossRuns，代价是被替换的文本会统一采用其所在的第一个 run
// 的格式。
func (d *Document) ReplaceText(old, new string) int {
	if d.src == nil || old == "" {
		return 0
	}

	count := 0
	escapedOld := ooxml.EscapeXMLText(old)
	escapedNew := ooxml.EscapeXMLText(new)

	d.body = wtRe.ReplaceAllStringFunc(d.body, func(m string) string {
		sub := wtRe.FindStringSubmatch(m)
		if sub == nil {
			return m
		}
		attrs, text := sub[1], sub[2]

		n := strings.Count(text, escapedOld)
		if n == 0 {
			return m
		}
		count += n

		replaced := strings.ReplaceAll(text, escapedOld, escapedNew)

		// 若替换后的文本首尾出现空白，需保留空白，
		// 否则 Word 会将其裁掉。
		if attrs == "" && needsSpacePreserve(replaced) {
			attrs = ` xml:space="preserve"`
		}
		return "<w:t" + attrs + ">" + replaced + "</w:t>"
	})

	return count
}

// paragraphRe 匹配单个 <w:p> 元素（含自闭合形式）。
var paragraphRe = regexp.MustCompile(`(?s)<w:p(?:\s[^>]*)?(?:/>|>.*?</w:p>)`)

// ReplaceTextAcrossRuns 替换文本，且能匹配跨越多个 run 边界的内容。
//
// Word 会因修订标记、拼写检查状态或格式变化而把一句话拆散到多个 run 中，
// 此时 ReplaceText 会漏掉这类文本。本方法在段落层面先把各 run 的文字
// 拼接起来匹配，再将结果写回。
//
// 代价是格式会被压平：被改写的段落中，所有文本统一采用其第一个 run 的格式。
// 因此仅在 ReplaceText 无法命中目标时才使用它。返回值为发生改写的段落数。
func (d *Document) ReplaceTextAcrossRuns(old, new string) int {
	if d.src == nil || old == "" {
		return 0
	}

	count := 0
	escapedNew := ooxml.EscapeXMLText(new)

	d.body = paragraphRe.ReplaceAllStringFunc(d.body, func(para string) string {
		// 拼接本段落中所有 run 的文本。
		matches := wtRe.FindAllStringSubmatch(para, -1)
		if len(matches) == 0 {
			return para
		}

		var joined strings.Builder
		for _, m := range matches {
			joined.WriteString(unescapeXML(m[2]))
		}
		full := joined.String()

		if !strings.Contains(full, old) {
			return para
		}
		count++

		updated := strings.ReplaceAll(full, old, new)
		_ = escapedNew

		// 将全部文本写入第一个 <w:t>，其余清空。
		// 这样可保留第一个 run 的格式，并维持段落结构完整。
		first := true
		return wtRe.ReplaceAllStringFunc(para, func(m string) string {
			sub := wtRe.FindStringSubmatch(m)
			if sub == nil {
				return m
			}
			attrs := sub[1]
			if first {
				first = false
				if needsSpacePreserve(updated) && !strings.Contains(attrs, "xml:space") {
					attrs += ` xml:space="preserve"`
				}
				return "<w:t" + attrs + ">" + ooxml.EscapeXMLText(updated) + "</w:t>"
			}
			return "<w:t" + attrs + "></w:t>"
		})
	})

	return count
}

// SetParagraphText 替换第 index 个段落的全部文本，下标从 0 开始。
//
// 该段落第一个 run 的格式会被保留并应用于新文本。若下标越界则返回错误。
func (d *Document) SetParagraphText(index int, text string) error {
	if d.src == nil {
		return fmt.Errorf("docgo: SetParagraphText 仅适用于以 OpenFile 打开的文档")
	}

	paras := paragraphRe.FindAllStringIndex(d.body, -1)
	if index < 0 || index >= len(paras) {
		return fmt.Errorf("docgo: 段落下标 %d 越界（共 %d 个段落）", index, len(paras))
	}

	start, end := paras[index][0], paras[index][1]
	para := d.body[start:end]

	if !wtRe.MatchString(para) {
		return fmt.Errorf("docgo: 第 %d 个段落不含可编辑的文本 run", index)
	}

	first := true
	updated := wtRe.ReplaceAllStringFunc(para, func(m string) string {
		sub := wtRe.FindStringSubmatch(m)
		if sub == nil {
			return m
		}
		attrs := sub[1]
		if first {
			first = false
			if needsSpacePreserve(text) && !strings.Contains(attrs, "xml:space") {
				attrs += ` xml:space="preserve"`
			}
			return "<w:t" + attrs + ">" + ooxml.EscapeXMLText(text) + "</w:t>"
		}
		return "<w:t" + attrs + "></w:t>"
	})

	d.body = d.body[:start] + updated + d.body[end:]
	return nil
}

// SaveAs 将文档写入 path。
//
// 打开模式下，仅 word/document.xml 会被重写，其余部件按原始字节复制，
// 从而完整保留图片、样式、页眉页脚等内容。
// 新建模式下，行为与 Save 一致。
func (d *Document) SaveAs(path string) error {
	if d.src == nil {
		return d.Save(path)
	}

	d.src.SetPart(documentPart, []byte(d.body))

	if err := d.src.Save(path); err != nil {
		return fmt.Errorf("docgo: %w", err)
	}
	return nil
}

// SaveBytes 返回文档序列化后的字节流，保真规则与 SaveAs 相同。
func (d *Document) SaveBytes() ([]byte, error) {
	if d.src == nil {
		return d.Bytes()
	}

	d.src.SetPart(documentPart, []byte(d.body))

	data, err := d.src.WriteBytes()
	if err != nil {
		return nil, fmt.Errorf("docgo: %w", err)
	}
	return data, nil
}

// needsSpacePreserve 判断字符串首尾是否存在需要显式保留的空白。
func needsSpacePreserve(s string) bool {
	if s == "" {
		return false
	}
	return s != strings.TrimSpace(s)
}

// unescapeXML 还原 XML 中的基本实体引用。
func unescapeXML(s string) string {
	r := strings.NewReplacer(
		"&lt;", "<",
		"&gt;", ">",
		"&quot;", `"`,
		"&apos;", "'",
		"&amp;", "&", // 必须放在最后，避免二次解码
	)
	return r.Replace(s)
}

// extractTagText 返回 XML 中首个 <tag>...</tag> 的内容，并还原实体引用。
// 与 extractTag 的区别在于会做反转义，适用于读取可能含特殊字符的元数据。
func extractTagText(xml, tag string) string {
	return unescapeXML(extractTag(xml, tag))
}
