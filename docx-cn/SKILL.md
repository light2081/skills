---
name: docx-cn
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of 'Word doc', 'word document', '.docx', or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a 'report', 'memo', 'letter', 'template', or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
license: Proprietary. LICENSE.txt has complete terms
---
# DOCX creation, editing, and analysis

## Overview

A .docx file is a ZIP archive containing XML files.

## Quick Reference

| Task                   | Approach                                                            |
| ---------------------- | ------------------------------------------------------------------- |
| Read/analyze content   | `pandoc` or unpack for raw XML                                    |
| Create new document    | Use `docx-js` - see Creating New Documents below                  |
| Edit existing document | Unpack → edit XML → repack - see Editing Existing Documents below |

### Converting .doc to .docx

Legacy `.doc` files must be converted before editing:

```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content

```bash
# Text extraction with tracked changes
pandoc --track-changes=all document.docx -o output.md

# Raw XML access
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes

To produce a clean document with all tracked changes accepted (requires LibreOffice):

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## Creating New Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Setup

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        InternalHyperlink, Bookmark, FootnoteReferenceRun, PositionalTab,
        PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
        TabStopType, TabStopPosition, Column, SectionType,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation

After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.

```bash
python scripts/office/validate.py doc.docx
```

### Page Size

```javascript
// 始终使用 A4，上下页边距1英寸，左右1.25英寸
sections: [{
  properties: {
    page: {
      size: { width: 11906, height: 16838 }, // A4
      margin: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 1020, footer: 1020 }
    }
  },
  children: [/* content */]
}]
```

**常用页面尺寸（DXA，1440 DXA = 1英寸）：**

| 纸张 | 宽 | 高 | 正文宽（左右1.25英寸边距） |
|------|----|----|--------------------------|
| A4（默认） | 11,906 | 16,838 | 8,306 |
| US Letter | 12,240 | 15,840 | 8,640 |

**横向：** docx-js 内部会交换宽高，传竖向尺寸 + `orientation: PageOrientation.LANDSCAPE` 即可：
```javascript
size: {
  width: 11906,  // 传短边
  height: 16838, // 传长边
  orientation: PageOrientation.LANDSCAPE
},
// 正文宽 = 16838 - 左边距 - 右边距
```

### Styles (Template — MANDATORY for ALL documents)

**所有新建文档必须使用以下样式，源自标准模板 `word文档模板.docm`。**

#### 全局字体与段落默认值

| 属性 | 值 |
|------|-----|
| 中文字体 | 宋体 |
| 西文字体 | Times New Roman |
| 字号 | 12pt（`size: 24`） |
| 颜色 | 黑色（`000000`） |
| 行距 | 1.5倍（`line: 360, lineRule: "auto"`） |
| 首行缩进 | 2字符（`firstLineChars: 200`） |
| 对齐 | 两端对齐（`AlignmentType.BOTH`） |

#### 标题样式（heading 1~8）

| 样式 | 字号 | 段前/段后 | 首行缩进 | 大纲级别 | 特殊 |
|------|------|----------|---------|---------|------|
| Heading 1 | 18pt (36) | 100/75 twip | 0 | 0 | 段前分页，keepNext，keepLines |
| Heading 2 | 16pt (32) | 100/75 twip | 0 | 1 | — |
| Heading 3 | 15pt (30) | 100/75 twip | 0 | 2 | — |
| Heading 4 | 14pt (28) | 100/75 twip | 0 | 3 | — |
| Heading 5 | 14pt (28) | 100/75 twip | 0 | 4 | 左缩进851，悬挂851 |
| Heading 6 | 14pt (28) | 100/75 twip | 0 | 5 | 左缩进1134 |
| Heading 7 | 14pt (28) | 100/75 twip | 0 | 6 | 左缩进1276 |
| Heading 8 | 14pt (28) | 100/75 twip | 0 | 7 | 左缩进1418 |

所有标题均加粗，字体宋体/Times New Roman，无首行缩进（Heading 5~8 有左缩进但无首行缩进）。

#### 图注、表注

| 样式 ID | 对齐 | 字号 | 加粗 | 颜色 | 编号格式 | 备注 |
|---------|------|------|------|------|---------|------|
| FigureCaption | 居中 | 12pt | 是 | #000000 | `Fig 1`、`Fig 2`... | 段后50 twip |
| TableCaption | 居中 | 12pt | 是 | #000000 | `Table 1`、`Table 2`... | 段前163 twip，keepNext |

#### 列表（NEVER use unicode bullets）

```javascript
// WRONG: never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })

// CORRECT: use numbering config
numbering: {
  config: [
    // 正文列表
    { reference: "list-1",   // 1级序号: "1."  left=620, hanging=420
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 620, hanging: 420 } } } }] },
    { reference: "list-2",   // 2级序号: "（1）"  left=1137, hanging=737
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "（%1）",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1137, hanging: 737 } } } }] },
    { reference: "list-3",   // 3级序号: "1)"  left=1020, hanging=420
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1020, hanging: 420 } } } }] },
    { reference: "bullet-1", // 1级项目符号  left=620
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 620, hanging: 420 } } } }] },
    { reference: "bullet-2", // 2级项目符号  left=820
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 820, hanging: 420 } } } }] },
    { reference: "bullet-3", // 3级项目符号  left=1020
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1020, hanging: 420 } } } }] },
    // 表格内列表（缩进更小）
    { reference: "table-list-1",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 420, hanging: 420 } } } }] },
    { reference: "table-list-2",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "（%1）",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 840, hanging: 420 } } } }] },
    { reference: "table-list-3",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1260, hanging: 420 } } } }] },
    { reference: "table-bullet-1",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 420, hanging: 420 } } } }] },
    { reference: "table-bullet-2",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 840, hanging: 420 } } } }] },
    { reference: "table-bullet-3",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1260, hanging: 420 } } } }] },
    // 图注/表注自动编号
    { reference: "fig-caption",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "Fig %1",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 0 } } } }] },
    { reference: "table-caption",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "Table %1",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 0 } } } }] },
  ]
}
// Same reference = 连续编号；Different reference = 从1重新开始
```

#### 共享样式模块

所有文档生成脚本必须通过 `require` 引入共享样式，禁止在脚本中硬编码字体、颜色、字号等常量。

```javascript
const {
  FONT, COLOR, SIZE, SPACING, PAGE,
  BORDER_SINGLE, TABLE_BORDERS,
  NUMBERING_CONFIG, DOC_DEFAULT_STYLES,
} = require('D:/workspace/AIcoding/daily/skill/docx-cn/scripts/styles.js');
```

> **FONT 注意**：`styles.js` 使用 `{ ascii, hAnsi, eastAsia }` 而非 `{ name, eastAsia }`。
> docx-js 9.x 中 `name` 会将所有字体槽（含 eastAsia）统一覆盖，导致中文字体失效。

#### 完整 Document 样式配置

```javascript
const { AlignmentType, BorderStyle, Document, Header, Footer,
        Paragraph, TextRun, PageNumber } = require("docx");
const {
  FONT, PAGE, NUMBERING_CONFIG, DOC_DEFAULT_STYLES,
} = require('D:/workspace/AIcoding/daily/skill/docx-cn/scripts/styles.js');

const doc = new Document({
  numbering: { config: NUMBERING_CONFIG },
  styles: DOC_DEFAULT_STYLES,
  sections: [{
    properties: {
      page: { size: PAGE.size, margin: PAGE.margin }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1 } },
          alignment: AlignmentType.CENTER,
          spacing: { line: 240, lineRule: "auto" },
          indent: { firstLine: 0, firstLineChars: 0 },
          children: [new TextRun({ text: "文档标题", bold: true, size: 22, font: FONT })]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          indent: { firstLine: 360 },
          spacing: { after: 120 },
          children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, font: FONT })]
        })
      ]})
    },
    children: [/* content */]
  }]
});
```

#### 表格样式

```javascript
const { TABLE_BORDERS, COLOR } =
  require('D:/workspace/AIcoding/daily/skill/docx-cn/scripts/styles.js');
const { ShadingType } = require("docx");

// 标准表格：全边框单线 sz=4，无填充色，12pt 宋体/Times New Roman
// tableBorders 直接使用 TABLE_BORDERS（已包含六面边框）
// 单元格 shading：
//   普通单元格 → 不设置 shading（无填充）
//   表头行 → shading: { fill: COLOR.tableHeader, type: ShadingType.CLEAR }
//   代码背景 → shading: { fill: COLOR.codeBg,      type: ShadingType.CLEAR }
// 表格内文字段落：无首行缩进，行距继承文档默认（1.5倍）
```

### Tables

**CRITICAL: Tables need dual widths** - set both `columnWidths` on the table AND `width` on each cell. Without both, tables render incorrectly on some platforms.

```javascript
// CRITICAL: Always set table width for consistent rendering
// CRITICAL: Use ShadingType.CLEAR (not SOLID) to prevent black backgrounds
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

new Table({
  width: { size: 9360, type: WidthType.DXA }, // Always use DXA (percentages break in Google Docs)
  columnWidths: [4680, 4680], // Must sum to table width (DXA: 1440 = 1 inch)
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: 4680, type: WidthType.DXA }, // Also set on each cell
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, // CLEAR not SOLID
          margins: { top: 80, bottom: 80, left: 120, right: 120 }, // Cell padding (internal, not added to width)
          children: [new Paragraph({ children: [new TextRun("Cell")] })]
        })
      ]
    })
  ]
})
```

**Table width calculation:**

Always use `WidthType.DXA` — `WidthType.PERCENTAGE` breaks in Google Docs.

```javascript
// Table width = sum of columnWidths = content width
// US Letter with 1" margins: 12240 - 2880 = 9360 DXA
width: { size: 9360, type: WidthType.DXA },
columnWidths: [7000, 2360]  // Must sum to table width
```

**Width rules:**

- **Always use `WidthType.DXA`** — never `WidthType.PERCENTAGE` (incompatible with Google Docs)
- Table width must equal the sum of `columnWidths`
- Cell `width` must match corresponding `columnWidth`
- Cell `margins` are internal padding - they reduce content area, not add to cell width
- For full-width tables: use content width (page width minus left and right margins)

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  children: [new ImageRun({
    type: "png", // Required: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" } // All three required
  })]
})
```

### Page Breaks

```javascript
// CRITICAL: PageBreak must be inside a Paragraph
new Paragraph({ children: [new PageBreak()] })

// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Hyperlinks

```javascript
// External link
new Paragraph({
  children: [new ExternalHyperlink({
    children: [new TextRun({ text: "Click here", style: "Hyperlink" })],
    link: "https://example.com",
  })]
})

// Internal link (bookmark + reference)
// 1. Create bookmark at destination
new Paragraph({ heading: HeadingLevel.HEADING_1, children: [
  new Bookmark({ id: "chapter1", children: [new TextRun("Chapter 1")] }),
]})
// 2. Link to it
new Paragraph({ children: [new InternalHyperlink({
  children: [new TextRun({ text: "See Chapter 1", style: "Hyperlink" })],
  anchor: "chapter1",
})]})
```

### Footnotes

```javascript
const doc = new Document({
  footnotes: {
    1: { children: [new Paragraph("Source: Annual Report 2024")] },
    2: { children: [new Paragraph("See appendix for methodology")] },
  },
  sections: [{
    children: [new Paragraph({
      children: [
        new TextRun("Revenue grew 15%"),
        new FootnoteReferenceRun(1),
        new TextRun(" using adjusted metrics"),
        new FootnoteReferenceRun(2),
      ],
    })]
  }]
});
```

### Tab Stops

```javascript
// Right-align text on same line (e.g., date opposite a title)
new Paragraph({
  children: [
    new TextRun("Company Name"),
    new TextRun("\tJanuary 2025"),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
})

// Dot leader (e.g., TOC-style)
new Paragraph({
  children: [
    new TextRun("Introduction"),
    new TextRun({ children: [
      new PositionalTab({
        alignment: PositionalTabAlignment.RIGHT,
        relativeTo: PositionalTabRelativeTo.MARGIN,
        leader: PositionalTabLeader.DOT,
      }),
      "3",
    ]}),
  ],
})
```

### Multi-Column Layouts

```javascript
// Equal-width columns
sections: [{
  properties: {
    column: {
      count: 2,          // number of columns
      space: 720,        // gap between columns in DXA (720 = 0.5 inch)
      equalWidth: true,
      separate: true,    // vertical line between columns
    },
  },
  children: [/* content flows naturally across columns */]
}]

// Custom-width columns (equalWidth must be false)
sections: [{
  properties: {
    column: {
      equalWidth: false,
      children: [
        new Column({ width: 5400, space: 720 }),
        new Column({ width: 3240 }),
      ],
    },
  },
  children: [/* content */]
}]
```

Force a column break with a new section using `type: SectionType.NEXT_COLUMN`.

### Table of Contents

```javascript
// CRITICAL: Headings must use HeadingLevel ONLY - no custom styles
new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })
```

### Headers/Footers

```javascript
sections: [{
  properties: {
    page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } // 1440 = 1 inch
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Critical Rules for docx-js

- **Set page size explicitly** - docx-js defaults to A4; use US Letter (12240 x 15840 DXA) for US documents
- **Landscape: pass portrait dimensions** - docx-js swaps width/height internally; pass short edge as `width`, long edge as `height`, and set `orientation: PageOrientation.LANDSCAPE`
- **Never use `\n`** - use separate Paragraph elements
- **Never use unicode bullets** - use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** - standalone creates invalid XML
- **ImageRun requires `type`** - always specify png/jpg/etc
- **Always set table `width` with DXA** - never use `WidthType.PERCENTAGE` (breaks in Google Docs)
- **Tables need dual widths** - `columnWidths` array AND cell `width`, both must match
- **Table width = sum of columnWidths** - for DXA, ensure they add up exactly
- **Always add cell margins** - use `margins: { top: 80, bottom: 80, left: 120, right: 120 }` for readable padding
- **Use `ShadingType.CLEAR`** - never SOLID for table shading
- **Never use tables as dividers/rules** - cells have minimum height and render as empty boxes (including in headers/footers); use `border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } }` on a Paragraph instead. For two-column footers, use tab stops (see Tab Stops section), not tables
- **TOC requires HeadingLevel only** - no custom styles on heading paragraphs
- **Override built-in styles** - use exact IDs: "Heading1", "Heading2", etc.
- **Include `outlineLevel`** - required for TOC (0 for H1, 1 for H2, etc.)

---

## Editing Existing Documents

**Follow all 3 steps in order.**

### Step 1: Unpack

```bash
python scripts/office/unpack.py document.docx unpacked/
```

Extracts XML, pretty-prints, merges adjacent runs, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing. Use `--merge-runs false` to skip run merging.

### Step 2: Edit XML

Edit files in `unpacked/word/`. See XML Reference below for patterns.

**Use "Claude" as the author** for tracked changes and comments, unless the user explicitly requests use of a different name.

**Use the Edit tool directly for string replacement. Do not write Python scripts.** Scripts introduce unnecessary complexity. The Edit tool shows exactly what is being replaced.

**CRITICAL: Use smart quotes for new content.** When adding text with apostrophes or quotes, use XML entities to produce smart quotes:

```xml
<!-- Use these entities for professional typography -->
<w:t>Here’s a quote: “Hello”</w:t>
```

| Entity       | Character                      |
| ------------ | ------------------------------ |
| `&#x2018;` | ‘ (left single)               |
| `&#x2019;` | ’ (right single / apostrophe) |
| `&#x201C;` | “ (left double)               |
| `&#x201D;` | ” (right double)              |

**Adding comments:** Use `comment.py` to handle boilerplate across multiple XML files (text must be pre-escaped XML):

```bash
python scripts/comment.py unpacked/ 0 "Comment text with & and ’"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0  # reply to comment 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"  # custom author name
```

Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack

```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```

Validates with auto-repair, condenses XML, and creates DOCX. Use `--validate false` to skip.

**Auto-repair will fix:**

- `durableId` >= 0x7FFFFFFF (regenerates valid ID)
- Missing `xml:space="preserve"` on `<w:t>` with whitespace

**Auto-repair won't fix:**

- Malformed XML, invalid element nesting, missing relationships, schema violations

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings. Don't inject tracked change tags inside a run.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.

---

## XML Reference

### Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: Add `xml:space="preserve"` to `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)

### Tracked Changes

**Insertion:**

```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**

```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Inside `<w:del>`**: Use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

**Minimal edits** - only mark what changes:

```xml
<!-- Change "30 days" to "60 days" -->
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting entire paragraphs/list items** - when removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Add `<w:del/>` inside `<w:pPr><w:rPr>`:

```xml
<w:p>
  <w:pPr>
    <w:numPr>...</w:numPr>  <!-- list numbering if present -->
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```

Without the `<w:del/>` in `<w:pPr><w:rPr>`, accepting changes leaves an empty paragraph/list item.

**Rejecting another author's insertion** - nest deletion inside their insertion:

```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion** - add insertion after (don't modify their deletion):

```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py` (see Step 2), add markers to document.xml. For replies, use `--parent` flag and nest markers inside the parent's.

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment markers are direct children of w:p, never inside w:r -->
<w:commentRangeStart w:id="0"/>
<w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted</w:delText></w:r>
</w:del>
<w:r><w:t> more text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment 0 with reply 1 nested inside -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

### Images

1. Add image file to `word/media/`
2. Add relationship to `word/_rels/document.xml.rels`:

```xml
<Relationship Id="rId5" Type=".../image" Target="media/image1.png"/>
```

3. Add content type to `[Content_Types].xml`:

```xml
<Default Extension="png" ContentType="image/png"/>
```

4. Reference in document.xml:

```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="914400" cy="914400"/>  <!-- EMUs: 914400 = 1 inch -->
    <a:graphic>
      <a:graphicData uri=".../picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

---

## Dependencies

- **pandoc**: Text extraction
- **docx**: `npm install -g docx` (new documents)
- **LibreOffice**: PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images
