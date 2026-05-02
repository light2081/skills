/**
 * 共享样式模块 — 基于 word文档模板.docm
 * 所有文档生成脚本通过 require 此文件获取统一样式配置
 */

const {
  AlignmentType, BorderStyle, LevelFormat, ShadingType,
  PageOrientation, WidthType,
} = require("docx");

// ─── 字体 ────────────────────────────────────────────────────────────────────
const FONT = { ascii: "Times New Roman", hAnsi: "Times New Roman", eastAsia: "宋体" };

// ─── 颜色 ────────────────────────────────────────────────────────────────────
const COLOR = {
  black:       "000000",
  tableHeader: "A6A6A6",  // 表头灰色填充
  codeBg:      "E5E5E5",  // 代码块背景
  border:      "auto",
};

// ─── 字号（单位：半磅，12pt = 24） ──────────────────────────────────────────
const SIZE = {
  body:     24,  // 12pt 正文
  h1:       36,  // 18pt
  h2:       32,  // 16pt
  h3:       30,  // 15pt
  h4:       28,  // 14pt
  h5:       28,  // 14pt
  h6:       28,  // 14pt
  h7:       28,  // 14pt
  h8:       28,  // 14pt
  caption:  24,  // 12pt 图注/表注
  header:   22,  // 11pt 页眉
  footer:   18,  //  9pt 页脚
  toc:      24,  // 12pt 目录
  title:    72,  // 36pt 封面主标题
  subtitle: 72,  // 36pt 封面副标题
  tocTitle: 32,  // 16pt 目录标题
};

// ─── 行距 ────────────────────────────────────────────────────────────────────
const SPACING = {
  line15: { line: 360, lineRule: "auto" },  // 1.5倍
  line1:  { line: 240, lineRule: "auto" },  // 单倍
};

// ─── 页面设置 ─────────────────────────────────────────────────────────────────
const PAGE = {
  size:   { width: 11906, height: 16838 },  // A4
  margin: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 1020, footer: 1020 },
};

// ─── 边框 ────────────────────────────────────────────────────────────────────
const BORDER_SINGLE = { style: BorderStyle.SINGLE, size: 4, color: COLOR.border };
const TABLE_BORDERS = {
  top: BORDER_SINGLE, bottom: BORDER_SINGLE,
  left: BORDER_SINGLE, right: BORDER_SINGLE,
  insideH: BORDER_SINGLE, insideV: BORDER_SINGLE,
};

// ─── 编号配置 ─────────────────────────────────────────────────────────────────
const NUMBERING_CONFIG = [
  // 正文序号
  { reference: "list-1",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 620, hanging: 420 } } } }] },
  { reference: "list-2",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "（%1）",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 1137, hanging: 737 } } } }] },
  { reference: "list-3",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 1020, hanging: 420 } } } }] },
  // 正文项目符号
  { reference: "bullet-1",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 620, hanging: 420 } } } }] },
  { reference: "bullet-2",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 820, hanging: 420 } } } }] },
  { reference: "bullet-3",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 1020, hanging: 420 } } } }] },
  // 表格内序号
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
  // 表格内项目符号
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
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 0 } } } }] },
  { reference: "table-caption",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "Table %1",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 0 } } } }] },
];

// ─── 段落样式列表 ─────────────────────────────────────────────────────────────
const PARAGRAPH_STYLES = [
  { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h1, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { firstLine: 0, firstLineChars: 0 }, outlineLevel: 0,
      pageBreakBefore: true, keepNext: true, keepLines: true } },
  { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h2, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { firstLine: 0, firstLineChars: 0 }, outlineLevel: 1 } },
  { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h3, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { firstLine: 0, firstLineChars: 0 }, outlineLevel: 2 } },
  { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h4, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { firstLine: 0, firstLineChars: 0 }, outlineLevel: 3 } },
  { id: "Heading5", name: "Heading 5", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h5, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { left: 851, hanging: 851, firstLine: 0, firstLineChars: 0 }, outlineLevel: 4 } },
  { id: "Heading6", name: "Heading 6", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h6, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { left: 1134, firstLine: 0, firstLineChars: 0 }, outlineLevel: 5 } },
  { id: "Heading7", name: "Heading 7", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h7, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { left: 1276, firstLine: 0, firstLineChars: 0 }, outlineLevel: 6 } },
  { id: "Heading8", name: "Heading 8", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.h8, bold: true, font: FONT },
    paragraph: { spacing: { before: 100, after: 75 },
      indent: { left: 1418, firstLine: 0, firstLineChars: 0 }, outlineLevel: 7 } },
  { id: "FigureCaption", name: "Figure Caption", basedOn: "Normal", next: "Normal",
    run: { size: SIZE.caption, bold: true, font: FONT, color: COLOR.black },
    paragraph: { alignment: AlignmentType.CENTER,
      indent: { firstLine: 0, firstLineChars: 0 }, spacing: { after: 50 } } },
  { id: "TableCaption", name: "Table Caption", basedOn: "Normal", next: "Normal",
    run: { size: SIZE.caption, bold: true, font: FONT, color: COLOR.black },
    paragraph: { alignment: AlignmentType.CENTER,
      indent: { firstLine: 0, firstLineChars: 0 },
      spacing: { before: 163, ...SPACING.line15 }, keepNext: true } },
  { id: "FigureParagraph", name: "Figure Paragraph", basedOn: "Normal", next: "Normal",
    paragraph: { alignment: AlignmentType.CENTER,
      indent: { firstLine: 0, firstLineChars: 0 } } },
  { id: "TOCHeading", name: "TOC Heading", basedOn: "Normal", next: "Normal",
    run: { size: SIZE.tocTitle, bold: true, font: FONT },
    paragraph: { alignment: AlignmentType.CENTER,
      indent: { firstLine: 0, firstLineChars: 0 } } },
  { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.title, bold: true, font: FONT },
    paragraph: { alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 60 },
      indent: { firstLine: 0, firstLineChars: 0 } } },
  { id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: SIZE.subtitle, font: FONT },
    paragraph: { alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 60, ...SPACING.line1 },
      indent: { firstLine: 0, firstLineChars: 0 } } },
];

// ─── 文档默认样式 ─────────────────────────────────────────────────────────────
const DOC_DEFAULT_STYLES = {
  default: {
    document: {
      run: { font: FONT, size: SIZE.body, color: COLOR.black },
      paragraph: {
        spacing: SPACING.line15,
        indent: { firstLineChars: 200 },
        alignment: AlignmentType.BOTH,
      },
    },
  },
  paragraphStyles: PARAGRAPH_STYLES,
};

module.exports = {
  FONT, COLOR, SIZE, SPACING, PAGE,
  BORDER_SINGLE, TABLE_BORDERS,
  NUMBERING_CONFIG, PARAGRAPH_STYLES, DOC_DEFAULT_STYLES,
};
