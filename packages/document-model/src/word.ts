export type WordAlignment = 'left' | 'center' | 'right' | 'justify';

export interface RunStyle {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  verticalAlign?: 'super' | 'sub' | 'baseline';
  letterSpacingPx?: number;
  allCaps?: boolean;
  smallCaps?: boolean;
}

export interface DocImage {
  id?: string;
  src: string;
  widthPx: number;
  heightPx: number;
  alt?: string;
  /** Existing OOXML relationship + media part, preserved on save. */
  relId?: string;
  part?: string;
  name?: string;
  /** Base64 data URL for newly inserted images. */
  dataUrl?: string;
}

export interface DocRun {
  text?: string;
  style: RunStyle;
  image?: DocImage;
  /** Explicit line break attached to this run. */
  lineBreak?: boolean;
  tab?: boolean;
  /** Field code such as PAGE / NUMPAGES, resolved by the renderer. */
  field?: string;
}

export interface ParagraphStyle {
  alignment?: WordAlignment;
  indentLeftPx?: number;
  indentRightPx?: number;
  firstLineIndentPx?: number;
  spaceBeforePx?: number;
  spaceAfterPx?: number;
  /** Multiplier, e.g. 1.5. */
  lineHeight?: number;
  /** Exact line height in CSS px (fixed line spacing). */
  lineHeightExactPx?: number;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
  outlineLevel?: number;
  backgroundColor?: string;
}

export interface ListInfo {
  numId?: string;
  level: number;
  marker: string;
  ordered: boolean;
}

export interface DocParagraph {
  type: 'paragraph';
  id: string;
  runs: DocRun[];
  style: ParagraphStyle;
  list?: ListInfo;
  styleId?: string;
  /** Source paragraph index for stable round-trips. */
  sourceIndex?: number;
}

export interface TableBorderSide {
  style: 'none' | 'single' | 'double' | 'dashed' | 'dotted';
  widthPx: number;
  color: string;
}

export interface TableBorders {
  top?: TableBorderSide;
  right?: TableBorderSide;
  bottom?: TableBorderSide;
  left?: TableBorderSide;
  insideH?: TableBorderSide;
  insideV?: TableBorderSide;
}

export interface DocTableCell {
  id: string;
  blocks: DocBlock[];
  colSpan: number;
  rowSpan: number;
  widthPx?: number;
  backgroundColor?: string;
  borders?: TableBorders;
  verticalAlign?: 'top' | 'middle' | 'bottom';
}

export interface DocTableRow {
  id: string;
  cells: DocTableCell[];
  heightPx?: number;
  isHeader?: boolean;
}

export interface DocTable {
  type: 'table';
  id: string;
  rows: DocTableRow[];
  columnWidthsPx: number[];
  widthPct?: number;
  alignment?: WordAlignment;
  borders?: TableBorders;
  styleId?: string;
}

export interface PageBreakBlock {
  type: 'pageBreak';
  id: string;
}

export type DocBlock = DocParagraph | DocTable | PageBreakBlock;

export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter?: number;
}

export interface PageSetup {
  widthPx: number;
  heightPx: number;
  orientation: 'portrait' | 'landscape';
  sizeName: string;
  margins: PageMargins;
  columns?: number;
}

export interface HeaderFooter {
  blocks: DocBlock[];
  distancePx: number;
  pageNumbers?: boolean;
}

export interface DocMeta {
  title?: string;
  author?: string;
  subject?: string;
  created?: string;
  modified?: string;
  lastModifiedBy?: string;
}

export interface DocModel {
  version: 1;
  page: PageSetup;
  blocks: DocBlock[];
  header?: HeaderFooter;
  footer?: HeaderFooter;
  defaultFont: string;
  defaultFontSizePt: number;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  /** Map relationship id → media info for images referenced by runs. */
  images: Record<string, { src: string; part?: string }>;
  meta: DocMeta;
  hasTrackedChanges?: boolean;
  hasComments?: boolean;
  hasContentControls?: boolean;
  stats?: { words: number; characters: number };
}

export const A4_PAGE: PageSetup = {
  widthPx: 793.7,
  heightPx: 1122.52,
  orientation: 'portrait',
  sizeName: 'A4',
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
};

export function createBlankDoc(): DocModel {
  return {
    version: 1,
    page: { ...A4_PAGE, margins: { ...A4_PAGE.margins } },
    blocks: [
      {
        type: 'paragraph',
        id: 'p-1',
        runs: [{ text: '', style: { fontSizePt: 10.5 } }],
        style: { alignment: 'left', lineHeight: 1.6 },
      },
    ],
    defaultFont: '宋体',
    defaultFontSizePt: 10.5,
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    images: {},
    meta: { title: '未命名文档' },
    stats: { words: 0, characters: 0 },
  };
}
