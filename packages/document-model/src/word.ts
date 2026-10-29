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
  /** 功能：保存图片的文字环绕方式，支持嵌入、左右环绕和上下型；时间：2026-10-24；作者：lq。 */
  wrap?: 'inline' | 'square-left' | 'square-right' | 'top-bottom';
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
  /** 功能：结构化邮件合并字段名称，与普通文字独立保存；时间：2026-10-24；作者：lq。 */
  mergeField?: string;
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
  /** 功能：最小值行距（w:lineRule="atLeast"），低于该高度时按内容撑开；时间：2026-10-29；作者：lq。 */
  lineHeightAtLeastPx?: number;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
  outlineLevel?: number;
  backgroundColor?: string;
  /** 功能：保留段落上下左右边框，用于兼容 Word 横线与签名栏；时间：2026-10-27；作者：lq。 */
  borders?: TableBorders;
}

export interface ListInfo {
  numId?: string;
  level: number;
  marker: string;
  ordered: boolean;
}

/**
 * 功能：文档级样式定义，供「样式库」批量套用并在保存时写入 styles.xml；时间：2026-10-29；作者：lq。
 *
 * 设计要点：
 * 1. `id` 直接作为 OOXML 的 `w:styleId`，内置预设沿用 Word 的内置 ID（Heading1/Title…），
 *    这样打开 Word 文档时其标题段落能被识别，反向保存后 Word 也仍视其为内置样式。
 * 2. 只保存「样式自身」的格式，不保存继承结果；应用时再与文档默认字体合并，
 *    避免文档换默认字体后样式仍停留在旧字体上。
 * 3. `origin` 仅用于 UI 分组与删除权限，不写入文件。
 */
export interface DocStyleDefinition {
  id: string;
  name: string;
  /** OOXML 内置样式名（如 heading 1），写入 styles.xml 后 Word 会按内置样式处理。 */
  builtinName?: string;
  paragraph: ParagraphStyle;
  run: RunStyle;
  basedOn?: string;
  /** 列表型样式：应用时同时写入段落列表标记。 */
  list?: { ordered: boolean; level: number };
  origin?: 'preset' | 'custom' | 'imported';
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
  /** 功能：保留段尾分节信息，让批量报告重新打开后仍独立编号；时间：2026-10-24；作者：lq。 */
  sectionEnd?: DocSection;
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
  /** 功能：页码起始值，独立于物理页序；时间：2026-10-24；作者：lq。 */
  pageNumberStart?: number;
}

export interface HeaderFooter {
  blocks: DocBlock[];
  distancePx: number;
  pageNumbers?: boolean;
  /** 功能：仅重写经用户编辑的页眉页脚，保留导入文档的其他变体；时间：2026-10-24；作者：lq。 */
  edited?: boolean;
  /** 功能：保留页眉页脚部件路径以正确复用图片关系；时间：2026-10-24；作者：lq。 */
  sourcePart?: string;
}

/** 功能：描述一个报告节的页面与页眉页脚；时间：2026-10-24；作者：lq。 */
export interface DocSection { page: PageSetup; header?: HeaderFooter; footer?: HeaderFooter }

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
  /** 功能：页面渲染时的节总页数，不写入文档包；时间：2026-10-24；作者：lq。 */
  sectionPageCount?: number;
  /** 功能：本文档自定义/导入的样式，与内置预设合并后供样式库使用；时间：2026-10-29；作者：lq。 */
  styles?: DocStyleDefinition[];
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
