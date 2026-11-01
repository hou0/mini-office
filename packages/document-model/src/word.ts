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

/**
 * 功能：记录 wp:anchor 的锚定与环绕参数，让浮动图片按 Word 语义定位而不是简单 float；
 * 时间：2026-10-31；作者：lq。
 *
 * Word 的浮动图片是「锚点段落 + 相对基准偏移 + 环绕距离」三部分决定位置的，
 * 只保存 wrap 方向会让图片挤占后续段落的行盒，造成正文莫名换行。
 */
export interface ImageAnchor {
  /** 水平基准：column（栏）/ margin（页边距）/ page（页面）/ character（字符）。 */
  hRelative?: string;
  /** 水平偏移，px；负值表示向左越界。 */
  hOffsetPx?: number;
  /** 水平对齐（与偏移二选一）。 */
  hAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  /** 垂直基准：paragraph（段落）/ line（行）/ margin / page。 */
  vRelative?: string;
  /** 垂直偏移，px。 */
  vOffsetPx?: number;
  vAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
  /** 环绕类型，none 表示浮于文字上/下方不占位。 */
  wrapType?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** 环绕侧：both / left / right / largest。 */
  wrapText?: string;
  /** 衬于文字下方。 */
  behindDoc?: boolean;
  /** 四周环绕距离（EMU 转 px）。 */
  distTopPx?: number;
  distBottomPx?: number;
  distLeftPx?: number;
  distRightPx?: number;
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
  /** 功能：浮动图片的锚定参数，保存时原样写回；时间：2026-10-31；作者：lq。 */
  anchor?: ImageAnchor;
}

export interface MergeFieldFormat {
  num?: string;        // \# 格式，如 "#,##0.00"
  date?: string;       // \@ 格式，如 "yyyy年M月d日"
  case?: 'upper' | 'lower' | 'caps' | 'firstcap'; // \* 开关
  before?: string;     // \b 非空前缀
  after?: string;      // \f 非空后缀
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
  mergeFormat?: MergeFieldFormat;
  /** Word 邮件合并控制域（解析后占位文本，合并时消耗） */
  mergeNext?: boolean;         // NEXT
  mergeNextIf?: string;        // NEXTIF 条件表达式
  mergeSkipIf?: string;        // SKIPIF 条件表达式
  mergeRecord?: boolean;       // MERGEREC
  mergeSeq?: boolean;          // MERGESEQ
}

/**
 * 功能：Word 制表位（w:tabs/w:tab），标尺与正文 Tab 定位共用同一份数据；时间：2026-10-30；作者：lq。
 *
 * `posPx` 相对段落左边界（即正文内容框左沿，不含缩进），与 OOXML 的 `w:pos` 语义一致，
 * 因此解析后可直接与浏览器实测位置比较，不需要再做一次坐标换算。
 */
export type TabAlignment = 'left' | 'center' | 'right' | 'decimal';

export type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'heavy';

export interface TabStop {
  posPx: number;
  alignment: TabAlignment;
  leader?: TabLeader;
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
  /** 功能：段落自定义制表位，为空时按文档默认制表位（defaultTabStopPx）等距推进；时间：2026-10-30；作者：lq。 */
  tabStops?: TabStop[];
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
  /**
   * 功能：标记该段落是目录域（TOC），渲染时按文档标题实时生成条目；时间：2026-10-30；作者：lq。
   *
   * 目录用「一个段落承载全部条目」而不是新增块类型，原因是分页、测量、撤销、查找替换等
   * 现有流程都以段落为最小单位，这样改造面最小且不会漏掉某个分支。
   */
  toc?: { levels: number };
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

/** 功能：邮件合并筛选条件，面板筛选与 Word 兼容输出共用同一结构；时间：2026-10-29；作者：lq。 */
export interface DocMergeCondition {
  field: string;
  op: 'contains' | 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte' | 'empty' | 'notEmpty' | 'notContains';
  value: string;
  /** 与上一条件的组合方式，仅第 2 条起生效；对齐 Word 查询选项的「且 / 或」。 */
  conjunction?: 'and' | 'or';
}

export interface DocMeta {
  title?: string;
  author?: string;
  subject?: string;
  created?: string;
  modified?: string;
  lastModifiedBy?: string;
  /** 邮件合并面板持久化的筛选/排序/输出状态（JSON 字符串），保存在 settings.xml 的文档变量中。 */
  mailMergeState?: string;
  /** 写给 Word / WPS 的邮件合并数据源与筛选，保存模板时转成标准 w:mailMerge，不参与 mini-office 自身恢复。 */
  mailMergeTemplate?: { dataSourcePath: string; sheetName: string; fields: string[]; conditions: DocMergeCondition[] };
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
  /** 功能：文档默认制表位间距（settings.xml 的 w:defaultTabStop），未设置时按 0.5 英寸处理；时间：2026-10-30；作者：lq。 */
  defaultTabStopPx?: number;
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
