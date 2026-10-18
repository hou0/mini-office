export type CellScalar = string | number | boolean | null;

export type HorizontalAlign = 'general' | 'left' | 'center' | 'right';
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export interface BorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none';
  color?: string;
}

export interface CellBorders {
  top?: BorderSide;
  right?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
}

export interface CellStyle {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fill?: string;
  horizontalAlign?: HorizontalAlign;
  verticalAlign?: VerticalAlign;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  numFmt?: string;
  numFmtId?: number;
  borders?: CellBorders;
  /** Original cellXfs index, used to preserve untouched style references. */
  xfIndex?: number;
}

export interface Cell {
  value: CellScalar;
  formula?: string;
  styleId?: string;
}

export interface ComputedCell {
  value: CellScalar;
  display: string;
  error?: string;
}

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface FreezePane {
  rows: number;
  cols: number;
}

export interface SheetModel {
  id: string;
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  order: number;
  cells: Record<string, Cell>;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  merges: CellRange[];
  freeze?: FreezePane;
  hiddenRows: Record<number, true>;
  hiddenCols: Record<number, true>;
  autoFilter?: CellRange;
  dimension: { rows: number; cols: number };
  defaultColWidth: number;
  defaultRowHeight: number;
  showGridLines: boolean;
  tabColor?: string;
}

export interface WorkbookModel {
  version: 1;
  sheets: SheetModel[];
  activeSheetId: string;
  styles: Record<string, CellStyle>;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  defaultFont: string;
  defaultFontSizePt: number;
  date1904?: boolean;
  definedNames?: Record<string, string>;
  csv?: boolean;
  meta: { title?: string; author?: string; created?: string; modified?: string };
  /** Transient calculation cache; never serialized to XLSX. */
  computed?: Record<string, Record<string, ComputedCell>>;
  stats?: { cells: number; sheets: number };
}

export const DEFAULT_COL_WIDTH = 88;
export const DEFAULT_ROW_HEIGHT = 24;

export function createSheet(name: string, order: number): SheetModel {
  return {
    id: `sheet-${order + 1}`,
    name,
    state: 'visible',
    order,
    cells: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    hiddenRows: {},
    hiddenCols: {},
    dimension: { rows: 100, cols: 26 },
    defaultColWidth: DEFAULT_COL_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    showGridLines: true,
  };
}

export function createBlankWorkbook(): WorkbookModel {
  const sheet = createSheet('Sheet1', 0);
  return {
    version: 1,
    sheets: [sheet],
    activeSheetId: sheet.id,
    styles: {
      default: {
        fontFamily: '等线',
        fontSizePt: 11,
        numFmt: 'General',
        horizontalAlign: 'general',
        verticalAlign: 'bottom',
      },
    },
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: { title: '未命名工作簿' },
    stats: { cells: 0, sheets: 1 },
  };
}

export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): { row: number; col: number } {
  const [row, col] = key.split(':').map(Number);
  return { row, col };
}

export function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
}

export function normalizeRange(range: CellRange): CellRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}
