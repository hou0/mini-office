import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  createBlankWorkbook,
  defaultSheetPageSetup,
  type BorderSide,
  type Cell,
  type CellBorders,
  type CellScalar,
  type CellStyle,
  type SheetModel,
  type WorkbookModel,
} from '@mini-office/document-model';
import {
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  mixColors,
  normalizeHex,
  parseXmlDom,
  readThemeColors,
  relationshipMap,
  type OoxmlPackage,
  type Relationship,
} from '@mini-office/ooxml';
import { baseName, stripExtension } from '@mini-office/shared';
import { BUILTIN_NUMBER_FORMATS } from './format';
import { recalculateWorkbook } from './formula/engine';
import { translateFormula } from './formula/parser';

const EXCEL_INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const THEME_INDEX_TO_SCHEME = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

function excelColor(element: Element | null, theme: Record<string, string>): string | undefined {
  if (!element) return undefined;
  const rgb = getAttr(element, 'rgb');
  if (rgb) {
    const normalized = normalizeHex(rgb);
    if (normalized) return normalized;
  }
  const themeIndex = getAttr(element, 'theme');
  if (themeIndex !== undefined) {
    const scheme = THEME_INDEX_TO_SCHEME[Number(themeIndex)] ?? 'dk1';
    const base = theme[scheme];
    if (base) {
      const tint = Number(getAttr(element, 'tint') ?? '0');
      if (Number.isFinite(tint) && tint !== 0) {
        return tint > 0 ? mixColors(base, '#FFFFFF', Math.min(1, tint)) : mixColors(base, '#000000', Math.min(1, -tint));
      }
      return base;
    }
  }
  const indexed = getAttr(element, 'indexed');
  if (indexed !== undefined) {
    const hex = EXCEL_INDEXED_COLORS[Number(indexed)];
    if (hex) return `#${hex}`;
  }
  return undefined;
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = parseXmlDom(xml);
  const root = firstDescendant(doc, 'sst') ?? doc.documentElement;
  const result: string[] = [];
  for (const item of Array.from(root.childNodes)) {
    if (!(item instanceof Element) || item.tagName !== 'si') continue;
    let text = '';
    for (const child of Array.from(item.childNodes)) {
      if (!(child instanceof Element)) continue;
      if (child.tagName === 't') text += child.textContent ?? '';
      else if (child.tagName === 'r') {
        for (const runChild of Array.from(child.childNodes)) {
          if (runChild instanceof Element && runChild.tagName === 't') text += runChild.textContent ?? '';
        }
      }
    }
    result.push(text);
  }
  return result;
}

interface ParsedStyles {
  styles: Record<string, CellStyle>;
  defaultStyleId: string;
}

function parseStyles(xml: string, theme: Record<string, string>): ParsedStyles {
  const styles: Record<string, CellStyle> = {};
  if (!xml) {
    styles.default = {
      fontFamily: '等线',
      fontSizePt: 11,
      numFmt: 'General',
      horizontalAlign: 'general',
      verticalAlign: 'bottom',
    };
    return { styles, defaultStyleId: 'default' };
  }
  const doc = parseXmlDom(xml);
  const root = firstDescendant(doc, 'styleSheet') ?? doc.documentElement;

  const numFmts = new Map<number, string>(Object.entries(BUILTIN_NUMBER_FORMATS).map(([id, fmt]) => [Number(id), fmt]));
  for (const item of Array.from(firstChild(root, 'numFmts')?.childNodes ?? [])) {
    if (!(item instanceof Element) || item.tagName !== 'numFmt') continue;
    const id = Number(getAttr(item, 'numFmtId'));
    const code = getAttr(item, 'formatCode');
    if (Number.isFinite(id) && code) numFmts.set(id, code);
  }

  interface FontInfo {
    fontFamily?: string;
    fontSizePt?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    color?: string;
  }
  const fonts: FontInfo[] = [];
  for (const fontElement of Array.from(firstChild(root, 'fonts')?.childNodes ?? [])) {
    if (!(fontElement instanceof Element) || fontElement.tagName !== 'font') continue;
    const font: FontInfo = {};
    const name = getAttr(firstChild(fontElement, 'name'), 'val');
    if (name) font.fontFamily = name;
    const size = Number(getAttr(firstChild(fontElement, 'sz'), 'val'));
    if (Number.isFinite(size)) font.fontSizePt = size;
    if (firstChild(fontElement, 'b')) font.bold = true;
    if (firstChild(fontElement, 'i')) font.italic = true;
    if (firstChild(fontElement, 'u')) font.underline = true;
    if (firstChild(fontElement, 'strike')) font.strike = true;
    const color = excelColor(firstChild(fontElement, 'color'), theme);
    if (color) font.color = color;
    fonts.push(font);
  }

  const fills: Array<{ color?: string }> = [];
  for (const fillElement of Array.from(firstChild(root, 'fills')?.childNodes ?? [])) {
    if (!(fillElement instanceof Element) || fillElement.tagName !== 'fill') continue;
    const pattern = firstChild(fillElement, 'patternFill');
    const type = getAttr(pattern, 'patternType');
    const foreground = firstChild(pattern, 'fgColor');
    const background = firstChild(pattern, 'bgColor');
    fills.push({
      color:
        type === 'solid'
          ? excelColor(foreground, theme) ?? excelColor(background, theme)
          : undefined,
    });
  }

  const borderStyleMap: Record<string, BorderSide['style']> = {
    thin: 'thin',
    medium: 'medium',
    thick: 'thick',
    double: 'double',
    dashed: 'dashed',
    dotted: 'dotted',
    hair: 'thin',
    mediumDashed: 'dashed',
    dashDot: 'dashed',
    slantDashDot: 'dashed',
  };
  const borders: CellBorders[] = [];
  for (const borderElement of Array.from(firstChild(root, 'borders')?.childNodes ?? [])) {
    if (!(borderElement instanceof Element) || borderElement.tagName !== 'border') continue;
    const parsed: CellBorders = {};
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      const sideElement = firstChild(borderElement, side);
      const style = getAttr(sideElement, 'style');
      if (!style || style === 'none') continue;
      parsed[side] = {
        style: borderStyleMap[style] ?? 'thin',
        color: excelColor(firstChild(sideElement, 'color'), theme) ?? '#000000',
      };
    }
    borders.push(parsed);
  }

  const xfs: Element[] = Array.from(firstChild(root, 'cellXfs')?.childNodes ?? []).filter(
    (node): node is Element => node instanceof Element && node.tagName === 'xf',
  );
  if (xfs.length === 0) {
    styles.default = {
      fontFamily: '等线',
      fontSizePt: 11,
      numFmt: 'General',
      horizontalAlign: 'general',
      verticalAlign: 'bottom',
    };
    return { styles, defaultStyleId: 'default' };
  }

  xfs.forEach((xf, index) => {
    const font = fonts[Number(getAttr(xf, 'fontId') ?? '0')] ?? {};
    const fill = fills[Number(getAttr(xf, 'fillId') ?? '0')];
    const border = borders[Number(getAttr(xf, 'borderId') ?? '0')];
    const numFmtId = Number(getAttr(xf, 'numFmtId') ?? '0');
    const alignment = firstChild(xf, 'alignment');
    const horizontalRaw = getAttr(alignment, 'horizontal');
    const verticalRaw = getAttr(alignment, 'vertical');
    const horizontalMap: Record<string, CellStyle['horizontalAlign']> = {
      general: 'general',
      left: 'left',
      center: 'center',
      right: 'right',
      fill: 'left',
      justify: 'left',
      centerContinuous: 'center',
      distributed: 'left',
    };
    const verticalMap: Record<string, CellStyle['verticalAlign']> = {
      top: 'top',
      center: 'middle',
      bottom: 'bottom',
      justify: 'middle',
      distributed: 'middle',
    };
    const style: CellStyle = {
      fontFamily: font.fontFamily ?? '等线',
      fontSizePt: font.fontSizePt ?? 11,
      bold: font.bold,
      italic: font.italic,
      underline: font.underline,
      strike: font.strike,
      color: font.color,
      fill: fill?.color,
      numFmtId,
      numFmt: numFmts.get(numFmtId) ?? 'General',
      horizontalAlign: horizontalMap[horizontalRaw ?? 'general'] ?? 'general',
      verticalAlign: verticalMap[verticalRaw ?? 'bottom'] ?? 'bottom',
      wrapText: getAttr(alignment, 'wrapText') === '1',
      indent: Number(getAttr(alignment, 'indent') ?? '0') || undefined,
      borders: border && Object.keys(border).length > 0 ? border : undefined,
      xfIndex: index,
    };
    styles[`s${index}`] = style;
  });
  return { styles, defaultStyleId: 's0' };
}

interface SheetDescriptor {
  name: string;
  state: SheetModel['state'];
  order: number;
  part: string;
}

async function parseSheet(
  pkg: OoxmlPackage,
  descriptor: SheetDescriptor,
  sharedStrings: string[],
  styles: Record<string, CellStyle>,
  defaultStyleId: string,
  theme: Record<string, string>,
): Promise<SheetModel> {
  const xml = await pkg.text(descriptor.part);
  const doc = parseXmlDom(xml);
  const worksheet = firstDescendant(doc, 'worksheet') ?? doc.documentElement;

  const sheet: SheetModel = {
    id: `sheet-${descriptor.order + 1}-${descriptor.name.replace(/\s+/g, '_')}`,
    name: descriptor.name,
    state: descriptor.state,
    order: descriptor.order,
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
    pageSetup: defaultSheetPageSetup(),
  };

  const sheetViews = firstChild(worksheet, 'sheetViews');
  const sheetView = firstChild(sheetViews, 'sheetView');
  if (sheetView) {
    if (getAttr(sheetView, 'showGridLines') === '0') sheet.showGridLines = false;
    const pane = firstChild(sheetView, 'pane');
    if (pane) {
      const xSplit = Number(getAttr(pane, 'xSplit') ?? '0');
      const ySplit = Number(getAttr(pane, 'ySplit') ?? '0');
      if (xSplit > 0 || ySplit > 0) sheet.freeze = { rows: ySplit, cols: xSplit };
    }
    const tabColor = excelColor(firstChild(sheetView, 'tabColor'), theme);
    if (tabColor) sheet.tabColor = tabColor;
  }

  const dimension = getAttr(firstChild(worksheet, 'dimension'), 'ref');

  // 功能：读取 XLSX 页面方向、缩放、页边距和打印辅助项；时间：2026-10-24；作者：lq。
  const pageSetup = firstChild(worksheet, 'pageSetup');
  const pageMargins = firstChild(worksheet, 'pageMargins');
  const printOptions = firstChild(worksheet, 'printOptions');
  const setup = sheet.pageSetup ?? defaultSheetPageSetup();
  sheet.pageSetup = {
    ...setup,
    orientation: getAttr(pageSetup, 'orientation') === 'portrait' ? 'portrait' : 'landscape',
    scalePercent: Number(getAttr(pageSetup, 'scale') ?? setup.scalePercent) || setup.scalePercent,
    fitToWidth: Number(getAttr(pageSetup, 'fitToWidth') ?? '') || undefined,
    fitToHeight: Number(getAttr(pageSetup, 'fitToHeight') ?? '') || undefined,
    firstPageNumber: Number(getAttr(pageSetup, 'firstPageNumber') ?? '') || undefined,
    margins: {
      left: Number(getAttr(pageMargins, 'left') ?? setup.margins.left) || setup.margins.left,
      right: Number(getAttr(pageMargins, 'right') ?? setup.margins.right) || setup.margins.right,
      top: Number(getAttr(pageMargins, 'top') ?? setup.margins.top) || setup.margins.top,
      bottom: Number(getAttr(pageMargins, 'bottom') ?? setup.margins.bottom) || setup.margins.bottom,
      header: Number(getAttr(pageMargins, 'header') ?? setup.margins.header) || setup.margins.header,
      footer: Number(getAttr(pageMargins, 'footer') ?? setup.margins.footer) || setup.margins.footer,
    },
    printGridLines: getAttr(printOptions, 'gridLines') !== '0',
    printHeadings: getAttr(printOptions, 'headings') === '1',
  };
  let maxRow = 0;
  let maxCol = 0;
  if (dimension) {
    const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(dimension);
    if (match) {
      maxCol = Math.max(0, columnIndexFromLabel(match[3]));
      maxRow = Math.max(0, Number(match[4]) - 1);
    }
  }

  const cols = firstChild(worksheet, 'cols');
  for (const col of Array.from(cols?.childNodes ?? [])) {
    if (!(col instanceof Element) || col.tagName !== 'col') continue;
    const min = Number(getAttr(col, 'min') ?? '1');
    const max = Number(getAttr(col, 'max') ?? String(min));
    const width = Number(getAttr(col, 'width') ?? '0');
    const hidden = getAttr(col, 'hidden') === '1';
    const pixelWidth = width > 0 ? Math.round(width * 7 + 5) : undefined;
    for (let index = min - 1; index < max; index += 1) {
      if (pixelWidth) sheet.colWidths[index] = pixelWidth;
      if (hidden) sheet.hiddenCols[index] = true;
      maxCol = Math.max(maxCol, index);
    }
  }

  const inlineSharedFormula = new Map<string, { formula: string; row: number; col: number }>();
  const cellsContainer = firstChild(worksheet, 'sheetData');
  const rowElements = Array.from(cellsContainer?.childNodes ?? []).filter(
    (node): node is Element => node instanceof Element && node.tagName === 'row',
  );

  // First pass: collect shared formula masters.
  for (const rowElement of rowElements) {
    for (const cellElement of Array.from(rowElement.childNodes)) {
      if (!(cellElement instanceof Element) || cellElement.tagName !== 'c') continue;
      const formulaElement = firstChild(cellElement, 'f');
      if (!formulaElement) continue;
      if (getAttr(formulaElement, 't') !== 'shared') continue;
      const si = getAttr(formulaElement, 'si');
      const text = formulaElement.textContent ?? '';
      if (!si || !text) continue;
      const reference = getAttr(cellElement, 'r') ?? 'A1';
      const position = parseReference(reference);
      if (!position) continue;
      inlineSharedFormula.set(si, { formula: text, row: position.row, col: position.col });
    }
  }

  for (const rowElement of rowElements) {
    const rowNumber = Number(getAttr(rowElement, 'r') ?? '0') - 1;
    if (rowNumber < 0) continue;
    const height = Number(getAttr(rowElement, 'ht') ?? '0');
    if (height > 0) sheet.rowHeights[rowNumber] = Math.round(height * (4 / 3));
    if (getAttr(rowElement, 'hidden') === '1') sheet.hiddenRows[rowNumber] = true;
    maxRow = Math.max(maxRow, rowNumber);

    for (const cellElement of Array.from(rowElement.childNodes)) {
      if (!(cellElement instanceof Element) || cellElement.tagName !== 'c') continue;
      const reference = getAttr(cellElement, 'r');
      if (!reference) continue;
      const position = parseReference(reference);
      if (!position) continue;
      const { row, col } = position;
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
      const type = getAttr(cellElement, 't') ?? 'n';
      const styleIdRaw = getAttr(cellElement, 's');
      const styleId = styleIdRaw !== undefined ? `s${styleIdRaw}` : defaultStyleId;
      const style = styles[styleId];
      const cell: Cell = { value: null, styleId: styles[styleId] ? styleId : undefined };

      const formulaElement = firstChild(cellElement, 'f');
      if (formulaElement) {
        let formulaText = formulaElement.textContent ?? '';
        const sharedType = getAttr(formulaElement, 't');
        const si = getAttr(formulaElement, 'si');
        if (sharedType === 'shared' && si) {
          if (!formulaText) {
            const master = inlineSharedFormula.get(si);
            if (master) {
              formulaText = translateFormula(`=${master.formula}`, row - master.row, col - master.col).slice(1);
            }
          }
        }
        if (formulaText) cell.formula = formulaText.startsWith('=') ? formulaText : `=${formulaText}`;
      }

      if (type === 'inlineStr') {
        const inline = firstChild(cellElement, 'is');
        cell.value = inline?.textContent ?? '';
      } else {
        const valueElement = firstChild(cellElement, 'v');
        const raw = valueElement?.textContent ?? '';
        if (type === 's') {
          const index = Number(raw);
          cell.value = Number.isFinite(index) ? sharedStrings[index] ?? '' : '';
        } else if (type === 'str') {
          cell.value = raw;
        } else if (type === 'b') {
          cell.value = raw === '1' || raw.toLowerCase() === 'true';
        } else if (type === 'e') {
          cell.value = raw;
        } else if (type === 'd') {
          const parsed = new Date(raw);
          cell.value = Number.isNaN(parsed.getTime()) ? raw : Math.round(parsed.getTime() / 86_400_000) + 25_569;
        } else if (raw !== '') {
          const numeric = Number(raw);
          cell.value = Number.isFinite(numeric) ? numeric : raw;
        }
      }

      if (cell.formula && cell.value === null) cell.value = null;
      if (cell.value !== null || cell.formula) sheet.cells[`${row}:${col}`] = cell;
      void style;
    }
  }

  for (const mergeElement of Array.from(firstChild(worksheet, 'mergeCells')?.childNodes ?? [])) {
    if (!(mergeElement instanceof Element) || mergeElement.tagName !== 'mergeCell') continue;
    const ref = getAttr(mergeElement, 'ref');
    if (!ref) continue;
    const [start, end] = ref.split(':');
    const first = parseReference(start);
    const last = parseReference(end ?? start);
    if (first && last) {
      sheet.merges.push({
        r1: Math.min(first.row, last.row),
        c1: Math.min(first.col, last.col),
        r2: Math.max(first.row, last.row),
        c2: Math.max(first.col, last.col),
      });
      maxRow = Math.max(maxRow, first.row, last.row);
      maxCol = Math.max(maxCol, first.col, last.col);
    }
  }

  const autoFilter = getAttr(firstChild(worksheet, 'autoFilter'), 'ref');
  if (autoFilter) {
    const [start, end] = autoFilter.split(':');
    const first = parseReference(start);
    const last = parseReference(end ?? start);
    if (first && last) {
      sheet.autoFilter = {
        r1: Math.min(first.row, last.row),
        c1: Math.min(first.col, last.col),
        r2: Math.max(first.row, last.row),
        c2: Math.max(first.col, last.col),
      };
    }
  }

  sheet.dimension = {
    rows: Math.min(1_048_576, Math.max(maxRow + 80, 100)),
    cols: Math.min(16_384, Math.max(maxCol + 20, 26)),
  };
  return sheet;
}

function columnIndexFromLabel(label: string): number {
  let index = 0;
  for (const char of label.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64);
  return index - 1;
}

function parseReference(reference: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(reference.trim());
  if (!match) return null;
  return { row: Number(match[2]) - 1, col: columnIndexFromLabel(match[1]) };
}

export async function parseWorkbookPackage(
  pkg: OoxmlPackage,
  fileName: string,
): Promise<WorkbookModel> {
  const theme = await readThemeColors(pkg, 'xl/theme/theme1.xml');
  const sharedStrings = pkg.has('xl/sharedStrings.xml')
    ? parseSharedStrings(await pkg.text('xl/sharedStrings.xml'))
    : [];
  const parsedStyles = pkg.has('xl/styles.xml')
    ? parseStyles(await pkg.text('xl/styles.xml'), theme)
    : parseStyles('', theme);

  const workbookXml = await pkg.text('xl/workbook.xml');
  const workbookDoc = parseXmlDom(workbookXml);
  const workbookRoot = firstDescendant(workbookDoc, 'workbook') ?? workbookDoc.documentElement;
  const workbookPr = firstChild(workbookRoot, 'workbookPr');
  const date1904 = getAttr(workbookPr, 'date1904') === '1';
  const rels = await relationshipMap(pkg, 'xl/workbook.xml');
  const sheetElements = Array.from(firstChild(workbookRoot, 'sheets')?.childNodes ?? []).filter(
    (node): node is Element => node instanceof Element && node.tagName === 'sheet',
  );

  const descriptors: SheetDescriptor[] = [];
  sheetElements.forEach((sheetElement, index) => {
    const rid = getAttr(sheetElement, 'r:id') ?? getAttrLocal(sheetElement, 'id') ?? '';
    const part = rels.get(rid)?.resolved;
    if (!part) return;
    const stateRaw = getAttr(sheetElement, 'state') ?? 'visible';
    descriptors.push({
      name: getAttr(sheetElement, 'name') ?? `Sheet${index + 1}`,
      state: stateRaw === 'hidden' || stateRaw === 'veryHidden' ? stateRaw : 'visible',
      order: index,
      part,
    });
  });

  const sheets: SheetModel[] = [];
  for (const descriptor of descriptors) {
    sheets.push(
      await parseSheet(pkg, descriptor, sharedStrings, parsedStyles.styles, parsedStyles.defaultStyleId, theme),
    );
  }
  if (sheets.length === 0) {
    const fallback = createBlankWorkbook().sheets[0];
    fallback.name = 'Sheet1';
    sheets.push(fallback);
  }

  const definedNames: Record<string, string> = {};
  for (const definedName of Array.from(firstChild(workbookRoot, 'definedNames')?.childNodes ?? [])) {
    if (!(definedName instanceof Element) || definedName.tagName !== 'definedName') continue;
    const name = getAttr(definedName, 'name');
    const value = definedName.textContent?.trim();
    if (name && value && !name.startsWith('_xlnm.')) definedNames[name] = value;
  }

  const workbookView = firstChild(firstChild(workbookRoot, 'bookViews'), 'workbookView');
  const activeTab = Number(getAttr(workbookView, 'activeTab') ?? '0');
  const activeSheet = sheets[Math.min(Math.max(0, activeTab), sheets.length - 1)];

  const model: WorkbookModel = {
    version: 1,
    sheets,
    activeSheetId: activeSheet.id,
    styles: parsedStyles.styles,
    themeColors: theme,
    themeFonts: { major: '', minor: '' },
    defaultFont: parsedStyles.styles[parsedStyles.defaultStyleId]?.fontFamily ?? '等线',
    defaultFontSizePt: parsedStyles.styles[parsedStyles.defaultStyleId]?.fontSizePt ?? 11,
    date1904,
    definedNames: Object.keys(definedNames).length ? definedNames : undefined,
    meta: { title: stripExtension(baseName(fileName)) },
    stats: {
      cells: sheets.reduce((sum, sheet) => sum + Object.keys(sheet.cells).length, 0),
      sheets: sheets.length,
    },
  };
  return recalculateWorkbook(model);
}

// ---- CSV ---------------------------------------------------------------

function detectDelimiter(text: string): string {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? '';
  let inQuotes = false;
  let comma = 0;
  let semicolon = 0;
  let tab = 0;
  for (const char of firstLine) {
    if (char === '"') inQuotes = !inQuotes;
    else if (!inQuotes) {
      if (char === ',') comma += 1;
      else if (char === ';') semicolon += 1;
      else if (char === '\t') tab += 1;
    }
  }
  if (tab > comma && tab >= semicolon) return '\t';
  if (semicolon > comma) return ';';
  return ',';
}

export function parseCsvRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let index = 0;
  const input = text.replace(/^\uFEFF/, '');
  while (index < input.length) {
    const char = input[index];
    if (inQuotes) {
      if (char === '"') {
        if (input[index + 1] === '"') {
          field += '"';
          index += 2;
          continue;
        }
        inQuotes = false;
        index += 1;
        continue;
      }
      field += char;
      index += 1;
      continue;
    }
    if (char === '"') {
      inQuotes = true;
      index += 1;
      continue;
    }
    if (char === delimiter) {
      row.push(field);
      field = '';
      index += 1;
      continue;
    }
    if (char === '\r') {
      if (input[index + 1] === '\n') index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      index += 1;
      continue;
    }
    if (char === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      index += 1;
      continue;
    }
    field += char;
    index += 1;
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function parseCsvScalar(text: string): CellScalar {
  if (text === '') return null;
  const trimmed = text.trim();
  if (/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) return Number(trimmed);
  if (/^TRUE$/i.test(trimmed)) return true;
  if (/^FALSE$/i.test(trimmed)) return false;
  return text;
}

export function parseCsv(text: string, fileName: string): WorkbookModel {
  const rows = parseCsvRows(text, detectDelimiter(text));
  const sheet = createBlankWorkbook().sheets[0];
  sheet.name = stripExtension(baseName(fileName)) || 'Sheet1';
  sheet.cells = {};
  let maxCol = 0;
  rows.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (value === '') return;
      maxCol = Math.max(maxCol, colIndex);
      const parsed = parseCsvScalar(value);
      if (typeof parsed === 'string' && parsed.startsWith('=')) {
        sheet.cells[`${rowIndex}:${colIndex}`] = { value: null, formula: parsed };
      } else if (parsed !== null) {
        sheet.cells[`${rowIndex}:${colIndex}`] = { value: parsed };
      }
    });
  });
  sheet.dimension = {
    rows: Math.max(100, rows.length + 50),
    cols: Math.max(26, maxCol + 10),
  };
  const base = createBlankWorkbook();
  const model: WorkbookModel = {
    ...base,
    sheets: [{ ...sheet, id: 'sheet-csv-1', order: 0 }],
    activeSheetId: 'sheet-csv-1',
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: { title: stripExtension(baseName(fileName)) },
    csv: true,
    stats: { cells: Object.keys(sheet.cells).length, sheets: 1 },
  };
  return recalculateWorkbook(model);
}

export function createSheetDescriptor(
  sheets: WorkbookModel['sheets'],
): typeof sheets {
  return sheets;
}
