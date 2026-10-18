import JSZip from 'jszip';
import { BUILTIN_NUMBER_FORMATS, formatCellValue } from './format';
import type {
  Cell,
  CellBorders,
  CellStyle,
  SheetModel,
  WorkbookModel,
} from '@mini-office/document-model';
import { escapeXml, getAttrLocal, parseXmlDom, xmlDeclaration } from '@mini-office/ooxml';
import { recalculateWorkbook } from './formula/engine';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const CORE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const APP_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

function columnLabel(index: number): string {
  let n = index + 1;
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function argb(hex: string | undefined): string | undefined {
  if (!hex) return undefined;
  const cleaned = hex.replace('#', '');
  if (/^[0-9a-fA-F]{6}$/.test(cleaned)) return `FF${cleaned.toUpperCase()}`;
  if (/^[0-9a-fA-F]{8}$/.test(cleaned)) return cleaned.toUpperCase();
  return undefined;
}

interface GeneratedStyles {
  xml: string;
  index: Map<string, number>;
}

function borderSideXml(tag: string, side: CellBorders['top']): string {
  if (!side || side.style === 'none') return `<${tag}/>`;
  const styleMap: Record<string, string> = {
    thin: 'thin',
    medium: 'medium',
    thick: 'thick',
    double: 'double',
    dashed: 'dashed',
    dotted: 'dotted',
  };
  return `<${tag} style="${styleMap[side.style] ?? 'thin'}"><color rgb="${argb(side.color ?? '#000000') ?? 'FF000000'}"/></${tag}>`;
}

function borderXml(border: CellBorders | undefined): string {
  return `<border>${borderSideXml('left', border?.left)}${borderSideXml('right', border?.right)}${borderSideXml('top', border?.top)}${borderSideXml('bottom', border?.bottom)}<diagonal/></border>`;
}

function buildStyles(workbook: WorkbookModel): GeneratedStyles {
  const styles = workbook.styles;
  const styleEntries = Object.entries(styles);
  const defaultId = styles.default ? 'default' : styleEntries[0]?.[0] ?? 'default';
  const orderedIds = [
    defaultId,
    ...styleEntries.map(([id]) => id).filter((id) => id !== defaultId),
  ];

  const fontKeys = new Map<string, number>();
  const fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];
  const fillKeys = new Map<string, number>([
    ['none', 0],
    ['gray125', 1],
  ]);
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const borderKeys = new Map<string, number>([['empty', 0]]);
  const numFmtIds = new Map<string, number>();
  const customNumFmts: Array<{ id: number; code: string }> = [];
  let nextNumFmtId = 164;

  for (const [id, pattern] of Object.entries(BUILTIN_NUMBER_FORMATS)) {
    numFmtIds.set(pattern, Number(id));
  }
  for (const style of Object.values(styles)) {
    const pattern = style.numFmt;
    if (!pattern || pattern === 'General' || numFmtIds.has(pattern)) continue;
    const existing = customNumFmts.find((item) => item.code === pattern);
    if (existing) {
      numFmtIds.set(pattern, existing.id);
      continue;
    }
    const idNumber =
      style.numFmtId && style.numFmtId >= 164 ? style.numFmtId : nextNumFmtId++;
    customNumFmts.push({ id: idNumber, code: pattern });
    numFmtIds.set(pattern, idNumber);
  }

  const fontXmls: string[] = [];
  const addFont = (style: CellStyle): number => {
    const key = JSON.stringify([
      style.fontFamily ?? workbook.defaultFont,
      style.fontSizePt ?? workbook.defaultFontSizePt,
      !!style.bold,
      !!style.italic,
      !!style.underline,
      !!style.strike,
      style.color ?? '',
    ]);
    const existing = fontKeys.get(key);
    if (existing !== undefined) return existing;
    const parts = [
      `<sz val="${style.fontSizePt ?? workbook.defaultFontSizePt}"/>`,
      style.color ? `<color rgb="${argb(style.color)}"/>` : '<color theme="1"/>',
      `<name val="${escapeXml(style.fontFamily ?? workbook.defaultFont)}"/>`,
      '<family val="2"/>',
    ];
    if (style.bold) parts.push('<b/>');
    if (style.italic) parts.push('<i/>');
    if (style.underline) parts.push('<u/>');
    if (style.strike) parts.push('<strike/>');
    const index = fontXmls.length;
    fontXmls.push(`<font>${parts.join('')}</font>`);
    fontKeys.set(key, index);
    return index;
  };

  const addFill = (style: CellStyle): number => {
    if (!style.fill) return 0;
    const key = style.fill.toUpperCase();
    const existing = fillKeys.get(key);
    if (existing !== undefined) return existing;
    const index = fills.length;
    fills.push(
      `<fill><patternFill patternType="solid"><fgColor rgb="${argb(style.fill)}"/><bgColor indexed="64"/></patternFill></fill>`,
    );
    fillKeys.set(key, index);
    return index;
  };

  const addBorder = (style: CellStyle): number => {
    const key = JSON.stringify(style.borders ?? {});
    const existing = borderKeys.get(key);
    if (existing !== undefined) return existing;
    const index = borders.length;
    borders.push(borderXml(style.borders));
    borderKeys.set(key, index);
    return index;
  };

  const xfXmls: string[] = [];
  const index = new Map<string, number>();
  for (const id of orderedIds) {
    const style = styles[id] ?? {};
    const fontId = addFont(style);
    const fillId = addFill(style);
    const borderId = addBorder(style);
    const pattern = style.numFmt ?? 'General';
    const numFmtId =
      style.numFmtId && !numFmtIds.has(pattern) ? style.numFmtId : numFmtIds.get(pattern) ?? 0;
    const alignmentParts: string[] = [];
    if (style.horizontalAlign && style.horizontalAlign !== 'general') {
      alignmentParts.push(`horizontal="${style.horizontalAlign}"`);
    }
    if (style.verticalAlign && style.verticalAlign !== 'bottom') {
      alignmentParts.push(
        `vertical="${style.verticalAlign === 'middle' ? 'center' : style.verticalAlign}"`,
      );
    }
    if (style.wrapText) alignmentParts.push('wrapText="1"');
    if (style.indent) alignmentParts.push(`indent="${style.indent}"`);
    const alignment = alignmentParts.length ? `<alignment ${alignmentParts.join(' ')}/>` : '';
    xfXmls.push(
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${alignment ? ' applyAlignment="1"' : ''}>${alignment}</xf>`,
    );
    index.set(id, xfXmls.length - 1);
  }
  if (!index.has('default')) index.set('default', 0);

  const numFmtXml = customNumFmts.length
    ? `<numFmts count="${customNumFmts.length}">${customNumFmts
        .map(
          (item) =>
            `<numFmt numFmtId="${item.id}" formatCode="${escapeXml(item.code)}"/>`,
        )
        .join('')}</numFmts>`
    : '';

  const xml = `${xmlDeclaration()}<styleSheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${numFmtXml}<fonts count="${fontXmls.length}">${fontXmls.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfXmls.length}">${xfXmls.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  return { xml, index };
}

interface SheetComputed {
  value: string | number | boolean | null;
  display: string;
  error?: string;
}

function cellXml(
  cell: Cell,
  computed: SheetComputed | undefined,
  styleIndex: number,
): string {
  const styleAttr = styleIndex > 0 ? ` s="${styleIndex}"` : '';
  if (cell.formula) {
    const formula = escapeXml(cell.formula.replace(/^=/, ''));
    if (computed?.error) {
      return `<c${styleAttr} t="e"><f>${formula}</f><v>${escapeXml(computed.error)}</v></c>`;
    }
    const value = computed?.value ?? cell.value;
    if (value === null || value === undefined) {
      return `<c${styleAttr}><f>${formula}</f></c>`;
    }
    if (typeof value === 'number') {
      return `<c${styleAttr}><f>${formula}</f><v>${Number.isFinite(value) ? value : 0}</v></c>`;
    }
    if (typeof value === 'boolean') {
      return `<c${styleAttr} t="b"><f>${formula}</f><v>${value ? 1 : 0}</v></c>`;
    }
    return `<c${styleAttr} t="str"><f>${formula}</f><v>${escapeXml(String(value))}</v></c>`;
  }
  const value = cell.value;
  if (value === null || value === undefined) {
    return styleAttr ? `<c${styleAttr}/>` : '';
  }
  if (typeof value === 'number') {
    return `<c${styleAttr}><v>${Number.isFinite(value) ? value : 0}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c${styleAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  if (/^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A|CIRC!)$/i.test(value)) {
    return `<c${styleAttr} t="e"><v>${escapeXml(value)}</v></c>`;
  }
  return `<c${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function sheetXml(
  sheet: SheetModel,
  workbook: WorkbookModel,
  styleIndex: Map<string, number>,
): string {
  const computed = workbook.computed?.[sheet.id] ?? {};
  const rows = new Map<number, Array<{ col: number; cell: Cell }>>();
  let maxRow = 0;
  let maxCol = 0;
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const [row, col] = key.split(':').map(Number);
    if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
    const list = rows.get(row) ?? [];
    list.push({ col, cell });
    rows.set(row, list);
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  for (const merge of sheet.merges) {
    maxRow = Math.max(maxRow, merge.r2);
    maxCol = Math.max(maxCol, merge.c2);
  }

  const rowXml: string[] = [];
  for (const [rowIndex, cells] of [...rows.entries()].sort((a, b) => a[0] - b[0])) {
    cells.sort((a, b) => a.col - b.col);
    const height = sheet.rowHeights[rowIndex];
    const attrs = [
      `r="${rowIndex + 1}"`,
      height ? `ht="${Math.round(height * 0.75 * 100) / 100}"` : '',
      height ? 'customHeight="1"' : '',
      sheet.hiddenRows[rowIndex] ? 'hidden="1"' : '',
    ]
      .filter(Boolean)
      .join(' ');
    const cellXmls = cells
      .map(({ col, cell }) => {
        const reference = `${columnLabel(col)}${rowIndex + 1}`;
        const style = styleIndex.get(cell.styleId ?? 'default') ?? 0;
        const inner = cellXml(cell, computed[cellKey(rowIndex, col)], style);
        if (!inner) return '';
        return inner.replace(/^<c/, `<c r="${reference}"`);
      })
      .join('');
    if (cellXmls) rowXml.push(`<row ${attrs}>${cellXmls}</row>`);
    else if (height || sheet.hiddenRows[rowIndex]) rowXml.push(`<row ${attrs}/>`);
  }

  const colsXml: string[] = [];
  const colWidths = Object.entries(sheet.colWidths)
    .map(([col, width]) => ({ col: Number(col), width }))
    .sort((a, b) => a.col - b.col);
  for (const { col, width } of colWidths) {
    const chars = Math.max(1, Math.round(((width - 5) / 7) * 100) / 100);
    colsXml.push(
      `<col min="${col + 1}" max="${col + 1}" width="${chars}" customWidth="1"${sheet.hiddenCols[col] ? ' hidden="1"' : ''}/>`,
    );
  }
  for (const [col] of Object.entries(sheet.hiddenCols)) {
    if (sheet.colWidths[Number(col)]) continue;
    colsXml.push(`<col min="${Number(col) + 1}" max="${Number(col) + 1}" width="8.43" hidden="1"/>`);
  }

  const freeze = sheet.freeze;
  const pane = freeze
    ? `<pane xSplit="${freeze.cols}" ySplit="${freeze.rows}" topLeftCell="${columnLabel(freeze.cols)}${freeze.rows + 1}" activePane="${freeze.rows > 0 && freeze.cols > 0 ? 'bottomRight' : freeze.rows > 0 ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
    : '';
  const sheetViews = `<sheetViews><sheetView${sheet.showGridLines ? '' : ' showGridLines="0"'} workbookViewId="0">${pane}</sheetView></sheetViews>`;

  const mergeXml = sheet.merges.length
    ? `<mergeCells count="${sheet.merges.length}">${sheet.merges
        .map(
          (merge) =>
            `<mergeCell ref="${columnLabel(merge.c1)}${merge.r1 + 1}:${columnLabel(merge.c2)}${merge.r2 + 1}"/>`,
        )
        .join('')}</mergeCells>`
    : '';

  const filter = sheet.autoFilter
    ? `<autoFilter ref="${columnLabel(sheet.autoFilter.c1)}${sheet.autoFilter.r1 + 1}:${columnLabel(sheet.autoFilter.c2)}${sheet.autoFilter.r2 + 1}"/>`
    : '';

  const dimension = maxRow > 0 || maxCol > 0 ? `A1:${columnLabel(maxCol)}${maxRow + 1}` : 'A1';
  const tabColor = sheet.tabColor
    ? `<sheetPr><tabColor rgb="${argb(sheet.tabColor)}"/></sheetPr>`
    : '';

  return `${xmlDeclaration()}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${tabColor}<dimension ref="${dimension}"/>${sheetViews}<sheetFormatPr defaultRowHeight="15"/>${
    colsXml.length ? `<cols>${colsXml.join('')}</cols>` : ''
  }<sheetData>${rowXml.join('')}</sheetData>${mergeXml}${filter}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

interface RelationshipInfo {
  id: string;
  type: string;
  target: string;
}

function parseRelationships(xml: string | undefined): RelationshipInfo[] {
  if (!xml) return [];
  try {
    const doc = parseXmlDom(xml);
    return Array.from(doc.getElementsByTagName('Relationship')).map((element) => ({
      id: getAttrLocal(element, 'Id') ?? '',
      type: getAttrLocal(element, 'Type') ?? '',
      target: getAttrLocal(element, 'Target') ?? '',
    }));
  } catch {
    return [];
  }
}

function workbookXml(workbook: WorkbookModel): string {
  const activeIndex = Math.max(
    0,
    workbook.sheets.findIndex((sheet) => sheet.id === workbook.activeSheetId),
  );
  const sheets = workbook.sheets
    .map(
      (sheet, index) =>
        `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}"${sheet.state !== 'visible' ? ` state="${sheet.state}"` : ''} r:id="rIdSheet${index + 1}"/>`,
    )
    .join('');
  const definedNames =
    workbook.definedNames && Object.keys(workbook.definedNames).length
      ? `<definedNames>${Object.entries(workbook.definedNames)
          .map(
            ([name, value]) =>
              `<definedName name="${escapeXml(name)}">${escapeXml(value)}</definedName>`,
          )
          .join('')}</definedNames>`
      : '';
  return `${xmlDeclaration()}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27626"/><workbookPr${workbook.date1904 ? ' date1904="1"' : ''}/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000" activeTab="${activeIndex}"/></bookViews><sheets>${sheets}</sheets>${definedNames}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
}

function workbookRelsXml(
  sourceRels: RelationshipInfo[],
  workbook: WorkbookModel,
): string {
  const preserved = sourceRels.filter(
    (rel) =>
      !rel.type.endsWith('/worksheet') &&
      !rel.type.endsWith('/styles') &&
      !rel.type.endsWith('/sharedStrings') &&
      !rel.type.endsWith('/calcChain'),
  );
  const rels: RelationshipInfo[] = workbook.sheets.map((_sheet, index) => ({
    id: `rIdSheet${index + 1}`,
    type: `${REL_NS}/worksheet`,
    target: `worksheets/sheet${index + 1}.xml`,
  }));
  rels.push({ id: 'rIdStyles', type: `${REL_NS}/styles`, target: 'styles.xml' });
  for (const rel of preserved) {
    if (rels.some((item) => item.id === rel.id)) continue;
    rels.push(rel);
  }
  return `${xmlDeclaration()}<Relationships xmlns="${PKG_REL_NS}">${rels
    .map(
      (rel) =>
        `<Relationship Id="${escapeXml(rel.id)}" Type="${escapeXml(rel.type)}" Target="${escapeXml(rel.target)}"/>`,
    )
    .join('')}</Relationships>`;
}

function contentTypesXml(sheetCount: number): string {
  const sheets = Array.from(
    { length: sheetCount },
    (_, index) =>
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
  ).join('');
  return `${xmlDeclaration()}<Types xmlns="${CONTENT_TYPE_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
}

function packageRelsXml(): string {
  return `${xmlDeclaration()}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${DOC_REL_NS}" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${CORE_REL_NS}" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${APP_REL_NS}" Target="docProps/app.xml"/></Relationships>`;
}

function coreXml(workbook: WorkbookModel): string {
  const now = new Date().toISOString();
  return `${xmlDeclaration()}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(workbook.meta.title ?? 'Mini Office 工作簿')}</dc:title><dc:creator>${escapeXml(workbook.meta.author ?? 'Mini Office')}</dc:creator><cp:lastModifiedBy>Mini Office</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${workbook.meta.created ?? now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${workbook.meta.modified ?? now}</dcterms:modified></cp:coreProperties>`;
}

function appXml(): string {
  return `${xmlDeclaration()}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Mini Office</Application><AppVersion>0.1</AppVersion></Properties>`;
}

function updateContentTypes(xml: string, sheetCount: number): string {
  let next = xml;
  for (let index = 1; index <= sheetCount; index += 1) {
    if (next.includes(`PartName="/xl/worksheets/sheet${index}.xml"`)) continue;
    next = next.replace(
      '</Types>',
      `<Override PartName="/xl/worksheets/sheet${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    );
  }
  return next;
}

export interface SerializeWorkbookOptions {
  sourceBytes?: Uint8Array;
  name: string;
}

export async function serializeXlsx(
  workbook: WorkbookModel,
  options: SerializeWorkbookOptions,
): Promise<Uint8Array> {
  const recalculated = recalculateWorkbook(workbook);
  const zip = options.sourceBytes ? await JSZip.loadAsync(options.sourceBytes) : new JSZip();
  const source = !!options.sourceBytes;
  const sourceRelsXml = source
    ? await zip.file('xl/_rels/workbook.xml.rels')?.async('text')
    : undefined;
  const sourceRels = parseRelationships(sourceRelsXml);
  const sourceContentTypes = source
    ? await zip.file('[Content_Types].xml')?.async('text')
    : undefined;
  const styles = buildStyles(recalculated);
  const styleIndex = styles.index;

  if (!source) {
    zip.file('[Content_Types].xml', contentTypesXml(recalculated.sheets.length));
    zip.file('_rels/.rels', packageRelsXml());
    zip.file('docProps/core.xml', coreXml(recalculated));
    zip.file('docProps/app.xml', appXml());
  } else {
    zip.remove('xl/calcChain.xml');
    zip.remove('xl/sharedStrings.xml');
    zip.remove('xl/worksheets');
    let contentTypes = sourceContentTypes ?? contentTypesXml(recalculated.sheets.length);
    contentTypes = contentTypes
      .replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '')
      .replace(/<Override[^>]*PartName="\/xl\/sharedStrings\.xml"[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', updateContentTypes(contentTypes, recalculated.sheets.length));
    if (!zip.file('docProps/core.xml')) zip.file('docProps/core.xml', coreXml(recalculated));
    if (!zip.file('docProps/app.xml')) zip.file('docProps/app.xml', appXml());
  }

  zip.file('xl/workbook.xml', workbookXml(recalculated));
  zip.file('xl/_rels/workbook.xml.rels', workbookRelsXml(sourceRels, recalculated));
  zip.file('xl/styles.xml', styles.xml);
  recalculated.sheets.forEach((sheet, index) => {
    zip.file(`xl/worksheets/sheet${index + 1}.xml`, sheetXml(sheet, recalculated, styleIndex));
  });

  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function csvScalar(value: string | number | boolean | null): string {
  if (value === null || value === undefined) return '';
  return String(value);
}

export function serializeCsv(workbook: WorkbookModel): Uint8Array {
  const sheet =
    workbook.sheets.find((item) => item.id === workbook.activeSheetId) ?? workbook.sheets[0];
  let maxRow = 0;
  let maxCol = 0;
  for (const key of Object.keys(sheet.cells)) {
    const [row, col] = key.split(':').map(Number);
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  const computed = workbook.computed?.[sheet.id] ?? {};
  const lines: string[] = [];
  for (let row = 0; row <= maxRow; row += 1) {
    const fields: string[] = [];
    for (let col = 0; col <= maxCol; col += 1) {
      const result = computed[cellKey(row, col)];
      const cell = sheet.cells[cellKey(row, col)];
      const value = result
        ? result.display
        : cell?.value === null || cell?.value === undefined
          ? ''
          : formatCellValue(cell.value, workbook.styles[cell.styleId ?? 'default']);
      const text = csvScalar(value);
      fields.push(/[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    lines.push(fields.join(','));
  }
  return new TextEncoder().encode(`\uFEFF${lines.join('\r\n')}`);
}

