import { createBlankWorkbook, type CellScalar, type DocBlock, type DocModel, type DocRun, type WorkbookModel } from '@mini-office/document-model';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { formatCellValue, parseCsvRows, parseWorkbookPackage } from '@mini-office/editor-sheet';
import { mapParagraphs } from './doc-utils';

/** 功能：合并记录保留源行号、显示值及筛选原值，避免编号/金额失真；时间：2026-10-24；作者：lq。 */
export interface MergeRecord { rowNumber: number; values: Record<string, string>; raw: Record<string, CellScalar>; errors: string[] }
export interface MergeData { fields: string[]; records: MergeRecord[] }
export interface MergeCondition { field: string; op: 'contains' | 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte' | 'empty' | 'notEmpty'; value: string }
export const MAX_MERGE_RECORDS = 2000;
export const MAX_MERGE_BATCH = 200;

/** 功能：读取 XLSX 或 UTF-8/GB18030 CSV，CSV 所有值按文本保留，不执行公式；时间：2026-10-24；作者：lq。 */
export async function readMergeWorkbook(bytes: Uint8Array, name: string): Promise<WorkbookModel> {
  if (bytes.byteLength > 50 * 1024 * 1024) throw new Error('数据文件不能超过 50 MB。');
  if (/\.csv$/i.test(name)) {
    let text: string;
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch { text = new TextDecoder('gb18030', { fatal: true }).decode(bytes); }
    let quotes = false;
    for (let i = 0; i < text.length; i++) if (text[i] === '"') {
      if (quotes && text[i + 1] === '"') i++;
      else quotes = !quotes;
    }
    if (quotes) throw new Error('CSV 引号没有闭合，请检查数据文件。');
    const candidates = [',', '\t', ';'].map((delimiter) => parseCsvRows(text, delimiter));
    const rows = candidates.sort((a, b) => (b[0]?.length ?? 0) - (a[0]?.length ?? 0))[0];
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.name = name.replace(/\.csv$/i, '');
    rows.forEach((row, r) => row.forEach((value, c) => { if (value !== '') sheet.cells[`${r}:${c}`] = { value }; }));
    sheet.dimension = { rows: rows.length, cols: rows.reduce((max, row) => Math.max(max, row.length), 0) };
    return workbook;
  }
  if (!/\.xlsx$/i.test(name)) throw new Error('请选择 XLSX 或 CSV 文件，旧版 XLS 请先另存为 XLSX。');
  const pkg = await loadOoxmlPackage(bytes);
  try { return await parseWorkbookPackage(pkg, name); } finally { pkg.dispose(); }
}

/** 功能：只提取实际有值区域，校验表头、合并单元格和公式结果；时间：2026-10-24；作者：lq。 */
export function extractMergeData(workbook: WorkbookModel, sheetId: string, headerRow: number): MergeData {
  const sheet = workbook.sheets.find((item) => item.id === sheetId);
  if (!sheet) throw new Error('请选择工作表。');
  if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 1048576) throw new Error('表头行号应为正整数。');
  const row = headerRow - 1;
  const keys = Object.entries(sheet.cells).filter(([, cell]) => cell.formula || (cell.value !== null && cell.value !== '')).map(([key]) => key.split(':').map(Number)).filter(([r]) => r >= row);
  if (!keys.length) throw new Error('表头行及其下方没有数据。');
  const cols = [...new Set(keys.map(([, c]) => c))].sort((a, b) => a - b);
  if (cols.length > 200) throw new Error('最多支持 200 个数据字段。');
  const fields = cols.map((c) => String(sheet.cells[`${row}:${c}`]?.value ?? '').trim());
  if (fields.some((field) => !field)) throw new Error('有数据的列缺少表头，请为每列填写唯一名称。');
  if (new Set(fields).size !== fields.length) throw new Error('表头名称重复，请改为不同名称。');
  if (sheet.merges.some((range) => range.r2 >= row && range.r1 <= keys.reduce((max, [r]) => Math.max(max, r), row) && cols.some((c) => c >= range.c1 && c <= range.c2))) throw new Error('数据区域含合并单元格，请取消合并后导入。');
  const sourceRows = [...new Set(keys.map(([r]) => r))].filter((r) => r > row).sort((a, b) => a - b);
  if (sourceRows.length > MAX_MERGE_RECORDS) throw new Error(`最多导入 ${MAX_MERGE_RECORDS} 条记录，请拆分数据文件。`);
  const records = sourceRows.map((r): MergeRecord => {
    const values: Record<string, string> = Object.create(null);
    const raw: Record<string, CellScalar> = Object.create(null);
    const errors: string[] = [];
    cols.forEach((c, index) => {
      const cell = sheet.cells[`${r}:${c}`];
      const computed = workbook.computed?.[sheetId]?.[`${r}:${c}`];
      const value = cell?.formula ? computed?.value ?? cell.value : cell?.value ?? null;
      const field = fields[index];
      if (computed?.error || (cell?.formula && value === null) || (typeof value === 'string' && /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|N\/A|NUM!|NULL!|SPILL!)/.test(value))) errors.push(`${field}：公式结果无效，请在 Excel 中计算后保存`);
      if (typeof value === 'number' && (!Number.isFinite(value) || (Number.isInteger(value) && Math.abs(value) >= 1e15))) errors.push(`${field}：数值超过可保证的精度，请核对原值并使用文本保存长编号`);
      raw[field] = value;
      const style = cell?.styleId ? workbook.styles[cell.styleId] : undefined;
      // 功能：保留常见 0000 编号格式的前导零；时间：2026-10-24；作者：lq。
      values[field] = typeof value === 'number' && Number.isInteger(value) && /^0+$/.test(style?.numFmt ?? '') ? `${value < 0 ? '-' : ''}${String(Math.abs(value)).padStart(style!.numFmt!.length, '0')}` : formatCellValue(value, style, workbook.date1904);
    });
    return { rowNumber: r + 1, values, raw, errors };
  }).filter((record) => record.errors.length > 0 || Object.values(record.values).some((value) => value.trim() !== ''));
  if (!records.length) throw new Error('表头下方没有可合并的记录。');
  return { fields, records };
}

/** 功能：按原始数值比较、按显示文字搜索，多条件同时满足；时间：2026-10-24；作者：lq。 */
export function filterMergeRecords(records: MergeRecord[], conditions: MergeCondition[], search = ''): MergeRecord[] {
  return records.filter((record) => (!search || Object.values(record.values).some((value) => value.toLocaleLowerCase().includes(search.toLocaleLowerCase()))) && conditions.every(({ field, op, value }) => {
    const text = record.values[field] ?? '';
    if (op === 'empty') return !text.trim();
    if (op === 'notEmpty') return !!text.trim();
    if (op === 'contains') return text.toLocaleLowerCase().includes(value.toLocaleLowerCase());
    if (op === 'eq') return text === value;
    if (op === 'ne') return text !== value;
    const raw = record.raw[field];
    if (raw === null || raw === undefined || String(raw).trim() === '' || !value.trim()) return false;
    const a = Number(raw), b = Number(value);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
    return op === 'gt' ? a > b : op === 'gte' ? a >= b : op === 'lt' ? a < b : a <= b;
  }));
}

/** 功能：递归收集正文、表格和页眉页脚字段名称；时间：2026-10-24；作者：lq。 */
export function collectMergeFields(doc: DocModel): string[] {
  const fields = new Set<string>();
  const visit = (blocks: DocBlock[]) => { mapParagraphs(blocks, (paragraph) => { paragraph.runs.forEach((run) => { if (run.mergeField) fields.add(run.mergeField); }); return paragraph; }); };
  visit(doc.blocks); visit(doc.header?.blocks ?? []); visit(doc.footer?.blocks ?? []);
  return [...fields];
}

/** 功能：逐记录核对字段缺列、必填空值及计算错误，不静默生成残缺报告；时间：2026-10-24；作者：lq。 */
export function mergeRecordIssues(doc: DocModel, data: MergeData, record: MergeRecord, allowEmpty = false): string[] {
  return [...record.errors, ...collectMergeFields(doc).flatMap((field) => !data.fields.includes(field) ? [`缺少数据列：${field}`] : !allowEmpty && !record.values[field]?.trim() ? [`字段为空：${field}`] : [])];
}

/** 功能：从模板生成独立文档副本，替换结构化域并保留样式、图片和分页；时间：2026-10-24；作者：lq。 */
export function mergeDocument(template: DocModel, record: MergeRecord): DocModel {
  const replace = (blocks: DocBlock[]) => mapParagraphs(blocks, (paragraph) => ({ ...paragraph, runs: paragraph.runs.flatMap<DocRun>((run) => {
    if (!run.mergeField) return { ...run, style: { ...run.style } };
    if (!Object.hasOwn(record.values, run.mergeField)) throw new Error(`缺少数据列：${run.mergeField}`);
    // 功能：CSV 单元格内换行和制表符转换成 Word 原生换行/制表符；时间：2026-10-24；作者：lq。
    return record.values[run.mergeField].replace(/\r\n?/g, '\n').split(/(\n|\t)/).map((text) => text === '\n' ? { lineBreak: true, style: { ...run.style } } : text === '\t' ? { tab: true, style: { ...run.style } } : { text, style: { ...run.style } });
  }) }));
  const doc = { ...template, blocks: replace(template.blocks), page: { ...template.page, margins: { ...template.page.margins }, pageNumberStart: template.page.pageNumberStart ?? 1 } };
  for (const kind of ['header', 'footer'] as const) if (template[kind]) {
    const original = template[kind]!;
    const hasFields = collectMergeFields({ ...template, blocks: original.blocks, header: undefined, footer: undefined }).length > 0;
    doc[kind] = { ...original, blocks: replace(original.blocks), edited: original.edited || hasFields };
  }
  return doc;
}

/** 功能：按记录生成安全且不重名的文件名，避免路径穿越与覆盖；时间：2026-10-24；作者：lq。 */
export function mergeFileName(pattern: string, record: MergeRecord, index: number): string {
  const value = pattern.replace(/\{([^{}]+)\}/g, (_, field: string) => field === '序号' ? String(index + 1).padStart(3, '0') : record.values[field] ?? '');
  const safe = value.replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/^\.+|[. ]+$/g, '').slice(0, 100).trim() || '报告';
  return `${String(index + 1).padStart(3, '0')}_${safe}`;
}
