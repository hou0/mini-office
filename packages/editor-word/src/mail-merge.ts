import { createBlankWorkbook, type CellScalar, type DocBlock, type DocMergeCondition, type DocModel, type DocRun, type DocTable, type DocTableRow, type MergeFieldFormat, type WorkbookModel } from '@mini-office/document-model';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { formatCellValue, parseCsvRows, parseWorkbookPackage, recalculateWorkbookAsync } from '@mini-office/editor-sheet';
import { mapParagraphs } from './doc-utils';
import { uid } from '@mini-office/shared';

/** 功能：合并记录保留源行号、显示值及筛选原值，避免编号/金额失真；时间：2026-10-24；作者：lq。 */
export interface MergeRecord { rowNumber: number; values: Record<string, string>; raw: Record<string, CellScalar>; errors: string[] }
export interface MergeData { fields: string[]; records: MergeRecord[] }
export interface MergeCondition extends DocMergeCondition {}
export const MAX_MERGE_RECORDS = 2000;
export const MAX_MERGE_BATCH = 200;

export type MergeReadStage = 'reading' | 'parsing' | 'calculating' | 'complete';
export type MergeReadProgress = (stage: MergeReadStage, completed: number, total: number) => void;

/** 功能：读取 XLSX/XLS（OOXML 内容或 Excel HTML 导出）及 UTF-8/GB18030 CSV；时间：2026-10-27；作者：lq。 */
export async function readMergeWorkbook(
  bytes: Uint8Array,
  name: string,
  onProgress?: MergeReadProgress,
): Promise<WorkbookModel> {
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
    onProgress?.('complete', 1, 1);
    return workbook;
  }
  if (!/\.(xlsx|xls|xlsm|xltx)$/i.test(name)) throw new Error('请选择 XLS、XLSX 或 CSV 文件。');
  onProgress?.('reading', 0, 1);
  try {
    const pkg = await loadOoxmlPackage(bytes);
    try {
      // 功能：数据源走轻量模式，跳过图片与筛选还原等邮件合并不需要的解析；时间：2026-10-30；作者：lq。
      const workbook = await parseWorkbookPackage(pkg, name, {
        formulaCellsOnly: true,
        calculate: false,
        lightweight: true,
        onProgress: (stage, completed, total) => onProgress?.(stage, completed, total),
      });
      // 功能：公式重算改为分批异步执行（每 200 个单元格让出主线程并回报进度），
      // 避免大工作簿导入时界面长时间停在“正在读取…”；时间：2026-10-29；作者：lq。
      onProgress?.('calculating', 0, 1);
      try {
        return await recalculateWorkbookAsync(workbook, {
          formulaCellsOnly: true,
          onProgress: (completed, total) => onProgress?.('calculating', completed, total),
        });
      } catch {
        // 功能：公式引擎遇到不支持的函数或异常时退回未计算的工作簿（保留缓存值/公式），不阻断导入；时间：2026-10-29；作者：lq。
        return workbook;
      }
    } finally { pkg.dispose(); }
  } catch {
    // 功能：对 Excel 常见的“网页格式 .xls”导出提供离线导入；时间：2026-10-27；作者：lq。
    const text = new TextDecoder('utf-8').decode(bytes);
    if (!/<(?:html|table|tr|td)\b/i.test(text)) throw new Error('无法读取该 XLS 文件。请在 Excel 中另存为 XLSX 后再导入。');
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.name = name.replace(/\.(xls|xlsx|xlsm|xltx)$/i, '') || 'Sheet1';
    const rows = [...text.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)].map((row) =>
      [...row[1].matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi)].map((cell) => cell[1].replace(/<[^>]+>/g, '').replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').trim()),
    );
    rows.forEach((row, r) => row.forEach((value, c) => { if (value !== '') sheet.cells[`${r}:${c}`] = { value }; }));
    sheet.dimension = { rows: rows.length, cols: rows.reduce((max, row) => Math.max(max, row.length), 0) };
    onProgress?.('complete', 1, 1);
    return workbook;
  }
}

/** 功能：把完整导入过程显示为明确阶段，避免大文件读取时界面只有笼统的加载提示；时间：2026-10-27；作者：lq。 */
export function mergeReadProgressText(stage: MergeReadStage, completed: number, total: number): string {
  if (stage === 'reading') return '正在读取并解压 XLSX…';
  if (stage === 'parsing') return `正在解析工作表 ${completed} / ${total}…`;
  if (stage === 'calculating') return total ? `正在计算公式 ${completed} / ${total}…` : '正在计算公式…';
  return '数据读取完成';
}

/** 功能：直接在调用线程解析 XLSX/CSV，不再使用 Web Worker。
 * Tauri/打包 WebView 及部分 Vite 配置下 new Worker 会静默失败（脚本不加载或消息不回传），
 * 导致 Promise 永久 pending、界面一直“读取中”；改为直接主线程解析可彻底根治；时间：2026-10-29；作者：lq。 */
export async function readMergeWorkbookOffThread(
  bytes: Uint8Array,
  name: string,
  onProgress?: MergeReadProgress,
): Promise<WorkbookModel> {
  return readMergeWorkbook(bytes, name, onProgress);
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
      // 功能：公式重算失败时回退到文件内已缓存的结果（与 Excel/WPS 打开文件时的行为一致），
      // 避免把“本机引擎暂不支持该公式”误报成“公式结果无效”；时间：2026-10-30；作者：lq。
      const cached = cell?.value ?? null;
      const value = cell?.formula ? (computed?.error ? cached : (computed?.value ?? cached)) : cached;
      const field = fields[index];
      if ((cell?.formula && value === null) || (typeof value === 'string' && /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|N\/A|NUM!|NULL!|SPILL!)/.test(value))) errors.push(`${field}：公式结果无效，请在 Excel 中计算后保存`);
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

/** 功能：按原始数值比较、按显示文字搜索，多条件按「且 / 或」组合（对齐 Word 查询选项）；时间：2026-10-29；作者：lq。 */
export function filterMergeRecords(records: MergeRecord[], conditions: MergeCondition[], search = ''): MergeRecord[] {
  const match = (record: MergeRecord, { field, op, value }: MergeCondition): boolean => {
    const text = record.values[field] ?? '';
    if (op === 'empty') return !text.trim();
    if (op === 'notEmpty') return !!text.trim();
    if (op === 'contains') return text.toLocaleLowerCase().includes(value.toLocaleLowerCase());
    if (op === 'notContains') return !text.toLocaleLowerCase().includes(value.toLocaleLowerCase());
    if (op === 'eq') return text === value;
    if (op === 'ne') return text !== value;
    const raw = record.raw[field];
    if (raw === null || raw === undefined || String(raw).trim() === '' || !value.trim()) return false;
    const a = Number(raw), b = Number(value);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
    return op === 'gt' ? a > b : op === 'gte' ? a >= b : op === 'lt' ? a < b : a <= b;
  };
  return records.filter((record) => {
    if (search && !Object.values(record.values).some((value) => value.toLocaleLowerCase().includes(search.toLocaleLowerCase()))) return false;
    if (!conditions.length) return true;
    let result = match(record, conditions[0]);
    for (let i = 1; i < conditions.length; i++) {
      const next = match(record, conditions[i]);
      result = conditions[i].conjunction === 'or' ? result || next : result && next;
    }
    return result;
  });
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

function rowHasMergeField(row: DocTableRow, key: 'mergeNext' | 'mergeField' | 'mergeRecord' | 'mergeSeq'): boolean {
  return row.cells.some((cell) =>
    cell.blocks.some((block) => block.type === 'paragraph' && block.runs.some((run) => run[key])),
  );
}

function findConditionInRows(rows: DocTableRow[], key: 'mergeNextIf' | 'mergeSkipIf'): string | undefined {
  for (const row of rows) {
    for (const cell of row.cells) {
      for (const block of cell.blocks) {
        if (block.type !== 'paragraph') continue;
        const run = block.runs.find((r) => r[key]);
        if (run) return run[key];
      }
    }
  }
  return undefined;
}

function cloneRows(rows: DocTableRow[]): DocTableRow[] {
  return rows.map((row) => ({
    ...row,
    id: uid('r'),
    cells: row.cells.map((cell) => ({
      ...cell,
      id: uid('c'),
      blocks: cell.blocks.map((block) =>
        block.type === 'paragraph'
          ? { ...block, id: uid('p'), runs: block.runs.map((run) => ({ ...run, style: { ...run.style } })) }
          : block,
      ),
    })),
  }));
}

// 功能：深拷贝非 NEXT 块并分配独立 id，使同一块在合并后的多页中各自渲染、互不干扰；时间：2026-10-29；作者：lq。
function cloneBlock(block: DocBlock): DocBlock {
  if (block.type === 'table') {
    return { ...block, id: uid('tbl'), rows: cloneRows(block.rows) };
  }
  return { ...block, id: uid('blk') };
}

function splitValueRuns(value: string, run: DocRun): DocRun[] {
  return value.replace(/\r\n?/g, '\n').split(/(\n|\t)/).map((text): DocRun =>
    text === '\n' ? { lineBreak: true, style: { ...run.style } }
    : text === '\t' ? { tab: true, style: { ...run.style } }
    : { text, style: { ...run.style } });
}

function formatMergeValue(value: string, format?: MergeFieldFormat): string {
  if (value === '' || value === null || value === undefined) return value ?? '';
  let out = value;
  if (format?.num) {
    const n = Number(out);
    if (Number.isFinite(n)) {
      const decimals = format.num.includes('.') ? (format.num.split('.')[1]?.match(/\d/g)?.length ?? 0) : 0;
      out = n.toFixed(decimals);
      const [int, frac] = out.split('.');
      out = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (frac ? `.${frac}` : '');
    }
  }
  if (format?.case === 'upper') out = out.toUpperCase();
  if (format?.case === 'lower') out = out.toLowerCase();
  if (format?.case === 'caps') out = out.charAt(0).toUpperCase() + out.slice(1).toLowerCase();
  if (format?.case === 'firstcap') out = out.replace(/\b\w/g, (c) => c.toUpperCase());
  if (format?.before && out !== '') out = format.before + out;
  if (format?.after && out !== '') out = out + format.after;
  return out;
}

function evalMergeIf(condition: string, record: MergeRecord): boolean {
  const match = /^\s*([^\s<>]+)\s*(<>|=|>=|<=|>|<)\s*(?:"((?:[^"\\]|\\.)*)"|([^\s]+))\s*$/.exec(condition.trim());
  if (!match) return false;
  const field = match[1];
  const op = match[2];
  const value = (match[3] ?? match[4] ?? '').replace(/\\(["\\])/g, '$1');
  const text = record.values[field] ?? '';
  switch (op) {
    case '=': return text === value;
    case '<>': return text !== value;
    case '>': { const a = Number(text); const b = Number(value); return Number.isFinite(a) && Number.isFinite(b) && a > b; }
    case '>=': { const a = Number(text); const b = Number(value); return Number.isFinite(a) && Number.isFinite(b) && a >= b; }
    case '<': { const a = Number(text); const b = Number(value); return Number.isFinite(a) && Number.isFinite(b) && a < b; }
    case '<=': { const a = Number(text); const b = Number(value); return Number.isFinite(a) && Number.isFinite(b) && a <= b; }
    default: return false;
  }
}

function fillBlockRows(rows: DocTableRow[], record: MergeRecord, seqCounter: { n: number }, allowEmpty: boolean): void {
  for (const row of rows) {
    for (const cell of row.cells) {
      cell.blocks = mapParagraphs(cell.blocks, (paragraph) => ({
        ...paragraph,
        runs: paragraph.runs.flatMap<DocRun>((run): DocRun[] => {
          if (run.mergeField) {
            const raw = Object.hasOwn(record.values, run.mergeField) ? record.values[run.mergeField] : '';
            return splitValueRuns(formatMergeValue(raw, run.mergeFormat), run);
          }
          if (run.mergeRecord) return splitValueRuns(String(record.rowNumber), run);
          if (run.mergeSeq) {
            seqCounter.n += 1;
            return splitValueRuns(String(seqCounter.n), run);
          }
          if (run.mergeNext || run.mergeNextIf || run.mergeSkipIf) return [{ text: '', style: { ...run.style } }];
          return [{ ...run, style: { ...run.style } }];
        }),
      }));
    }
  }
}

function clearBlockMergeFields(rows: DocTableRow[]): void {
  for (const row of rows) {
    for (const cell of row.cells) {
      cell.blocks = mapParagraphs(cell.blocks, (paragraph) => ({
        ...paragraph,
        runs: paragraph.runs.flatMap<DocRun>((run): DocRun[] => {
          if (run.mergeField || run.mergeRecord || run.mergeSeq || run.mergeNext || run.mergeNextIf || run.mergeSkipIf) {
            return [{ text: '', style: { ...run.style } }];
          }
          return [{ ...run, style: { ...run.style } }];
        }),
      }));
    }
  }
}

/** 功能：规划 NEXT 模板的表头、记录块、尾部与记录分配，供分页输出复用；时间：2026-10-28；作者：lq。 */
function planTableRows(
  table: DocTable,
  records: MergeRecord[],
): {
  headerRows: DocTableRow[];
  blocks: DocTableRow[][];
  tailRows: DocTableRow[];
  assignment: Array<MergeRecord | undefined>;
} | undefined {
  const nextIndexes = table.rows
    .map((row, index) => (rowHasMergeField(row, 'mergeNext') ? index : -1))
    .filter((index) => index >= 0);
  if (nextIndexes.length === 0) return undefined;

  let lastMergeRow = -1;
  table.rows.forEach((row, index) => {
    if (rowHasMergeField(row, 'mergeField') || rowHasMergeField(row, 'mergeRecord') || rowHasMergeField(row, 'mergeSeq')) {
      lastMergeRow = index;
    }
  });
  if (lastMergeRow < 0) lastMergeRow = table.rows.length - 1;

  const headerRows = table.rows.slice(0, nextIndexes[0]);
  const tailRows = table.rows.slice(lastMergeRow + 1);
  const blocks: DocTableRow[][] = [];
  nextIndexes.forEach((start, index) => {
    const end = index + 1 < nextIndexes.length ? nextIndexes[index + 1] - 1 : lastMergeRow;
    if (start <= end) blocks.push(table.rows.slice(start, end + 1));
  });

  const pending: MergeRecord[] = [];
  for (const record of records) {
    const skipped = blocks.some((block) => {
      const condition = findConditionInRows(block, 'mergeSkipIf');
      return condition ? evalMergeIf(condition, record) : false;
    });
    if (!skipped) pending.push(record);
  }

  const assignment: Array<MergeRecord | undefined> = [];
  let recordIndex = 0;
  let blockIndex = 0;
  while (blockIndex < blocks.length) {
    const record = recordIndex < pending.length ? pending[recordIndex] : undefined;
    assignment.push(record);
    if (record) {
      const nextIf = findConditionInRows(blocks[blockIndex], 'mergeNextIf');
      if (nextIf && !evalMergeIf(nextIf, record)) {
        // NEXTIF 为假：下一块继续使用当前记录
      } else {
        recordIndex += 1;
      }
    } else {
      recordIndex += 1;
    }
    blockIndex += 1;
  }
  // 记录数超过块数：继续按块排布，每条记录占一块
  while (recordIndex < pending.length) {
    assignment.push(pending[recordIndex]);
    recordIndex += 1;
  }
  return { headerRows, blocks, tailRows, assignment };
}

export function mergeNextDocument(
  template: DocModel,
  records: MergeRecord[],
  options: { allowEmpty?: boolean } = {},
): DocModel {
  const allowEmpty = options.allowEmpty ?? true;
  const nextTables = template.blocks.filter(
    (block): block is DocTable =>
      block.type === 'table' && block.rows.some((row) => rowHasMergeField(row, 'mergeNext')),
  );
  if (nextTables.length === 0) return template;

  type TablePlan = { block: DocTable; plan: NonNullable<ReturnType<typeof planTableRows>> };
  const plans: TablePlan[] = [];
  for (const block of template.blocks) {
    if (block.type !== 'table' || !nextTables.includes(block)) continue;
    const plan = planTableRows(block, records);
    if (plan) plans.push({ block, plan });
  }
  if (!plans.length) return template;

  // 功能：校验所有 NEXT 表的记录块数一致，避免多张表记录窗口错位时静默丢记录；时间：2026-10-29；作者：lq。
  const firstBlockCount = plans[0].plan.blocks.length;
  if (plans.some((entry) => entry.plan.blocks.length !== firstBlockCount)) {
    throw new Error(
      `多张 NEXT 表的记录块数不一致（${plans.map((entry) => entry.plan.blocks.length).join('/')}），无法合并，请调整模板。`,
    );
  }

  // 功能：按页组装——同一页内所有 NEXT 表共用该页的记录窗口，每页输出一组表格，避免所有记录堆进一个超高表格；时间：2026-10-28；作者：lq。
  const perPage = Math.max(1, plans[0].plan.blocks.length);
  const pageCount = Math.max(1, ...plans.map((entry) => Math.ceil(entry.plan.assignment.length / perPage)));
  const seqCounter = { n: 0 };
  const result: DocBlock[] = [];
  for (let page = 0; page < pageCount; page += 1) {
    if (page > 0) result.push({ type: 'pageBreak', id: uid('pb') });
    for (const block of template.blocks) {
      const entry = plans.find((item) => item.block === block);
      if (!entry) {
        // 功能：非 NEXT 块（标题段落、静态表、页脚说明等）作为每页的页面框架重复出现，
        // 使表单标题/编号在每一页都保留；每页独立克隆 id，避免多页共用同一 id 触发重复键与 blockMap 折叠；时间：2026-10-29；作者：lq。
        result.push(cloneBlock(block));
        continue;
      }
      const window = entry.plan.assignment.slice(page * perPage, page * perPage + perPage);
      // 功能：每页的表头/表尾行也独立复制，避免多页共用同一批行与单元格 id，导致 React 键重复与跨页选区串页；时间：2026-10-28；作者：lq。
      const rows: DocTableRow[] = cloneRows(entry.plan.headerRows);
      for (let i = 0; i < perPage; i += 1) {
        const source = entry.plan.blocks[i] ?? entry.plan.blocks[entry.plan.blocks.length - 1];
        const blockRows = cloneRows(source);
        const record = window[i];
        if (record) fillBlockRows(blockRows, record, seqCounter, allowEmpty);
        else clearBlockMergeFields(blockRows);
        rows.push(...blockRows);
      }
      rows.push(...cloneRows(entry.plan.tailRows));
      // 功能：每页表格使用独立 id，React 按 key 区分多页表格，避免重复键导致子节点被丢弃；时间：2026-10-28；作者：lq。
      result.push({ ...entry.block, id: uid('tbl'), rows });
    }
  }
  return { ...template, blocks: result };
}

/** 检测模板是否包含 NEXT 记录边界域（仅扫描顶层表格行） */
export function hasNextMergeField(doc: DocModel): boolean {
  return doc.blocks.some(
    (block) => block.type === 'table' && block.rows.some((row) => rowHasMergeField(row, 'mergeNext')),
  );
}

/** 统计模板顶层表格中的 NEXT 块数，即每页可容纳的记录条数 */
export function countNextBlocks(doc: DocModel): number {
  let count = 0;
  for (const block of doc.blocks) {
    if (block.type !== 'table') continue;
    for (const row of block.rows) {
      if (rowHasMergeField(row, 'mergeNext')) count += 1;
    }
  }
  return count;
}

/** 功能：把页号范围换算成记录区间（1-based，含两端）；每页 pageSize 条记录，起止页颠倒时自动交换；时间：2026-10-30；作者：lq。 */
export function pageRangeToRecordRange(start: number, end: number, pageSize: number): [number, number] {
  const size = Math.max(1, Math.floor(pageSize) || 1);
  const first = Math.max(1, Math.floor(start) || 1);
  const last = Math.max(1, Math.floor(end) || 1);
  const low = Math.min(first, last);
  const high = Math.max(first, last);
  return [(low - 1) * size + 1, high * size];
}

/** 按记录范围从选中记录列表中选取子集；范围语义学 Office 邮件合并“打印记录”对话框 */
export type MergeRangeKind = 'all' | 'currentRecord' | 'recordRange';

/** 按记录序号选取记录：all=全部；currentRecord=当前预览记录；recordRange=从第 start 条到第 end 条（1-based，自动交换越界钳制） */
export function recordsByRange(
  records: MergeRecord[],
  range: MergeRangeKind,
  previewIndex: number,
  start = 1,
  end = 1,
): MergeRecord[] {
  if (range === 'all' || records.length === 0) return records;
  if (range === 'currentRecord') {
    const i = Math.min(Math.max(0, previewIndex), records.length - 1);
    return records.slice(i, i + 1);
  }
  const from = Math.max(1, Math.min(start, end));
  const to = Math.min(records.length, Math.max(start, end));
  if (from > records.length) return [];
  return records.slice(from - 1, to);
}
