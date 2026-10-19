/** 功能：结构编辑、工作表更名及排序时维护公式引用；时间：2026-10-18；作者：lq。 */
import type { WorkbookModel } from '@mini-office/document-model';
import { columnLabel } from './parser';
import { tokenize, type CellReferenceToken } from './tokenizer';

export type ReferenceEdit =
  | { kind: 'rows' | 'columns'; sheet: string; at: number; count: number; deleting?: boolean }
  | { kind: 'rename'; sheet: string; name: string }
  | { kind: 'delete-sheet'; sheet: string };

/** 功能：仅替换引用 token，保留字符串、运算优先级和原始函数文本；时间：2026-10-18；作者：lq。 */
function rewrite(formula: string, transform: (ref: CellReferenceToken) => CellReferenceToken | null): string {
  const offset = formula.startsWith('=') ? 1 : 0;
  const source = formula.slice(offset);
  let tokens;
  try {
    tokens = tokenize(source);
  } catch {
    throw new Error('公式包含暂不支持调整的引用语法，请先在 Excel 中处理，再进行结构操作。');
  }
  let output = formula;
  for (const token of tokens.reverse()) {
    if (token.type !== 'ref') continue;
    const next = transform(token.value);
    if (next === token.value) continue;
    const text = next ? referenceText(next) : '#REF!';
    output = output.slice(0, token.pos + offset) + text + output.slice(token.end + offset);
  }
  return output;
}

/** 功能：正确输出绝对引用、整行整列和带引号的工作表名；时间：2026-10-18；作者：lq。 */
function referenceText(ref: CellReferenceToken): string {
  const prefix = ref.sheet === undefined ? '' : `'${ref.sheet.replace(/'/g, "''")}'!`;
  const col = (n: number, absolute: boolean) => `${absolute ? '$' : ''}${columnLabel(n)}`;
  const row = (n: number, absolute: boolean) => `${absolute ? '$' : ''}${n + 1}`;
  if (ref.wholeColumn) return `${prefix}${col(ref.c1, ref.colAbs1)}:${col(ref.c2, ref.colAbs2)}`;
  if (ref.wholeRow) return `${prefix}${row(ref.r1, ref.rowAbs1)}:${row(ref.r2, ref.rowAbs2)}`;
  const first = `${col(ref.c1, ref.colAbs1)}${row(ref.r1, ref.rowAbs1)}`;
  const last = `${col(ref.c2, ref.colAbs2)}${row(ref.r2, ref.rowAbs2)}`;
  return prefix + (first === last ? first : `${first}:${last}`);
}

/** 功能：删除引用区间时缩短范围，全部删除时返回无效引用；时间：2026-10-18；作者：lq。 */
function editInterval(first: number, last: number, at: number, count: number, deleting: boolean): [number, number] | null {
  const low = Math.min(first, last);
  const high = Math.max(first, last);
  let start: number;
  let end: number;
  if (!deleting) {
    start = low >= at ? low + count : low;
    end = high >= at ? high + count : high;
  } else {
    const removedEnd = at + count - 1;
    if (low >= at && high <= removedEnd) return null;
    start = low < at ? low : low > removedEnd ? low - count : at;
    end = high > removedEnd ? high - count : high < at ? high : at - 1;
  }
  return first <= last ? [start, end] : [end, start];
}

/** 功能：结构位移同时调整绝对和相对引用，跨表只调整目标工作表；时间：2026-10-18；作者：lq。 */
export function editFormulaReferences(formula: string, currentSheet: string, edit: ReferenceEdit): string {
  return rewrite(formula, (ref) => {
    if ((ref.sheet ?? currentSheet).toLowerCase() !== edit.sheet.toLowerCase()) return ref;
    if (edit.kind === 'rename') return ref.sheet === undefined ? ref : { ...ref, sheet: edit.name };
    if (edit.kind === 'delete-sheet') return null;
    if ((edit.kind === 'rows' && ref.wholeColumn) || (edit.kind === 'columns' && ref.wholeRow)) return ref;
    const rows = edit.kind === 'rows';
    const interval = editInterval(rows ? ref.r1 : ref.c1, rows ? ref.r2 : ref.c2, edit.at, edit.count, !!edit.deleting);
    if (!interval || interval.some((n) => n < 0 || n > (rows ? 1_048_575 : 16_383))) return null;
    return rows ? { ...ref, r1: interval[0], r2: interval[1] } : { ...ref, c1: interval[0], c2: interval[1] };
  });
}

/** 功能：重写全工作簿公式和命名区域，并清除旧计算缓存；时间：2026-10-18；作者：lq。 */
export function editWorkbookReferences(workbook: WorkbookModel, edit: ReferenceEdit): WorkbookModel {
  return {
    ...workbook,
    computed: undefined,
    sheets: workbook.sheets.map((sheet) => ({
      ...sheet,
      cells: Object.fromEntries(Object.entries(sheet.cells).map(([key, cell]) => [key, cell.formula
        ? { ...cell, formula: editFormulaReferences(cell.formula, sheet.name, edit) }
        : cell])),
    })),
    definedNames: workbook.definedNames && Object.fromEntries(Object.entries(workbook.definedNames).map(([name, formula]) =>
      [name, editFormulaReferences(formula, workbook.sheets[0]?.name ?? '', edit)])),
  };
}

/** 功能：排序按移动行数平移相对引用，绝对引用保持原位，越界输出 REF；时间：2026-10-18；作者：lq。 */
export function moveSortedFormula(formula: string, deltaRow: number): string {
  if (!deltaRow) return formula;
  return rewrite(formula, (ref) => {
    if (ref.wholeColumn) return ref;
    const next = { ...ref, r1: ref.r1 + (ref.rowAbs1 ? 0 : deltaRow), r2: ref.r2 + (ref.rowAbs2 ? 0 : deltaRow) };
    return next.r1 < 0 || next.r2 < 0 || next.r1 > 1_048_575 || next.r2 > 1_048_575 ? null : next;
  });
}
