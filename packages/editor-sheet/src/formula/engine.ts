import type {
  Cell,
  CellScalar,
  ComputedCell,
  SheetModel,
  WorkbookModel,
} from '@mini-office/document-model';
import {
  ERROR_CODES,
  FormulaError,
  compareText,
  error,
  flattenRange,
  intersectScalar,
  isError,
  isRange,
  unwrap,
  type EvalValue,
  type RangeStats,
  type RangeValue,
} from './types';
import { parseFormula, type Expr } from './parser';
import { getFunction, type FunctionContext } from './functions';
import { formatCellValue } from '../format';

const MAX_RANGE_CELLS = 300_000;

/** 单列前缀累计统计，用于把 MAX($A$4:An)+1 这类公式从 O(n²) 降为 O(n)。 */
interface ColumnStats {
  /** 已累计到的最后一行（不含）。 */
  lastRow: number;
  count: number;
  sum: number;
  max: number;
  min: number;
  product: number;
  nonEmpty: number;
  error?: FormulaError;
  /** 首个错误所在行，-1 表示没有错误。 */
  errorRow: number;
}

/** 功能：缓存公式语法树，避免同一公式文本被重复解析；时间：2026-10-27；作者：lq。 */
const formulaAstCache = new Map<string, Expr>();

function cachedParseFormula(formula: string): Expr {
  const cached = formulaAstCache.get(formula);
  if (cached) return cached;
  const ast = parseFormula(formula);
  if (formulaAstCache.size >= 50_000) formulaAstCache.clear();
  formulaAstCache.set(formula, ast);
  return ast;
}

function isErrorString(value: string): boolean {
  return /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A|CIRC!)$/i.test(value);
}

function scalarOf(value: EvalValue): CellScalar {
  const scalar = unwrap(value);
  if (isRange(scalar)) return scalarOf(scalar);
  if (isError(scalar)) return scalar.code;
  return scalar;
}

export function evaluateFormula(
  formula: string,
  workbook: WorkbookModel,
  sheetName: string,
  row = 0,
  col = 0,
): EvalValue {
  const sheet = workbook.sheets.find((item) => item.name.toLowerCase() === sheetName.toLowerCase());
  if (!sheet) return error(ERROR_CODES.REF);
  const evaluator = new WorkbookEvaluator(workbook);
  return evaluator.evaluateFormulaAt(formula, sheet, row, col);
}

class WorkbookEvaluator {
  private readonly workbook: WorkbookModel;
  private readonly sheetById = new Map<string, SheetModel>();
  private readonly sheetByName = new Map<string, SheetModel>();
  /** 功能：按工作表缓存单元格求值结果并使用数值键，避免区域扫描时反复拼接字符串；时间：2026-10-27；作者：lq。 */
  private readonly cache = new Map<string, Map<number, EvalValue>>();
  /** 功能：缓存同形状区域对象，让查找类函数可以复用匹配结果；时间：2026-10-27；作者：lq。 */
  private readonly ranges = new Map<string, RangeValue>();
  /** 功能：缓存单列前缀统计，把 MAX(A1:An) 这类公式从 O(n²) 降为 O(n)；时间：2026-10-27；作者：lq。 */
  private readonly columnStats = new Map<string, ColumnStats>();
  private readonly stack = new Set<string>();
  private readonly nameStack = new Set<string>();
  private readonly now = new Date();
  private readonly date1904: boolean;

  constructor(workbook: WorkbookModel) {
    this.workbook = workbook;
    this.date1904 = workbook.date1904 === true;
    for (const sheet of workbook.sheets) {
      this.sheetById.set(sheet.id, sheet);
      this.sheetByName.set(sheet.name.toLowerCase(), sheet);
    }
  }

  private resolveSheet(name?: string): SheetModel | undefined {
    if (!name) return undefined;
    return this.sheetByName.get(name.toLowerCase());
  }

  private rawCellValue(sheet: SheetModel, row: number, col: number): EvalValue {
    if (row < 0 || col < 0 || row > 1_048_575 || col > 16_383) {
      return error(ERROR_CODES.REF);
    }
    return this.evaluateCellAt(sheet.id, row, col);
  }

  evaluateCell(sheetId: string, key: string): EvalValue {
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    return this.evaluateCellAt(sheetId, row, col);
  }

  private evaluateCellAt(sheetId: string, row: number, col: number): EvalValue {
    const numericKey = row * 16_384 + col;
    let sheetCache = this.cache.get(sheetId);
    if (!sheetCache) {
      sheetCache = new Map<number, EvalValue>();
      this.cache.set(sheetId, sheetCache);
    }
    const cached = sheetCache.get(numericKey);
    if (cached !== undefined) return cached;
    const stackKey = `${sheetId}!${row}:${col}`;
    if (this.stack.has(stackKey)) return error(ERROR_CODES.CIRC);
    const sheet = this.sheetById.get(sheetId);
    if (!sheet) return error(ERROR_CODES.REF);
    const cell = sheet.cells[`${row}:${col}`];
    if (!cell) {
      sheetCache.set(numericKey, null);
      return null;
    }
    this.stack.add(stackKey);
    let value: EvalValue;
    try {
      if (cell.formula) {
        const ast = cachedParseFormula(cell.formula);
        value = this.evaluateExpr(ast, sheet, row, col);
        if (isRange(value)) value = intersectScalar(value, row, col);
      } else {
        value = cell.value;
        if (typeof value === 'string' && isErrorString(value)) {
          value = new FormulaError(value.toUpperCase() as FormulaError['code']);
        }
      }
    } catch (caught) {
      value =
        caught instanceof FormulaError
          ? caught
          : new FormulaError(ERROR_CODES.VALUE, caught instanceof Error ? caught.message : String(caught));
    } finally {
      this.stack.delete(stackKey);
    }
    sheetCache.set(numericKey, value);
    return value;
  }

  evaluateFormulaAt(formula: string, sheet: SheetModel, row: number, col: number): EvalValue {
    try {
      return this.evaluateExpr(cachedParseFormula(formula), sheet, row, col);
    } catch (caught) {
      return caught instanceof FormulaError ? caught : error(ERROR_CODES.VALUE);
    }
  }

  private createRange(
    sheetName: string | undefined,
    r1: number,
    c1: number,
    r2: number,
    c2: number,
    currentSheet: SheetModel,
    wholeColumn: boolean,
    wholeRow: boolean,
  ): EvalValue {
    const sheet = sheetName ? this.resolveSheet(sheetName) : currentSheet;
    if (!sheet) return error(ERROR_CODES.REF);
    let top = Math.min(r1, r2);
    let bottom = Math.max(r1, r2);
    let left = Math.min(c1, c2);
    let right = Math.max(c1, c2);
    if (wholeColumn) {
      top = 0;
      bottom = Math.max(0, Math.min(bottom, sheet.dimension.rows - 1));
    }
    if (wholeRow) {
      left = 0;
      right = Math.max(0, Math.min(right, sheet.dimension.cols - 1));
    }
    top = Math.max(0, top);
    left = Math.max(0, left);
    bottom = Math.min(1_048_575, bottom);
    right = Math.min(16_383, right);
    const rows = bottom - top + 1;
    const cols = right - left + 1;
    if (rows <= 0 || cols <= 0) return error(ERROR_CODES.REF);
    if (rows * cols > MAX_RANGE_CELLS) return error(ERROR_CODES.NUM);
    const cacheKey = `${sheet.id}:${top}:${left}:${bottom}:${right}`;
    const cached = this.ranges.get(cacheKey);
    if (cached) {
      // 功能：同一区域可能由整列/整行引用创建，需补齐来源标记以支持隐式交集；时间：2026-10-28；作者：lq。
      if ((wholeColumn && !cached.wholeColumn) || (wholeRow && !cached.wholeRow)) {
        const marked: RangeValue = {
          kind: 'range',
          sheetName: cached.sheetName,
          r1: cached.r1,
          c1: cached.c1,
          r2: cached.r2,
          c2: cached.c2,
          rows: cached.rows,
          cols: cached.cols,
          wholeColumn: wholeColumn || cached.wholeColumn,
          wholeRow: wholeRow || cached.wholeRow,
          get cells(): EvalValue[][] {
            return cached.cells;
          },
          cellAt: cached.cellAt,
        };
        this.ranges.set(cacheKey, marked);
        return marked;
      }
      return cached;
    }
    const evaluator = this;
    let materialized: EvalValue[][] | undefined;
    const range: RangeValue = {
      kind: 'range',
      sheetName: sheet.name,
      r1: top,
      c1: left,
      r2: bottom,
      c2: right,
      rows,
      cols,
      wholeColumn: wholeColumn || undefined,
      wholeRow: wholeRow || undefined,
      // 功能：区域默认惰性求值，只有旧接口直接读取 cells 时才整体展开；时间：2026-10-27；作者：lq。
      get cells(): EvalValue[][] {
        if (!materialized) {
          const grid: EvalValue[][] = [];
          for (let row = top; row <= bottom; row += 1) {
            const line: EvalValue[] = [];
            for (let col = left; col <= right; col += 1) line.push(evaluator.rawCellValue(sheet, row, col));
            grid.push(line);
          }
          materialized = grid;
        }
        return materialized;
      },
      cellAt: (row: number, col: number) => evaluator.rawCellValue(sheet, row, col),
    };
    this.ranges.set(cacheKey, range);
    return range;
  }

  /** 功能：统计区域内的数值单元格，单列区域走前缀缓存，其余区域直接扫描；时间：2026-10-27；作者：lq。 */
  private rangeStats(range: RangeValue): RangeStats {
    const sheet = this.resolveSheet(range.sheetName);
    if (!sheet) {
      return { count: 0, sum: 0, max: 0, min: 0, product: 1, nonEmpty: 0, error: error(ERROR_CODES.REF) };
    }
    if (range.c1 === range.c2) return this.columnStatsFor(sheet, range.c1, range.r1, range.r2);
    return this.scanRangeStats(sheet, range.r1, range.c1, range.r2, range.c2);
  }

  /** 功能：按列前缀累计统计，MAX($A$4:An)+1 这类公式只扫描一次整列；时间：2026-10-27；作者：lq。 */
  private columnStatsFor(sheet: SheetModel, col: number, startRow: number, endRow: number): RangeStats {
    const cacheKey = `${sheet.id}:${col}:${startRow}`;
    let stats = this.columnStats.get(cacheKey);
    if (!stats) {
      stats = {
        lastRow: startRow - 1,
        count: 0,
        sum: 0,
        max: -Infinity,
        min: Infinity,
        product: 1,
        nonEmpty: 0,
        errorRow: -1,
      };
      this.columnStats.set(cacheKey, stats);
    }
    if (stats.error && endRow >= stats.errorRow) return this.toRangeStats(stats);
    if (stats.lastRow === endRow) return this.toRangeStats(stats);
    if (stats.lastRow > endRow) return this.scanRangeStats(sheet, startRow, col, endRow, col);
    while (stats.lastRow < endRow) {
      const row = stats.lastRow + 1;
      const value = this.rawCellValue(sheet, row, col);
      if (isError(value)) {
        if (stats.lastRow < row) {
          stats.error = value;
          stats.errorRow = row;
          stats.lastRow = row;
        }
        break;
      }
      // 嵌套求值可能已经推进过累计位置，此时跳过避免重复累计。
      if (stats.lastRow >= row) continue;
      this.accumulateStats(stats, value);
      stats.lastRow = row;
    }
    return this.toRangeStats(stats);
  }

  private scanRangeStats(sheet: SheetModel, r1: number, c1: number, r2: number, c2: number): RangeStats {
    const stats: ColumnStats = {
      lastRow: r2,
      count: 0,
      sum: 0,
      max: -Infinity,
      min: Infinity,
      product: 1,
      nonEmpty: 0,
      errorRow: -1,
    };
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) {
        const value = this.rawCellValue(sheet, row, col);
        if (isError(value)) {
          stats.error = value;
          stats.errorRow = row;
          return this.toRangeStats(stats);
        }
        this.accumulateStats(stats, value);
      }
    }
    return this.toRangeStats(stats);
  }

  private accumulateStats(stats: ColumnStats, value: EvalValue): void {
    if (value !== null && value !== undefined && value !== '') stats.nonEmpty += 1;
    if (typeof value !== 'number') return;
    stats.count += 1;
    stats.sum += value;
    if (value > stats.max) stats.max = value;
    if (value < stats.min) stats.min = value;
    stats.product *= value;
  }

  private toRangeStats(stats: ColumnStats): RangeStats {
    return {
      count: stats.count,
      sum: stats.sum,
      max: stats.count ? stats.max : 0,
      min: stats.count ? stats.min : 0,
      product: stats.count ? stats.product : 1,
      nonEmpty: stats.nonEmpty,
      error: stats.error,
    };
  }

  private evaluateExpr(expr: Expr, sheet: SheetModel, row: number, col: number): EvalValue {
    switch (expr.type) {
      case 'number':
        return expr.value;
      case 'string':
        return expr.value;
      case 'boolean':
        return expr.value;
      case 'error':
        return new FormulaError(expr.value);
      case 'empty':
        return null;
      case 'ref': {
        const target = expr.sheet ? this.resolveSheet(expr.sheet) : sheet;
        if (!target) return error(ERROR_CODES.REF);
        return this.rawCellValue(target, expr.row, expr.col);
      }
      case 'range':
        return this.createRange(
          expr.sheet,
          expr.r1,
          expr.c1,
          expr.r2,
          expr.c2,
          sheet,
          !!expr.wholeColumn,
          !!expr.wholeRow,
        );
      case 'name': {
        const defined = this.workbook.definedNames?.[expr.name] ?? this.workbook.definedNames?.[expr.name.toUpperCase()];
        if (!defined) return error(ERROR_CODES.NAME);
        if (this.nameStack.has(expr.name)) return error(ERROR_CODES.CIRC);
        this.nameStack.add(expr.name);
        try {
          const definition = defined.startsWith('=') ? defined : `=${defined}`;
          return this.evaluateExpr(parseFormula(definition), sheet, row, col);
        } catch {
          return error(ERROR_CODES.NAME);
        } finally {
          this.nameStack.delete(expr.name);
        }
      }
      case 'unary': {
        const value = this.toNumber(this.evaluateExpr(expr.expr, sheet, row, col));
        if (isError(value)) return value;
        return expr.op === '-' ? -value : value;
      }
      case 'percent': {
        const value = this.toNumber(this.evaluateExpr(expr.expr, sheet, row, col));
        if (isError(value)) return value;
        return value / 100;
      }
      case 'binary':
        return this.evaluateBinary(expr, sheet, row, col);
      case 'func': {
        const definition = getFunction(expr.name);
        if (!definition) return error(ERROR_CODES.NAME);
        if (expr.args.length < definition.minArgs || expr.args.length > definition.maxArgs) {
          return error(ERROR_CODES.VALUE);
        }
        const context: FunctionContext = {
          workbook: this.workbook,
          sheet,
          sheetName: sheet.name,
          date1904: this.date1904,
          now: this.now,
          currentRow: row,
          currentCol: col,
          evalArg: (arg) => this.evaluateExpr(arg, sheet, row, col),
          createRange: (sheetName, r1, c1, r2, c2) => {
            const result = this.createRange(sheetName, r1, c1, r2, c2, sheet, false, false);
            if (isError(result)) throw result;
            return result as RangeValue;
          },
          rangeStats: (range) => this.rangeStats(range),
          resolveSheet: (name) => (name ? this.resolveSheet(name) : sheet),
        };
        try {
          const result = definition.call(expr.args, context);
          return isRange(result) ? unwrap(result) : result;
        } catch (caught) {
          if (caught instanceof FormulaError) return caught;
          return error(ERROR_CODES.VALUE);
        }
      }
      default:
        return error(ERROR_CODES.NAME);
    }
  }

  private evaluateBinary(
    expr: Extract<Expr, { type: 'binary' }>,
    sheet: SheetModel,
    row: number,
    col: number,
  ): EvalValue {
    const leftRaw = this.evaluateExpr(expr.left, sheet, row, col);
    const rightRaw = this.evaluateExpr(expr.right, sheet, row, col);
    const left = intersectScalar(leftRaw, row, col);
    const right = intersectScalar(rightRaw, row, col);
    if (isError(left)) return left;
    if (isError(right)) return right;
    if (expr.op === '&') {
      return `${this.toText(left)}${this.toText(right)}`;
    }
    if (['=', '<>', '<', '>', '<=', '>='].includes(expr.op)) {
      return this.compare(left, right, expr.op);
    }
    const a = this.toNumber(left);
    const b = this.toNumber(right);
    if (isError(a)) return a;
    if (isError(b)) return b;
    switch (expr.op) {
      case '+':
        return a + b;
      case '-':
        return a - b;
      case '*':
        return a * b;
      case '/':
        return b === 0 ? error(ERROR_CODES.DIV_ZERO) : a / b;
      case '^': {
        const result = a ** b;
        return Number.isFinite(result) ? result : error(ERROR_CODES.NUM);
      }
      default:
        return error(ERROR_CODES.VALUE);
    }
  }

  private compare(left: EvalValue, right: EvalValue, op: string): boolean {
    if (isError(left)) return false;
    if (isError(right)) return false;
    let comparison = 0;
    if (left === null && right === null) comparison = 0;
    else if (left === null) comparison = typeof right === 'string' ? -1 : -1;
    else if (right === null) comparison = 1;
    else if (typeof left === 'number' || typeof right === 'number' || typeof left === 'boolean' || typeof right === 'boolean') {
      const a = this.toNumber(left);
      const b = this.toNumber(right);
      if (isError(a) || isError(b)) {
        comparison = compareText(this.toText(left), this.toText(right));
      } else {
        comparison = a - b;
      }
    } else {
      comparison = compareText(this.toText(left), this.toText(right));
    }
    switch (op) {
      case '=':
        return comparison === 0;
      case '<>':
        return comparison !== 0;
      case '<':
        return comparison < 0;
      case '>':
        return comparison > 0;
      case '<=':
        return comparison <= 0;
      case '>=':
        return comparison >= 0;
      default:
        return false;
    }
  }

  private toNumber(value: EvalValue): number | FormulaError {
    const scalar = unwrap(value);
    if (isError(scalar)) return scalar;
    if (scalar === null || scalar === undefined) return 0;
    if (typeof scalar === 'number') return Number.isFinite(scalar) ? scalar : error(ERROR_CODES.NUM);
    if (typeof scalar === 'boolean') return scalar ? 1 : 0;
    if (typeof scalar === 'string') {
      const text = scalar.trim();
      if (!text) return 0;
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/.test(text)) {
        return text.endsWith('%') ? Number(text.slice(0, -1)) / 100 : Number(text);
      }
      return error(ERROR_CODES.VALUE);
    }
    return error(ERROR_CODES.VALUE);
  }

  private toText(value: EvalValue): string {
    const scalar = unwrap(value);
    if (isError(scalar)) return scalar.code;
    if (scalar === null || scalar === undefined) return '';
    if (typeof scalar === 'boolean') return scalar ? 'TRUE' : 'FALSE';
    if (typeof scalar === 'number') return formatCellValue(scalar, undefined, this.date1904);
    return String(scalar);
  }

  private createComputedCell(cell: Cell, value: EvalValue): ComputedCell {
    const style = cell.styleId ? this.workbook.styles[cell.styleId] : undefined;
    const scalar = scalarOf(value);
    return {
      value: scalar,
      display: isError(value) ? value.code : formatCellValue(scalar, style, this.date1904),
      error: isError(value) ? value.code : undefined,
    };
  }

  calculate(formulaCellsOnly = false): WorkbookModel {
    const computed: Record<string, Record<string, ComputedCell>> = {};
    for (const sheet of this.workbook.sheets) {
      const sheetComputed: Record<string, ComputedCell> = {};
      const keys = Object.keys(sheet.cells);
      for (const key of keys) {
        const cell = sheet.cells[key];
        // 功能：邮件合并等只需要公式结果时，不为普通单元格重复生成显示缓存；时间：2026-10-27；作者：lq。
        if (formulaCellsOnly && !cell.formula) continue;
        const value = this.evaluateCell(sheet.id, key);
        sheetComputed[key] = this.createComputedCell(cell, value);
      }
      computed[sheet.id] = sheetComputed;
    }
    return { ...this.workbook, computed };
  }

  /** 功能：分批计算公式并让出主线程，避免大工作簿导入时界面长期停在读取中；时间：2026-10-27；作者：lq。 */
  async calculateAsync(
    formulaCellsOnly = false,
    onProgress?: (completed: number, total: number) => void,
  ): Promise<WorkbookModel> {
    const computed: Record<string, Record<string, ComputedCell>> = {};
    let total = 0;
    for (const sheet of this.workbook.sheets) {
      for (const key of Object.keys(sheet.cells)) {
        if (!formulaCellsOnly || sheet.cells[key].formula) total += 1;
      }
    }
    let completed = 0;
    onProgress?.(0, total);
    if (total > 0) await new Promise<void>((resolve) => setTimeout(resolve, 0));
    for (const sheet of this.workbook.sheets) {
      const sheetComputed: Record<string, ComputedCell> = {};
      for (const key of Object.keys(sheet.cells)) {
        const cell = sheet.cells[key];
        if (formulaCellsOnly && !cell.formula) continue;
        const value = this.evaluateCell(sheet.id, key);
        sheetComputed[key] = this.createComputedCell(cell, value);
        completed += 1;
        if (completed % 200 === 0) {
          onProgress?.(completed, total);
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
        }
      }
      computed[sheet.id] = sheetComputed;
    }
    onProgress?.(completed, total);
    return { ...this.workbook, computed };
  }
}

export function recalculateWorkbook(
  workbook: WorkbookModel,
  options?: { formulaCellsOnly?: boolean },
): WorkbookModel {
  return new WorkbookEvaluator(workbook).calculate(options?.formulaCellsOnly === true);
}

/** 功能：异步分批执行与同步公式引擎相同的计算规则，并报告计算进度；时间：2026-10-27；作者：lq。 */
export async function recalculateWorkbookAsync(
  workbook: WorkbookModel,
  options?: {
    formulaCellsOnly?: boolean;
    onProgress?: (completed: number, total: number) => void;
  },
): Promise<WorkbookModel> {
  return new WorkbookEvaluator(workbook).calculateAsync(
    options?.formulaCellsOnly === true,
    options?.onProgress,
  );
}

export function evaluateCellValue(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  col: number,
): EvalValue {
  const evaluator = new WorkbookEvaluator(workbook);
  return evaluator.evaluateCell(sheetId, `${row}:${col}`);
}

export function formulaValueToCell(value: EvalValue): CellScalar {
  return scalarOf(value);
}

export function collectFormulaCells(workbook: WorkbookModel): Array<{ sheetId: string; key: string; formula: string }> {
  const result: Array<{ sheetId: string; key: string; formula: string }> = [];
  for (const sheet of workbook.sheets) {
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (cell.formula) result.push({ sheetId: sheet.id, key, formula: cell.formula });
    }
  }
  return result;
}

export function isFormulaCell(cell: Cell | undefined): boolean {
  return !!cell?.formula;
}

export { flattenRange, isError, isRange, unwrap };
