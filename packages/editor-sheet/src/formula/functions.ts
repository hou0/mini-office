import type { SheetModel, WorkbookModel } from '@mini-office/document-model';
import {
  ERROR_CODES,
  FormulaError,
  error,
  flattenRange,
  isError,
  isRange,
  unwrap,
  type EvalValue,
  type RangeValue,
} from './types';
import type { Expr } from './parser';
import {
  addMonths,
  endOfMonthSerial,
  makeDateSerial,
  nowSerial,
  serialToLocalParts,
  todaySerial,
} from '../dates';
import { formatNumeric } from '../format';

export interface FunctionContext {
  workbook: WorkbookModel;
  sheet: SheetModel;
  sheetName: string;
  date1904: boolean;
  now: Date;
  currentRow: number;
  currentCol: number;
  evalArg: (expr: Expr) => EvalValue;
  createRange: (
    sheetName: string | undefined,
    r1: number,
    c1: number,
    r2: number,
    c2: number,
  ) => RangeValue;
  resolveSheet: (name?: string) => SheetModel | undefined;
}

export interface FormulaFunction {
  minArgs: number;
  maxArgs: number;
  call: (args: Expr[], ctx: FunctionContext) => EvalValue;
}

const registry = new Map<string, FormulaFunction>();

export function registerFunction(name: string, fn: FormulaFunction): void {
  registry.set(name.toUpperCase(), fn);
}

export function getFunction(name: string): FormulaFunction | undefined {
  const normalized = name.toUpperCase().replace(/^_XLFN\./, '').replace(/^_XLWS\./, '');
  return registry.get(normalized);
}

// 功能：向公式智能提示提供实际已注册的函数，避免推荐尚未实现的函数；时间：2026-10-24；作者：lq。
export function listFunctions(): Array<{ name: string; minArgs: number; maxArgs: number }> {
  return Array.from(registry, ([name, definition]) => ({
    name,
    minArgs: definition.minArgs,
    maxArgs: definition.maxArgs,
  }));
}

function fn(
  minArgs: number,
  maxArgs: number,
  call: FormulaFunction['call'],
): FormulaFunction {
  return { minArgs, maxArgs, call };
}

function evaluateAll(args: Expr[], ctx: FunctionContext): EvalValue[] {
  return args.map((arg) => ctx.evalArg(arg));
}

function firstError(values: EvalValue[]): FormulaError | null {
  for (const value of values) {
    if (isError(value)) return value;
    if (isRange(value)) {
      for (const row of value.cells) {
        for (const cell of row) if (isError(cell)) return cell;
      }
    }
  }
  return null;
}

export function toNumberValue(value: EvalValue): number | FormulaError {
  if (isError(value)) return value;
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

export function toTextValue(value: EvalValue): string | FormulaError {
  if (isError(value)) return value;
  const scalar = unwrap(value);
  if (isError(scalar)) return scalar;
  if (scalar === null || scalar === undefined) return '';
  if (typeof scalar === 'boolean') return scalar ? 'TRUE' : 'FALSE';
  if (typeof scalar === 'number') return formatNumeric(scalar, 'General');
  return String(scalar);
}

export function toBooleanValue(value: EvalValue): boolean | FormulaError {
  if (isError(value)) return value;
  const scalar = unwrap(value);
  if (isError(scalar)) return scalar;
  if (scalar === null || scalar === undefined) return false;
  if (typeof scalar === 'boolean') return scalar;
  if (typeof scalar === 'number') return scalar !== 0;
  if (typeof scalar === 'string') {
    const upper = scalar.trim().toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    if (upper === '') return false;
    const numeric = Number(upper);
    if (Number.isFinite(numeric)) return numeric !== 0;
  }
  return error(ERROR_CODES.VALUE);
}

function numberArg(args: Expr[], index: number, ctx: FunctionContext, fallback?: number): number | FormulaError {
  if (index >= args.length) {
    if (fallback !== undefined) return fallback;
    return error(ERROR_CODES.VALUE);
  }
  return toNumberValue(ctx.evalArg(args[index]));
}

function textArg(args: Expr[], index: number, ctx: FunctionContext, fallback = ''): string | FormulaError {
  if (index >= args.length) return fallback;
  return toTextValue(ctx.evalArg(args[index]));
}

function booleanArg(args: Expr[], index: number, ctx: FunctionContext, fallback = false): boolean | FormulaError {
  if (index >= args.length) return fallback;
  return toBooleanValue(ctx.evalArg(args[index]));
}

interface NumericCollection {
  numbers: number[];
  error?: FormulaError;
}

function collectNumbers(values: EvalValue[], includeScalarText = false): NumericCollection {
  const numbers: number[] = [];
  for (const value of values) {
    if (isError(value)) return { numbers, error: value };
    const items = isRange(value) ? flattenRange(value) : [value];
    const fromRange = isRange(value);
    for (const item of items) {
      if (isError(item)) return { numbers, error: item };
      if (typeof item === 'number') numbers.push(item);
      else if (!fromRange && typeof item === 'boolean') numbers.push(item ? 1 : 0);
      else if (!fromRange && includeScalarText && typeof item === 'string' && item.trim() !== '') {
        const numeric = Number(item);
        if (Number.isFinite(numeric)) numbers.push(numeric);
      }
    }
  }
  return { numbers };
}

function roundHalfAway(value: number, digits: number): number {
  const factor = 10 ** digits;
  const scaled = value * factor;
  const rounded = Math.sign(scaled) * Math.round(Math.abs(scaled));
  return rounded / factor;
}

function aggregateNumbers(args: Expr[], ctx: FunctionContext): NumericCollection {
  return collectNumbers(evaluateAll(args, ctx));
}

function textOf(value: EvalValue): string {
  const scalar = unwrap(value);
  if (isError(scalar)) return scalar.code;
  if (scalar === null) return '';
  if (typeof scalar === 'boolean') return scalar ? 'TRUE' : 'FALSE';
  if (typeof scalar === 'number') return formatNumeric(scalar, 'General');
  return String(scalar);
}

function compareValues(a: EvalValue, b: EvalValue): number {
  const left = unwrap(a);
  const right = unwrap(b);
  if (isError(left)) return 1;
  if (isError(right)) return -1;
  if (left === null && right === null) return 0;
  const leftNumber = typeof left === 'number' ? left : left === null ? 0 : Number(left);
  const rightNumber = typeof right === 'number' ? right : right === null ? 0 : Number(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber) && typeof left !== 'boolean') {
    return leftNumber - rightNumber;
  }
  const leftText = textOf(left);
  const rightText = textOf(right);
  return leftText.localeCompare(rightText, 'zh-CN', { sensitivity: 'accent' });
}

function equalValues(a: EvalValue, b: EvalValue): boolean {
  const left = unwrap(a);
  const right = unwrap(b);
  if (left === null && right === null) return true;
  if (typeof left === 'string' && typeof right === 'string') {
    return left.localeCompare(right, 'zh-CN', { sensitivity: 'accent' }) === 0;
  }
  if (typeof left === 'boolean' || typeof right === 'boolean') {
    const lb = typeof left === 'boolean' ? left : String(left).toUpperCase() === 'TRUE';
    const rb = typeof right === 'boolean' ? right : String(right).toUpperCase() === 'TRUE';
    return lb === rb;
  }
  return compareValues(left, right) === 0;
}

interface Criteria {
  op: '=' | '<>' | '>' | '>=' | '<' | '<=';
  operand: EvalValue;
  wildcard?: RegExp;
}

function wildcardRegex(pattern: string): RegExp {
  let source = '^';
  for (const char of pattern) {
    if (char === '*') source += '.*';
    else if (char === '?') source += '.';
    else if (char === '~') continue;
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  source += '$';
  return new RegExp(source, 'i');
}

function parseCriteria(value: EvalValue): Criteria | FormulaError {
  if (isError(value)) return value;
  const scalar = unwrap(value);
  if (typeof scalar !== 'string') return { op: '=', operand: scalar };
  const match = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(scalar);
  const op = (match?.[1] as Criteria['op']) ?? '=';
  const operandText = match?.[2] ?? '';
  let operand: EvalValue = operandText;
  if (operandText === '') operand = null;
  else if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(operandText)) {
    operand = Number(operandText);
  } else if (/^TRUE$/i.test(operandText)) operand = true;
  else if (/^FALSE$/i.test(operandText)) operand = false;
  const wildcard =
    typeof operand === 'string' && (operand.includes('*') || operand.includes('?'))
      ? wildcardRegex(operand)
      : undefined;
  return { op, operand, wildcard };
}

function matchesCriteria(value: EvalValue, criteria: Criteria): boolean {
  const scalar = unwrap(value);
  if (criteria.wildcard) {
    if (criteria.op === '<>') return !criteria.wildcard.test(textOf(scalar));
    return criteria.wildcard.test(textOf(scalar));
  }
  if (criteria.op === '=') {
    if (criteria.operand === null) return scalar === null || scalar === '';
    if (typeof criteria.operand === 'string') {
      return textOf(scalar).localeCompare(criteria.operand, 'zh-CN', { sensitivity: 'accent' }) === 0;
    }
    return equalValues(scalar, criteria.operand);
  }
  if (criteria.op === '<>') {
    if (criteria.operand === null) return scalar !== null && scalar !== '';
    return !equalValues(scalar, criteria.operand);
  }
  const comparison = compareValues(scalar, criteria.operand);
  switch (criteria.op) {
    case '>':
      return comparison > 0;
    case '>=':
      return comparison >= 0;
    case '<':
      return comparison < 0;
    case '<=':
      return comparison <= 0;
    default:
      return false;
  }
}

function filteredValues(
  criteriaRange: RangeValue,
  criteriaValue: EvalValue,
): { values: EvalValue[]; matches: boolean[]; error?: FormulaError } {
  const criteria = parseCriteria(criteriaValue);
  if (criteria instanceof FormulaError) return { values: [], matches: [], error: criteria };
  const values = flattenRange(criteriaRange);
  return { values, matches: values.map((value) => matchesCriteria(value, criteria)) };
}

// ---- Math & statistics -------------------------------------------------

registerFunction('SUM', fn(0, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return collection.numbers.reduce((sum, value) => sum + value, 0);
}));

registerFunction('AVERAGE', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  if (collection.numbers.length === 0) return error(ERROR_CODES.DIV_ZERO);
  return collection.numbers.reduce((sum, value) => sum + value, 0) / collection.numbers.length;
}));

registerFunction('COUNT', fn(0, 255, (args, ctx) => {
  let count = 0;
  for (const value of evaluateAll(args, ctx)) {
    if (isError(value)) continue;
    const items = isRange(value) ? flattenRange(value) : [value];
    for (const item of items) if (typeof item === 'number') count += 1;
  }
  return count;
}));

registerFunction('COUNTA', fn(0, 255, (args, ctx) => {
  let count = 0;
  for (const value of evaluateAll(args, ctx)) {
    if (isError(value)) {
      count += 1;
      continue;
    }
    const items = isRange(value) ? flattenRange(value) : [value];
    for (const item of items) {
      if (item !== null && item !== undefined && item !== '') count += 1;
    }
  }
  return count;
}));

registerFunction('MAX', fn(0, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return collection.numbers.length ? Math.max(...collection.numbers) : 0;
}));

registerFunction('MIN', fn(0, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return collection.numbers.length ? Math.min(...collection.numbers) : 0;
}));

registerFunction('PRODUCT', fn(0, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  if (!collection.numbers.length) return 0;
  return collection.numbers.reduce((product, value) => product * value, 1);
}));

registerFunction('MEDIAN', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  if (!collection.numbers.length) return error(ERROR_CODES.NUM);
  const sorted = [...collection.numbers].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}));

function variance(numbers: number[], sample: boolean): number {
  if (numbers.length === 0 || (sample && numbers.length === 1)) return NaN;
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const divisor = sample ? numbers.length - 1 : numbers.length;
  return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / divisor;
}

registerFunction('STDEV', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  const result = Math.sqrt(variance(collection.numbers, true));
  return Number.isNaN(result) ? error(ERROR_CODES.DIV_ZERO) : result;
}));
registerFunction('STDEV.S', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  const result = Math.sqrt(variance(collection.numbers, true));
  return Number.isNaN(result) ? error(ERROR_CODES.DIV_ZERO) : result;
}));
registerFunction('STDEVP', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return Math.sqrt(variance(collection.numbers, false));
}));
registerFunction('STDEV.P', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return Math.sqrt(variance(collection.numbers, false));
}));
registerFunction('VAR', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return variance(collection.numbers, true);
}));
registerFunction('VARP', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  return variance(collection.numbers, false);
}));

registerFunction('ROUND', fn(2, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const digits = numberArg(args, 1, ctx);
  if (isError(value)) return value;
  if (isError(digits)) return digits;
  return roundHalfAway(value, Math.trunc(digits));
}));

registerFunction('ROUNDUP', fn(2, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const digits = numberArg(args, 1, ctx);
  if (isError(value)) return value;
  if (isError(digits)) return digits;
  const factor = 10 ** Math.trunc(digits);
  return (Math.sign(value) * Math.ceil(Math.abs(value) * factor)) / factor;
}));

registerFunction('ROUNDDOWN', fn(2, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const digits = numberArg(args, 1, ctx);
  if (isError(value)) return value;
  if (isError(digits)) return digits;
  const factor = 10 ** Math.trunc(digits);
  return (Math.sign(value) * Math.floor(Math.abs(value) * factor)) / factor;
}));

registerFunction('INT', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.floor(value);
}));
registerFunction('TRUNC', fn(1, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const digits = numberArg(args, 1, ctx, 0);
  if (isError(value)) return value;
  if (isError(digits)) return digits;
  const factor = 10 ** Math.trunc(digits);
  return Math.trunc(value * factor) / factor;
}));
registerFunction('ABS', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.abs(value);
}));
registerFunction('SIGN', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.sign(value);
}));
registerFunction('MOD', fn(2, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const divisor = numberArg(args, 1, ctx);
  if (isError(value)) return value;
  if (isError(divisor)) return divisor;
  if (divisor === 0) return error(ERROR_CODES.DIV_ZERO);
  return value - divisor * Math.floor(value / divisor);
}));
registerFunction('POWER', fn(2, 2, (args, ctx) => {
  const base = numberArg(args, 0, ctx);
  const exponent = numberArg(args, 1, ctx);
  if (isError(base)) return base;
  if (isError(exponent)) return exponent;
  const result = base ** exponent;
  return Number.isFinite(result) ? result : error(ERROR_CODES.NUM);
}));
registerFunction('SQRT', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  if (isError(value)) return value;
  return value < 0 ? error(ERROR_CODES.NUM) : Math.sqrt(value);
}));
registerFunction('EXP', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.exp(value);
}));
registerFunction('LN', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  if (isError(value)) return value;
  return value <= 0 ? error(ERROR_CODES.NUM) : Math.log(value);
}));
registerFunction('LOG', fn(1, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const base = numberArg(args, 1, ctx, 10);
  if (isError(value)) return value;
  if (isError(base)) return base;
  if (value <= 0 || base <= 0) return error(ERROR_CODES.NUM);
  return Math.log(value) / Math.log(base);
}));
registerFunction('LOG10', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  if (isError(value)) return value;
  return value <= 0 ? error(ERROR_CODES.NUM) : Math.log10(value);
}));
registerFunction('FLOOR', fn(1, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const significance = numberArg(args, 1, ctx, 1);
  if (isError(value)) return value;
  if (isError(significance)) return significance;
  if (significance === 0) return error(ERROR_CODES.DIV_ZERO);
  return Math.floor(value / significance) * significance;
}));
registerFunction('CEILING', fn(1, 2, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  const significance = numberArg(args, 1, ctx, 1);
  if (isError(value)) return value;
  if (isError(significance)) return significance;
  if (significance === 0) return error(ERROR_CODES.DIV_ZERO);
  return Math.ceil(value / significance) * significance;
}));
registerFunction('PI', fn(0, 0, () => Math.PI));
registerFunction('RAND', fn(0, 0, () => Math.random()));
registerFunction('RANDBETWEEN', fn(2, 2, (args, ctx) => {
  const low = numberArg(args, 0, ctx);
  const high = numberArg(args, 1, ctx);
  if (isError(low)) return low;
  if (isError(high)) return high;
  return Math.floor(Math.random() * (Math.floor(high) - Math.ceil(low) + 1)) + Math.ceil(low);
}));
registerFunction('GCD', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  const gcd = (a: number, b: number): number => (b === 0 ? Math.abs(a) : gcd(b, a % b));
  return collection.numbers.map(Math.trunc).reduce((result, value) => gcd(result, value), 0);
}));
registerFunction('LCM', fn(1, 255, (args, ctx) => {
  const collection = aggregateNumbers(args, ctx);
  if (collection.error) return collection.error;
  const gcd = (a: number, b: number): number => (b === 0 ? Math.abs(a) : gcd(b, a % b));
  return collection.numbers
    .map(Math.trunc)
    .reduce((result, value) => (result === 0 || value === 0 ? 0 : Math.abs(result * value) / gcd(result, value)), 1);
}));
registerFunction('FACT', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  if (isError(value)) return value;
  let result = 1;
  for (let index = 2; index <= Math.floor(value); index += 1) result *= index;
  return result;
}));

// ---- Logical -----------------------------------------------------------

registerFunction('IF', fn(2, 3, (args, ctx) => {
  const condition = toBooleanValue(ctx.evalArg(args[0]));
  if (isError(condition)) return condition;
  if (condition) return ctx.evalArg(args[1]);
  return args[2] ? ctx.evalArg(args[2]) : false;
}));

registerFunction('IFERROR', fn(2, 2, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  return isError(value) ? ctx.evalArg(args[1]) : value;
}));

registerFunction('IFNA', fn(2, 2, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  return isError(value) && value.code === ERROR_CODES.NA ? ctx.evalArg(args[1]) : value;
}));

registerFunction('IFS', fn(2, 255, (args, ctx) => {
  for (let index = 0; index + 1 < args.length; index += 2) {
    const condition = toBooleanValue(ctx.evalArg(args[index]));
    if (isError(condition)) return condition;
    if (condition) return ctx.evalArg(args[index + 1]);
  }
  return error(ERROR_CODES.NA);
}));

registerFunction('SWITCH', fn(3, 255, (args, ctx) => {
  const target = ctx.evalArg(args[0]);
  for (let index = 1; index + 1 < args.length; index += 2) {
    const candidate = ctx.evalArg(args[index]);
    if (equalValues(target, candidate)) return ctx.evalArg(args[index + 1]);
  }
  if (args.length % 2 === 0) return ctx.evalArg(args.at(-1)!);
  return error(ERROR_CODES.NA);
}));

registerFunction('AND', fn(1, 255, (args, ctx) => {
  let result = true;
  for (const value of evaluateAll(args, ctx)) {
    if (isError(value)) return value;
    for (const item of isRange(value) ? flattenRange(value) : [value]) {
      if (item === null || item === '') continue;
      const bool = toBooleanValue(item);
      if (isError(bool)) return bool;
      result = result && bool;
    }
  }
  return result;
}));

registerFunction('OR', fn(1, 255, (args, ctx) => {
  let result = false;
  for (const value of evaluateAll(args, ctx)) {
    if (isError(value)) return value;
    for (const item of isRange(value) ? flattenRange(value) : [value]) {
      if (item === null || item === '') continue;
      const bool = toBooleanValue(item);
      if (isError(bool)) return bool;
      result = result || bool;
    }
  }
  return result;
}));

registerFunction('NOT', fn(1, 1, (args, ctx) => {
  const value = toBooleanValue(ctx.evalArg(args[0]));
  return isError(value) ? value : !value;
}));

registerFunction('TRUE', fn(0, 0, () => true));
registerFunction('FALSE', fn(0, 0, () => false));

// ---- Conditional aggregations -----------------------------------------

registerFunction('COUNTIF', fn(2, 2, (args, ctx) => {
  const range = ctx.evalArg(args[0]);
  if (!isRange(range)) return error(ERROR_CODES.VALUE);
  const result = filteredValues(range, ctx.evalArg(args[1]));
  if (result.error) return result.error;
  return result.matches.filter(Boolean).length;
}));

registerFunction('COUNTIFS', fn(2, 255, (args, ctx) => {
  if (args.length % 2 !== 0) return error(ERROR_CODES.VALUE);
  const arrays: Array<{ values: EvalValue[]; criteria: Criteria }> = [];
  for (let index = 0; index < args.length; index += 2) {
    const range = ctx.evalArg(args[index]);
    if (!isRange(range)) return error(ERROR_CODES.VALUE);
    const criteria = parseCriteria(ctx.evalArg(args[index + 1]));
    if (criteria instanceof FormulaError) return criteria;
    arrays.push({ values: flattenRange(range), criteria });
  }
  const length = Math.min(...arrays.map((item) => item.values.length));
  let count = 0;
  for (let index = 0; index < length; index += 1) {
    if (arrays.every((item) => matchesCriteria(item.values[index], item.criteria))) count += 1;
  }
  return count;
}));

registerFunction('SUMIF', fn(2, 3, (args, ctx) => {
  const range = ctx.evalArg(args[0]);
  if (!isRange(range)) return error(ERROR_CODES.VALUE);
  const criteria = parseCriteria(ctx.evalArg(args[1]));
  if (criteria instanceof FormulaError) return criteria;
  const sumRange = args[2] ? ctx.evalArg(args[2]) : range;
  if (!isRange(sumRange)) return error(ERROR_CODES.VALUE);
  const values = flattenRange(range);
  const sums = flattenRange(sumRange);
  let total = 0;
  for (let index = 0; index < values.length; index += 1) {
    if (!matchesCriteria(values[index], criteria)) continue;
    const value = sums[index];
    if (isError(value)) return value;
    if (typeof value === 'number') total += value;
  }
  return total;
}));

registerFunction('SUMIFS', fn(3, 255, (args, ctx) => {
  if (args.length % 2 === 0) return error(ERROR_CODES.VALUE);
  const sumRange = ctx.evalArg(args[0]);
  if (!isRange(sumRange)) return error(ERROR_CODES.VALUE);
  const sumValues = flattenRange(sumRange);
  const conditions: Array<{ values: EvalValue[]; criteria: Criteria }> = [];
  for (let index = 1; index < args.length; index += 2) {
    const range = ctx.evalArg(args[index]);
    if (!isRange(range)) return error(ERROR_CODES.VALUE);
    const criteria = parseCriteria(ctx.evalArg(args[index + 1]));
    if (criteria instanceof FormulaError) return criteria;
    conditions.push({ values: flattenRange(range), criteria });
  }
  let total = 0;
  for (let index = 0; index < sumValues.length; index += 1) {
    if (!conditions.every((condition) => matchesCriteria(condition.values[index], condition.criteria))) continue;
    const value = sumValues[index];
    if (isError(value)) return value;
    if (typeof value === 'number') total += value;
  }
  return total;
}));

registerFunction('AVERAGEIF', fn(2, 3, (args, ctx) => {
  const range = ctx.evalArg(args[0]);
  if (!isRange(range)) return error(ERROR_CODES.VALUE);
  const criteria = parseCriteria(ctx.evalArg(args[1]));
  if (criteria instanceof FormulaError) return criteria;
  const averageRange = args[2] ? ctx.evalArg(args[2]) : range;
  if (!isRange(averageRange)) return error(ERROR_CODES.VALUE);
  const values = flattenRange(range);
  const numbers = flattenRange(averageRange);
  let total = 0;
  let count = 0;
  for (let index = 0; index < values.length; index += 1) {
    if (!matchesCriteria(values[index], criteria)) continue;
    const value = numbers[index];
    if (isError(value)) return value;
    if (typeof value === 'number') {
      total += value;
      count += 1;
    }
  }
  return count === 0 ? error(ERROR_CODES.DIV_ZERO) : total / count;
}));

registerFunction('AVERAGEIFS', fn(3, 255, (args, ctx) => {
  if (args.length % 2 === 0) return error(ERROR_CODES.VALUE);
  const averageRange = ctx.evalArg(args[0]);
  if (!isRange(averageRange)) return error(ERROR_CODES.VALUE);
  const numbers = flattenRange(averageRange);
  const conditions: Array<{ values: EvalValue[]; criteria: Criteria }> = [];
  for (let index = 1; index < args.length; index += 2) {
    const range = ctx.evalArg(args[index]);
    if (!isRange(range)) return error(ERROR_CODES.VALUE);
    const criteria = parseCriteria(ctx.evalArg(args[index + 1]));
    if (criteria instanceof FormulaError) return criteria;
    conditions.push({ values: flattenRange(range), criteria });
  }
  let total = 0;
  let count = 0;
  for (let index = 0; index < numbers.length; index += 1) {
    if (!conditions.every((condition) => matchesCriteria(condition.values[index], condition.criteria))) continue;
    const value = numbers[index];
    if (isError(value)) return value;
    if (typeof value === 'number') {
      total += value;
      count += 1;
    }
  }
  return count === 0 ? error(ERROR_CODES.DIV_ZERO) : total / count;
}));

for (const [name, mode] of [
  ['MAXIFS', 'max'],
  ['MINIFS', 'min'],
] as const) {
  registerFunction(name, fn(3, 255, (args, ctx) => {
    const targetRange = ctx.evalArg(args[0]);
    if (!isRange(targetRange)) return error(ERROR_CODES.VALUE);
    const targetValues = flattenRange(targetRange);
    const conditions: Array<{ values: EvalValue[]; criteria: Criteria }> = [];
    for (let index = 1; index < args.length; index += 2) {
      const range = ctx.evalArg(args[index]);
      if (!isRange(range)) return error(ERROR_CODES.VALUE);
      const criteria = parseCriteria(ctx.evalArg(args[index + 1]));
      if (criteria instanceof FormulaError) return criteria;
      conditions.push({ values: flattenRange(range), criteria });
    }
    const matches: number[] = [];
    for (let index = 0; index < targetValues.length; index += 1) {
      if (!conditions.every((condition) => matchesCriteria(condition.values[index], condition.criteria))) continue;
      const value = targetValues[index];
      if (isError(value)) return value;
      if (typeof value === 'number') matches.push(value);
    }
    if (!matches.length) return 0;
    return mode === 'max' ? Math.max(...matches) : Math.min(...matches);
  }));
}

registerFunction('SUMPRODUCT', fn(1, 255, (args, ctx) => {
  const arrays: EvalValue[][] = [];
  let length = -1;
  for (const arg of args) {
    const value = ctx.evalArg(arg);
    if (!isRange(value)) return error(ERROR_CODES.VALUE);
    const flat = flattenRange(value);
    if (length === -1) length = flat.length;
    else if (flat.length !== length) return error(ERROR_CODES.VALUE);
    arrays.push(flat);
  }
  let total = 0;
  for (let index = 0; index < length; index += 1) {
    let product = 1;
    for (const array of arrays) {
      const value = array[index];
      if (isError(value)) return value;
      product *= typeof value === 'number' ? value : 0;
    }
    total += product;
  }
  return total;
}));

// ---- Lookup ------------------------------------------------------------

registerFunction('VLOOKUP', fn(3, 4, (args, ctx) => {
  const lookup = ctx.evalArg(args[0]);
  const table = ctx.evalArg(args[1]);
  const column = numberArg(args, 2, ctx);
  if (isError(column)) return column;
  const approximate = booleanArg(args, 3, ctx, true);
  if (isError(approximate)) return approximate;
  if (!isRange(table)) return error(ERROR_CODES.VALUE);
  if (column < 1 || column > table.cols) return error(ERROR_CODES.REF);
  let rowIndex = -1;
  for (let index = 0; index < table.rows; index += 1) {
    const candidate = table.cells[index][0];
    if (approximate) {
      if (compareValues(candidate, lookup) <= 0) rowIndex = index;
    } else if (equalValues(candidate, lookup)) {
      rowIndex = index;
      break;
    }
  }
  if (rowIndex < 0) return error(ERROR_CODES.NA);
  return table.cells[rowIndex][Math.trunc(column) - 1] ?? null;
}));

registerFunction('HLOOKUP', fn(3, 4, (args, ctx) => {
  const lookup = ctx.evalArg(args[0]);
  const table = ctx.evalArg(args[1]);
  const row = numberArg(args, 2, ctx);
  const approximate = booleanArg(args, 3, ctx, true);
  if (isError(row)) return row;
  if (isError(approximate)) return approximate;
  if (!isRange(table)) return error(ERROR_CODES.VALUE);
  if (row < 1 || row > table.rows) return error(ERROR_CODES.REF);
  let columnIndex = -1;
  for (let index = 0; index < table.cols; index += 1) {
    const candidate = table.cells[0][index];
    if (approximate) {
      if (compareValues(candidate, lookup) <= 0) columnIndex = index;
    } else if (equalValues(candidate, lookup)) {
      columnIndex = index;
      break;
    }
  }
  if (columnIndex < 0) return error(ERROR_CODES.NA);
  return table.cells[Math.trunc(row) - 1][columnIndex] ?? null;
}));

registerFunction('XLOOKUP', fn(3, 6, (args, ctx) => {
  const lookup = ctx.evalArg(args[0]);
  const lookupRange = ctx.evalArg(args[1]);
  const returnRange = ctx.evalArg(args[2]);
  if (!isRange(lookupRange) || !isRange(returnRange)) return error(ERROR_CODES.VALUE);
  const lookupValues = flattenRange(lookupRange);
  const returnValues = flattenRange(returnRange);
  const matchMode = args[4] !== undefined ? numberArg(args, 4, ctx, 0) : 0;
  if (isError(matchMode)) return matchMode;
  let index = -1;
  if (matchMode === 2) {
    if (typeof lookup === 'string') {
      const regex = wildcardRegex(lookup);
      index = lookupValues.findIndex((value) => regex.test(textOf(value)));
    }
  } else {
    index = lookupValues.findIndex((value) => equalValues(value, lookup));
    if (index < 0 && (matchMode === -1 || matchMode === 1)) {
      for (let cursor = 0; cursor < lookupValues.length; cursor += 1) {
        const comparison = compareValues(lookupValues[cursor], lookup);
        if (matchMode === -1 && comparison <= 0) index = cursor;
        if (matchMode === 1 && comparison >= 0) {
          index = cursor;
          break;
        }
      }
    }
  }
  if (index < 0) return args[3] ? ctx.evalArg(args[3]) : error(ERROR_CODES.NA);
  return returnValues[index] ?? error(ERROR_CODES.REF);
}));

registerFunction('INDEX', fn(2, 3, (args, ctx) => {
  const range = ctx.evalArg(args[0]);
  if (!isRange(range)) return error(ERROR_CODES.VALUE);
  const row = numberArg(args, 1, ctx);
  if (isError(row)) return row;
  const col = args[2] ? numberArg(args, 2, ctx) : 1;
  if (isError(col)) return col;
  if (row < 0 || col < 0) return error(ERROR_CODES.VALUE);
  if (row === 0 && col > 0) {
    return ctx.createRange(range.sheetName, range.r1, range.c1 + Math.trunc(col) - 1, range.r2, range.c1 + Math.trunc(col) - 1);
  }
  if (col === 0 && row > 0) {
    return ctx.createRange(range.sheetName, range.r1 + Math.trunc(row) - 1, range.c1, range.r1 + Math.trunc(row) - 1, range.c2);
  }
  if (row > range.rows || col > range.cols) return error(ERROR_CODES.REF);
  return range.cells[Math.trunc(row) - 1]?.[Math.trunc(col) - 1] ?? null;
}));

registerFunction('MATCH', fn(2, 3, (args, ctx) => {
  const lookup = ctx.evalArg(args[0]);
  const range = ctx.evalArg(args[1]);
  const matchType = numberArg(args, 2, ctx, 1);
  if (isError(matchType)) return matchType;
  if (!isRange(range)) return error(ERROR_CODES.VALUE);
  const values = flattenRange(range);
  if (matchType === 0) {
    const index = values.findIndex((value) => equalValues(value, lookup));
    return index < 0 ? error(ERROR_CODES.NA) : index + 1;
  }
  let result = -1;
  for (let index = 0; index < values.length; index += 1) {
    const comparison = compareValues(values[index], lookup);
    if (matchType > 0 && comparison <= 0) result = index;
    if (matchType < 0 && comparison >= 0) result = index;
  }
  return result < 0 ? error(ERROR_CODES.NA) : result + 1;
}));

// ---- Text --------------------------------------------------------------

registerFunction('LEFT', fn(1, 2, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const count = numberArg(args, 1, ctx, 1);
  if (isError(text)) return text;
  if (isError(count)) return count;
  return Array.from(text).slice(0, Math.max(0, Math.trunc(count))).join('');
}));
registerFunction('RIGHT', fn(1, 2, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const count = numberArg(args, 1, ctx, 1);
  if (isError(text)) return text;
  if (isError(count)) return count;
  const characters = Array.from(text);
  return characters.slice(Math.max(0, characters.length - Math.trunc(count))).join('');
}));
registerFunction('MID', fn(3, 3, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const start = numberArg(args, 1, ctx);
  const count = numberArg(args, 2, ctx);
  if (isError(text)) return text;
  if (isError(start)) return start;
  if (isError(count)) return count;
  return Array.from(text)
    .slice(Math.max(0, Math.trunc(start) - 1), Math.max(0, Math.trunc(start) - 1) + Math.max(0, Math.trunc(count)))
    .join('');
}));
registerFunction('LEN', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  return isError(text) ? text : Array.from(text).length;
}));
registerFunction('CONCAT', fn(1, 255, (args, ctx) => {
  let result = '';
  for (const value of evaluateAll(args, ctx)) {
    if (isError(value)) return value;
    for (const item of isRange(value) ? flattenRange(value) : [value]) {
      const text = textOf(item);
      if (isErrorStringText(text)) return error(ERROR_CODES.VALUE);
      result += text;
    }
  }
  return result;
}));
registerFunction('CONCATENATE', fn(1, 255, (args, ctx) => {
  let result = '';
  for (const value of evaluateAll(args, ctx)) {
    const text = toTextValue(value);
    if (isError(text)) return text;
    result += text;
  }
  return result;
}));
registerFunction('TRIM', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  return isError(text) ? text : text.replace(/\s+/g, ' ').trim();
}));
registerFunction('LOWER', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  return isError(text) ? text : text.toLowerCase();
}));
registerFunction('UPPER', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  return isError(text) ? text : text.toUpperCase();
}));
registerFunction('PROPER', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  if (isError(text)) return text;
  return text.replace(/\b\w/g, (char) => char.toUpperCase());
}));
registerFunction('SUBSTITUTE', fn(3, 4, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const oldText = textArg(args, 1, ctx);
  const newText = textArg(args, 2, ctx);
  if (isError(text)) return text;
  if (isError(oldText)) return oldText;
  if (isError(newText)) return newText;
  if (!oldText) return text;
  if (args[3] !== undefined) {
    const instance = numberArg(args, 3, ctx);
    if (isError(instance)) return instance;
    let count = 0;
    return text.replaceAll(oldText, (match) => {
      count += 1;
      return count === Math.trunc(instance) ? newText : match;
    });
  }
  return text.split(oldText).join(newText);
}));
registerFunction('REPLACE', fn(4, 4, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const start = numberArg(args, 1, ctx);
  const count = numberArg(args, 2, ctx);
  const newText = textArg(args, 3, ctx);
  if (isError(text)) return text;
  if (isError(start)) return start;
  if (isError(count)) return count;
  if (isError(newText)) return newText;
  const characters = Array.from(text);
  return characters.slice(0, Math.max(0, Math.trunc(start) - 1)).join('') + newText + characters.slice(Math.max(0, Math.trunc(start) - 1) + Math.max(0, Math.trunc(count))).join('');
}));
registerFunction('FIND', fn(2, 3, (args, ctx) => {
  const findText = textArg(args, 0, ctx);
  const within = textArg(args, 1, ctx);
  const start = numberArg(args, 2, ctx, 1);
  if (isError(findText)) return findText;
  if (isError(within)) return within;
  if (isError(start)) return start;
  const index = within.indexOf(findText, Math.max(0, Math.trunc(start) - 1));
  return index < 0 ? error(ERROR_CODES.VALUE) : Array.from(within.slice(0, index)).length + 1;
}));
registerFunction('SEARCH', fn(2, 3, (args, ctx) => {
  const findText = textArg(args, 0, ctx);
  const within = textArg(args, 1, ctx);
  const start = numberArg(args, 2, ctx, 1);
  if (isError(findText)) return findText;
  if (isError(within)) return within;
  if (isError(start)) return start;
  const index = within.toLowerCase().indexOf(findText.toLowerCase(), Math.max(0, Math.trunc(start) - 1));
  return index < 0 ? error(ERROR_CODES.VALUE) : Array.from(within.slice(0, index)).length + 1;
}));
registerFunction('EXACT', fn(2, 2, (args, ctx) => {
  const a = textArg(args, 0, ctx);
  const b = textArg(args, 1, ctx);
  if (isError(a)) return a;
  if (isError(b)) return b;
  return a === b;
}));
registerFunction('TEXTJOIN', fn(3, 255, (args, ctx) => {
  const delimiter = textArg(args, 0, ctx);
  const ignoreEmpty = booleanArg(args, 1, ctx);
  if (isError(delimiter)) return delimiter;
  if (isError(ignoreEmpty)) return ignoreEmpty;
  const parts: string[] = [];
  for (let index = 2; index < args.length; index += 1) {
    const value = ctx.evalArg(args[index]);
    if (isError(value)) return value;
    for (const item of isRange(value) ? flattenRange(value) : [value]) {
      const text = textOf(item);
      if (ignoreEmpty && text === '') continue;
      parts.push(text);
    }
  }
  return parts.join(delimiter);
}));
registerFunction('REPT', fn(2, 2, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  const count = numberArg(args, 1, ctx);
  if (isError(text)) return text;
  if (isError(count)) return count;
  return text.repeat(Math.max(0, Math.trunc(count)));
}));
registerFunction('TEXT', fn(2, 2, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  const pattern = textArg(args, 1, ctx);
  if (isError(pattern)) return pattern;
  const scalar = unwrap(value);
  if (isError(scalar)) return scalar;
  if (typeof scalar === 'number') return formatNumeric(scalar, pattern, ctx.date1904);
  return textOf(scalar);
}));
registerFunction('VALUE', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  if (isError(text)) return text;
  const cleaned = text.replace(/[¥￥$€£,\s]/g, '');
  if (/^-?\d+(\.\d+)?%$/.test(cleaned)) return Number(cleaned.slice(0, -1)) / 100;
  const numeric = Number(cleaned);
  return Number.isFinite(numeric) ? numeric : error(ERROR_CODES.VALUE);
}));
registerFunction('CHAR', fn(1, 1, (args, ctx) => {
  const code = numberArg(args, 0, ctx);
  if (isError(code)) return code;
  return String.fromCharCode(Math.trunc(code));
}));
registerFunction('CODE', fn(1, 1, (args, ctx) => {
  const text = textArg(args, 0, ctx);
  if (isError(text)) return text;
  return text.charCodeAt(0) || error(ERROR_CODES.VALUE);
}));

// ---- Date & time -------------------------------------------------------

registerFunction('TODAY', fn(0, 0, (_args, ctx) => todaySerial(ctx.now, ctx.date1904)));
registerFunction('NOW', fn(0, 0, (_args, ctx) => nowSerial(ctx.now, ctx.date1904)));
registerFunction('DATE', fn(3, 3, (args, ctx) => {
  const year = numberArg(args, 0, ctx);
  const month = numberArg(args, 1, ctx);
  const day = numberArg(args, 2, ctx);
  if (isError(year)) return year;
  if (isError(month)) return month;
  if (isError(day)) return day;
  return makeDateSerial(Math.trunc(year), Math.trunc(month), Math.trunc(day), ctx.date1904);
}));
registerFunction('TIME', fn(3, 3, (args, ctx) => {
  const hour = numberArg(args, 0, ctx);
  const minute = numberArg(args, 1, ctx);
  const second = numberArg(args, 2, ctx);
  if (isError(hour)) return hour;
  if (isError(minute)) return minute;
  if (isError(second)) return second;
  return ((Math.trunc(hour) * 3600 + Math.trunc(minute) * 60 + Math.trunc(second)) % 86_400) / 86_400;
}));
const datePartFunction = (name: string, part: 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second' | 'weekday') => {
  registerFunction(name, fn(1, 2, (args, ctx) => {
    const serial = numberArg(args, 0, ctx);
    if (isError(serial)) return serial;
    const parts = serialToLocalParts(serial, ctx.date1904);
    if (part === 'weekday') {
      const type = numberArg(args, 1, ctx, 1);
      if (isError(type)) return type;
      if (type === 1) return parts.weekday + 1;
      if (type === 2) return ((parts.weekday + 6) % 7) + 1;
      if (type === 3) return (parts.weekday + 6) % 7;
      return parts.weekday + 1;
    }
    return parts[part];
  }));
};
datePartFunction('YEAR', 'year');
datePartFunction('MONTH', 'month');
datePartFunction('DAY', 'day');
datePartFunction('HOUR', 'hour');
datePartFunction('MINUTE', 'minute');
datePartFunction('SECOND', 'second');
datePartFunction('WEEKDAY', 'weekday');
registerFunction('EDATE', fn(2, 2, (args, ctx) => {
  const serial = numberArg(args, 0, ctx);
  const months = numberArg(args, 1, ctx);
  if (isError(serial)) return serial;
  if (isError(months)) return months;
  return addMonths(serial, Math.trunc(months), ctx.date1904);
}));
registerFunction('EOMONTH', fn(2, 2, (args, ctx) => {
  const serial = numberArg(args, 0, ctx);
  const months = numberArg(args, 1, ctx);
  if (isError(serial)) return serial;
  if (isError(months)) return months;
  return endOfMonthSerial(addMonths(serial, Math.trunc(months), ctx.date1904), ctx.date1904);
}));
registerFunction('DATEDIF', fn(3, 3, (args, ctx) => {
  const start = numberArg(args, 0, ctx);
  const end = numberArg(args, 1, ctx);
  const unitText = textArg(args, 2, ctx);
  if (isError(start)) return start;
  if (isError(end)) return end;
  if (isError(unitText)) return unitText;
  const unit = unitText.toUpperCase();
  const startDate = serialToLocalParts(start, ctx.date1904);
  const endDate = serialToLocalParts(end, ctx.date1904);
  if (end < start) return error(ERROR_CODES.NUM);
  if (unit === 'D') return end - start;
  if (unit === 'Y') return endDate.year - startDate.year - (endDate.month < startDate.month || (endDate.month === startDate.month && endDate.day < startDate.day) ? 1 : 0);
  if (unit === 'M') {
    let months = (endDate.year - startDate.year) * 12 + (endDate.month - startDate.month);
    if (endDate.day < startDate.day) months -= 1;
    return months;
  }
  if (unit === 'YM') {
    let months = endDate.month - startDate.month;
    if (endDate.day < startDate.day) months -= 1;
    return (months + 12) % 12;
  }
  if (unit === 'MD') return endDate.day - startDate.day;
  if (unit === 'YD') {
    let days = end - start;
    days %= 365;
    return days;
  }
  return error(ERROR_CODES.VALUE);
}));

// ---- Information -------------------------------------------------------

registerFunction('ISBLANK', fn(1, 1, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  if (isRange(value)) return false;
  return value === null || value === '';
}));
registerFunction('ISNUMBER', fn(1, 1, (args, ctx) => {
  const value = unwrap(ctx.evalArg(args[0]));
  return typeof value === 'number';
}));
registerFunction('ISTEXT', fn(1, 1, (args, ctx) => {
  const value = unwrap(ctx.evalArg(args[0]));
  return typeof value === 'string' && !isErrorStringText(value);
}));
registerFunction('ISLOGICAL', fn(1, 1, (args, ctx) => typeof unwrap(ctx.evalArg(args[0])) === 'boolean'));
registerFunction('ISERROR', fn(1, 1, (args, ctx) => isError(unwrap(ctx.evalArg(args[0])))));
registerFunction('ISERR', fn(1, 1, (args, ctx) => {
  const errorValue = ctx.evalArg(args[0]);
  return isError(errorValue) && errorValue.code !== ERROR_CODES.NA;
}));
registerFunction('ISNA', fn(1, 1, (args, ctx) => {
  const errorValue = ctx.evalArg(args[0]);
  return isError(errorValue) && errorValue.code === ERROR_CODES.NA;
}));
registerFunction('ISEVEN', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.abs(Math.trunc(value)) % 2 === 0;
}));
registerFunction('ISODD', fn(1, 1, (args, ctx) => {
  const value = numberArg(args, 0, ctx);
  return isError(value) ? value : Math.abs(Math.trunc(value)) % 2 === 1;
}));
registerFunction('N', fn(1, 1, (args, ctx) => {
  const value = unwrap(ctx.evalArg(args[0]));
  if (isError(value)) return value;
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  return 0;
}));
registerFunction('T', fn(1, 1, (args, ctx) => {
  const value = unwrap(ctx.evalArg(args[0]));
  return typeof value === 'string' ? value : '';
}));
registerFunction('NA', fn(0, 0, () => error(ERROR_CODES.NA)));
registerFunction('ROW', fn(0, 1, (args, ctx) => {
  if (args.length === 0) return ctx.currentRow + 1;
  const value = ctx.evalArg(args[0]);
  if (isRange(value)) return value.r1 + 1;
  return ctx.currentRow + 1;
}));
registerFunction('COLUMN', fn(0, 1, (args, ctx) => {
  if (args.length === 0) return ctx.currentCol + 1;
  const value = ctx.evalArg(args[0]);
  if (isRange(value)) return value.c1 + 1;
  return ctx.currentCol + 1;
}));
registerFunction('ROWS', fn(1, 1, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  return isRange(value) ? value.rows : 1;
}));
registerFunction('COLUMNS', fn(1, 1, (args, ctx) => {
  const value = ctx.evalArg(args[0]);
  return isRange(value) ? value.cols : 1;
}));

function isErrorStringText(value: string): boolean {
  return /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A|CIRC!)$/i.test(value);
}

export { evaluateAll, firstError, aggregateNumbers, toTextValue as toText };

