import { ERROR_CODES, FormulaError } from './types';

export interface CellReferenceToken {
  sheet?: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  rowAbs1: boolean;
  colAbs1: boolean;
  rowAbs2: boolean;
  colAbs2: boolean;
  wholeColumn?: boolean;
  wholeRow?: boolean;
}

export type Token =
  | { type: 'number'; value: number; pos: number }
  | { type: 'string'; value: string; pos: number }
  | { type: 'boolean'; value: boolean; pos: number }
  | { type: 'error'; value: string; pos: number }
  // 功能：保留引用的原文结束位置，避免结构编辑改写其他公式文本；时间：2026-10-18；作者：lq。
  | { type: 'ref'; value: CellReferenceToken; pos: number; end: number }
  | { type: 'name'; value: string; pos: number }
  | { type: 'op'; value: string; pos: number }
  | { type: 'lparen'; pos: number }
  | { type: 'rparen'; pos: number }
  | { type: 'comma'; pos: number }
  | { type: 'eof'; pos: number };

const ERROR_LITERALS = [
  ERROR_CODES.DIV_ZERO,
  ERROR_CODES.VALUE,
  ERROR_CODES.REF,
  ERROR_CODES.NAME,
  ERROR_CODES.NUM,
  ERROR_CODES.NA,
  ERROR_CODES.NULL,
  ERROR_CODES.CIRC,
];

const CELL_PATTERN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/;
const COLUMN_PATTERN = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
const ROW_PATTERN = /^(\$?)(\d+):(\$?)(\d+)/;
const SHEET_PREFIX_PATTERN = /^(?:'((?:[^']|'')+)'|([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*))!/;

function columnToIndex(label: string): number {
  let index = 0;
  for (const char of label.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return Math.max(0, index - 1);
}

export class FormulaTokenizer {
  private readonly source: string;
  private position = 0;

  constructor(source: string) {
    this.source = source;
    this.position = 0;
  }

  private skipWhitespace(): void {
    while (this.position < this.source.length && /\s/.test(this.source[this.position])) {
      this.position += 1;
    }
  }

  private tryRef(start: number): { token: CellReferenceToken; end: number } | null {
    const input = this.source;
    let position = start;
    let sheet: string | undefined;

    if (input[position] === "'") {
      let match = position + 1;
      let name = '';
      let closed = false;
      while (match < input.length) {
        if (input[match] === "'") {
          if (input[match + 1] === "'") {
            name += "'";
            match += 2;
            continue;
          }
          closed = true;
          match += 1;
          break;
        }
        name += input[match];
        match += 1;
      }
      if (!closed || input[match] !== '!') return null;
      sheet = name;
      position = match + 1;
    } else {
      const rest = input.slice(position);
      const prefix = SHEET_PREFIX_PATTERN.exec(rest);
      if (prefix) {
        sheet = prefix[1] ?? prefix[2];
        position += prefix[0].length;
      }
    }

    const rest = input.slice(position);
    // 功能：列字母前的 $ 锁定列，行号前的 $ 锁定行；时间：2026-10-18；作者：lq。
    const cell = CELL_PATTERN.exec(rest);
    if (cell) {
      let end = position + cell[0].length;
      const rangeTail = /^:(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(input.slice(end));
      if (rangeTail) {
        end += rangeTail[0].length;
        const next = input[end];
        if (next && /[A-Za-z0-9_]/.test(next)) return null;
        if (next === '(') return null;
        return {
          token: {
            sheet,
            rowAbs1: cell[3] === '$',
            colAbs1: cell[1] === '$',
            r1: Number(cell[4]) - 1,
            c1: columnToIndex(cell[2]),
            rowAbs2: rangeTail[3] === '$',
            colAbs2: rangeTail[1] === '$',
            r2: Number(rangeTail[4]) - 1,
            c2: columnToIndex(rangeTail[2]),
          },
          end,
        };
      }
      const next = input[end];
      if (next && /[A-Za-z0-9_]/.test(next)) return null;
      if (next === '(') return null;
      return {
        token: {
          sheet,
          rowAbs1: cell[3] === '$',
          colAbs1: cell[1] === '$',
          r1: Number(cell[4]) - 1,
          c1: columnToIndex(cell[2]),
          r2: Number(cell[4]) - 1,
          c2: columnToIndex(cell[2]),
          rowAbs2: cell[3] === '$',
          colAbs2: cell[1] === '$',
        },
        end,
      };
    }

    const wholeColumn = COLUMN_PATTERN.exec(rest);
    if (wholeColumn) {
      let end = position + wholeColumn[0].length;
      if (input[end] === '(') return null;
      return {
        token: {
          sheet,
          rowAbs1: false,
          colAbs1: wholeColumn[1] === '$',
          r1: 0,
          c1: columnToIndex(wholeColumn[2]),
          rowAbs2: false,
          colAbs2: wholeColumn[3] === '$',
          r2: 1_048_575,
          c2: columnToIndex(wholeColumn[4]),
          wholeColumn: true,
        },
        end,
      };
    }

    const wholeRow = ROW_PATTERN.exec(rest);
    if (wholeRow) {
      const end = position + wholeRow[0].length;
      if (input[end] === '(') return null;
      return {
        token: {
          sheet,
          rowAbs1: wholeRow[1] === '$',
          colAbs1: false,
          r1: Number(wholeRow[2]) - 1,
          c1: 0,
          rowAbs2: wholeRow[3] === '$',
          colAbs2: false,
          r2: Number(wholeRow[4]) - 1,
          c2: 16_383,
          wholeRow: true,
        },
        end,
      };
    }

    if (sheet) {
      // A sheet prefix without a reference is invalid.
      throw new FormulaError(ERROR_CODES.REF, `无效的引用：${input.slice(start, position)}`);
    }
    return null;
  }

  next(): Token {
    this.skipWhitespace();
    const start = this.position;
    if (start >= this.source.length) return { type: 'eof', pos: start };
    const input = this.source;
    const char = input[start];

    if (char === '"') {
      let text = '';
      let position = start + 1;
      while (position < input.length) {
        if (input[position] === '"') {
          if (input[position + 1] === '"') {
            text += '"';
            position += 2;
            continue;
          }
          position += 1;
          this.position = position;
          return { type: 'string', value: text, pos: start };
        }
        text += input[position];
        position += 1;
      }
      throw new FormulaError(ERROR_CODES.VALUE, '字符串缺少结束引号');
    }

    for (const literal of ERROR_LITERALS) {
      if (input.slice(start, start + literal.length).toUpperCase() === literal) {
        this.position = start + literal.length;
        return { type: 'error', value: literal, pos: start };
      }
    }

    if (char === '(') {
      this.position += 1;
      return { type: 'lparen', pos: start };
    }
    if (char === ')') {
      this.position += 1;
      return { type: 'rparen', pos: start };
    }
    if (char === ',' || char === ';') {
      this.position += 1;
      return { type: 'comma', pos: start };
    }

    const operator = /^(<=|>=|<>|[+\-*/^&%=<>])/.exec(input.slice(start));
    if (operator) {
      this.position = start + operator[0].length;
      return { type: 'op', value: operator[0], pos: start };
    }

    const reference = this.tryRef(start);
    if (reference) {
      this.position = reference.end;
      return { type: 'ref', value: reference.token, pos: start, end: reference.end };
    }

    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(input.slice(start));
    if (number) {
      this.position = start + number[0].length;
      return { type: 'number', value: Number(number[0]), pos: start };
    }

    const name = /^[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*/.exec(input.slice(start));
    if (name) {
      this.position = start + name[0].length;
      const upper = name[0].toUpperCase();
      if (upper === 'TRUE') return { type: 'boolean', value: true, pos: start };
      if (upper === 'FALSE') return { type: 'boolean', value: false, pos: start };
      return { type: 'name', value: name[0], pos: start };
    }

    throw new FormulaError(ERROR_CODES.VALUE, `无法识别的字符：${char}`);
  }
}

export function tokenize(source: string): Token[] {
  const tokenizer = new FormulaTokenizer(source);
  const tokens: Token[] = [];
  for (let index = 0; index < 100_000; index += 1) {
    const token = tokenizer.next();
    tokens.push(token);
    if (token.type === 'eof') return tokens;
  }
  throw new FormulaError(ERROR_CODES.VALUE, '公式过长');
}
