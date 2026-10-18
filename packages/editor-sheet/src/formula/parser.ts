import { ERROR_CODES, FormulaError, type ErrorCode } from './types';
import { tokenize, type CellReferenceToken, type Token } from './tokenizer';

export type Expr =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'error'; value: ErrorCode }
  | { type: 'empty' }
  | {
      type: 'ref';
      sheet?: string;
      row: number;
      col: number;
      rowAbs: boolean;
      colAbs: boolean;
    }
  | {
      type: 'range';
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
  | { type: 'name'; name: string }
  | { type: 'unary'; op: '+' | '-'; expr: Expr }
  | { type: 'percent'; expr: Expr }
  | { type: 'binary'; op: string; left: Expr; right: Expr }
  | { type: 'func'; name: string; args: Expr[] };

const PRECEDENCE: Record<string, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
};

class Parser {
  private readonly tokens: Token[];
  private index = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  private peek(): Token {
    return this.tokens[this.index] ?? { type: 'eof', pos: 0 };
  }

  private next(): Token {
    const token = this.peek();
    this.index += 1;
    return token;
  }

  private expect(type: Token['type']): Token {
    const token = this.next();
    if (token.type !== type) {
      throw new FormulaError(ERROR_CODES.NAME, `公式语法错误：期望 ${type}`);
    }
    return token;
  }

  parse(): Expr {
    const expression = this.parseExpression(0);
    if (this.peek().type !== 'eof') {
      throw new FormulaError(ERROR_CODES.NAME, '公式包含多余的内容');
    }
    return expression;
  }

  private parseExpression(minPrecedence: number): Expr {
    let left = this.parseUnary();
    for (;;) {
      const token = this.peek();
      if (token.type !== 'op') break;
      const precedence = PRECEDENCE[token.value];
      if (precedence === undefined || precedence < minPrecedence) break;
      this.next();
      const right = this.parseExpression(precedence + 1);
      left = { type: 'binary', op: token.value, left, right };
    }
    return left;
  }

  private parseUnary(): Expr {
    const token = this.peek();
    if (token.type === 'op' && (token.value === '-' || token.value === '+')) {
      this.next();
      return { type: 'unary', op: token.value, expr: this.parseUnary() };
    }
    let expression = this.parsePrimary();
    while (this.peek().type === 'op' && (this.peek() as { value: string }).value === '%') {
      this.next();
      expression = { type: 'percent', expr: expression };
    }
    return expression;
  }

  private parsePrimary(): Expr {
    const token = this.next();
    switch (token.type) {
      case 'number':
        return { type: 'number', value: token.value };
      case 'string':
        return { type: 'string', value: token.value };
      case 'boolean':
        return { type: 'boolean', value: token.value };
      case 'error':
        return { type: 'error', value: token.value as ErrorCode };
      case 'ref':
        return referenceToExpr(token.value);
      case 'name': {
        if (this.peek().type === 'lparen') {
          this.next();
          const args: Expr[] = [];
          if (this.peek().type !== 'rparen') {
            for (;;) {
              if (this.peek().type === 'comma') {
                args.push({ type: 'empty' });
                this.next();
                continue;
              }
              args.push(this.parseExpression(0));
              if (this.peek().type === 'comma') {
                this.next();
                continue;
              }
              break;
            }
          }
          this.expect('rparen');
          return { type: 'func', name: token.value, args };
        }
        return { type: 'name', name: token.value };
      }
      case 'lparen': {
        const expression = this.parseExpression(0);
        this.expect('rparen');
        return expression;
      }
      default:
        throw new FormulaError(ERROR_CODES.NAME, '公式语法错误');
    }
  }
}

function referenceToExpr(ref: CellReferenceToken): Expr {
  if (ref.r1 === ref.r2 && ref.c1 === ref.c2 && !ref.wholeRow && !ref.wholeColumn) {
    return {
      type: 'ref',
      sheet: ref.sheet,
      row: ref.r1,
      col: ref.c1,
      rowAbs: ref.rowAbs1,
      colAbs: ref.colAbs1,
    };
  }
  return {
    type: 'range',
    sheet: ref.sheet,
    r1: ref.r1,
    c1: ref.c1,
    r2: ref.r2,
    c2: ref.c2,
    rowAbs1: ref.rowAbs1,
    colAbs1: ref.colAbs1,
    rowAbs2: ref.rowAbs2,
    colAbs2: ref.colAbs2,
    wholeColumn: ref.wholeColumn,
    wholeRow: ref.wholeRow,
  };
}

export function parseFormula(formula: string): Expr {
  const source = formula.startsWith('=') ? formula.slice(1) : formula;
  if (!source.trim()) throw new FormulaError(ERROR_CODES.VALUE, '公式为空');
  return new Parser(tokenize(source)).parse();
}

export function columnLabel(index: number): string {
  let n = index + 1;
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function refToText(sheet: string | undefined, row: number, col: number, rowAbs: boolean, colAbs: boolean): string {
  const prefix = sheet ? `'${sheet.replace(/'/g, "''")}'!` : '';
  return `${prefix}${colAbs ? '$' : ''}${columnLabel(col)}${rowAbs ? '$' : ''}${row + 1}`;
}

export function printFormula(expr: Expr): string {
  switch (expr.type) {
    case 'number':
      return String(expr.value);
    case 'string':
      return `"${expr.value.replace(/"/g, '""')}"`;
    case 'boolean':
      return expr.value ? 'TRUE' : 'FALSE';
    case 'error':
      return expr.value;
    case 'empty':
      return '';
    case 'ref':
      return refToText(expr.sheet, expr.row, expr.col, expr.rowAbs, expr.colAbs);
    case 'range':
      return `${refToText(expr.sheet, expr.r1, expr.c1, expr.rowAbs1, expr.colAbs1)}:${refToText(
        undefined,
        expr.r2,
        expr.c2,
        expr.rowAbs2,
        expr.colAbs2,
      )}`;
    case 'name':
      return expr.name;
    case 'unary':
      return `${expr.op}${printFormula(expr.expr)}`;
    case 'percent':
      return `${printFormula(expr.expr)}%`;
    case 'binary':
      return `(${printFormula(expr.left)}${expr.op}${printFormula(expr.right)})`;
    case 'func':
      return `${expr.name}(${expr.args.map(printFormula).join(',')})`;
    default:
      return '';
  }
}

export interface ReferenceUse {
  sheet?: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function collectReferences(expr: Expr, output: ReferenceUse[] = []): ReferenceUse[] {
  switch (expr.type) {
    case 'ref':
      output.push({ sheet: expr.sheet, r1: expr.row, c1: expr.col, r2: expr.row, c2: expr.col });
      break;
    case 'range':
      output.push({ sheet: expr.sheet, r1: expr.r1, c1: expr.c1, r2: expr.r2, c2: expr.c2 });
      break;
    case 'unary':
    case 'percent':
      collectReferences(expr.expr, output);
      break;
    case 'binary':
      collectReferences(expr.left, output);
      collectReferences(expr.right, output);
      break;
    case 'func':
      for (const arg of expr.args) collectReferences(arg, output);
      break;
    default:
      break;
  }
  return output;
}

function translateReference(expr: Expr, deltaRow: number, deltaCol: number): Expr {
  switch (expr.type) {
    case 'ref':
      return {
        ...expr,
        row: expr.rowAbs ? expr.row : Math.max(0, expr.row + deltaRow),
        col: expr.colAbs ? expr.col : Math.max(0, expr.col + deltaCol),
      };
    case 'range':
      return {
        ...expr,
        r1: expr.rowAbs1 ? expr.r1 : Math.max(0, expr.r1 + deltaRow),
        c1: expr.colAbs1 ? expr.c1 : Math.max(0, expr.c1 + deltaCol),
        r2: expr.rowAbs2 ? expr.r2 : Math.max(0, expr.r2 + deltaRow),
        c2: expr.colAbs2 ? expr.c2 : Math.max(0, expr.c2 + deltaCol),
      };
    case 'unary':
      return { ...expr, expr: translateReference(expr.expr, deltaRow, deltaCol) };
    case 'percent':
      return { ...expr, expr: translateReference(expr.expr, deltaRow, deltaCol) };
    case 'binary':
      return {
        ...expr,
        left: translateReference(expr.left, deltaRow, deltaCol),
        right: translateReference(expr.right, deltaRow, deltaCol),
      };
    case 'func':
      return {
        ...expr,
        args: expr.args.map((arg) => translateReference(arg, deltaRow, deltaCol)),
      };
    default:
      return expr;
  }
}

/** Translate relative references in a formula, used by copy/paste and shared formulas. */
export function translateFormula(formula: string, deltaRow: number, deltaCol: number): string {
  if (!formula) return formula;
  try {
    const ast = parseFormula(formula);
    return `=${printFormula(translateReference(ast, deltaRow, deltaCol))}`;
  } catch {
    return formula;
  }
}

export function normalizeFormula(formula: string): string {
  try {
    const ast = parseFormula(formula);
    return `=${printFormula(ast)}`;
  } catch {
    return formula;
  }
}
