import { describe, expect, it } from 'vitest';
import {
  createBlankWorkbook,
  createSheet,
  type SheetModel,
  type WorkbookModel,
} from '@mini-office/document-model';
import {
  recalculateWorkbook,
  translateFormula,
  parseFormula,
  collectReferences,
  evaluateFormula,
} from '@mini-office/editor-sheet';
import { formatNumeric } from '../packages/editor-sheet/src/format';

function sheet(name: string, cells: Record<string, string | number | boolean>): SheetModel {
  const model = createSheet(name, 0);
  model.id = name;
  model.cells = {};
  for (const [address, value] of Object.entries(cells)) {
    const [col, row] = [address.replace(/\d+/g, ''), Number(address.replace(/\D+/g, '')) - 1];
    const colIndex = col
      .toUpperCase()
      .split('')
      .reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
    model.cells[`${row}:${colIndex}`] =
      typeof value === 'string' && value.startsWith('=')
        ? { value: null, formula: value }
        : { value };
    model.dimension.rows = Math.max(model.dimension.rows, row + 1);
    model.dimension.cols = Math.max(model.dimension.cols, colIndex + 1);
  }
  return model;
}

function workbook(sheets: SheetModel[]): WorkbookModel {
  const base = createBlankWorkbook();
  return {
    ...base,
    sheets,
    activeSheetId: sheets[0].id,
    styles: {
      default: { fontFamily: '等线', fontSizePt: 11, numFmt: 'General' },
      money: { numFmt: '¥#,##0.00' },
    },
  };
}

function value(model: WorkbookModel, sheetIndex: number, address: string): string | number | boolean | null {
  const calculated = recalculateWorkbook(model);
  const target = calculated.sheets[sheetIndex];
  const col = address.replace(/\d+/g, '');
  const row = Number(address.replace(/\D+/g, '')) - 1;
  const colIndex = col
    .toUpperCase()
    .split('')
    .reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  return calculated.computed?.[target.id]?.[`${row}:${colIndex}`]?.value ?? null;
}

function display(model: WorkbookModel, sheetIndex: number, address: string): string {
  const calculated = recalculateWorkbook(model);
  const target = calculated.sheets[sheetIndex];
  const col = address.replace(/\d+/g, '');
  const row = Number(address.replace(/\D+/g, '')) - 1;
  const colIndex = col
    .toUpperCase()
    .split('')
    .reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  return calculated.computed?.[target.id]?.[`${row}:${colIndex}`]?.display ?? '';
}

describe('formula engine', () => {
  it('evaluates arithmetic and precedence', () => {
    const model = workbook([sheet('Sheet1', { A1: '=1+2*3', A2: '=(1+2)*3', A3: '=-2^2', A4: '=10/4' })]);
    expect(value(model, 0, 'A1')).toBe(7);
    expect(value(model, 0, 'A2')).toBe(9);
    expect(value(model, 0, 'A3')).toBe(4);
    expect(value(model, 0, 'A4')).toBe(2.5);
  });

  it('evaluates cell references and ranges', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: 1,
        A2: 2,
        A3: 3,
        B1: '=SUM(A1:A3)',
        B2: '=AVERAGE(A1:A3)',
        B3: '=COUNT(A1:A3)',
        B4: '=MAX(A1:A3)',
        B5: '=MIN(A1:A3)',
        B6: '=A1+A2*2',
      }),
    ]);
    expect(value(model, 0, 'B1')).toBe(6);
    expect(value(model, 0, 'B2')).toBe(2);
    expect(value(model, 0, 'B3')).toBe(3);
    expect(value(model, 0, 'B4')).toBe(3);
    expect(value(model, 0, 'B5')).toBe(1);
    expect(value(model, 0, 'B6')).toBe(5);
  });

  it('supports logical functions and error handling', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: 10,
        A2: '=IF(A1>5,"big","small")',
        A3: '=IFERROR(1/0,"safe")',
        A4: '=IFS(A1>20,"x",A1>5,"y")',
        A5: '=AND(A1>5,A1<20)',
        A6: '=OR(A1<5,A1=10)',
      }),
    ]);
    expect(value(model, 0, 'A2')).toBe('big');
    expect(value(model, 0, 'A3')).toBe('safe');
    expect(value(model, 0, 'A4')).toBe('y');
    expect(value(model, 0, 'A5')).toBe(true);
    expect(value(model, 0, 'A6')).toBe(true);
  });

  it('supports VLOOKUP, XLOOKUP, INDEX and MATCH', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: '张三',
        B1: 88,
        A2: '李四',
        B2: 92,
        A3: '王五',
        B3: 79,
        D1: '=VLOOKUP("李四",A1:B3,2,FALSE)',
        D2: '=XLOOKUP("王五",A1:A3,B1:B3)',
        D3: '=INDEX(B1:B3,MATCH("张三",A1:A3,0))',
      }),
    ]);
    expect(value(model, 0, 'D1')).toBe(92);
    expect(value(model, 0, 'D2')).toBe(79);
    expect(value(model, 0, 'D3')).toBe(88);
  });

  it('supports text functions', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: '  Mini Office  ',
        A2: '=TRIM(A1)',
        A3: '=LEFT("合同编号",2)',
        A4: '=RIGHT("ABCD",2)',
        A5: '=MID("ABCDEF",2,3)',
        A6: '=LEN("中文abc")',
        A7: '=CONCAT("A","-",1)',
        A8: '=UPPER("abc")',
      }),
    ]);
    expect(value(model, 0, 'A2')).toBe('Mini Office');
    expect(value(model, 0, 'A3')).toBe('合同');
    expect(value(model, 0, 'A4')).toBe('CD');
    expect(value(model, 0, 'A5')).toBe('BCD');
    expect(value(model, 0, 'A6')).toBe(5);
    expect(value(model, 0, 'A7')).toBe('A-1');
    expect(value(model, 0, 'A8')).toBe('ABC');
  });

  it('supports date functions and round trips', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: '=DATE(2026,10,18)',
        A2: '=YEAR(A1)',
        A3: '=MONTH(A1)',
        A4: '=DAY(A1)',
        A5: '=DATE(2026,13,1)',
        A6: '=YEAR(A5)',
        A7: '=MONTH(A5)',
        A8: '=DAY(A5)',
      }),
    ]);
    expect(value(model, 0, 'A2')).toBe(2026);
    expect(value(model, 0, 'A3')).toBe(10);
    expect(value(model, 0, 'A4')).toBe(18);
    expect(value(model, 0, 'A6')).toBe(2027);
    expect(value(model, 0, 'A7')).toBe(1);
    expect(value(model, 0, 'A8')).toBe(1);
  });

  it('supports SUMIF and COUNTIF criteria', () => {
    const model = workbook([
      sheet('Sheet1', {
        A1: '水果',
        B1: '数量',
        A2: '苹果',
        B2: 10,
        A3: '香蕉',
        B3: 20,
        A4: '苹果',
        B4: 30,
        D1: '=SUMIF(A2:A4,"苹果",B2:B4)',
        D2: '=COUNTIF(B2:B4,">15")',
        D3: '=SUMIF(B2:B4,">=20")',
      }),
    ]);
    expect(value(model, 0, 'D1')).toBe(40);
    expect(value(model, 0, 'D2')).toBe(2);
    expect(value(model, 0, 'D3')).toBe(50);
  });

  it('resolves cross-sheet references', () => {
    const model = workbook([
      sheet('Sheet1', { A1: 10 }),
      sheet('Sheet2', { A1: '=Sheet1!A1*2', A2: "='Sheet1'!A1+5" }),
    ]);
    expect(value(model, 1, 'A1')).toBe(20);
    expect(value(model, 1, 'A2')).toBe(15);
  });

  it('detects circular references', () => {
    const model = workbook([sheet('Sheet1', { A1: '=A1+1', A2: '=A3', A3: '=A2' })]);
    expect(display(model, 0, 'A1')).toBe('#CIRC!');
    expect(display(model, 0, 'A2')).toBe('#CIRC!');
  });

  it('propagates errors through arithmetic', () => {
    const model = workbook([sheet('Sheet1', { A1: '=1/0', A2: '=A1+1', A3: '=IFERROR(A2,"x")' })]);
    expect(display(model, 0, 'A1')).toBe('#DIV/0!');
    expect(display(model, 0, 'A2')).toBe('#DIV/0!');
    expect(value(model, 0, 'A3')).toBe('x');
  });

  it('translates shared formulas', () => {
    expect(translateFormula('=A1+$B$2', 1, 0)).toContain('A2');
    expect(translateFormula('=A1+$B$2', 1, 0)).toContain('$B$2');
    const ast = parseFormula('=SUM(A1:B2)+Sheet2!C3');
    const refs = collectReferences(ast);
    expect(refs).toHaveLength(2);
    expect(refs[1].sheet).toBe('Sheet2');
  });

  it('formats numbers and dates', () => {
    expect(formatNumeric(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatNumeric(0.125, '0.0%')).toBe('12.5%');
    expect(formatNumeric(12, '¥#,##0.00')).toBe('¥12.00');
    expect(formatNumeric(46313, 'yyyy-MM-dd')).toBe('2026-10-18');
  });

  it('evaluates formulas without persisting computed caches', () => {
    const model = workbook([sheet('Sheet1', { A1: 2, A2: 3, A3: '=A1*A2' })]);
    const result = evaluateFormula('=A1+A2', model, 'Sheet1');
    expect(result).toBe(5);
  });
});
