import { describe, expect, it } from 'vitest';
import { createBlankWorkbook, type WorkbookModel } from '@mini-office/document-model';
import {
  loadOoxmlPackage,
  parseXmlDom,
  firstDescendant,
} from '@mini-office/ooxml';
import { parseWorkbookPackage } from '../packages/editor-sheet/src/parse';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';
import { recalculateWorkbook } from '../packages/editor-sheet/src/formula/engine';

function demoWorkbook(): WorkbookModel {
  const base = createBlankWorkbook();
  const sheet = base.sheets[0];
  sheet.id = 's1';
  sheet.name = '合同台账';
  sheet.cells = {
    '0:0': { value: '编号' },
    '0:1': { value: '金额', styleId: 'bold' },
    '1:0': { value: 'HT-001' },
    '1:1': { value: 12000, styleId: 'money' },
    '2:0': { value: 'HT-002' },
    '2:1': { value: 30000, styleId: 'money' },
    '3:1': { value: null, formula: '=SUM(B2:B3)', styleId: 'money' },
    '4:1': { value: 46313, styleId: 'date' },
    '5:1': { value: null, formula: '=B4*2' },
  };
  sheet.merges = [{ r1: 0, c1: 2, r2: 0, c2: 3 }];
  sheet.cells['0:2'] = { value: '合并表头', styleId: 'header' };
  sheet.freeze = { rows: 0, cols: 1 };
  sheet.colWidths = { 0: 120, 1: 100 };
  sheet.rowHeights = { 0: 30 };
  sheet.autoFilter = { r1: 0, c1: 0, r2: 5, c2: 1 };
  const second = {
    ...sheet,
    id: 's2',
    name: '汇总',
    order: 1,
    cells: {
      '0:0': { value: '合计' },
      '0:1': { value: null, formula: '=SUM(合同台账!B2:B4)' },
    },
    merges: [],
    freeze: undefined,
    autoFilter: undefined,
    colWidths: {},
    rowHeights: {},
    hiddenRows: {},
    hiddenCols: {},
  };
  return {
    ...base,
    sheets: [sheet, second],
    activeSheetId: 's1',
    styles: {
      default: { fontFamily: '等线', fontSizePt: 11, numFmt: 'General' },
      bold: { fontFamily: '等线', fontSizePt: 11, bold: true, numFmt: 'General' },
      header: { fontFamily: '等线', fontSizePt: 12, bold: true, fill: '#FFF2CC', horizontalAlign: 'center' },
      money: { fontFamily: '等线', fontSizePt: 11, numFmt: '¥#,##0.00', horizontalAlign: 'right' },
      date: { fontFamily: '等线', fontSizePt: 11, numFmt: 'yyyy-MM-dd', horizontalAlign: 'center' },
    },
    definedNames: { 数据区域: '合同台账!$B$2:$B$5' },
    stats: { cells: Object.keys(sheet.cells).length + 2, sheets: 2 },
  };
}

describe('xlsx round trip', () => {
  it('serializes and reparses values, formulas, styles, merges and freeze panes', async () => {
    const original = recalculateWorkbook(demoWorkbook());
    expect(original.computed?.s1?.['3:1']?.value).toBe(42000);
    const bytes = await serializeXlsx(original, { name: 'demo.xlsx' });
    expect(bytes.byteLength).toBeGreaterThan(1000);
    const pkg = await loadOoxmlPackage(bytes);
    try {
      expect(pkg.has('xl/workbook.xml')).toBe(true);
      const parsed = await parseWorkbookPackage(pkg, 'demo.xlsx');
      expect(parsed.sheets).toHaveLength(2);
      const sheet = parsed.sheets.find((item) => item.name === '合同台账')!;
      expect(sheet).toBeTruthy();
      expect(sheet.cells['0:0']?.value).toBe('编号');
      expect(sheet.cells['1:1']?.value).toBe(12000);
      expect(sheet.cells['3:1']?.formula).toBe('=SUM(B2:B3)');
      expect(parsed.computed?.[sheet.id]?.['3:1']?.value).toBe(42000);
      expect(parsed.computed?.[sheet.id]?.['4:1']?.display).toBe('2026-10-18');
      expect(sheet.merges).toEqual([{ r1: 0, c1: 2, r2: 0, c2: 3 }]);
      expect(sheet.freeze).toEqual({ rows: 0, cols: 1 });
      expect(sheet.colWidths[0]).toBe(120);
      expect(sheet.rowHeights[0]).toBe(30);
      const amountStyle = parsed.styles[sheet.cells['1:1']?.styleId ?? ''];
      expect(amountStyle?.numFmt).toBe('¥#,##0.00');
      expect(amountStyle?.horizontalAlign).toBe('right');
      const headerStyle = parsed.styles[sheet.cells['0:2']?.styleId ?? ''];
      expect(headerStyle?.fill).toBe('#FFF2CC');
      expect(headerStyle?.bold).toBe(true);
      const summary = parsed.sheets.find((item) => item.name === '汇总')!;
      expect(parsed.computed?.[summary.id]?.['0:1']?.value).toBe(84000);
    } finally {
      pkg.dispose();
    }
  });

  it('can round trip a serialized package twice', async () => {
    const first = await serializeXlsx(recalculateWorkbook(demoWorkbook()), { name: 'demo.xlsx' });
    const pkg = await loadOoxmlPackage(first);
    const parsed = await parseWorkbookPackage(pkg, 'demo.xlsx');
    pkg.dispose();
    const second = await serializeXlsx(parsed, { name: 'demo.xlsx', sourceBytes: first });
    const secondPackage = await loadOoxmlPackage(second);
    try {
      const again = await parseWorkbookPackage(secondPackage, 'demo.xlsx');
      expect(again.sheets).toHaveLength(2);
      const summary = again.sheets.find((item) => item.name === '汇总')!;
      expect(again.computed?.[summary.id]?.['0:1']?.value).toBe(84000);
      const doc = parseXmlDom(await secondPackage.text('xl/workbook.xml'));
      expect(firstDescendant(doc, 'workbook')).toBeTruthy();
    } finally {
      secondPackage.dispose();
    }
  });
});
