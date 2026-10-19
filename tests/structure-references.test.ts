/** 功能：验证结构变更后的真实引用与计算结果，包含绝对引用和跨表往返；时间：2026-10-18；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { createBlankWorkbook, createSheet } from '@mini-office/document-model';
import { insertRows, deleteRows, insertColumns, deleteColumns, renameSheet, deleteSheet, sortRange } from '../packages/editor-sheet/src/sheet-utils';
import { recalculateWorkbook } from '../packages/editor-sheet/src/formula/engine';
import { translateFormula } from '../packages/editor-sheet/src/formula/parser';
import { editFormulaReferences } from '../packages/editor-sheet/src/formula/references';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';
import { sheetFormatHandler } from '../packages/editor-sheet/src/handler';

/** 功能：构造有跨表和混合锁定引用的工作簿；时间：2026-10-18；作者：lq。 */
function sample() {
  const book = createBlankWorkbook();
  const sheet = book.sheets[0];
  sheet.name = '明细';
  sheet.cells = { '0:0': { value: 10 }, '1:0': { value: 20 }, '2:0': { value: 30 }, '0:1': { value: null, formula: '=SUM($A$1:A3)' } };
  const second = createSheet('汇总', 1);
  second.cells = { '0:0': { value: null, formula: "=SUM('明细'!$A$1:$A$3)" }, '0:1': { value: null, formula: '=A1' } };
  book.sheets.push(second);
  book.definedNames = { 总额: "'明细'!$A$1:$A$3" };
  return book;
}

describe('结构编辑引用维护', () => {
  it('插行同步更新本表、跨表、命名区域，另一表本地引用保持原样', () => {
    const original = sample();
    const book = insertRows(original, original.sheets[0].id, 1);
    expect(book.sheets[0].cells['0:1'].formula).toBe('=SUM($A$1:A4)');
    expect(book.sheets[1].cells['0:0'].formula).toBe("=SUM('明细'!$A$1:$A$4)");
    expect(book.sheets[1].cells['0:1'].formula).toBe('=A1');
    expect(book.definedNames?.总额).toBe("'明细'!$A$1:$A$4");
    expect(recalculateWorkbook(book).computed?.[book.sheets[1].id]['0:0'].value).toBe(60);
    expect(original.sheets[0].cells['0:1'].formula).toBe('=SUM($A$1:A3)');
  });

  it('删除范围中间部分会缩短引用，全部删除返回 REF', () => {
    const original = sample();
    const partial = deleteRows(original, original.sheets[0].id, 1);
    expect(partial.sheets[1].cells['0:0'].formula).toBe("=SUM('明细'!$A$1:$A$2)");
    expect(recalculateWorkbook(partial).computed?.[partial.sheets[1].id]['0:0'].value).toBe(40);
    const all = deleteRows(original, original.sheets[0].id, 0, 3);
    expect(all.sheets[1].cells['0:0'].formula).toBe('=SUM(#REF!)');
  });

  it('插列和删列同时调整绝对、相对引用', () => {
    const book = sample();
    const next = insertColumns(book, book.sheets[0].id, 0);
    expect(next.sheets[0].cells['0:2'].formula).toBe('=SUM($B$1:B3)');
    expect(next.sheets[1].cells['0:0'].formula).toBe("=SUM('明细'!$B$1:$B$3)");
    expect(deleteColumns(next, book.sheets[0].id, 1).sheets[1].cells['0:0'].formula).toBe('=SUM(#REF!)');
  });

  it('保留公式字符串、空格和优先级，整列引用不因插行变长', () => {
    const edit = { kind: 'rows' as const, sheet: '明细', at: 1, count: 1 };
    expect(editFormulaReferences('=IF(A2>0, "A2", -(A2 + 3))', '明细', edit)).toBe('=IF(A3>0, "A2", -(A3 + 3))');
    expect(editFormulaReferences('=SUM($A:$C)', '明细', edit)).toBe('=SUM($A:$C)');
    expect(editFormulaReferences('=SUM($1:$3)', '明细', edit)).toBe('=SUM($1:$4)');
  });

  it('保留混合引用锁定语义', () => {
    expect(translateFormula('=$A1+A$1', 1, 1)).toBe('=($A2+B$1)');
    expect(translateFormula('=SUM($A1:B$2)', 1, 1)).toBe('=SUM($A2:C$2)');
  });

  it('更名处理中文、空格、引号和大小写不同的引用', () => {
    const book = sample();
    book.sheets[1].cells['1:0'] = { value: null, formula: '=SUM(明细!A1:A3)' };
    const next = renameSheet(book, book.sheets[0].id, "本月 O'Brien");
    expect(next.sheets[1].cells['0:0'].formula).toBe("=SUM('本月 O''Brien'!$A$1:$A$3)");
    expect(next.definedNames?.总额).toBe("'本月 O''Brien'!$A$1:$A$3");
    expect(recalculateWorkbook(next).computed?.[next.sheets[1].id]['0:0'].value).toBe(60);
  });

  it('拒绝重复、非法名称及不可安全改写的语法，原模型不变', () => {
    const book = sample();
    expect(() => renameSheet(book, book.sheets[0].id, '汇总')).toThrow('已存在');
    expect(() => renameSheet(book, book.sheets[0].id, '非法/名称')).toThrow('工作表名');
    book.sheets[1].cells['0:0'].formula = '=Table1[Amount]';
    expect(() => insertRows(book, book.sheets[0].id, 0)).toThrow('暂不支持');
    expect(book.sheets[0].cells['0:0'].value).toBe(10);
  });

  it('删除工作表使相关公式和命名区域返回 REF', () => {
    const book = sample();
    const next = deleteSheet(book, book.sheets[0].id);
    expect(next.sheets[0].cells['0:0'].formula).toBe('=SUM(#REF!)');
    expect(next.definedNames?.总额).toBe('#REF!');
    expect(recalculateWorkbook(next).computed?.[next.sheets[0].id]['0:0'].value).toBe('#REF!');
  });

  it('排序按行移动相对公式并保留绝对行', () => {
    const book = sample();
    const sheet = book.sheets[0];
    sheet.cells = { '0:0': { value: 2 }, '0:1': { value: null, formula: '=A1+$C$1' },
      '1:0': { value: 1 }, '1:1': { value: null, formula: '=A2+$C$1' }, '0:2': { value: 10 } };
    const next = sortRange(book, sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }, 0, true);
    expect(next.sheets[0].cells['0:1'].formula).toBe('=A1+$C$1');
    expect(next.sheets[0].cells['1:1'].formula).toBe('=A2+$C$1');
    const computed = recalculateWorkbook(next).computed?.[sheet.id];
    expect(computed?.['0:1'].value).toBe(11);
    expect(computed?.['1:1'].value).toBe(12);
  });

  it('结构编辑和更名后保存重开保持跨表计算结果', async () => {
    const book = sample();
    const edited = renameSheet(insertRows(book, book.sheets[0].id, 0), book.sheets[0].id, '新 明细');
    const bytes = await serializeXlsx(edited, { name: 'edited.xlsx' });
    const reopened = await sheetFormatHandler.open({ name: 'edited.xlsx', bytes });
    const result = recalculateWorkbook(reopened.data);
    expect(result.computed?.[result.sheets[1].id]['0:0'].value).toBe(60);
  });
});
