// 功能：回归表头选区视图稳定性、多列筛选与 XLSX 条件保存；时间：2026-10-24；作者：lq。
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { createBlankWorkbook } from '@mini-office/document-model';
import { SheetGrid } from '../packages/editor-sheet/src/SheetGrid';
import { applyFilter, sortRange, tableRangeForSelection, type Selection } from '../packages/editor-sheet/src/sheet-utils';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';
import { parseWorkbookPackage } from '../packages/editor-sheet/src/parse';
import { loadOoxmlPackage } from '@mini-office/ooxml';

// 功能：提供包含空白、手动隐藏和关联数据的工作簿；时间：2026-10-24；作者：lq。
function fixture() {
  const book = createBlankWorkbook();
  const sheet = book.sheets[0];
  sheet.cells = {
    '0:0': { value: '部门' }, '0:1': { value: '金额' },
    '1:0': { value: '研发' }, '1:1': { value: 20 },
    '2:0': { value: '研发' }, '2:1': { value: 10 },
    '3:0': { value: '销售' }, '3:1': { value: 20 },
    '4:0': { value: '销售' },
  };
  sheet.autoFilter = { r1: 0, c1: 0, r2: 4, c2: 1 };
  return book;
}

describe('表头选区与筛选', () => {
  it('整行、整列、全选保持滚动位置，普通单元格导航仍会滚动', () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    const book = fixture();
    const host = document.createElement('div');
    document.body.append(host);
    const root = createRoot(host);
    const scrollRef = { current: null as HTMLDivElement | null };
    let selection: Selection = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    const onSelectionChange = (next: Selection) => { selection = next; render(); };
    const render = () => root.render(createElement(SheetGrid, {
      workbook: book, sheet: book.sheets[0], selection, editing: null, scrollElementRef: scrollRef,
      onSelectionChange, onStartEdit: vi.fn(), onCommitEdit: vi.fn(), onCancelEdit: vi.fn(),
      onColumnResize: vi.fn(), onRowResize: vi.fn(), onCellContextMenu: vi.fn(),
      onHeaderContextMenu: vi.fn(), onAction: vi.fn(), onFill: vi.fn(),
    }));
    act(render);
    const scroller = scrollRef.current!;
    Object.defineProperties(scroller, { clientWidth: { value: 600 }, clientHeight: { value: 400 } });
    scroller.scrollLeft = 150; scroller.scrollTop = 75;
    const mouse = (element: Element) => act(() => { element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); });
    try {
      mouse([...host.querySelectorAll('div')].find((el) => el.textContent === '2')!);
      expect(selection.kind).toBe('row');
      expect([scroller.scrollLeft, scroller.scrollTop]).toEqual([150, 75]);
      mouse([...host.querySelectorAll('div')].find((el) => el.textContent === 'C')!);
      expect(selection.kind).toBe('col');
      expect([scroller.scrollLeft, scroller.scrollTop]).toEqual([150, 75]);
      mouse(host.querySelector('[aria-label="全选工作表"]')!);
      const end = selection.focus;
      mouse(host.querySelector('[aria-label="全选工作表"]')!);
      expect(selection.focus).toEqual(end);
      expect([scroller.scrollLeft, scroller.scrollTop]).toEqual([150, 75]);
      act(() => { selection = { anchor: { row: 90, col: 20 }, focus: { row: 90, col: 20 } }; render(); });
      expect(scroller.scrollLeft).toBeGreaterThan(150);
      expect(scroller.scrollTop).toBeGreaterThan(75);
    } finally { act(() => root.unmount()); host.remove(); }
  });

  it('从整行表头启用筛选使用实际数据边界，不包含合并标题与空白列', () => {
    const book = fixture();
    const sheet = book.sheets[0];
    sheet.autoFilter = undefined;
    sheet.merges = [{ r1: 0, c1: 0, r2: 0, c2: 1 }];
    expect(tableRangeForSelection(sheet, { kind: 'row', anchor: { row: 1, col: 0 }, focus: { row: 1, col: 99 } })).toEqual({ r1: 1, c1: 0, r2: 4, c2: 1 });
    expect(tableRangeForSelection(sheet, { anchor: { row: 2, col: 1 }, focus: { row: 2, col: 1 } }).r1).toBe(1);
  });

  it('组合两列条件，清除其中一列时保留另一列与手动隐藏行', () => {
    let book = fixture();
    const id = book.sheets[0].id;
    const range = { r1: 1, c1: 0, r2: 4, c2: 1 };
    book.sheets[0].hiddenRows = { 2: true };
    book = applyFilter(book, book.sheets[0], 0, new Set(['研发']), range);
    book = applyFilter(book, book.sheets[0], 1, new Set(['20']), range);
    expect(book.sheets[0].hiddenRows).toEqual({ 2: true, 3: true, 4: true });
    book = applyFilter(book, book.sheets[0], 0, null, range);
    expect(book.sheets[0].hiddenRows).toEqual({ 2: true, 4: true });
    book = applyFilter(book, book.sheets[0], 1, null, range);
    expect(book.sheets[0].hiddenRows).toEqual({ 2: true });
    expect(book.sheets[0].id).toBe(id);
  });

  it('降序时空白仍在末尾且整行关联数据不变', () => {
    const book = fixture();
    const next = sortRange(book, book.sheets[0], { r1: 1, c1: 0, r2: 4, c2: 1 }, 1, false);
    expect(next.sheets[0].cells['4:0'].value).toBe('销售');
    expect(next.sheets[0].cells['4:1']).toBeUndefined();
    expect(next.sheets[0].cells['3:1'].value).toBe(10);
    expect(next.sheets[0].cells['0:1'].value).toBe('金额');
  });

  it('筛选条件保存重开后可继续勾选和清除', async () => {
    let book = fixture();
    book = applyFilter(book, book.sheets[0], 0, new Set(['研发']), { r1: 1, c1: 0, r2: 4, c2: 1 });
    const bytes = await serializeXlsx(book, { name: '筛选.xlsx' });
    const pkg = await loadOoxmlPackage(bytes);
    const reopened = await parseWorkbookPackage(pkg, '筛选.xlsx');
    expect(reopened.sheets[0].filterCriteria).toEqual({ 0: ['研发'] });
    expect(reopened.sheets[0].filterBaseHiddenRows).toEqual({});
    expect(reopened.sheets[0].hiddenRows[3]).toBe(true);
    const cleared = applyFilter(reopened, reopened.sheets[0], 0, null, { r1: 1, c1: 0, r2: 4, c2: 1 });
    expect(cleared.sheets[0].hiddenRows).toEqual({});
  });
  it('格式化金额筛选保存重开后仍显示金额，并保留空白选项', async () => {
    let book = fixture();
    book.styles.money = { numFmt: '¥0.00' };
    book.sheets[0].cells['1:1'].styleId = 'money';
    book = applyFilter(book, book.sheets[0], 1, new Set(['¥20.00', '']), { r1: 1, c1: 0, r2: 4, c2: 1 });
    const bytes = await serializeXlsx(book, { name: '金额.xlsx' });
    const pkg = await loadOoxmlPackage(bytes);
    const reopened = await parseWorkbookPackage(pkg, '金额.xlsx');
    expect(reopened.sheets[0].filterCriteria?.[1]).toContain('¥20.00');
    expect(reopened.sheets[0].filterCriteria?.[1]).toContain('');
    expect(reopened.sheets[0].hiddenRows[4]).toBeUndefined();
  });

});
