// 功能：验证格式设置、工作表原位改名/颜色和图片的 XLSX 往返保存；时间：2026-10-24；作者：lq。
import { act, createElement, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { createBlankWorkbook, createSheet, type WorkbookModel } from '@mini-office/document-model';
import { analyzeCapabilities, loadOoxmlPackage } from '@mini-office/ooxml';
import { applyCellFormat, cellStyle, renameSheet } from '../packages/editor-sheet/src/sheet-utils';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';
import { parseWorkbookPackage } from '../packages/editor-sheet/src/parse';
import { formatNumeric } from '../packages/editor-sheet/src/format';
import { SheetTabs } from '../packages/editor-sheet/src/SheetTabs';
import { buildSheetPrintPages } from '../packages/editor-sheet/src/SheetEditor';
import { CellFormatDialog } from '../packages/editor-sheet/src/CellFormatDialog';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
// 功能：模拟受控输入的真实编辑事件；时间：2026-10-24；作者：lq。
function inputValue(input: HTMLInputElement, value: string) {
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}
function mount(node: React.ReactElement) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const host = document.createElement('div'); document.body.append(host);
  const root = createRoot(host); act(() => root.render(node));
  return { host, dispose: () => { act(() => root.unmount()); host.remove(); } };
}

describe('单元格格式和工作表图片', () => {
  it('格式一次应用到选区，外框不污染内部边框或选区外单元格', () => {
    const book = createBlankWorkbook(); const id = book.sheets[0].id;
    book.sheets[0].cells['4:4'] = { value: '未选中' };
    const next = applyCellFormat(book, id, { r1: 0, c1: 0, r2: 2, c2: 2 }, { bold: true, fill: '#FACA15', numFmt: '0.00', shrinkToFit: true }, { preset: 'outline', side: { style: 'double', color: '#D92D20' } });
    const sheet = next.sheets[0];
    expect(cellStyle(next, sheet.cells['1:1']).borders).toBeUndefined();
    expect(cellStyle(next, sheet.cells['0:0']).borders).toMatchObject({ top: { style: 'double' }, left: { style: 'double' } });
    expect(cellStyle(next, sheet.cells['2:2']).borders).toMatchObject({ bottom: { color: '#D92D20' }, right: { style: 'double' } });
    expect(sheet.cells['4:4']).toEqual(book.sheets[0].cells['4:4']);
    expect(book.sheets[0].cells['0:0']).toBeUndefined();
  });

  it('格式弹窗切换页签和编辑草稿时不提交，确认才应用', () => {
    const apply = vi.fn(); const close = vi.fn();
    const view = mount(createElement(CellFormatDialog, { style: { numFmt: 'General' }, value: 12.34, onApply: apply, onClose: close }));
    try {
      const tab = [...document.querySelectorAll('[role="tab"]')].find((el) => el.textContent === '字体') as HTMLButtonElement;
      act(() => tab.click());
      const bold = [...document.querySelectorAll('label')].find((el) => el.textContent === '加粗')!.querySelector('input')!;
      act(() => bold.click());
      expect(apply).not.toHaveBeenCalled();
      act(() => [...document.querySelectorAll('button')].find((el) => el.textContent === '确定')!.click());
      expect(apply).toHaveBeenCalledWith({ bold: true }, undefined);
    } finally { view.dispose(); }
  });

  it('图片、尺寸锚点、颜色和格式在全新保存与覆盖保存后均保留', async () => {
    const book = createBlankWorkbook();
    book.sheets[0].tabColor = '#D92D20';
    book.sheets[0].images = [{ id: 'i1', name: '印章 & 图片.png', src: PNG, row: 3, col: 2, offsetX: 7.5, offsetY: 4, width: 180, height: 120 }];
    book.styles.format = { numFmt: '0.00%', shrinkToFit: true, verticalAlign: 'middle', borders: { bottom: { style: 'dashed', color: '#3478F6' } } };
    book.sheets[0].cells['0:0'] = { value: 0.25, styleId: 'format' };
    let bytes = await serializeXlsx(book, { name: '图片.xlsx' });
    for (let run = 0; run < 2; run++) {
      const pkg = await loadOoxmlPackage(bytes);
      const report = await analyzeCapabilities(pkg, 'xlsx', { fileName: '图片.xlsx' });
      expect(report.features.some((item) => item.code === 'xlsx.images' && item.supported)).toBe(true);
      const reopened = await parseWorkbookPackage(pkg, '图片.xlsx');
      expect(reopened.sheets[0].images?.[0]).toMatchObject({ name: '印章 & 图片.png', src: PNG, row: 3, col: 2, offsetX: expect.closeTo(7.5, 3), offsetY: 4, width: 180, height: 120 });
      expect(reopened.sheets[0].tabColor).toBe('#D92D20');
      expect(cellStyle(reopened, reopened.sheets[0].cells['0:0'])).toMatchObject({ numFmt: '0.00%', shrinkToFit: true, verticalAlign: 'middle', borders: { bottom: { style: 'dashed' } } });
      bytes = await serializeXlsx(reopened, { name: '图片.xlsx', sourceBytes: bytes });
      pkg.dispose();
    }
  });

  it('删除图片后覆盖保存会清除对应图片资源和绘图声明', async () => {
    const book = createBlankWorkbook();
    book.sheets[0].images = [{ id: 'i1', name: '删除.png', src: PNG, row: 0, col: 0, offsetX: 0, offsetY: 0, width: 20, height: 20 }];
    const source = await serializeXlsx(book, { name: '删除.xlsx' });
    book.sheets[0].images = [];
    const pkg = await loadOoxmlPackage(await serializeXlsx(book, { name: '删除.xlsx', sourceBytes: source }));
    expect(pkg.names().some((name) => /xl\/(media|drawings)\/.*mini-office-/.test(name))).toBe(false);
    expect(await pkg.text('[Content_Types].xml')).not.toContain('mini-office-sheet-');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).not.toContain('<drawing');
    pkg.dispose();
  });

  it('读取 Excel 双单元格图片锚点和没有 sheetViews 的标签颜色', async () => {
    const book = createBlankWorkbook();
    book.sheets[0].tabColor = '#3478F6';
    book.sheets[0].images = [{ id: 'i1', name: '图.png', src: PNG, row: 1, col: 1, offsetX: 0, offsetY: 0, width: 100, height: 100 }];
    const pkg = await loadOoxmlPackage(await serializeXlsx(book, { name: '图.xlsx' }));
    const path = 'xl/drawings/mini-office-sheet-1.xml';
    const drawing = (await pkg.text(path)).replaceAll('oneCellAnchor', 'twoCellAnchor').replace(/<xdr:ext[^>]+\/>/, '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>');
    pkg.zip.file(path, drawing);
    pkg.zip.file('xl/worksheets/sheet1.xml', (await pkg.text('xl/worksheets/sheet1.xml')).replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, ''));
    const changed = await loadOoxmlPackage(await pkg.zip.generateAsync({ type: 'uint8array' }));
    const reopened = await parseWorkbookPackage(changed, '图.xlsx');
    expect(reopened.sheets[0].tabColor).toBe('#3478F6');
    expect(reopened.sheets[0].images?.[0]).toMatchObject({ width: 176, height: 72 });
    changed.dispose(); pkg.dispose();
  });

  it('只有图片的工作表也能生成包含完整图片的打印页面', () => {
    const book = createBlankWorkbook(); const sheet = book.sheets[0];
    sheet.images = [{ id: 'i1', name: '签章.png', src: PNG, row: 3, col: 2, offsetX: 4, offsetY: 3, width: 180, height: 120 }];
    const host = document.createElement('div');
    try {
      const pages = buildSheetPrintPages(book, sheet, host);
      expect(pages.length).toBeGreaterThan(0);
      const image = pages[0].querySelector('img')!;
      expect(image.src).toBe(PNG);
      expect(image.alt).toBe('签章.png');
      expect(parseFloat(image.style.height)).toBeGreaterThan(0);
      expect(pages[0].querySelectorAll('tr').length).toBeGreaterThan(8);
    } finally { host.remove(); }
  });

  // 功能：验证打印页面保留单元格填充、合并关系并使用直接像素缩放；时间：2026-10-25；作者：lq。
  it('打印预览保留彩色填充且不使用会模糊文字的 transform 缩放', () => {
    const book = createBlankWorkbook(); const sheet = book.sheets[0];
    sheet.pageSetup = { ...sheet.pageSetup!, fitToWidth: 1 };
    sheet.cells['0:0'] = { value: '彩色标题', styleId: 'fill' };
    sheet.cells['1:0'] = { value: '内容' };
    book.styles.fill = { fill: '#FACA15', bold: true };
    sheet.merges.push({ r1: 0, c1: 0, r2: 0, c2: 1 });
    const host = document.createElement('div');
    try {
      const page = buildSheetPrintPages(book, sheet, host)[0];
      const table = page.querySelector('table')!;
      expect(table.style.transform).toBe('');
      expect(page.querySelector('td')?.style.background).toBe('rgb(250, 202, 21)');
      expect(page.querySelector('td')?.colSpan).toBe(2);
    } finally { host.remove(); }
  });

  it('文本、分数、科学计数和会计负数显示正确', () => {
    expect(formatNumeric(12.34, '@')).toBe('12.34');
    expect(formatNumeric(1.25, '# ?/?')).toBe('1 1/4');
    expect(formatNumeric(-0.5, '# ?/?')).toBe('-1/2');
    expect(formatNumeric(1250, '0.00E+00')).toBe('1.25E+03');
    expect(formatNumeric(-12.5, '¥0.00;(¥0.00)')).toBe('(¥12.50)');
  });
});

describe('工作表标签', () => {
  it('双击在标签上输入改名，Esc 取消，Enter 提交并维护跨表公式', () => {
    let final: WorkbookModel | undefined;
    // 功能：模拟真实工作簿受控状态与改名校验；时间：2026-10-24；作者：lq。
    function Harness() {
      const [book, setBook] = useState(() => {
        const data = createBlankWorkbook(); data.sheets.push(createSheet('汇总', 1));
        data.sheets[1].cells['0:0'] = { value: null, formula: '=Sheet1!A1' }; return data;
      });
      final = book;
      return createElement(SheetTabs, { sheets: book.sheets, activeId: book.sheets[0].id, readonly: false, onSwitch: vi.fn(), onAdd: vi.fn(), onDelete: vi.fn(), onColor: vi.fn(), onRename: (id, name) => setBook(renameSheet(book, id, name)) });
    }
    const view = mount(createElement(Harness));
    const start = () => act(() => view.host.querySelector('[role="tab"]')!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })));
    try {
      start(); const first = view.host.querySelector('input')!; inputValue(first, '取消改名');
      act(() => first.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
      expect(final!.sheets[0].name).toBe('Sheet1');
      start(); const input = view.host.querySelector('input')!; inputValue(input, '费用报销');
      act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })));
      expect(final!.sheets[0].name).toBe('费用报销');
      expect(final!.sheets[1].cells['0:0'].formula).toContain('费用报销');
      expect(view.host.querySelector('input')).toBeNull();
      start(); const duplicate = view.host.querySelector('input')!; inputValue(duplicate, '汇总');
      act(() => duplicate.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })));
      expect(final!.sheets[0].name).toBe('费用报销');
      expect(view.host.querySelector('[role="alert"]')?.textContent).toContain('已存在');
      act(() => duplicate.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
    } finally { view.dispose(); }
  });

  it('右键非当前标签可设置对应标签颜色，不修改当前工作表', () => {
    const book = createBlankWorkbook(); book.sheets.push(createSheet('第二张', 1));
    const color = vi.fn();
    const view = mount(createElement(SheetTabs, { sheets: book.sheets, activeId: book.sheets[0].id, readonly: false, onSwitch: vi.fn(), onAdd: vi.fn(), onDelete: vi.fn(), onRename: vi.fn(), onColor: color }));
    try {
      act(() => view.host.querySelectorAll('[role="tab"]')[1].dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 100, clientY: 100 })));
      act(() => [...document.querySelectorAll('button[role="menuitem"]')].find((el) => el.textContent === '工作表标签颜色…')!.dispatchEvent(new MouseEvent('click', { bubbles: true })));
      act(() => document.querySelector<HTMLButtonElement>('[aria-label="标签颜色 #D92D20"]')!.click());
      expect(color).toHaveBeenCalledWith(book.sheets[1].id, '#D92D20');
    } finally { view.dispose(); }
  });
});
