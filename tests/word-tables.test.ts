/** 功能：验证合并网格的结构操作、尺寸及 DOCX 往返，防止编辑丢失数据；时间：2026-10-19；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocTable } from '@mini-office/document-model';
import { makeParagraph, tableCellPlacements } from '../packages/editor-word/src/doc-utils';
import { editTable, formatTable, tableSelectionRect, type TableSelection } from '../packages/editor-word/src/table-editing';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { parseDocx } from '../packages/editor-word/src/parse';

/** 功能：构造每格带独立内容的三行三列表；时间：2026-10-19；作者：lq。 */
function fixture(): DocTable {
  return { type: 'table', id: 't', columnWidthsPx: [90, 100, 110], rows: Array.from({ length: 3 }, (_, r) => ({ id: `r${r}`, cells: Array.from({ length: 3 }, (_, c) => ({ id: `${r}-${c}`, colSpan: 1, rowSpan: 1, blocks: [makeParagraph([{ text: `${r}-${c}`, style: {} }])] })) })) };
}
function select(anchorId: string, focusId = anchorId): TableSelection { return { tableId: 't', anchorId, focusId }; }
/** 功能：检查每个逻辑坐标只有一个单元格占据，且没有空洞或越界；时间：2026-10-19；作者：lq。 */
function valid(table: DocTable) {
  const seen = new Set<string>();
  for (const p of tableCellPlacements(table)) {
    for (let r = p.row; r < p.row + p.cell.rowSpan; r++) for (let c = p.col; c < p.col + p.cell.colSpan; c++) {
      const key = `${r}:${c}`;
      expect(seen.has(key)).toBe(false); expect(r).toBeLessThan(table.rows.length); expect(c).toBeLessThan(table.columnWidthsPx.length); seen.add(key);
    }
  }
  expect(seen.size).toBe(table.rows.length * table.columnWidthsPx.length);
}
function texts(table: DocTable) { return tableCellPlacements(table).flatMap(({ cell }) => cell.blocks.flatMap((block) => block.type === 'paragraph' ? block.runs.map((r) => r.text ?? '') : [])).filter(Boolean).sort(); }

describe('Word 表格编辑', () => {
  it('矩形合并保留所有文字，拆分后恢复完整网格且不修改原表', () => {
    const table = fixture();
    const merged = editTable(table, select('0-0', '1-1'), 'merge')!;
    expect(merged.rows[0].cells[0]).toMatchObject({ colSpan: 2, rowSpan: 2, widthPx: 190 });
    expect(texts(merged)).toEqual(texts(table)); valid(merged);
    const split = editTable(merged, select('0-0'), 'split')!;
    valid(split); expect(split.rows.every((row) => row.cells.length === 3)).toBe(true);
    expect(texts(split)).toEqual(texts(table)); expect(table.rows[0].cells[0].colSpan).toBe(1);
  });
  it('选区遇到部分合并区域时扩展成完整矩形', () => {
    const table = editTable(fixture(), select('0-0', '1-1'), 'merge')!;
    expect(tableSelectionRect(table, select('0-2', '2-0'))).toEqual({ top: 0, left: 0, bottom: 3, right: 3 });
  });
  it('穿过合并区域插入行列时延伸合并关系并补齐空格', () => {
    let table = editTable(fixture(), select('0-0', '1-1'), 'merge')!;
    table = editTable(table, select('0-2'), 'row-below')!;
    expect(table.rows[0].cells[0].rowSpan).toBe(3); valid(table);
    table = editTable(table, select('2-0'), 'column-after')!;
    expect(table.rows[0].cells[0].colSpan).toBe(3); valid(table);
    expect(texts(table)).toEqual(texts(fixture()));
  });
  it('删除纵向合并起始行时将内容移动到剩余合并区域', () => {
    let table = editTable(fixture(), select('0-0', '1-0'), 'merge')!;
    table = editTable(table, select('0-1'), 'row-delete')!;
    valid(table); expect(table.rows[0].cells[0].rowSpan).toBe(1);
    expect(texts(table)).toContain('0-0'); expect(texts(table)).toContain('1-0');
    expect(texts(table)).not.toContain('0-1');
  });
  it('删除横向合并起始列时保留剩余区域内容', () => {
    let table = editTable(fixture(), select('0-0', '0-1'), 'merge')!;
    table = editTable(table, select('1-0'), 'column-delete')!;
    valid(table); expect(table.columnWidthsPx).toEqual([100, 110]);
    expect(texts(table)).toContain('0-0'); expect(texts(table)).toContain('0-1');
  });
  it('删除全部行或列移除表格', () => {
    expect(editTable(fixture(), select('0-0', '2-2'), 'row-delete')).toBeNull();
    expect(editTable(fixture(), select('0-0', '2-2'), 'column-delete')).toBeNull();
  });
  it('尺寸修改同步合并单元格总宽度，样式仅作用选区', () => {
    const table = editTable(fixture(), select('0-0', '1-1'), 'merge')!;
    const formatted = formatTable(table, select('0-0'), { columnWidth: 120, rowHeight: 48, cells: { backgroundColor: '#ff0000', verticalAlign: 'middle' } });
    expect(formatted.columnWidthsPx).toEqual([120, 120, 110]);
    expect(formatted.rows.map((r) => r.heightPx)).toEqual([48, 48, undefined]);
    expect(formatted.rows[0].cells[0]).toMatchObject({ widthPx: 240, backgroundColor: '#ff0000', verticalAlign: 'middle' });
    expect(formatted.rows[0].cells[1].backgroundColor).toBeUndefined(); valid(formatted);
  });
  it('均分及适应页宽后保留结构和内容', () => {
    let table = editTable(fixture(), select('0-0', '2-2'), 'distribute-columns')!;
    expect(table.columnWidthsPx).toEqual([100, 100, 100]);
    table = editTable(table, select('0-0'), 'fit-page', 600)!;
    expect(table.columnWidthsPx).toEqual([200, 200, 200]); valid(table);
  });
  it('插列保持总宽度，大数值列宽自动收缩其他列以保持在页内', () => {
    const inserted = editTable(fixture(), select('0-0'), 'column-after')!;
    expect(inserted.columnWidthsPx.reduce((a, b) => a + b, 0)).toBeCloseTo(300);
    const resized = formatTable(fixture(), select('0-0'), { columnWidth: 500 }, 300);
    expect(resized.columnWidthsPx).toEqual([252, 24, 24]); valid(resized);
    const all = formatTable(fixture(), select('0-0', '2-2'), { columnWidth: 500 }, 300);
    expect(all.columnWidthsPx).toEqual([100, 100, 100]);
  });
  it('全表合并保留空的被覆盖行，DOCX 保存重开两次仍保持网格、样式和文字', async () => {
    const original = fixture();
    const merged = editTable(original, select('0-0', '2-2'), 'merge')!;
    let table = formatTable(merged, select('0-0'), { columnWidth: 130, rowHeight: 42, cells: { backgroundColor: '#FFEEDD', verticalAlign: 'bottom' } });
    table = editTable(table, select('0-0'), 'header')!;
    for (let i = 0; i < 2; i++) {
      const doc = { ...createBlankDoc(), blocks: [table] };
      const reopened = await parseDocx(await serializeDocx(doc, { name: 'table.docx' }));
      table = reopened.blocks.find((b): b is DocTable => b.type === 'table')!;
      expect(table.rows).toHaveLength(3); expect(table.rows.map((r) => r.heightPx)).toEqual([42, 42, 42]);
      expect(table.rows[0].cells[0]).toMatchObject({ rowSpan: 3, colSpan: 3, widthPx: 390, backgroundColor: '#FFEEDD', verticalAlign: 'bottom' });
      expect(table.rows[0].isHeader).toBe(true); expect(table.columnWidthsPx).toEqual([130, 130, 130]);
      expect(texts(table)).toEqual(texts(original)); valid(table);
    }
  });
});
