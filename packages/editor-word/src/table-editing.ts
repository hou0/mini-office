/** 功能：以逻辑网格实现表格选择、合并拆分及跨合并区域的行列编辑；时间：2026-10-19；作者：lq。 */
import type { DocBlock, DocTable, DocTableCell } from '@mini-office/document-model';
import { uid } from '@mini-office/shared';
import { makeParagraph, tableCellPlacements, totalTableColumns, type CellPlacement } from './doc-utils';

export interface TableSelection { tableId: string; anchorId: string; focusId: string }
export interface TableRect { top: number; bottom: number; left: number; right: number }
export type TableAction = 'select-table' | 'select-rows' | 'select-columns' | 'merge' | 'split' | 'row-above' | 'row-below' | 'row-delete' | 'column-before' | 'column-after' | 'column-delete' | 'table-delete' | 'distribute-rows' | 'distribute-columns' | 'fit-page' | 'header';

/** 功能：查找包含嵌套表格的目标表；时间：2026-10-19；作者：lq。 */
export function findTable(blocks: DocBlock[], id: string): DocTable | undefined {
  for (const block of blocks) {
    if (block.type !== 'table') continue;
    if (block.id === id) return block;
    for (const row of block.rows) for (const cell of row.cells) {
      const found = findTable(cell.blocks, id);
      if (found) return found;
    }
  }
}

/** 功能：扩展矩形选区，确保不会切开已有合并单元格；时间：2026-10-19；作者：lq。 */
export function tableSelectionRect(table: DocTable, selection: TableSelection): TableRect | undefined {
  const placements = tableCellPlacements(table);
  const a = placements.find(({ cell }) => cell.id === selection.anchorId);
  const b = placements.find(({ cell }) => cell.id === selection.focusId);
  if (!a || !b) return;
  const rect = { top: Math.min(a.row, b.row), bottom: Math.max(a.row + a.cell.rowSpan, b.row + b.cell.rowSpan), left: Math.min(a.col, b.col), right: Math.max(a.col + a.cell.colSpan, b.col + b.cell.colSpan) };
  let changed = true;
  while (changed) {
    changed = false;
    for (const p of placements) {
      if (!intersects(p, rect)) continue;
      const next = { top: Math.min(rect.top, p.row), bottom: Math.max(rect.bottom, p.row + p.cell.rowSpan), left: Math.min(rect.left, p.col), right: Math.max(rect.right, p.col + p.cell.colSpan) };
      if (JSON.stringify(next) !== JSON.stringify(rect)) { Object.assign(rect, next); changed = true; }
    }
  }
  return rect;
}

/** 功能：检测单元格是否落在矩形选区内；时间：2026-10-19；作者：lq。 */
export function intersects(p: CellPlacement, rect: TableRect): boolean {
  return p.row < rect.bottom && p.row + p.cell.rowSpan > rect.top && p.col < rect.right && p.col + p.cell.colSpan > rect.left;
}

/** 功能：创建独立空白单元格并继承视觉属性；时间：2026-10-19；作者：lq。 */
function blankCell(base?: DocTableCell): DocTableCell {
  return { ...base, id: uid('cell'), blocks: [makeParagraph()], colSpan: 1, rowSpan: 1, widthPx: undefined };
}

/** 功能：按网格坐标重建行，并同步合并单元格宽度；时间：2026-10-19；作者：lq。 */
function rebuild(table: DocTable, placements: CellPlacement[]): DocTable {
  const rows = table.rows.map((row, r) => ({ ...row, cells: placements.filter((p) => p.row === r).sort((a, b) => a.col - b.col).map((p) => ({ ...p.cell, widthPx: table.columnWidthsPx.slice(p.col, p.col + p.cell.colSpan).reduce((a, b) => a + b, 0) })) }));
  return { ...table, rows };
}

/** 功能：统一处理结构命令，保留未删除区域的文本、图片与合并关系；时间：2026-10-19；作者：lq。 */
export function editTable(table: DocTable, selection: TableSelection, action: TableAction, pageWidth = 600): DocTable | null {
  const rect = tableSelectionRect(table, selection);
  if (!rect) return table;
  let ps = tableCellPlacements(table).map((p) => ({ ...p, cell: { ...p.cell } }));
  const chosen = ps.filter((p) => intersects(p, rect));
  let next = { ...table, rows: table.rows.map((row) => ({ ...row })), columnWidthsPx: Array.from({ length: totalTableColumns(table) }, (_, i) => table.columnWidthsPx[i] ?? 100) };
  if (action === 'table-delete') return null;
  if (action === 'merge') {
    if (chosen.length < 2) return table;
    const first = chosen[0];
    first.cell = { ...first.cell, blocks: chosen.flatMap((p) => p.cell.blocks), colSpan: rect.right - rect.left, rowSpan: rect.bottom - rect.top };
    first.row = rect.top; first.col = rect.left;
    ps = ps.filter((p) => !chosen.includes(p) || p === first);
  } else if (action === 'split') {
    for (const p of chosen) {
      for (let r = 0; r < p.cell.rowSpan; r++) for (let c = 0; c < p.cell.colSpan; c++) {
        if (r || c) ps.push({ row: p.row + r, col: p.col + c, cell: blankCell(p.cell) });
      }
      p.cell.colSpan = 1; p.cell.rowSpan = 1;
    }
  } else if (action.startsWith('row-') || action.startsWith('column-')) {
    const rowAxis = action.startsWith('row-');
    const deleting = action.endsWith('delete');
    const start = rowAxis ? rect.top : rect.left;
    const end = rowAxis ? rect.bottom : rect.right;
    const count = rowAxis ? next.rows.length : next.columnWidthsPx.length;
    const at = deleting || action.endsWith('above') || action.endsWith('before') ? start : end;
    const amount = deleting ? end - start : 1;
    if (deleting && amount === count) return null;
    const output: CellPlacement[] = [];
    for (const p of ps) {
      const position = rowAxis ? p.row : p.col;
      const span = rowAxis ? p.cell.rowSpan : p.cell.colSpan;
      let newPosition = position, newSpan = span;
      if (deleting) {
        const overlap = Math.max(0, Math.min(position + span, end) - Math.max(position, start));
        newSpan -= overlap;
        newPosition = position >= end ? position - amount : position >= start ? start : position;
      } else if (position >= at) newPosition++;
      else if (position + span > at) newSpan++;
      if (newSpan <= 0) continue;
      if (rowAxis) { p.row = newPosition; p.cell.rowSpan = newSpan; }
      else { p.col = newPosition; p.cell.colSpan = newSpan; }
      output.push(p);
    }
    ps = output;
    if (rowAxis) next.rows.splice(at, deleting ? amount : 0, ...(!deleting ? [{ id: uid('row'), cells: [] }] : []));
    else {
      // 功能：插列时保持整表总宽度，避免超出页边距导致裁切；时间：2026-10-19；作者：lq。
      const oldWidth = next.columnWidthsPx.reduce((sum, width) => sum + width, 0);
      next.columnWidthsPx.splice(at, deleting ? amount : 0, ...(!deleting ? [next.columnWidthsPx[Math.min(at, count - 1)]] : []));
      if (!deleting) {
        const newWidth = next.columnWidthsPx.reduce((sum, width) => sum + width, 0);
        next.columnWidthsPx = next.columnWidthsPx.map((width) => width * oldWidth / newWidth);
      }
    }
    if (!deleting) {
      const crossCount = rowAxis ? next.columnWidthsPx.length : next.rows.length;
      for (let cross = 0; cross < crossCount; cross++) {
        const r = rowAxis ? at : cross, c = rowAxis ? cross : at;
        if (!ps.some((p) => r >= p.row && r < p.row + p.cell.rowSpan && c >= p.col && c < p.col + p.cell.colSpan)) ps.push({ row: r, col: c, cell: blankCell() });
      }
    }
  } else if (action === 'distribute-columns' || action === 'fit-page') {
    const left = action === 'fit-page' ? 0 : rect.left, right = action === 'fit-page' ? next.columnWidthsPx.length : rect.right;
    const width = action === 'fit-page' ? pageWidth : next.columnWidthsPx.slice(left, right).reduce((a, b) => a + b, 0);
    next.columnWidthsPx = next.columnWidthsPx.map((w, c) => c >= left && c < right ? Math.max(24, width / (right - left)) : w);
    next.widthPct = undefined;
  } else if (action === 'distribute-rows') {
    const height = next.rows.slice(rect.top, rect.bottom).reduce((sum, row) => sum + (row.heightPx ?? 28), 0) / (rect.bottom - rect.top);
    next.rows = next.rows.map((row, r) => r >= rect.top && r < rect.bottom ? { ...row, heightPx: height } : row);
  } else if (action === 'header') {
    const enabled = !next.rows[rect.top].isHeader;
    next.rows = next.rows.map((row, r) => r < rect.bottom ? { ...row, isHeader: enabled } : row);
  }
  return rebuild(next, ps);
}

/** 功能：统一更新选区样式和尺寸，并移除会覆盖手动列宽的旧单元格宽度；时间：2026-10-19；作者：lq。 */
export function formatTable(table: DocTable, selection: TableSelection, format: { columnWidth?: number; rowHeight?: number; cells?: Partial<Pick<DocTableCell, 'backgroundColor' | 'verticalAlign' | 'borders'>>; alignment?: DocTable['alignment'] }, maxWidth = Number.POSITIVE_INFINITY): DocTable {
  const rect = tableSelectionRect(table, selection);
  if (!rect) return table;
  const ps = tableCellPlacements(table).map((p) => ({ ...p, cell: intersects(p, rect) ? { ...p.cell, ...format.cells } : p.cell }));
  const next = { ...table, alignment: format.alignment ?? table.alignment, columnWidthsPx: Array.from({ length: totalTableColumns(table) }, (_, c) => format.columnWidth !== undefined && c >= rect.left && c < rect.right ? Math.max(24, format.columnWidth) : table.columnWidthsPx[c] ?? 100), rows: table.rows.map((row, r) => format.rowHeight !== undefined && r >= rect.top && r < rect.bottom ? { ...row, heightPx: Math.max(16, format.rowHeight) } : row) };
  if (format.columnWidth !== undefined) {
    next.widthPct = undefined;
    // 功能：数值调宽时限制在页面内，并按比例收缩未选列，防止内容被纸张边缘裁切；时间：2026-10-19；作者：lq。
    if (next.columnWidthsPx.reduce((sum, width) => sum + width, 0) > maxWidth) {
      const selectedCount = rect.right - rect.left;
      let remaining = next.columnWidthsPx.map((_, c) => c).filter((c) => c < rect.left || c >= rect.right);
      const minimum = Math.min(24, maxWidth / next.columnWidthsPx.length);
      const width = Math.max(minimum, Math.min(format.columnWidth, (maxWidth - minimum * remaining.length) / selectedCount));
      for (let c = rect.left; c < rect.right; c++) next.columnWidthsPx[c] = width;
      let budget = maxWidth - width * selectedCount;
      while (remaining.length) {
        const sum = remaining.reduce((total, c) => total + next.columnWidthsPx[c], 0);
        const narrow = remaining.filter((c) => next.columnWidthsPx[c] / sum * budget < minimum);
        if (!narrow.length) {
          for (const c of remaining) next.columnWidthsPx[c] = next.columnWidthsPx[c] / sum * budget;
          break;
        }
        for (const c of narrow) { next.columnWidthsPx[c] = minimum; budget -= minimum; }
        remaining = remaining.filter((c) => !narrow.includes(c));
      }
    }
  }
  return rebuild(next, ps);
}
