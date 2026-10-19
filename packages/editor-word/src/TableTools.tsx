/** 功能：提供选中表格的行列结构、尺寸、合并和单元格格式工具；时间：2026-10-19；作者：lq。 */
import { useState } from 'react';
import type { DocTable } from '@mini-office/document-model';
import { Button } from '@mini-office/ui';
import { pxToMm, mmToPx } from '@mini-office/ooxml';
import { tableCellPlacements } from './doc-utils';
import { intersects, tableSelectionRect, type TableAction, type TableSelection, type formatTable } from './table-editing';

type Format = Parameters<typeof formatTable>[2];

/** 功能：用毫米输入提交有效尺寸，允许暂时清空输入并防止非法数值；时间：2026-10-19；作者：lq。 */
function Dimension({ label, value, onChange }: { label: string; value: number; onChange: (px: number) => void }) {
  const [draft, setDraft] = useState<string>();
  const shown = draft ?? String(Math.round(pxToMm(value) * 10) / 10);
  return <label className="flex items-center gap-1 text-xs">{label}<input aria-label={label} type="number" min="1" max="500" step="0.5" className="w-16 rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] px-1 py-1" value={shown}
    onChange={(event) => setDraft(event.target.value)}
    onBlur={() => { const n = Number(draft); if (draft?.trim() && Number.isFinite(n) && n > 0 && n <= 500) onChange(mmToPx(n)); setDraft(undefined); }}
    onKeyDown={(event) => { if (event.key === 'Enter') event.currentTarget.blur(); if (event.key === 'Escape') { event.preventDefault(); setDraft(undefined); } }} />毫米</label>;
}

/** 功能：工具栏保持表格选区，按上下文启用合并和拆分；时间：2026-10-19；作者：lq。 */
export function TableTools({ table, selection, onAction, onFormat }: { table: DocTable; selection: TableSelection; onAction: (action: TableAction) => void; onFormat: (format: Format) => void }) {
  const rect = tableSelectionRect(table, selection);
  if (!rect) return null;
  const cells = tableCellPlacements(table).filter((p) => intersects(p, rect));
  const actions: [TableAction, string][] = [['select-table', '选择整个表格'], ['select-rows', '选择当前行'], ['select-columns', '选择当前列'], ['row-above', '上方插行'], ['row-below', '下方插行'], ['column-before', '左侧插列'], ['column-after', '右侧插列'], ['row-delete', '删除行'], ['column-delete', '删除列'], ['distribute-rows', '平均行高'], ['distribute-columns', '平均列宽'], ['fit-page', '适应页宽'], ['header', '重复标题行'], ['table-delete', '删除表格']];
  return <div contentEditable={false} role="toolbar" aria-label="表格工具" className="flex shrink-0 flex-wrap items-center gap-2 border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-3 py-2" onMouseDown={(event) => { if ((event.target as HTMLElement).closest('button')) event.preventDefault(); }}>
    <span className="text-xs text-[var(--mo-text-muted)]">表格 · {rect.bottom - rect.top} 行 × {rect.right - rect.left} 列</span>
    <Button size="sm" disabled={cells.length < 2} onClick={() => onAction('merge')}>合并单元格</Button>
    <Button size="sm" disabled={!cells.some((p) => p.cell.colSpan > 1 || p.cell.rowSpan > 1)} onClick={() => onAction('split')}>拆分单元格</Button>
    <Dimension key={`w-${selection.anchorId}-${selection.focusId}`} label="列宽" value={table.columnWidthsPx[rect.left] ?? 100} onChange={(columnWidth) => onFormat({ columnWidth })} />
    <Dimension key={`h-${selection.anchorId}-${selection.focusId}`} label="行高" value={table.rows[rect.top]?.heightPx ?? 28} onChange={(rowHeight) => onFormat({ rowHeight })} />
    <select aria-label="表格操作" className="rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] p-1 text-xs" value="" onChange={(event) => onAction(event.target.value as TableAction)}><option value="" disabled>行列与表格操作</option>{actions.map(([id, label]) => <option key={id} value={id}>{label}</option>)}</select>
    <select aria-label="单元格垂直对齐" value={cells[0]?.cell.verticalAlign ?? 'top'} className="bg-[var(--mo-surface)] text-xs" onChange={(event) => onFormat({ cells: { verticalAlign: event.target.value as 'top' | 'middle' | 'bottom' } })}><option value="top">顶端对齐</option><option value="middle">垂直居中</option><option value="bottom">底端对齐</option></select>
    <select aria-label="表格对齐" value={table.alignment ?? 'left'} className="bg-[var(--mo-surface)] text-xs" onChange={(event) => onFormat({ alignment: event.target.value as 'left' | 'center' | 'right' })}><option value="left">表格居左</option><option value="center">表格居中</option><option value="right">表格居右</option></select>
    <label className="flex items-center gap-1 text-xs">底色<input aria-label="单元格底色" type="color" value={cells[0]?.cell.backgroundColor ?? '#ffffff'} onChange={(event) => onFormat({ cells: { backgroundColor: event.target.value } })} className="h-6 w-7" /></label>
    <Button size="sm" onClick={() => onFormat({ cells: { backgroundColor: undefined } })}>清除底色</Button>
    <select aria-label="单元格边框" value="" className="bg-[var(--mo-surface)] text-xs" onChange={(event) => { const border = { style: event.target.value === 'none' ? 'none' as const : 'single' as const, widthPx: 1, color: '#000000' }; onFormat({ cells: { borders: { top: border, bottom: border, left: border, right: border } } }); }}><option value="" disabled>边框</option><option value="all">所有边框</option><option value="none">无边框</option></select>
    <span className="text-[10px] text-[var(--mo-text-muted)]">拖选文字或 Shift + 点击选择多个单元格；拖动右边线或下边线调整尺寸</span>
  </div>;
}
