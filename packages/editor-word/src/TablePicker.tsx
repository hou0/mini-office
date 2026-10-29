/**
 * 功能：提供 Office 风格的表格行列网格与自定义尺寸，保留编辑器插入位置。
 * 时间：2026-10-19；作者：lq
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ChevronDown, Table2 } from 'lucide-react';
import { IconButton } from '@mini-office/ui';

export function TablePicker({ disabled, onInsert }: {
  disabled: boolean;
  onInsert: (rows: number, columns: number) => void;
}) {
  const [open, setOpen] = useState(false);
  const [preview, setPreview] = useState({ rows: 1, columns: 1 });
  const [rows, setRows] = useState('3');
  const [columns, setColumns] = useState('3');
  const [position, setPosition] = useState({ left: 0, top: 0 });
  const anchorRef = useRef<HTMLDivElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const validRows = Number(rows);
  const validColumns = Number(columns);
  const valid = Number.isInteger(validRows) && validRows >= 1 && validRows <= 100
    && Number.isInteger(validColumns) && validColumns >= 1 && validColumns <= 63;

  // 功能：将弹层放在可视区域内，避免工具栏滚动容器裁剪；时间：2026-10-19；作者：lq。
  useLayoutEffect(() => {
    if (!open) return;
    const anchor = anchorRef.current?.getBoundingClientRect();
    const panel = panelRef.current;
    if (!anchor || !panel) return;
    setPosition({
      left: Math.max(8, Math.min(anchor.left, window.innerWidth - panel.offsetWidth - 8)),
      top: anchor.bottom + panel.offsetHeight + 6 <= window.innerHeight - 8
        ? anchor.bottom + 6 : Math.max(8, anchor.top - panel.offsetHeight - 6),
    });
    panel.querySelector<HTMLButtonElement>('[data-table-cell]')?.focus({ preventScroll: true });
  }, [open]);

  // 功能：点击外部、按 Escape 或视口变化时收起，并支持键盘返回工具栏；时间：2026-10-19；作者：lq。
  useEffect(() => {
    if (!open) return;
    const close = () => setOpen(false);
    const outside = (event: PointerEvent) => {
      const target = event.target as Node;
      if (!panelRef.current?.contains(target) && !anchorRef.current?.contains(target)) close();
    };
    const escape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      close();
      anchorRef.current?.querySelector('button')?.focus({ preventScroll: true });
    };
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('keydown', escape);
    window.addEventListener('resize', close);
    window.addEventListener('scroll', close, true);
    return () => {
      document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('keydown', escape);
      window.removeEventListener('resize', close);
      window.removeEventListener('scroll', close, true);
    };
  }, [open]);

  // 功能：仅提交有效的行列尺寸，关闭后交由编辑器建立表格及撤销记录；时间：2026-10-19；作者：lq。
  const insert = (rowCount: number, columnCount: number) => {
    if (disabled) return;
    setOpen(false);
    onInsert(rowCount, columnCount);
  };

  return (
    <div ref={anchorRef} className="inline-flex shrink-0">
      <IconButton label="插入表格" disabled={disabled} active={open} aria-haspopup="dialog" aria-expanded={open}
        onClick={() => { setPreview({ rows: 1, columns: 1 }); setOpen(!open); }}>
        <Table2 size={15} /><ChevronDown size={9} />
      </IconButton>
      {open && !disabled ? createPortal(
        <div ref={panelRef} role="dialog" aria-label="插入表格"
          className="mo-pop-in fixed z-[100] w-[264px] max-h-[calc(100vh-16px)] overflow-y-auto rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-raised)] p-3 text-[12px] text-[var(--mo-text-primary)] shadow-pop"
          style={position}>
          <div className="mb-2 font-medium" aria-live="polite">{preview.columns} 列 × {preview.rows} 行表格</div>
          <div role="group" aria-label="选择表格行列" className="grid grid-cols-10 gap-1">
            {Array.from({ length: 80 }, (_, index) => {
              const row = Math.floor(index / 10) + 1;
              const column = index % 10 + 1;
              const selected = row <= preview.rows && column <= preview.columns;
              return <button key={index} type="button" data-table-cell
                aria-label={`${column} 列 ${row} 行`} tabIndex={row === preview.rows && column === preview.columns ? 0 : -1}
                className="h-5 w-5 rounded-sm border focus-visible:outline focus-visible:outline-2 focus-visible:outline-[var(--mo-accent)]"
                style={{ borderColor: selected ? 'var(--mo-accent)' : 'var(--mo-line-strong)',
                  background: selected ? 'var(--mo-accent-soft)' : 'var(--mo-surface)' }}
                onMouseEnter={() => setPreview({ rows: row, columns: column })}
                onFocus={() => setPreview({ rows: row, columns: column })}
                onMouseDown={(event) => event.preventDefault()}
                onKeyDown={(event) => {
                  const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -10, ArrowDown: 10 };
                  if (!(event.key in offsets)) return;
                  event.preventDefault();
                  const next = Math.max(0, Math.min(79, index + offsets[event.key]));
                  panelRef.current?.querySelectorAll<HTMLButtonElement>('[data-table-cell]')[next]?.focus();
                }}
                onClick={() => insert(row, column)} />;
            })}
          </div>
          <form className="mt-3 border-t border-[var(--mo-line)] pt-3"
            onSubmit={(event) => { event.preventDefault(); if (valid) insert(validRows, validColumns); }}>
            <div className="mb-2 font-medium">自定义表格</div>
            <div className="flex gap-3">
              <label className="flex min-w-0 flex-1 items-center gap-2">列数
                <input type="number" min={1} max={63} required value={columns}
                  onChange={(event) => setColumns(event.target.value)} className="mo-field h-7 w-full min-w-0" />
              </label>
              <label className="flex min-w-0 flex-1 items-center gap-2">行数
                <input type="number" min={1} max={100} required value={rows}
                  onChange={(event) => setRows(event.target.value)} className="mo-field h-7 w-full min-w-0" />
              </label>
            </div>
            <div className="my-2 text-[10px] text-[var(--mo-text-muted)]">1–63 列，1–100 行</div>
            <button type="submit" disabled={!valid}
              className="h-7 w-full rounded-md bg-[var(--mo-accent)] text-[var(--mo-accent-ink)] disabled:opacity-40">插入表格</button>
          </form>
        </div>, document.body,
      ) : null}
    </div>
  );
}
