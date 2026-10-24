// 功能：工作表浮动图片支持选中、拖动、等比例缩放和删除，操作结束后一次提交历史；时间：2026-10-24；作者：lq。
import { useRef, useState } from 'react';
import type { SheetImage } from '@mini-office/document-model';

export function SheetImageLayer({ images, colOffsets, rowOffsets, readonly, zoom, onChange, onDelete }: {
  images: SheetImage[]; colOffsets: number[]; rowOffsets: number[]; readonly: boolean; zoom: number;
  onChange: (id: string, patch: Partial<SheetImage>) => void; onDelete: (id: string) => void;
}) {
  const [selected, setSelected] = useState<string | null>(null);
  const [preview, setPreview] = useState<{ id: string; left: number; top: number; width: number; height: number } | null>(null);
  const drag = useRef<{ id: string; x: number; y: number; left: number; top: number; width: number; height: number; resize: boolean } | null>(null);
  const previewRef = useRef<typeof preview>(null);
  // 功能：将拖动后的像素位置换算为标准单元格锚点，保存重开保持位置；时间：2026-10-24；作者：lq。
  const anchorAt = (offsets: number[], value: number) => {
    let index = 0;
    while (index < offsets.length - 2 && offsets[index + 1] <= value) index++;
    return { index, offset: value - offsets[index] };
  };
  return <>{images.map((image) => {
    const box = preview?.id === image.id ? preview : { left: (colOffsets[image.col] ?? 0) + image.offsetX, top: (rowOffsets[image.row] ?? 0) + image.offsetY, width: image.width, height: image.height };
    return <div key={image.id} data-sheet-image={image.id} tabIndex={0} role="group" aria-label={`图片 ${image.name}`} className="absolute outline-none" style={{ left: box.left, top: box.top, width: box.width, height: box.height, zIndex: 15, outline: selected === image.id ? '2px solid var(--mo-accent)' : undefined, touchAction: 'none' }}
      onFocus={() => setSelected(image.id)} onBlur={(event) => { if (!drag.current && !event.currentTarget.contains(event.relatedTarget as Node | null)) setSelected(null); }}
      onPointerDown={(event) => {
        event.stopPropagation(); event.preventDefault(); event.currentTarget.focus({ preventScroll: true }); setSelected(image.id);
        if (readonly || event.button !== 0) return;
        event.currentTarget.setPointerCapture(event.pointerId);
        drag.current = { id: image.id, x: event.clientX, y: event.clientY, ...box, resize: !!(event.target as HTMLElement).closest('[data-image-resize]') };
        previewRef.current = { id: image.id, ...box };
      }}
      onPointerMove={(event) => {
        const start = drag.current;
        if (!start || start.id !== image.id) return;
        const dx = (event.clientX - start.x) / zoom; const dy = (event.clientY - start.y) / zoom;
        const factor = Math.max(16 / start.width, 16 / start.height, 1 + (Math.abs(dx / start.width) > Math.abs(dy / start.height) ? dx / start.width : dy / start.height));
        const next = { id: image.id, left: Math.max(0, start.left + (start.resize ? 0 : dx)), top: Math.max(0, start.top + (start.resize ? 0 : dy)), width: start.resize ? start.width * factor : start.width, height: start.resize ? start.height * factor : start.height };
        previewRef.current = next; setPreview(next);
      }}
      onPointerUp={(event) => {
        if (!drag.current || drag.current.id !== image.id) return;
        const next = previewRef.current; drag.current = null; previewRef.current = null; setPreview(null);
        event.currentTarget.releasePointerCapture(event.pointerId);
        if (!next || next.width === image.width && next.height === image.height && next.left === (colOffsets[image.col] ?? 0) + image.offsetX && next.top === (rowOffsets[image.row] ?? 0) + image.offsetY) return;
        const col = anchorAt(colOffsets, next.left); const row = anchorAt(rowOffsets, next.top);
        onChange(image.id, { col: col.index, row: row.index, offsetX: col.offset, offsetY: row.offset, width: next.width, height: next.height });
      }}
      onPointerCancel={() => { drag.current = null; previewRef.current = null; setPreview(null); }}
      onKeyDown={(event) => {
        event.stopPropagation();
        if (readonly) return;
        if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); onDelete(image.id); }
        else if (event.key === 'Escape') { setSelected(null); event.currentTarget.blur(); }
        else if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
          event.preventDefault(); const step = event.shiftKey ? 10 : 1;
          const col = anchorAt(colOffsets, Math.max(0, box.left + (event.key === 'ArrowRight' ? step : event.key === 'ArrowLeft' ? -step : 0)));
          const row = anchorAt(rowOffsets, Math.max(0, box.top + (event.key === 'ArrowDown' ? step : event.key === 'ArrowUp' ? -step : 0)));
          onChange(image.id, { col: col.index, row: row.index, offsetX: col.offset, offsetY: row.offset });
        }
      }}>
      <img src={image.src} alt={image.name} draggable={false} className="h-full w-full select-none" style={{ objectFit: 'fill', pointerEvents: 'none' }} />
      {selected === image.id && !readonly ? <><button type="button" aria-label={`删除图片 ${image.name}`} className="absolute right-0 top-0 h-6 w-6 bg-[var(--mo-surface)] text-[var(--mo-danger)] shadow" onPointerDown={(event) => event.stopPropagation()} onClick={(event) => { event.stopPropagation(); onDelete(image.id); }}>×</button><span data-image-resize role="button" aria-label="缩放图片" className="absolute -bottom-1 -right-1 h-3 w-3 cursor-nwse-resize border border-white bg-[var(--mo-accent)]" /></> : null}
    </div>;
  })}</>;
}
