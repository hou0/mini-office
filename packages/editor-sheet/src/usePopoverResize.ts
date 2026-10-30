// 功能：为筛选浮层提供四边四角鼠标缩放和本次运行内的尺寸记忆；时间：2026-10-30；作者：lq。
import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';

export interface PopoverSize {
  width: number;
  /** null 表示随内容自适应高度，首次拖动上下边或角落时锁定为实测高度。 */
  height: number | null;
}

export type ResizeAxis = 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'nw';

export const RESIZE_HANDLES: ReadonlyArray<{ axis: ResizeAxis; label: string; className: string }> = [
  { axis: 'n', label: '上边', className: 'left-2 right-2 top-0 h-1.5 cursor-ns-resize' },
  { axis: 's', label: '下边', className: 'left-2 right-2 bottom-0 h-1.5 cursor-ns-resize' },
  { axis: 'w', label: '左边', className: 'top-2 bottom-2 left-0 w-1.5 cursor-ew-resize' },
  { axis: 'e', label: '右边', className: 'top-2 bottom-2 right-0 w-1.5 cursor-ew-resize' },
  { axis: 'nw', label: '左上角', className: 'left-0 top-0 h-3 w-3 cursor-nwse-resize' },
  { axis: 'ne', label: '右上角', className: 'right-0 top-0 h-3 w-3 cursor-nesw-resize' },
  { axis: 'sw', label: '左下角', className: 'left-0 bottom-0 h-3 w-3 cursor-nesw-resize' },
  { axis: 'se', label: '右下角', className: 'right-0 bottom-0 h-3 w-3 cursor-nwse-resize' },
];

const MARGIN = 8;
const ANCHOR_GAP = 4;
const DEFAULT_WIDTH = 280;
const DEFAULT_MIN_WIDTH = 240;
const DEFAULT_MIN_HEIGHT = 220;

// 功能：把数值限制在区间内，并兼容 min 大于 max 的极端视口；时间：2026-10-30；作者：lq。
function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), Math.max(min, max));
}

interface DragOrigin {
  axis: ResizeAxis;
  pointerX: number;
  pointerY: number;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface UsePopoverResizeOptions {
  anchor: { left: number; bottom: number };
  initial?: PopoverSize | null;
  minWidth?: number;
  minHeight?: number;
  onCommit?: (size: PopoverSize) => void;
}

export interface ResizeHandleHandlers {
  onPointerDown: (event: ReactPointerEvent<HTMLElement>) => void;
  onPointerMove: (event: ReactPointerEvent<HTMLElement>) => void;
  onPointerUp: (event: ReactPointerEvent<HTMLElement>) => void;
  onPointerCancel: (event: ReactPointerEvent<HTMLElement>) => void;
  onLostPointerCapture: (event: ReactPointerEvent<HTMLElement>) => void;
}

// 功能：维护浮层尺寸和位置，拖边只改单轴、拖角改双轴，左/上边拖动时保持对侧锚定；时间：2026-10-30；作者：lq。
export function usePopoverResize(options: UsePopoverResizeOptions): {
  size: PopoverSize;
  left: number;
  top: number;
  handleProps: (axis: ResizeAxis) => ResizeHandleHandlers;
} {
  const { anchor, initial, minWidth = DEFAULT_MIN_WIDTH, minHeight = DEFAULT_MIN_HEIGHT, onCommit } = options;
  const [size, setSize] = useState<PopoverSize>(() => ({
    width: initial ? clamp(initial.width, minWidth, window.innerWidth - MARGIN * 2) : DEFAULT_WIDTH,
    height: initial?.height ?? null,
  }));
  const [left, setLeft] = useState(() => {
    const width = initial ? clamp(initial.width, minWidth, window.innerWidth - MARGIN * 2) : DEFAULT_WIDTH;
    return clamp(anchor.left, MARGIN, window.innerWidth - width - MARGIN);
  });
  const [top, setTop] = useState(() => clamp(anchor.bottom + ANCHOR_GAP, MARGIN, window.innerHeight - 480));
  const dragRef = useRef<DragOrigin | null>(null);
  const sizeRef = useRef<PopoverSize>(size);
  const commitRef = useRef(onCommit);
  commitRef.current = onCommit;

  // 功能：拖拽中禁用文本选择，组件卸载时兜底恢复；时间：2026-10-30；作者：lq。
  useEffect(() => () => { document.body.style.userSelect = ''; }, []);

  const finish = () => {
    if (!dragRef.current) return;
    dragRef.current = null;
    document.body.style.userSelect = '';
    commitRef.current?.(sizeRef.current);
  };

  const onPointerDown = (axis: ResizeAxis) => (event: ReactPointerEvent<HTMLElement>) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const rect = event.currentTarget.parentElement?.getBoundingClientRect();
    const current = sizeRef.current;
    dragRef.current = {
      axis,
      pointerX: event.clientX,
      pointerY: event.clientY,
      left: rect?.left ?? left,
      top: rect?.top ?? top,
      width: rect?.width || current.width,
      height: rect?.height || current.height || minHeight,
    };
    event.currentTarget.setPointerCapture?.(event.pointerId);
    document.body.style.userSelect = 'none';
  };

  const onPointerMove = (event: ReactPointerEvent<HTMLElement>) => {
    const origin = dragRef.current;
    if (!origin) return;
    const dx = event.clientX - origin.pointerX;
    const dy = event.clientY - origin.pointerY;
    const { axis } = origin;
    const affectsX = axis.includes('e') || axis.includes('w');
    const affectsY = axis.includes('n') || axis.includes('s');

    let nextLeft = origin.left;
    let width = sizeRef.current.width;
    if (affectsX) {
      const wanted = axis.includes('w') ? origin.width - dx : origin.width + dx;
      if (axis.includes('w')) {
        width = clamp(wanted, minWidth, origin.left + origin.width - MARGIN);
        nextLeft = origin.left + origin.width - width;
      } else {
        width = clamp(wanted, minWidth, window.innerWidth - MARGIN - origin.left);
      }
    }

    let nextTop = origin.top;
    let height = sizeRef.current.height;
    if (affectsY) {
      const wanted = axis.includes('n') ? origin.height - dy : origin.height + dy;
      const maxHeight = window.innerHeight - MARGIN * 2;
      if (axis.includes('n')) {
        height = clamp(wanted, minHeight, Math.min(maxHeight, origin.top + origin.height - MARGIN));
        nextTop = origin.top + origin.height - height;
      } else {
        height = clamp(wanted, minHeight, Math.min(maxHeight, window.innerHeight - MARGIN - origin.top));
      }
    }

    const next: PopoverSize = { width, height };
    sizeRef.current = next;
    setSize(next);
    setLeft(nextLeft);
    setTop(nextTop);
  };

  return {
    size,
    left,
    top,
    handleProps: (axis: ResizeAxis) => ({
      onPointerDown: onPointerDown(axis),
      onPointerMove,
      onPointerUp: finish,
      onPointerCancel: finish,
      onLostPointerCapture: finish,
    }),
  };
}
