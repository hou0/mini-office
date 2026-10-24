/**
 * 功能：为 Excel 工作表提供与 Word/Excel 原版一致的快捷键和修饰键滚轮缩放。
 * 时间：2026-10-24；作者：lq。
 */
import { useEffect, useLayoutEffect, useRef, type RefObject } from 'react';
import { isMac } from '@mini-office/shared';

export type SheetZoomAction = 'in' | 'out' | 'reset';
export const SHEET_ZOOM_MIN = 0.4;
export const SHEET_ZOOM_MAX = 2.5;

export function normalizeSheetZoom(value: number): number {
  return Math.max(SHEET_ZOOM_MIN, Math.min(SHEET_ZOOM_MAX, Number((Number.isFinite(value) ? value : 1).toFixed(2))));
}

export function sheetZoomAction(event: KeyboardEvent, mac = isMac()): SheetZoomAction | null {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || event.altKey) return null;
  if (mac ? !event.metaKey || event.ctrlKey : !event.ctrlKey || event.metaKey) return null;
  if (event.key === '+' || event.key === '=' || event.code === 'NumpadAdd') return 'in';
  if (!event.shiftKey && (event.key === '-' || event.code === 'NumpadSubtract')) return 'out';
  if (!event.shiftKey && event.key === '0') return 'reset';
  return null;
}

interface SheetZoomShortcutOptions {
  editorRef: RefObject<HTMLDivElement>;
  scrollRef: RefObject<HTMLDivElement>;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  enabled: boolean;
}

export function useSheetZoomShortcuts({ editorRef, scrollRef, zoom, onZoomChange, enabled }: SheetZoomShortcutOptions) {
  const latest = useRef({ zoom, onZoomChange });
  useLayoutEffect(() => { latest.current = { zoom, onZoomChange }; }, [zoom, onZoomChange]);
  useEffect(() => {
    if (!enabled) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    let wheelDelta = 0;
    let wheelTime = 0;
    // 功能：统一键盘和滚轮的缩放步长，连续操作不等待 React 重绘；时间：2026-10-24；作者：lq。
    const apply = (action: SheetZoomAction) => {
      const next = normalizeSheetZoom(action === 'reset' ? 1 : latest.current.zoom + (action === 'in' ? 0.1 : -0.1));
      if (next === latest.current.zoom) return;
      latest.current.zoom = next;
      latest.current.onZoomChange(next);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (document.querySelector('[role="dialog"][aria-modal="true"]')) return;
      const target = event.target;
      if (target instanceof Node && target !== document && target !== document.body && !editorRef.current?.contains(target)) return;
      const action = sheetZoomAction(event);
      if (!action) return;
      event.preventDefault();
      event.stopPropagation();
      apply(action);
    };
    // 功能：按住 Command/Ctrl 滚轮缩放工作表，普通滚动继续用于移动表格；时间：2026-10-24；作者：lq。
    const onWheel = (event: WheelEvent) => {
      if (event.defaultPrevented || event.altKey || event.shiftKey || !event.cancelable || (!event.ctrlKey && !(isMac() && event.metaKey))) return;
      if (document.querySelector('[role="dialog"][aria-modal="true"]')) return;
      event.preventDefault();
      event.stopPropagation();
      if (Date.now() - wheelTime > 180) wheelDelta = 0;
      wheelTime = Date.now();
      const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? scroll.clientHeight : 1);
      if (Math.sign(delta) !== Math.sign(wheelDelta)) wheelDelta = 0;
      wheelDelta += delta;
      if (Math.abs(wheelDelta) < 32) return;
      apply(wheelDelta < 0 ? 'in' : 'out');
      wheelDelta = 0;
    };
    window.addEventListener('keydown', onKeyDown, true);
    scroll.addEventListener('wheel', onWheel, { capture: true, passive: false });
    return () => {
      window.removeEventListener('keydown', onKeyDown, true);
      scroll.removeEventListener('wheel', onWheel, true);
    };
  }, [editorRef, scrollRef, enabled]);
}
