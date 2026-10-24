/**
 * 功能：按 Microsoft Word 的平台快捷键缩放文档，拦截浏览器默认缩放并保留编辑焦点。
 * 时间：2026-10-24；作者：lq。
 * 依据：https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-word-95ef89dd-7142-4b50-afb2-f762f663ceb2
 */
import { useEffect, useLayoutEffect, useRef, type RefObject } from 'react';
import { isMac } from '@mini-office/shared';

export type WordZoomAction = 'in' | 'out' | 'reset';
export const WORD_ZOOM_MIN = 0.4;
export const WORD_ZOOM_MAX = 2.5;
const ZOOM_STEP = 0.1;

// 功能：统一键盘、滚轮和工具栏的比例边界，消除连续加减的浮点误差；时间：2026-10-24；作者：lq。
export function normalizeWordZoom(value: number): number {
  return Math.max(WORD_ZOOM_MIN, Math.min(WORD_ZOOM_MAX, Number((Number.isFinite(value) ? value : 1).toFixed(2))));
}

// 功能：匹配 Mac 的 Command 或 Windows 的 Ctrl，兼容主键盘加号、等号和数字键盘；时间：2026-10-24；作者：lq。
export function wordZoomAction(event: KeyboardEvent, mac = isMac()): WordZoomAction | null {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || event.altKey) return null;
  if (mac ? !event.metaKey || event.ctrlKey : !event.ctrlKey || event.metaKey) return null;
  if (event.key === '+' || event.key === '=' || event.code === 'NumpadAdd') return 'in';
  if (!event.shiftKey && (event.key === '-' || event.code === 'NumpadSubtract')) return 'out';
  if (!event.shiftKey && event.key === '0') return 'reset';
  return null;
}

interface WordZoomShortcutsOptions {
  editorRef: RefObject<HTMLDivElement>;
  scrollRef: RefObject<HTMLDivElement>;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  enabled: boolean;
}

export function useWordZoomShortcuts({ editorRef, scrollRef, zoom, onZoomChange, enabled }: WordZoomShortcutsOptions) {
  const latest = useRef({ zoom, onZoomChange });
  useLayoutEffect(() => { latest.current = { zoom, onZoomChange }; }, [zoom, onZoomChange]);

  useEffect(() => {
    if (!enabled) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    let wheelDelta = 0;
    let wheelTime = 0;
    // 功能：连续按键/滚轮同步累积最新比例，不依赖 React 下一次渲染；时间：2026-10-24；作者：lq。
    const apply = (action: WordZoomAction) => {
      const next = normalizeWordZoom(action === 'reset' ? 1 : latest.current.zoom + (action === 'in' ? ZOOM_STEP : -ZOOM_STEP));
      if (next === latest.current.zoom) return;
      latest.current.zoom = next;
      latest.current.onZoomChange(next);
    };
    // 功能：快捷键限于当前 Word 工作区和空焦点，弹窗及其他编辑器保留原有按键；时间：2026-10-24；作者：lq。
    const onKeyDown = (event: KeyboardEvent) => {
      if (document.querySelector('[role="dialog"][aria-modal="true"]')) return;
      const target = event.target;
      if (target instanceof Node && target !== document && target !== document.body && !editorRef.current?.contains(target)) return;
      const action = wordZoomAction(event);
      if (!action) return;
      event.preventDefault();
      event.stopPropagation();
      apply(action);
    };
    // 功能：Word 的修饰键+滚轮缩放；兼容触控板以 Ctrl+wheel 发送的捏合事件，普通滚动保持原行为；时间：2026-10-24；作者：lq。
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
