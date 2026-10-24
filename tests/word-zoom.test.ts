/**
 * 功能：验证 Word 平台快捷键、滚轮、缩放边界及焦点/选区保护。
 * 时间：2026-10-24；作者：lq。
 */
import { act, createElement, useRef, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { normalizeWordZoom, useWordZoomShortcuts, wordZoomAction } from '../packages/editor-word/src/zoom-shortcuts';

describe('Word 缩放快捷键', () => {
  it('按微软官方定义匹配两平台的加减和 0，兼容数字键盘', () => {
    for (const mac of [true, false]) {
      const mod = mac ? { metaKey: true } : { ctrlKey: true };
      for (const key of ['+', '=']) expect(wordZoomAction(new KeyboardEvent('keydown', { key, ...mod }), mac)).toBe('in');
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: '+', shiftKey: true, ...mod }), mac)).toBe('in');
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: '-', ...mod }), mac)).toBe('out');
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: '0', ...mod }), mac)).toBe('reset');
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: 'Add', code: 'NumpadAdd', ...mod }), mac)).toBe('in');
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: 'Subtract', code: 'NumpadSubtract', ...mod }), mac)).toBe('out');
      for (const extra of [{ altKey: true }, { isComposing: true }, { key: 'b' }, { key: '_' }, { key: '0', shiftKey: true }]) {
        expect(wordZoomAction(new KeyboardEvent('keydown', { key: '+', ...mod, ...extra }), mac)).toBeNull();
      }
      expect(wordZoomAction(new KeyboardEvent('keydown', { key: '+', ...(mac ? { ctrlKey: true } : { metaKey: true }) }), mac)).toBeNull();
    }
  });

  it('边界与浮点修正统一，异常值不污染视图', () => {
    expect(normalizeWordZoom(0.3)).toBe(0.4);
    expect(normalizeWordZoom(3)).toBe(2.5);
    expect(normalizeWordZoom(1.20000000002)).toBe(1.2);
    expect(normalizeWordZoom(NaN)).toBe(1);
  });
});

describe('Word 文档缩放交互', () => {
  let host: HTMLDivElement;
  let root: Root;
  let content: HTMLDivElement;
  let scroll: HTMLDivElement;
  let platform = 'MacIntel';
  const changed = vi.fn();

  // 功能：以受控文档视图模拟真实缩放，验证事件取消和监听器生命周期；时间：2026-10-24；作者：lq。
  function Harness({ enabled = true, readonly = false }: { enabled?: boolean; readonly?: boolean }) {
    const editorRef = useRef<HTMLDivElement>(null);
    const scrollRef = useRef<HTMLDivElement>(null);
    const [zoom, setZoom] = useState(1);
    useWordZoomShortcuts({ editorRef, scrollRef, zoom, enabled, onZoomChange: (next) => { changed(next); setZoom(next); } });
    return createElement('div', { ref: editorRef },
      createElement('output', null, Math.round(zoom * 100)),
      createElement('div', { ref: scrollRef, 'data-scroll': true },
        createElement('div', { 'data-content': true, tabIndex: 0, contentEditable: !readonly, suppressContentEditableWarning: true }, '测试文档'),
      ),
    );
  }
  function mount(options: { enabled?: boolean; readonly?: boolean } = {}) {
    act(() => root.render(createElement(Harness, options)));
    content = host.querySelector('[data-content]')!;
    scroll = host.querySelector('[data-scroll]')!;
    content.focus();
  }
  function key(name: string, target: EventTarget = content, extra: KeyboardEventInit = {}) {
    const event = new KeyboardEvent('keydown', { key: name, metaKey: true, bubbles: true, cancelable: true, ...extra });
    act(() => target.dispatchEvent(event));
    return event;
  }
  function wheel(deltaY: number, extra: WheelEventInit = {}) {
    const event = new WheelEvent('wheel', { deltaY, metaKey: true, bubbles: true, cancelable: true, ...extra });
    act(() => content.dispatchEvent(event));
    return event;
  }
  function percent() { return host.querySelector('output')!.textContent; }

  beforeEach(() => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    platform = 'MacIntel';
    vi.spyOn(navigator, 'platform', 'get').mockImplementation(() => platform);
    changed.mockClear(); host = document.createElement('div'); document.body.append(host); root = createRoot(host);
  });
  afterEach(() => { act(() => root.unmount()); host.remove(); window.getSelection()?.removeAllRanges(); vi.restoreAllMocks(); });

  it('缩放文档时取消浏览器缩放，保留正文、焦点和文字选区', () => {
    mount();
    const range = document.createRange(); range.setStart(content.firstChild!, 1); range.setEnd(content.firstChild!, 3);
    window.getSelection()!.removeAllRanges(); window.getSelection()!.addRange(range);
    expect(key('=').defaultPrevented).toBe(true); expect(percent()).toBe('110');
    expect(key('+').defaultPrevented).toBe(true); expect(percent()).toBe('120');
    expect(key('-').defaultPrevented).toBe(true); expect(percent()).toBe('110');
    key('0'); expect(percent()).toBe('100');
    expect(document.activeElement).toBe(content);
    expect(window.getSelection()!.toString()).toBe('试文');
    expect(content.textContent).toBe('测试文档');
  });

  it('Windows 使用 Ctrl 缩放，Mac 的 Command 不会被错误匹配', () => {
    platform = 'Win32';
    mount();
    expect(key('+').defaultPrevented).toBe(false);
    key('+', content, { ctrlKey: true, metaKey: false }); expect(percent()).toBe('110');
    key('0', content, { ctrlKey: true, metaKey: false }); expect(percent()).toBe('100');
  });

  it('连续快捷键累计并限制在 40% 到 250%，到达边界仍取消默认行为', () => {
    mount();
    act(() => { for (let i = 0; i < 30; i += 1) content.dispatchEvent(new KeyboardEvent('keydown', { key: '+', metaKey: true, bubbles: true, cancelable: true })); });
    expect(percent()).toBe('250'); expect(key('+').defaultPrevented).toBe(true);
    for (let i = 0; i < 30; i += 1) key('-');
    expect(percent()).toBe('40'); key('0'); expect(percent()).toBe('100');
  });

  it('修饰键滚轮和触控板缩放，普通滚动不受影响', () => {
    mount();
    expect(wheel(-100).defaultPrevented).toBe(true); expect(percent()).toBe('110');
    wheel(100); expect(percent()).toBe('100');
    wheel(-16, { metaKey: false, ctrlKey: true }); expect(percent()).toBe('100');
    wheel(-16, { metaKey: false, ctrlKey: true }); expect(percent()).toBe('110');
    expect(wheel(100, { metaKey: false }).defaultPrevented).toBe(false); expect(percent()).toBe('110');
  });

  it('只读正文仍可缩放；工作区外、弹窗和暂停状态保留原有行为', () => {
    mount({ readonly: true }); key('+'); expect(percent()).toBe('110');
    const external = document.createElement('input'); document.body.append(external);
    try { expect(key('+', external).defaultPrevented).toBe(false); } finally { external.remove(); }
    const dialog = document.createElement('div'); dialog.setAttribute('role', 'dialog'); dialog.setAttribute('aria-modal', 'true'); document.body.append(dialog);
    try { expect(key('+').defaultPrevented).toBe(false); expect(wheel(-100).defaultPrevented).toBe(false); } finally { dialog.remove(); }
    mount({ enabled: false }); expect(key('+').defaultPrevented).toBe(false); expect(wheel(-100).defaultPrevented).toBe(false);
    expect(percent()).toBe('110');
  });

  it('正文尚未聚焦时也可缩放，组件卸载后移除监听器', () => {
    mount(); expect(key('+', document.body).defaultPrevented).toBe(true); expect(percent()).toBe('110');
    act(() => root.render(null)); changed.mockClear();
    expect(key('+', document.body).defaultPrevented).toBe(false); expect(changed).not.toHaveBeenCalled();
  });
});
