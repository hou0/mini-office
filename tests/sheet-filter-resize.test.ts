// 功能：回归筛选浮层四边四角鼠标缩放、锚定和本次运行内尺寸记忆；时间：2026-10-30；作者：lq。
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { ColumnFilterMenu } from '../packages/editor-sheet/src/ColumnFilterMenu';

function renderMenu() {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const host = document.createElement('div');
  document.body.append(host);
  const onClose = vi.fn();
  const root = createRoot(host);
  act(() => root.render(createElement(ColumnFilterMenu, {
    anchor: { left: 100, bottom: 100 },
    title: '部门',
    values: ['研发', '销售', ''],
    selected: new Set(['研发']),
    onSelected: vi.fn(),
    onSort: vi.fn(),
    onApply: vi.fn(),
    onClear: vi.fn(),
    onClose,
  })));
  const dialog = document.body.querySelector('[role="dialog"]') as HTMLElement;
  // 功能：jsdom 没有布局，按当前内联样式模拟浮层实际矩形；时间：2026-10-30；作者：lq。
  dialog.getBoundingClientRect = () => {
    const left = Number.parseFloat(dialog.style.left) || 100;
    const top = Number.parseFloat(dialog.style.top) || 104;
    const width = Number.parseFloat(dialog.style.width) || 280;
    const height = Number.parseFloat(dialog.style.height) || 400;
    return { left, top, width, height, right: left + width, bottom: top + height, x: left, y: top, toJSON: () => ({}) } as DOMRect;
  };
  return { root, host, dialog, onClose };
}

function drag(handle: Element, from: [number, number], to: [number, number]) {
  const fire = (type: string, clientX: number, clientY: number) => act(() => {
    handle.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, clientX, clientY }));
  });
  fire('pointerdown', from[0], from[1]);
  fire('pointermove', to[0], to[1]);
  fire('pointerup', to[0], to[1]);
}

describe('筛选浮层缩放', () => {
  it('四边四角拖动改变尺寸，左/上边保持对侧锚定，并限制最小尺寸', () => {
    const { dialog, onClose } = renderMenu();
    expect(dialog.style.width).toBe('280px');
    expect(dialog.style.height).toBe('');
    expect(dialog.style.left).toBe('100px');
    expect(dialog.style.top).toBe('104px');

    // 右下角：宽高同时增加。
    drag(dialog.querySelector('[data-resize-axis="se"]')!, [380, 504], [440, 534]);
    expect(dialog.style.width).toBe('340px');
    expect(dialog.style.height).toBe('430px');
    expect(dialog.style.left).toBe('100px');
    expect(dialog.style.top).toBe('104px');

    // 左边：右侧锚定不动，left 向左移动。
    drag(dialog.querySelector('[data-resize-axis="w"]')!, [100, 300], [60, 300]);
    expect(dialog.style.width).toBe('380px');
    expect(dialog.style.left).toBe('60px');

    // 上边：下侧锚定不动，top 向下移动。
    drag(dialog.querySelector('[data-resize-axis="n"]')!, [200, 104], [200, 154]);
    expect(dialog.style.height).toBe('380px');
    expect(dialog.style.top).toBe('154px');

    // 右下角反向拖过头：停在最小 240 × 220。
    drag(dialog.querySelector('[data-resize-axis="se"]')!, [500, 500], [10, 10]);
    expect(dialog.style.width).toBe('240px');
    expect(dialog.style.height).toBe('220px');
    expect(onClose).not.toHaveBeenCalled();
  });

  it('重新打开时沿用本次运行内最后一次尺寸', () => {
    const { dialog } = renderMenu();
    expect(dialog.style.width).toBe('240px');
    expect(dialog.style.height).toBe('220px');
  });
});
