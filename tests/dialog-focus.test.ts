/** 功能：回归验证弹窗输入不会因父组件更新而失焦，并使用最新关闭回调；时间：2026-10-19；作者：lq。 */
import { act, createElement, useState, type ChangeEvent } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Dialog } from '../packages/ui/src/dialog';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  vi.useFakeTimers();
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
});
afterEach(() => {
  act(() => root.unmount());
  container.remove();
  vi.useRealTimers();
});

/** 功能：复现 PDF 导出表单每次输入产生新 onClose 回调的使用方式；时间：2026-10-19；作者：lq。 */
function ExportForm({ onClose }: { onClose: (name: string) => void }) {
  const [name, setName] = useState('报价.pdf');
  return createElement(Dialog, {
    open: true,
    title: '导出为 PDF',
    onClose: () => onClose(name),
    children: createElement('input', {
      'aria-label': '文件名',
      value: name,
      onChange: (event: ChangeEvent<HTMLInputElement>) => setName(event.target.value),
    }),
  });
}

/** 功能：模拟连续编辑并触发受控表单重渲染及延时焦点处理；时间：2026-10-19；作者：lq。 */
function inputValue(input: HTMLInputElement, value: string) {
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
  act(() => vi.advanceTimersByTime(50));
}

describe('弹窗表单焦点', () => {
  it('文件名连续输入、删除和全量替换后仍保持输入焦点', () => {
    act(() => root.render(createElement(ExportForm, { onClose: vi.fn() })));
    act(() => vi.advanceTimersByTime(50));
    const input = document.querySelector<HTMLInputElement>('[aria-label="文件名"]')!;
    input.focus();
    for (const value of ['报价a.pdf', '报价ab.pdf', '报价a.pdf', '', '卫星影像与瓦片服务报价.pdf']) {
      inputValue(input, value);
      expect(input.value).toBe(value);
      expect(document.activeElement).toBe(input);
    }
  });

  it('输入导致回调更新后，Escape 读取最新文件名并关闭', () => {
    const close = vi.fn();
    act(() => root.render(createElement(ExportForm, { onClose: close })));
    act(() => vi.advanceTimersByTime(50));
    const input = document.querySelector<HTMLInputElement>('[aria-label="文件名"]')!;
    input.focus();
    inputValue(input, '新的报价.pdf');
    act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
    expect(close).toHaveBeenCalledOnce();
    expect(close).toHaveBeenCalledWith('新的报价.pdf');
  });

  it('刚打开时用户提前点击输入框，初始聚焦定时器不会抢走焦点', () => {
    act(() => root.render(createElement(ExportForm, { onClose: vi.fn() })));
    const input = document.querySelector<HTMLInputElement>('[aria-label="文件名"]')!;
    input.focus();
    act(() => vi.advanceTimersByTime(50));
    expect(document.activeElement).toBe(input);
  });
});
