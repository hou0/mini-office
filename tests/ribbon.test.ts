/**
 * 功能：验证方案 A 的 Ribbon 功能——宿主内置功能区，由 chrome.toolbar='ribbon' 触发。
 * 通过真实挂载 Ribbon 与 WordToolbar（jsdom）断言标签页渲染与切换，证明插件只需翻转开关即可生效。
 * 时间：2026-10-31；作者：lq
 */
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Ribbon } from '../packages/ui/src/Ribbon';
import { WordToolbar, type WordToolbarProps } from '../packages/editor-word/src/WordToolbar';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
});
afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

const fn = () => vi.fn();
function makeProps(chromeToolbar: 'modern' | 'classic-two-row' | 'ribbon'): WordToolbarProps {
  return {
    doc: {
      defaultFont: '默认',
      defaultFontSizePt: 11,
      page: {
        widthPx: 794,
        heightPx: 1123,
        orientation: 'portrait',
        margins: { top: 96, bottom: 96, left: 96, right: 96, header: 48, footer: 48, gutter: 0 },
        columns: 1,
        sizeName: 'A4',
      },
    } as unknown as WordToolbarProps['doc'],
    readonly: false,
    activeStyle: {},
    canUndo: false,
    canRedo: false,
    zoom: 100,
    pageCount: 1,
    currentPage: 1,
    findOpen: false,
    propertiesOpen: false,
    showRuler: true,
    formatPainterActive: false,
    systemFonts: [],
    onInlineCommand: fn(),
    onColorFormat: fn(),
    onFontFormat: fn(),
    onParagraphFormat: fn(),
    onInsertTable: fn(),
    onInsertImage: fn(),
    onPageBreak: fn(),
    onHeaderFooter: fn(),
    onMailMerge: fn(),
    onUndo: fn(),
    onRedo: fn(),
    onZoomChange: fn(),
    onFitWidth: fn(),
    onToggleFind: fn(),
    onToggleProperties: fn(),
    onToggleRuler: fn(),
    onExport: fn(),
    onPrint: fn(),
    onShowCompatibility: fn(),
    onAddNumbering: fn(),
    onAddBullets: fn(),
    onFormatPainter: fn(),
    onOpenParagraphDialog: fn(),
    chromeToolbar,
  } as unknown as WordToolbarProps;
}

describe('Ribbon 通用外壳', () => {
  it('渲染标签页，且仅展示当前标签页的分组', () => {
    const tabs = [
      { id: 'a', label: 'A', groups: [{ id: 'g1', items: [{ id: 'x', label: 'X', onClick: () => {} }] }] },
      { id: 'b', label: 'B', groups: [{ id: 'g2', items: [{ id: 'y', label: 'Y', onClick: () => {} }] }] },
    ];
    act(() => root.render(createElement(Ribbon, { tabs })));
    expect(container.textContent).toContain('A');
    expect(container.textContent).toContain('X');
    expect(container.textContent).not.toContain('Y');
  });

  it('点击标签页切换可见分组', () => {
    const tabs = [
      { id: 'a', label: 'A', groups: [{ id: 'g1', items: [{ id: 'x', label: 'X', onClick: () => {} }] }] },
      { id: 'b', label: 'B', groups: [{ id: 'g2', items: [{ id: 'y', label: 'Y', onClick: () => {} }] }] },
    ];
    act(() => root.render(createElement(Ribbon, { tabs })));
    const tabB = [...container.querySelectorAll('button')].find((b) => b.textContent === 'B')!;
    act(() => tabB.dispatchEvent(new MouseEvent('click', { bubbles: true })));
    expect(container.textContent).toContain('Y');
    expect(container.textContent).not.toContain('X');
  });
});

describe('WordToolbar 随 chrome.toolbar 分支', () => {
  it("chrome.toolbar='ribbon' 渲染标签页功能区（开始/插入/布局/视图）", () => {
    act(() => root.render(createElement(WordToolbar, makeProps('ribbon'))));
    expect(container.textContent).toContain('开始');
    expect(container.textContent).toContain('插入');
    expect(container.textContent).toContain('页面布局');
    expect(container.textContent).toContain('视图');
    // 功能：验证补充的 WPS 风格按钮已接入——纯图标按钮无文字，改按 aria-label 断言
    const label = (name: string) => container.querySelector(`button[aria-label="${name}"]`);
    expect(label('上标')).toBeTruthy();
    expect(label('下标')).toBeTruthy();
    expect(label('清除格式')).toBeTruthy();
    // 功能：剪贴板三连已接入（粘贴带文字；剪切/复制为纯图标）
    expect(container.textContent).toContain('粘贴');
    expect(label('复制')).toBeTruthy();
  });

  it('ribbon 下点击「插入」切到对应分组，且隐藏非当前标签内容', () => {
    act(() => root.render(createElement(WordToolbar, makeProps('ribbon'))));
    // 撤销/重做在顶栏（图标按钮，aria-label 可查），始终可见
    expect(container.querySelector('[aria-label="撤销"]')).toBeTruthy();
    expect(container.textContent).not.toContain('页眉页脚'); // 仅在「插入」标签
    const insertTab = [...container.querySelectorAll('button')].find((b) => b.textContent === '插入')!;
    act(() => insertTab.dispatchEvent(new MouseEvent('click', { bubbles: true })));
    expect(container.textContent).toContain('页眉页脚');
    expect(container.textContent).not.toContain('项目符号'); // 「开始」标签内容已隐藏
  });

  it("chrome.toolbar='modern' 走扁平工具栏（无标签页），证明分支生效", () => {
    act(() => root.render(createElement(WordToolbar, makeProps('modern'))));
    expect(container.textContent).not.toContain('开始');
  });
});
