import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createBlankWorkbook, useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { StatusBar } from '../packages/editor-shell/src/chrome';

const initialState = useWorkspaceStore.getState();
let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  useWorkspaceStore.setState(initialState);
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
});

afterEach(() => {
  act(() => root.unmount());
  container.remove();
  vi.restoreAllMocks();
  useWorkspaceStore.setState(initialState);
});

function session(overrides: Partial<DocumentSession> = {}): DocumentSession {
  return {
    id: 'word-1', kind: 'docx', displayName: '测试.docx', readonly: false,
    capability: { level: 3, warnings: [], features: [], missingFonts: [], scannedAt: 0 },
    state: 'EDIT_READY', data: { stats: { words: 12, characters: 20 } } as DocumentSession['data'],
    dirty: false, isNew: true, restoredFromSnapshot: false, canSaveInPlace: false,
    saveAsCopyRequired: false, sourceSize: 0, createdAt: 0, updatedAt: 0,
    view: { zoom: 1, page: 0, pageCount: 5 },
    ...overrides,
  };
}

function render(active: DocumentSession, other?: DocumentSession) {
  useWorkspaceStore.setState({
    sessions: { [active.id]: active, ...(other ? { [other.id]: other } : {}) },
    activeId: active.id,
  });
  act(() => root.render(createElement(StatusBar)));
}

function button(label: string) {
  return container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;
}

function click(element: HTMLElement) {
  expect(element).toBeTruthy();
  act(() => element.dispatchEvent(new MouseEvent('click', { bubbles: true })));
}

function changeRange(value: string) {
  const range = container.querySelector<HTMLInputElement>('input[type="range"]')!;
  const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  act(() => {
    setValue.call(range, value);
    range.dispatchEvent(new Event('input', { bubbles: true }));
    range.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

function changeSelect(value: string) {
  const select = container.querySelector<HTMLSelectElement>('select[aria-label="缩放比例"]')!;
  act(() => {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

describe('Word 状态栏页数与缩放', () => {
  it('左侧以 0-based view.page 显示页码并同步总页数，省略 Word 冗余信息', () => {
    render(session());
    expect(container.querySelector('[aria-label="文档页数"]')?.textContent).toBe('第 1 页 / 共 5 页');
    expect(container.querySelector('footer > div:first-child')?.firstElementChild?.getAttribute('aria-label')).toBe('文档页数');
    expect(container.textContent).toContain('12 字 · 20 字符');
    for (const label of ['自动恢复', 'UTF-8', '命令', 'Web 模式']) expect(container.textContent).not.toContain(label);
    act(() => useWorkspaceStore.getState().updateView('word-1', { page: 3, pageCount: 9 }));
    expect(container.querySelector('[aria-label="文档页数"]')?.textContent).toBe('第 4 页 / 共 9 页');
  });

  it('未同步页数与缩放时默认为第 1 页和 100%，越界页码显示在页数范围内', () => {
    render(session({ view: {} }));
    expect(container.querySelector('[aria-label="文档页数"]')?.textContent).toBe('第 1 页 / 共 1 页');
    expect(container.querySelector<HTMLSelectElement>('select')?.value).toBe('100');
    act(() => useWorkspaceStore.getState().updateView('word-1', { page: 8, pageCount: 3 }));
    expect(container.querySelector('[aria-label="文档页数"]')?.textContent).toBe('第 3 页 / 共 3 页');
  });

  it('最右侧按适应宽度、减号、滑块、加号、比例顺序提供中文可访问控件', () => {
    render(session({ view: { zoom: 1.23, page: 0, pageCount: 5 } }));
    const group = container.querySelector('[role="group"][aria-label="Word 缩放"]')!;
    expect(group).toBe(container.querySelector('footer')?.lastElementChild);
    const controls = Array.from(group.querySelectorAll('button, input, select'));
    expect(controls.map((control) => control.getAttribute('aria-label')))
      .toEqual(['适应宽度', '缩小', '缩放滑块', '放大', '缩放比例']);
    for (const control of controls) expect(control.getAttribute('title')).toMatch(/[\u4e00-\u9fff]/);
    const range = group.querySelector<HTMLInputElement>('input')!;
    expect([range.min, range.max, range.step, range.value]).toEqual(['40', '250', '5', '123']);
    expect(range.getAttribute('aria-valuetext')).toBe('123%');
    const select = group.querySelector<HTMLSelectElement>('select')!;
    expect(select.value).toBe('123');
    expect(select.selectedOptions[0].textContent).toBe('123%');
    expect(Array.from(select.options).map((option) => Number(option.value)))
      .toEqual([40, 50, 75, 90, 100, 110, 123, 125, 150, 175, 200, 250]);
  });

  it('加减按 5% 更新活动会话，不弄脏文档或改动其它会话的 view', () => {
    render(session(), session({ id: 'word-2', view: { zoom: 1.5, page: 2, pageCount: 7 } }));
    click(button('放大'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view).toEqual({ zoom: 1.05, page: 0, pageCount: 5 });
    click(button('缩小'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1);
    expect(useWorkspaceStore.getState().sessions['word-1'].dirty).toBe(false);
    expect(useWorkspaceStore.getState().sessions['word-2'].view).toEqual({ zoom: 1.5, page: 2, pageCount: 7 });
  });

  it('滑块与比例下拉均调用 updateView 并与外部缩放同步', () => {
    render(session());
    changeRange('165');
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1.65);
    expect(container.querySelector<HTMLSelectElement>('select')?.value).toBe('165');
    changeSelect('75');
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(0.75);
    expect(container.querySelector<HTMLInputElement>('input[type="range"]')?.value).toBe('75');
    act(() => useWorkspaceStore.getState().updateView('word-1', { zoom: 1.17 }));
    expect(container.querySelector<HTMLSelectElement>('select')?.value).toBe('117');
    expect(container.querySelector<HTMLInputElement>('input[type="range"]')?.value).toBe('117');
  });

  it('从任意比例增减时消除浮点误差并限制在 40% 到 250%', () => {
    render(session({ view: { zoom: 0.42 } }));
    click(button('缩小'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(0.4);
    expect(button('缩小').disabled).toBe(true);
    click(button('缩小'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(0.4);
    act(() => useWorkspaceStore.getState().updateView('word-1', { zoom: 2.48 }));
    click(button('放大'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(2.5);
    expect(button('放大').disabled).toBe(true);
    act(() => useWorkspaceStore.getState().updateView('word-1', { zoom: 1.17 }));
    click(button('放大'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1.22);
  });

  it.each([[0.1, '40'], [3, '250'], [Number.NaN, '100'], [Number.POSITIVE_INFINITY, '100']])
    ('显示外部异常缩放 %s 时本地归一化为 %s%%', (zoom, expected) => {
      render(session({ view: { zoom } }));
      expect(container.querySelector<HTMLSelectElement>('select')?.value).toBe(expected);
    });

  it('适应宽度向 window 派发包含当前 sessionId 的事件，不直接猜测缩放值', () => {
    render(session());
    const listener = vi.fn();
    window.addEventListener('mini-office:word-fit-width', listener);
    try {
      click(button('适应宽度'));
      expect(listener).toHaveBeenCalledOnce();
      const event = listener.mock.calls[0][0] as CustomEvent;
      expect(event.detail).toEqual({ sessionId: 'word-1' });
      expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1);
      act(() => useWorkspaceStore.setState({ sessions: { 'word-2': session({ id: 'word-2' }) }, activeId: 'word-2' }));
      click(button('适应宽度'));
      expect((listener.mock.calls[1][0] as CustomEvent).detail).toEqual({ sessionId: 'word-2' });
    } finally {
      window.removeEventListener('mini-office:word-fit-width', listener);
    }
  });

  it('只读 Word 仍能缩放和适应宽度，按钮鼠标按下不抢走编辑焦点', () => {
    render(session({ readonly: true }));
    expect(button('适应宽度').disabled).toBe(false);
    expect(button('放大').disabled).toBe(false);
    for (const label of ['适应宽度', '缩小', '放大']) {
      const event = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
      act(() => button(label).dispatchEvent(event));
      expect(event.defaultPrevented).toBe(true);
    }
    click(button('放大'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1.05);
    expect(container.textContent).toContain('只读保真模式');
  });

  it('Excel 与 Word 一样把缩放放在右下角，但不显示 Word 专属的页码与适应宽度', () => {
    render(session({ kind: 'xlsx', data: createBlankWorkbook() }));
    expect(container.querySelector('[aria-label="工作表缩放"]')).not.toBeNull();
    expect(container.querySelector('[aria-label="缩放滑块"]')).not.toBeNull();
    expect(container.querySelector('[aria-label="缩放比例"]')).not.toBeNull();
    expect(container.querySelector('[aria-label="适应宽度"]')).toBeNull();
    expect(container.querySelector('[aria-label="文档页数"]')).toBeNull();
    // 工作表缩放同样写回 session.view.zoom
    click(button('放大'));
    expect(useWorkspaceStore.getState().sessions['word-1'].view.zoom).toBe(1.05);
    act(() => useWorkspaceStore.setState({ activeId: null }));
    expect(container.textContent).toContain('就绪');
    expect(container.textContent).toContain('自动恢复');
    expect(container.querySelector('[aria-label="工作表缩放"]')).toBeNull();
  });
});
