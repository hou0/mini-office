/**
 * 功能：验证函数提示上下文、局部替换、键盘补全和输入法事件的真实交互。
 * 时间：2026-10-24；作者：lq。
 */
import { act, createElement, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FormulaInput } from '../packages/editor-sheet/src/FormulaInput';
import { applyFormulaCompletion, getFormulaCompletion } from '../packages/editor-sheet/src/formula/completion';
import { listFunctions } from '../packages/editor-sheet/src/formula/functions';
import { createBlankWorkbook } from '@mini-office/document-model';
import { SheetGrid, type SheetGridProps } from '../packages/editor-sheet/src/SheetGrid';

describe('公式补全上下文', () => {
  it('从一个字母开始匹配已实现函数，支持小写、点号和数字函数名', () => {
    const context = getFormulaCompletion('=s', 2)!;
    expect(context.suggestions.map((item) => item.name)).toContain('SUM');
    expect(context.suggestions.find((item) => item.name === 'SUM')?.signature).toBe('SUM(数值1, [数值2], …)');
    expect(getFormulaCompletion('=STDEV.', 7)?.suggestions.map((item) => item.name)).toEqual(['STDEV.S', 'STDEV.P']);
    expect(getFormulaCompletion('=LOG1', 5)?.suggestions[0].name).toBe('LOG10');
    for (const { name } of listFunctions()) {
      const suggestion = getFormulaCompletion(`=${name}`, name.length + 1)?.suggestions.find((item) => item.name === name);
      expect(suggestion, name).toBeDefined();
      expect(suggestion?.description, name).not.toBe('工作表函数');
    }
  });

  it('普通文本、字符串、工作表引用、区域引用和未知函数不弹提示', () => {
    for (const text of ['SUM', '=', '=A1', '=SUM(A1:S', '=Sheet1!S', '=S!A1', '=S:A1', '=UNKNOWN', '=IF(A1="S', '=\'S', '=IF(A1="""S']) {
      expect(getFormulaCompletion(text, text.length), text).toBeNull();
    }
    expect(getFormulaCompletion('=SUMIF(A1:A3,"=S",B1:B3)', 15)).toBeNull();
  });

  it('嵌套公式和运算符后补全，已有括号和后半段保持不变', () => {
    const text = '=IF(A1>0,su(A1:A3),0)';
    const context = getFormulaCompletion(text, text.indexOf('su') + 2)!;
    const sum = context.suggestions.find((item) => item.name === 'SUM')!;
    const result = applyFormulaCompletion(text, context, sum);
    expect(result.text).toBe('=IF(A1>0,SUM(A1:A3),0)');
    expect(result.text.slice(0, result.caret)).toBe('=IF(A1>0,SUM(');
    expect(getFormulaCompletion('=A1+S', 5)?.suggestions.some((item) => item.name === 'SUM')).toBe(true);
  });

  it('在函数名中间补全会替换整个名称，零参数函数自动补齐括号', () => {
    const context = getFormulaCompletion('=sum(A1)', 2)!;
    expect(applyFormulaCompletion('=sum(A1)', context, context.suggestions.find((item) => item.name === 'SUM')!).text).toBe('=SUM(A1)');
    const today = getFormulaCompletion('=tod', 4)!;
    expect(applyFormulaCompletion('=tod', today, today.suggestions[0])).toEqual({ text: '=TODAY()', caret: 8 });
  });
});

describe('公式输入交互', () => {
  let host: HTMLDivElement;
  let root: Root;
  let input: HTMLInputElement;
  const commit = vi.fn();

  // 功能：通过受控输入和 DOM 事件验证补全不会意外提交或失去焦点；时间：2026-10-24；作者：lq。
  function mount(initial = '=S', plain = false, readOnly = false) {
    function Harness() {
      const [value, setValue] = useState(initial);
      return createElement(FormulaInput, { value, onValueChange: setValue, onKeyDown: commit, plain, readOnly, 'aria-label': '公式输入' });
    }
    act(() => root.render(createElement(Harness)));
    input = host.querySelector('input')!;
    act(() => { input.focus(); input.setSelectionRange(initial.length, initial.length); input.dispatchEvent(new Event('select', { bubbles: true })); });
  }
  function key(name: string, extra: KeyboardEventInit = {}) {
    act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true, ...extra })));
  }
  beforeEach(() => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    host = document.createElement('div'); document.body.append(host);
    root = createRoot(host); commit.mockClear();
  });
  afterEach(() => { act(() => root.unmount()); host.remove(); });

  it('两种输入样式都支持方向键和 Tab 补全，保留光标且不提交', () => {
    mount('=S', true);
    const options = [...document.querySelectorAll('[role="option"]')];
    const nextName = options[1].querySelector('span')!.textContent;
    key('ArrowDown'); key('Tab');
    expect(input.value).toBe(`=${nextName}(`);
    expect(document.activeElement).toBe(input);
    expect(input.selectionStart).toBe(input.value.length);
    expect(commit).not.toHaveBeenCalled();
    expect(document.querySelector('[role="listbox"]')).toBeNull();
    key('Enter'); expect(commit).toHaveBeenCalledTimes(1);
  });

  it('鼠标选择 SUM 保持输入焦点、不触发 blur 提交', () => {
    mount();
    const option = [...document.querySelectorAll('[role="option"]')].find((item) => item.querySelector('span')?.textContent === 'SUM')!;
    act(() => {
      option.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      option.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    });
    expect(input.value).toBe('=SUM(');
    expect(document.activeElement).toBe(input);
    expect(commit).not.toHaveBeenCalled();
  });

  it('Escape 先关闭列表，再次 Escape 交回编辑器；输入法确认不补全或提交', () => {
    mount(); key('Enter', { isComposing: true });
    expect(input.value).toBe('=S'); expect(commit).not.toHaveBeenCalled();
    key('Escape'); expect(document.querySelector('[role="listbox"]')).toBeNull();
    expect(commit).not.toHaveBeenCalled();
    key('Escape'); expect(commit).toHaveBeenCalledTimes(1);
  });

  it('回车补全零参数函数后，下一次回车才提交；失焦及只读时隐藏提示', () => {
    mount('=TOD'); key('Enter'); expect(input.value).toBe('=TODAY()');
    expect(commit).not.toHaveBeenCalled();
    key('Enter'); expect(commit).toHaveBeenCalledTimes(1);
    act(() => input.blur()); expect(document.querySelector('[role="listbox"]')).toBeNull();
  });

  it('只读输入不显示候选列表', () => {
    mount('=S', false, true);
    expect(document.querySelector('[role="listbox"]')).toBeNull();
  });

  // 功能：回归单元格首字符初始化的焦点/光标竞态，无需等待 20ms 定时器即可继续输入；时间：2026-10-24；作者：lq。
  it('单元格开始输入等号后立即聚焦在末尾，后续输入不会被重新定位', () => {
    const workbook = createBlankWorkbook();
    const props: SheetGridProps = {
      workbook, sheet: workbook.sheets[0],
      selection: { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } },
      editing: { row: 0, col: 0, initialText: '=' },
      onSelectionChange: vi.fn(), onStartEdit: vi.fn(), onCommitEdit: vi.fn(),
      onCancelEdit: vi.fn(), onColumnResize: vi.fn(), onRowResize: vi.fn(),
      onCellContextMenu: vi.fn(), onHeaderContextMenu: vi.fn(), onAction: vi.fn(), onFill: vi.fn(),
    };
    act(() => root.render(createElement(SheetGrid, props)));
    input = host.querySelector('input')!;
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('=');
    expect(input.selectionStart).toBe(1);
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    act(() => {
      setValue.call(input, '=S');
      input.setSelectionRange(2, 2);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(input.selectionStart).toBe(2);
    expect(document.querySelector('[role="listbox"]')).not.toBeNull();
    key('Tab');
    expect(input.value).toBe('=SUM(');
    expect(props.onCommitEdit).not.toHaveBeenCalled();
  });
});
