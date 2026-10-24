/**
 * 功能：为公式栏和单元格编辑共用函数智能提示，支持键盘、鼠标和输入法。
 * 时间：2026-10-24；作者：lq。
 */
import {
  forwardRef, useId, useLayoutEffect, useRef, useState,
  type CSSProperties, type InputHTMLAttributes,
} from 'react';
import { createPortal } from 'react-dom';
import { Input } from '@mini-office/ui';
import { applyFormulaCompletion, getFormulaCompletion, type FunctionSuggestion } from './formula/completion';

interface FormulaInputProps extends Omit<InputHTMLAttributes<HTMLInputElement>, 'value' | 'onChange'> {
  value: string;
  onValueChange: (value: string) => void;
  plain?: boolean;
}

export const FormulaInput = forwardRef<HTMLInputElement, FormulaInputProps>(function FormulaInput(
  { value, onValueChange, plain = false, onKeyDown, onFocus, onBlur, onSelect,
    onCompositionStart, onCompositionEnd, ...props }, forwardedRef,
) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const pendingCaret = useRef<number | null>(null);
  const [focused, setFocused] = useState(false);
  const [composing, setComposing] = useState(false);
  const [caret, setCaret] = useState(value.length);
  const [collapsed, setCollapsed] = useState(true);
  const [dismissed, setDismissed] = useState<string | null>(null);
  const [choice, setChoice] = useState({ key: '', index: 0 });
  const [position, setPosition] = useState<CSSProperties>({});
  const listId = useId();
  const context = getFormulaCompletion(value, caret);
  const contextKey = `${value}:${caret}`;
  const open = focused && collapsed && !composing && !props.disabled && !props.readOnly
    && !!context && dismissed !== contextKey;
  const selected = choice.key === contextKey ? Math.min(choice.index, (context?.suggestions.length ?? 1) - 1) : 0;

  // 功能：受控值更新后恢复补全光标，避免在公式中间编辑时跳到末尾；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (pendingCaret.current === null || !inputRef.current) return;
    inputRef.current.setSelectionRange(pendingCaret.current, pendingCaret.current);
    pendingCaret.current = null;
  }, [value]);

  // 功能：弹层挂载至 body，跟随输入框滚动/缩放定位，并在窗口底部自动向上展开；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (!open) return;
    const update = () => {
      const input = inputRef.current;
      if (!input) return;
      const rect = input.getBoundingClientRect();
      const width = Math.min(420, window.innerWidth - 16);
      const below = Math.max(0, window.innerHeight - rect.bottom - 12);
      const above = Math.max(0, rect.top - 12);
      const upwards = below < 250 && above > below;
      setPosition({
        position: 'fixed', zIndex: 1000, width,
        left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)),
        ...(upwards ? { bottom: window.innerHeight - rect.top + 4 } : { top: rect.bottom + 4 }),
        maxHeight: Math.min(330, upwards ? above : below),
      });
    };
    update();
    window.addEventListener('resize', update);
    window.addEventListener('scroll', update, true);
    return () => {
      window.removeEventListener('resize', update);
      window.removeEventListener('scroll', update, true);
    };
  }, [open, value, caret]);

  useLayoutEffect(() => {
    if (open) listRef.current?.querySelector('[aria-selected="true"]')?.scrollIntoView?.({ block: 'nearest' });
  }, [open, selected, contextKey]);

  // 功能：选中函数仅补全文本，保持输入焦点，不触发单元格提交；时间：2026-10-24；作者：lq。
  const accept = (suggestion: FunctionSuggestion) => {
    if (!context) return;
    const next = applyFormulaCompletion(value, context, suggestion);
    pendingCaret.current = next.caret;
    setCaret(next.caret);
    setDismissed(null);
    onValueChange(next.text);
  };
  const Component = plain ? 'input' : Input;

  return (
    <>
      <Component
        {...props}
        ref={(element) => {
          inputRef.current = element;
          if (typeof forwardedRef === 'function') forwardedRef(element);
          else if (forwardedRef) forwardedRef.current = element;
        }}
        value={value}
        autoComplete="off"
        spellCheck={false}
        role="combobox"
        aria-autocomplete="list"
        aria-expanded={open}
        aria-controls={open ? listId : undefined}
        aria-activedescendant={open ? `${listId}-${selected}` : undefined}
        onFocus={(event) => {
          setFocused(true);
          setCaret(event.currentTarget.selectionStart ?? value.length);
          setCollapsed(event.currentTarget.selectionStart === event.currentTarget.selectionEnd);
          setDismissed(null);
          onFocus?.(event);
        }}
        onBlur={(event) => { setFocused(false); onBlur?.(event); }}
        onSelect={(event) => {
          setCaret(event.currentTarget.selectionStart ?? value.length);
          setCollapsed(event.currentTarget.selectionStart === event.currentTarget.selectionEnd);
          onSelect?.(event);
        }}
        onChange={(event) => {
          setCaret(event.currentTarget.selectionStart ?? event.target.value.length);
          setCollapsed(event.currentTarget.selectionStart === event.currentTarget.selectionEnd);
          setDismissed(null);
          onValueChange(event.target.value);
        }}
        onCompositionStart={(event) => { setComposing(true); onCompositionStart?.(event); }}
        onCompositionEnd={(event) => { setComposing(false); onCompositionEnd?.(event); }}
        onKeyDown={(event) => {
          // 功能：输入法确认和带修饰键的原有操作交由输入框处理；时间：2026-10-24；作者：lq。
          if (composing || event.nativeEvent.isComposing || event.keyCode === 229) return;
          if (open && context && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
            if (['ArrowDown', 'ArrowUp', 'Enter', 'Tab', 'Escape'].includes(event.key)) {
              event.preventDefault();
              event.stopPropagation();
              if (event.key === 'Escape') setDismissed(contextKey);
              else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                const delta = event.key === 'ArrowDown' ? 1 : -1;
                setChoice({ key: contextKey, index: (selected + delta + context.suggestions.length) % context.suggestions.length });
              } else accept(context.suggestions[selected]);
              return;
            }
          }
          onKeyDown?.(event);
        }}
      />
      {open && context ? createPortal(
        <div style={position} className="mo-no-print flex flex-col overflow-hidden rounded-lg border border-[var(--mo-line-strong)] bg-[var(--mo-surface)] text-[12px] text-[var(--mo-text-primary)] shadow-[var(--mo-shadow-pop)]">
          <div className="shrink-0 border-b border-[var(--mo-line)] px-3 py-2 text-[11px] text-[var(--mo-text-muted)]">
            函数提示 · ↑↓ 选择 · Tab / 回车补全 · Esc 关闭
          </div>
          <div ref={listRef} id={listId} role="listbox" aria-label="公式函数建议" className="min-h-0 overflow-y-auto p-1">
            {context.suggestions.map((suggestion, index) => (
              <div
                key={suggestion.name}
                id={`${listId}-${index}`}
                role="option"
                aria-selected={selected === index}
                onMouseEnter={() => setChoice({ key: contextKey, index })}
                onMouseDown={(event) => { event.preventDefault(); event.stopPropagation(); }}
                onClick={(event) => { event.stopPropagation(); accept(suggestion); }}
                className={`cursor-pointer rounded-md px-2 py-2 ${selected === index ? 'bg-[var(--mo-accent-soft)]' : ''}`}
              >
                <div className="flex items-baseline gap-3">
                  <span className="font-mono font-semibold text-[var(--mo-accent)]">{suggestion.name}</span>
                  <span className="text-[var(--mo-text-muted)]">{suggestion.description}</span>
                </div>
                <div className="mt-1 break-words font-mono text-[11px] text-[var(--mo-text-muted)]">{suggestion.signature}</div>
              </div>
            ))}
          </div>
        </div>, document.body,
      ) : null}
    </>
  );
});
