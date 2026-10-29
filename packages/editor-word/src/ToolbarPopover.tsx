/**
 * 功能：工具栏浮层容器，供字号、行距等组合框复用定位与关闭逻辑；时间：2026-10-29；作者：lq。
 *
 * 为什么不用现成的 Dropdown：下拉菜单只能渲染单行菜单项，字号与行距需要自定义面板
 * （手动输入框 + 分组网格）。这里复用同一套定位/关闭规则，避免两套浮层行为不一致。
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cn } from '@mini-office/ui';

export function ToolbarPopover({
  trigger,
  children,
  width = 208,
  align = 'left',
  disabled,
  open: controlledOpen,
  onOpenChange,
}: {
  trigger: (props: { open: boolean; toggle: () => void }) => ReactNode;
  children: (props: { close: () => void }) => ReactNode;
  width?: number;
  align?: 'left' | 'right';
  disabled?: boolean;
  /** 传入后由外部控制开合，便于命令面板等入口直接打开面板。 */
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
}) {
  const [uncontrolledOpen, setUncontrolledOpen] = useState(false);
  const isControlled = controlledOpen !== undefined;
  const open = isControlled ? controlledOpen : uncontrolledOpen;
  const setOpen = useCallback(
    (next: boolean) => {
      if (!isControlled) setUncontrolledOpen(next);
      onOpenChange?.(next);
    },
    [isControlled, onOpenChange],
  );
  const anchorRef = useRef<HTMLDivElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: 0, top: 0 });

  const close = useCallback(() => setOpen(false), [setOpen]);

  const updatePosition = useCallback(() => {
    const anchor = anchorRef.current;
    if (!anchor) return;
    const rect = anchor.getBoundingClientRect();
    const height = panelRef.current?.offsetHeight ?? 240;
    const margin = 8;
    let left = align === 'right' ? rect.right - width : rect.left;
    left = Math.min(Math.max(margin, left), window.innerWidth - width - margin);
    let top = rect.bottom + 4;
    if (top + height > window.innerHeight - margin) {
      top = Math.max(margin, rect.top - height - 4);
    }
    setPosition({ left, top });
  }, [align, width]);

  useLayoutEffect(() => {
    if (!open) return;
    updatePosition();
    const frame = requestAnimationFrame(updatePosition);
    return () => cancelAnimationFrame(frame);
  }, [open, updatePosition]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (anchorRef.current?.contains(target) || panelRef.current?.contains(target)) return;
      setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false);
    };
    const reposition = () => updatePosition();
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', reposition);
    window.addEventListener('scroll', reposition, true);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', reposition);
      window.removeEventListener('scroll', reposition, true);
    };
  }, [open, updatePosition, setOpen]);

  return (
    <div ref={anchorRef} className="relative inline-flex shrink-0">
      {trigger({ open, toggle: () => (disabled ? undefined : setOpen(!open)) })}
      {open
        ? createPortal(
            <div
              ref={panelRef}
              style={{ left: position.left, top: position.top, width }}
              className={cn(
                'mo-glass mo-pop-in fixed z-[95] overflow-hidden rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-raised)] p-1 shadow-pop',
              )}
            >
              {children({ close })}
            </div>,
            document.body,
          )
        : null}
    </div>
  );
}
