import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface TooltipState {
  label: string;
  left: number;
  top: number;
  placement: 'above' | 'below';
}

/**
 * A single portal-based tooltip host for all toolbars. Toolbars scroll
 * horizontally, so ordinary absolutely-positioned tooltips get clipped by
 * `overflow-x-auto`; this host always renders in the body layer.
 */
export function TooltipHost() {
  const [tooltip, setTooltip] = useState<TooltipState | null>(null);
  const timer = useRef<number | null>(null);

  useEffect(() => {
    const clear = () => {
      if (timer.current !== null) {
        window.clearTimeout(timer.current);
        timer.current = null;
      }
    };
    const hide = () => {
      clear();
      setTooltip(null);
    };
    const show = (target: HTMLElement) => {
      clear();
      const label = target.dataset.tooltip?.trim();
      if (!label) return;
      timer.current = window.setTimeout(() => {
        const rect = target.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;
        const above = rect.bottom > window.innerHeight - 72;
        const left = Math.min(
          Math.max(rect.left + rect.width / 2, 60),
          window.innerWidth - 60,
        );
        setTooltip({
          label,
          left,
          top: above ? rect.top - 7 : rect.bottom + 7,
          placement: above ? 'above' : 'below',
        });
      }, 320);
    };

    const onPointerOver = (event: PointerEvent) => {
      const target = (event.target as HTMLElement | null)?.closest<HTMLElement>('[data-tooltip]');
      if (!target) {
        hide();
        return;
      }
      if (event.relatedTarget instanceof Node && target.contains(event.relatedTarget)) return;
      show(target);
    };
    const onPointerOut = (event: PointerEvent) => {
      const target = (event.target as HTMLElement | null)?.closest<HTMLElement>('[data-tooltip]');
      if (!target) return;
      if (event.relatedTarget instanceof Node && target.contains(event.relatedTarget)) return;
      hide();
    };
    const onPointerDown = () => hide();
    const onScroll = () => hide();

    document.addEventListener('pointerover', onPointerOver, true);
    document.addEventListener('pointerout', onPointerOut, true);
    document.addEventListener('pointerdown', onPointerDown, true);
    window.addEventListener('scroll', onScroll, true);
    window.addEventListener('resize', onScroll);
    return () => {
      clear();
      document.removeEventListener('pointerover', onPointerOver, true);
      document.removeEventListener('pointerout', onPointerOut, true);
      document.removeEventListener('pointerdown', onPointerDown, true);
      window.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onScroll);
    };
  }, []);

  if (!tooltip) return null;
  return createPortal(
    <div
      role="tooltip"
      className="pointer-events-none fixed z-[120] max-w-[320px] -translate-x-1/2 whitespace-nowrap rounded-md bg-[var(--mo-text-primary)] px-2 py-1 text-[11px] font-medium text-[var(--mo-surface)] shadow-pop"
      style={{
        left: tooltip.left,
        top: tooltip.top,
        transform: `translate(-50%, ${tooltip.placement === 'above' ? '-100%' : '0'})`,
      }}
    >
      {tooltip.label}
    </div>,
    document.body,
  );
}
