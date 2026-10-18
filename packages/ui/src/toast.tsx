import { useEffect } from 'react';
import { AlertTriangle, CheckCircle2, Info, X, XCircle } from 'lucide-react';
import { cn } from './primitives';

export interface ToastItem {
  id: string;
  level: 'info' | 'success' | 'warning' | 'error';
  title: string;
  message?: string;
}

const ICONS = {
  info: Info,
  success: CheckCircle2,
  warning: AlertTriangle,
  error: XCircle,
};

const TONES = {
  info: 'text-[var(--mo-accent)]',
  success: 'text-[var(--mo-success)]',
  warning: 'text-[var(--mo-warning)]',
  error: 'text-[var(--mo-danger)]',
};

export function Toast({
  toast,
  onDismiss,
  duration = 4600,
}: {
  toast: ToastItem;
  onDismiss: (id: string) => void;
  duration?: number;
}) {
  useEffect(() => {
    const timer = setTimeout(() => onDismiss(toast.id), duration);
    return () => clearTimeout(timer);
  }, [toast.id, duration, onDismiss]);

  const Icon = ICONS[toast.level];
  return (
    <div
      role="status"
      className="mo-glass mo-pop-in pointer-events-auto flex w-[330px] items-start gap-2.5 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] px-3 py-2.5 shadow-pop"
    >
      <Icon size={16} className={cn('mt-0.5 shrink-0', TONES[toast.level])} />
      <div className="min-w-0 flex-1">
        <div className="text-[12px] font-medium leading-5 text-[var(--mo-text-primary)]">{toast.title}</div>
        {toast.message ? (
          <div className="mt-0.5 break-words text-[11px] leading-4 text-[var(--mo-text-muted)]">{toast.message}</div>
        ) : null}
      </div>
      <button
        type="button"
        aria-label="关闭通知"
        onClick={() => onDismiss(toast.id)}
        className="mt-0.5 rounded p-0.5 text-[var(--mo-text-faint)] hover:bg-[var(--mo-surface-2)]"
      >
        <X size={13} />
      </button>
    </div>
  );
}

export function ToastViewport({
  toasts,
  onDismiss,
}: {
  toasts: ToastItem[];
  onDismiss: (id: string) => void;
}) {
  return (
    <div className="pointer-events-none fixed bottom-9 right-4 z-[100] flex flex-col items-end gap-2">
      {toasts.map((toast) => (
        <Toast key={toast.id} toast={toast} onDismiss={onDismiss} />
      ))}
    </div>
  );
}
