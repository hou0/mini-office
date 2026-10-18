import clsx, { type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import {
  forwardRef,
  type ButtonHTMLAttributes,
  type InputHTMLAttributes,
  type ReactNode,
  type SelectHTMLAttributes,
} from 'react';

export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}

export type ButtonVariant = 'primary' | 'secondary' | 'ghost' | 'danger' | 'soft';
export type ButtonSize = 'sm' | 'md' | 'icon';

const BUTTON_VARIANTS: Record<ButtonVariant, string> = {
  primary:
    'bg-[var(--mo-accent)] text-[var(--mo-accent-ink)] hover:bg-[var(--mo-accent-hover)] shadow-[0_1px_2px_rgba(16,24,40,0.08)]',
  secondary:
    'bg-[var(--mo-surface)] text-[var(--mo-text-primary)] border border-[var(--mo-line-strong)] hover:bg-[var(--mo-surface-2)]',
  ghost: 'text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]',
  soft: 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)] hover:brightness-95',
  danger: 'bg-[var(--mo-danger)] text-white hover:brightness-110',
};

const BUTTON_SIZES: Record<ButtonSize, string> = {
  sm: 'h-7 px-2.5 text-[12px] rounded-md gap-1.5',
  md: 'h-8 px-3.5 text-[13px] rounded-md gap-2',
  icon: 'h-7 w-7 rounded-md',
};

export interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: ButtonVariant;
  size?: ButtonSize;
  icon?: ReactNode;
  loading?: boolean;
}

export const Button = forwardRef<HTMLButtonElement, ButtonProps>(function Button(
  { className, variant = 'secondary', size = 'md', icon, loading, children, disabled, ...rest },
  ref,
) {
  return (
    <button
      ref={ref}
      type="button"
      disabled={disabled || loading}
      className={cn(
        'inline-flex select-none items-center justify-center whitespace-nowrap font-medium transition-colors duration-fast disabled:cursor-not-allowed disabled:opacity-45',
        BUTTON_VARIANTS[variant],
        BUTTON_SIZES[size],
        className,
      )}
      {...rest}
    >
      {loading ? <Spinner size={14} className="shrink-0" /> : icon ? <span className="shrink-0">{icon}</span> : null}
      {children}
    </button>
  );
});

export interface IconButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  label: string;
  active?: boolean;
  size?: 'sm' | 'md';
}

export const IconButton = forwardRef<HTMLButtonElement, IconButtonProps>(function IconButton(
  { className, label, active, size = 'sm', children, ...rest },
  ref,
) {
  return (
    <button
      ref={ref}
      type="button"
      aria-label={label}
      data-tooltip={label}
      data-active={active ? 'true' : undefined}
      className={cn(
        'inline-flex items-center justify-center rounded-md border border-transparent text-[var(--mo-text-primary)] transition-colors duration-fast hover:bg-[var(--mo-surface-2)] disabled:pointer-events-none disabled:opacity-40',
        size === 'sm' ? 'h-7 w-7' : 'h-8 w-8',
        active && 'border-[color-mix(in_srgb,var(--mo-accent)_26%,transparent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
        className,
      )}
      {...rest}
    >
      {children}
    </button>
  );
});

export const Spinner = ({ size = 16, className }: { size?: number; className?: string }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    className={cn('animate-[mo-spin_0.8s_linear_infinite]', className)}
    aria-hidden
  >
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.22" strokeWidth="2.5" />
    <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" />
  </svg>
);

export const Input = forwardRef<HTMLInputElement, InputHTMLAttributes<HTMLInputElement>>(function Input(
  { className, ...rest },
  ref,
) {
  const label = rest['aria-label'];
  return (
    <input
      ref={ref}
      data-tooltip={typeof label === 'string' ? label : undefined}
      className={cn('mo-field w-full outline-none', className)}
      {...rest}
    />
  );
});

export const Select = forwardRef<HTMLSelectElement, SelectHTMLAttributes<HTMLSelectElement>>(function Select(
  { className, children, ...rest },
  ref,
) {
  const label = rest['aria-label'];
  return (
    <select
      ref={ref}
      data-tooltip={typeof label === 'string' ? label : undefined}
      className={cn('mo-field w-full pr-6 outline-none', className)}
      {...rest}
    >
      {children}
    </select>
  );
});

export const Textarea = forwardRef<HTMLTextAreaElement, React.TextareaHTMLAttributes<HTMLTextAreaElement>>(
  function Textarea({ className, ...rest }, ref) {
    return (
      <textarea
        ref={ref}
        className={cn(
          'w-full resize-none rounded-md border border-[var(--mo-line-strong)] bg-[var(--mo-surface)] px-2.5 py-2 text-[13px] outline-none focus:border-[var(--mo-accent)]',
          className,
        )}
        {...rest}
      />
    );
  },
);

export const Kbd = ({ children, className }: { children: ReactNode; className?: string }) => (
  <kbd
    className={cn(
      'inline-flex h-5 min-w-5 items-center justify-center rounded border border-[var(--mo-line)] bg-[var(--mo-surface-2)] px-1 font-ui text-[10px] font-medium text-[var(--mo-text-muted)]',
      className,
    )}
  >
    {children}
  </kbd>
);

export function Tooltip({ label, children }: { label: string; children: ReactNode }) {
  return (
    <span data-tooltip={label} className="inline-flex">
      {children}
    </span>
  );
}

export function Switch({
  checked,
  onChange,
  label,
  disabled,
}: {
  checked: boolean;
  onChange: (checked: boolean) => void;
  label: string;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      role="switch"
      aria-checked={checked}
      aria-label={label}
      disabled={disabled}
      onClick={() => onChange(!checked)}
      className={cn(
        'relative inline-flex h-[22px] w-[38px] shrink-0 items-center rounded-full transition-colors duration-fast disabled:opacity-40',
        checked ? 'bg-[var(--mo-accent)]' : 'bg-[var(--mo-line-strong)]',
      )}
    >
      <span
        className={cn(
          'inline-block h-[18px] w-[18px] transform rounded-full shadow-sm transition-transform duration-fast',
          checked ? 'translate-x-[18px] bg-[var(--mo-accent-ink)]' : 'translate-x-[2px] bg-white',
        )}
      />
    </button>
  );
}

export function SegmentedControl<T extends string>({
  value,
  options,
  onChange,
  className,
  size = 'sm',
}: {
  value: T;
  options: Array<{ value: T; label: ReactNode; title?: string }>;
  onChange: (value: T) => void;
  className?: string;
  size?: 'sm' | 'md';
}) {
  return (
    <div
      role="tablist"
      className={cn(
        'inline-flex items-center gap-0.5 rounded-md bg-[var(--mo-surface-2)] p-0.5',
        size === 'sm' ? 'h-7' : 'h-8',
        className,
      )}
    >
      {options.map((option) => (
        <button
          key={option.value}
          type="button"
          role="tab"
          aria-selected={value === option.value}
          title={option.title}
          onClick={() => onChange(option.value)}
          className={cn(
            'inline-flex h-full items-center justify-center rounded-[5px] px-2.5 text-[12px] font-medium transition-colors duration-fast',
            value === option.value
              ? 'bg-[var(--mo-surface)] text-[var(--mo-text-primary)] shadow-sm'
              : 'text-[var(--mo-text-muted)] hover:text-[var(--mo-text-primary)]',
          )}
        >
          {option.label}
        </button>
      ))}
    </div>
  );
}

export function Field({
  label,
  hint,
  children,
  className,
}: {
  label: string;
  hint?: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    <label className={cn('flex flex-col gap-1.5', className)}>
      <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">{label}</span>
      {children}
      {hint ? <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">{hint}</span> : null}
    </label>
  );
}

export function EmptyState({
  icon,
  title,
  description,
  action,
  className,
}: {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
  className?: string;
}) {
  return (
    <div className={cn('flex flex-col items-center justify-center gap-3 px-6 py-12 text-center', className)}>
      {icon ? (
        <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-[var(--mo-surface-2)] text-[var(--mo-text-muted)]">
          {icon}
        </div>
      ) : null}
      <div className="space-y-1">
        <div className="text-[14px] font-medium text-[var(--mo-text-primary)]">{title}</div>
        {description ? (
          <div className="max-w-sm text-[12px] leading-5 text-[var(--mo-text-muted)]">{description}</div>
        ) : null}
      </div>
      {action}
    </div>
  );
}

export function ProgressBar({ value, className }: { value: number; className?: string }) {
  return (
    <div className={cn('h-1.5 w-full overflow-hidden rounded-full bg-[var(--mo-surface-2)]', className)}>
      <div
        className="h-full rounded-full bg-[var(--mo-accent)] transition-[width] duration-base"
        style={{ width: `${Math.max(0, Math.min(100, value * 100))}%` }}
      />
    </div>
  );
}

export function Divider({ className }: { className?: string }) {
  return <div className={cn('h-px w-full bg-[var(--mo-line)]', className)} />;
}

export function Badge({
  children,
  tone = 'neutral',
  className,
}: {
  children: ReactNode;
  tone?: 'neutral' | 'accent' | 'success' | 'warning' | 'danger';
  className?: string;
}) {
  const tones: Record<string, string> = {
    neutral: 'bg-[var(--mo-surface-2)] text-[var(--mo-text-muted)]',
    accent: 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
    success: 'bg-[var(--mo-success-soft)] text-[var(--mo-success)]',
    warning: 'bg-[var(--mo-warning-soft)] text-[var(--mo-warning)]',
    danger: 'bg-[var(--mo-danger-soft)] text-[var(--mo-danger)]',
  };
  return (
    <span
      className={cn(
        'inline-flex h-5 items-center rounded px-1.5 text-[10px] font-medium leading-none',
        tones[tone],
        className,
      )}
    >
      {children}
    </span>
  );
}
