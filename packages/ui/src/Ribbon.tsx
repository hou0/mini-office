import { useState, type ReactNode } from 'react';
import { ArrowUpRight, ChevronDown, ChevronUp } from 'lucide-react';
import { cn } from './primitives';

/**
 * 通用 Ribbon（功能区）外壳。
 *
 * 设计约束：纯展示 + 派发，不持有任何文档状态——标签页切换由本组件内部 state 维护，
 * 按钮动作全部由调用方通过 `RibbonButton.onClick` 注入（与现有扁平工具栏同源）。
 * 这样三个编辑器（Word/Sheet/Slides）只需提供各自的分组描述，不用重复实现 tab 条与分组布局。
 *
 * 主题化：所有外观都走可主题化令牌（`--mo-surface*` / `--mo-line` / `--mo-text-*` /
 * `--mo-accent*`），WPS 这类外观插件改令牌即可整条换肤，无需注入 CSS。
 */
export interface RibbonButton {
  id: string;
  kind?: 'button';
  label: string;
  icon?: ReactNode;
  /** WPS 风格：大号按钮（如"粘贴"），图标与标签更大、占两行高度。 */
  large?: boolean;
  /** 仅显示图标、隐藏下方文字标签（WPS 的"格式刷/增大字号/减小字号"即纯图标）。 */
  hideLabel?: boolean;
  /** 悬停提示里附带的快捷键说明，如 'Ctrl+B'（配合全局 TooltipHost 显示）。 */
  shortcut?: string;
  active?: boolean;
  disabled?: boolean;
  title?: string;
  onClick?: () => void;
  onDoubleClick?: () => void;
}

export interface RibbonInline {
  id: string;
  kind: 'inline';
  content: ReactNode;
}

export type RibbonItem = RibbonButton | RibbonInline;

export interface RibbonGroup {
  id: string;
  label?: string;
  /** 组内排布方向：'row' 横向（默认，适合图标组），'col' 纵向（适合对齐/缩进小图标）。 */
  direction?: 'row' | 'col';
  /** WPS「字体」式双行分组：前 N 个 item 占第一行（如下拉框），其余换行。 */
  firstRowCount?: number;
  /** WPS 风格：分组右下角的对话框启动器箭头，点击打开对应设置面板。 */
  launch?: () => void;
  items: RibbonItem[];
}

export interface RibbonTab {
  id: string;
  label: string;
  groups: RibbonGroup[];
}

export function Ribbon({
  tabs,
  defaultTab,
  /** 功能：标签行左侧插槽——WPS 布局的「文件」按钮与撤销/重做放这里；时间：2026-10-31。 */
  topBarLeft,
  /** 功能：标签行右侧插槽——WPS 布局的查找替换输入区放这里；时间：2026-10-31。 */
  topBarRight,
}: {
  tabs: RibbonTab[];
  defaultTab?: string;
  topBarLeft?: ReactNode;
  topBarRight?: ReactNode;
}) {
  const [active, setActive] = useState(defaultTab ?? tabs[0]?.id);
  const [collapsed, setCollapsed] = useState(false);
  const current = tabs.find((tab) => tab.id === active) ?? tabs[0];
  if (!current) return null;

  // WPS 风格：图标按钮固定 36px，图标区 18px，标签 9px。
  const renderItem = (item: RibbonItem) =>
    item.kind === 'inline' ? (
      <div key={item.id} className="flex items-center">
        {item.content}
      </div>
    ) : (
      <button
        key={item.id}
        type="button"
        title={item.title ?? item.label}
        data-tooltip={item.shortcut ? `${item.title ?? item.label} (${item.shortcut})` : item.title ?? item.label}
        aria-label={item.label}
        disabled={item.disabled}
        data-active={item.active ? 'true' : undefined}
        onClick={item.onClick}
        onDoubleClick={item.onDoubleClick}
        className={cn(
          'flex shrink-0 flex-col items-center justify-center gap-0.5 rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:cursor-not-allowed disabled:opacity-40',
          item.large ? 'h-12 w-12' : 'h-9 w-9',
          item.active && 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
        )}
      >
        {item.icon ? (
          <span className={cn('flex items-center justify-center', item.large ? 'h-6 w-6' : 'h-5 w-5')}>
            {item.icon}
          </span>
        ) : null}
        {!item.hideLabel ? (
          <span className={cn('max-w-[56px] truncate leading-none', item.large ? 'text-[11px]' : 'text-[11px]')}>
            {item.label}
          </span>
        ) : null}
      </button>
    );

  return (
    <div className="mo-glass mo-no-print shrink-0 border-b border-[var(--mo-line)] bg-[var(--mo-surface)]">
      {/* 标签页条：active 标签与功能区同色、底部 accent 边框，模拟 WPS 的"标签连体"效果。 */}
      <div className="flex items-stretch border-b border-[var(--mo-line)] pl-1">
        {topBarLeft ? <div className="flex shrink-0 items-center gap-0.5 pr-2">{topBarLeft}</div> : null}
        {tabs.map((tab) => {
          const isActive = tab.id === current.id;
          return (
            <button
              key={tab.id}
              type="button"
              onClick={() => setActive(tab.id)}
              data-active={isActive ? 'true' : undefined}
              className={cn(
                'relative -mb-px border-b-2 px-3.5 py-1 text-[13px] transition-colors',
                isActive
                  ? 'border-[var(--mo-accent)] bg-[var(--mo-surface)] font-semibold text-[var(--mo-text-primary)]'
                  : 'border-transparent text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]',
              )}
            >
              {tab.label}
            </button>
          );
        })}
        <div className="flex-1" />
        {topBarRight ? <div className="flex shrink-0 items-center gap-1 pr-1">{topBarRight}</div> : null}
        <button
          type="button"
          aria-label={collapsed ? '展开功能区' : '折叠功能区'}
          title={collapsed ? '展开功能区' : '折叠功能区'}
          data-tooltip={collapsed ? '展开功能区' : '折叠功能区'}
          onClick={() => setCollapsed((c) => !c)}
          className="flex w-7 shrink-0 items-center justify-center text-[var(--mo-text-muted)] transition-colors hover:text-[var(--mo-accent)]"
        >
          {collapsed ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
        </button>
      </div>
      {/* 功能区内容：分组等高、右侧细竖线分隔，分组名居中固定在底部；折叠时仅显示标签页。 */}
      {!collapsed ? (
      <div className="mo-scroll flex items-stretch gap-0 overflow-x-auto px-2 py-1.5">
        {current.groups.map((group, index) => {
          const isLast = index === current.groups.length - 1;
          return (
            <div
              key={group.id}
              className="relative flex shrink-0 flex-col pr-3 last:pr-0"
            >
              <div className="flex min-h-[54px] flex-1 flex-col justify-center gap-1">
                {group.firstRowCount != null ? (
                  <>
                    <div className="flex items-center gap-1">
                      {group.items.slice(0, group.firstRowCount).map(renderItem)}
                    </div>
                    <div className="flex flex-wrap items-center gap-1">
                      {group.items.slice(group.firstRowCount).map(renderItem)}
                    </div>
                  </>
                ) : (
                  <div
                    className={cn(
                      'flex flex-wrap content-start gap-1',
                      group.direction === 'col' ? 'flex-col' : 'flex-row items-center',
                    )}
                  >
                    {group.items.map(renderItem)}
                  </div>
                )}
              </div>
              {group.label ? (
                <div className="mt-0.5 pr-3 text-center text-[11px] leading-none text-[var(--mo-text-faint)]">
                  {group.label}
                </div>
              ) : null}
              {group.launch ? (
                <button
                  type="button"
                  aria-label={`${group.label ?? ''} 设置`}
                  title={`${group.label ?? ''} 设置`}
                  onClick={group.launch}
                  className="absolute bottom-1 right-1 flex h-4 w-4 items-center justify-center text-[var(--mo-text-faint)] transition-colors hover:text-[var(--mo-accent)]"
                >
                  <ArrowUpRight size={12} />
                </button>
              ) : null}
              {!isLast ? (
                <span className="pointer-events-none absolute inset-y-2.5 right-0 w-px bg-gradient-to-b from-transparent via-[var(--mo-line)] to-transparent" />
              ) : null}
            </div>
          );
        })}
      </div>
      ) : null}
    </div>
  );
}
