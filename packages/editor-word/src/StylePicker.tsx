/**
 * 功能：工具栏「样式」入口——样式名按钮 + 样式库浮层；时间：2026-10-29；作者：lq。
 *
 * 为什么单独成文件：WordToolbar 已经承载大量排版控件，样式库自带 16 个预设卡片与
 * 样式集列表，逻辑独立；把「触发按钮 + 浮层 + 面板」收在一处，工具栏只负责摆位置。
 */
import { ChevronDown, Type } from 'lucide-react';
import { Tooltip } from '@mini-office/ui';
import { ToolbarPopover } from './ToolbarPopover';
import { StyleGallery, type StyleGalleryProps } from './StyleGallery';

export interface StylePickerProps extends StyleGalleryProps {
  /** 当前段落样式名，未识别时为「自定义格式」。 */
  currentStyleName: string;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function StylePicker({ currentStyleName, open, onOpenChange, readonly, ...gallery }: StylePickerProps) {
  return (
    <ToolbarPopover
      width={452}
      open={open}
      onOpenChange={onOpenChange}
      disabled={readonly}
      trigger={({ toggle }) => (
        <Tooltip label={readonly ? '样式库（只读）' : '样式库：批量套用预设样式'}>
          <button
            type="button"
            aria-haspopup="dialog"
            aria-expanded={open}
            aria-label="样式库"
            disabled={readonly}
            onClick={toggle}
            className="flex h-7 w-[104px] shrink-0 items-center gap-1 rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] px-1.5 text-[11px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
          >
            <Type size={13} className="shrink-0 text-[var(--mo-text-muted)]" />
            <span className="min-w-0 flex-1 truncate text-left">{currentStyleName}</span>
            <ChevronDown size={12} className="shrink-0 text-[var(--mo-text-faint)]" />
          </button>
        </Tooltip>
      )}
    >
      {() => <StyleGallery {...gallery} readonly={readonly} />}
    </ToolbarPopover>
  );
}
