import type { DocParagraph, DocModel, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { FONT_CHOICES } from '@mini-office/design-system';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  Download,
  FileWarning,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  MoreHorizontal,
  Printer,
  Redo2,
  Search,
  Strikethrough,
  Table2,
  Underline,
  Undo2,
  ZoomIn,
  PanelRight,
  Scissors,
} from 'lucide-react';
import { ColorField, Dropdown, IconButton, Kbd, Select, Tooltip } from '@mini-office/ui';
import type { MenuItem } from '@mini-office/ui';

const FONT_SIZES = [9, 10.5, 11, 12, 14, 16, 18, 22, 24, 28, 36, 48];

export interface WordToolbarProps {
  doc: DocModel;
  readonly: boolean;
  activeParagraph?: DocParagraph;
  activeStyle: RunStyle;
  canUndo: boolean;
  canRedo: boolean;
  zoom: number;
  pageCount: number;
  currentPage: number;
  findOpen: boolean;
  propertiesOpen: boolean;
  onInlineCommand: (command: 'bold' | 'italic' | 'underline' | 'strikeThrough') => void;
  onColorFormat: (patch: Partial<RunStyle>, kind: 'color' | 'highlight') => void;
  onFontFormat: (patch: Partial<RunStyle>) => void;
  onParagraphFormat: (patch: Partial<ParagraphStyle>) => void;
  onInsertTable: () => void;
  onInsertImage: () => void;
  onPageBreak: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onZoomChange: (zoom: number) => void;
  onFitWidth: () => void;
  onToggleFind: () => void;
  onToggleProperties: () => void;
  onExport: () => void;
  onPrint: () => void;
  onShowCompatibility: () => void;
  onAddNumbering: () => void;
  onAddBullets: () => void;
  onBeforeInteraction?: () => void;
}

function Divider() {
  return <span className="mx-1 h-4 w-px shrink-0 bg-[var(--mo-line)]" aria-hidden />;
}

export function WordToolbar(props: WordToolbarProps) {
  const { readonly, activeStyle, activeParagraph } = props;
  const paragraphStyle = activeParagraph?.style ?? {};
  const moreItems: MenuItem[] = [
    {
      id: 'page-break',
      label: '插入分页符',
      icon: <Scissors size={14} />,
      disabled: readonly,
      onSelect: props.onPageBreak,
    },
    {
      id: 'print',
      label: '打印',
      icon: <Printer size={14} />,
      onSelect: props.onPrint,
    },
    {
      id: 'export',
      label: '导出 PDF',
      icon: <Download size={14} />,
      onSelect: props.onExport,
    },
    { type: 'separator' },
    {
      id: 'compat',
      label: '兼容性报告',
      icon: <FileWarning size={14} />,
      onSelect: props.onShowCompatibility,
    },
    {
      id: 'properties',
      label: props.propertiesOpen ? '隐藏属性面板' : '显示属性面板',
      icon: <PanelRight size={14} />,
      onSelect: props.onToggleProperties,
    },
  ];

  return (
    <div
      className="mo-glass mo-no-print flex h-11 shrink-0 items-center gap-0.5 overflow-x-auto border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-2"
      onPointerDownCapture={() => props.onBeforeInteraction?.()}
      onMouseDown={(event) => {
        if ((event.target as HTMLElement).closest('button')) event.preventDefault();
      }}
    >
      <Tooltip label={`撤销 ${KbdLabel('Z')}`}>
        <IconButton label="撤销" disabled={!props.canUndo || readonly} onClick={props.onUndo}>
          <Undo2 size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`重做 ${KbdLabel('Shift+Z')}`}>
        <IconButton label="重做" disabled={!props.canRedo || readonly} onClick={props.onRedo}>
          <Redo2 size={15} />
        </IconButton>
      </Tooltip>
      <Divider />
      <Select
        aria-label="字体"
        className="h-7 w-[132px] shrink-0 border-transparent bg-transparent text-[12px] hover:bg-[var(--mo-surface-2)]"
        value={activeStyle.fontFamily ?? '默认'}
        disabled={readonly}
        onChange={(event) => {
          const value = event.target.value;
          props.onFontFormat({ fontFamily: value === '默认' ? undefined : value });
        }}
      >
        {FONT_CHOICES.map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </Select>
      <Select
        aria-label="字号"
        className="h-7 w-[64px] shrink-0 border-transparent bg-transparent text-[12px] hover:bg-[var(--mo-surface-2)]"
        value={String(activeStyle.fontSizePt ?? props.doc.defaultFontSizePt)}
        disabled={readonly}
        onChange={(event) => {
          const value = Number(event.target.value);
          if (value) props.onFontFormat({ fontSizePt: value });
        }}
      >
        {[...new Set([...(activeStyle.fontSizePt ? [activeStyle.fontSizePt] : []), ...FONT_SIZES])]
          .sort((a, b) => a - b)
          .map((size) => (
            <option key={size} value={size}>
              {size}
            </option>
          ))}
      </Select>
      <Divider />
      <Tooltip label={`加粗 ${KbdLabel('B')}`}>
        <IconButton
          label="加粗"
          active={!!activeStyle.bold}
          disabled={readonly}
          onClick={() => props.onInlineCommand('bold')}
        >
          <Bold size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`斜体 ${KbdLabel('I')}`}>
        <IconButton
          label="斜体"
          active={!!activeStyle.italic}
          disabled={readonly}
          onClick={() => props.onInlineCommand('italic')}
        >
          <Italic size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`下划线 ${KbdLabel('U')}`}>
        <IconButton
          label="下划线"
          active={!!activeStyle.underline}
          disabled={readonly}
          onClick={() => props.onInlineCommand('underline')}
        >
          <Underline size={15} />
        </IconButton>
      </Tooltip>
      <IconButton
        label="删除线"
        active={!!activeStyle.strike}
        disabled={readonly}
        onClick={() => props.onInlineCommand('strikeThrough')}
      >
        <Strikethrough size={15} />
      </IconButton>
      <ColorField
        label="字体颜色"
        value={activeStyle.color}
        onChange={(color) => props.onColorFormat({ color }, 'color')}
      />
      <ColorField
        label="突出显示"
        value={activeStyle.highlight}
        onChange={(color) => props.onColorFormat({ highlight: color }, 'highlight')}
      />
      <Divider />
      <IconButton
        label="左对齐"
        active={paragraphStyle.alignment === 'left' || !paragraphStyle.alignment}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'left' })}
      >
        <AlignLeft size={15} />
      </IconButton>
      <IconButton
        label="居中"
        active={paragraphStyle.alignment === 'center'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'center' })}
      >
        <AlignCenter size={15} />
      </IconButton>
      <IconButton
        label="右对齐"
        active={paragraphStyle.alignment === 'right'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'right' })}
      >
        <AlignRight size={15} />
      </IconButton>
      <IconButton
        label="两端对齐"
        active={paragraphStyle.alignment === 'justify'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'justify' })}
      >
        <AlignJustify size={15} />
      </IconButton>
      <Select
        aria-label="行距"
        className="h-7 w-[78px] shrink-0 border-transparent bg-transparent text-[12px] hover:bg-[var(--mo-surface-2)]"
        value={
          paragraphStyle.lineHeightExactPx
            ? `${paragraphStyle.lineHeightExactPx}px`
            : String(paragraphStyle.lineHeight ?? 1.5)
        }
        disabled={readonly}
        onChange={(event) => {
          const value = event.target.value;
          if (value.endsWith('px')) {
            props.onParagraphFormat({ lineHeight: undefined, lineHeightExactPx: parseFloat(value) });
          } else {
            props.onParagraphFormat({ lineHeight: Number(value), lineHeightExactPx: undefined });
          }
        }}
      >
        <option value="1">单倍行距</option>
        <option value="1.15">1.15 倍</option>
        <option value="1.5">1.5 倍</option>
        <option value="2">2 倍</option>
        <option value="24px">固定 18pt</option>
      </Select>
      <IconButton
        label="增加缩进"
        disabled={readonly}
        onClick={() =>
          props.onParagraphFormat({
            indentLeftPx: (paragraphStyle.indentLeftPx ?? 0) + 24,
          })
        }
      >
        <IndentIncrease size={15} />
      </IconButton>
      <IconButton
        label="减少缩进"
        disabled={readonly}
        onClick={() =>
          props.onParagraphFormat({
            indentLeftPx: Math.max(0, (paragraphStyle.indentLeftPx ?? 0) - 24),
          })
        }
      >
        <IndentDecrease size={15} />
      </IconButton>
      <Divider />
      <IconButton label="项目符号" disabled={readonly} onClick={props.onAddBullets}>
        <List size={15} />
      </IconButton>
      <IconButton label="编号" disabled={readonly} onClick={props.onAddNumbering}>
        <span className="text-[11px] font-semibold">1.</span>
      </IconButton>
      <IconButton label="插入表格" disabled={readonly} onClick={props.onInsertTable}>
        <Table2 size={15} />
      </IconButton>
      <IconButton label="插入图片" disabled={readonly} onClick={props.onInsertImage}>
        <ImagePlus size={15} />
      </IconButton>
      <Divider />
      <Tooltip label={`查找替换 ${KbdLabel('F')}`}>
        <IconButton label="查找替换" active={props.findOpen} onClick={props.onToggleFind}>
          <Search size={15} />
        </IconButton>
      </Tooltip>
      <IconButton label="属性面板" active={props.propertiesOpen} onClick={props.onToggleProperties}>
        <PanelRight size={15} />
      </IconButton>
      <Tooltip label="打印 ⌘P">
        <IconButton label="打印" onClick={props.onPrint}>
          <Printer size={15} />
        </IconButton>
      </Tooltip>
      <div className="ml-auto flex shrink-0 items-center gap-1 pl-2">
        <span className="whitespace-nowrap pr-1 text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {props.pageCount} 页 · 第 {props.currentPage} 页
        </span>
        <IconButton label="适应宽度" onClick={props.onFitWidth}>
          <ZoomIn size={15} />
        </IconButton>
        <Select
          aria-label="缩放"
          className="h-7 w-[70px] border-transparent bg-transparent text-[11px]"
          value={String(Math.round(props.zoom * 100))}
          onChange={(event) => props.onZoomChange(Number(event.target.value) / 100)}
        >
          {[50, 75, 90, 100, 110, 125, 150, 175, 200].map((percent) => (
            <option key={percent} value={percent}>
              {percent}%
            </option>
          ))}
        </Select>
        <IconButton
          label={props.readonly ? '只读模式' : '兼容性报告'}
          onClick={props.onShowCompatibility}
        >
          <FileWarning size={15} className={props.readonly ? 'text-[var(--mo-warning)]' : ''} />
        </IconButton>
        <Dropdown
          align="right"
          minWidth={210}
          items={moreItems}
          trigger={({ toggle }) => (
            <IconButton label="更多操作" onClick={toggle}>
              <MoreHorizontal size={16} />
            </IconButton>
          )}
        />
      </div>
    </div>
  );
}

function KbdLabel(key: string): string {
  return navigator.platform.toLowerCase().includes('mac') ? `⌘${key}` : `Ctrl+${key}`;
}
