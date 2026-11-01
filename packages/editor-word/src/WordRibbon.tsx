import type { ReactNode } from 'react';
import type { WordToolbarProps } from './WordToolbar';
import { Ribbon, type RibbonTab } from '@mini-office/ui';
import {
  AlignCenter,
  FileOutput,
  Menu,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AArrowDown,
  AArrowUp,
  Bold,
  ClipboardPaste,
  Columns3,
  Copy,
  FileText,
  FileWarning,
  FlipHorizontal2,
  Highlighter,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  Mails,
  Baseline,
  ChevronDown,
  PanelRight,
  PanelTop,
  Printer,
  Replace,
  Redo2,
  Ruler,
  Scissors,
  Search,
  Strikethrough,
  Subscript,
  Superscript,
  RemoveFormatting,
  Type,
  Underline,
  Undo2,
  MessageSquarePlus,
  FileDiff,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { ColorField, Dropdown, Select } from '@mini-office/ui';
import { mergeFontChoices } from './use-system-fonts';
import { TablePicker } from './TablePicker';
import { FontSizePicker } from './FontSizePicker';
import { LineSpacingPicker } from './LineSpacingPicker';
import { StylePicker } from './StylePicker';

/**
 * Word 功能区（Ribbon）布局。
 *
 * 复用 WordToolbar 的全部动作回调与子组件（样式库/字体/字号/颜色/表格/行距），
 * 只把"扁平一行"重组成"标签页 + 分组"。所有状态仍由编辑器持有，这里不引入新状态。
 */
export function WordRibbon(props: WordToolbarProps) {
  const { readonly, activeStyle, activeParagraph } = props;
  const paragraphStyle = activeParagraph?.style ?? {};
  const disabled = readonly;

  const fontSelect = (
    <span className="relative inline-flex shrink-0 items-center">
      <Select
        aria-label="字体"
        className="h-7 w-[120px] shrink-0 appearance-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] pl-2 pr-6 text-[12px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
        value={activeStyle.fontFamily ?? '默认'}
        disabled={disabled}
        onChange={(event) => {
          const value = event.target.value;
          props.onFontFormat({ fontFamily: value === '默认' ? undefined : value });
        }}
      >
        {mergeFontChoices(props.systemFonts ?? [], activeStyle.fontFamily, props.doc.defaultFont).map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </Select>
      <ChevronDown size={12} className="pointer-events-none absolute right-1.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-muted)]" />
    </span>
  );

  const colorField = (
    <ColorField
      label="字体颜色"
      icon={<Baseline size={16} />}
      indicatorColor={activeStyle.color ?? '#1D2939'}
      disabled={disabled}
      value={activeStyle.color}
      onChange={(color) => props.onColorFormat({ color }, 'color')}
    />
  );

  const highlightField = (
    <ColorField
      label="文本背景颜色（突出显示）"
      icon={<Highlighter size={16} />}
      indicatorColor={activeStyle.highlight ?? 'transparent'}
      disabled={disabled}
      value={activeStyle.highlight}
      onChange={(color) => props.onColorFormat({ highlight: color }, 'highlight')}
    />
  );

  const moreMenu = (
    <Dropdown
      align="right"
      minWidth={210}
      items={[
        { id: 'mail-merge', label: '邮件合并 / 批量报告', icon: <Mails size={14} />, disabled, onSelect: props.onMailMerge },
        { id: 'header-footer', label: '页眉、页脚与页码', icon: <PanelTop size={14} />, disabled, onSelect: props.onHeaderFooter },
        { id: 'page-break', label: '插入分页符', icon: <Scissors size={14} />, disabled, onSelect: props.onPageBreak },
        { id: 'print', label: '打印', icon: <Printer size={14} />, onSelect: props.onPrint },
        { type: 'separator' },
        { id: 'compat', label: '兼容性报告', icon: <FileWarning size={14} />, onSelect: props.onShowCompatibility },
        { id: 'properties', label: props.propertiesOpen ? '隐藏属性面板' : '显示属性面板', icon: <PanelRight size={14} />, onSelect: props.onToggleProperties },
      ]}
      trigger={({ toggle }) => (
        <button
          type="button"
          aria-label="更多操作"
          onClick={toggle}
          className="flex h-9 w-9 shrink-0 flex-col items-center justify-center gap-0.5 rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)]"
        >
          <span className="text-[13px] leading-none">⋯</span>
          <span className="max-w-[44px] truncate text-[9px] leading-none">更多</span>
        </button>
      )}
    />
  );

  const mmToPx = (mm: number) => Math.round((mm / 25.4) * 96);
  const SYMBOLS = ['§', '¶', '•', '†', '‡', '±', '×', '÷', '≈', '≠', '∞', '≤', '≥', '∑', '√', '©', '®', '™', '…', '★', '☆', '◆', '●'];
  const MARGIN_PRESETS: { name: string; mm: number }[] = [
    { name: '普通', mm: 25.4 },
    { name: '窄', mm: 12.7 },
    { name: '适中', mm: 19.05 },
    { name: '宽', mm: 31.75 },
  ];
  const PAPER_SIZES: { name: string; wMm: number; hMm: number }[] = [
    { name: 'A4', wMm: 210, hMm: 297 },
    { name: 'A3', wMm: 297, hMm: 420 },
    { name: 'Letter', wMm: 215.9, hMm: 279.4 },
  ];
  const marginPatch = (mm: number) => {
    const px = mmToPx(mm);
    const m = props.doc.page.margins;
    return { margins: { ...m, top: px, bottom: px, left: px, right: px } };
  };
  const paperPatch = (wMm: number, hMm: number) => {
    const landscape = props.doc.page.orientation === 'landscape';
    return {
      widthPx: mmToPx(landscape ? hMm : wMm),
      heightPx: mmToPx(landscape ? wMm : hMm),
      sizeName: `${wMm}x${hMm}`,
    };
  };

  // 功能：文件菜单（WPS 布局的左上角入口）——导出/打印/兼容性/属性；时间：2026-10-31。
  const fileDropdown = (
    <Dropdown
      align="left"
      minWidth={200}
      items={[
        { id: 'export', label: '导出 PDF', icon: <FileOutput size={14} />, onSelect: props.onExport },
        { id: 'print', label: '打印', icon: <Printer size={14} />, onSelect: props.onPrint },
        { type: 'separator' },
        { id: 'compat', label: '兼容性报告', icon: <FileWarning size={14} />, onSelect: props.onShowCompatibility },
        { id: 'properties', label: props.propertiesOpen ? '隐藏属性面板' : '显示属性面板', icon: <PanelRight size={14} />, onSelect: props.onToggleProperties },
      ]}
      trigger={({ toggle }) => (
        <button
          type="button"
          aria-label="文件"
          onClick={toggle}
          className="flex h-8 shrink-0 items-center gap-1.5 rounded-md px-2 text-[13px] font-medium text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)]"
        >
          <Menu size={15} />
          文件
        </button>
      )}
    />
  );

  const topBarLeft = (
    <>
      {fileDropdown}
      <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" aria-hidden />
      <button
        type="button"
        aria-label="撤销"
        title="撤销 (Ctrl+Z)"
        disabled={!props.canUndo || disabled}
        onClick={props.onUndo}
        className="flex h-8 w-8 items-center justify-center rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Undo2 size={15} />
      </button>
      <button
        type="button"
        aria-label="重做"
        title="重做 (Ctrl+Y)"
        disabled={!props.canRedo || disabled}
        onClick={props.onRedo}
        className="flex h-8 w-8 items-center justify-center rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Redo2 size={15} />
      </button>
    </>
  );

  const topBarRight = props.onToggleFind ? (
    <button
      type="button"
      aria-label="查找替换"
      onClick={props.onToggleFind}
      className="flex h-8 items-center gap-1.5 rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2.5 text-[12px] text-[var(--mo-text-muted)] transition-colors hover:bg-[var(--mo-surface-2)]"
    >
      <Search size={13} />
      查找替换
    </button>
  ) : null;

  const tabs: RibbonTab[] = [
    {
      id: 'home',
      label: '开始',
      groups: [
        {
          id: 'clipboard',
          label: '剪贴板',
          items: [
            { id: 'paste', label: '粘贴', icon: <ClipboardPaste size={18} />, large: true, shortcut: 'Ctrl+V', disabled, onClick: () => props.onPaste?.() },
            { id: 'cut', label: '剪切', icon: <Scissors size={16} />, hideLabel: true, shortcut: 'Ctrl+X', disabled, onClick: () => props.onCut?.() },
            { id: 'copy', label: '复制', icon: <Copy size={16} />, hideLabel: true, shortcut: 'Ctrl+C', disabled, onClick: () => props.onCopy?.() },
            {
              id: 'format-painter',
              label: '格式刷',
              icon: <PaintbrushIcon />,
              hideLabel: true,
              disabled,
              active: props.formatPainterActive,
              onClick: () => props.onFormatPainter(false),
              onDoubleClick: () => props.onFormatPainter(true),
            },
          ],
        },
        {
          id: 'font',
          label: '字体',
          firstRowCount: 4,
          items: [
            { id: 'font-family', kind: 'inline', content: fontSelect },
            {
              id: 'font-size',
              kind: 'inline',
              content: (
                <FontSizePicker
                  value={activeStyle.fontSizePt ?? props.doc.defaultFontSizePt}
                  disabled={disabled}
                  onChange={(pt) => props.onFontFormat({ fontSizePt: pt })}
                />
              ),
            },
            { id: 'size-up', label: '增大字号', icon: <AArrowUp size={16} />, hideLabel: true, disabled, onClick: () => props.onFontFormat({ fontSizePt: Math.round((activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) + 1) }) },
            { id: 'size-down', label: '减小字号', icon: <AArrowDown size={16} />, hideLabel: true, disabled, onClick: () => props.onFontFormat({ fontSizePt: Math.max(1, Math.round((activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) - 1)) }) },
            { id: 'bold', label: '加粗', icon: <Bold size={16} />, hideLabel: true, shortcut: 'Ctrl+B', active: !!activeStyle.bold, disabled, onClick: () => props.onInlineCommand('bold') },
            { id: 'italic', label: '斜体', icon: <Italic size={16} />, hideLabel: true, shortcut: 'Ctrl+I', active: !!activeStyle.italic, disabled, onClick: () => props.onInlineCommand('italic') },
            { id: 'underline', label: '下划线', icon: <Underline size={16} />, hideLabel: true, shortcut: 'Ctrl+U', active: !!activeStyle.underline, disabled, onClick: () => props.onInlineCommand('underline') },
            { id: 'strike', label: '删除线', icon: <Strikethrough size={16} />, hideLabel: true, active: !!activeStyle.strike, disabled, onClick: () => props.onInlineCommand('strikeThrough') },
            { id: 'color', kind: 'inline', content: colorField },
            { id: 'highlight', kind: 'inline', content: highlightField },
            { id: 'superscript', label: '上标', icon: <Superscript size={16} />, hideLabel: true, active: activeStyle.verticalAlign === 'super', disabled, onClick: () => props.onInlineCommand('superscript') },
            { id: 'subscript', label: '下标', icon: <Subscript size={16} />, hideLabel: true, active: activeStyle.verticalAlign === 'sub', disabled, onClick: () => props.onInlineCommand('subscript') },
            { id: 'clear', label: '清除格式', icon: <RemoveFormatting size={16} />, hideLabel: true, disabled, onClick: () => props.onClearFormat?.() },
          ],
        },
        {
          id: 'paragraph',
          label: '段落',
          launch: props.onOpenParagraphDialog,
          firstRowCount: 4,
          items: [
            { id: 'bullets', label: '项目符号', icon: <List size={16} />, hideLabel: true, disabled, onClick: props.onAddBullets },
            { id: 'numbering', label: '编号', icon: <NumberingIcon />, hideLabel: true, disabled, onClick: props.onAddNumbering },
            { id: 'indent-dec', label: '减少缩进', icon: <IndentDecrease size={16} />, hideLabel: true, disabled, onClick: () => props.onParagraphFormat({ indentLeftPx: Math.max(0, (paragraphStyle.indentLeftPx ?? 0) - 24) }) },
            { id: 'indent-inc', label: '增加缩进', icon: <IndentIncrease size={16} />, hideLabel: true, disabled, onClick: () => props.onParagraphFormat({ indentLeftPx: (paragraphStyle.indentLeftPx ?? 0) + 24 }) },
            { id: 'align-left', label: '左对齐', icon: <AlignLeft size={16} />, hideLabel: true, active: paragraphStyle.alignment === 'left' || !paragraphStyle.alignment, disabled, onClick: () => props.onParagraphFormat({ alignment: 'left' }) },
            { id: 'align-center', label: '居中', icon: <AlignCenter size={16} />, hideLabel: true, active: paragraphStyle.alignment === 'center', disabled, onClick: () => props.onParagraphFormat({ alignment: 'center' }) },
            { id: 'align-right', label: '右对齐', icon: <AlignRight size={16} />, hideLabel: true, active: paragraphStyle.alignment === 'right', disabled, onClick: () => props.onParagraphFormat({ alignment: 'right' }) },
            { id: 'align-justify', label: '两端对齐', icon: <AlignJustify size={16} />, hideLabel: true, active: paragraphStyle.alignment === 'justify', disabled, onClick: () => props.onParagraphFormat({ alignment: 'justify' }) },
            {
              id: 'line-spacing',
              kind: 'inline',
              content: (
                <LineSpacingPicker
                  style={paragraphStyle}
                  disabled={disabled}
                  onChange={props.onParagraphFormat}
                  onOpenOptions={props.onOpenParagraphDialog}
                />
              ),
            },
          ],
        },
        {
          id: 'styles',
          label: '样式',
          items: [
            { id: 'style', kind: 'inline', content: props.stylePicker ? <StylePicker {...props.stylePicker} /> : null },
          ],
        },
        {
          id: 'edit',
          label: '编辑',
          items: [
            { id: 'find', label: '查找', icon: <Search size={15} />, active: props.findOpen, onClick: props.onToggleFind },
            { id: 'replace', label: '替换', icon: <Replace size={15} />, onClick: props.onToggleFind },
          ],
        },
      ],
    },
    {
      id: 'insert',
      label: '插入',
      groups: [
        {
          id: 'pages',
          label: '页',
          items: [
            { id: 'page-break', label: '分页符', icon: <Scissors size={15} />, disabled, onClick: props.onPageBreak },
            { id: 'header-footer', label: '页眉页脚', icon: <PanelTop size={15} />, disabled, onClick: props.onHeaderFooter },
          ],
        },
        {
          id: 'tables',
          label: '表格',
          items: [{ id: 'table', kind: 'inline', content: <TablePicker disabled={disabled} onInsert={props.onInsertTable} /> }],
        },
        {
          id: 'illustrations',
          label: '插图',
          items: [{ id: 'image', label: '图片', icon: <ImagePlus size={15} />, disabled, onClick: props.onInsertImage }],
        },
        {
          id: 'mail',
          label: '邮件',
          items: [{ id: 'mail-merge', label: '邮件合并', icon: <Mails size={15} />, disabled, onClick: props.onMailMerge }],
        },
        {
          id: 'text',
          label: '文本',
          items: [
            {
              id: 'symbol',
              kind: 'inline',
              content: (
                <Dropdown
                  minWidth={240}
                  items={SYMBOLS.map((ch) => ({ id: `sym-${ch}`, label: ch, onSelect: () => props.onInsertSymbol?.(ch) }))}
                  trigger={({ toggle }) => <MenuButton icon={<Type size={15} />} label="符号" toggle={toggle} disabled={disabled} />}
                />
              ),
            },
          ],
        },
      ],
    },
    {
      id: 'layout',
      label: '页面布局',
      groups: [
        {
          id: 'page-setup',
          label: '页面设置',
          items: [
            {
              id: 'orientation',
              kind: 'inline',
              content: (
                <Dropdown
                  minWidth={150}
                  items={[
                    { id: 'portrait', label: '纵向', checked: props.doc.page.orientation !== 'landscape', onSelect: () => props.onPageChange?.({ orientation: 'portrait' }) },
                    { id: 'landscape', label: '横向', checked: props.doc.page.orientation === 'landscape', onSelect: () => props.onPageChange?.({ orientation: 'landscape' }) },
                  ]}
                  trigger={({ toggle }) => <MenuButton icon={<FlipHorizontal2 size={15} />} label="方向" toggle={toggle} disabled={disabled} />}
                />
              ),
            },
            {
              id: 'margin',
              kind: 'inline',
              content: (
                <Dropdown
                  minWidth={150}
                  items={MARGIN_PRESETS.map((p) => ({ id: `m-${p.name}`, label: p.name, onSelect: () => props.onPageChange?.(marginPatch(p.mm)) }))}
                  trigger={({ toggle }) => <MenuButton icon={<Ruler size={15} />} label="页边距" toggle={toggle} disabled={disabled} />}
                />
              ),
            },
            {
              id: 'paper',
              kind: 'inline',
              content: (
                <Dropdown
                  minWidth={150}
                  items={PAPER_SIZES.map((p) => ({ id: `p-${p.name}`, label: p.name, onSelect: () => props.onPageChange?.(paperPatch(p.wMm, p.hMm)) }))}
                  trigger={({ toggle }) => <MenuButton icon={<FileText size={15} />} label="纸张" toggle={toggle} disabled={disabled} />}
                />
              ),
            },
            {
              id: 'columns',
              kind: 'inline',
              content: (
                <Dropdown
                  minWidth={150}
                  items={[
                    { id: 'c1', label: '一栏', checked: (props.doc.page.columns ?? 1) === 1, onSelect: () => props.onPageChange?.({ columns: 1 }) },
                    { id: 'c2', label: '两栏', checked: (props.doc.page.columns ?? 1) === 2, onSelect: () => props.onPageChange?.({ columns: 2 }) },
                    { id: 'c3', label: '三栏', checked: (props.doc.page.columns ?? 1) === 3, onSelect: () => props.onPageChange?.({ columns: 3 }) },
                  ]}
                  trigger={({ toggle }) => <MenuButton icon={<Columns3 size={15} />} label="分栏" toggle={toggle} disabled={disabled} />}
                />
              ),
            },
          ],
        },
      ],
    },
    {
      id: 'review',
      label: '审阅',
      groups: [
        {
          id: 'comments',
          label: '批注',
          items: [
            { id: 'add-comment', label: '新建批注', icon: <MessageSquarePlus size={15} />, disabled: disabled || !props.onAddComment, onClick: props.onAddComment },
            { id: 'review-panel', label: '审阅面板', icon: <FileDiff size={15} />, active: props.reviewOpen, disabled: disabled || !props.onToggleReview, onClick: props.onToggleReview },
          ],
        },
        {
          id: 'changes',
          label: '修订',
          items: [
            { id: 'review-hint', kind: 'inline', content: (
              <span className="max-w-[180px] text-[11px] leading-tight text-[var(--mo-text-faint)]">
                打开审阅面板可逐条接受或拒绝修订
              </span>
            ) },
          ],
        },
      ],
    },
    {
      id: 'view',
      label: '视图',
      groups: [
        {
          id: 'show',
          label: '显示',
          items: [
            { id: 'ruler', label: '标尺', icon: <Ruler size={15} />, active: props.showRuler, onClick: props.onToggleRuler },
            { id: 'properties', label: '属性', icon: <PanelRight size={15} />, active: props.propertiesOpen, onClick: props.onToggleProperties },
          ],
        },
        {
          id: 'zoom',
          label: '显示比例',
          items: [
            { id: 'zoom-out', label: '缩小', icon: <ZoomOut size={15} />, onClick: () => props.onZoomChange(Math.max(50, (props.zoom ?? 100) - 10)) },
            { id: 'zoom-reset', label: '100%', icon: <ZoomIn size={15} />, onClick: () => props.onZoomChange(100) },
            { id: 'fit-width', label: '适应宽度', icon: <ZoomIn size={15} />, onClick: props.onFitWidth },
          ],
        },
        {
          id: 'tools',
          label: '工具',
          items: [
            { id: 'print', label: '打印', icon: <Printer size={15} />, onClick: props.onPrint },
            { id: 'compat', label: '兼容报告', icon: <FileWarning size={15} />, onClick: props.onShowCompatibility },
            { id: 'more', kind: 'inline', content: moreMenu },
          ],
        },
      ],
    },
  ];

  return <Ribbon tabs={tabs} defaultTab="home" topBarLeft={topBarLeft} topBarRight={topBarRight} />;
}

function PaintbrushIcon() {
  return (
    <svg width={15} height={15} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
      <path d="M18 2l4 4-9 9-4-4 9-9z" />
      <path d="M2 22l4-1 1-4" />
    </svg>
  );
}

function NumberingIcon() {
  return <span className="text-[11px] font-semibold">1.</span>;
}

/** 功能：Ribbon 内下拉型按钮的统一外观（图标 + 下方文字标签）；时间：2026-10-31；作者：lq。 */
function MenuButton({ icon, label, toggle, disabled }: { icon: ReactNode; label: string; toggle: () => void; disabled?: boolean }) {
  return (
    <button
      type="button"
      aria-label={label}
      disabled={disabled}
      onClick={toggle}
      className="flex h-9 w-9 shrink-0 flex-col items-center justify-center gap-0.5 rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
    >
      {icon}
      <span className="max-w-[44px] truncate text-[9px] leading-none">{label}</span>
    </button>
  );
}
