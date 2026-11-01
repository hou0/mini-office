import type { SheetToolbarProps } from './SheetToolbar';
import { Ribbon, type RibbonTab } from '@mini-office/ui';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronDown,
  ClipboardPaste,
  Columns3,
  Copy,
  Download,
  FileWarning,
  Italic,
  ImagePlus,
  ListFilter,
  Merge,
  PanelRight,
  Printer,
  Redo2,
  Rows3,
  Scissors,
  Search,
  Settings2,
  Snowflake,
  Strikethrough,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react';
import { ColorField, Dropdown, Select } from '@mini-office/ui';
import { FONT_CHOICES } from '@mini-office/design-system';
import { cellAt, cellStyle } from './sheet-utils';
import { BorderMenu } from './BorderMenu';

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20,24, 28, 36];

const NUMBER_FORMATS = [
  { label: '常规', format: 'General' },
  { label: '数值 0.00', format: '0.00' },
  { label: '千分位 #,##0.00', format: '#,##0.00' },
  { label: '整数 #,##0', format: '#,##0' },
  { label: '货币 ¥#,##0.00', format: '¥#,##0.00' },
  { label: '百分比 0.00%', format: '0.00%' },
  { label: '日期 yyyy-MM-dd', format: 'yyyy-MM-dd' },
  { label: '时间 hh:mm:ss', format: 'hh:mm:ss' },
  { label: '文本 @', format: '@' },
];

export function SheetRibbon(props: SheetToolbarProps) {
  const { workbook, sheet, range, active, readonly } = props;
  const style = cellStyle(workbook, cellAt(sheet, active.row, active.col));
  const merged = sheet.merges.some(
    (merge) => merge.r1 <= range.r2 && merge.r2 >= range.r1 && merge.c1 <= range.c2 && merge.c2 >= range.c1,
  );
  const frozen = !!sheet.freeze && (sheet.freeze.rows > 0 || sheet.freeze.cols > 0);
  const formatValue = style.numFmt ?? 'General';
  const disabled = readonly;

  const fontSelect = (
    <span className="relative inline-flex shrink-0 items-center">
      <Select
        aria-label="字体"
        className="h-7 w-[110px] shrink-0 appearance-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] pl-2 pr-6 text-[12px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
        value={style.fontFamily ?? workbook.defaultFont}
        disabled={disabled}
        onChange={(event) => props.onStyle({ fontFamily: event.target.value })}
      >
        {[workbook.defaultFont, ...FONT_CHOICES.filter((font) => font !== '默认')].filter(
          (font, index, list) => list.indexOf(font) === index,
        ).map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </Select>
      <ChevronDown size={12} className="pointer-events-none absolute right-1.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-muted)]" />
    </span>
  );

  const sizeSelect = (
    <span className="relative inline-flex shrink-0 items-center">
      <Select
        aria-label="字号"
        className="h-7 w-[54px] shrink-0 appearance-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] pl-2 pr-6 text-[12px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
        value={String(style.fontSizePt ?? workbook.defaultFontSizePt)}
        disabled={disabled}
        onChange={(event) => props.onStyle({ fontSizePt: Number(event.target.value) })}
      >
        {[...new Set([style.fontSizePt ?? workbook.defaultFontSizePt, ...FONT_SIZES])]
          .sort((a, b) => a - b)
          .map((size) => (
            <option key={size} value={size}>
              {size}
            </option>
          ))}
      </Select>
      <ChevronDown size={12} className="pointer-events-none absolute right-1.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-muted)]" />
    </span>
  );

  const numberFormatSelect = (
    <span className="relative inline-flex shrink-0 items-center">
      <Select
        aria-label="数字格式"
        className="h-7 w-[110px] shrink-0 appearance-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] pl-2 pr-6 text-[11px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
        value={formatValue}
        disabled={disabled}
        onChange={(event) => props.onNumberFormat(event.target.value)}
      >
      {[...NUMBER_FORMATS, ...(NUMBER_FORMATS.some((item) => item.format === formatValue) ? [] : [{ label: formatValue, format: formatValue }])].map(
        (item) => (
          <option key={item.format} value={item.format}>
            {item.label}
          </option>
        ),
      )}
    </Select>
    <ChevronDown size={12} className="pointer-events-none absolute right-1.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-muted)]" />
    </span>
    );

    const dataMenu = (
    <Dropdown
      minWidth={210}
      items={[
        { type: 'label', label: '行' },
        { id: 'insert-row', label: '在上方插入行', icon: <Rows3 size={13} />, disabled, onSelect: props.onInsertRow },
        { id: 'delete-row', label: '删除当前行', icon: <Trash2Icon />, disabled, onSelect: props.onDeleteRow },
        { type: 'separator' },
        { type: 'label', label: '列' },
        { id: 'insert-col', label: '在左侧插入列', icon: <Columns3 size={13} />, disabled, onSelect: props.onInsertColumn },
        { id: 'delete-col', label: '删除当前列', icon: <Trash2Icon />, disabled, onSelect: props.onDeleteColumn },
      ]}
      trigger={({ toggle }) => (
        <button type="button" aria-label="数据与行列" disabled={disabled} onClick={toggle}
          className="flex h-9 w-9 shrink-0 flex-col items-center justify-center gap-0.5 rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:opacity-40">
          <Rows3 size={15} />
          <span className="max-w-[44px] truncate text-[9px] leading-none">行列</span>
        </button>
      )}
    />
  );

  const sortFilterMenu = (
    <Dropdown
      minWidth={190}
      items={[
        { id: 'asc', label: '升序', icon: <ArrowUpAZIcon />, disabled, onSelect: () => props.onSort(true) },
        { id: 'desc', label: '降序', icon: <ArrowDownAZIcon />, disabled, onSelect: () => props.onSort(false) },
        { id: 'custom', label: '自定义排序…', disabled, onSelect: props.onCustomSort },
        { type: 'separator' },
        { id: 'filter', label: sheet.autoFilter ? '关闭筛选' : '筛选', icon: <ListFilter size={14} />, disabled, onSelect: props.onFilter },
        { id: 'clear', label: '清除筛选条件', disabled: disabled || !sheet.autoFilter, onSelect: props.onClearFilter },
      ]}
      trigger={({ toggle }) => (
        <button type="button" aria-label="排序和筛选" disabled={disabled} onClick={toggle}
          className="flex h-9 w-9 shrink-0 flex-col items-center justify-center gap-0.5 rounded-md text-[var(--mo-text-primary)] transition-colors hover:bg-[var(--mo-surface-2)] disabled:opacity-40">
          <ListFilter size={15} />
          <span className="max-w-[44px] truncate text-[9px] leading-none">排序</span>
        </button>
      )}
    />
  );

  const tabs: RibbonTab[] = [
    {
      id: 'home',
      label: '开始',
      groups: [
        {
          id: 'clipboard',
          label: '剪贴板',
          items: [
            { id: 'undo', label: '撤销', icon: <Undo2 size={15} />, disabled: !props.canUndo || disabled, onClick: props.onUndo },
            { id: 'redo', label: '重做', icon: <Redo2 size={15} />, disabled: !props.canRedo || disabled, onClick: props.onRedo },
            { id: 'paste', label: '粘贴', icon: <ClipboardPaste size={15} />, disabled, onClick: props.onPaste },
            { id: 'cut', label: '剪切', icon: <Scissors size={16} />, hideLabel: true, disabled, onClick: props.onCut },
            { id: 'copy', label: '复制', icon: <Copy size={16} />, hideLabel: true, disabled, onClick: props.onCopy },
          ],
        },
        {
          id: 'font',
          label: '字体',
          items: [
            { id: 'font-family', kind: 'inline', content: fontSelect },
            { id: 'font-size', kind: 'inline', content: sizeSelect },
            { id: 'bold', label: '加粗', icon: <Bold size={16} />, hideLabel: true, active: !!style.bold, disabled, onClick: () => props.onStyle({ bold: !style.bold }) },
            { id: 'italic', label: '斜体', icon: <Italic size={16} />, hideLabel: true, active: !!style.italic, disabled, onClick: () => props.onStyle({ italic: !style.italic }) },
            { id: 'underline', label: '下划线', icon: <Underline size={16} />, hideLabel: true, active: !!style.underline, disabled, onClick: () => props.onStyle({ underline: !style.underline }) },
            { id: 'strike', label: '删除线', icon: <Strikethrough size={16} />, hideLabel: true, active: !!style.strike, disabled, onClick: () => props.onStyle({ strike: !style.strike }) },
          ],
        },
        {
          id: 'color',
          label: '颜色',
          items: [
            { id: 'font-color', kind: 'inline', content: <ColorField label="字体颜色" value={style.color} onChange={(color) => props.onStyle({ color })} /> },
            { id: 'fill-color', kind: 'inline', content: <ColorField label="填充颜色" value={style.fill} onChange={(fill) => props.onStyle({ fill })} /> },
          ],
        },
        {
          id: 'align',
          label: '对齐',
          items: [
            { id: 'align-left', label: '左对齐', icon: <AlignLeft size={16} />, hideLabel: true, active: style.horizontalAlign === 'left', disabled, onClick: () => props.onStyle({ horizontalAlign: 'left' }) },
            { id: 'align-center', label: '居中', icon: <AlignCenter size={16} />, hideLabel: true, active: style.horizontalAlign === 'center', disabled, onClick: () => props.onStyle({ horizontalAlign: 'center' }) },
            { id: 'align-right', label: '右对齐', icon: <AlignRight size={16} />, hideLabel: true, active: style.horizontalAlign === 'right', disabled, onClick: () => props.onStyle({ horizontalAlign: 'right' }) },
            { id: 'wrap', label: '自动换行', icon: <WrapText size={16} />, hideLabel: true, active: !!style.wrapText, disabled, onClick: () => props.onStyle({ wrapText: !style.wrapText }) },
          ],
        },
        {
          id: 'number',
          label: '数字',
          items: [
            { id: 'numfmt', kind: 'inline', content: numberFormatSelect },
            { id: 'merge', label: '合并', icon: <Merge size={15} />, active: merged, disabled, onClick: props.onMergeToggle },
            { id: 'freeze', label: '冻结', icon: <Snowflake size={15} />, active: frozen, onClick: props.onFreeze },
          ],
        },
        {
          id: 'border',
          label: '边框',
          items: [{ id: 'border', kind: 'inline', content: <BorderMenu readonly={disabled} weight={props.borderWeight} color={props.borderColor} onWeight={props.onBorderWeight} onColor={props.onBorderColor} onApply={props.onBorders} onMore={props.onMoreBorders} /> }],
        },
      ],
    },
    {
      id: 'insert',
      label: '插入',
      groups: [
        {
          id: 'data-media',
          label: '数据与媒体',
          items: [
            { id: 'rows-cols', kind: 'inline', content: dataMenu },
            { id: 'sort-filter', kind: 'inline', content: sortFilterMenu },
            { id: 'image', label: '图片', icon: <ImagePlus size={15} />, disabled, onClick: props.onInsertImage },
            { id: 'find', label: '查找', icon: <Search size={15} />, onClick: props.onFind },
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
          label: '显示与输出',
          launch: props.onPageSetup,
          items: [
            { id: 'properties', label: '属性面板', icon: <PanelRight size={15} />, active: props.propertiesOpen, onClick: props.onToggleProperties },
            { id: 'page-setup', label: '页面设置', icon: <Settings2 size={15} />, onClick: props.onPageSetup },
            { id: 'export', label: '导出PDF', icon: <Download size={15} />, onClick: props.onExport },
            { id: 'print', label: '打印', icon: <Printer size={15} />, onClick: props.onPrint },
            { id: 'compat', label: '兼容报告', icon: <FileWarning size={15} />, onClick: props.onCompatibility },
          ],
        },
      ],
    },
  ];

  return <Ribbon tabs={tabs} defaultTab="home" />;
}

function Trash2Icon() {
  return (
    <svg width={13} height={13} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6" />
    </svg>
  );
}

function ArrowUpAZIcon() {
  return (
    <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 8l4-4 4 4M7 4v16" />
    </svg>
  );
}

function ArrowDownAZIcon() {
  return (
    <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 16l4 4 4-4M7 4v16" />
    </svg>
  );
}
