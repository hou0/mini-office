import type { CellRange, CellStyle, SheetModel, WorkbookModel } from '@mini-office/document-model';
import { columnLabel } from '@mini-office/shared';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowUpAZ,
  Bold,
  Columns3,
  Download,
  FileWarning,
  Italic,
  ListFilter,
  Merge,
  PanelRight,
  Printer,
  Redo2,
  Rows3,
  Search,
  Snowflake,
  Strikethrough,
  Trash2,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react';
import {
  ColorField,
  Dropdown,
  IconButton,
  Input,
  Select,
  Tooltip,
  type MenuItem,
} from '@mini-office/ui';
import { FONT_CHOICES } from '@mini-office/design-system';
import { cellAt, cellStyle } from './sheet-utils';

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

export interface NumberFormatOption {
  label: string;
  format: string;
}

export const NUMBER_FORMATS: NumberFormatOption[] = [
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

export type BorderPreset = 'all' | 'outline' | 'bottom' | 'none';

export interface SheetToolbarProps {
  workbook: WorkbookModel;
  sheet: SheetModel;
  range: CellRange;
  active: { row: number; col: number };
  formulaText: string;
  readonly: boolean;
  canUndo: boolean;
  canRedo: boolean;
  zoom: number;
  propertiesOpen: boolean;
  onFormulaChange: (value: string) => void;
  onFormulaCommit: () => void;
  onJump: (ref: string) => void;
  onStyle: (patch: Partial<CellStyle>) => void;
  onNumberFormat: (format: string) => void;
  onBorders: (preset: BorderPreset) => void;
  onMergeToggle: () => void;
  onFreeze: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onInsertColumn: () => void;
  onDeleteColumn: () => void;
  onSort: (ascending: boolean) => void;
  onFilter: () => void;
  onFind: () => void;
  onExport: () => void;
  onPrint: () => void;
  onCompatibility: () => void;
  onToggleProperties: () => void;
  onZoom: (zoom: number) => void;
  onUndo: () => void;
  onRedo: () => void;
  onAddSheet: () => void;
}

export function SheetToolbar(props: SheetToolbarProps) {
  const { workbook, sheet, range, active, readonly } = props;
  const style = cellStyle(workbook, cellAt(sheet, active.row, active.col));
  const rangeLabel =
    range.r1 === range.r2 && range.c1 === range.c2
      ? `${columnLabel(range.c1)}${range.r1 + 1}`
      : `${columnLabel(range.c1)}${range.r1 + 1}:${columnLabel(range.c2)}${range.r2 + 1}`;
  const merged = sheet.merges.some(
    (merge) =>
      merge.r1 <= range.r2 && merge.r2 >= range.r1 && merge.c1 <= range.c2 && merge.c2 >= range.c1,
  );
  const frozen = !!sheet.freeze && (sheet.freeze.rows > 0 || sheet.freeze.cols > 0);

  const dataMenu: MenuItem[] = [
    { type: 'label', label: '行' },
    { id: 'insert-row', label: '在上方插入行', icon: <Rows3 size={13} />, disabled: readonly, onSelect: props.onInsertRow },
    { id: 'delete-row', label: '删除当前行', icon: <Trash2 size={13} />, disabled: readonly, onSelect: props.onDeleteRow },
    { type: 'separator' },
    { type: 'label', label: '列' },
    { id: 'insert-col', label: '在左侧插入列', icon: <Columns3 size={13} />, disabled: readonly, onSelect: props.onInsertColumn },
    { id: 'delete-col', label: '删除当前列', icon: <Trash2 size={13} />, disabled: readonly, onSelect: props.onDeleteColumn },
    { type: 'separator' },
    { id: 'sort-asc', label: '按当前列升序排序', icon: <ArrowUpAZ size={13} />, disabled: readonly, onSelect: () => props.onSort(true) },
    { id: 'sort-desc', label: '按当前列降序排序', icon: <ArrowDownAZ size={13} />, disabled: readonly, onSelect: () => props.onSort(false) },
    { id: 'filter', label: '筛选…', icon: <ListFilter size={13} />, onSelect: props.onFilter },
  ];

  const formatValue = style.numFmt ?? 'General';
  const borderMenu: MenuItem[] = [
    { id: 'all', label: '所有框线', onSelect: () => props.onBorders('all') },
    { id: 'outline', label: '外侧框线', onSelect: () => props.onBorders('outline') },
    { id: 'bottom', label: '下边框', onSelect: () => props.onBorders('bottom') },
    { id: 'none', label: '无框线', onSelect: () => props.onBorders('none') },
  ];

  return (
    <div className="mo-glass mo-no-print shrink-0 border-b border-[var(--mo-line)] bg-[var(--mo-surface)]">
      <div
        className="flex h-11 items-center gap-0.5 overflow-x-auto px-2"
        onMouseDown={(event) => {
          if ((event.target as HTMLElement).closest('button')) event.preventDefault();
        }}
      >
        <Tooltip label="撤销 ⌘Z">
          <IconButton label="撤销" disabled={!props.canUndo || readonly} onClick={props.onUndo}>
            <Undo2 size={15} />
          </IconButton>
        </Tooltip>
        <Tooltip label="重做 ⇧⌘Z">
          <IconButton label="重做" disabled={!props.canRedo || readonly} onClick={props.onRedo}>
            <Redo2 size={15} />
          </IconButton>
        </Tooltip>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <Select
          aria-label="字体"
          className="h-7 w-[120px] shrink-0 border-transparent bg-transparent text-[12px]"
          value={style.fontFamily ?? workbook.defaultFont}
          disabled={readonly}
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
        <Select
          aria-label="字号"
          className="h-7 w-[58px] shrink-0 border-transparent bg-transparent text-[12px]"
          value={String(style.fontSizePt ?? workbook.defaultFontSizePt)}
          disabled={readonly}
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
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <IconButton label="加粗" active={!!style.bold} disabled={readonly} onClick={() => props.onStyle({ bold: !style.bold })}>
          <Bold size={15} />
        </IconButton>
        <IconButton label="斜体" active={!!style.italic} disabled={readonly} onClick={() => props.onStyle({ italic: !style.italic })}>
          <Italic size={15} />
        </IconButton>
        <IconButton label="下划线" active={!!style.underline} disabled={readonly} onClick={() => props.onStyle({ underline: !style.underline })}>
          <Underline size={15} />
        </IconButton>
        <IconButton label="删除线" active={!!style.strike} disabled={readonly} onClick={() => props.onStyle({ strike: !style.strike })}>
          <Strikethrough size={15} />
        </IconButton>
        <ColorField
          label="字体颜色"
          value={style.color}
          onChange={(color) => props.onStyle({ color })}
        />
        <ColorField
          label="填充颜色"
          value={style.fill}
          onChange={(fill) => props.onStyle({ fill })}
        />
        <Dropdown
          minWidth={160}
          items={borderMenu}
          trigger={({ toggle }) => (
            <IconButton label="边框" onClick={toggle}>
              <span className="grid h-3.5 w-3.5 grid-cols-2 gap-[1px]">
                <span className="border border-current" />
                <span className="border border-current" />
                <span className="border border-current" />
                <span className="border border-current" />
              </span>
            </IconButton>
          )}
        />
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <IconButton label="左对齐" active={style.horizontalAlign === 'left'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'left' })}>
          <AlignLeft size={15} />
        </IconButton>
        <IconButton label="居中" active={style.horizontalAlign === 'center'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'center' })}>
          <AlignCenter size={15} />
        </IconButton>
        <IconButton label="右对齐" active={style.horizontalAlign === 'right'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'right' })}>
          <AlignRight size={15} />
        </IconButton>
        <IconButton label="自动换行" active={!!style.wrapText} disabled={readonly} onClick={() => props.onStyle({ wrapText: !style.wrapText })}>
          <WrapText size={15} />
        </IconButton>
        <Select
          aria-label="数字格式"
          className="h-7 w-[118px] shrink-0 border-transparent bg-transparent text-[11px]"
          value={formatValue}
          disabled={readonly}
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
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <IconButton label={merged ? '取消合并' : '合并单元格'} active={merged} disabled={readonly} onClick={props.onMergeToggle}>
          <Merge size={15} />
        </IconButton>
        <IconButton label="冻结窗格" active={frozen} onClick={props.onFreeze}>
          <Snowflake size={15} />
        </IconButton>
        <Dropdown
          minWidth={210}
          items={dataMenu}
          trigger={({ toggle }) => (
            <IconButton label="数据与行列" onClick={toggle}>
              <ListFilter size={15} />
            </IconButton>
          )}
        />
        <Tooltip label="查找替换 ⌘F">
          <IconButton label="查找" onClick={props.onFind}>
            <Search size={15} />
          </IconButton>
        </Tooltip>
        <IconButton label="属性面板" active={props.propertiesOpen} onClick={props.onToggleProperties}>
          <PanelRight size={15} />
        </IconButton>
        <div className="ml-auto flex shrink-0 items-center gap-1 pl-2">
          <span className="whitespace-nowrap text-[10px] text-[var(--mo-text-muted)]">
            {rangeLabel}
          </span>
          <Select
            aria-label="缩放"
            className="h-7 w-[68px] border-transparent bg-transparent text-[11px]"
            value={String(Math.round(props.zoom * 100))}
            onChange={(event) => props.onZoom(Number(event.target.value) / 100)}
          >
            {[75, 90, 100, 110, 125, 150].map((value) => (
              <option key={value} value={value}>
                {value}%
              </option>
            ))}
          </Select>
          <IconButton label="打印" onClick={props.onPrint}>
            <Printer size={15} />
          </IconButton>
          <IconButton label="导出 PDF" onClick={props.onExport}>
            <Download size={15} />
          </IconButton>
          <IconButton label="兼容性报告" onClick={props.onCompatibility}>
            <FileWarning size={15} />
          </IconButton>
        </div>
      </div>

      <div className="flex h-9 items-center gap-2 border-t border-[var(--mo-line)] px-2">
        <Input
          aria-label="名称框"
          className="h-6 w-[92px] shrink-0 text-center font-mono text-[11px]"
          defaultValue={rangeLabel}
          key={`${range.c1}-${range.r1}-${range.c2}-${range.r2}`}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              props.onJump(event.currentTarget.value);
              event.currentTarget.blur();
            }
          }}
        />
        <span className="shrink-0 font-serif text-[12px] italic text-[var(--mo-text-muted)]">fx</span>
        <Input
          aria-label="公式栏"
          className="h-6 min-w-0 flex-1 font-mono text-[12px]"
          value={props.formulaText}
          disabled={readonly}
          onChange={(event) => props.onFormulaChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              props.onFormulaCommit();
            }
          }}
          placeholder="输入值或公式，例如 =SUM(A1:A3)"
        />
      </div>
    </div>
  );
}
