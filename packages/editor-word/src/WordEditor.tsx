import {
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react';
import {
  findParagraph,
  collectParagraphs,
  computeDocStats,
  makeParagraph,
  paragraphImages,
  splitRunsAtOffset,
  tableCellPlacements,
  totalTableColumns,
  updateParagraph,
  type CellPlacement,
} from './doc-utils';
import {
  getCaretOffset,
  selectedParagraphSlices,
  getSelectionOffsets,
  replaceInRuns,
  runsPlainText,
  selectRangeByOffset,
  serializeDomRuns,
  setCaretAtOffset,
  findInParagraph,
  type FindHit,
} from './editing';
import {
  paragraphTextStyle,
  paragraphWrapperStyle,
  renderParagraphHtml,
} from './html';
import { replaceDocumentSelection, selectionHtml, selectionText } from './document-selection';
import { pageMetrics, paginateDocument, type PageSlice } from './pagination';
import { WordToolbar } from './WordToolbar';
// 功能：接入页眉页脚草稿设置；时间：2026-10-24；作者：lq。
import { HeaderFooterDialog } from './HeaderFooterDialog';
// 功能：接入模板字段和批量报告流程；时间：2026-10-24；作者：lq。
import { MailMergePanel } from './MailMergePanel';
// 功能：接入表格选区与网格编辑工具；时间：2026-10-19；作者：lq。
import { TableTools } from './TableTools';
import { editTable, findTable, formatTable, intersects, tableSelectionRect, type TableAction, type TableSelection } from './table-editing';
import { useWorkspaceStore, type WordSession } from '@mini-office/document-model';
import { printCurrentView } from '@mini-office/platform-adapter';
import type {
  DocBlock,
  DocImage,
  DocModel,
  DocParagraph,
  DocRun,
  DocTable,
  DocTableCell,
  ParagraphStyle,
  RunStyle,
  TableBorderSide,
} from '@mini-office/document-model';
import {
  Button,
  ContextMenu,
  Dialog,
  Field,
  IconButton,
  Input,
  ProgressBar,
  SegmentedControl,
  Select,
  Switch,
  type MenuItem,
} from '@mini-office/ui';
import { exportElementsToPdf, type ExportProgress } from '@mini-office/export-service';
import { PAGE_SIZES_PX, mmToPx, pxToMm } from '@mini-office/ooxml';
import {
  ensureExtension,
  formatBytes,
  stripExtension,
  uid,
} from '@mini-office/shared';
import {
  ChevronDown,
  ChevronUp,
  FileWarning,
  Replace,
  Search,
  X,
} from 'lucide-react';

const PAGE_SIZE_OPTIONS = ['A4', 'A3', 'A5', 'Letter'] as const;

interface HistoryState {
  past: DocModel[];
  future: DocModel[];
  lastKey?: string;
  lastTime: number;
}

const historyStore = new Map<string, HistoryState>();

function historyFor(id: string): HistoryState {
  const existing = historyStore.get(id);
  if (existing) return existing;
  const created: HistoryState = { past: [], future: [], lastTime: 0 };
  historyStore.set(id, created);
  return created;
}

function baseRunStyle(paragraph: DocParagraph, doc: DocModel): RunStyle {
  const first = paragraph.runs.find((run) => run.text)?.style ?? paragraph.runs[0]?.style ?? {};
  return {
    fontFamily: doc.defaultFont,
    fontSizePt: doc.defaultFontSizePt,
    ...first,
  };
}

function styleAtOffset(paragraph: DocParagraph, offset: number): RunStyle {
  let consumed = 0;
  for (const run of paragraph.runs) {
    const length = run.image || run.tab || run.lineBreak || run.mergeField ? 1 : (run.text ?? '').length;
    if (offset <= consumed + length) return run.style;
    consumed += length;
  }
  return paragraph.runs.at(-1)?.style ?? {};
}

function atomicLength(runs: DocRun[]): number {
  return runs.reduce((sum, run) => {
    if (run.image || run.tab || run.lineBreak || run.mergeField) return sum + 1;
    return sum + (run.text ?? '').length;
  }, 0);
}

function blockContainsParagraph(block: DocBlock, paragraphId: string): boolean {
  if (block.type === 'paragraph') return block.id === paragraphId;
  if (block.type !== 'table') return false;
  return block.rows.some((row) =>
    row.cells.some((cell) => cell.blocks.some((child) => blockContainsParagraph(child, paragraphId))),
  );
}

function topLevelBlockIndex(doc: DocModel, paragraphId: string | undefined): number {
  if (!paragraphId) return doc.blocks.length - 1;
  return doc.blocks.findIndex((block) => blockContainsParagraph(block, paragraphId));
}

function insertAfterParagraph(
  blocks: DocBlock[],
  paragraphId: string,
  paragraph: DocParagraph,
): { blocks: DocBlock[]; inserted: boolean } {
  const result: DocBlock[] = [];
  let inserted = false;
  for (const block of blocks) {
    if (block.type === 'paragraph' && block.id === paragraphId && !inserted) {
      result.push(block, paragraph);
      inserted = true;
      continue;
    }
    if (block.type === 'table' && !inserted) {
      let tableInserted = false;
      const table: DocTable = {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => {
            if (tableInserted) return cell;
            const nested = insertAfterParagraph(cell.blocks, paragraphId, paragraph);
            if (nested.inserted) {
              tableInserted = true;
              return { ...cell, blocks: nested.blocks };
            }
            return cell;
          }),
        })),
      };
      if (tableInserted) inserted = true;
      result.push(table);
      continue;
    }
    result.push(block);
  }
  return { blocks: result, inserted };
}

function removeTopLevelBlock(blocks: DocBlock[], blockId: string): DocBlock[] {
  return blocks.filter((block) => block.id !== blockId);
}

function replaceTableInBlocks(blocks: DocBlock[], tableId: string, updater: (table: DocTable) => DocTable | null): DocBlock[] {
  const result: DocBlock[] = [];
  for (const block of blocks) {
    if (block.type === 'table') {
      if (block.id === tableId) {
        const next = updater(block);
        if (next) result.push(next);
        continue;
      }
      const nextTable: DocTable = {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            blocks: replaceTableInBlocks(cell.blocks, tableId, updater),
          })),
        })),
      };
      result.push(nextTable);
      continue;
    }
    result.push(block);
  }
  return result;
}

function borderCss(border: TableBorderSide | undefined): string | undefined {
  if (!border || border.style === 'none' || border.widthPx <= 0) return undefined;
  const cssStyle =
    border.style === 'double'
      ? 'double'
      : border.style === 'dashed'
        ? 'dashed'
        : border.style === 'dotted'
          ? 'dotted'
          : 'solid';
  return `${Math.max(0.75, border.widthPx)}px ${cssStyle} ${border.color ?? '#000000'}`;
}

function cellBorderStyle(cell: DocTable['rows'][number]['cells'][number], table: DocTable, rowIndex: number, col: number): CSSProperties {
  const tableBorders = table.borders;
  const cellBorders = cell.borders;
  const isLastRow = rowIndex + cell.rowSpan >= table.rows.length;
  const isLastCol = col + cell.colSpan >= totalTableColumns(table);
  const top = cellBorders?.top ?? (rowIndex === 0 ? tableBorders?.top : tableBorders?.insideH);
  const bottom = cellBorders?.bottom ?? (isLastRow ? tableBorders?.bottom : tableBorders?.insideH);
  const left = cellBorders?.left ?? (col === 0 ? tableBorders?.left : tableBorders?.insideV);
  const right = cellBorders?.right ?? (isLastCol ? tableBorders?.right : tableBorders?.insideV);
  return {
    borderTop: borderCss(top),
    borderBottom: borderCss(bottom),
    borderLeft: borderCss(left),
    borderRight: borderCss(right),
  };
}

interface ParagraphHandlers {
  syncToken: number;
  onInput: (paragraphId: string, element: HTMLElement) => void;
  onBlur: (paragraphId: string) => void;
  onKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>, paragraphId: string, element: HTMLDivElement) => void;
  onTableContextMenu: (event: React.MouseEvent, table: DocTable, rowIndex: number, cellId: string) => void;
  // 功能：共享选区及尺寸拖动入口；时间：2026-10-19；作者：lq。
  tableSelection?: TableSelection | null;
  onTableResize: (event: React.PointerEvent, table: DocTable, index: number, axis: 'row' | 'column') => void;
  onPaste?: (event: React.ClipboardEvent<HTMLDivElement>) => void;
}

const StaticParagraph = memo(function StaticParagraph({
  paragraph,
  doc,
  pageNumber,
  totalPages,
}: {
  paragraph: DocParagraph;
  doc: DocModel;
  pageNumber: number;
  totalPages: number;
}) {
  const html = useMemo(
    () => renderParagraphHtml(paragraph, { pageNumber, totalPages, sectionPages: doc.sectionPageCount }),
    [paragraph, pageNumber, totalPages, doc.sectionPageCount],
  );
  return (
    <div style={paragraphWrapperStyle(paragraph)}>
      <div
        style={paragraphTextStyle(paragraph, doc.defaultFont, doc.defaultFontSizePt)}
        dangerouslySetInnerHTML={{ __html: html }}
      />
    </div>
  );
});

const EditableParagraph = memo(function EditableParagraph({
  paragraph,
  doc,
  pageNumber,
  totalPages,
  readonly,
  handlers,
}: {
  paragraph: DocParagraph;
  doc: DocModel;
  pageNumber: number;
  totalPages: number;
  readonly: boolean;
  handlers: ParagraphHandlers;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const lastSyncToken = useRef(handlers.syncToken);
  const html = useMemo(
    () => renderParagraphHtml(paragraph, { pageNumber, totalPages, sectionPages: doc.sectionPageCount }),
    [paragraph, pageNumber, totalPages, doc.sectionPageCount],
  );

  useLayoutEffect(() => {
    const element = ref.current;
    if (!element) return;
    // 功能：共享编辑宿主下按选区判断焦点，避免跨段格式化时重置选区；时间：2026-10-18；作者：lq。
    const selection = window.getSelection();
    const focused = element.contains(document.activeElement) || !!(
      document.activeElement?.closest('[data-word-editor]') && selection?.rangeCount && selection.getRangeAt(0).intersectsNode(element)
    );
    const force = handlers.syncToken !== lastSyncToken.current;
    if (focused && !force) return;
    const offset = focused ? getCaretOffset(element) : 0;
    if (element.innerHTML !== html) element.innerHTML = html;
    if (focused && selection?.isCollapsed) {
      element.focus({ preventScroll: true });
      setCaretAtOffset(element, offset);
    }
    lastSyncToken.current = handlers.syncToken;
  }, [handlers.syncToken, html]);

  return (
    <div style={paragraphWrapperStyle(paragraph)}>
      <div
        ref={ref}
        // 功能：继承整篇文档的编辑宿主，允许跨段和跨页选区；时间：2026-10-18；作者：lq。
        tabIndex={-1}
        suppressContentEditableWarning
        spellCheck={false}
        data-paragraph-id={paragraph.id}
        role="textbox"
        aria-multiline="true"
        aria-label="文档段落"
        style={{ ...paragraphTextStyle(paragraph, doc.defaultFont, doc.defaultFontSizePt), outline: 'none' }}
        onInput={(event) => handlers.onInput(paragraph.id, event.currentTarget)}
        onBlur={() => handlers.onBlur(paragraph.id)}
      />
    </div>
  );
});

function TableGrid({
  table,
  doc,
  pageNumber,
  totalPages,
  readonly,
  handlers,
  staticMode,
}: {
  table: DocTable;
  doc: DocModel;
  pageNumber: number;
  totalPages: number;
  readonly: boolean;
  handlers?: ParagraphHandlers;
  staticMode?: boolean;
}) {
  const placements = useMemo(() => tableCellPlacements(table), [table]);
  const columnCount = totalTableColumns(table);
  const widths = useMemo(
    () =>
      Array.from({ length: columnCount }, (_, index) => table.columnWidthsPx[index] ?? 100),
    [table.columnWidthsPx, columnCount],
  );
  const totalWidth = widths.reduce((sum, width) => sum + width, 0);
  const gridStyle: CSSProperties = {
    display: 'grid',
    gridTemplateColumns: widths.map((width) => `${width}px`).join(' '),
    // 功能：按行高设置最小轨道高度，内容较多时自动撑开；时间：2026-10-19；作者：lq。
    gridTemplateRows: table.rows.map((row) => `minmax(${row.heightPx ?? 28}px, auto)`).join(' '),
    width: totalWidth,
    maxWidth: '100%',
    marginLeft: table.alignment === 'center' ? 'auto' : table.alignment === 'right' ? 'auto' : undefined,
    marginRight: table.alignment === 'right' ? '0' : table.alignment === 'center' ? 'auto' : undefined,
  };

  // 功能：高亮矩形单元格选区，不影响打印和 PDF；时间：2026-10-19；作者：lq。
  const selectionRect = !readonly && !staticMode && handlers?.tableSelection?.tableId === table.id ? tableSelectionRect(table, handlers.tableSelection) : undefined;
  return (
    <div style={gridStyle} data-table-id={table.id}>
      {placements.map((placement: CellPlacement) => {
        const { cell, row, col } = placement;
        const backgroundColor = cell.backgroundColor;
        const justify =
          cell.verticalAlign === 'middle' ? 'center' : cell.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start';
        return (
          <div
            key={cell.id}
            data-cell-id={cell.id}
            onContextMenu={
              handlers
                ? (event) => handlers.onTableContextMenu(event, table, row, cell.id)
                : undefined
            }
            style={{
              // 功能：单元格选区轮廓与边缘拖动定位；时间：2026-10-19；作者：lq。
              position: 'relative',
              boxShadow: selectionRect && intersects(placement, selectionRect) ? 'inset 0 0 0 2px #3478f6' : undefined,
              gridColumn: `${col + 1} / span ${cell.colSpan}`,
              gridRow: `${row + 1} / span ${cell.rowSpan}`,
              display: 'flex',
              flexDirection: 'column',
              justifyContent: justify,
              minWidth: 0,
              padding: '3px 6px',
              boxSizing: 'border-box',
              backgroundColor,
              ...cellBorderStyle(cell, table, row, col),
            }}
          >
            {/* 功能：提供列右边界、行下边界拖柄；时间：2026-10-19；作者：lq。 */}
            {!readonly && !staticMode && handlers ? <>
              <span data-table-resize contentEditable={false} title="拖动调整列宽" onPointerDown={(event) => handlers.onTableResize(event, table, col + cell.colSpan - 1, 'column')} style={{ position: 'absolute', zIndex: 2, right: -3, top: 0, bottom: 6, width: 6, cursor: 'col-resize', touchAction: 'none' }} />
              <span data-table-resize contentEditable={false} title="拖动调整行高" onPointerDown={(event) => handlers.onTableResize(event, table, row + cell.rowSpan - 1, 'row')} style={{ position: 'absolute', zIndex: 2, bottom: -3, left: 0, right: 6, height: 6, cursor: 'row-resize', touchAction: 'none' }} />
            </> : null}
            {cell.blocks.map((block) =>
              block.type === 'paragraph' ? (
                staticMode || !handlers ? (
                  <StaticParagraph
                    key={block.id}
                    paragraph={block}
                    doc={doc}
                    pageNumber={pageNumber}
                    totalPages={totalPages}
                  />
                ) : (
                  <EditableParagraph
                    key={block.id}
                    paragraph={block}
                    doc={doc}
                    pageNumber={pageNumber}
                    totalPages={totalPages}
                    readonly={readonly}
                    handlers={handlers}
                  />
                )
              ) : block.type === 'table' ? (
                <TableGrid
                  key={block.id}
                  table={block}
                  doc={doc}
                  pageNumber={pageNumber}
                  totalPages={totalPages}
                  readonly={readonly}
                  handlers={handlers}
                  staticMode={staticMode}
                />
              ) : null,
            )}
          </div>
        );
      })}
    </div>
  );
}

function HeaderFooterView({
  blocks,
  doc,
  pageNumber,
  totalPages,
}: {
  blocks: DocBlock[];
  doc: DocModel;
  pageNumber: number;
  totalPages: number;
}) {
  return (
    <div contentEditable={false}>
      {blocks.map((block) =>
        block.type === 'paragraph' ? (
          <StaticParagraph
            key={block.id}
            paragraph={block}
            doc={doc}
            pageNumber={pageNumber}
            totalPages={totalPages}
          />
        ) : block.type === 'table' ? (
          <TableGrid
            key={block.id}
            table={block}
            doc={doc}
            pageNumber={pageNumber}
            totalPages={totalPages}
            readonly
            staticMode
          />
        ) : null,
      )}
    </div>
  );
}

/** 功能：为预览和批量输出提供与正文一致的静态渲染；时间：2026-10-24；作者：lq。 */
function renderMergeBlocks(blocks: DocBlock[], doc: DocModel, pageNumber: number, totalPages: number) {
  return <HeaderFooterView blocks={blocks} doc={doc} pageNumber={pageNumber} totalPages={totalPages} />;
}

function FindBar({
  query,
  replacement,
  caseSensitive,
  count,
  active,
  readonly,
  onQuery,
  onReplacement,
  onCaseSensitive,
  onNext,
  onPrev,
  onReplace,
  onReplaceAll,
  onClose,
}: {
  query: string;
  replacement: string;
  caseSensitive: boolean;
  count: number;
  active: number;
  readonly: boolean;
  onQuery: (value: string) => void;
  onReplacement: (value: string) => void;
  onCaseSensitive: (value: boolean) => void;
  onNext: () => void;
  onPrev: () => void;
  onReplace: () => void;
  onReplaceAll: () => void;
  onClose: () => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    inputRef.current?.focus();
  }, []);
  return (
    <div className="mo-pop-in absolute right-4 top-3 z-30 w-[330px] rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2.5 shadow-pop">
      <div className="flex items-center gap-1.5">
        <div className="relative flex-1">
          <Search size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
          <Input
            ref={inputRef}
            value={query}
            onChange={(event) => onQuery(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                if (event.shiftKey) onPrev();
                else onNext();
              }
              if (event.key === 'Escape') onClose();
            }}
            placeholder="查找内容"
            className="h-7 pl-7 text-[12px]"
          />
        </div>
        <span className="w-14 whitespace-nowrap text-center text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {count > 0 ? `${active + 1}/${count}` : '0/0'}
        </span>
        <IconButton label="上一个" onClick={onPrev} disabled={count === 0}>
          <ChevronUp size={14} />
        </IconButton>
        <IconButton label="下一个" onClick={onNext} disabled={count === 0}>
          <ChevronDown size={14} />
        </IconButton>
        <IconButton label="关闭" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </div>
      {!readonly ? (
        <div className="mt-2 flex items-center gap-1.5">
          <div className="relative flex-1">
            <Replace size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
            <Input
              value={replacement}
              onChange={(event) => onReplacement(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  onReplace();
                }
                if (event.key === 'Escape') onClose();
              }}
              placeholder="替换为"
              className="h-7 pl-7 text-[12px]"
            />
          </div>
          <Button size="sm" onClick={onReplace} disabled={count === 0}>
            替换
          </Button>
          <Button size="sm" onClick={onReplaceAll} disabled={count === 0}>
            全部替换
          </Button>
        </div>
      ) : null}
      <label className="mt-2 flex cursor-pointer items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(event) => onCaseSensitive(event.target.checked)}
          className="h-3 w-3 accent-[var(--mo-accent)]"
        />
        区分大小写
      </label>
    </div>
  );
}

function PropertiesPanel({
  doc,
  session,
  activeParagraph,
  pageCount,
  onPageChange,
  onParagraphChange,
  onShowCompatibility,
  onClose,
}: {
  doc: DocModel;
  session: WordSession;
  activeParagraph?: DocParagraph;
  pageCount: number;
  onPageChange: (page: DocModel['page']) => void;
  onParagraphChange: (patch: Partial<ParagraphStyle>) => void;
  onShowCompatibility: () => void;
  onClose: () => void;
}) {
  const stats = useMemo(() => computeDocStats(doc), [doc]);
  const page = doc.page;
  const setSize = (name: string) => {
    const size = PAGE_SIZES_PX[name as keyof typeof PAGE_SIZES_PX];
    if (!size) return;
    const width = page.orientation === 'landscape' ? size.height : size.width;
    const height = page.orientation === 'landscape' ? size.width : size.height;
    onPageChange({ ...page, sizeName: name, widthPx: width, heightPx: height });
  };
  const setOrientation = (value: 'portrait' | 'landscape') => {
    if (value === page.orientation) return;
    onPageChange({
      ...page,
      orientation: value,
      widthPx: page.heightPx,
      heightPx: page.widthPx,
    });
  };
  const setMargin = (key: 'top' | 'right' | 'bottom' | 'left', millimeters: number) => {
    onPageChange({
      ...page,
      margins: { ...page.margins, [key]: mmToPx(Math.max(0, millimeters)) },
    });
  };
  const activeStyle = activeParagraph?.style ?? {};
  return (
    <aside className="mo-glass mo-no-print flex w-[268px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]">
      <div className="flex h-10 items-center justify-between border-b border-[var(--mo-line)] px-3">
        <span className="text-[12px] font-semibold">文档属性</span>
        <IconButton label="关闭属性面板" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </div>
      <div className="space-y-4 p-3">
        <section className="space-y-2">
          <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">页面</h3>
          <Field label="纸张尺寸">
            <Select value={page.sizeName} onChange={(event) => setSize(event.target.value)}>
              {PAGE_SIZE_OPTIONS.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
              <option value="自定义">自定义</option>
            </Select>
          </Field>
          <Field label="方向">
            <SegmentedControl
              value={page.orientation}
              onChange={setOrientation}
              options={[
                { value: 'portrait', label: '纵向' },
                { value: 'landscape', label: '横向' },
              ]}
              className="w-full"
            />
          </Field>
          <div className="grid grid-cols-2 gap-2">
            {(['top', 'right', 'bottom', 'left'] as const).map((key) => (
              <Field key={key} label={`页边距 ${key === 'top' ? '上' : key === 'right' ? '右' : key === 'bottom' ? '下' : '左'} (mm)`}>
                <Input
                  type="number"
                  min={0}
                  step={1}
                  value={Math.round(pxToMm(page.margins[key]))}
                  onChange={(event) => setMargin(key, Number(event.target.value))}
                />
              </Field>
            ))}
          </div>
        </section>

        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">段落</h3>
          <div className="grid grid-cols-2 gap-2">
            <Field label="段前 (pt)">
              <Input
                type="number"
                min={0}
                step={1}
                value={Math.round((activeStyle.spaceBeforePx ?? 0) * 0.75)}
                onChange={(event) => onParagraphChange({ spaceBeforePx: (Number(event.target.value) / 0.75) || 0 })}
                disabled={!activeParagraph}
              />
            </Field>
            <Field label="段后 (pt)">
              <Input
                type="number"
                min={0}
                step={1}
                value={Math.round((activeStyle.spaceAfterPx ?? 0) * 0.75)}
                onChange={(event) => onParagraphChange({ spaceAfterPx: (Number(event.target.value) / 0.75) || 0 })}
                disabled={!activeParagraph}
              />
            </Field>
          </div>
          <Field label="行距">
            <Select
              value={
                activeStyle.lineHeightExactPx
                  ? 'exact'
                  : String(activeStyle.lineHeight ?? 1.5)
              }
              onChange={(event) => {
                const value = event.target.value;
                if (value === 'exact') {
                  onParagraphChange({ lineHeightExactPx: 24, lineHeight: undefined });
                } else {
                  onParagraphChange({ lineHeight: Number(value), lineHeightExactPx: undefined });
                }
              }}
              disabled={!activeParagraph}
            >
              <option value="1">单倍行距</option>
              <option value="1.15">1.15 倍</option>
              <option value="1.5">1.5 倍</option>
              <option value="2">2 倍</option>
              <option value="exact">固定值</option>
            </Select>
          </Field>
          <Field label="首行缩进 (字符)">
            <Input
              type="number"
              step={0.5}
              min={-10}
              value={Number(((activeStyle.firstLineIndentPx ?? 0) / (doc.defaultFontSizePt * 1.333)).toFixed(1))}
              onChange={(event) =>
                onParagraphChange({
                  firstLineIndentPx: Number(event.target.value) * doc.defaultFontSizePt * 1.333,
                })
              }
              disabled={!activeParagraph}
            />
          </Field>
          <label className="flex items-center justify-between text-[11px]">
            <span>段前分页</span>
            <Switch
              checked={!!activeStyle.pageBreakBefore}
              onChange={(checked) => onParagraphChange({ pageBreakBefore: checked })}
              label="段前分页"
              disabled={!activeParagraph}
            />
          </label>
        </section>

        <section className="space-y-1.5 border-t border-[var(--mo-line)] pt-3">
          <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">统计</h3>
          <InfoRow label="页数" value={String(pageCount)} />
          <InfoRow label="字数" value={String(stats.words)} />
          <InfoRow label="字符数" value={String(stats.characters)} />
          <InfoRow label="文件大小" value={formatBytes(session.sourceSize)} />
          <InfoRow label="格式级别" value={`L${session.capability.level}`} />
        </section>

        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <button
            type="button"
            onClick={onShowCompatibility}
            className="flex w-full items-center gap-2 rounded-md bg-[var(--mo-surface-2)] px-2 py-2 text-left text-[11px] hover:bg-[var(--mo-surface-3)]"
          >
            <FileWarning size={14} className="text-[var(--mo-warning)]" />
            <span className="flex-1">查看兼容性报告</span>
            <span className="text-[var(--mo-text-faint)]">{session.capability.warnings.length}</span>
          </button>
          {session.capability.warnings.slice(0, 3).map((warning) => (
            <p key={warning.code} className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
              · {warning.message}
            </p>
          ))}
        </section>
      </div>
    </aside>
  );
}

function InfoRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between text-[11px]">
      <span className="text-[var(--mo-text-muted)]">{label}</span>
      <span className="tabular-nums text-[var(--mo-text-primary)]">{value}</span>
    </div>
  );
}

interface TableMenuState {
  x: number;
  y: number;
  tableId: string;
  rowIndex: number;
}

export function WordEditor({ session }: { session: WordSession }) {
  const doc = session.data as DocModel;
  const readonly = session.readonly;
  const updateData = useWorkspaceStore((state) => state.updateData);
  const updateView = useWorkspaceStore((state) => state.updateView);
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const exportBytes = useWorkspaceStore((state) => state.exportBytes);
  const openCompatibility = useWorkspaceStore((state) => state.openDialog);
  const notify = useWorkspaceStore((state) => state.notify);

  const scrollRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<HTMLDivElement>(null);
  const measureRef = useRef<HTMLDivElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const focusRequestRef = useRef<{ paragraphId: string; offset: number } | null>(null);
  const docRef = useRef(doc);
  docRef.current = doc;

  const [layoutRevision, setLayoutRevision] = useState(0);
  const [pages, setPages] = useState<PageSlice[]>([{ index: 0, blocks: doc.blocks }]);
  const [zoom, setZoom] = useState(() => session.view.zoom ?? 1);
  const [currentPage, setCurrentPage] = useState(1);
  const [activeParagraphId, setActiveParagraphId] = useState<string>();
  const [activeStyle, setActiveStyle] = useState<RunStyle>({
    fontFamily: doc.defaultFont,
    fontSizePt: doc.defaultFontSizePt,
  });
  const [historyVersion, setHistoryVersion] = useState(0);
  const [syncToken, setSyncToken] = useState(0);
  const savedSelectionRef = useRef<{ paragraphId: string; start: number; end: number } | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [findActive, setFindActive] = useState(0);
  // 功能：保存独立于文本焦点的表格矩形选区；时间：2026-10-19；作者：lq。
  const explicitTableSelectionRef = useRef(false);
  const [tableSelection, setTableSelection] = useState<TableSelection | null>(null);
  const [tableMenu, setTableMenu] = useState<TableMenuState | null>(null);
  const [exportOpen, setExportOpen] = useState(false);
  // 功能：按双击区域打开设置，每次打开重建草稿；时间：2026-10-24；作者：lq。
  const [headerFooterRegion, setHeaderFooterRegion] = useState<'header' | 'footer' | null>(null);
  const [mailMergeOpen, setMailMergeOpen] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportName, setExportName] = useState('');
  const [exportRange, setExportRange] = useState('');
  const [exportProgress, setExportProgress] = useState<ExportProgress | null>(null);
  const cancelExportRef = useRef(false);
  const scrollRafRef = useRef<number | null>(null);

  const latestDoc = useCallback((): DocModel => {
    const latest = useWorkspaceStore.getState().sessions[session.id]?.data as DocModel | undefined;
    return latest ?? docRef.current;
  }, [session.id]);

  const history = historyStore.get(session.id) ?? historyFor(session.id);
  void historyVersion;
  const canUndo = history.past.length > 0;
  const canRedo = history.future.length > 0;

  const pushHistory = useCallback(
    (previous: DocModel, key?: string) => {
      const stack = historyFor(session.id);
      const now = Date.now();
      if (key && stack.lastKey === key && now - stack.lastTime < 900) {
        stack.lastTime = now;
        return;
      }
      stack.past.push(previous);
      if (stack.past.length > 80) stack.past.shift();
      stack.future = [];
      stack.lastKey = key;
      stack.lastTime = now;
      setHistoryVersion((value) => value + 1);
    },
    [session.id],
  );

  const commit = useCallback(
    (next: DocModel, options?: { key?: string; reflow?: boolean; history?: boolean }) => {
      if (readonly) return;
      const previous = latestDoc();
      if (options?.history !== false) pushHistory(previous, options?.key);
      updateData(session.id, { ...next, stats: computeDocStats(next) });
      if (options?.reflow) setLayoutRevision((value) => value + 1);
    },
    [latestDoc, pushHistory, readonly, session.id, updateData],
  );

  const undo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const previous = stack.past.pop();
    if (!previous) return;
    stack.future.push(latestDoc());
    stack.lastKey = undefined;
    updateData(session.id, previous);
    setHistoryVersion((value) => value + 1);
    setSyncToken((value) => value + 1);
    setLayoutRevision((value) => value + 1);
  }, [latestDoc, readonly, session.id, updateData]);

  const redo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const next = stack.future.pop();
    if (!next) return;
    stack.past.push(latestDoc());
    stack.lastKey = undefined;
    updateData(session.id, next);
    setHistoryVersion((value) => value + 1);
    setSyncToken((value) => value + 1);
    setLayoutRevision((value) => value + 1);
  }, [latestDoc, readonly, session.id, updateData]);

  const rememberSelection = useCallback((range: Range, element: HTMLElement) => {
    const paragraphId = element.dataset.paragraphId;
    if (!paragraphId) return;
    const offsets = getSelectionOffsets(element);
    savedSelectionRef.current = {
      paragraphId,
      start: offsets?.start ?? getCaretOffset(element),
      end: offsets?.end ?? getCaretOffset(element),
    };
    savedRangeRef.current = range.cloneRange();
  }, []);

  const captureSelection = useCallback(() => {
    const selection = window.getSelection();
    const container = pagesRef.current;
    if (!selection || selection.rangeCount === 0 || !container) return;
    const range = selection.getRangeAt(0);
    if (!container.contains(range.startContainer) || !container.contains(range.endContainer)) return;
    const node = range.startContainer;
    const element = (node instanceof Element ? node : node.parentElement)?.closest<HTMLElement>(
      '[data-paragraph-id]',
    );
    if (!element || !container.contains(element)) return;
    rememberSelection(range, element);
  }, [rememberSelection]);

  const refreshSelection = useCallback(() => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    // 功能：从原生拖选的两端解析同表单元格，支持跨单元格合并；时间：2026-10-19；作者：lq。
    const cellAt = (node: Node | null) => (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('[data-cell-id]');
    const anchorCell = cellAt(selection.anchorNode), focusCell = cellAt(selection.focusNode);
    if (!explicitTableSelectionRef.current && anchorCell && focusCell && pagesRef.current?.contains(anchorCell)) {
      const table = anchorCell.closest<HTMLElement>('[data-table-id]');
      if (table && table === focusCell.closest('[data-table-id]')) setTableSelection({ tableId: table.dataset.tableId!, anchorId: anchorCell.dataset.cellId!, focusId: focusCell.dataset.cellId! });
    } else if (!explicitTableSelectionRef.current && pagesRef.current?.contains(selection.anchorNode)) setTableSelection(null);
    const node = selection.anchorNode;
    const element = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>(
      '[data-paragraph-id]',
    );
    if (!element || !pagesRef.current?.contains(element)) return;
    const paragraphId = element.dataset.paragraphId;
    if (!paragraphId) return;
    const current = latestDoc();
    const paragraph = findParagraph(current, paragraphId);
    if (!paragraph) return;
    const focusOutside = !!document.activeElement && !element.contains(document.activeElement);
    if (!(focusOutside && selection.isCollapsed)) {
      const range = selection.getRangeAt(0);
      if (pagesRef.current?.contains(range.startContainer) && pagesRef.current?.contains(range.endContainer)) {
        rememberSelection(range, element);
      }
    }
    const offset = getCaretOffset(element);
    setActiveParagraphId(paragraphId);
    setActiveStyle(styleAtOffset(paragraph, offset));
  }, [latestDoc, rememberSelection]);

  useEffect(() => {
    document.addEventListener('selectionchange', refreshSelection);
    return () => document.removeEventListener('selectionchange', refreshSelection);
  }, [refreshSelection]);

  const activeParagraph = activeParagraphId
    ? findParagraph(doc, activeParagraphId)
    : undefined;

  const metrics = useMemo(() => pageMetrics(doc.page), [doc.page]);
  const effectiveZoom = exporting ? 1 : zoom;

  const blockMap = useMemo(() => new Map(doc.blocks.map((block) => [block.id, block])), [doc.blocks]);

  useLayoutEffect(() => {
    const container = measureRef.current;
    if (!container) return;
    const heights = new Map<string, number>();
    container.querySelectorAll<HTMLElement>('[data-measure-id]').forEach((element) => {
      const id = element.getAttribute('data-measure-id');
      if (id) heights.set(id, element.getBoundingClientRect().height);
    });
    setPages(paginateDocument(docRef.current, heights));
  }, [
    layoutRevision,
    metrics.contentHeight,
    doc.page.widthPx,
    doc.page.heightPx,
    doc.page.margins.top,
    doc.page.margins.right,
    doc.page.margins.bottom,
    doc.page.margins.left,
    metrics.contentWidth,
    doc.blocks.length,
  ]);

  useLayoutEffect(() => {
    const request = focusRequestRef.current;
    if (!request) return;
    const element = findParagraphElement(pagesRef.current, request.paragraphId);
    if (element) {
      element.focus({ preventScroll: true });
      setCaretAtOffset(element, request.offset);
      element.scrollIntoView({ block: 'nearest' });
    }
    focusRequestRef.current = null;
  }, [pages, syncToken]);

  useEffect(() => {
    const container = measureRef.current;
    if (!container || typeof ResizeObserver === 'undefined') return;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const observer = new ResizeObserver(() => {
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => setLayoutRevision((value) => value + 1), 240);
    });
    observer.observe(container);
    return () => {
      if (timer) clearTimeout(timer);
      observer.disconnect();
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    if (document.fonts) {
      document.fonts.ready
        .then(() => {
          if (!cancelled) setLayoutRevision((value) => value + 1);
        })
        .catch(() => {});
    }
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const handleFind = () => setFindOpen((open) => !open);
    const handleExport = () => {
      setExportName(`${stripExtension(session.displayName)}.pdf`);
      setExportRange('');
      setExportOpen(true);
    };
    const handlePrint = async () => {
      const ok = await printCurrentView();
      if (!ok) notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
    };
    window.addEventListener('mo:find', handleFind);
    window.addEventListener('mo:export-pdf', handleExport);
    window.addEventListener('mo:print', handlePrint);
    return () => {
      window.removeEventListener('mo:find', handleFind);
      window.removeEventListener('mo:export-pdf', handleExport);
      window.removeEventListener('mo:print', handlePrint);
    };
  }, [notify, session.displayName]);

  const handleScroll = useCallback(() => {
    if (scrollRafRef.current !== null) return;
    scrollRafRef.current = requestAnimationFrame(() => {
      scrollRafRef.current = null;
      const container = scrollRef.current;
      if (!container) return;
      const containerTop = container.getBoundingClientRect().top;
      const wrappers = container.querySelectorAll<HTMLElement>('[data-word-page-wrapper]');
      let page = 1;
      wrappers.forEach((wrapper, index) => {
        const top = wrapper.getBoundingClientRect().top - containerTop + container.scrollTop;
        if (top <= container.scrollTop + 80) page = index + 1;
      });
      setCurrentPage(page);
    });
  }, []);

  const changeZoom = useCallback(
    (value: number) => {
      const next = Math.max(0.4, Math.min(2.5, Number(value.toFixed(2))));
      setZoom(next);
      updateView(session.id, { zoom: next });
    },
    [session.id, updateView],
  );

  const fitWidth = useCallback(() => {
    const container = scrollRef.current;
    if (!container) return;
    const available = container.clientWidth - 64;
    changeZoom(available / doc.page.widthPx);
  }, [changeZoom, doc.page.widthPx]);

  const findHits = useMemo<FindHit[]>(() => {
    if (!findQuery) return [];
    const hits: FindHit[] = [];
    for (const paragraph of collectParagraphs(doc)) {
      for (const hit of findInParagraph(paragraph, findQuery, caseSensitive)) {
        hits.push({ paragraphId: paragraph.id, index: hit.index, length: hit.length });
      }
    }
    return hits;
  }, [doc, findQuery, caseSensitive]);

  const goToHit = useCallback(
    (index: number) => {
      if (findHits.length === 0) return;
      const normalized = ((index % findHits.length) + findHits.length) % findHits.length;
      setFindActive(normalized);
      const hit = findHits[normalized];
      const element = findParagraphElement(pagesRef.current, hit.paragraphId);
      if (!element) return;
      selectRangeByOffset(element, hit.index, hit.index + hit.length, false);
      element.scrollIntoView({ block: 'center', behavior: 'smooth' });
    },
    [findHits],
  );

  const selectedParagraphIds = useCallback((): string[] => {
    const container = pagesRef.current;
    const selection = window.getSelection();
    if (!container || !selection || selection.rangeCount === 0) {
      return activeParagraphId ? [activeParagraphId] : [];
    }
    const range = selection.getRangeAt(0);
    if (!container.contains(range.commonAncestorContainer)) {
      return activeParagraphId ? [activeParagraphId] : [];
    }
    const ids: string[] = [];
    container.querySelectorAll<HTMLElement>('[data-paragraph-id]').forEach((element) => {
      if (!range.intersectsNode(element)) return;
      const id = element.dataset.paragraphId;
      if (id && !ids.includes(id)) ids.push(id);
    });
    return ids.length > 0 ? ids : activeParagraphId ? [activeParagraphId] : [];
  }, [activeParagraphId]);

  const selectionInsideEditor = useCallback((): boolean => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return false;
    return !!pagesRef.current?.contains(selection.anchorNode);
  }, []);

  const restoreSelection = useCallback((): boolean => {
    const selection = window.getSelection();
    const container = pagesRef.current;
    if (selection && !selection.isCollapsed && container?.contains(selection.anchorNode)) {
      return true;
    }
    const savedRange = savedRangeRef.current;
    if (
      selection &&
      container &&
      savedRange &&
      savedRange.startContainer.isConnected &&
      savedRange.endContainer.isConnected &&
      container.contains(savedRange.startContainer) &&
      container.contains(savedRange.endContainer)
    ) {
      const node = savedRange.startContainer;
      const element = (node instanceof Element ? node : node.parentElement)?.closest<HTMLElement>(
        '[data-paragraph-id]',
      );
      element?.focus({ preventScroll: true });
      selection.removeAllRanges();
      selection.addRange(savedRange);
      return true;
    }
    const saved = savedSelectionRef.current;
    if (!saved || !container) return selectionInsideEditor();
    const element = findParagraphElement(container, saved.paragraphId);
    if (!element) return false;
    element.focus({ preventScroll: true });
    selectRangeByOffset(element, saved.start, saved.end);
    return true;
  }, [selectionInsideEditor]);

  const triggerSelectionInput = useCallback(() => {
    const container = pagesRef.current;
    if (!container) return;
    const selection = window.getSelection();
    container.querySelectorAll<HTMLElement>('[data-paragraph-id]').forEach((element) => {
      if (selection && selection.rangeCount > 0 && !selection.getRangeAt(0).intersectsNode(element)) return;
      element.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }, []);

  const execCommand = useCallback((command: string, value?: string): boolean => {
    try {
      return document.execCommand(command, false, value);
    } catch {
      return false;
    }
  }, []);

  const applyInlineCommand = useCallback(
    (command: 'bold' | 'italic' | 'underline' | 'strikeThrough') => {
      if (readonly) return;
      if (!restoreSelection()) {
        notify('info', '请先选择文本', '在文档中选择要设置格式的文字。');
        return;
      }
      execCommand('styleWithCSS', 'true');
      const ok = execCommand(command);
      if (!ok) {
        notify('warning', '无法应用格式', '当前浏览器不支持该排版命令。');
        return;
      }
      triggerSelectionInput();
      setActiveStyle((style) => ({
        ...style,
        ...(command === 'bold' ? { bold: document.queryCommandState('bold') } : {}),
        ...(command === 'italic' ? { italic: document.queryCommandState('italic') } : {}),
        ...(command === 'underline' ? { underline: document.queryCommandState('underline') } : {}),
        ...(command === 'strikeThrough' ? { strike: document.queryCommandState('strikeThrough') } : {}),
      }));
    },
    [execCommand, notify, readonly, restoreSelection, triggerSelectionInput],
  );

  const applyFontFormat = useCallback(
    (patch: Partial<RunStyle>) => {
      if (readonly) return;
      if (!restoreSelection()) {
        notify('info', '请先选择文本', '在文档中选择要调整字体的文字。');
        return;
      }
      execCommand('styleWithCSS', 'false');
      if (patch.fontFamily !== undefined) {
        execCommand('fontName', patch.fontFamily || doc.defaultFont);
      }
      if (patch.fontSizePt) {
        execCommand('fontSize', '7');
        pagesRef.current?.querySelectorAll<HTMLFontElement>('font[size="7"]').forEach((font) => {
          font.style.fontSize = `${patch.fontSizePt}pt`;
          font.removeAttribute('size');
        });
      }
      triggerSelectionInput();
      refreshSelection();
    },
    [doc.defaultFont, execCommand, notify, readonly, refreshSelection, restoreSelection, triggerSelectionInput],
  );

  const applyColorFormat = useCallback(
    (patch: Partial<RunStyle>, kind: 'color' | 'highlight') => {
      if (readonly) return;
      if (!restoreSelection()) {
        notify('info', '请先选择文本', '在文档中选择文字后再设置颜色。');
        return;
      }
      execCommand('styleWithCSS', 'true');
      if (kind === 'color') {
        execCommand('foreColor', patch.color ?? '#1D2939');
      } else {
        execCommand('hiliteColor', patch.highlight ?? 'transparent');
      }
      triggerSelectionInput();
      refreshSelection();
    },
    [execCommand, notify, readonly, refreshSelection, restoreSelection, triggerSelectionInput],
  );

  const handleParagraphInput = useCallback(
    (paragraphId: string, element: HTMLElement) => {
      const current = latestDoc();
      const paragraph = findParagraph(current, paragraphId);
      if (!paragraph) return;
      const fallback = baseRunStyle(paragraph, current);
      const runs = serializeDomRuns(element, fallback, paragraphImages(paragraph));
      const next = updateParagraph(current, paragraphId, (existing) => ({ ...existing, runs }));
      commit(next, { key: `type:${paragraphId}` });
      setActiveParagraphId(paragraphId);
      setActiveStyle(runs.find((run) => run.text)?.style ?? fallback);
    },
    [commit, latestDoc],
  );

  const handleParagraphBlur = useCallback(() => {
    setTimeout(() => {
      // 功能：鼠标跨段拖选或段落内移焦时不重新分页和强制移动光标；时间：2026-10-18；作者：lq。
      const selection = window.getSelection();
      if (selection?.rangeCount && !selection.isCollapsed && pagesRef.current?.contains(selection.anchorNode)) return;
      const active = document.activeElement as HTMLElement | null;
      if (active && pagesRef.current?.contains(active)) return;
      const editable = active?.closest<HTMLElement>('[data-paragraph-id]');
      if (editable?.dataset.paragraphId) {
        focusRequestRef.current = {
          paragraphId: editable.dataset.paragraphId,
          offset: getCaretOffset(editable),
        };
      }
      setLayoutRevision((value) => value + 1);
    }, 0);
  }, []);

  const splitParagraph = useCallback(
    (paragraphId: string, element: HTMLElement) => {
      const current = latestDoc();
      const paragraph = findParagraph(current, paragraphId);
      if (!paragraph) return;
      const offset = getCaretOffset(element);
      const fallback = baseRunStyle(paragraph, current);
      const [before, after] = splitRunsAtOffset(paragraph.runs, offset, fallback);
      const nextParagraph = makeParagraph(after, { style: { ...paragraph.style } });
      const updated = updateParagraph(current, paragraphId, (existing) => ({
        ...existing,
        runs: before,
      }));
      const inserted = insertAfterParagraph(updated.blocks, paragraphId, nextParagraph);
      const next: DocModel = {
        ...updated,
        blocks: inserted.inserted ? inserted.blocks : [...updated.blocks, nextParagraph],
      };
      commit(next, { reflow: true });
      focusRequestRef.current = { paragraphId: nextParagraph.id, offset: 0 };
      setActiveParagraphId(nextParagraph.id);
    },
    [commit, latestDoc],
  );

  const mergePrevious = useCallback(
    (paragraphId: string) => {
      const current = latestDoc();
      const index = current.blocks.findIndex((block) => block.type === 'paragraph' && block.id === paragraphId);
      if (index <= 0) return false;
      const previous = current.blocks[index - 1];
      const currentBlock = current.blocks[index];
      if (previous.type !== 'paragraph' || currentBlock.type !== 'paragraph') return false;
      const mergedRuns = [...previous.runs, ...currentBlock.runs];
      const updated = updateParagraph(current, previous.id, (paragraph) => ({
        ...paragraph,
        runs: mergedRuns,
      }));
      const blocks = current.blocks.filter((block) => block.id !== paragraphId);
      const next: DocModel = { ...updated, blocks };
      commit(next, { reflow: true });
      focusRequestRef.current = { paragraphId: previous.id, offset: atomicLength(previous.runs) };
      setActiveParagraphId(previous.id);
      return true;
    },
    [commit, latestDoc],
  );

  const handleKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>, paragraphId: string, element: HTMLDivElement) => {
      if (readonly) return;
      const mod = event.metaKey || event.ctrlKey;
      if (mod && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
        return;
      }
      if (mod && ['b', 'i', 'u'].includes(event.key.toLowerCase())) {
        event.preventDefault();
        applyInlineCommand(
          event.key.toLowerCase() === 'b'
            ? 'bold'
            : event.key.toLowerCase() === 'i'
              ? 'italic'
              : 'underline',
        );
        return;
      }
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        splitParagraph(paragraphId, element);
        return;
      }
      if (event.key === 'Backspace') {
        if (getCaretOffset(element) === 0) {
          if (mergePrevious(paragraphId)) event.preventDefault();
        }
        return;
      }
      if (event.key === 'Tab') {
        event.preventDefault();
        execCommand('insertText', '    ');
      }
    },
    [applyInlineCommand, execCommand, mergePrevious, readonly, redo, splitParagraph, undo],
  );

  const handlePaste = useCallback((event: React.ClipboardEvent<HTMLDivElement>) => {
    const text = event.clipboardData.getData('text/plain');
    if (!text) return;
    event.preventDefault();
    execCommand('insertText', text.replace(/\r?\n/g, ' '));
  }, [execCommand]);

  /** 功能：跨段选区始终在模型层替换，随后统一分页并恢复插入点；时间：2026-10-18；作者：lq。 */
  const replaceCrossSelection = useCallback((text: string): boolean => {
    const container = pagesRef.current;
    const selection = window.getSelection();
    if (readonly || !container || !selection?.rangeCount) return false;
    const slices = selectedParagraphSlices(container, selection.getRangeAt(0));
    if (slices.length < 2) return false;
    const result = replaceDocumentSelection(latestDoc(), slices, text);
    if (!result) return false;
    savedRangeRef.current = null;
    savedSelectionRef.current = null;
    selection.removeAllRanges();
    focusRequestRef.current = { paragraphId: result.paragraphId, offset: result.offset };
    commit(result.doc, { reflow: true });
    setSyncToken((value) => value + 1);
    return true;
  }, [commit, latestDoc, readonly]);

  /** 功能：键盘全选限定正文，跨段删除/输入由模型接管；时间：2026-10-18；作者：lq。 */
  const handleDocumentKeyDown = useCallback((event: ReactKeyboardEvent<HTMLDivElement>) => {
    const root = pagesRef.current;
    if (!root) return;
    // 功能：Tab 在表格单元格间导航，末格 Tab 自动追加一行；时间：2026-10-19；作者：lq。
    if (!readonly && event.key === 'Tab') {
      const node = window.getSelection()?.anchorNode;
      const cell = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('[data-cell-id]');
      const tableElement = cell?.closest<HTMLElement>('[data-table-id]');
      const current = latestDoc();
      const table = tableElement ? findTable(current.blocks, tableElement.dataset.tableId!) : undefined;
      if (table && cell) {
        event.preventDefault(); event.stopPropagation();
        const placements = tableCellPlacements(table);
        const index = placements.findIndex((p) => p.cell.id === cell.dataset.cellId);
        let target: DocTableCell | undefined = placements[index + (event.shiftKey ? -1 : 1)]?.cell;
        if (!target && !event.shiftKey) {
          const last = placements[index];
          const next = editTable(table, { tableId: table.id, anchorId: last.cell.id, focusId: last.cell.id }, 'row-below');
          if (next) {
            target = next.rows.at(-1)?.cells[0];
            commit({ ...current, blocks: replaceTableInBlocks(current.blocks, table.id, () => next) }, { reflow: true });
          }
        }
        const paragraph = target?.blocks.find((block): block is DocParagraph => block.type === 'paragraph');
        if (paragraph) {
          focusRequestRef.current = { paragraphId: paragraph.id, offset: 0 };
          setTableSelection({ tableId: table.id, anchorId: target!.id, focusId: target!.id });
          setSyncToken((value) => value + 1);
        }
        return;
      }
    }
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
      const paragraphs = root.querySelectorAll<HTMLElement>('[data-paragraph-id]');
      if (!paragraphs.length) return;
      event.preventDefault();
      event.stopPropagation();
      const range = document.createRange();
      range.setStart(paragraphs[0], 0);
      range.setEnd(paragraphs[paragraphs.length - 1], paragraphs[paragraphs.length - 1].childNodes.length);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      captureSelection();
      return;
    }
    if (!readonly && ['Backspace', 'Delete', 'Enter'].includes(event.key) && replaceCrossSelection(event.key === 'Enter' ? '\n' : '')) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const node = window.getSelection()?.anchorNode;
    const paragraph = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLDivElement>('[data-paragraph-id]');
    if (paragraph && root.contains(paragraph)) handleKeyDown(event, paragraph.dataset.paragraphId!, paragraph);
  }, [captureSelection, commit, handleKeyDown, latestDoc, readonly, replaceCrossSelection]);

  /** 功能：beforeinput 覆盖键入、触屏删除及跨段替换，阻止浏览器合并分页容器；时间：2026-10-18；作者：lq。 */
  useEffect(() => {
    const root = pagesRef.current;
    if (!root) return;
    const beforeInput = (event: InputEvent) => {
      if (!event.cancelable || readonly) return;
      if (event.inputType.startsWith('insert') && event.inputType !== 'insertFromPaste') {
        const text = ['insertParagraph', 'insertLineBreak'].includes(event.inputType) ? '\n' : event.data;
        if (text !== null && replaceCrossSelection(text)) event.preventDefault();
      } else if (event.inputType.startsWith('delete') && replaceCrossSelection('')) event.preventDefault();
    };
    root.addEventListener('beforeinput', beforeInput);
    return () => root.removeEventListener('beforeinput', beforeInput);
  }, [readonly, replaceCrossSelection]);

  /** 功能：跨段复制/剪切保留正文文本和文字样式，取消系统对分页结构的直接删除；时间：2026-10-18；作者：lq。 */
  const handleDocumentCopy = useCallback((event: React.ClipboardEvent<HTMLDivElement>, cut = false) => {
    const root = pagesRef.current;
    const selection = window.getSelection();
    if (!root || !selection?.rangeCount) return;
    const slices = selectedParagraphSlices(root, selection.getRangeAt(0));
    if (slices.length < 2) return;
    event.preventDefault();
    const current = latestDoc();
    event.clipboardData.setData('text/plain', selectionText(current, slices));
    event.clipboardData.setData('text/html', selectionHtml(current, slices));
    if (cut && !readonly) replaceCrossSelection('');
  }, [latestDoc, readonly, replaceCrossSelection]);

  const applyParagraphFormat = useCallback(
    (patch: Partial<ParagraphStyle>) => {
      if (readonly) return;
      const ids = selectedParagraphIds();
      if (ids.length === 0) {
        notify('info', '没有选中的段落', '请将光标放在要调整的段落中。');
        return;
      }
      let next = latestDoc();
      for (const id of ids) {
        next = updateParagraph(next, id, (paragraph) => ({
          ...paragraph,
          style: { ...paragraph.style, ...patch },
        }));
      }
      commit(next, { reflow: true, key: `paragraph:${patch.alignment ?? patch.lineHeight ?? 'format'}` });
    },
    [commit, latestDoc, notify, readonly, selectedParagraphIds],
  );

  const updatePage = useCallback(
    (page: DocModel['page']) => {
      const current = latestDoc();
      commit({ ...current, page }, { reflow: true });
    },
    [commit, latestDoc],
  );

  /** 功能：在正文或表格当前光标插入原子合并域，统一撤销且恢复光标；时间：2026-10-24；作者：lq。 */
  const insertMergeField = useCallback((field: string) => {
    if (readonly) return;
    captureSelection();
    const current = latestDoc();
    const saved = savedSelectionRef.current;
    const id = saved?.paragraphId ?? activeParagraphId;
    const paragraph = id ? findParagraph(current, id) : undefined;
    if (!paragraph) { notify('info', '请先选择插入位置', '在正文或表格单元格中点击，再点击右侧字段。'); return; }
    const start = saved && saved.paragraphId === id ? saved.start : runsPlainText(paragraph.runs).length;
    const end = saved && saved.paragraphId === id ? saved.end : start;
    const style = styleAtOffset(paragraph, start);
    const [before] = splitRunsAtOffset(paragraph.runs, start, style);
    const [, after] = splitRunsAtOffset(paragraph.runs, end, style);
    const next = updateParagraph(current, paragraph.id, (value) => ({ ...value, runs: [...before, { text: `«${field}»`, mergeField: field, style }, ...after] }));
    commit(next, { reflow: true });
    savedSelectionRef.current = { paragraphId: paragraph.id, start: start + 1, end: start + 1 };
    focusRequestRef.current = { paragraphId: paragraph.id, offset: start + 1 };
    setSyncToken((value) => value + 1);
  }, [activeParagraphId, captureSelection, commit, latestDoc, notify, readonly]);

  // 功能：校验行列尺寸并按页面内容宽度生成等宽表格；时间：2026-10-19；作者：lq。
  const insertTable = useCallback((rows: number, columns: number) => {
    if (readonly || !Number.isInteger(rows) || !Number.isInteger(columns)
      || rows < 1 || rows > 100 || columns < 1 || columns > 63) return;
    const current = latestDoc();
    const columnWidth = metrics.contentWidth / columns;
    const table: DocTable = {
      type: 'table',
      id: uid('tbl'),
      columnWidthsPx: Array.from({ length: columns }, () => columnWidth),
      alignment: 'left',
      borders: {
        top: { style: 'single', widthPx: 0.75, color: '#1D2939' },
        bottom: { style: 'single', widthPx: 0.75, color: '#1D2939' },
        left: { style: 'single', widthPx: 0.75, color: '#1D2939' },
        right: { style: 'single', widthPx: 0.75, color: '#1D2939' },
        insideH: { style: 'single', widthPx: 0.75, color: '#1D2939' },
        insideV: { style: 'single', widthPx: 0.75, color: '#1D2939' },
      },
      rows: Array.from({ length: rows }, () => ({
        id: uid('row'),
        cells: Array.from({ length: columns }, () => ({
          id: uid('cell'),
          blocks: [makeParagraph()],
          colSpan: 1,
          rowSpan: 1,
        })),
      })),
    };
    const index = topLevelBlockIndex(current, activeParagraphId);
    const blocks = [...current.blocks];
    blocks.splice(Math.max(0, index) + 1, 0, table);
    commit({ ...current, blocks }, { reflow: true });
    setTableMenu(null);
  }, [activeParagraphId, commit, latestDoc, metrics.contentWidth, readonly]);

  const insertPageBreak = useCallback(() => {
    if (readonly) return;
    const current = latestDoc();
    const index = topLevelBlockIndex(current, activeParagraphId);
    const blocks = [...current.blocks];
    blocks.splice(Math.max(0, index) + 1, 0, { type: 'pageBreak', id: uid('break') });
    commit({ ...current, blocks }, { reflow: true });
  }, [activeParagraphId, commit, latestDoc, readonly]);

  const handleImageFile = useCallback(
    (file: File) => {
      if (readonly) return;
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = String(reader.result ?? '');
        const image = new Image();
        image.onload = () => {
          const maxWidth = metrics.contentWidth * 0.8;
          const scale = Math.min(1, maxWidth / Math.max(1, image.naturalWidth));
          const widthPx = image.naturalWidth * scale;
          const heightPx = image.naturalHeight * scale;
          const docImage: DocImage = {
            id: uid('img'),
            src: dataUrl,
            dataUrl,
            widthPx,
            heightPx,
            name: file.name,
            alt: file.name,
          };
          const current = latestDoc();
          const paragraph = makeParagraph([{ text: '', style: {}, image: docImage }], {
            style: { alignment: 'center' },
          });
          const index = topLevelBlockIndex(current, activeParagraphId);
          const blocks = [...current.blocks];
          blocks.splice(Math.max(0, index) + 1, 0, paragraph);
          commit({ ...current, blocks }, { reflow: true });
          setActiveParagraphId(paragraph.id);
        };
        image.onerror = () => notify('error', '图片读取失败', '无法识别该图片文件。');
        image.src = dataUrl;
      };
      reader.onerror = () => notify('error', '图片读取失败', file.name);
      reader.readAsDataURL(file);
    },
    [activeParagraphId, commit, latestDoc, metrics.contentWidth, notify, readonly],
  );

  const applyList = useCallback(
    (ordered: boolean) => {
      if (readonly) return;
      const ids = selectedParagraphIds();
      if (ids.length === 0) {
        notify('info', '请选择段落', '选择要设置为列表的段落。');
        return;
      }
      let next = latestDoc();
      const allSame = ids.every((id) => {
        const paragraph = findParagraph(next, id);
        return paragraph?.list?.ordered === ordered;
      });
      let counter = 0;
      for (const id of ids) {
        counter += 1;
        next = updateParagraph(next, id, (paragraph) => {
          if (allSame) return { ...paragraph, list: undefined };
          return {
            ...paragraph,
            list: {
              level: 0,
              marker: ordered ? `${counter}.` : '•',
              ordered,
            },
            style: {
              ...paragraph.style,
              indentLeftPx: Math.max(paragraph.style.indentLeftPx ?? 0, 24),
            },
          };
        });
      }
      commit(next, { reflow: true });
    },
    [commit, latestDoc, notify, readonly, selectedParagraphIds],
  );

  const handleTableContextMenu = useCallback(
    (event: React.MouseEvent, table: DocTable, rowIndex: number, cellId: string) => {
      if (readonly) return;
      event.preventDefault();
      // 功能：右键保留已有矩形选区，否则定位当前单元格；时间：2026-10-19；作者：lq。
      explicitTableSelectionRef.current = true;
      const rect = tableSelection?.tableId === table.id ? tableSelectionRect(table, tableSelection) : undefined;
      const placement = tableCellPlacements(table).find((p) => p.cell.id === cellId);
      if (!rect || !placement || !intersects(placement, rect)) setTableSelection({ tableId: table.id, anchorId: cellId, focusId: cellId });
      setTableMenu({ x: event.clientX, y: event.clientY, tableId: table.id, rowIndex });
    },
    [readonly, tableSelection],
  );

  const mutateTable = useCallback(
    (tableId: string, updater: (table: DocTable) => DocTable | null) => {
      const current = latestDoc();
      // 功能：结构编辑刷新段落 DOM，避免合并后仍显示旧内容；时间：2026-10-19；作者：lq。
      const blocks = replaceTableInBlocks(current.blocks, tableId, updater);
      commit({ ...current, blocks: blocks.length ? blocks : [makeParagraph()] }, { reflow: true });
      setSyncToken((value) => value + 1);
    },
    [commit, latestDoc],
  );

  /** 功能：执行表格命令并在结构变化后保留有效选区；时间：2026-10-19；作者：lq。 */
  const runTableAction = useCallback((action: TableAction) => {
    if (readonly || !tableSelection) return;
    if (action.startsWith('select-')) {
      const table = findTable(latestDoc().blocks, tableSelection.tableId);
      const rect = table && tableSelectionRect(table, tableSelection);
      if (!table || !rect) return;
      const area = { top: action === 'select-rows' ? rect.top : 0, bottom: action === 'select-rows' ? rect.bottom : table.rows.length, left: action === 'select-columns' ? rect.left : 0, right: action === 'select-columns' ? rect.right : totalTableColumns(table) };
      const cells = tableCellPlacements(table).filter((p) => intersects(p, area));
      if (cells.length) {
        explicitTableSelectionRef.current = true;
        window.getSelection()?.removeAllRanges();
        setTableSelection({ tableId: table.id, anchorId: cells[0].cell.id, focusId: cells.at(-1)!.cell.id });
      }
      return;
    }
    let nextSelection: TableSelection | null = null;
    let focusParagraph: string | undefined;
    mutateTable(tableSelection.tableId, (table) => {
      const rect = tableSelectionRect(table, tableSelection);
      const next = editTable(table, tableSelection, action, metrics.contentWidth);
      if (next) {
        const placements = tableCellPlacements(next);
        const cell = placements.find((p) => p.cell.id === tableSelection.anchorId) ?? placements.find((p) => p.row >= Math.min(rect?.top ?? 0, next.rows.length - 1)) ?? placements[0];
        if (cell) {
          nextSelection = { tableId: next.id, anchorId: cell.cell.id, focusId: cell.cell.id };
          focusParagraph = cell.cell.blocks.find((block): block is DocParagraph => block.type === 'paragraph')?.id;
        }
      }
      return next;
    });
    // 功能：结构操作后将光标放回有效单元格，允许继续输入；时间：2026-10-19；作者：lq。
    if (focusParagraph) focusRequestRef.current = { paragraphId: focusParagraph, offset: 0 };
    explicitTableSelectionRef.current = true;
    window.getSelection()?.removeAllRanges();
    savedRangeRef.current = null;
    savedSelectionRef.current = null;
    setTableSelection(nextSelection);
    setTableMenu(null);
  }, [readonly, tableSelection, latestDoc, mutateTable, metrics.contentWidth]);

  /** 功能：拖动边缘时即时预览，松开仅产生一次可撤销尺寸变更；时间：2026-10-19；作者：lq。 */
  const handleTableResize = useCallback((event: React.PointerEvent, table: DocTable, index: number, axis: 'row' | 'column') => {
    if (readonly || event.button !== 0) return;
    event.preventDefault(); event.stopPropagation();
    const grid = event.currentTarget.closest<HTMLElement>('[data-table-id]');
    if (!grid) return;
    const start = axis === 'column' ? event.clientX : event.clientY;
    const tracks = axis === 'column' ? Array.from({ length: totalTableColumns(table) }, (_, c) => table.columnWidthsPx[c] ?? 100) : getComputedStyle(grid).gridTemplateRows.split(' ').map(Number.parseFloat);
    const originalStyle = grid.getAttribute('style');
    const initial = tracks[index] || (axis === 'column' ? 100 : 28);
    let value = initial;
    const move = (pointer: PointerEvent) => {
      const max = axis === 'column' ? (tracks[index + 1] !== undefined ? initial + tracks[index + 1] - 24 : Math.max(initial, metrics.contentWidth - tracks.reduce((sum, width, i) => sum + (i === index ? 0 : width), 0))) : 1900;
      value = Math.min(max, Math.max(axis === 'column' ? 24 : 16, initial + ((axis === 'column' ? pointer.clientX : pointer.clientY) - start) / effectiveZoom));
      const preview = [...tracks]; preview[index] = value;
      if (axis === 'column' && preview[index + 1] !== undefined) preview[index + 1] = tracks[index + 1] + initial - value;
      if (axis === 'column') { grid.style.gridTemplateColumns = preview.map((w) => `${w}px`).join(' '); grid.style.width = `${preview.reduce((a, b) => a + b, 0)}px`; }
      else grid.style.gridTemplateRows = preview.map((h) => `minmax(${h}px, auto)`).join(' ');
    };
    const finish = (pointer: PointerEvent) => {
      window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', finish); window.removeEventListener('pointercancel', finish);
      if (originalStyle !== null) grid.setAttribute('style', originalStyle);
      if (pointer.type === 'pointercancel' || value === initial) return;
      mutateTable(table.id, (current) => {
        // 功能：只变更指定逻辑行列，保留其余尺寸与合并关系；时间：2026-10-19；作者：lq。
        const widths = [...current.columnWidthsPx];
        if (axis === 'column') {
          widths[index] = value;
          if (widths[index + 1] !== undefined) widths[index + 1] = tracks[index + 1] + initial - value;
        }
        return { ...current, widthPct: axis === 'column' ? undefined : current.widthPct, columnWidthsPx: widths, rows: current.rows.map((row, r) => ({ ...row, heightPx: axis === 'row' && r === index ? value : row.heightPx, cells: row.cells.map((cell) => ({ ...cell, widthPx: undefined })) })) };
      });
    };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', finish); window.addEventListener('pointercancel', finish);
  }, [readonly, effectiveZoom, metrics.contentWidth, mutateTable]);

  const activeTable = tableSelection ? findTable(doc.blocks, tableSelection.tableId) : undefined;

  const fontFallback = useMemo(() => {
    const fallbackParagraph =
      activeParagraph ??
      (doc.blocks.find((block): block is DocParagraph => block.type === 'paragraph') ?? makeParagraph());
    return baseRunStyle(fallbackParagraph, doc);
  }, [activeParagraph, doc]);

  const paragraphHandlers: ParagraphHandlers = useMemo(
    () => ({
      syncToken,
      onInput: handleParagraphInput,
      onBlur: handleParagraphBlur,
      onKeyDown: handleKeyDown,
      onTableContextMenu: handleTableContextMenu,
      tableSelection,
      onTableResize: handleTableResize,
      onPaste: handlePaste,
    }),
    [handleParagraphBlur, handleParagraphInput, handleKeyDown, handlePaste, handleTableContextMenu, handleTableResize, tableSelection, syncToken],
  );

  const handleReplaceCurrent = useCallback(() => {
    if (readonly || findHits.length === 0) return;
    const hit = findHits[Math.min(findActive, findHits.length - 1)];
    const current = latestDoc();
    let replaced = false;
    const next = updateParagraph(current, hit.paragraphId, (paragraph) => {
      const result = replaceInRuns(paragraph.runs, findQuery, replacement, 1, caseSensitive);
      if (result.count > 0) replaced = true;
      return { ...paragraph, runs: result.runs };
    });
    if (replaced) commit(next, { reflow: true });
  }, [caseSensitive, commit, findActive, findHits, findQuery, latestDoc, readonly, replacement]);

  const handleReplaceAll = useCallback(() => {
    if (readonly || !findQuery) return;
    const current = latestDoc();
    let total = 0;
    let next = current;
    for (const paragraph of collectParagraphs(current)) {
      const result = replaceInRuns(paragraph.runs, findQuery, replacement, Number.POSITIVE_INFINITY, caseSensitive);
      if (result.count > 0) {
        total += result.count;
        next = updateParagraph(next, paragraph.id, (item) => ({ ...item, runs: result.runs }));
      }
    }
    if (total > 0) {
      commit(next, { reflow: true });
      notify('success', `已替换 ${total} 处`, replacement ? `替换为“${replacement}”` : '已删除匹配内容');
    } else {
      notify('info', '没有可替换的内容');
    }
  }, [caseSensitive, commit, findQuery, latestDoc, notify, readonly, replacement]);

  const handleRepaginate = useCallback(() => setLayoutRevision((value) => value + 1), []);

  const pageElements = useCallback(() => {
    const container = pagesRef.current;
    if (!container) return [] as HTMLElement[];
    return Array.from(container.querySelectorAll<HTMLElement>('[data-word-page]'));
  }, []);

  const handleExportPdf = useCallback(async () => {
    if (exporting) return;
    const pagesToExport = pageElements();
    if (pagesToExport.length === 0) {
      notify('warning', '没有可导出的页面');
      return;
    }
    cancelExportRef.current = false;
    setExporting(true);
    setExportProgress({ phase: 'preparing', current: 0, total: pagesToExport.length, message: '正在准备导出…' });
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    try {
      const bytes = await exportElementsToPdf({
        pages: pageElements(),
        fileName: ensureExtension(stripExtension(exportName || session.displayName), '.pdf'),
        title: session.displayName,
        pageRange: exportRange,
        // 功能：使用导出服务默认 300 DPI 无损画质；时间：2026-10-19；作者：lq。
        pageWidthPx: doc.page.widthPx,
        pageHeightPx: doc.page.heightPx,
        onProgress: setExportProgress,
        isCancelled: () => cancelExportRef.current,
      });
      await exportBytes(
        bytes,
        ensureExtension(stripExtension(exportName || session.displayName), '.pdf'),
        'application/pdf',
      );
      setExportOpen(false);
      setExportProgress(null);
    } catch (error) {
      if ((error as { code?: string })?.code === 'CANCELLED') {
        notify('info', '已取消导出');
      } else {
        notify('error', '导出 PDF 失败', error instanceof Error ? error.message : String(error));
      }
      setExportProgress(null);
    } finally {
      setExporting(false);
    }
  }, [doc.page.heightPx, doc.page.widthPx, exportBytes, exportName, exportRange, exporting, notify, pageElements, session.displayName]);

  useEffect(() => {
    setExportName(`${stripExtension(session.displayName)}.pdf`);
  }, [session.displayName]);

  // 功能：右键菜单与工具栏共用安全的网格编辑命令；时间：2026-10-19；作者：lq。
  const tableMenuItems: MenuItem[] = tableMenu && tableSelection && activeTable
    ? ([['merge', '合并选中单元格'], ['split', '拆分已合并单元格'], ['row-above', '在上方插入行'], ['row-below', '在下方插入行'], ['column-before', '在左侧插入列'], ['column-after', '在右侧插入列'], ['row-delete', '删除选中行'], ['column-delete', '删除选中列'], ['table-delete', '删除表格']] as [TableAction, string][]).map(([id, label]) => ({ id, label, danger: id.endsWith('delete'), onSelect: () => runTableAction(id) }))
    : [];

  return (
    <div className="flex h-full min-h-0 w-full">
      <div className="flex min-w-0 flex-1 flex-col">
        {settings.showToolbar ? (
        <WordToolbar
          doc={doc}
          readonly={readonly}
          activeParagraph={activeParagraph}
          activeStyle={activeParagraph ? activeStyle : fontFallback}
          canUndo={canUndo}
          canRedo={canRedo}
          zoom={zoom}
          pageCount={pages.length}
          currentPage={currentPage}
          findOpen={findOpen}
          propertiesOpen={settings.showProperties}
          onInlineCommand={applyInlineCommand}
          onColorFormat={applyColorFormat}
          onFontFormat={applyFontFormat}
          onParagraphFormat={applyParagraphFormat}
          onInsertTable={insertTable}
          onInsertImage={() => imageInputRef.current?.click()}
          onPageBreak={insertPageBreak}
          onHeaderFooter={() => setHeaderFooterRegion('footer')}
          onMailMerge={() => setMailMergeOpen((value) => !value)}
          onUndo={undo}
          onRedo={redo}
          onZoomChange={changeZoom}
          onFitWidth={fitWidth}
          onToggleFind={() => setFindOpen((open) => !open)}
          onToggleProperties={() => updateSettings({ showProperties: !settings.showProperties })}
          onExport={() => {
            setExportName(`${stripExtension(session.displayName)}.pdf`);
            setExportRange('');
            setExportOpen(true);
          }}
          onPrint={() => void (async () => {
            const ok = await printCurrentView();
            if (!ok) notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
          })()}
          onShowCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
          onAddBullets={() => applyList(false)}
          onAddNumbering={() => applyList(true)}
          onBeforeInteraction={captureSelection}
        />
        ) : null}
        {/* 功能：选中表格自动显示专属工具；时间：2026-10-19；作者：lq。 */}
        {!readonly && activeTable && tableSelection ? <TableTools table={activeTable} selection={tableSelection} onAction={runTableAction} onFormat={(format) => mutateTable(activeTable.id, (table) => formatTable(table, tableSelection, format, metrics.contentWidth))} /> : null}
        <div
          ref={scrollRef}
          onScroll={handleScroll}
          className="mo-word-scroll relative flex-1 overflow-auto bg-[var(--mo-canvas)]"
        >
          <div
            ref={pagesRef}
            data-word-editor
            contentEditable={!readonly}
            suppressContentEditableWarning
            role="textbox"
            aria-label="文档正文"
            aria-multiline="true"
            tabIndex={0}
            // 功能：Shift 点击扩展单元格选区，普通点击定位表格；时间：2026-10-19；作者：lq。
            onPointerDownCapture={(event) => {
              if (readonly || event.button !== 0 || (event.target as HTMLElement).closest('[data-table-resize]')) return;
              const cell = (event.target as HTMLElement).closest<HTMLElement>('[data-cell-id]');
              const table = cell?.closest<HTMLElement>('[data-table-id]');
              if (!cell || !table) { explicitTableSelectionRef.current = false; setTableSelection(null); return; }
              if (event.shiftKey && tableSelection && tableSelection.tableId === table.dataset.tableId) {
                event.preventDefault();
                explicitTableSelectionRef.current = true;
                window.getSelection()?.removeAllRanges();
                setTableSelection({ ...tableSelection, focusId: cell.dataset.cellId! });
              } else {
                explicitTableSelectionRef.current = false;
                setTableSelection({ tableId: table.dataset.tableId!, anchorId: cell.dataset.cellId!, focusId: cell.dataset.cellId! });
              }
            }}
            onKeyDown={handleDocumentKeyDown}
            onInput={(event) => {
              // 功能：共享编辑区的原生 input 事件按当前段落写回模型；时间：2026-10-18；作者：lq。
              if (event.target !== event.currentTarget) return;
              const node = window.getSelection()?.anchorNode;
              const paragraph = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('[data-paragraph-id]');
              if (paragraph && event.currentTarget.contains(paragraph)) handleParagraphInput(paragraph.dataset.paragraphId!, paragraph);
            }}
            onPaste={(event) => {
              if (readonly) { event.preventDefault(); return; }
              if (replaceCrossSelection(event.clipboardData.getData('text/plain'))) event.preventDefault();
              else handlePaste(event);
            }}
            onCopy={(event) => handleDocumentCopy(event)}
            onCut={(event) => handleDocumentCopy(event, true)}
            onBlur={handleParagraphBlur}
            className="mo-selectable outline-none mo-word-pages mx-auto flex min-h-full w-fit flex-col items-center gap-6 px-8 py-8"
          >
            {pages.map((page, index) => (
              <WordPageView
                key={page.index}
                page={page}
                index={index}
                total={pages.length}
                blockMap={blockMap}
                doc={page.section ? { ...doc, ...page.section, sectionPageCount: page.sectionTotal } : doc}
                zoom={effectiveZoom}
                readonly={readonly || exporting}
                handlers={paragraphHandlers}
                onEditHeaderFooter={setHeaderFooterRegion}
              />
            ))}
          </div>
          {findOpen ? (
            <FindBar
              query={findQuery}
              replacement={replacement}
              caseSensitive={caseSensitive}
              count={findHits.length}
              active={findActive}
              readonly={readonly}
              onQuery={(value) => {
                setFindQuery(value);
                setFindActive(0);
              }}
              onReplacement={setReplacement}
              onCaseSensitive={setCaseSensitive}
              onNext={() => goToHit(findActive + 1)}
              onPrev={() => goToHit(findActive - 1)}
              onReplace={handleReplaceCurrent}
              onReplaceAll={handleReplaceAll}
              onClose={() => setFindOpen(false)}
            />
          ) : null}
        </div>

        <div
          ref={measureRef}
          aria-hidden
          style={{
            position: 'absolute',
            left: -100000,
            top: 0,
            width: metrics.contentWidth,
            visibility: 'hidden',
            pointerEvents: 'none',
          }}
        >
          {doc.blocks.map((block) =>
            block.type === 'pageBreak' ? null : (
              <div key={block.id} data-measure-id={block.id} style={{ display: 'flow-root' }}>
                {block.type === 'paragraph' ? (
                  <StaticParagraph
                    paragraph={block}
                    doc={doc}
                    pageNumber={1}
                    totalPages={pages.length}
                  />
                ) : (
                  <TableGrid
                    table={block}
                    doc={doc}
                    pageNumber={1}
                    totalPages={pages.length}
                    readonly
                    staticMode
                  />
                )}
              </div>
            ),
          )}
        </div>
      </div>

      {/* 功能：侧栏隐藏时保留本次数据快照和选择状态；时间：2026-10-24；作者：lq。 */}
      <MailMergePanel open={mailMergeOpen && !readonly} session={session} doc={doc} onClose={() => setMailMergeOpen(false)} onInsertField={insertMergeField} renderBlocks={renderMergeBlocks} />
      {settings.showProperties && !mailMergeOpen ? (
        <PropertiesPanel
          doc={doc}
          session={session}
          activeParagraph={activeParagraph}
          pageCount={pages.length}
          onPageChange={updatePage}
          onParagraphChange={applyParagraphFormat}
          onShowCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
          onClose={() => updateSettings({ showProperties: false })}
        />
      ) : null}

      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml"
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) handleImageFile(file);
          event.target.value = '';
        }}
      />

      {tableMenu ? (
        <ContextMenu
          x={tableMenu.x}
          y={tableMenu.y}
          items={tableMenuItems}
          onClose={() => setTableMenu(null)}
        />
      ) : null}

      {/* 功能：应用设置作为单次可撤销编辑，取消时不修改正文；时间：2026-10-24；作者：lq。 */}
      {headerFooterRegion && !readonly ? <HeaderFooterDialog doc={doc} pageCount={pages.length} initialRegion={headerFooterRegion} onClose={() => setHeaderFooterRegion(null)} onApply={(patch) => { commit({ ...latestDoc(), ...patch }, { reflow: true }); setHeaderFooterRegion(null); }} /> : null}
      <Dialog
        open={exportOpen}
        onClose={() => {
          if (!exporting) setExportOpen(false);
        }}
        title="导出为 PDF"
        description="使用当前分页版面生成 PDF，页数与屏幕预览保持一致。"
        width="sm"
        footer={
          <>
            <Button onClick={() => setExportOpen(false)} disabled={exporting}>
              取消
            </Button>
            <Button variant="primary" loading={exporting} onClick={handleExportPdf}>
              开始导出
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <Field label="文件名">
            <Input value={exportName} onChange={(event) => setExportName(event.target.value)} />
          </Field>
          <Field label="页码范围" hint={`留空导出全部 ${pages.length} 页，例如 1-3,5`}>
            <Input
              value={exportRange}
              onChange={(event) => setExportRange(event.target.value)}
              placeholder={`1-${pages.length}`}
              disabled={exporting}
            />
          </Field>
          <div className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
            纸张：{doc.page.sizeName} · {doc.page.orientation === 'portrait' ? '纵向' : '横向'} ·{' '}
            {Math.round(pxToMm(doc.page.widthPx))} × {Math.round(pxToMm(doc.page.heightPx))} mm
          </div>
          {exportProgress ? (
            <div className="space-y-1.5">
              <ProgressBar value={exportProgress.total > 0 ? exportProgress.current / exportProgress.total : 0} />
              <div className="flex justify-between text-[10px] text-[var(--mo-text-muted)]">
                <span>{exportProgress.message}</span>
                <span>
                  {exportProgress.current}/{exportProgress.total}
                </span>
              </div>
            </div>
          ) : null}
          {exporting ? (
            <Button
              size="sm"
              variant="ghost"
              className="text-[var(--mo-danger)]"
              onClick={() => {
                cancelExportRef.current = true;
              }}
            >
              取消导出
            </Button>
          ) : null}
        </div>
      </Dialog>

      <style>{`
        /* 功能：导出打印隐藏表格选区与交互柄；时间：2026-10-19；作者：lq。 */
        @media print {
          [data-table-resize] { display: none !important; }
          [data-cell-id] { box-shadow: none !important; }
          body * { visibility: hidden !important; }
          .mo-word-pages, .mo-word-pages * { visibility: visible !important; }
          .mo-word-pages {
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            display: block !important;
            width: 100% !important;
            padding: 0 !important;
            gap: 0 !important;
          }
          .mo-word-page-scale { width: auto !important; height: auto !important; }
          .mo-word-page {
            position: relative !important;
            left: auto !important;
            top: auto !important;
            transform: none !important;
            box-shadow: none !important;
            margin: 0 auto !important;
            page-break-after: always;
            break-after: page;
          }
          /* 功能：仅最后一张纸取消分页，避免每页都是包装层末元素而失去分页；时间：2026-10-24；作者：lq。 */
          .mo-word-page-scale:last-child .mo-word-page { page-break-after: auto; break-after: auto; }
          .mo-word-scroll { overflow: visible !important; }
          /* 功能：打印纸张跟随文档尺寸，保持页眉页脚位置；时间：2026-10-24；作者：lq。 */
          @page { size: ${doc.page.widthPx}px ${doc.page.heightPx}px; margin: 0; }
        }
      `}</style>
    </div>
  );
}

function WordPageView({
  page,
  index,
  total,
  blockMap,
  doc,
  zoom,
  readonly,
  handlers,
  onEditHeaderFooter,
}: {
  page: PageSlice;
  index: number;
  total: number;
  blockMap: Map<string, DocBlock>;
  doc: DocModel;
  zoom: number;
  readonly: boolean;
  handlers: ParagraphHandlers;
  /** 功能：页面双击进入页眉或页脚编辑；时间：2026-10-24；作者：lq。 */
  onEditHeaderFooter: (region: 'header' | 'footer') => void;
}) {
  const pageNumber = (page.sectionPage ?? index) + (doc.page.pageNumberStart ?? 1);
  const blocks = page.blocks
    .map((block) => blockMap.get(block.id) ?? block)
    .filter((block) => block.type !== 'pageBreak');
  const margins = doc.page.margins;
  const contentWidth = Math.max(120, doc.page.widthPx - margins.left - margins.right);
  return (
    <div
      data-word-page-wrapper
      className="mo-word-page-scale relative shrink-0"
      style={{ width: doc.page.widthPx * zoom, height: doc.page.heightPx * zoom }}
    >
      <div
        data-word-page={index + 1}
        onDoubleClick={(event) => {
          // 功能：保留正文双击选词，仅在上下页边距打开设置；时间：2026-10-24；作者：lq。
          if (readonly) return;
          const y = (event.clientY - event.currentTarget.getBoundingClientRect().top) / zoom;
          if (y < margins.top) onEditHeaderFooter('header');
          else if (y > doc.page.heightPx - margins.bottom) onEditHeaderFooter('footer');
        }}
        className="mo-word-page mo-paper"
        style={{
          width: doc.page.widthPx,
          height: doc.page.heightPx,
          transform: `scale(${zoom})`,
          transformOrigin: 'top left',
          position: 'absolute',
          left: 0,
          top: 0,
          boxShadow: 'var(--mo-shadow-page)',
          overflow: 'hidden',
        }}
      >
        {doc.header ? (
          <div
            style={{
              position: 'absolute',
              left: margins.left,
              right: margins.right,
              top: margins.header,
              maxHeight: Math.max(24, margins.top - margins.header),
              overflow: 'hidden',
            }}
          >
            <HeaderFooterView
              blocks={doc.header.blocks}
              doc={doc}
              pageNumber={pageNumber}
              totalPages={total}
            />
          </div>
        ) : null}
        <div
          style={{
            position: 'absolute',
            left: margins.left,
            top: margins.top,
            width: contentWidth,
          }}
        >
          {blocks.map((block) =>
            block.type === 'paragraph' ? (
              <EditableParagraph
                key={block.id}
                paragraph={block}
                doc={doc}
                pageNumber={pageNumber}
                totalPages={total}
                readonly={readonly}
                handlers={handlers}
              />
            ) : block.type === 'table' ? (
              <TableGrid
                key={block.id}
                table={block}
                doc={doc}
                pageNumber={pageNumber}
                totalPages={total}
                readonly={readonly}
                handlers={handlers}
              />
            ) : null,
          )}
        </div>
        {doc.footer ? (
          <div
            style={{
              position: 'absolute',
              left: margins.left,
              right: margins.right,
              bottom: margins.footer,
              maxHeight: Math.max(24, margins.bottom - margins.footer),
              overflow: 'hidden',
            }}
          >
            <HeaderFooterView
              blocks={doc.footer.blocks}
              doc={doc}
              pageNumber={pageNumber}
              totalPages={total}
            />
          </div>
        ) : null}
      </div>
    </div>
  );
}

function findParagraphElement(container: HTMLElement | null, paragraphId: string): HTMLElement | null {
  if (!container) return null;
  const elements = container.querySelectorAll<HTMLElement>('[data-paragraph-id]');
  for (const element of Array.from(elements)) {
    if (element.dataset.paragraphId === paragraphId) return element;
  }
  return null;
}
