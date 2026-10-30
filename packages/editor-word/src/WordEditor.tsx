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
  mapParagraphs,
  mergeFragmentRuns,
  paragraphImages,
  runLength,
  runsLength,
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
import { pageMetrics, paginateDocument, type PageSlice, type ParagraphBreaker, type TableLayout } from './pagination';
import { WordToolbar } from './WordToolbar';
// 功能：Word 标尺（段落缩进可视化拖拽）；时间：2026-10-30；作者：hou0。
import { WordRuler } from './WordRuler';
import { type RulerIndents, type RulerMetrics } from './word-ruler';
// 功能：接入段落对话框与格式刷所需的排版工具；时间：2026-10-29；作者：lq。
import { ParagraphDialog } from './ParagraphDialog';
import {
  applyRunStyleToRuns,
  mergeParagraphStyle,
  pickParagraphStyle,
} from './typography';
// 功能：接入样式库（预设样式、批量套用、样式集）；时间：2026-10-29；作者：lq。
import {
  addStyleSet,
  applyStyleToParagraph,
  BUILTIN_STYLES,
  clearParagraphFormatting,
  createStyleFromParagraph,
  detectStyleId,
  documentStyles as documentStyleList,
  effectiveStyles,
  findStyle,
  importStyleSet,
  readStyleSets,
  removeStyleSet,
  resolveStyleRun,
  type WordStyleSet,
} from './styles';
import { NewStyleDialog, type NewStyleBase } from './NewStyleDialog';
import { SaveStyleSetDialog } from './StyleSetDialog';
// 功能：复用 Word 原版平台缩放快捷键，比例与文档视图同步；时间：2026-10-24；作者：lq。
import { normalizeWordZoom, useWordZoomShortcuts } from './zoom-shortcuts';
// 功能：接入页眉页脚草稿设置；时间：2026-10-24；作者：lq。
import { HeaderFooterDialog } from './HeaderFooterDialog';
// 功能：接入模板字段和批量报告流程；时间：2026-10-24；作者：lq。
import { MailMergePanel } from './MailMergePanel';
import { MailMergeCanvas, MergePreviewBar } from './MailMergePreview';
import { useMailMergeController } from './mail-merge-controller';
// 功能：接入表格选区与网格编辑工具；时间：2026-10-19；作者：lq。
import { TableTools } from './TableTools';
import { editTable, findTable, formatTable, intersects, tableSelectionRect, type TableAction, type TableSelection } from './table-editing';
import { useWorkspaceStore, type WordSession } from '@mini-office/document-model';
import { clearPrintRoots, printPages, renderPrintPreview } from '@mini-office/platform-adapter';
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
  WordTabStop,
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
  // 功能：快捷键提示按平台显示 ⌘ / Ctrl；时间：2026-10-29；作者：lq。
  MOD_KEY,
} from '@mini-office/shared';
import {
  ChevronDown,
  ChevronUp,
  FileWarning,
  Pilcrow,
  Replace,
  Search,
  X,
} from 'lucide-react';

// 功能：提供常用办公纸张尺寸，页面布局修改即时重新分页；时间：2026-10-24；作者：lq。
const PAGE_SIZE_OPTIONS = ['A4', 'A3', 'A5', 'Letter', 'Legal', 'B5'] as const;

/** 功能：从行内样式里取正数像素值，取不到时回退；用于读取每页所在节的实际纸张；时间：2026-10-29；作者：lq。 */
function positiveOr(value: string | undefined, fallback: number): number {
  const parsed = Number.parseFloat(value ?? '');
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

/** 功能：把节内页面设置切换到指定方向（互换宽高），方向一致时原样返回；时间：2026-10-29；作者：lq。 */
function withOrientation(page: DocModel['page'], orientation: 'portrait' | 'landscape'): DocModel['page'] {
  if (page.orientation === orientation) return page;
  return { ...page, orientation, widthPx: page.heightPx, heightPx: page.widthPx };
}

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

/** 功能：在正文、表格及嵌套表格中查找图片；时间：2026-10-24；作者：lq。 */
function findImageInBlocks(blocks: DocBlock[], imageId: string): DocImage | undefined {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      const run = block.runs.find((item) => item.image && (item.image.id === imageId || item.image.relId === imageId));
      if (run?.image) return run.image;
    } else if (block.type === 'table') {
      const found = block.rows.flatMap((row) => row.cells).map((cell) => findImageInBlocks(cell.blocks, imageId)).find(Boolean);
      if (found) return found;
    }
  }
  return undefined;
}

/** 功能：统一修改图片尺寸与环绕属性并保留其他文档结构；时间：2026-10-24；作者：lq。 */
function updateImageInBlocks(blocks: DocBlock[], imageId: string, patch: Partial<DocImage>): DocBlock[] {
  return blocks.map((block) => {
    if (block.type === 'paragraph') return { ...block, runs: block.runs.map((run) => run.image && (run.image.id === imageId || run.image.relId === imageId) ? { ...run, image: { ...run.image, ...patch } } : run) };
    if (block.type === 'table') return { ...block, rows: block.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell, blocks: updateImageInBlocks(cell.blocks, imageId, patch) })) })) };
    return block;
  });
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
  // 功能：headEnd 表示该片段在模型段落中的结束偏移，用于跨页首片编辑时保留续页尾部；时间：2026-10-30；作者：lq。
  onInput: (paragraphId: string, element: HTMLElement, headEnd?: number) => void;
  onBlur: (paragraphId: string) => void;
  onKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>, paragraphId: string, element: HTMLDivElement) => void;
  onTableContextMenu: (event: React.MouseEvent, table: DocTable, rowIndex: number, cellId: string) => void;
  // 功能：共享选区及尺寸拖动入口；时间：2026-10-19；作者：lq。
  tableSelection?: TableSelection | null;
  onTableResize: (event: React.PointerEvent, table: DocTable, index: number, axis: 'row' | 'column') => void;
  onPaste?: (event: React.ClipboardEvent<HTMLDivElement>) => void;
  /** 功能：记录当前选中的图片对象，供边框手柄与右键菜单使用；时间：2026-10-24；作者：lq。 */
  onImageSelect?: (imageId: string) => void;
  /** 功能：处理图片边框手柄拖动，实时预览并在松开时保存尺寸；时间：2026-10-24；作者：lq。 */
  onImageResizeStart?: (event: React.PointerEvent, imageId: string, handle: string) => void;
  /** 功能：将图片环绕与对象操作入口放入图片右键菜单；时间：2026-10-24；作者：lq。 */
  onImageContextMenu?: (event: React.MouseEvent, imageId: string) => void;
  selectedImageId?: string;
}

/**
 * 功能：转义 CSS 选择器中的特殊字符，用于按书签名定位目录跳转目标；时间：2026-10-30；作者：hou0。
 */
function cssEscape(value: string): string {
  if (typeof window !== 'undefined' && window.CSS && typeof window.CSS.escape === 'function') {
    return window.CSS.escape(value);
  }
  return value.replace(/["\\]/g, '\\$&');
}

/** 功能：点击目录项（data-anchor）滚动到对应书签标题，编辑态与只读态共用；时间：2026-10-30；作者：hou0。 */
function scrollToAnchor(currentTarget: HTMLElement, link: HTMLElement): void {
  const name = link.getAttribute('data-anchor');
  if (!name) return;
  const editorRoot = currentTarget.closest('[data-word-editor]');
  editorRoot
    ?.querySelector<HTMLElement>(`[data-bookmark~="${cssEscape(name)}"]`)
    ?.scrollIntoView({ behavior: 'smooth', block: 'start' });
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
        data-bookmark={paragraph.bookmarks?.length ? paragraph.bookmarks.join(' ') : undefined}
        style={paragraphTextStyle(paragraph, doc.defaultFont, doc.defaultFontSizePt)}
        onClick={(event) => {
          // 功能：只读段落同样支持点击目录项跳转；时间：2026-10-30；作者：hou0。
          const link = (event.target as HTMLElement).closest<HTMLElement>('[data-anchor]');
          if (link) {
            event.preventDefault();
            scrollToAnchor(event.currentTarget as HTMLElement, link);
          }
        }}
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
  headEnd,
}: {
  paragraph: DocParagraph;
  doc: DocModel;
  pageNumber: number;
  totalPages: number;
  readonly: boolean;
  handlers: ParagraphHandlers;
  /** 功能：跨页首片在模型段落中的结束偏移，编辑回写时据此保留续页尾部；时间：2026-10-30；作者：lq。 */
  headEnd?: number;
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

  useEffect(() => {
    ref.current?.querySelectorAll<HTMLElement>('[data-image-frame]').forEach((frame) => {
      const selected = !!handlers.selectedImageId && frame.dataset.imageId === handlers.selectedImageId;
      frame.style.outline = selected ? '2px solid #3478f6' : '';
      frame.style.outlineOffset = selected ? '2px' : '';
      frame.dataset.selected = selected ? 'true' : 'false';
      frame.querySelectorAll<HTMLElement>('[data-image-handle]').forEach((handle) => {
        handle.style.opacity = selected ? '1' : '0';
        handle.style.pointerEvents = selected ? 'auto' : 'none';
      });
    });
  }, [handlers.selectedImageId, html]);

  return (
    <div style={paragraphWrapperStyle(paragraph)}>
      <div
        ref={ref}
        // 功能：继承整篇文档的编辑宿主，允许跨段和跨页选区；时间：2026-10-18；作者：lq。
        tabIndex={-1}
        suppressContentEditableWarning
        spellCheck={false}
        data-paragraph-id={paragraph.id}
        // 功能：跨页首片的模型边界，供文档根的兜底提交复用；时间：2026-10-30；作者：lq。
        data-head-end={headEnd}
        data-bookmark={paragraph.bookmarks?.length ? paragraph.bookmarks.join(' ') : undefined}
        role="textbox"
        aria-multiline="true"
        aria-label="文档段落"
        style={{ ...paragraphTextStyle(paragraph, doc.defaultFont, doc.defaultFontSizePt), outline: 'none' }}
        onInput={(event) => handlers.onInput(paragraph.id, event.currentTarget, headEnd)}
        onPointerDown={(event) => {
          const handle = (event.target as HTMLElement).closest<HTMLElement>('[data-image-handle]');
          const frame = (event.target as HTMLElement).closest<HTMLElement>('[data-image-frame]');
          if (handle && frame && handlers.onImageResizeStart) {
            event.preventDefault();
            event.stopPropagation();
            handlers.onImageResizeStart(event, frame.dataset.imageId ?? '', handle.dataset.imageHandle ?? 'se');
          }
        }}
        // 功能：点击图片进入图片上下文工具，避免把图片当作普通文本处理；时间：2026-10-24；作者：lq。
        onClick={(event) => {
          // 功能：点击目录项（data-anchor）滚动到对应书签标题；时间：2026-10-30；作者：hou0。
          const link = (event.target as HTMLElement).closest<HTMLElement>('[data-anchor]');
          if (link) {
            event.preventDefault();
            scrollToAnchor(event.currentTarget as HTMLElement, link);
            return;
          }
          const image = (event.target as HTMLElement).closest<HTMLImageElement>('img[data-image-id]');
          if (image) {
            event.preventDefault();
            handlers.onImageSelect?.(image.dataset.imageId ?? '');
          }
        }}
        onContextMenu={(event) => {
          const frame = (event.target as HTMLElement).closest<HTMLElement>('[data-image-frame]');
          if (frame) handlers.onImageContextMenu?.(event, frame.dataset.imageId ?? '');
        }}
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
  onOpenParagraphDialog,
  onShowCompatibility,
  onClose,
}: {
  doc: DocModel;
  session: WordSession;
  activeParagraph?: DocParagraph;
  pageCount: number;
  onPageChange: (page: DocModel['page']) => void;
  onParagraphChange: (patch: Partial<ParagraphStyle>) => void;
  onOpenParagraphDialog: () => void;
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

        {/* 功能：段落格式属于局部文本样式，属性面板只保留入口，具体设置在工具栏“段落”中完成；时间：2026-10-29；作者：lq。 */}
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">段落格式</h3>
          <p className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
            段落格式只作用于当前选区，因此在工具栏“段落”里设置：缩进、行距（含固定值）、段前段后间距与段前分页。
          </p>
          <button
            type="button"
            onClick={onOpenParagraphDialog}
            disabled={!activeParagraph}
            className="flex w-full items-center gap-2 rounded-md border border-[var(--mo-line)] px-2 py-2 text-left text-[11px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
          >
            <Pilcrow size={13} className="text-[var(--mo-text-muted)]" />
            <span className="flex-1">打开段落设置…</span>
            <span className="text-[var(--mo-text-faint)]">{MOD_KEY}⌥M</span>
          </button>
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

interface ImageMenuState {
  x: number;
  y: number;
  imageId: string;
}

export function WordEditor({ session }: { session: WordSession }) {
  const doc = session.data as DocModel;
  const readonly = session.readonly;
  const updateData = useWorkspaceStore((state) => state.updateData);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const updateView = useWorkspaceStore((state) => state.updateView);
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const exportBytes = useWorkspaceStore((state) => state.exportBytes);
  const openCompatibility = useWorkspaceStore((state) => state.openDialog);
  const notify = useWorkspaceStore((state) => state.notify);

  // 功能：限定文档缩放快捷键的工作区范围；时间：2026-10-24；作者：lq。
  const editorRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<HTMLDivElement>(null);
  const measureRef = useRef<HTMLDivElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const focusRequestRef = useRef<{ paragraphId: string; offset: number } | null>(null);
  const docRef = useRef(doc);
  docRef.current = doc;

  const [layoutRevision, setLayoutRevision] = useState(0);
  const [pages, setPages] = useState<PageSlice[]>([{ index: 0, blocks: doc.blocks }]);
  // 功能：前置打印预览状态，与 Excel 保持一致的分页列表体验；时间：2026-10-29；作者：lq。
  const printPreviewRef = useRef<HTMLDivElement>(null);
  const printPreviewCleanupRef = useRef<(() => void) | null>(null);
  const [printPreviewOpen, setPrintPreviewOpen] = useState(false);
  const [printPreviewCount, setPrintPreviewCount] = useState(0);
  const [previewScaleMode, setPreviewScaleMode] = useState<'fit' | 'actual'>('fit');
  const [zoom, setZoom] = useState(() => session.view.zoom ?? 1);
  const [currentPage, setCurrentPage] = useState(1);
  const [activeParagraphId, setActiveParagraphId] = useState<string>();
  // 功能：Word 标尺——段落缩进拖拽，默认显示，工具栏可开关；时间：2026-10-30；作者：hou0。
  const [showRuler, setShowRuler] = useState(true);
  const [rulerMetrics, setRulerMetrics] = useState<RulerMetrics | null>(null);
  const rulerHostRef = useRef<HTMLDivElement>(null);
  // 功能：记录当前图片选中状态，驱动边框手柄与右键菜单；时间：2026-10-24；作者：lq。
  const [selectedImageId, setSelectedImageId] = useState<string>();
  const [activeStyle, setActiveStyle] = useState<RunStyle>({
    fontFamily: doc.defaultFont,
    fontSizePt: doc.defaultFontSizePt,
  });
  const [historyVersion, setHistoryVersion] = useState(0);
  const [syncToken, setSyncToken] = useState(0);
  const savedSelectionRef = useRef<{ paragraphId: string; start: number; end: number } | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const [findOpen, setFindOpen] = useState(false);
  // 功能：格式刷与段落对话框状态，供工具栏与属性面板共用；时间：2026-10-29；作者：lq。
  const [formatPainter, setFormatPainter] = useState<{ run: RunStyle; paragraph: ParagraphStyle; sticky: boolean } | null>(null);
  const [paragraphDialogOpen, setParagraphDialogOpen] = useState(false);
  // 功能：样式库面板与样式集状态，样式集从本机读取一次；时间：2026-10-29；作者：lq。
  const [stylePickerOpen, setStylePickerOpen] = useState(false);
  const [showMoreStyles, setShowMoreStyles] = useState(false);
  const [newStyleOpen, setNewStyleOpen] = useState(false);
  const [styleSetDialogOpen, setStyleSetDialogOpen] = useState(false);
  const [styleSets, setStyleSets] = useState<WordStyleSet[]>(() => readStyleSets());
  const [findQuery, setFindQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [findActive, setFindActive] = useState(0);
  // 功能：保存独立于文本焦点的表格矩形选区；时间：2026-10-19；作者：lq。
  const explicitTableSelectionRef = useRef(false);
  const [tableSelection, setTableSelection] = useState<TableSelection | null>(null);
  const [tableMenu, setTableMenu] = useState<TableMenuState | null>(null);
  // 功能：保存图片右键菜单位置与目标对象；时间：2026-10-24；作者：lq。
  const [imageMenu, setImageMenu] = useState<ImageMenuState | null>(null);
  const [exportOpen, setExportOpen] = useState(false);
  // 功能：按双击区域打开设置，每次打开重建草稿；时间：2026-10-24；作者：lq。
  const [headerFooterRegion, setHeaderFooterRegion] = useState<'header' | 'footer' | null>(null);
  const [mailMergeOpen, setMailMergeOpen] = useState(false);
  const mailMerge = useMailMergeController({
    session, doc, renderBlocks: renderMergeBlocks,
    // 功能：合并 Word 直接在软件窗口显示新文档，不先弹保存框；时间：2026-10-30；作者：lq。
    onOpenDocument: async (mergedDoc) => {
      const id = await createDocument('docx');
      if (!id) return;
      updateData(id, mergedDoc, { dirty: true });
    },
  });
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
    setActiveStyle(resolveStyleRun({ run: styleAtOffset(paragraph, offset) }, current));
  }, [latestDoc, rememberSelection]);

  useEffect(() => {
    document.addEventListener('selectionchange', refreshSelection);
    return () => document.removeEventListener('selectionchange', refreshSelection);
  }, [refreshSelection]);

  const activeParagraph = activeParagraphId
    ? findParagraph(doc, activeParagraphId)
    : undefined;

  // 功能：样式库数据源——内置预设 + 文档内样式（文档定义优先覆盖同名预设）；时间：2026-10-29；作者：lq。
  const styleList = useMemo(() => effectiveStyles(doc), [doc.styles]);
  const presetStyles = useMemo(() => styleList.slice(0, BUILTIN_STYLES.length), [styleList]);
  const extraStyles = useMemo(() => documentStyleList(doc), [doc.styles]);
  const currentStyleId = activeParagraph ? detectStyleId(activeParagraph, styleList, doc) : undefined;
  const currentStyleName = findStyle(styleList, currentStyleId)?.name ?? '自定义格式';

  const metrics = useMemo(() => pageMetrics(doc.page), [doc.page]);
  const effectiveZoom = exporting ? 1 : zoom;

  // 功能：Word 标尺——缩进标记取值、页面内容盒测量与拖拽提交；时间：2026-10-30；作者：hou0。
  const rulerIndents: RulerIndents = {
    left: activeParagraph?.style?.indentLeftPx ?? 0,
    firstLine: activeParagraph?.style?.firstLineIndentPx ?? 0,
    right: activeParagraph?.style?.indentRightPx ?? 0,
  };

  const measureRuler = useCallback(() => {
    if (!showRuler) return;
    const pageEl = pagesRef.current?.querySelector<HTMLElement>('[data-word-page-wrapper]');
    const host = rulerHostRef.current;
    if (!pageEl || !host) return;
    const pageRect = pageEl.getBoundingClientRect();
    const hostRect = host.getBoundingClientRect();
    const margins = doc.page.margins;
    const contentW = Math.max(120, doc.page.widthPx - margins.left - margins.right);
    setRulerMetrics({
      contentLeft: pageRect.left + margins.left * effectiveZoom - hostRect.left,
      contentWidth: contentW * effectiveZoom,
      zoom: effectiveZoom,
    });
  }, [doc.page, effectiveZoom, showRuler]);

  useEffect(() => {
    measureRuler();
    const scroller = scrollRef.current;
    if (!scroller) return;
    scroller.addEventListener('scroll', measureRuler, { passive: true });
    let ro: ResizeObserver | undefined;
    if (typeof ResizeObserver !== 'undefined') {
      ro = new ResizeObserver(measureRuler);
      ro.observe(scroller);
    }
    return () => {
      scroller.removeEventListener('scroll', measureRuler);
      ro?.disconnect();
    };
  }, [measureRuler]);

  const onRulerChange = useCallback(
    (next: RulerIndents) => {
      if (readonly || !activeParagraphId) return;
      const updated = updateParagraph(latestDoc(), activeParagraphId, (paragraph) => ({
        ...paragraph,
        style: mergeParagraphStyle(paragraph.style, {
          indentLeftPx: next.left,
          firstLineIndentPx: next.firstLine,
          indentRightPx: next.right,
        }),
      }));
      commit(updated, { reflow: true, key: 'ruler-indent' });
    },
    [readonly, activeParagraphId, commit, latestDoc],
  );

  const onTabStopsChange = useCallback(
    (next: WordTabStop[]) => {
      if (readonly || !activeParagraphId) return;
      const updated = updateParagraph(latestDoc(), activeParagraphId, (paragraph) => ({
        ...paragraph,
        style: mergeParagraphStyle(paragraph.style, { tabStops: next }),
      }));
      commit(updated, { reflow: true, key: 'ruler-tabstop' });
    },
    [readonly, activeParagraphId, commit, latestDoc],
  );

  const blockMap = useMemo(() => new Map(doc.blocks.map((block) => [block.id, block])), [doc.blocks]);

  useLayoutEffect(() => {
    const container = measureRef.current;
    if (!container) return;
    const heights = new Map<string, number>();
    const tables = new Map<string, TableLayout>();
    // 功能：缓存段落渲染外壳，供跨页切分时克隆测量前缀高度；时间：2026-10-30；作者：lq。
    const wrapperByParagraph = new Map<string, HTMLElement>();
    container.querySelectorAll<HTMLElement>('[data-measure-id]').forEach((element) => {
      const id = element.getAttribute('data-measure-id');
      if (id) heights.set(id, element.getBoundingClientRect().height);
      if (!id) return;
      // 功能：读取表格每行实测高度，用于长表格跨页切分并重复表头；时间：2026-10-29；作者：lq。
      const grid = element.querySelector<HTMLElement>('[data-table-id]');
      const block = blockMap.get(id);
      if (!grid || !block || block.type !== 'table') {
        if (block && block.type === 'paragraph') {
          const wrapper = element.firstElementChild as HTMLElement | null;
          if (wrapper) wrapperByParagraph.set(id, wrapper);
        }
        return;
      }
      const rowHeights = getComputedStyle(grid).gridTemplateRows
        .split(' ')
        .map((value) => parseFloat(value))
        .filter((value) => Number.isFinite(value) && value > 0);
      if (rowHeights.length !== block.rows.length) return;
      let headerRows = 0;
      while (block.rows[headerRows]?.isHeader) headerRows += 1;
      tables.set(id, { rowHeights, headerRows });
    });

    // 功能：段落断页器——按真实 DOM 测量前缀高度做二分，找到本页能容纳的字符断点；段内分页符作为强制断点；时间：2026-10-30；作者：lq。
    const probeRoot = document.createElement('div');
    probeRoot.style.cssText = `position:absolute;left:-100000px;top:0;width:${metrics.contentWidth}px;visibility:hidden;pointer-events:none;`;
    document.body.appendChild(probeRoot);
    const measurePrefix = (paragraph: DocParagraph, offset: number, baseWrapper: HTMLElement): number => {
      const [headRuns] = splitRunsAtOffset(paragraph.runs, offset, {});
      const inner = (baseWrapper.firstElementChild as HTMLElement).cloneNode(false) as HTMLElement;
      inner.innerHTML = renderParagraphHtml({ ...paragraph, runs: headRuns }, {});
      const probeWrapper = baseWrapper.cloneNode(false) as HTMLElement;
      probeWrapper.style.paddingTop = `${paragraph.style.spaceBeforePx ?? 0}px`;
      probeWrapper.style.paddingBottom = `${paragraph.style.spaceAfterPx ?? 0}px`;
      probeWrapper.innerHTML = '';
      probeWrapper.appendChild(inner);
      const box = document.createElement('div');
      box.style.display = 'flow-root';
      box.appendChild(probeWrapper);
      probeRoot.appendChild(box);
      const h = box.getBoundingClientRect().height;
      probeRoot.removeChild(box);
      return h;
    };
    const toContinuation = (p: DocParagraph): DocParagraph => ({
      ...p,
      style: { ...p.style, firstLineIndentPx: 0, spaceBeforePx: 0, spaceAfterPx: 0 },
      list: undefined,
    });
    /** 功能：找出段落中第一个段内分页符所在的模型偏移；时间：2026-10-30；作者：lq。 */
    const forcedBreakOffset = (paragraph: DocParagraph): number => {
      let acc = 0;
      for (const run of paragraph.runs) {
        if (run.pageBreak) return acc;
        acc += runLength(run);
      }
      return -1;
    };
    const breaker: ParagraphBreaker = {
      needsBreak: (paragraph) => paragraph.runs.some((run) => run.pageBreak),
      split: (paragraph, maxHeight) => {
        const baseId = paragraph.id.replace(/(?:-cont)+$/, '') || paragraph.id;
        const wrapper = wrapperByParagraph.get(baseId);
        if (!wrapper) return null;
        const total = runsLength(paragraph.runs);
        if (total <= 0) return null;
        // 功能：强制断点之前的内容必须留在本页，二分上界不能越过它；时间：2026-10-30；作者：lq。
        const forced = forcedBreakOffset(paragraph);
        const limit = forced >= 0 ? forced : total;
        let lo = 0;
        let hi = limit;
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (measurePrefix(paragraph, mid, wrapper) <= maxHeight + 0.5) lo = mid;
          else hi = mid - 1;
        }
        // 功能：强制断点处放得下就按它断开；分页符本身归入续页，避免续页再次从同一断点起断；时间：2026-10-30；作者：lq。
        const forcedHit = lo >= limit && forced >= 0 && forced < total;
        const headEnd = forcedHit ? forced : lo;
        const tailStart = forcedHit ? forced + 1 : headEnd;
        if (headEnd <= 0 && tailStart <= 0) return null;
        if (tailStart >= total && headEnd <= 0) return null;
        const [headRuns] = splitRunsAtOffset(paragraph.runs, headEnd, {});
        const [, tailRuns] = splitRunsAtOffset(paragraph.runs, tailStart, {});
        const head = { ...paragraph, runs: headRuns };
        const tail = toContinuation({ ...paragraph, runs: tailRuns });
        return {
          head,
          tail,
          headEnd,
          // 功能：段内分页符断出的续片必须另起一页；高度切出的续片则接在本页首片之后；时间：2026-10-30；作者：lq。
          forceNewPage: forcedHit,
          headHeight: measurePrefix(head, headEnd, wrapper),
          tailHeight: measurePrefix(tail, total - tailStart, wrapper),
        };
      },
    };
    setPages(paginateDocument(docRef.current, heights, tables, breaker));
    document.body.removeChild(probeRoot);
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

  // 功能：把屏幕页面克隆成 1:1 打印页，前置预览与系统打印共用同一份页面，页数必然一致；时间：2026-10-29；作者：lq。
  const buildWordPrintPages = useCallback((): { pages: HTMLElement[]; sizes: Array<{ widthPx: number; heightPx: number }> } => {
    const source = pagesRef.current;
    if (!source) return { pages: [], sizes: [] };
    const wrappers = Array.from(source.querySelectorAll<HTMLElement>('[data-word-page-wrapper]'));
    const sizes: Array<{ widthPx: number; heightPx: number }> = [];
    const pages = wrappers.map((wrapper) => {
      const clone = wrapper.cloneNode(true) as HTMLElement;
      clone.removeAttribute('contenteditable');
      const paper = clone.querySelector<HTMLElement>('.mo-word-page');
      // 功能：每页按自己所在节的纸张取尺寸，多节混合方向时不会被强行拉成同一张纸；时间：2026-10-29；作者：lq。
      const width = positiveOr(paper?.style.width, doc.page.widthPx);
      const height = positiveOr(paper?.style.height, doc.page.heightPx);
      sizes.push({ widthPx: width, heightPx: height });
      clone.style.width = `${width}px`;
      clone.style.height = `${height}px`;
      if (paper) {
        // 功能：打印/预览时复位屏幕缩放，保证 1:1 落在纸张上；时间：2026-10-29；作者：lq。
        paper.style.transform = 'none';
        paper.style.width = `${width}px`;
        paper.style.height = `${height}px`;
        paper.style.boxShadow = 'none';
        paper.style.margin = '0';
      }
      clone.querySelectorAll<HTMLElement>('[contenteditable]').forEach((element) => element.removeAttribute('contenteditable'));
      return clone;
    });
    return { pages, sizes };
  }, [doc.page.heightPx, doc.page.widthPx]);

  const printWord = useCallback(async () => {
    const { pages, sizes } = buildWordPrintPages();
    if (pages.length === 0) return false;
    // 功能：清理上次打印残留的样式，避免本次打印被旧的 @media print 规则干扰；时间：2026-10-29；作者：lq。
    clearPrintRoots();
    return printPages({
      pages,
      pageSizes: sizes,
      widthPx: doc.page.widthPx,
      heightPx: doc.page.heightPx,
      title: session.displayName,
    });
  }, [buildWordPrintPages, doc.page.heightPx, doc.page.widthPx, session.displayName]);

  const openPrintPreview = useCallback(() => {
    printPreviewCleanupRef.current?.();
    setPrintPreviewOpen(true);
  }, []);

  const closePrintPreview = useCallback(() => {
    setPrintPreviewOpen(false);
    printPreviewCleanupRef.current?.();
    printPreviewCleanupRef.current = null;
  }, []);

  useEffect(() => {
    if (!printPreviewOpen || !printPreviewRef.current) return;
    const { pages, sizes } = buildWordPrintPages();
    // 功能：预览与打印走同一条构建路径，所见即所得；时间：2026-10-29；作者：lq。
    const cleanup = renderPrintPreview({
      host: printPreviewRef.current,
      pages,
      pageSizes: sizes,
      widthPx: doc.page.widthPx,
      heightPx: doc.page.heightPx,
      scale: previewScaleMode === 'actual' ? 1 : undefined,
    });
    setPrintPreviewCount(pages.length);
    printPreviewCleanupRef.current = cleanup;
    return () => {
      cleanup();
      printPreviewCleanupRef.current = null;
    };
  }, [printPreviewOpen, previewScaleMode, buildWordPrintPages, doc.page.widthPx, doc.page.heightPx, pages]);

  const confirmPrintPreview = useCallback(async () => {
    closePrintPreview();
    const ok = await printWord();
    if (!ok) notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
  }, [closePrintPreview, notify, printWord]);

  // 功能：统计文档实际用到的纸张方向，判断是否属于「多节混合方向」；时间：2026-10-29；作者：lq。
  const pageOrientations = useMemo(() => {
    const set = new Set<'portrait' | 'landscape'>();
    for (const slice of pages) set.add((slice.section?.page ?? doc.page).orientation);
    return set;
  }, [pages, doc.page]);
  const mixedOrientation = pageOrientations.size > 1;
  const documentOrientation: 'portrait' | 'landscape' | null =
    pageOrientations.size === 1 ? ([...pageOrientations][0] ?? 'portrait') : null;

  // 功能：把方向应用到整篇文档（含所有分节），对应 Word「页面设置 → 应用于：整篇文档」；时间：2026-10-29；作者：lq。
  const applyDocumentOrientation = useCallback(
    (orientation: 'portrait' | 'landscape') => {
      if (readonly) return;
      const current = latestDoc();
      const blocks = current.blocks.map((block) =>
        block.type === 'paragraph' && block.sectionEnd
          ? { ...block, sectionEnd: { ...block.sectionEnd, page: withOrientation(block.sectionEnd.page, orientation) } }
          : block,
      );
      const changedSections = blocks.some((block, index) => block !== current.blocks[index]);
      const page = withOrientation(current.page, orientation);
      if (!changedSections && page === current.page) return;
      commit({ ...current, page, blocks }, { reflow: true, key: 'page-orientation' });
    },
    [commit, latestDoc, readonly],
  );

  useEffect(() => {
    const handleFind = () => setFindOpen((open) => !open);
    const handleExport = () => {
      setExportName(`${stripExtension(session.displayName)}.pdf`);
      setExportRange('');
      setExportOpen(true);
    };
    const handlePrint = () => openPrintPreview();
    // 功能：命令面板 / 菜单「样式库」直接打开浮层；时间：2026-10-29；作者：lq。
    const handleStyles = () => setStylePickerOpen(true);
    window.addEventListener('mo:find', handleFind);
    window.addEventListener('mo:export-pdf', handleExport);
    window.addEventListener('mo:print', handlePrint);
    window.addEventListener('mo:styles', handleStyles);
    return () => {
      window.removeEventListener('mo:find', handleFind);
      window.removeEventListener('mo:export-pdf', handleExport);
      window.removeEventListener('mo:print', handlePrint);
      window.removeEventListener('mo:styles', handleStyles);
    };
  }, [openPrintPreview, session.displayName]);

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
      const next = normalizeWordZoom(value);
      setZoom(next);
      updateView(session.id, { zoom: next });
    },
    [session.id, updateView],
  );

  // 功能：只读文档同样允许缩放，导出时暂停快捷键以保护固定的导出版式；时间：2026-10-24；作者：lq。
  useWordZoomShortcuts({ editorRef, scrollRef, zoom, onZoomChange: changeZoom, enabled: !exporting });

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
    (paragraphId: string, element: HTMLElement, headEnd?: number) => {
      const current = latestDoc();
      const paragraph = findParagraph(current, paragraphId);
      if (!paragraph) return;
      const fallback = baseRunStyle(paragraph, current);
      const domRuns = serializeDomRuns(element, fallback, paragraphImages(paragraph));
      // 功能：跨页首片只含段落前缀，编辑后按边界回写，保留被切到续页的尾部；时间：2026-10-30；作者：lq。
      const runs = headEnd === undefined
        ? domRuns
        : mergeFragmentRuns(paragraph.runs, headEnd, domRuns, fallback);
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
      // 功能：回车分段时继承样式关联，保证新段落仍属于同一标题层级；时间：2026-10-29；作者：lq。
      const nextParagraph = makeParagraph(after, {
        style: { ...paragraph.style },
        styleId: paragraph.styleId,
        list: paragraph.list ? { ...paragraph.list } : undefined,
      });
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

  const insertPageBreak = useCallback(() => {
    if (readonly) return;
    const current = latestDoc();
    const index = topLevelBlockIndex(current, activeParagraphId);
    const blocks = [...current.blocks];
    blocks.splice(Math.max(0, index) + 1, 0, { type: 'pageBreak', id: uid('break') });
    commit({ ...current, blocks }, { reflow: true });
  }, [activeParagraphId, commit, latestDoc, readonly]);

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
      // 功能：Ctrl/Cmd+Enter 插入分页符、Shift+Enter 插入段内换行，与 Word/WPS 键位保持一致；时间：2026-10-29；作者：lq。
      if (event.key === 'Enter' && mod) {
        event.preventDefault();
        event.stopPropagation();
        insertPageBreak();
        return;
      }
      if (event.key === 'Enter' && event.shiftKey) {
        event.preventDefault();
        event.stopPropagation();
        if (!execCommand('insertLineBreak')) execCommand('insertHTML', '<br>');
        triggerSelectionInput();
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
    [applyInlineCommand, execCommand, insertPageBreak, mergePrevious, readonly, redo, splitParagraph, triggerSelectionInput, undo],
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
    // 功能：Esc 退出格式刷，避免误把后续选择都套上同一格式；时间：2026-10-29；作者：lq。
    if (event.key === 'Escape' && formatPainter) {
      event.preventDefault();
      setFormatPainter(null);
      notify('info', '已退出格式刷');
      return;
    }
    // 功能：⌘/Ctrl+⌥+M 打开段落对话框，与 Word for Mac 一致；时间：2026-10-29；作者：lq。
    if ((event.metaKey || event.ctrlKey) && event.altKey && event.key.toLowerCase() === 'm') {
      event.preventDefault();
      event.stopPropagation();
      setParagraphDialogOpen(true);
      return;
    }
    // 功能：⌘/Ctrl+⌥+1~4 套用标题 1~4，⌘/Ctrl+⇧+N 回到正文，与 Word 键位一致；时间：2026-10-29；作者：lq。
    // 说明：这两个组合键在 onKeyDownCapture 中处理（见 handleStyleShortcuts），避免与正文键入逻辑互相干扰。
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
  }, [captureSelection, commit, formatPainter, handleKeyDown, latestDoc, notify, readonly, replaceCrossSelection]);

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

  /** 功能：套用样式后按原偏移恢复选区，避免 DOM 重绘导致跨段选区丢失；时间：2026-10-29；作者：lq。 */
  const restoreSliceSelection = useCallback(
    (slices: Array<{ paragraphId: string; start: number; end: number }>) => {
      if (slices.length === 0) return;
      requestAnimationFrame(() => {
        const container = pagesRef.current;
        if (!container) return;
        const first = findParagraphElement(container, slices[0].paragraphId);
        const last = findParagraphElement(container, slices[slices.length - 1].paragraphId);
        if (!first || !last) return;
        first.focus({ preventScroll: true });
        if (slices.length === 1) {
          selectRangeByOffset(first, slices[0].start, slices[0].end);
          return;
        }
        const range = document.createRange();
        range.setStart(first, 0);
        range.setEnd(last, last.childNodes.length);
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        savedSelectionRef.current = {
          paragraphId: slices[0].paragraphId,
          start: slices[0].start,
          end: slices[0].start,
        };
      });
    },
    [],
  );

  /**
   * 功能：把选中段落批量套用为指定样式；时间：2026-10-29；作者：lq。
   *
   * 与 WPS 的差异：WPS 每次套用都会丢掉选区，无法连续比较多个样式；这里在提交后按原偏移
   * 恢复选区，面板也保持打开，可以连点几个样式直到满意。整次套用只写一条撤销记录。
   */
  const applyNamedStyle = useCallback(
    (styleId: string) => {
      if (readonly) return;
      const style = findStyle(styleList, styleId);
      if (!style) return;
      restoreSelection();
      const ids = selectedParagraphIds();
      if (ids.length === 0) {
        notify('info', '请先选择段落', '把光标放到要套用样式的段落中，或拖选多个段落。');
        return;
      }
      const container = pagesRef.current;
      const selection = window.getSelection();
      const slices =
        container && selection?.rangeCount
          ? selectedParagraphSlices(container, selection.getRangeAt(0))
          : [];
      const current = latestDoc();
      let next = current;
      let counter = 0;
      for (const id of ids) {
        counter += 1;
        const marker = style.list ? (style.list.ordered ? `${counter}.` : '•') : undefined;
        next = updateParagraph(next, id, (paragraph) =>
          applyStyleToParagraph(paragraph, style, { doc: current, marker }),
        );
      }
      commit(next, { reflow: true, key: 'apply-style' });
      setActiveStyle(resolveStyleRun(style, current));
      setSyncToken((value) => value + 1);
      restoreSliceSelection(slices);
    },
    [
      commit,
      latestDoc,
      notify,
      readonly,
      restoreSelection,
      restoreSliceSelection,
      selectedParagraphIds,
      styleList,
    ],
  );

  /** 功能：样式快捷键——⌘/Ctrl+⌥+1~4 套用标题、⌘/Ctrl+⇧+N 回到正文，键位与 Word 一致；时间：2026-10-29；作者：lq。 */
  const handleStyleShortcuts = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if (readonly) return;
      const mod = event.metaKey || event.ctrlKey;
      if (!mod) return;
      if (event.altKey && /^[1-4]$/.test(event.key)) {
        event.preventDefault();
        event.stopPropagation();
        applyNamedStyle(`Heading${event.key}`);
        return;
      }
      if (event.shiftKey && event.key.toLowerCase() === 'n') {
        event.preventDefault();
        event.stopPropagation();
        applyNamedStyle('Normal');
      }
    },
    [applyNamedStyle, readonly],
  );

  /** 功能：清除选中段落的全部排版格式，回到文档默认字体；时间：2026-10-29；作者：lq。 */
  const clearFormatting = useCallback(() => {
    if (readonly) return;
    restoreSelection();
    const ids = selectedParagraphIds();
    if (ids.length === 0) {
      notify('info', '没有选中的段落', '请将光标放在要清除格式的段落中。');
      return;
    }
    const container = pagesRef.current;
    const selection = window.getSelection();
    const slices =
      container && selection?.rangeCount ? selectedParagraphSlices(container, selection.getRangeAt(0)) : [];
    const current = latestDoc();
    let next = current;
    for (const id of ids) {
      next = updateParagraph(next, id, (paragraph) => clearParagraphFormatting(paragraph, current));
    }
    commit(next, { reflow: true, key: 'clear-format' });
    setActiveStyle({ fontFamily: current.defaultFont, fontSizePt: current.defaultFontSizePt });
    setSyncToken((value) => value + 1);
    restoreSliceSelection(slices);
  }, [commit, latestDoc, notify, readonly, restoreSelection, restoreSliceSelection, selectedParagraphIds]);

  /** 功能：以当前段落格式创建自定义样式，并立即套用到选中段落；时间：2026-10-29；作者：lq。 */
  const createStyle = useCallback(
    (name: string, base: NewStyleBase) => {
      if (readonly) return;
      restoreSelection();
      const current = latestDoc();
      const ids = selectedParagraphIds();
      const target = (ids[0] ? findParagraph(current, ids[0]) : undefined) ?? activeParagraph;
      if (!target) {
        notify('info', '请先选择段落', '把光标放到要作为样式蓝本的段落中。');
        return;
      }
      const source: DocParagraph =
        base === 'current'
          ? target
          : { ...target, style: {}, runs: [{ text: '', style: {} }], list: undefined };
      const definition = createStyleFromParagraph(name, source, current);
      const withStyle: DocModel = { ...current, styles: [...(current.styles ?? []), definition] };
      const targets = ids.length > 0 ? ids : [target.id];
      const next = targets.reduce(
        (acc, id) =>
          updateParagraph(acc, id, (paragraph) =>
            applyStyleToParagraph(paragraph, definition, { doc: current }),
          ),
        withStyle,
      );
      commit(next, { reflow: true });
      setSyncToken((value) => value + 1);
      notify('success', `已创建样式「${definition.name}」`, '样式已保存在本文档，可在样式库的「本文档样式」中重复使用。');
    },
    [activeParagraph, commit, latestDoc, notify, readonly, restoreSelection, selectedParagraphIds],
  );

  /** 功能：删除本文档样式；引用该样式的段落保留现有格式，仅解除关联；时间：2026-10-29；作者：lq。 */
  const deleteStyle = useCallback(
    (styleId: string) => {
      if (readonly) return;
      const current = latestDoc();
      const styles = (current.styles ?? []).filter((style) => style.id !== styleId);
      const blocks = mapParagraphs(current.blocks, (paragraph) =>
        paragraph.styleId === styleId ? { ...paragraph, styleId: undefined } : paragraph,
      );
      commit({ ...current, styles, blocks }, { reflow: true });
      notify('info', '已删除样式', '使用该样式的段落会保留当前格式，只是不再关联样式。');
    },
    [commit, latestDoc, notify, readonly],
  );

  /** 功能：把本文档的自定义样式保存为可复用的样式集；时间：2026-10-29；作者：lq。 */
  const openSaveStyleSet = useCallback(() => {
    if (extraStyles.length === 0) {
      notify('info', '还没有自定义样式', '先用「新建样式」创建样式，或从样式集导入，再保存为样式集。');
      return;
    }
    setStylePickerOpen(false);
    setStyleSetDialogOpen(true);
  }, [extraStyles.length, notify]);

  const saveStyleSet = useCallback(
    (name: string) => {
      setStyleSets(addStyleSet(name, extraStyles));
      notify('success', '样式集已保存', `「${name}」包含 ${extraStyles.length} 个样式，可在任意文档的样式库中套用。`);
    },
    [extraStyles, notify],
  );

  /** 功能：把样式集导入当前文档（不覆盖已有样式，冲突时重新分配 ID）；时间：2026-10-29；作者：lq。 */
  const applyStyleSetToDoc = useCallback(
    (set: WordStyleSet) => {
      if (readonly) return;
      const current = latestDoc();
      commit({ ...current, styles: importStyleSet(set, current.styles) });
      notify('success', `已导入「${set.name}」`, `${set.styles.length} 个样式已加入本文档样式库。`);
    },
    [commit, latestDoc, notify, readonly],
  );

  const deleteStyleSet = useCallback((id: string) => {
    setStyleSets(removeStyleSet(id));
  }, []);

  /**
   * 功能：格式刷——先复制光标处的字体与段落格式，再套用到目标选区；时间：2026-10-29；作者：lq。
   * 单击为单次套用，双击（sticky）可连续套用，再次点击或按 Esc 退出，行为与 Word / WPS 一致。
   */
  const startFormatPainter = useCallback(
    (sticky: boolean) => {
      if (readonly) return;
      if (formatPainter && !sticky) {
        setFormatPainter(null);
        notify('info', '已退出格式刷');
        return;
      }
      captureSelection();
      const current = latestDoc();
      const saved = savedSelectionRef.current;
      const id = saved?.paragraphId ?? activeParagraphId;
      const paragraph = id ? findParagraph(current, id) : undefined;
      const run = paragraph
        ? {
            ...baseRunStyle(paragraph, current),
            ...styleAtOffset(paragraph, saved && saved.paragraphId === paragraph.id ? saved.start : 0),
          }
        : { ...activeStyle };
      setFormatPainter({ run, paragraph: pickParagraphStyle(paragraph?.style), sticky });
      notify(
        'info',
        sticky ? '格式刷已锁定' : '已复制格式',
        sticky ? '连续拖选可多次套用，按 Esc 或再次点击格式刷退出。' : '在目标文字上拖选即可套用格式。',
      );
    },
    [activeParagraphId, activeStyle, captureSelection, formatPainter, latestDoc, notify, readonly],
  );

  const applyFormatPainter = useCallback(() => {
    const painter = formatPainter;
    const root = pagesRef.current;
    if (!painter || readonly || !root) return false;
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return false;
    const range = selection.getRangeAt(0);
    if (range.collapsed || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return false;
    const slices = selectedParagraphSlices(root, range);
    if (slices.length === 0) return false;
    let next = latestDoc();
    for (const slice of slices) {
      next = updateParagraph(next, slice.paragraphId, (paragraph) => ({
        ...paragraph,
        style: mergeParagraphStyle(paragraph.style, painter.paragraph),
        runs: applyRunStyleToRuns(paragraph.runs, slice.start, slice.end, painter.run, baseRunStyle(paragraph, next)),
      }));
    }
    commit(next, { reflow: true, key: 'format-painter' });
    setActiveStyle({ ...painter.run });
    // 功能：模型改动后强制刷新受控段落 DOM，再恢复选区以便连续套用；时间：2026-10-29；作者：lq。
    setSyncToken((value) => value + 1);
    const first = slices[0];
    requestAnimationFrame(() => {
      const element = findParagraphElement(pagesRef.current, first.paragraphId);
      if (!element) return;
      element.focus({ preventScroll: true });
      selectRangeByOffset(element, first.start, first.end);
    });
    if (!painter.sticky) setFormatPainter(null);
    return true;
  }, [commit, formatPainter, latestDoc, readonly]);

  const updatePage = useCallback(
    (page: DocModel['page']) => {
      const current = latestDoc();
      commit({ ...current, page }, { reflow: true });
    },
    [commit, latestDoc],
  );
  const updateImage = useCallback((imageId: string, patch: Partial<DocImage>) => {
    if (readonly || !imageId) return;
    const current = latestDoc();
    commit({ ...current, blocks: updateImageInBlocks(current.blocks, imageId, patch) }, { reflow: true, key: `image:${imageId}` });
  }, [commit, latestDoc, readonly]);

  // 功能：拖动图片边框手柄调整尺寸，角点默认保持原始比例；时间：2026-10-24；作者：lq。
  const handleImageResizeStart = useCallback((event: React.PointerEvent, imageId: string, handle: string) => {
    if (readonly || event.button !== 0 || !imageId) return;
    const image = findImageInBlocks(latestDoc().blocks, imageId);
    const frame = event.currentTarget.closest<HTMLElement>('[data-image-frame]')
      ?? (event.target as HTMLElement).closest<HTMLElement>('[data-image-frame]');
    if (!image || !frame) return;
    const startX = event.clientX;
    const startY = event.clientY;
    const startWidth = Math.max(8, image.widthPx);
    const startHeight = Math.max(8, image.heightPx);
    const ratio = startWidth / startHeight;
    const horizontal = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
    const vertical = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
    const isCorner = handle.length === 2;
    const originalFrameStyle = frame.getAttribute('style');
    const preview = frame.querySelector<HTMLImageElement>('img[data-image-id]');
    const originalImageStyle = preview?.getAttribute('style') ?? null;
    let nextWidth = startWidth;
    let nextHeight = startHeight;
    const applyPreview = (width: number, height: number) => {
      frame.style.width = `${width}px`;
      if (preview) {
        preview.style.width = `${width}px`;
        preview.style.height = `${height}px`;
      }
    };
    const move = (pointer: PointerEvent) => {
      const dx = (pointer.clientX - startX) / effectiveZoom;
      const dy = (pointer.clientY - startY) / effectiveZoom;
      const widthFromX = Math.max(8, startWidth + horizontal * dx);
      const heightFromY = Math.max(8, startHeight + vertical * dy);
      if (isCorner && !pointer.shiftKey) {
        const widthScale = widthFromX / startWidth;
        const heightScale = heightFromY / startHeight;
        const scale = Math.max(widthScale, heightScale, 8 / Math.min(startWidth, startHeight));
        nextWidth = Math.max(8, startWidth * scale);
        nextHeight = Math.max(8, nextWidth / ratio);
      } else {
        nextWidth = horizontal === 0 ? startWidth : widthFromX;
        nextHeight = vertical === 0 ? startHeight : heightFromY;
      }
      applyPreview(nextWidth, nextHeight);
    };
    const finish = (pointer: PointerEvent) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', finish);
      window.removeEventListener('pointercancel', finish);
      if (originalFrameStyle !== null) frame.setAttribute('style', originalFrameStyle);
      if (preview && originalImageStyle !== null) preview.setAttribute('style', originalImageStyle);
      if (pointer.type === 'pointercancel') return;
      if (Math.abs(nextWidth - startWidth) < 0.5 && Math.abs(nextHeight - startHeight) < 0.5) return;
      updateImage(imageId, { widthPx: nextWidth, heightPx: nextHeight });
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', finish);
    window.addEventListener('pointercancel', finish);
  }, [effectiveZoom, latestDoc, readonly, updateImage]);

  // 功能：打开图片专属右键菜单，把文字环绕等低频操作移出常驻工具栏；时间：2026-10-24；作者：lq。
  const handleImageContextMenu = useCallback((event: React.MouseEvent, imageId: string) => {
    if (readonly || !imageId) return;
    event.preventDefault();
    event.stopPropagation();
    setSelectedImageId(imageId);
    setImageMenu({ x: event.clientX, y: event.clientY, imageId });
  }, [readonly]);

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
      onImageSelect: setSelectedImageId,
      onImageResizeStart: handleImageResizeStart,
      onImageContextMenu: handleImageContextMenu,
      selectedImageId,
    }),
    [handleImageContextMenu, handleImageResizeStart, handleParagraphBlur, handleParagraphInput, handleKeyDown, handlePaste, handleTableContextMenu, handleTableResize, selectedImageId, tableSelection, syncToken],
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
  // 功能：提供图片右键环绕菜单，集中低频对象布局操作；时间：2026-10-24；作者：lq。
  const imageMenuItems: MenuItem[] = imageMenu ? [
    { type: 'label', label: '图片文字环绕' },
    { id: 'inline', label: '嵌入型', checked: (findImageInBlocks(doc.blocks, imageMenu.imageId)?.wrap ?? 'inline') === 'inline', onSelect: () => updateImage(imageMenu.imageId, { wrap: 'inline' }) },
    { id: 'square-left', label: '左侧环绕', checked: findImageInBlocks(doc.blocks, imageMenu.imageId)?.wrap === 'square-left', onSelect: () => updateImage(imageMenu.imageId, { wrap: 'square-left' }) },
    { id: 'square-right', label: '右侧环绕', checked: findImageInBlocks(doc.blocks, imageMenu.imageId)?.wrap === 'square-right', onSelect: () => updateImage(imageMenu.imageId, { wrap: 'square-right' }) },
    { id: 'top-bottom', label: '上下型', checked: findImageInBlocks(doc.blocks, imageMenu.imageId)?.wrap === 'top-bottom', onSelect: () => updateImage(imageMenu.imageId, { wrap: 'top-bottom' }) },
    { type: 'separator' },
    { id: 'image-done', label: '完成图片选择', onSelect: () => setSelectedImageId(undefined) },
  ] : [];

  return (
    <div ref={editorRef} className="flex h-full min-h-0 w-full">
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
          showRuler={showRuler}
          onToggleRuler={() => setShowRuler((visible) => !visible)}
          onExport={() => {
            setExportName(`${stripExtension(session.displayName)}.pdf`);
            setExportRange('');
            setExportOpen(true);
          }}
          onPrint={() => openPrintPreview()}
          onShowCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
          onAddBullets={() => applyList(false)}
          onAddNumbering={() => applyList(true)}
          onFormatPainter={startFormatPainter}
          formatPainterActive={!!formatPainter}
          onOpenParagraphDialog={() => setParagraphDialogOpen(true)}
          stylePicker={{
            presets: presetStyles,
            documentStyles: extraStyles,
            currentStyleId,
            currentStyleName,
            showMore: showMoreStyles,
            defaultFont: doc.defaultFont,
            defaultFontSizePt: doc.defaultFontSizePt,
            readonly,
            styleSets,
            open: stylePickerOpen,
            onOpenChange: setStylePickerOpen,
            onApply: applyNamedStyle,
            onClearFormatting: clearFormatting,
            onNewStyle: () => {
              // 功能：新建/保存对话框需要独占焦点，先收起样式库浮层；时间：2026-10-29；作者：lq。
              setStylePickerOpen(false);
              setNewStyleOpen(true);
            },
            onToggleMore: () => setShowMoreStyles((value) => !value),
            onDeleteStyle: deleteStyle,
            onSaveStyleSet: openSaveStyleSet,
            onApplyStyleSet: applyStyleSetToDoc,
            onDeleteStyleSet: deleteStyleSet,
            onBeforeInteraction: captureSelection,
          }}
          onBeforeInteraction={captureSelection}
        />
        ) : null}
        {mailMerge.previewOn ? <MergePreviewBar controller={mailMerge} /> : null}
        {/* 功能：选中表格自动显示专属工具；时间：2026-10-19；作者：lq。 */}
        {!readonly && activeTable && tableSelection ? <TableTools table={activeTable} selection={tableSelection} onAction={runTableAction} onFormat={(format) => mutateTable(activeTable.id, (table) => formatTable(table, tableSelection, format, metrics.contentWidth))} /> : null}
        {/* 功能：Word 标尺——对齐页面内容盒，拖拽调整当前段落缩进；时间：2026-10-30；作者：hou0。 */}
        {!readonly && !mailMerge.previewOn && showRuler ? (
          <div ref={rulerHostRef} className="relative">
            {rulerMetrics ? <WordRuler metrics={rulerMetrics} indents={rulerIndents} onChange={onRulerChange} tabStops={activeParagraph?.style?.tabStops ?? []} onTabStopsChange={onTabStopsChange} /> : null}
          </div>
        ) : null}
        {mailMerge.previewOn ? (
          <MailMergeCanvas controller={mailMerge} doc={doc} />
        ) : (
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
            // 功能：样式快捷键在捕获阶段处理，优先于正文键入与浏览器默认行为；时间：2026-10-29；作者：lq。
            onKeyDownCapture={handleStyleShortcuts}
            // 功能：格式刷激活时，拖选结束即套用复制的格式；时间：2026-10-29；作者：lq。
            onPointerUp={() => {
              if (formatPainter) applyFormatPainter();
            }}
            onInput={(event) => {
              // 功能：共享编辑区的原生 input 事件按当前段落写回模型；时间：2026-10-18；作者：lq。
              if (event.target !== event.currentTarget) return;
              const node = window.getSelection()?.anchorNode;
              const paragraph = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('[data-paragraph-id]');
              if (paragraph && event.currentTarget.contains(paragraph)) {
                // 功能：兜底提交同样按片段边界回写，避免跨页首片编辑截断模型尾部；时间：2026-10-30；作者：lq。
                const headEnd = paragraph.dataset.headEnd ? Number(paragraph.dataset.headEnd) : undefined;
                handleParagraphInput(paragraph.dataset.paragraphId!, paragraph, headEnd);
              }
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
        )}

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
      <MailMergePanel open={mailMergeOpen && !readonly} session={session} doc={doc} controller={mailMerge} onClose={() => setMailMergeOpen(false)} onInsertField={insertMergeField} renderBlocks={renderMergeBlocks} />
      {settings.showProperties && !mailMergeOpen && !mailMerge.previewOn ? (
        <PropertiesPanel
          doc={doc}
          session={session}
          activeParagraph={activeParagraph}
          pageCount={pages.length}
          onPageChange={updatePage}
          onParagraphChange={applyParagraphFormat}
          onOpenParagraphDialog={() => setParagraphDialogOpen(true)}
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
      {imageMenu ? (
        <ContextMenu
          x={imageMenu.x}
          y={imageMenu.y}
          items={imageMenuItems}
          onClose={() => setImageMenu(null)}
        />
      ) : null}

      {/* 功能：应用设置作为单次可撤销编辑，取消时不修改正文；时间：2026-10-24；作者：lq。 */}
      {headerFooterRegion && !readonly ? <HeaderFooterDialog doc={doc} pageCount={pages.length} initialRegion={headerFooterRegion} onClose={() => setHeaderFooterRegion(null)} onApply={(patch) => { commit({ ...latestDoc(), ...patch }, { reflow: true }); setHeaderFooterRegion(null); }} /> : null}
      {/* 功能：段落对话框统一提交缩进、间距与行距，只产生一条撤销记录；时间：2026-10-29；作者：lq。 */}
      {!readonly ? (
        <ParagraphDialog
          open={paragraphDialogOpen}
          style={activeParagraph?.style}
          fontSizePt={activeStyle.fontSizePt ?? doc.defaultFontSizePt}
          selectionHint={
            activeParagraph
              ? `将应用于当前选中的段落；行距与缩进按 ${activeStyle.fontSizePt ?? doc.defaultFontSizePt} 磅字号换算字符宽度。`
              : '请先把光标放到要调整的段落中。'
          }
          onClose={() => setParagraphDialogOpen(false)}
          onApply={(patch) => applyParagraphFormat(patch)}
        />
      ) : null}
      {/* 功能：新建样式与保存样式集，都以当前段落格式为蓝本；时间：2026-10-29；作者：lq。 */}
      {!readonly ? (
        <NewStyleDialog
          open={newStyleOpen}
          paragraph={activeParagraph}
          doc={doc}
          existingNames={styleList.map((style) => style.name)}
          defaultName={`样式 ${(doc.styles?.length ?? 0) + 1}`}
          onClose={() => setNewStyleOpen(false)}
          onCreate={createStyle}
        />
      ) : null}
      {!readonly ? (
        <SaveStyleSetDialog
          open={styleSetDialogOpen}
          defaultName={doc.meta.title ? `${doc.meta.title} 样式集` : '我的样式集'}
          styleCount={extraStyles.length}
          existingNames={styleSets.map((set) => set.name)}
          onClose={() => setStyleSetDialogOpen(false)}
          onSave={saveStyleSet}
        />
      ) : null}
      {/* 功能：打印前先给出与应用内分页一致的前置预览，避免系统预览缩略图看不清版式；时间：2026-10-29；作者：lq。 */}
      <Dialog
        open={printPreviewOpen}
        onClose={closePrintPreview}
        title="打印预览"
        description={`预览全部分页（共 ${printPreviewCount} 页），确认后调用系统打印。`}
        width="xl"
        footer={
          <>
            <Button onClick={closePrintPreview}>取消</Button>
            <Button variant="primary" onClick={() => void confirmPrintPreview()}>
              打印
            </Button>
          </>
        }
      >
        <div className="grid grid-cols-[180px_1fr] gap-4 text-[12px]">
          <div className="space-y-3 rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface-2)] p-3">
            <Field label="纸张">
              {mixedOrientation
                ? '多节混合'
                : `${doc.page.sizeName} · ${documentOrientation === 'landscape' ? '横向' : '纵向'}`}
            </Field>
            <Field label="尺寸">
              {mixedOrientation
                ? '按节保留'
                : `${Math.round(pxToMm(doc.page.widthPx))} × ${Math.round(pxToMm(doc.page.heightPx))} mm`}
            </Field>
            <Field label="总页数">{printPreviewCount} 页</Field>
            {/* 功能：方向切换作用于整篇文档（含所有分节），与 Word「页面设置 → 应用于整篇文档」一致；时间：2026-10-29；作者：lq。 */}
            <div className="space-y-1.5">
              <div className="text-[11px] text-[var(--mo-text-muted)]">方向</div>
              {mixedOrientation ? (
                <div className="space-y-1.5">
                  <div className="rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 py-1 text-[11px] leading-4 text-[var(--mo-text-muted)]">
                    本文档含多个节，方向按节保留
                  </div>
                  <div className="flex gap-1.5">
                    <Button size="sm" disabled={readonly} onClick={() => applyDocumentOrientation('portrait')}>
                      全部纵向
                    </Button>
                    <Button size="sm" disabled={readonly} onClick={() => applyDocumentOrientation('landscape')}>
                      全部横向
                    </Button>
                  </div>
                </div>
              ) : (
                <SegmentedControl
                  value={documentOrientation ?? 'portrait'}
                  onChange={applyDocumentOrientation}
                  options={[
                    { value: 'portrait', label: '纵向' },
                    { value: 'landscape', label: '横向' },
                  ]}
                  className="w-full"
                />
              )}
            </div>
            <div className="space-y-1.5">
              <div className="text-[11px] text-[var(--mo-text-muted)]">预览比例</div>
              <SegmentedControl
                value={previewScaleMode}
                onChange={setPreviewScaleMode}
                options={[
                  { value: 'fit', label: '适应页宽' },
                  { value: 'actual', label: '实际大小' },
                ]}
              />
            </div>
            <p className="leading-5 text-[var(--mo-text-muted)]">
              {mixedOrientation
                ? '每页按所属节的纸张输出，页码标签会标出该页方向。'
                : '分页位置与“导出 PDF”、系统打印完全一致。纸张与页边距可在右侧属性面板中调整。'}
            </p>
          </div>
          <div
            ref={printPreviewRef}
            className="relative min-h-[300px] max-h-[460px] overflow-auto rounded-md border border-[var(--mo-line)] bg-[#667085]"
            aria-label="文档打印预览"
          />
        </div>
      </Dialog>

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

      {/* 功能：打印样式统一由打印模块注入，避免与页面级 @media print 规则互相覆盖；时间：2026-10-29；作者：lq。 */}
      <style>{`
        @media print {
          /* 功能：强制打印引擎保留段落高亮、表格底纹和文字背景色；时间：2026-10-25；作者：lq。 */
          *, *::before, *::after { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          [data-table-resize] { display: none !important; }
          [data-cell-id] { box-shadow: none !important; }
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
  const fragments = page.fragments;
  const locked = page.locked;
  // 功能：跨页首片的模型边界，编辑回写时用于保留续页尾部；时间：2026-10-30；作者：lq。
  const fragmentEnds = page.fragmentEnds;
  // 功能：表格跨页片段是分页切片而不是原表，跳过原表映射；续页只读，首片仍可编辑原表；时间：2026-10-29；作者：lq。
  const blocks = page.blocks
    .map((block) => (fragments?.includes(block.id) ? block : blockMap.get(block.id) ?? block))
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
              locked?.includes(block.id) ? (
                // 功能：跨页续片段只读，避免编辑落到错误位置；时间：2026-10-30；作者：lq。
                <StaticParagraph key={block.id} paragraph={block} doc={doc} pageNumber={pageNumber} totalPages={total} />
              ) : (
                <EditableParagraph
                  key={block.id}
                  paragraph={block}
                  doc={doc}
                  pageNumber={pageNumber}
                  totalPages={total}
                  readonly={readonly}
                  handlers={handlers}
                  headEnd={fragmentEnds?.[block.id]}
                />
              )
            ) : block.type === 'table' ? (
              <TableGrid
                key={block.id}
                table={block}
                doc={doc}
                pageNumber={pageNumber}
                totalPages={total}
                readonly={readonly || locked?.includes(block.id) === true}
                handlers={locked?.includes(block.id) ? undefined : handlers}
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
