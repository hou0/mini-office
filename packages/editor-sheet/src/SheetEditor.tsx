import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
} from 'react';
import {
  cellKey,
  type CellRange,
  type CellStyle,
  type SheetModel,
  type SheetSession,
  type WorkbookModel,
} from '@mini-office/document-model';
import { useWorkspaceStore } from '@mini-office/document-model';
import { printCurrentView } from '@mini-office/platform-adapter';
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
import {
  ChevronDown,
  ChevronUp,
  FileWarning,
  Filter,
  Plus,
  Search,
  Trash2,
  X,
} from 'lucide-react';
import { exportElementsToPdf, type ExportProgress } from '@mini-office/export-service';
import {
  columnLabel,
  ensureExtension,
  formatBytes,
  stripExtension,
  uid,
} from '@mini-office/shared';
import { SheetGrid, type EditingCell, type GridAction } from './SheetGrid';
import { SheetToolbar, type BorderPreset } from './SheetToolbar';
import { recalculateWorkbook } from './formula/engine';
import {
  activeSheet as activeSheetOf,
  addSheet,
  applyCellUpdates,
  applyFilter,
  applyStyleToRange,
  cellAt,
  cellDisplay,
  cellRawInput,
  cellStyle,
  clearRangeContents,
  deleteColumns,
  deleteRows,
  deleteSheet,
  extractClipboard,
  findOrCreateStyle,
  insertColumns,
  insertRows,
  mergeSelection,
  parseRangeRef,
  pasteClipboard,
  pasteTsv,
  parseTsv,
  rangeArea,
  rangeRef,
  renameSheet,
  selectionRange,
  selectionStats,
  setCellInput,
  setColumnWidth,
  setFreeze,
  setRowHeight,
  sortRange,
  unmergeSelection,
  uniqueColumnValues,
  updateSheet,
  usedRange,
  type Selection,
  type SheetClipboard,
} from './sheet-utils';

interface HistoryEntry {
  past: WorkbookModel[];
  future: WorkbookModel[];
  lastKey?: string;
  lastTime: number;
}

const historyStore = new Map<string, HistoryEntry>();

function historyFor(id: string): HistoryEntry {
  const existing = historyStore.get(id);
  if (existing) return existing;
  const created: HistoryEntry = { past: [], future: [], lastTime: 0 };
  historyStore.set(id, created);
  return created;
}

interface FilterState {
  column: number;
  range: CellRange;
  autoFilterRange: CellRange;
  values: string[];
  selected: Set<string>;
}

interface ExportState {
  open: boolean;
  name: string;
  progress: ExportProgress | null;
  busy: boolean;
}

export function SheetEditor({ session }: { session: SheetSession }) {
  const workbook = session.data as WorkbookModel;
  const readonly = session.readonly;
  const updateData = useWorkspaceStore((state) => state.updateData);
  const updateView = useWorkspaceStore((state) => state.updateView);
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const exportBytes = useWorkspaceStore((state) => state.exportBytes);
  const openCompatibility = useWorkspaceStore((state) => state.openDialog);
  const notify = useWorkspaceStore((state) => state.notify);

  const [sheetId, setSheetId] = useState(() => {
    const saved = session.view.activeSheetId;
    return saved && workbook.sheets.some((sheet) => sheet.id === saved)
      ? saved
      : workbook.activeSheetId;
  });
  const sheet = workbook.sheets.find((item) => item.id === sheetId) ?? workbook.sheets[0];
  const [selection, setSelection] = useState<Selection>({
    anchor: { row: 0, col: 0 },
    focus: { row: 0, col: 0 },
  });
  const [editing, setEditing] = useState<EditingCell | null>(null);
  const [formulaText, setFormulaText] = useState('');
  const [historyVersion, setHistoryVersion] = useState(0);
  const [zoom, setZoom] = useState(() => session.view.zoom ?? 1);
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [findActive, setFindActive] = useState(0);
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [filterState, setFilterState] = useState<FilterState | null>(null);
  const [filterOpen, setFilterOpen] = useState(false);
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    kind: 'cell' | 'row' | 'col';
    row: number;
    col: number;
  } | null>(null);
  const [sheetDialog, setSheetDialog] = useState<{ type: 'rename' | 'delete'; sheetId: string; name: string } | null>(null);
  const [exportState, setExportState] = useState<ExportState>({
    open: false,
    name: '',
    progress: null,
    busy: false,
  });
  const [replaceText, setReplaceText] = useState('');
  const clipboardRef = useRef<SheetClipboard | null>(null);
  const clipboardTsvRef = useRef('');
  const cancelExportRef = useRef(false);

  const latestWorkbook = useCallback((): WorkbookModel => {
    const latest = useWorkspaceStore.getState().sessions[session.id]?.data as WorkbookModel | undefined;
    return latest ?? workbook;
  }, [session.id, workbook]);

  const history = historyFor(session.id);
  void historyVersion;
  const canUndo = history.past.length > 0;
  const canRedo = history.future.length > 0;

  const pushHistory = useCallback(
    (previous: WorkbookModel, key?: string) => {
      const stack = historyFor(session.id);
      const now = Date.now();
      if (key && stack.lastKey === key && now - stack.lastTime < 900) {
        stack.lastTime = now;
        return;
      }
      stack.past.push(previous);
      if (stack.past.length > 60) stack.past.shift();
      stack.future = [];
      stack.lastKey = key;
      stack.lastTime = now;
      setHistoryVersion((value) => value + 1);
    },
    [session.id],
  );

  const commit = useCallback(
    (next: WorkbookModel, options?: { key?: string; history?: boolean }) => {
      if (readonly) return;
      const previous = latestWorkbook();
      if (options?.history !== false) pushHistory(previous, options?.key);
      const recalculated = recalculateWorkbook(next);
      updateData(session.id, recalculated);
    },
    [latestWorkbook, pushHistory, readonly, session.id, updateData],
  );

  const undo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const previous = stack.past.pop();
    if (!previous) return;
    stack.future.push(latestWorkbook());
    stack.lastKey = undefined;
    updateData(session.id, previous);
    setEditing(null);
    setHistoryVersion((value) => value + 1);
  }, [latestWorkbook, readonly, session.id, updateData]);

  const redo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const next = stack.future.pop();
    if (!next) return;
    stack.past.push(latestWorkbook());
    stack.lastKey = undefined;
    updateData(session.id, next);
    setEditing(null);
    setHistoryVersion((value) => value + 1);
  }, [latestWorkbook, readonly, session.id, updateData]);

  const active = selection.focus;
  const range = selectionRange(selection);
  const activeCell = cellAt(sheet, active.row, active.col);

  useEffect(() => {
    if (editing) return;
    setFormulaText(cellRawInput(cellAt(sheet, active.row, active.col)));
  }, [active.col, active.row, editing, sheet, workbook]);

  useEffect(() => {
    updateView(session.id, { activeSheetId: sheet.id, zoom });
  }, [sheet.id, session.id, updateView, zoom]);

  const focusGrid = useCallback(() => {
    requestAnimationFrame(() => {
      const element = document.querySelector<HTMLElement>(`[data-sheet-grid="${session.id}"]`);
      element?.focus();
    });
  }, [session.id]);

  const handleSelectionChange = useCallback((next: Selection) => {
    setSelection(next);
  }, []);

  const startEdit = useCallback(
    (row: number, col: number, initialText?: string) => {
      if (readonly) return;
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const raw = initialText ?? cellRawInput(cellAt(currentSheet, row, col));
      setEditing({ row, col, initialText: raw });
      setFormulaText(raw);
      setSelection({ anchor: { row, col }, focus: { row, col } });
    },
    [latestWorkbook, readonly, sheet],
  );

  const commitEdit = useCallback(
    (value: string, move: 'down' | 'right' | 'up' | 'left' | 'none') => {
      if (!editing || readonly) return;
      const current = latestWorkbook();
      const next = setCellInput(current, sheet.id, editing.row, editing.col, value);
      const moved =
        move === 'down'
          ? { row: Math.min(sheet.dimension.rows - 1, editing.row + 1), col: editing.col }
          : move === 'right'
            ? { row: editing.row, col: Math.min(sheet.dimension.cols - 1, editing.col + 1) }
            : move === 'up'
              ? { row: Math.max(0, editing.row - 1), col: editing.col }
              : move === 'left'
                ? { row: editing.row, col: Math.max(0, editing.col - 1) }
                : { row: editing.row, col: editing.col };
      commit(next, { key: `cell:${editing.row}:${editing.col}` });
      setEditing(null);
      setSelection({ anchor: moved, focus: moved });
      focusGrid();
    },
    [commit, editing, focusGrid, latestWorkbook, readonly, sheet.dimension.cols, sheet.dimension.rows, sheet.id],
  );

  const cancelEdit = useCallback(() => {
    setEditing(null);
    setFormulaText(cellRawInput(cellAt(sheet, active.row, active.col)));
    focusGrid();
  }, [active.col, active.row, focusGrid, sheet]);

  const applyStyle = useCallback(
    (patch: Partial<CellStyle>) => {
      if (readonly) return;
      let target = range;
      if (rangeArea(target) > 20_000) {
        target = usedRange(sheet);
        notify('info', '选区过大', '已将格式应用到当前数据区域。');
      }
      const next = applyStyleToRange(latestWorkbook(), sheet.id, target, patch);
      commit(next, { key: `style:${JSON.stringify(patch).slice(0, 40)}` });
    },
    [commit, latestWorkbook, notify, range, readonly, sheet],
  );

  const applyBorders = useCallback(
    (preset: BorderPreset) => {
      if (readonly) return;
      const side = { style: 'thin' as const, color: '#1D2939' };
      let next = latestWorkbook();
      if (preset === 'none') {
        next = applyStyleToRange(next, sheet.id, range, { borders: undefined });
      } else if (preset === 'all') {
        next = applyStyleToRange(next, sheet.id, range, {
          borders: { top: side, right: side, bottom: side, left: side },
        });
      } else if (preset === 'bottom') {
        next = applyStyleToRange(next, sheet.id, range, {
          borders: { ...(cellStyle(next, cellAt(sheet, range.r1, range.c1)).borders ?? {}), bottom: side },
        });
      } else {
        const top: CellRange = { r1: range.r1, c1: range.c1, r2: range.r1, c2: range.c2 };
        const bottom: CellRange = { r1: range.r2, c1: range.c1, r2: range.r2, c2: range.c2 };
        const left: CellRange = { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c1 };
        const right: CellRange = { r1: range.r1, c1: range.c2, r2: range.r2, c2: range.c2 };
        next = applyStyleToRange(next, sheet.id, top, { borders: { top: side } });
        next = applyStyleToRange(next, sheet.id, bottom, { borders: { bottom: side } });
        next = applyStyleToRange(next, sheet.id, left, { borders: { left: side } });
        next = applyStyleToRange(next, sheet.id, right, { borders: { right: side } });
      }
      commit(next, { key: 'borders' });
    },
    [commit, latestWorkbook, range, readonly, sheet],
  );

  const toggleMerge = useCallback(() => {
    if (readonly) return;
    const merged = sheet.merges.some(
      (merge) =>
        merge.r1 <= range.r2 && merge.r2 >= range.r1 && merge.c1 <= range.c2 && merge.c2 >= range.c1,
    );
    const next = merged
      ? unmergeSelection(latestWorkbook(), sheet.id, range)
      : mergeSelection(latestWorkbook(), sheet.id, range);
    commit(next, { key: 'merge' });
  }, [commit, latestWorkbook, range, readonly, sheet]);

  const freeze = useCallback(() => {
    const current = sheet.freeze;
    const activeFrozen = !!current && (current.rows > 0 || current.cols > 0);
    if (activeFrozen) {
      commit(setFreeze(latestWorkbook(), sheet.id, 0, 0));
    } else {
      commit(setFreeze(latestWorkbook(), sheet.id, active.row, active.col));
    }
  }, [active.col, active.row, commit, latestWorkbook, sheet]);

  const insertRow = useCallback(() => {
    commit(insertRows(latestWorkbook(), sheet.id, active.row, 1), { key: 'structure' });
  }, [active.row, commit, latestWorkbook, sheet.id]);
  const deleteRow = useCallback(() => {
    commit(deleteRows(latestWorkbook(), sheet.id, active.row, 1), { key: 'structure' });
  }, [active.row, commit, latestWorkbook, sheet.id]);
  const insertColumn = useCallback(() => {
    commit(insertColumns(latestWorkbook(), sheet.id, active.col, 1), { key: 'structure' });
  }, [active.col, commit, latestWorkbook, sheet.id]);
  const deleteColumn = useCallback(() => {
    commit(deleteColumns(latestWorkbook(), sheet.id, active.col, 1), { key: 'structure' });
  }, [active.col, commit, latestWorkbook, sheet.id]);

  const sortByActive = useCallback(
    (ascending: boolean) => {
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const singleCell = range.r1 === range.r2 && range.c1 === range.c2;
      let target = singleCell ? usedRange(currentSheet) : range;
      if (singleCell && currentSheet.autoFilter) {
        // Auto filter ranges include the header row; sort only the data rows.
        target = { ...currentSheet.autoFilter, r1: currentSheet.autoFilter.r1 + 1 };
      }
      if (target.r2 <= target.r1) {
        notify('info', '没有可排序的数据', '请选择包含数据的区域。');
        return;
      }
      if (currentSheet.merges.some((merge) => merge.r1 <= target.r2 && merge.r2 >= target.r1 && merge.c1 <= target.c2 && merge.c2 >= target.c1)) {
        notify('warning', '无法排序', '排序区域包含合并单元格，请先取消合并。');
        return;
      }
      commit(sortRange(current, currentSheet, target, active.col, ascending), { key: 'sort' });
      notify('success', ascending ? '已按当前列升序排序' : '已按当前列降序排序');
    },
    [active.col, commit, latestWorkbook, notify, range, sheet],
  );

  const openFilter = useCallback(() => {
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
    const target = currentSheet.autoFilter ?? usedRange(currentSheet);
    if (target.r2 <= target.r1) {
      notify('info', '没有可筛选的数据');
      return;
    }
    const dataRange = { ...target, r1: target.r1 + 1 };
    const values = uniqueColumnValues(current, currentSheet, active.col, dataRange);
    setFilterState({
      column: active.col,
      range: dataRange,
      autoFilterRange: target,
      values,
      selected: new Set(values),
    });
    setFilterOpen(true);
  }, [active.col, latestWorkbook, notify, sheet]);

  const applyFilterDialog = useCallback(() => {
    if (!filterState) return;
    const current = latestWorkbook();
    const next = applyFilter(current, sheet, filterState.column, filterState.selected, filterState.range);
    commit(updateSheet(next, sheet.id, (item) => ({ ...item, autoFilter: filterState.autoFilterRange })), { key: 'filter' });
    setFilterOpen(false);
    notify('success', '筛选已应用');
  }, [commit, filterState, latestWorkbook, notify, sheet]);

  const clearFilter = useCallback(() => {
    const current = latestWorkbook();
    const target = sheet.autoFilter ?? usedRange(sheet);
    const next = applyFilter(current, sheet, active.col, null, { ...target, r1: target.r1 + 1 });
    commit(next, { key: 'filter-clear' });
    setFilterOpen(false);
    notify('info', '已清除筛选');
  }, [active.col, commit, latestWorkbook, notify, sheet]);

  const findHits = useMemo(() => {
    if (!findQuery) return [] as Array<{ row: number; col: number }>;
    const target = usedRange(sheet);
    const hits: Array<{ row: number; col: number }> = [];
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    for (let row = target.r1; row <= target.r2; row += 1) {
      for (let col = target.c1; col <= target.c2; col += 1) {
        const text = cellDisplay(workbook, sheet, row, col);
        if (!text) continue;
        const haystack = caseSensitive ? text : text.toLowerCase();
        if (haystack.includes(needle)) hits.push({ row, col });
        if (hits.length >= 2000) return hits;
      }
    }
    return hits;
  }, [caseSensitive, findQuery, sheet, workbook]);

  const goToHit = useCallback(
    (index: number) => {
      if (findHits.length === 0) return;
      const normalized = ((index % findHits.length) + findHits.length) % findHits.length;
      setFindActive(normalized);
      const hit = findHits[normalized];
      const next = { anchor: { row: hit.row, col: hit.col }, focus: { row: hit.row, col: hit.col } };
      setSelection(next);
      setFormulaText(cellRawInput(cellAt(sheet, hit.row, hit.col)));
      focusGrid();
    },
    [findHits, focusGrid, sheet],
  );

  const replaceCurrent = useCallback(() => {
    if (readonly || findHits.length === 0 || !findQuery) return;
    const hit = findHits[Math.min(findActive, findHits.length - 1)];
    const cell = cellAt(sheet, hit.row, hit.col);
    if (cell?.formula) {
      notify('warning', '公式单元格', '查找替换不会修改公式本身，请直接编辑公式。');
      return;
    }
    const display = cellDisplay(workbook, sheet, hit.row, hit.col);
    const haystack = caseSensitive ? display : display.toLowerCase();
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    const index = haystack.indexOf(needle);
    if (index < 0) return;
    const replaced = display.slice(0, index) + replaceText + display.slice(index + findQuery.length);
    const next = setCellInput(latestWorkbook(), sheet.id, hit.row, hit.col, replaced);
    commit(next, { key: `replace:${hit.row}:${hit.col}` });
  }, [caseSensitive, commit, findActive, findHits, findQuery, latestWorkbook, notify, readonly, replaceText, sheet, workbook]);

  const replaceAll = useCallback(() => {
    if (readonly || !findQuery) return;
    const current = latestWorkbook();
    let total = 0;
    const updates: Array<{ row: number; col: number; cell: ReturnType<typeof createReplacedCell> }> = [];
    const target = usedRange(sheet);
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    for (let row = target.r1; row <= target.r2; row += 1) {
      for (let col = target.c1; col <= target.c2; col += 1) {
        const cell = cellAt(sheet, row, col);
        if (!cell || cell.formula) continue;
        const display = cellDisplay(workbook, sheet, row, col);
        const haystack = caseSensitive ? display : display.toLowerCase();
        if (!haystack.includes(needle)) continue;
        const replaced = caseSensitive
          ? display.split(findQuery).join(replaceText)
          : replaceCaseInsensitive(display, findQuery, replaceText);
        if (replaced !== display) {
          updates.push({ row, col, cell: createReplacedCell(replaced, cell) });
          total += 1;
        }
      }
    }
    if (updates.length === 0) {
      notify('info', '没有可替换的内容');
      return;
    }
    commit(applyCellUpdates(current, sheet.id, updates), { key: 'replace-all' });
    notify('success', `已替换 ${total} 个单元格`);
  }, [caseSensitive, commit, findQuery, latestWorkbook, notify, readonly, replaceText, sheet, workbook]);

  const copySelection = useCallback(
    async (cut: boolean) => {
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const clipboard = extractClipboard(current, currentSheet, range);
      clipboardRef.current = clipboard;
      const tsv = clipboardToTsvText(current, currentSheet, clipboard);
      clipboardTsvRef.current = tsv;
      try {
        await navigator.clipboard.writeText(tsv);
      } catch {
        // Clipboard permissions may be denied; the internal clipboard still works.
      }
      if (cut && !readonly) {
        commit(clearRangeContents(current, sheet.id, range), { key: 'cut' });
      }
    },
    [commit, latestWorkbook, range, readonly, sheet],
  );

  const pasteSelection = useCallback(async () => {
    if (readonly) return;
    let text = '';
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = '';
    }
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
    if (text && text.trim() && text !== clipboardTsvRef.current) {
      const rows = parseTsv(text);
      commit(pasteTsv(current, currentSheet, active, rows), { key: 'paste' });
      notify('success', `已粘贴 ${rows.length} 行`);
      return;
    }
    const clipboard = clipboardRef.current;
    if (!clipboard) {
      notify('info', '剪贴板为空');
      return;
    }
    const target =
      range.r1 === range.r2 && range.c1 === range.c2
        ? {
            r1: active.row,
            c1: active.col,
            r2: active.row + (clipboard.range.r2 - clipboard.range.r1),
            c2: active.col + (clipboard.range.c2 - clipboard.range.c1),
          }
        : range;
    commit(pasteClipboard(current, currentSheet, target, clipboard), { key: 'paste' });
  }, [active, commit, latestWorkbook, notify, range, readonly, sheet]);

  const handleAction = useCallback(
    (action: GridAction) => {
      if (action === 'copy') void copySelection(false);
      else if (action === 'cut') void copySelection(true);
      else if (action === 'paste') void pasteSelection();
      else if (action === 'undo') undo();
      else if (action === 'redo') redo();
      else if (action === 'selectAll') {
        const target = usedRange(sheet);
        setSelection({ anchor: { row: target.r1, col: target.c1 }, focus: { row: target.r2, col: target.c2 } });
      } else if (action === 'clear') {
        if (!readonly) commit(clearRangeContents(latestWorkbook(), sheet.id, range), { key: 'clear' });
      } else if (action === 'find') setFindOpen(true);
    },
    [commit, copySelection, latestWorkbook, pasteSelection, range, readonly, redo, sheet, undo],
  );

  const handleFill = useCallback(
    (source: CellRange, target: CellRange) => {
      if (readonly) return;
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const clipboard = extractClipboard(current, currentSheet, source);
      const next = pasteClipboard(current, currentSheet, target, clipboard);
      commit(next, { key: 'fill' });
    },
    [commit, latestWorkbook, readonly, sheet],
  );

  const handleCellContextMenu = useCallback(
    (event: ReactMouseEvent, row: number, col: number) => {
      setContextMenu({ x: event.clientX, y: event.clientY, kind: 'cell', row, col });
    },
    [],
  );

  const handleHeaderContextMenu = useCallback(
    (event: ReactMouseEvent, kind: 'row' | 'col', index: number) => {
      setContextMenu({ x: event.clientX, y: event.clientY, kind, row: kind === 'row' ? index : active.row, col: kind === 'col' ? index : active.col });
    },
    [active.col, active.row],
  );

  const contextItems: MenuItem[] = contextMenu
    ? contextMenu.kind === 'cell'
      ? [
          { id: 'copy', label: '复制', shortcut: '⌘C', onSelect: () => void copySelection(false) },
          { id: 'cut', label: '剪切', shortcut: '⌘X', onSelect: () => void copySelection(true) },
          { id: 'paste', label: '粘贴', shortcut: '⌘V', onSelect: () => void pasteSelection() },
          { type: 'separator' },
          { id: 'insert-row', label: '在上方插入行', onSelect: insertRow },
          { id: 'insert-col', label: '在左侧插入列', onSelect: insertColumn },
          { id: 'delete-row', label: '删除当前行', danger: true, onSelect: deleteRow },
          { id: 'delete-col', label: '删除当前列', danger: true, onSelect: deleteColumn },
          { type: 'separator' },
          { id: 'clear', label: '清除内容', onSelect: () => handleAction('clear') },
          { id: 'freeze', label: '冻结到当前单元格', onSelect: freeze },
        ]
      : [
          ...(contextMenu.kind === 'row'
            ? [
                { id: 'insert', label: '在上方插入行', onSelect: insertRow },
                { id: 'delete', label: '删除行', danger: true, onSelect: deleteRow },
              ]
            : [
                { id: 'insert', label: '在左侧插入列', onSelect: insertColumn },
                { id: 'delete', label: '删除列', danger: true, onSelect: deleteColumn },
              ]),
          { type: 'separator' },
          { id: 'sort-asc', label: '升序排序', onSelect: () => sortByActive(true) },
          { id: 'sort-desc', label: '降序排序', onSelect: () => sortByActive(false) },
          { id: 'filter', label: '筛选…', onSelect: openFilter },
        ]
    : [];

  const activeSheetIndex = workbook.sheets.findIndex((item) => item.id === sheet.id);
  const stats = useMemo(() => selectionStats(workbook, sheet, range), [range, sheet, workbook]);

  const switchSheet = useCallback(
    (nextId: string) => {
      setSheetId(nextId);
      updateView(session.id, { activeSheetId: nextId });
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
      setEditing(null);
    },
    [session.id, updateView, workbook],
  );

  const addNewSheet = useCallback(() => {
    if (readonly) return;
    const next = addSheet(latestWorkbook());
    commit(next);
    const newSheet = next.sheets[next.sheets.length - 1];
    setSheetId(newSheet.id);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
  }, [commit, latestWorkbook, readonly]);

  const confirmSheetDialog = useCallback(() => {
    if (!sheetDialog) return;
    if (sheetDialog.type === 'rename') {
      commit(renameSheet(latestWorkbook(), sheetDialog.sheetId, sheetDialog.name), { key: 'rename-sheet' });
    } else {
      const next = deleteSheet(latestWorkbook(), sheetDialog.sheetId);
      commit(next);
      if (sheetDialog.sheetId === sheet.id) {
        setSheetId(next.sheets[0].id);
      }
    }
    setSheetDialog(null);
  }, [commit, latestWorkbook, sheetDialog, sheet.id]);

  const openExport = useCallback(() => {
    setExportState({ open: true, name: `${stripExtension(session.displayName)}.pdf`, progress: null, busy: false });
  }, [session.displayName]);

  const runExport = useCallback(async () => {
    if (exportState.busy) return;
    cancelExportRef.current = false;
    setExportState((state) => ({ ...state, busy: true, progress: { phase: 'preparing', current: 0, total: 1, message: '正在准备表格…' } }));
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    const container = document.createElement('div');
    container.dataset.sheetPrint = 'true';
    try {
      const pages = buildSheetPrintPages(workbook, sheet, container);
      if (pages.length === 0) throw new Error('没有可导出的内容');
      const bytes = await exportElementsToPdf({
        pages,
        fileName: ensureExtension(stripExtension(exportState.name || session.displayName), '.pdf'),
        title: session.displayName,
        scale: 2,
        pageWidthPx: 1123,
        pageHeightPx: 794,
        onProgress: (progress) => setExportState((state) => ({ ...state, progress })),
        isCancelled: () => cancelExportRef.current,
      });
      await exportBytes(
        bytes,
        ensureExtension(stripExtension(exportState.name || session.displayName), '.pdf'),
        'application/pdf',
      );
      setExportState({ open: false, name: '', progress: null, busy: false });
    } catch (error) {
      if ((error as { code?: string })?.code === 'CANCELLED') notify('info', '已取消导出');
      else notify('error', '导出 PDF 失败', error instanceof Error ? error.message : String(error));
      setExportState((state) => ({ ...state, busy: false, progress: null }));
    } finally {
      container.remove();
    }
  }, [exportBytes, exportState.busy, exportState.name, notify, session.displayName, sheet, workbook]);

  const printSheet = useCallback(async () => {
    const container = document.createElement('div');
    container.dataset.sheetPrint = 'true';
    buildSheetPrintPages(workbook, sheet, container);
    const style = document.createElement('style');
    style.dataset.sheetPrintStyle = 'true';
    style.textContent = `@media print {
      body * { visibility: hidden !important; }
      [data-sheet-print], [data-sheet-print] * { visibility: visible !important; }
      [data-sheet-print] { position: absolute !important; left: 0 !important; top: 0 !important; width: auto !important; }
      [data-sheet-print] > div { page-break-after: always; break-after: page; }
      @page { size: A4 landscape; margin: 0; }
    }`;
    document.head.appendChild(style);
    const cleanup = () => {
      container.remove();
      style.remove();
      window.removeEventListener('afterprint', cleanup);
    };
    window.addEventListener('afterprint', cleanup);
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
    const ok = await printCurrentView();
    if (!ok) notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
    setTimeout(cleanup, 30_000);
  }, [notify, sheet, workbook]);

  useEffect(() => {
    const handleFind = () => setFindOpen(true);
    const handleExport = () => openExport();
    const handlePrint = () => printSheet();
    window.addEventListener('mo:find', handleFind);
    window.addEventListener('mo:export-pdf', handleExport);
    window.addEventListener('mo:print', handlePrint);
    return () => {
      window.removeEventListener('mo:find', handleFind);
      window.removeEventListener('mo:export-pdf', handleExport);
      window.removeEventListener('mo:print', handlePrint);
    };
  }, [openExport, printSheet]);

  const formulaCommit = useCallback(() => {
    if (readonly) return;
    if (editing) {
      commitEdit(formulaText, 'none');
      return;
    }
    const next = setCellInput(latestWorkbook(), sheet.id, active.row, active.col, formulaText);
    commit(next, { key: `cell:${active.row}:${active.col}` });
  }, [active.col, active.row, commit, commitEdit, editing, formulaText, latestWorkbook, readonly, sheet.id]);

  const handleJump = useCallback(
    (ref: string) => {
      const parsed = parseRangeRef(ref);
      if (!parsed) {
        notify('warning', '无效的单元格地址', '请输入如 A1 或 A1:C10 的地址。');
        return;
      }
      setSelection({
        anchor: { row: parsed.r1, col: parsed.c1 },
        focus: { row: parsed.r2, col: parsed.c2 },
      });
      focusGrid();
    },
    [focusGrid, notify],
  );

  return (
    <div className="flex h-full min-h-0 w-full">
      <div className="flex min-w-0 flex-1 flex-col">
        {settings.showToolbar ? (
        <SheetToolbar
          workbook={workbook}
          sheet={sheet}
          range={range}
          active={active}
          formulaText={formulaText}
          readonly={readonly}
          canUndo={canUndo}
          canRedo={canRedo}
          zoom={zoom}
          propertiesOpen={settings.showProperties}
          onFormulaChange={setFormulaText}
          onFormulaCommit={formulaCommit}
          onJump={handleJump}
          onStyle={applyStyle}
          onNumberFormat={(format) => applyStyle({ numFmt: format })}
          onBorders={applyBorders}
          onMergeToggle={toggleMerge}
          onFreeze={freeze}
          onInsertRow={insertRow}
          onDeleteRow={deleteRow}
          onInsertColumn={insertColumn}
          onDeleteColumn={deleteColumn}
          onSort={sortByActive}
          onFilter={openFilter}
          onFind={() => setFindOpen(true)}
          onExport={openExport}
          onPrint={printSheet}
          onCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
          onToggleProperties={() => updateSettings({ showProperties: !settings.showProperties })}
          onZoom={setZoom}
          onUndo={undo}
          onRedo={redo}
          onAddSheet={addNewSheet}
        />
        ) : null}

        <div className="relative min-h-0 flex-1" style={{ zoom }}>
          <SheetGrid
            gridId={session.id}
            workbook={workbook}
            sheet={sheet}
            selection={selection}
            editing={editing}
            onSelectionChange={handleSelectionChange}
            onStartEdit={startEdit}
            onCommitEdit={commitEdit}
            onCancelEdit={cancelEdit}
            onColumnResize={(col, width) => commit(setColumnWidth(latestWorkbook(), sheet.id, col, width))}
            onRowResize={(row, height) => commit(setRowHeight(latestWorkbook(), sheet.id, row, height))}
            onCellContextMenu={handleCellContextMenu}
            onHeaderContextMenu={handleHeaderContextMenu}
            onAction={handleAction}
            onFill={handleFill}
          />
          {findOpen ? (
            <SheetFindBar
              query={findQuery}
              replaceText={replaceText}
              caseSensitive={caseSensitive}
              count={findHits.length}
              active={findActive}
              readonly={readonly}
              onQuery={(value) => {
                setFindQuery(value);
                setFindActive(0);
              }}
              onReplace={setReplaceText}
              onCaseSensitive={setCaseSensitive}
              onPrev={() => goToHit(findActive - 1)}
              onNext={() => goToHit(findActive + 1)}
              onReplaceCurrent={replaceCurrent}
              onReplaceAll={replaceAll}
              onClose={() => setFindOpen(false)}
            />
          ) : null}
        </div>

        <div className="mo-glass mo-no-print flex h-8 shrink-0 items-center gap-1 border-t border-[var(--mo-line)] bg-[var(--mo-surface)] px-2">
          <div className="mo-scroll flex min-w-0 flex-1 items-center gap-1 overflow-x-auto">
            {workbook.sheets.map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={() => switchSheet(item.id)}
                onDoubleClick={() => setSheetDialog({ type: 'rename', sheetId: item.id, name: item.name })}
                onContextMenu={(event) => {
                  event.preventDefault();
                  setSheetDialog({ type: 'rename', sheetId: item.id, name: item.name });
                }}
                className={`h-6 shrink-0 rounded-t-md border-b-2 px-3 text-[11px] transition-colors ${
                  item.id === sheet.id
                    ? 'border-[var(--mo-accent)] bg-[var(--mo-surface)] font-medium text-[var(--mo-text-primary)]'
                    : 'border-transparent text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]'
                }`}
                style={item.tabColor ? { color: item.tabColor } : undefined}
              >
                {item.name}
                {item.state !== 'visible' ? '（隐藏）' : ''}
              </button>
            ))}
            <IconButton label="添加工作表" onClick={addNewSheet} disabled={readonly}>
              <Plus size={13} />
            </IconButton>
          </div>
          <div className="flex shrink-0 items-center gap-3 pl-3 text-[10px] tabular-nums text-[var(--mo-text-muted)]">
            {stats.numericCount > 0 ? (
              <>
                <span>求和 {formatNumber(stats.sum)}</span>
                <span>平均 {formatNumber(stats.average)}</span>
                <span>计数 {stats.count}</span>
              </>
            ) : stats.count > 0 ? (
              <span>非空 {stats.count}</span>
            ) : (
              <span>
                {columnLabel(active.col)}
                {active.row + 1}
              </span>
            )}
            <span className="text-[var(--mo-text-faint)]">工作表 {activeSheetIndex + 1}/{workbook.sheets.length}</span>
          </div>
        </div>
      </div>

      {settings.showProperties ? (
        <SheetPropertiesPanel
          workbook={workbook}
          sheet={sheet}
          range={range}
          active={active}
          session={session}
          onStyle={applyStyle}
          onClose={() => updateSettings({ showProperties: false })}
          onShowCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
        />
      ) : null}

      {contextMenu ? (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          items={contextItems}
          onClose={() => setContextMenu(null)}
        />
      ) : null}

      <Dialog
        open={filterOpen}
        onClose={() => setFilterOpen(false)}
        title="筛选数据"
        description={filterState ? `按 ${columnLabel(filterState.column)} 列筛选 ${rangeRef(filterState.range)}` : ''}
        width="sm"
        footer={
          <>
            <Button variant="ghost" onClick={clearFilter}>
              清除筛选
            </Button>
            <Button onClick={() => setFilterOpen(false)}>取消</Button>
            <Button variant="primary" onClick={applyFilterDialog}>
              应用筛选
            </Button>
          </>
        }
      >
        {filterState ? (
          <div className="space-y-2">
            <label className="flex items-center justify-between rounded-md bg-[var(--mo-surface-2)] px-2 py-1.5 text-[12px]">
              <span>全选</span>
              <input
                type="checkbox"
                checked={filterState.selected.size === filterState.values.length && filterState.values.length > 0}
                onChange={(event) =>
                  setFilterState((state) =>
                    state
                      ? {
                          ...state,
                          selected: event.target.checked ? new Set(state.values) : new Set(),
                        }
                      : state,
                  )
                }
                className="h-3.5 w-3.5 accent-[var(--mo-accent)]"
              />
            </label>
            <div className="max-h-[260px] space-y-0.5 overflow-auto">
              {filterState.values.map((value) => (
                <label
                  key={value || '(空白)'}
                  className="flex items-center justify-between rounded px-2 py-1 text-[12px] hover:bg-[var(--mo-surface-2)]"
                >
                  <span className="truncate">{value || '(空白)'}</span>
                  <input
                    type="checkbox"
                    checked={filterState.selected.has(value)}
                    onChange={(event) =>
                      setFilterState((state) => {
                        if (!state) return state;
                        const selected = new Set(state.selected);
                        if (event.target.checked) selected.add(value);
                        else selected.delete(value);
                        return { ...state, selected };
                      })
                    }
                    className="h-3.5 w-3.5 accent-[var(--mo-accent)]"
                  />
                </label>
              ))}
            </div>
          </div>
        ) : null}
      </Dialog>

      <Dialog
        open={sheetDialog !== null}
        onClose={() => setSheetDialog(null)}
        title="工作表"
        width="sm"
        footer={
          <>
            <Button onClick={() => setSheetDialog(null)}>取消</Button>
            <Button
              variant={sheetDialog?.type === 'delete' ? 'danger' : 'primary'}
              onClick={confirmSheetDialog}
            >
              {sheetDialog?.type === 'delete' ? '删除' : '重命名'}
            </Button>
          </>
        }
      >
        {sheetDialog ? (
          <div className="space-y-3">
            {sheetDialog.type === 'rename' ? (
              <Field label="工作表名称">
                <Input
                  value={sheetDialog.name}
                  onChange={(event) =>
                    setSheetDialog((state) => (state ? { ...state, name: event.target.value } : state))
                  }
                  autoFocus
                />
              </Field>
            ) : (
              <p className="text-[12px]">确定删除工作表“{sheet.name}”吗？此操作可通过撤销恢复。</p>
            )}
            <div className="flex items-center gap-2">
              <Button
                size="sm"
                variant="ghost"
                className="text-[var(--mo-danger)]"
                onClick={() =>
                  setSheetDialog({ type: 'delete', sheetId: sheet.id, name: sheet.name })
                }
              >
                <Trash2 size={12} /> 删除工作表
              </Button>
              <span className="text-[10px] text-[var(--mo-text-muted)]">
                当前工作表共 {Object.keys(sheet.cells).length} 个有内容的单元格
              </span>
            </div>
          </div>
        ) : null}
      </Dialog>

      <Dialog
        open={exportState.open}
        onClose={() => {
          if (!exportState.busy) setExportState((state) => ({ ...state, open: false }));
        }}
        title="导出为 PDF"
        description="将当前工作表的数据区域按 A4 横向分页导出，保留字体、颜色、对齐与边框。"
        width="sm"
        footer={
          <>
            <Button onClick={() => setExportState((state) => ({ ...state, open: false }))} disabled={exportState.busy}>
              取消
            </Button>
            <Button variant="primary" loading={exportState.busy} onClick={() => void runExport()}>
              开始导出
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <Field label="文件名">
            <Input
              value={exportState.name}
              onChange={(event) => setExportState((state) => ({ ...state, name: event.target.value }))}
              disabled={exportState.busy}
            />
          </Field>
          <div className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
            工作表：{sheet.name} · 数据区域 {rangeRef(usedRange(sheet))} · 纸张 A4 横向
          </div>
          {exportState.progress ? (
            <div className="space-y-1.5">
              <ProgressBar
                value={
                  exportState.progress.total > 0
                    ? exportState.progress.current / exportState.progress.total
                    : 0
                }
              />
              <div className="flex justify-between text-[10px] text-[var(--mo-text-muted)]">
                <span>{exportState.progress.message}</span>
                <span>
                  {exportState.progress.current}/{exportState.progress.total}
                </span>
              </div>
            </div>
          ) : null}
          {exportState.busy ? (
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
    </div>
  );
}

function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return '—';
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)));
}

function replaceCaseInsensitive(text: string, needle: string, replacement: string): string {
  if (!needle) return text;
  const lower = text.toLowerCase();
  const target = needle.toLowerCase();
  let result = '';
  let index = 0;
  while (index < text.length) {
    const found = lower.indexOf(target, index);
    if (found < 0) {
      result += text.slice(index);
      break;
    }
    result += text.slice(index, found) + replacement;
    index = found + needle.length;
  }
  return result;
}

function createReplacedCell(value: string, existing: { styleId?: string }) {
  return { value, styleId: existing.styleId };
}

function clipboardToTsvText(
  workbook: WorkbookModel,
  sheet: SheetModel,
  clipboard: SheetClipboard,
): string {
  const rows: string[] = [];
  for (let row = clipboard.range.r1; row <= clipboard.range.r2; row += 1) {
    const fields: string[] = [];
    for (let col = clipboard.range.c1; col <= clipboard.range.c2; col += 1) {
      fields.push(cellDisplay(workbook, sheet, row, col));
    }
    rows.push(fields.join('\t'));
  }
  return rows.join('\n');
}

function SheetFindBar(props: {
  query: string;
  replaceText: string;
  caseSensitive: boolean;
  count: number;
  active: number;
  readonly: boolean;
  onQuery: (value: string) => void;
  onReplace: (value: string) => void;
  onCaseSensitive: (value: boolean) => void;
  onPrev: () => void;
  onNext: () => void;
  onReplaceCurrent: () => void;
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
            value={props.query}
            onChange={(event) => props.onQuery(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                if (event.shiftKey) props.onPrev();
                else props.onNext();
              }
              if (event.key === 'Escape') props.onClose();
            }}
            placeholder="查找内容"
            className="h-7 pl-7 text-[12px]"
          />
        </div>
        <span className="w-14 text-center text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {props.count > 0 ? `${props.active + 1}/${props.count}` : '0/0'}
        </span>
        <IconButton label="上一个" onClick={props.onPrev} disabled={props.count === 0}>
          <ChevronUp size={14} />
        </IconButton>
        <IconButton label="下一个" onClick={props.onNext} disabled={props.count === 0}>
          <ChevronDown size={14} />
        </IconButton>
        <IconButton label="关闭" onClick={props.onClose}>
          <X size={14} />
        </IconButton>
      </div>
      {!props.readonly ? (
        <div className="mt-2 flex items-center gap-1.5">
          <Input
            value={props.replaceText}
            onChange={(event) => props.onReplace(event.target.value)}
            placeholder="替换为"
            className="h-7 flex-1 text-[12px]"
          />
          <Button size="sm" onClick={props.onReplaceCurrent} disabled={props.count === 0}>
            替换
          </Button>
          <Button size="sm" onClick={props.onReplaceAll} disabled={props.count === 0}>
            全部
          </Button>
        </div>
      ) : null}
      <label className="mt-2 flex cursor-pointer items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
        <input
          type="checkbox"
          checked={props.caseSensitive}
          onChange={(event) => props.onCaseSensitive(event.target.checked)}
          className="h-3 w-3 accent-[var(--mo-accent)]"
        />
        区分大小写
      </label>
    </div>
  );
}

function SheetPropertiesPanel({
  workbook,
  sheet,
  range,
  active,
  session,
  onStyle,
  onClose,
  onShowCompatibility,
}: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  range: CellRange;
  active: { row: number; col: number };
  session: SheetSession;
  onStyle: (patch: Partial<CellStyle>) => void;
  onClose: () => void;
  onShowCompatibility: () => void;
}) {
  const cell = cellAt(sheet, active.row, active.col);
  const style = cellStyle(workbook, cell);
  return (
    <aside className="mo-glass mo-no-print flex w-[268px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]">
      <div className="flex h-10 items-center justify-between border-b border-[var(--mo-line)] px-3">
        <span className="text-[12px] font-semibold">单元格属性</span>
        <IconButton label="关闭属性面板" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </div>
      <div className="space-y-4 p-3 text-[11px]">
        <section className="space-y-1.5">
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">地址</span>
            <span className="font-mono">
              {columnLabel(active.col)}
              {active.row + 1}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">选区</span>
            <span className="font-mono">{rangeRef(range)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">内容</span>
            <span className="max-w-[150px] truncate">{cellRawInput(cell) || '（空）'}</span>
          </div>
        </section>
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">字体</h3>
          <Field label="字体">
            <Select value={style.fontFamily ?? workbook.defaultFont} onChange={(event) => onStyle({ fontFamily: event.target.value })}>
              {[workbook.defaultFont, '宋体', '微软雅黑', '等线', '思源黑体', 'Arial', 'Times New Roman'].filter(
                (value, index, list) => list.indexOf(value) === index,
              ).map((font) => (
                <option key={font} value={font}>
                  {font}
                </option>
              ))}
            </Select>
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Field label="字号">
              <Input
                type="number"
                value={style.fontSizePt ?? workbook.defaultFontSizePt}
                onChange={(event) => onStyle({ fontSizePt: Number(event.target.value) || 11 })}
              />
            </Field>
            <Field label="水平对齐">
              <Select
                value={style.horizontalAlign ?? 'general'}
                onChange={(event) => onStyle({ horizontalAlign: event.target.value as CellStyle['horizontalAlign'] })}
              >
                <option value="general">常规</option>
                <option value="left">左对齐</option>
                <option value="center">居中</option>
                <option value="right">右对齐</option>
              </Select>
            </Field>
          </div>
        </section>
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">数字格式</h3>
          <Input
            value={style.numFmt ?? 'General'}
            onChange={(event) => onStyle({ numFmt: event.target.value })}
          />
        </section>
        <section className="space-y-1.5 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">工作表</h3>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">名称</span>
            <span>{sheet.name}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">有效单元格</span>
            <span>{Object.keys(sheet.cells).length}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">冻结</span>
            <span>
              {sheet.freeze
                ? `${sheet.freeze.rows} 行 / ${sheet.freeze.cols} 列`
                : '未冻结'}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">文件大小</span>
            <span>{formatBytes(session.sourceSize)}</span>
          </div>
        </section>
        <button
          type="button"
          onClick={onShowCompatibility}
          className="flex w-full items-center gap-2 rounded-md bg-[var(--mo-surface-2)] px-2 py-2 text-left text-[11px] hover:bg-[var(--mo-surface-3)]"
        >
          <FileWarning size={14} className="text-[var(--mo-warning)]" />
          <span className="flex-1">查看兼容性报告</span>
          <span className="text-[var(--mo-text-faint)]">{session.capability.warnings.length}</span>
        </button>
      </div>
    </aside>
  );
}

export function buildSheetPrintPages(
  workbook: WorkbookModel,
  sheet: SheetModel,
  container: HTMLElement,
): HTMLElement[] {
  const PAGE_WIDTH = 1123;
  const PAGE_HEIGHT = 794;
  const MARGIN = 36;
  const contentWidth = PAGE_WIDTH - MARGIN * 2;
  const contentHeight = PAGE_HEIGHT - MARGIN * 2;
  const target = usedRange(sheet);
  const columnWidths: number[] = [];
  for (let col = target.c1; col <= target.c2; col += 1) {
    columnWidths.push(sheet.colWidths[col] ?? sheet.defaultColWidth);
  }
  const totalWidth = columnWidths.reduce((sum, width) => sum + width, 0) || 100;
  const scale = Math.min(1, contentWidth / totalWidth);
  const rowHeight = sheet.defaultRowHeight;
  const titleHeight = 30;
  const rowsPerPage = Math.max(
    1,
    Math.floor((contentHeight - titleHeight) / Math.max(10, rowHeight * scale)) - 1,
  );
  const headerRow = target.r2 > target.r1 ? target.r1 : null;
  const bodyStart = headerRow !== null ? target.r1 + 1 : target.r1;
  const pageRows: Array<{ start: number; end: number }> = [];
  if (bodyStart > target.r2) {
    pageRows.push({ start: target.r1, end: target.r2 });
  } else {
    for (let start = bodyStart; start <= target.r2; start += rowsPerPage) {
      pageRows.push({ start, end: Math.min(target.r2, start + rowsPerPage - 1) });
    }
  }
  if (pageRows.length === 0) pageRows.push({ start: target.r1, end: target.r1 });

  container.style.cssText = `position:fixed;left:-200000px;top:0;width:${PAGE_WIDTH}px;background:#fff;z-index:-1;`;
  const pages: HTMLElement[] = [];
  pageRows.forEach((chunk, pageIndex) => {
    const page = document.createElement('div');
    page.dataset.sheetPrintPage = 'true';
    page.style.cssText = `width:${PAGE_WIDTH}px;height:${PAGE_HEIGHT}px;background:#fff;color:#1D2939;box-sizing:border-box;padding:${MARGIN}px;font-family:等线,PingFang SC,sans-serif;position:relative;overflow:hidden;`;
    const title = document.createElement('div');
    title.style.cssText = `height:${titleHeight}px;display:flex;align-items:center;justify-content:space-between;font-size:12px;color:#667085;border-bottom:1px solid #E4E9F2;margin-bottom:6px;`;
    title.innerHTML = `<span style="font-weight:600;color:#1D2939;">${escapeHtml(sheet.name)}</span><span>Mini Office · 第 ${pageIndex + 1} / ${pageRows.length} 页</span>`;
    page.appendChild(title);

    const wrapper = document.createElement('div');
    wrapper.style.cssText = `width:${contentWidth}px;height:${contentHeight - titleHeight}px;overflow:hidden;`;
    const table = document.createElement('table');
    table.style.cssText = `border-collapse:collapse;table-layout:fixed;width:${totalWidth}px;transform:scale(${scale});transform-origin:top left;font-size:${workbook.defaultFontSizePt}pt;`;
    const colgroup = document.createElement('colgroup');
    columnWidths.forEach((width) => {
      const col = document.createElement('col');
      col.style.width = `${width}px`;
      colgroup.appendChild(col);
    });
    table.appendChild(colgroup);
    if (headerRow !== null) {
      const thead = document.createElement('thead');
      thead.appendChild(buildPrintRow(workbook, sheet, headerRow, target, true));
      table.appendChild(thead);
    }
    const tbody = document.createElement('tbody');
    for (let row = chunk.start; row <= chunk.end; row += 1) {
      tbody.appendChild(buildPrintRow(workbook, sheet, row, target, false));
    }
    table.appendChild(tbody);
    wrapper.appendChild(table);
    page.appendChild(wrapper);
    container.appendChild(page);
    pages.push(page);
  });
  document.body.appendChild(container);
  return pages;
}

function buildPrintRow(
  workbook: WorkbookModel,
  sheet: SheetModel,
  row: number,
  target: CellRange,
  header: boolean,
): HTMLTableRowElement {
  const tr = document.createElement('tr');
  for (let col = target.c1; col <= target.c2; col += 1) {
    const td = document.createElement('td');
    const cell = cellAt(sheet, row, col);
    const style = cellStyle(workbook, cell);
    const text = cellDisplay(workbook, sheet, row, col);
    td.textContent = text;
    const horizontal =
      style.horizontalAlign && style.horizontalAlign !== 'general'
        ? style.horizontalAlign
        : typeof cell?.value === 'number'
          ? 'right'
          : 'left';
    td.style.cssText = `border:1px solid #D0D5DD;padding:2px 5px;vertical-align:${
      style.verticalAlign ?? 'bottom'
    };text-align:${horizontal};overflow:hidden;white-space:pre-wrap;word-break:break-word;background:${
      header ? '#EEF2F9' : style.fill ?? '#FFFFFF'
    };color:${style.color ?? '#1D2939'};font-weight:${header || style.bold ? 700 : 400};font-style:${
      style.italic ? 'italic' : 'normal'
    };font-size:${style.fontSizePt ?? workbook.defaultFontSizePt}pt;font-family:${
      style.fontFamily ?? workbook.defaultFont
    };`;
    if (style.borders?.top && style.borders.top.style !== 'none') {
      td.style.borderTop = `${style.borders.top.style === 'medium' ? 2 : style.borders.top.style === 'thick' ? 3 : 1}px solid ${style.borders.top.color ?? '#000'}`;
    }
    if (style.borders?.right && style.borders.right.style !== 'none') {
      td.style.borderRight = `${style.borders.right.style === 'medium' ? 2 : style.borders.right.style === 'thick' ? 3 : 1}px solid ${style.borders.right.color ?? '#000'}`;
    }
    if (style.borders?.bottom && style.borders.bottom.style !== 'none') {
      td.style.borderBottom = `${style.borders.bottom.style === 'medium' ? 2 : style.borders.bottom.style === 'thick' ? 3 : 1}px solid ${style.borders.bottom.color ?? '#000'}`;
    }
    if (style.borders?.left && style.borders.left.style !== 'none') {
      td.style.borderLeft = `${style.borders.left.style === 'medium' ? 2 : style.borders.left.style === 'thick' ? 3 : 1}px solid ${style.borders.left.color ?? '#000'}`;
    }
    tr.appendChild(td);
  }
  return tr;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
