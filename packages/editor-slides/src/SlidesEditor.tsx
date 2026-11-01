// 功能：提供 PPT 页面编辑工作区、对象操作、备注、演示与导出；时间：2026-10-24；作者：lq。
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { SlideDeck, SlideModel, SlidesSession } from '@mini-office/document-model';
import { useWorkspaceStore } from '@mini-office/document-model';
import { isTauriRuntime, printCurrentView, setPresentationFullscreen } from '@mini-office/platform-adapter';
import {
  Button,
  Dialog,
  Field,
  IconButton,
  Input,
  ProgressBar,
  SegmentedControl,
  Tooltip,
} from '@mini-office/ui';
import {
  Bold,
  ChevronLeft,
  ChevronRight,
  Copy,
  Download,
  FileWarning,
  Image as ImageIcon,
  Italic,
  List,
  Maximize,
  MonitorPlay,
  PanelLeftClose,
  PanelLeftOpen,
  Play,
  Plus,
  StickyNote,
  Trash2,
  Type,
  Undo2,
  Redo2,
  X,
} from 'lucide-react';
import { exportElementsToPdf, exportElementsToZip, type ExportProgress } from '@mini-office/export-service';
import { stripExtension, uid } from '@mini-office/shared';
import { SlideRenderer, SlideSurface } from './SlideRenderer';
import { copySlide, replaceParagraphText, updateElementTree } from './editing';
// 功能：WPS 风格 Ribbon 布局——主题插件声明 chrome.toolbar='ribbon' 时启用；时间：2026-10-31；作者：lq。
import { SlidesRibbon } from './SlidesRibbon';

export function SlidesEditor({ session, chromeToolbar }: { session: SlidesSession; chromeToolbar?: 'modern' | 'classic-two-row' | 'ribbon' }) {
  const deck = session.data as SlideDeck;
  const updateView = useWorkspaceStore((state) => state.updateView);
  const updateData = useWorkspaceStore((state) => state.updateData);
  const settings = useWorkspaceStore((state) => state.settings);
  const exportBytes = useWorkspaceStore((state) => state.exportBytes);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const notify = useWorkspaceStore((state) => state.notify);

  const [index, setIndex] = useState(() => Math.min(session.view.slideIndex ?? 0, Math.max(0, deck.slides.length - 1)));
  const [thumbnailsOpen, setThumbnailsOpen] = useState(true);
  const [notesOpen, setNotesOpen] = useState(() => session.view.showNotes ?? false);
  const [presenting, setPresenting] = useState(false);
  const [presentNotes, setPresentNotes] = useState(false);
  const [scale, setScale] = useState(1);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportKind, setExportKind] = useState<'pdf' | 'images'>('pdf');
  const [exportName, setExportName] = useState('');
  const [exporting, setExporting] = useState(false);
  const [printing, setPrinting] = useState(false);
  const [progress, setProgress] = useState<ExportProgress | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  // 功能：记录当前正在编辑文本的对象，支持双击进入和失焦退出；时间：2026-10-28；作者：lq。
  const [editingId, setEditingId] = useState<string | null>(null);
  const cancelRef = useRef(false);
  const historyRef = useRef<{ past: SlideDeck[]; future: SlideDeck[] }>({ past: [], future: [] });

  const slide = deck.slides[index] ?? deck.slides[0];
  const canvasRef = useRef<HTMLDivElement>(null);
  const exportLayerRef = useRef<HTMLDivElement>(null);

  const updateDeck = useCallback(
    (change: (current: SlideDeck) => SlideDeck) => {
      if (session.readonly) return;
      updateData(session.id, (current) => {
        historyRef.current.past = [...historyRef.current.past.slice(-39), structuredClone(current as SlideDeck)];
        historyRef.current.future = [];
        return change(current as SlideDeck);
      }, { dirty: true });
    },
    [session.id, session.readonly, updateData],
  );

  const undo = useCallback(() => {
    const previous = historyRef.current.past.at(-1);
    if (!previous) return;
    const current = structuredClone(deck);
    historyRef.current.past = historyRef.current.past.slice(0, -1);
    historyRef.current.future = [...historyRef.current.future, current];
    updateData(session.id, previous, { dirty: true });
    setSelectedId(null);
  }, [deck, session.id, updateData]);

  const redo = useCallback(() => {
    const nextDeck = historyRef.current.future.at(-1);
    if (!nextDeck) return;
    historyRef.current.future = historyRef.current.future.slice(0, -1);
    historyRef.current.past = [...historyRef.current.past, structuredClone(deck)];
    updateData(session.id, nextDeck, { dirty: true });
    setSelectedId(null);
  }, [deck, session.id, updateData]);

  const changeText = useCallback((elementId: string, value: string) => {
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: updateElementTree(item.elements, elementId, (element) => {
          if (element.type === 'text') return { ...element, paragraphs: replaceParagraphText(element.paragraphs, value) };
          if (element.type === 'shape' && element.paragraphs) return { ...element, paragraphs: replaceParagraphText(element.paragraphs, value) };
          return element;
        }),
      }),
    }));
  }, [index, updateDeck]);

  const toggleTextStyle = useCallback((key: 'bold' | 'italic') => {
    if (!selectedId) return;
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: updateElementTree(item.elements, selectedId, (element) => {
          if (element.type !== 'text' && element.type !== 'shape') return element;
          const paragraphs = element.type === 'text' ? element.paragraphs : element.paragraphs ?? [];
          const enabled = !paragraphs.some((paragraph) => paragraph.runs.some((run) => run.style[key]));
          return { ...element, paragraphs: paragraphs.map((paragraph) => ({ ...paragraph, runs: paragraph.runs.map((run) => ({ ...run, style: { ...run.style, [key]: enabled } })) })) };
        }),
      }),
    }));
  }, [index, selectedId, updateDeck]);

  // 功能：拖拽选中对象时按画布缩放比例换算坐标，提供直觉的排版体验；时间：2026-10-24；作者：lq。
  const moveElement = useCallback((elementId: string, dx: number, dy: number) => {
    if (!selectedId || selectedId !== elementId) return;
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: updateElementTree(item.elements, elementId, (element) => ({ ...element, x: element.x + dx / Math.max(scale, 0.01), y: element.y + dy / Math.max(scale, 0.01) })),
      }),
    }));
  }, [index, scale, selectedId, updateDeck]);

  const moveSelected = useCallback((dx: number, dy: number) => {
    if (!selectedId) return;
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: updateElementTree(item.elements, selectedId, (element) => ({ ...element, x: element.x + dx, y: element.y + dy })),
      }),
    }));
  }, [index, selectedId, updateDeck]);

  const resizeSelected = useCallback((dw: number, dh: number) => {
    if (!selectedId) return;
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: updateElementTree(item.elements, selectedId, (element) => ({ ...element, w: Math.max(24, element.w + dw), h: Math.max(18, element.h + dh) })),
      }),
    }));
  }, [index, selectedId, updateDeck]);

  const deleteSelected = useCallback(() => {
    if (!selectedId) return;
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : { ...item, elements: updateElementTree(item.elements, selectedId, () => null) }),
    }));
    setSelectedId(null);
    setEditingId(null);
  }, [index, selectedId, updateDeck]);

  const addText = useCallback(() => {
    const id = uid('el-text');
    updateDeck((current) => ({
      ...current,
      slides: current.slides.map((item, itemIndex) => itemIndex !== index ? item : {
        ...item,
        elements: [...item.elements, {
          id,
          type: 'text' as const,
          x: current.widthPx * 0.18,
          y: current.heightPx * 0.42,
          w: current.widthPx * 0.64,
          h: current.heightPx * 0.14,
          paragraphs: [{ runs: [{ text: '双击编辑文字', style: { fontSizePt: 24, color: '#1D2939' } }], align: 'center' }],
          verticalAlign: 'middle' as const,
        }],
      }),
    }));
    setSelectedId(id);
    setEditingId(null);
  }, [index, updateDeck]);

  const duplicateSlide = useCallback(() => {
    const source = deck.slides[index];
    if (!source) return;
    const duplicate: SlideModel = { ...copySlide(source), index: index + 1 };
    updateDeck((current) => ({ ...current, slides: current.slides.flatMap((item, itemIndex) => itemIndex === index ? [item, duplicate] : [item]).map((item, itemIndex) => ({ ...item, index: itemIndex })) }));
    setIndex(index + 1);
    setSelectedId(null);
  }, [deck.slides, index, updateDeck]);

  const addSlide = useCallback(() => {
    const newSlide: SlideModel = { id: uid('slide'), index: deck.slides.length, title: `幻灯片 ${deck.slides.length + 1}`, elements: [] };
    updateDeck((current) => ({ ...current, slides: [...current.slides, newSlide].map((item, itemIndex) => ({ ...item, index: itemIndex })) }));
    setIndex(deck.slides.length);
    setSelectedId(null);
  }, [deck.slides.length, updateDeck]);

  const deleteSlide = useCallback(() => {
    if (deck.slides.length <= 1) return;
    updateDeck((current) => ({ ...current, slides: current.slides.filter((_, itemIndex) => itemIndex !== index).map((item, itemIndex) => ({ ...item, index: itemIndex })) }));
    setIndex(Math.min(index, deck.slides.length - 2));
    setSelectedId(null);
  }, [deck.slides.length, index, updateDeck]);

  const startPresentation = useCallback(async () => {
    const fullscreen = await setPresentationFullscreen(true);
    setPresenting(true);
    if (!fullscreen) {
      notify('info', '窗口内演示', '系统全屏不可用，已切换为窗口内演示模式。');
    }
  }, [notify]);

  const stopPresentation = useCallback(async () => {
    setPresenting(false);
    await setPresentationFullscreen(false);
  }, []);

  useEffect(() => () => {
    void setPresentationFullscreen(false);
  }, []);

  useEffect(() => {
    if (!presenting || isTauriRuntime()) return;
    const onChange = () => {
      if (!document.fullscreenElement) setPresenting(false);
    };
    document.addEventListener('fullscreenchange', onChange);
    return () => document.removeEventListener('fullscreenchange', onChange);
  }, [presenting]);

  const goTo = useCallback(
    (next: number) => {
      const clamped = Math.max(0, Math.min(deck.slides.length - 1, next));
      setIndex(clamped);
      updateView(session.id, { slideIndex: clamped });
    },
    [deck.slides.length, session.id, updateView],
  );

  const next = useCallback(() => goTo(index + 1), [goTo, index]);
  const prev = useCallback(() => goTo(index - 1), [goTo, index]);

  useEffect(() => {
    setExportName(`${stripExtension(session.displayName)}.pdf`);
  }, [session.displayName]);

  // Fit the stage to the available area.
  useEffect(() => {
    const element = canvasRef.current;
    if (!element) return;
    const update = () => {
      const rect = element.getBoundingClientRect();
      const horizontal = Math.max(200, rect.width - 64);
      const vertical = Math.max(160, rect.height - 48);
      setScale(Math.min(horizontal / deck.widthPx, vertical / deck.heightPx));
    };
    update();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(update);
    observer.observe(element);
    return () => observer.disconnect();
  }, [deck.heightPx, deck.widthPx, notesOpen]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement;
      if (['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName)) return;
      if (target.isContentEditable && editingId) {
        if (event.key === 'Escape') {
          event.preventDefault();
          target.blur();
          setEditingId(null);
        }
        return;
      }
      const mod = event.metaKey || event.ctrlKey;
      if (!presenting && mod && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        if (event.shiftKey) redo(); else undo();
        return;
      }
      if (!presenting && mod && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        redo();
        return;
      }
      if (!presenting && selectedId && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
        event.preventDefault();
        const distance = event.shiftKey ? 10 : 2;
        if (event.altKey) resizeSelected(event.key === 'ArrowLeft' ? -distance : event.key === 'ArrowRight' ? distance : 0, event.key === 'ArrowUp' ? -distance : event.key === 'ArrowDown' ? distance : 0);
        else moveSelected(event.key === 'ArrowLeft' ? -distance : event.key === 'ArrowRight' ? distance : 0, event.key === 'ArrowUp' ? -distance : event.key === 'ArrowDown' ? distance : 0);
        return;
      }
      if (!presenting && selectedId && (event.key === 'Backspace' || event.key === 'Delete')) {
        event.preventDefault();
        deleteSelected();
        return;
      }
      if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
        event.preventDefault();
        next();
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        prev();
      } else if (event.key === 'Home') {
        event.preventDefault();
        goTo(0);
      } else if (event.key === 'End') {
        event.preventDefault();
        goTo(deck.slides.length - 1);
      } else if (event.key === 'Escape' && presenting) {
        void stopPresentation();
      } else if ((event.key === 'n' || event.key === 'N') && presenting) {
        setPresentNotes((value) => !value);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [deck.slides.length, deleteSelected, editingId, goTo, moveSelected, next, presenting, prev, redo, resizeSelected, selectedId, stopPresentation, undo]);

  useEffect(() => {
    const onPresent = () => void startPresentation();
    const onExport = () => {
      setExportOpen(true);
      setExportKind('pdf');
      setProgress(null);
    };
    window.addEventListener('mo:present', onPresent);
    window.addEventListener('mo:export-pdf', onExport);
    return () => {
      window.removeEventListener('mo:present', onPresent);
      window.removeEventListener('mo:export-pdf', onExport);
    };
  }, [startPresentation]);

  useEffect(() => {
    updateView(session.id, { presenting, showNotes: notesOpen });
  }, [notesOpen, presenting, session.id, updateView]);

  useEffect(() => {
    document.body.style.overflow = presenting ? 'hidden' : '';
    return () => {
      document.body.style.overflow = '';
    };
  }, [presenting]);

  const stage = useMemo(
    () => <SlideRenderer deck={deck} slide={slide} scale={scale} editable={!presenting && !session.readonly} selectedId={selectedId} editingId={editingId} onSelect={(id) => { setSelectedId(id); setEditingId(null); }} onStartEditing={(id) => { setSelectedId(id); setEditingId(id); }} onEndEditing={() => setEditingId(null)} onTextChange={changeText} onMove={moveElement} />,
    [changeText, deck, editingId, moveElement, presenting, scale, selectedId, session.readonly, slide],
  );

  const printSlides = useCallback(async () => {
    setPrinting(true);
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    const cleanup = () => setPrinting(false);
    window.addEventListener('afterprint', cleanup, { once: true });
    const ok = await printCurrentView();
    if (!ok) {
      notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
      setPrinting(false);
      return;
    }
    setTimeout(cleanup, 20_000);
  }, [notify]);

  const startExport = useCallback(async () => {
    if (exporting) return;
    const layer = exportLayerRef.current;
    const pages = layer
      ? Array.from(layer.querySelectorAll<HTMLElement>('[data-slide-export]'))
      : [];
    if (pages.length === 0) {
      notify('error', '导出准备失败', '没有可导出的幻灯片。');
      return;
    }
    cancelRef.current = false;
    setExporting(true);
    setProgress({ phase: 'preparing', current: 0, total: pages.length, message: '正在准备导出…' });
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    try {
      if (exportKind === 'pdf') {
        const bytes = await exportElementsToPdf({
          pages,
          fileName: `${stripExtension(exportName || session.displayName)}.pdf`,
          title: session.displayName,
          scale: 2,
          pageWidthPx: deck.widthPx,
          pageHeightPx: deck.heightPx,
          onProgress: setProgress,
          isCancelled: () => cancelRef.current,
        });
        await exportBytes(bytes, `${stripExtension(exportName || session.displayName)}.pdf`, 'application/pdf');
      } else {
        const bytes = await exportElementsToZip({
          pages,
          scale: 2,
          onProgress: setProgress,
          isCancelled: () => cancelRef.current,
        });
        const zipName = `${stripExtension(exportName || session.displayName)}.zip`;
        await exportBytes(bytes, zipName, 'application/zip');
      }
      setExportOpen(false);
      setProgress(null);
    } catch (error) {
      if ((error as { code?: string })?.code === 'CANCELLED') {
        notify('info', '已取消导出');
      } else {
        notify('error', '导出失败', error instanceof Error ? error.message : String(error));
      }
      setProgress(null);
    } finally {
      setExporting(false);
    }
  }, [deck.heightPx, deck.widthPx, exportBytes, exportKind, exportName, exporting, notify, session.displayName]);

  if (!slide) {
    return (
      <div className="flex h-full items-center justify-center text-[13px] text-[var(--mo-text-muted)]">
        演示文稿没有幻灯片
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-0 w-full flex-col">
      {settings.showToolbar ? (
        chromeToolbar === 'ribbon' ? (
          <SlidesRibbon
            readonly={session.readonly}
            canUndo={!historyRef.current.past.length}
            canRedo={!historyRef.current.future.length}
            thumbnailsOpen={thumbnailsOpen}
            notesOpen={notesOpen}
            selectedId={selectedId}
            slideIndex={index}
            slideCount={deck.slides.length}
            onToggleThumbnails={() => setThumbnailsOpen((value) => !value)}
            onToggleNotes={() => setNotesOpen((value) => !value)}
            onPrev={prev}
            onNext={next}
            onBold={() => toggleTextStyle('bold')}
            onItalic={() => toggleTextStyle('italic')}
            onUndo={undo}
            onRedo={redo}
            onAddText={addText}
            onDuplicate={duplicateSlide}
            onAddSlide={addSlide}
            onDeleteSlide={deleteSlide}
            onPresent={() => void startPresentation()}
            onExport={() => openExport(setExportOpen, setExportKind, setProgress, setExportName, session.displayName)}
            onCompatibility={() => openDialog({ type: 'compatibility', sessionId: session.id })}
          />
        ) : (
      <div className="mo-glass mo-no-print flex h-11 shrink-0 items-center gap-1 border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-2">
        <IconButton label={thumbnailsOpen ? '隐藏缩略图' : '显示缩略图'} onClick={() => setThumbnailsOpen((value) => !value)}>
          {thumbnailsOpen ? <PanelLeftClose size={15} /> : <PanelLeftOpen size={15} />}
        </IconButton>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        {!session.readonly ? (
          <>
            <Tooltip label="粗体"><IconButton label="粗体" onClick={() => toggleTextStyle('bold')} disabled={!selectedId}><Bold size={15} /></IconButton></Tooltip>
            <Tooltip label="斜体"><IconButton label="斜体" onClick={() => toggleTextStyle('italic')} disabled={!selectedId}><Italic size={15} /></IconButton></Tooltip>
            <Tooltip label="撤销 (⌘Z)"><IconButton label="撤销" onClick={undo} disabled={!historyRef.current.past.length}><Undo2 size={15} /></IconButton></Tooltip>
            <Tooltip label="重做 (⇧⌘Z)"><IconButton label="重做" onClick={redo} disabled={!historyRef.current.future.length}><Redo2 size={15} /></IconButton></Tooltip>
          </>
        ) : null}
        <IconButton label="上一页" onClick={prev} disabled={index <= 0}>
          <ChevronLeft size={16} />
        </IconButton>
        <span className="min-w-[60px] text-center text-[11px] tabular-nums text-[var(--mo-text-muted)]">
          {index + 1} / {deck.slides.length}
        </span>
        <IconButton label="下一页" onClick={next} disabled={index >= deck.slides.length - 1}>
          <ChevronRight size={16} />
        </IconButton>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        {!session.readonly ? (
          <>
            <Tooltip label="插入文本框"><IconButton label="插入文本框" onClick={addText}><Type size={15} /></IconButton></Tooltip>
            <Tooltip label="复制当前页"><IconButton label="复制当前页" onClick={duplicateSlide}><Copy size={15} /></IconButton></Tooltip>
            <Tooltip label="新增空白页"><IconButton label="新增空白页" onClick={addSlide}><Plus size={15} /></IconButton></Tooltip>
            <Tooltip label="删除当前页"><IconButton label="删除当前页" onClick={deleteSlide} disabled={deck.slides.length <= 1}><Trash2 size={15} /></IconButton></Tooltip>
            <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
          </>
        ) : null}
        <div className="min-w-0 flex-1 truncate text-[12px] font-medium">{slide.title ?? `幻灯片 ${index + 1}`}</div>
        {!session.readonly && selectedId ? <span className="hidden text-[10px] text-[var(--mo-text-faint)] xl:block">方向键移动 · ⌥方向键缩放</span> : null}
        <span className="mr-2 rounded bg-[var(--mo-surface-2)] px-1.5 py-0.5 text-[10px] text-[var(--mo-text-muted)]">
          {deck.aspect === 'custom' ? `${Math.round(deck.widthPx)}×${Math.round(deck.heightPx)}` : deck.aspect}
        </span>
        <span className="mr-2 text-[10px] text-[var(--mo-text-faint)]">{Math.round(scale * 100)}%</span>
        <Tooltip label="演讲稿备注 (N)">
          <IconButton label="显示备注" active={notesOpen} onClick={() => setNotesOpen((value) => !value)}>
            <StickyNote size={15} />
          </IconButton>
        </Tooltip>
        <Tooltip label="F5 全屏演示">
          <Button size="sm" variant="soft" icon={<Play size={13} />} onClick={() => void startPresentation()}>
            演示
          </Button>
        </Tooltip>
        <IconButton label="导出" onClick={() => openExport(setExportOpen, setExportKind, setProgress, setExportName, session.displayName)}>
          <Download size={15} />
        </IconButton>
        <IconButton
          label="兼容性报告"
          onClick={() => openDialog({ type: 'compatibility', sessionId: session.id })}
        >
          <FileWarning size={15} className="text-[var(--mo-warning)]" />
        </IconButton>
      </div>
        )
      ) : null}

      <div className="flex min-h-0 flex-1">
        {thumbnailsOpen ? (
          <aside className="mo-glass mo-no-print flex w-[190px] shrink-0 flex-col border-r border-[var(--mo-line)] bg-[var(--mo-surface)]">
            <div className="flex h-9 items-center justify-between border-b border-[var(--mo-line)] px-3 text-[11px] font-medium text-[var(--mo-text-muted)]">
              <span>幻灯片</span>
              <span className="tabular-nums">{deck.slides.length}</span>
            </div>
            <div className="mo-scroll flex-1 space-y-2 p-2.5">
              {deck.slides.map((item, itemIndex) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => { goTo(itemIndex); setSelectedId(null); }}
                  className={`group flex w-full gap-2 rounded-md p-1 text-left transition-colors ${
                    itemIndex === index ? 'bg-[var(--mo-accent-soft)]' : 'hover:bg-[var(--mo-surface-2)]'
                  }`}
                >
                  <span className="w-4 shrink-0 pt-1 text-right text-[10px] tabular-nums text-[var(--mo-text-muted)]">
                    {itemIndex + 1}
                  </span>
                  <div className="min-w-0">
                    <div
                      className={`overflow-hidden rounded border bg-white shadow-sm ${
                        itemIndex === index ? 'border-[var(--mo-accent)]' : 'border-[var(--mo-line)]'
                      }`}
                      style={{ width: 142, height: (142 / deck.widthPx) * deck.heightPx }}
                    >
                      <SlideRenderer deck={deck} slide={item} scale={142 / deck.widthPx} />
                    </div>
                    <div className="mt-1 flex items-center gap-1">
                      <span className="min-w-0 flex-1 truncate text-[10px] text-[var(--mo-text-muted)]">
                        {item.title ?? `幻灯片 ${itemIndex + 1}`}
                      </span>
                      {item.notes ? <StickyNote size={9} className="shrink-0 text-[var(--mo-text-faint)]" /> : null}
                    </div>
                  </div>
                </button>
              ))}
            </div>
          </aside>
        ) : null}

        <div ref={canvasRef} onClick={() => setSelectedId(null)} className="mo-canvas relative flex min-w-0 flex-1 flex-col overflow-hidden bg-[var(--mo-canvas)]">
          <div className="flex min-h-0 flex-1 items-center justify-center p-6">
            <div className="shadow-pop" style={{ borderRadius: 2, overflow: 'hidden' }}>
              {stage}
            </div>
          </div>
          {notesOpen ? (
            <div className="mo-glass mo-no-print max-h-[28%] shrink-0 overflow-auto border-t border-[var(--mo-line)] bg-[var(--mo-surface)] p-3">
              <div className="mb-1 flex items-center gap-2 text-[11px] font-semibold text-[var(--mo-text-muted)]">
                <StickyNote size={12} /> 演讲者备注
              </div>
              {session.readonly ? (
                <p className="whitespace-pre-wrap text-[12px] leading-6 text-[var(--mo-text-primary)]">{slide.notes || '此幻灯片没有备注。'}</p>
              ) : (
                <textarea
                  value={slide.notes ?? ''}
                  onChange={(event) => updateDeck((current) => ({ ...current, slides: current.slides.map((item, itemIndex) => itemIndex === index ? { ...item, notes: event.target.value } : item) }))}
                  placeholder="写下这页的讲解提示…"
                  className="min-h-[64px] w-full resize-y rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface-2)] px-2.5 py-2 text-[12px] leading-6 text-[var(--mo-text-primary)] outline-none focus:border-[var(--mo-accent)]"
                />
              )}
            </div>
          ) : null}
        </div>
      </div>

      {presenting ? (
        <PresentationOverlay
          deck={deck}
          slide={slide}
          index={index}
          notesOpen={presentNotes}
          onNext={next}
          onPrev={prev}
          onExit={() => void stopPresentation()}
          onToggleNotes={() => setPresentNotes((value) => !value)}
        />
      ) : null}

      {printing ? (
        <div
          data-slide-print-container
          aria-hidden
          style={{ position: 'fixed', left: 0, top: 0, width: deck.widthPx, zIndex: -1, background: '#fff' }}
        >
          {deck.slides.map((item) => (
            <div
              key={`print-${item.id}`}
              data-slide-print-page
              style={{ width: deck.widthPx, height: deck.heightPx, overflow: 'hidden', pageBreakAfter: 'always' }}
            >
              <SlideSurface deck={deck} slide={item} />
            </div>
          ))}
          <style>{`
            @media print {
              body * { visibility: hidden !important; }
              [data-slide-print-container], [data-slide-print-container] * { visibility: visible !important; }
              [data-slide-print-container] { position: absolute !important; left: 0 !important; top: 0 !important; width: auto !important; background: #fff !important; }
              [data-slide-print-page] { page-break-after: always; break-after: page; }
              [data-slide-print-page]:last-child { page-break-after: auto; break-after: auto; }
              @page { size: ${deck.widthPx / 96}in ${deck.heightPx / 96}in; margin: 0; }
            }
          `}</style>
        </div>
      ) : null}

      {exportOpen ? (
        <div
          ref={exportLayerRef}
          aria-hidden
          style={{
            position: 'fixed',
            left: -200_000,
            top: 0,
            width: deck.widthPx,
            pointerEvents: 'none',
            opacity: 1,
          }}
        >
          {deck.slides.map((item) => (
            <div
              key={item.id}
              data-slide-export
              style={{ width: deck.widthPx, height: deck.heightPx, overflow: 'hidden' }}
            >
              <SlideSurface deck={deck} slide={item} />
            </div>
          ))}
        </div>
      ) : null}

      <Dialog
        open={exportOpen}
        onClose={() => {
          if (!exporting) setExportOpen(false);
        }}
        title="导出演示文稿"
        description="每页幻灯片按原始 16:9/4:3 比例导出，中文字体以当前系统字体渲染。"
        width="sm"
        footer={
          <>
            <Button onClick={() => setExportOpen(false)} disabled={exporting}>
              取消
            </Button>
            <Button variant="primary" loading={exporting} onClick={startExport}>
              开始导出
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <Field label="导出格式">
            <SegmentedControl
              value={exportKind}
              onChange={(value) => {
                setExportKind(value);
                setExportName(
                  `${stripExtension(session.displayName)}${value === 'pdf' ? '.pdf' : '-幻灯片.zip'}`,
                );
              }}
              options={[
                { value: 'pdf', label: <span className="inline-flex items-center gap-1"><MonitorPlay size={12} /> PDF</span> },
                { value: 'images', label: <span className="inline-flex items-center gap-1"><ImageIcon size={12} /> 图片 ZIP</span> },
              ]}
              className="w-full"
            />
          </Field>
          <Field label="文件名">
            <Input
              value={exportName}
              onChange={(event) => setExportName(event.target.value)}
              disabled={exporting}
            />
          </Field>
          <div className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
            共 {deck.slides.length} 页 · {deck.aspect === 'custom' ? '自定义比例' : deck.aspect} ·{' '}
            {Math.round(deck.widthPx)} × {Math.round(deck.heightPx)} px
          </div>
          {progress ? (
            <div className="space-y-1.5">
              <ProgressBar value={progress.total > 0 ? progress.current / progress.total : 0} />
              <div className="flex justify-between text-[10px] text-[var(--mo-text-muted)]">
                <span>{progress.message}</span>
                <span>
                  {progress.current}/{progress.total}
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
                cancelRef.current = true;
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

function openExport(
  setOpen: (open: boolean) => void,
  setKind: (kind: 'pdf' | 'images') => void,
  setProgress: (progress: ExportProgress | null) => void,
  setName: (name: string) => void,
  fileName: string,
): void {
  setKind('pdf');
  setProgress(null);
  setName(`${stripExtension(fileName)}.pdf`);
  setOpen(true);
}

function PresentationOverlay({
  deck,
  slide,
  index,
  notesOpen,
  onNext,
  onPrev,
  onExit,
  onToggleNotes,
}: {
  deck: SlideDeck;
  slide: SlideModel;
  index: number;
  notesOpen: boolean;
  onNext: () => void;
  onPrev: () => void;
  onExit: () => void;
  onToggleNotes: () => void;
}) {
  const [scale, setScale] = useState(1);
  useEffect(() => {
    const update = () => {
      const width = window.innerWidth - 48;
      const height = window.innerHeight - 48;
      setScale(Math.min(width / deck.widthPx, height / deck.heightPx));
    };
    update();
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, [deck.heightPx, deck.widthPx]);

  return (
    <div className="mo-fade-in fixed inset-0 z-[120] flex items-center justify-center overflow-hidden bg-black">
      <div className="absolute left-1/2 top-4 z-10 flex -translate-x-1/2 items-center gap-1 rounded-full border border-white/10 bg-black/40 px-2 py-1 text-white opacity-0 shadow-[0_10px_30px_rgba(0,0,0,.35)] backdrop-blur-md transition-opacity duration-300 hover:opacity-100">
        <IconButton label="上一页" onClick={onPrev} className="text-white hover:bg-white/20">
          <ChevronLeft size={18} />
        </IconButton>
        <span className="px-2 text-[12px] tabular-nums">
          {index + 1} / {deck.slides.length}
        </span>
        <IconButton label="下一页" onClick={onNext} className="text-white hover:bg-white/20">
          <ChevronRight size={18} />
        </IconButton>
        <span className="mx-1 h-4 w-px bg-white/30" />
        <IconButton label="备注" onClick={onToggleNotes} className="text-white hover:bg-white/20">
          <StickyNote size={16} />
        </IconButton>
        <IconButton label="退出演示 (Esc)" onClick={onExit} className="text-white hover:bg-white/20">
          <X size={18} />
        </IconButton>
      </div>
      <button type="button" aria-label="上一页" className="absolute left-0 top-0 h-full w-[18%] cursor-w-resize" onClick={onPrev} />
      <button type="button" aria-label="下一页" className="absolute right-0 top-0 h-full w-[18%] cursor-e-resize" onClick={onNext} />
      <div
        className="shadow-[0_30px_100px_rgba(0,0,0,0.6)]"
        style={{ width: deck.widthPx * scale, height: deck.heightPx * scale, overflow: 'hidden' }}
      >
        <div
          style={{
            width: deck.widthPx,
            height: deck.heightPx,
            transform: `scale(${scale})`,
            transformOrigin: 'top left',
          }}
        >
          <SlideSurface deck={deck} slide={slide} />
        </div>
      </div>
      {notesOpen ? (
        <div className="absolute bottom-6 left-6 max-h-[30%] w-[min(440px,42vw)] overflow-auto rounded-xl border border-white/10 bg-black/65 p-4 text-white shadow-[0_12px_40px_rgba(0,0,0,.35)] backdrop-blur-md">
          <div className="mb-1 flex items-center gap-2 text-[11px] uppercase tracking-wide text-white/60">
            <StickyNote size={12} /> 备注
          </div>
          <p className="whitespace-pre-wrap text-[13px] leading-6">{slide.notes || '无备注'}</p>
        </div>
      ) : null}
      <div className="pointer-events-none absolute bottom-4 left-1/2 -translate-x-1/2 rounded-full border border-white/10 bg-black/35 px-3 py-1 text-[10.5px] text-white/55 backdrop-blur-sm">
        ← → 换页　　N 备注　　Esc 退出
      </div>
    </div>
  );
}
