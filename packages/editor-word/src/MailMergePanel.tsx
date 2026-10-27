import { useEffect, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import { getSourceBytes, useWorkspaceStore, type DocBlock, type DocModel, type WordSession, type WorkbookModel } from '@mini-office/document-model';
import { getPlatformAdapter, printCurrentView } from '@mini-office/platform-adapter';
import { Button, Dialog, Field, Input, ProgressBar, Select } from '@mini-office/ui';
import { stripExtension } from '@mini-office/shared';
import { exportElementsToPdf } from '@mini-office/export-service';
import { mapParagraphs } from './doc-utils';
import { serializeDocx } from './serialize';
import { MergeReport, mountMergeReport, type MergeBlockRenderer } from './MergeReport';
import { collectMergeFields, extractMergeData, filterMergeRecords, MAX_MERGE_BATCH, mergeDocument, mergeFileName, readMergeWorkbook, type MergeCondition, type MergeData } from './mail-merge';

const OPERATORS: { value: MergeCondition['op']; label: string }[] = [
  { value: 'contains', label: '包含' }, { value: 'eq', label: '等于' }, { value: 'ne', label: '不等于' },
  { value: 'gt', label: '大于' }, { value: 'gte', label: '大于等于' }, { value: 'lt', label: '小于' }, { value: 'lte', label: '小于等于' },
  { value: 'empty', label: '为空' }, { value: 'notEmpty', label: '不为空' },
];
type OutputKind = 'print' | 'pdf' | 'docx' | 'docx-zip' | 'pdf-zip';
/** 功能：在同一文档切换标签时保留数据快照、筛选和勾选，不自动刷新源数据；时间：2026-10-24；作者：lq。 */
type MergeSource = { workbook: WorkbookModel; name: string };
interface MergeDraft { source?: MergeSource; sheetId: string; headerRow: number; conditions: MergeCondition[]; search: string; sort: string; descending: boolean; selected: number[]; allowEmpty: boolean; namePattern: string; output: OutputKind }
const mergeDrafts = new Map<string, MergeDraft>();


/** 功能：从数据选择到字段插入、条件筛选、真实预览和批量输出的一体化操作面板；时间：2026-10-24；作者：lq。 */
export function MailMergePanel({ open, session, doc, onClose, onInsertField, renderBlocks }: {
  open: boolean; session: WordSession; doc: DocModel; onClose: () => void;
  onInsertField: (field: string) => void; renderBlocks: MergeBlockRenderer;
}) {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const notify = useWorkspaceStore((state) => state.notify);
  const cached = useRef(mergeDrafts.get(session.id)).current;
  const [source, setSource] = useState<MergeSource | undefined>(cached?.source);
  const [sheetId, setSheetId] = useState(cached?.sheetId ?? '');
  const [headerRow, setHeaderRow] = useState(cached?.headerRow ?? 1);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [openWorkbook, setOpenWorkbook] = useState('');
  const [dialogOpen, setDialogOpen] = useState(false);
  const [conditions, setConditions] = useState<MergeCondition[]>(cached?.conditions ?? []);
  const [search, setSearch] = useState(cached?.search ?? '');
  const [sort, setSort] = useState(cached?.sort ?? '');
  const [descending, setDescending] = useState(cached?.descending ?? false);
  const [selected, setSelected] = useState<Set<number>>(new Set(cached?.selected ?? []));
  const [tablePage, setTablePage] = useState(0);
  const [previewIndex, setPreviewIndex] = useState(0);
  const [allowEmpty, setAllowEmpty] = useState(cached?.allowEmpty ?? false);
  const [namePattern, setNamePattern] = useState(cached?.namePattern ?? '报告_{序号}');
  const [output, setOutput] = useState<OutputKind>(cached?.output ?? 'print');
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState('');
  const [previewZoom, setPreviewZoom] = useState('fit');
  const [previewWidth, setPreviewWidth] = useState(440);
  const previewRef = useRef<HTMLDivElement>(null);
  const cancel = useRef(false);
  const printCleanup = useRef<(() => void) | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const workbooks = Object.values(sessions).filter((item) => item.kind === 'xlsx');
  const templateFields = useMemo(() => collectMergeFields(doc), [doc]);
  const result = useMemo<{ data?: MergeData; error?: string }>(() => {
    if (!source) return {};
    try { return { data: extractMergeData(source.workbook, sheetId, headerRow) }; }
    catch (cause) { return { error: cause instanceof Error ? cause.message : String(cause) }; }
  }, [source, sheetId, headerRow]);
  const data = result.data;
  // 功能：按数据对象变化初始化选择，避免重复挂载清空已恢复的勾选；时间：2026-10-24；作者：lq。
  const selectionDataRef = useRef(data);
  const filtered = useMemo(() => {
    const records = filterMergeRecords(data?.records ?? [], conditions, search);
    if (sort) records.sort((a, b) => {
      const left = a.raw[sort], right = b.raw[sort];
      const numeric = left !== null && right !== null && String(left).trim() !== '' && String(right).trim() !== '' && Number.isFinite(Number(left)) && Number.isFinite(Number(right));
      const comparison = numeric ? Number(left) - Number(right) : (a.values[sort] ?? '').localeCompare(b.values[sort] ?? '', 'zh-CN', { numeric: true });
      return (descending ? -1 : 1) * comparison;
    });
    return records;
  }, [data, conditions, search, sort, descending]);
  const chosen = useMemo(() => filtered.filter((record) => selected.has(record.rowNumber)), [filtered, selected]);
  const issues = useMemo(() => chosen.flatMap((record) => [
    ...record.errors,
    ...templateFields.flatMap((field) => !data?.fields.includes(field) ? [`缺少数据列：${field}`] : !allowEmpty && !record.values[field]?.trim() ? [`字段为空：${field}`] : []),
  ].map((message) => `第 ${record.rowNumber} 行：${message}`)), [chosen, templateFields, data, allowEmpty]);
  const current = chosen[Math.min(previewIndex, Math.max(0, chosen.length - 1))];
  const previewDoc = useMemo(() => {
    if (!current || !templateFields.length || templateFields.some((field) => !Object.hasOwn(current.values, field))) return undefined;
    return mergeDocument(doc, current);
  }, [current, doc, templateFields]);
  useEffect(() => {
    if (selectionDataRef.current === data) return;
    selectionDataRef.current = data;
    setSelected(new Set(data?.records.map((record) => record.rowNumber) ?? []));
    setConditions([]); setSearch(''); setSort(''); setPreviewIndex(0); setTablePage(0); setAllowEmpty(false);
  }, [data]);
  useEffect(() => { setTablePage(0); setPreviewIndex(0); }, [search, conditions, sort, descending]);
  useEffect(() => () => { cancel.current = true; printCleanup.current?.(); }, []);
  useEffect(() => {
    mergeDrafts.set(session.id, { source, sheetId, headerRow, conditions, search, sort, descending, selected: [...selected], allowEmpty, namePattern, output });
    for (const id of mergeDrafts.keys()) if (!sessions[id]) mergeDrafts.delete(id);
  }, [session.id, sessions, source, sheetId, headerRow, conditions, search, sort, descending, selected, allowEmpty, namePattern, output]);
  useEffect(() => {
    const element = previewRef.current;
    if (!dialogOpen || !open || !element) return;
    const observer = new ResizeObserver(() => setPreviewWidth(element.clientWidth));
    observer.observe(element);
    return () => observer.disconnect();
  }, [dialogOpen, open]);
  const chooseSource = (workbook: WorkbookModel, name: string) => {
    // 功能：冻结本次数据快照，后续 Excel 编辑不会悄悄改变预览与输出；时间：2026-10-24；作者：lq。
    setSource({ workbook: structuredClone(workbook), name });
    setSheetId(workbook.sheets.find((sheet) => sheet.state === 'visible')?.id ?? workbook.sheets[0]?.id ?? '');
    setHeaderRow(1); setError(''); setStatus('');
  };
  const importFile = async (file: File) => {
    setLoading(true); setError('');
    try {
      if (file.size > 50 * 1024 * 1024) throw new Error('数据文件不能超过 50 MB。');
      chooseSource(await readMergeWorkbook(new Uint8Array(await file.arrayBuffer()), file.name), file.name);
    } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); }
    finally { setLoading(false); }
  };
  const closeDialog = () => { if (busy) return; printCleanup.current?.(); printCleanup.current = null; setDialogOpen(false); };
  const multipleSections = doc.blocks.some((block) => block.type === 'paragraph' && block.sectionEnd);
  const badNameFields = [...namePattern.matchAll(/\{([^{}]+)\}/g)].map((match) => match[1]).filter((field) => field !== '序号' && !data?.fields.includes(field));
  const invalidName = (output === 'docx-zip' || output === 'pdf-zip') && badNameFields.length > 0;
  const canGenerate = !multipleSections && !invalidName && !!chosen.length && chosen.length <= MAX_MERGE_BATCH && !!templateFields.length && !issues.length;
  const generate = async () => {
    if (!canGenerate || busy) return;
    cancel.current = false; setBusy(true); setProgress(0); setError('');
    printCleanup.current?.(); printCleanup.current = null;
    const container = document.createElement('div');
    container.style.cssText = 'position:fixed;left:-100000px;top:0;background:white;color:black;';
    document.body.appendChild(container);
    const mounts: ReturnType<typeof mountMergeReport>[] = [];
    const cleanup = () => { mounts.forEach((mount) => mount.dispose()); container.remove(); };
    let keepForPrint = false;
    try {
      const records = [...chosen];
      const template = structuredClone(doc);
      const sourceBytes = getSourceBytes(session.id)?.slice();
      const zip = new JSZip();
      const documents: DocModel[] = [];
      let archiveBytes = 0;
      const addArchiveFile = (name: string, bytes: Uint8Array) => { archiveBytes += bytes.byteLength; if (archiveBytes > 200 * 1024 * 1024) throw new Error('报告文件合计超过 200 MB，请减少本批记录数。'); zip.file(name, bytes); };
      const allPages: HTMLElement[] = [];
      const baseName = `${stripExtension(session.displayName)}_合并报告`;
      for (let index = 0; index < records.length; index++) {
        if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
        const record = records[index];
        setStatus(`正在生成 ${index + 1} / ${records.length} 份（数据第 ${record.rowNumber} 行）…`);
        const merged = mergeDocument(template, record);
        const mount = mountMergeReport(merged, renderBlocks, container); mounts.push(mount);
        let layout;
        try { layout = await mount.ready; } catch (cause) { throw new Error(`第 ${record.rowNumber} 行：${cause instanceof Error ? cause.message : String(cause)}`); }
        if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
        // 功能：将实测的长表分页结果同步给 DOCX，成品与预览使用同一块序列；时间：2026-10-24；作者：lq。
        const pageBlocks: DocBlock[] = [];
        layout.slices.forEach((slice, pageIndex) => { if (pageIndex) pageBlocks.push({ type: 'pageBreak', id: `merge-break-${pageIndex}` }); pageBlocks.push(...slice.blocks); });
        const paged = { ...merged, blocks: pageBlocks };
        documents.push(paged);
        const fileName = mergeFileName(namePattern, record, index);
        if (output === 'docx-zip') addArchiveFile(`${fileName}.docx`, await serializeDocx(paged, { sourceBytes, name: fileName, materialized: true }));
        if (output === 'pdf-zip') addArchiveFile(`${fileName}.pdf`, await exportElementsToPdf({ pages: layout.pages, fileName, pageWidthPx: doc.page.widthPx, pageHeightPx: doc.page.heightPx, isCancelled: () => cancel.current }));
        if (output === 'print' || output === 'pdf') {
          allPages.push(...layout.pages);
          if (allPages.length > 500) throw new Error('合并结果超过 500 页，请缩小筛选范围后分批输出。');
        } else { mount.dispose(); mounts.pop(); }
        setProgress(Math.round((index + 1) / records.length * 85));
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      }
      if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
      setStatus('正在整理输出…');
      if (output === 'print') {
        container.className = 'mo-merge-print-root';
        allPages.at(-1)?.setAttribute('data-merge-last', 'true');
        const style = document.createElement('style');
        // 功能：只打印选中的报告，隐藏模板和交互界面，逐份独立分页；时间：2026-10-24；作者：lq。
        style.textContent = `@media print { html, body { height:auto !important;overflow:visible !important; } body > :not(.mo-merge-print-root) { display:none !important; } body * { visibility:hidden !important; } .mo-word-pages { display:none !important; } .mo-merge-print-root, .mo-merge-print-root * { visibility:visible !important; } .mo-merge-print-root { position:absolute !important;left:0 !important;top:0 !important;width:100% !important; } .mo-merge-print-root .mo-merge-measure { display:none !important; } .mo-merge-print-root [data-merge-page] { break-after:page;page-break-after:always;transform:none !important; } .mo-merge-print-root [data-merge-last] { break-after:auto;page-break-after:auto; } .mo-merge-print-root .flex { display:block !important; } @page { size:${doc.page.widthPx}px ${doc.page.heightPx}px; margin:0; } }`;
        container.appendChild(style);
        keepForPrint = true;
        printCleanup.current = cleanup;
        if (!await printCurrentView()) throw new Error('当前平台不能打开打印，请选择合并 PDF 后打印。');
        setStatus(`已准备 ${records.length} 份报告、${allPages.length} 页。打印面板中确认打印机和份数；关闭本窗口会清理临时打印页面。`);
      } else {
        let bytes: Uint8Array;
        let extension: string;
        let mime: string;
        if (output === 'docx') {
          // 功能：合并 DOCX 的总页数字段按报告所在节计算，编号逐份重启；时间：2026-10-24；作者：lq。
          const sections = documents.map((value) => {
            const convert = (blocks: DocBlock[]) => mapParagraphs(blocks, (paragraph) => ({ ...paragraph, runs: paragraph.runs.map((run) => run.field === 'NUMPAGES' ? { ...run, field: 'SECTIONPAGES' } : run) }));
            return { ...value, blocks: convert(value.blocks), header: value.header ? { ...value.header, blocks: convert(value.header.blocks), edited: true } : undefined, footer: value.footer ? { ...value.footer, blocks: convert(value.footer.blocks), edited: true } : undefined };
          });
          bytes = await serializeDocx(sections[0], { sourceBytes, name: baseName, sections, materialized: true });
          extension = 'docx'; mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        } else if (output === 'pdf') {
          bytes = await exportElementsToPdf({ pages: allPages, fileName: baseName, pageWidthPx: doc.page.widthPx, pageHeightPx: doc.page.heightPx, isCancelled: () => cancel.current, onProgress: (value) => { setStatus(value.message); setProgress(85 + Math.round(value.current / Math.max(1, value.total) * 12)); } });
          extension = 'pdf'; mime = 'application/pdf';
        } else {
          bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }, () => { if (cancel.current) throw new Error('已取消生成，本次结果未保存。'); });
          extension = 'zip'; mime = 'application/zip';
        }
        if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
        await getPlatformAdapter().saveBytes(bytes, { suggestedName: `${baseName}.${extension}`, forcePicker: true }, mime);
        setStatus(`已生成并保存 ${records.length} 份报告。`);
        notify('success', '邮件合并完成', `${records.length} 份报告已保存，模板和源数据保留。`);
      }
      setProgress(100);
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : String(cause);
      setError(message); setStatus('');
      if (keepForPrint) { keepForPrint = false; printCleanup.current = null; }
    } finally { if (!keepForPrint) cleanup(); setBusy(false); }
  };
  if (!open) return null;
  return <>
    <aside className="mo-no-print flex w-[350px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]" aria-label="邮件合并面板">
      <div className="flex items-center justify-between border-b border-[var(--mo-line)] p-3"><strong className="text-sm">邮件合并 / 批量报告</strong><Button size="sm" variant="ghost" onClick={onClose}>收起</Button></div>
      <div className="space-y-5 p-4 text-xs">
        <p className="leading-5 text-[var(--mo-text-muted)]">一行数据生成一份报告。先准备模板，在正文或表格中插入字段，再挑选要打印的记录。</p>
        <section className="space-y-3"><h3 className="font-semibold">1. 选择数据</h3>
          <Button className="w-full" disabled={loading} onClick={() => fileRef.current?.click()}>{loading ? '正在读取…' : '导入 Excel / CSV'}</Button>
          <input ref={fileRef} type="file" aria-label="导入合并数据" className="hidden" accept=".xlsx,.xls,.xlsm,.xltx,.csv" onChange={(event) => { const file = event.target.files?.[0]; if (file) void importFile(file); event.target.value = ''; }} />
          {workbooks.length ? <div className="flex gap-2"><Select aria-label="已打开的数据表" value={openWorkbook} onChange={(event) => setOpenWorkbook(event.target.value)}><option value="">选择已打开的 Excel</option>{workbooks.map((item) => <option key={item.id} value={item.id}>{item.displayName}</option>)}</Select><Button size="sm" disabled={!openWorkbook} onClick={() => { const value = sessions[openWorkbook]; if (value) chooseSource(value.data as WorkbookModel, value.displayName); }}>使用</Button></div> : null}
          {source ? <><p className="break-all text-[var(--mo-text-muted)]">数据快照：{source.name}</p><Field label="工作表"><Select aria-label="合并工作表" value={sheetId} onChange={(event) => setSheetId(event.target.value)}>{source.workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}</Select></Field><Field label="表头所在行"><Input aria-label="表头所在行" type="number" min={1} value={Number.isNaN(headerRow) ? '' : headerRow} onChange={(event) => setHeaderRow(event.target.value === '' ? NaN : Number(event.target.value))} /></Field>{data ? <p>{data.records.length} 条记录 · {data.fields.length} 个字段</p> : null}</> : <p className="leading-5 text-[var(--mo-text-muted)]">支持 XLS、XLSX 和 CSV，首行通常是字段名称。编号建议使用文本，日期与金额沿用 Excel 格式。</p>}
          {result.error ? <p role="alert" className="text-[var(--mo-danger)]">{result.error}</p> : null}
        </section>
        <section className="space-y-3 border-t border-[var(--mo-line)] pt-4"><h3 className="font-semibold">2. 在模板中插入字段</h3><p className="leading-5 text-[var(--mo-text-muted)]">将光标放在正文或表格单元格中，点击字段插入。生成时会替换为每行的真实数据。</p><div className="flex flex-wrap gap-2">{data?.fields.map((field) => <Button key={field} size="sm" onMouseDown={(event) => event.preventDefault()} onClick={() => onInsertField(field)}>{field}</Button>)}</div>{!data ? <p className="text-[var(--mo-text-muted)]">导入数据后显示可用字段。</p> : null}<p>模板已使用：{templateFields.length ? templateFields.join('、') : '尚未插入字段'}</p></section>
        <section className="space-y-3 border-t border-[var(--mo-line)] pt-4"><h3 className="font-semibold">3. 筛选、预览与批量输出</h3><Button variant="primary" className="w-full" disabled={!data || !templateFields.length || multipleSections} onClick={() => { setError(''); setStatus(''); setProgress(0); setDialogOpen(true); }}>选择记录并预览报告</Button><p className="leading-5 text-[var(--mo-text-muted)]">可按条件筛选、逐条勾选，预览真实分页后打印。每批最多 {MAX_MERGE_BATCH} 份。</p></section>
        {multipleSections ? <p role="alert" className="text-[var(--mo-danger)]">当前文档含多个节。请使用单份报告模板进行合并；批量成品可直接打开查看和打印。</p> : null}
        {error && !dialogOpen ? <p role="alert" className="text-[var(--mo-danger)]">{error}</p> : null}
      </div>
    </aside>
    <Dialog open={dialogOpen} onClose={closeDialog} closeOnBackdrop={!busy} width="xl" className="!max-w-[calc(100vw-32px)]" title="选择记录，批量生成报告" description="只有当前筛选结果中勾选的记录会输出，每份报告独立分页、重新编号。" footer={<><Button disabled={busy} onClick={closeDialog}>返回编辑模板</Button>{busy ? <Button onClick={() => { cancel.current = true; setStatus('正在取消，请稍候…'); }}>取消生成</Button> : <Button variant="primary" disabled={!canGenerate} onClick={() => void generate()}>{output === 'print' ? `打印选中 ${chosen.length} 份报告` : `生成选中 ${chosen.length} 份报告`}</Button>}</>}>
      <div className="grid gap-5 lg:grid-cols-[minmax(0,1fr)_minmax(330px,0.8fr)]">
        <div className="space-y-3">
          <fieldset disabled={busy} className="space-y-3">
            <div className="flex gap-2"><Input aria-label="搜索合并记录" placeholder="搜索所有字段…" value={search} onChange={(event) => setSearch(event.target.value)} /><Button size="sm" disabled={!data?.fields.length || conditions.length >= 5} onClick={() => setConditions((previous) => [...previous, { field: data!.fields[0], op: 'contains', value: '' }])}>添加条件</Button></div>
            {conditions.map((condition, index) => <div key={index} className="grid grid-cols-[1fr_100px_1fr_50px] gap-2"><Select aria-label={`条件 ${index + 1} 字段`} value={condition.field} onChange={(event) => setConditions((previous) => previous.map((item, i) => i === index ? { ...item, field: event.target.value } : item))}>{data?.fields.map((field) => <option key={field}>{field}</option>)}</Select><Select aria-label={`条件 ${index + 1} 比较`} value={condition.op} onChange={(event) => setConditions((previous) => previous.map((item, i) => i === index ? { ...item, op: event.target.value as MergeCondition['op'] } : item))}>{OPERATORS.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</Select><Input aria-label={`条件 ${index + 1} 值`} value={condition.value} disabled={condition.op === 'empty' || condition.op === 'notEmpty'} onChange={(event) => setConditions((previous) => previous.map((item, i) => i === index ? { ...item, value: event.target.value } : item))} /><Button size="sm" onClick={() => setConditions((previous) => previous.filter((_, i) => i !== index))}>删除</Button></div>)}
            {conditions.length ? <p className="text-xs text-[var(--mo-text-muted)]">所有条件同时满足；大于/小于使用原始数值比较。</p> : null}
            <div className="flex items-center gap-2 text-xs"><Select aria-label="记录排序字段" value={sort} onChange={(event) => setSort(event.target.value)}><option value="">按源数据顺序</option>{data?.fields.map((field) => <option key={field}>{field}</option>)}</Select><Button size="sm" disabled={!sort} onClick={() => setDescending((value) => !value)}>{descending ? '降序' : '升序'}</Button><Button size="sm" onClick={() => setSelected(new Set(filtered.map((record) => record.rowNumber)))}>全选筛选结果</Button><Button size="sm" onClick={() => setSelected(new Set())}>清空选择</Button>{filtered.length > MAX_MERGE_BATCH ? <Button size="sm" onClick={() => setSelected(new Set(filtered.slice(0, MAX_MERGE_BATCH).map((record) => record.rowNumber)))}>选择前 {MAX_MERGE_BATCH} 条</Button> : null}</div>
            <p className="text-xs font-medium">当前选中 {chosen.length} / 筛选结果 {filtered.length} / 全部 {data?.records.length ?? 0} 条</p>
            <div className="max-h-[240px] overflow-auto rounded-md border border-[var(--mo-line)]"><table className="w-full whitespace-nowrap text-left text-xs"><thead className="sticky top-0 bg-[var(--mo-surface-2)]"><tr><th className="p-2">选择</th><th className="p-2">源行</th>{data?.fields.map((field) => <th key={field} className="p-2">{field}</th>)}</tr></thead><tbody>{filtered.slice(tablePage * 25, tablePage * 25 + 25).map((record) => <tr key={record.rowNumber} className="border-t border-[var(--mo-line)]"><td className="p-2"><input type="checkbox" aria-label={`选择第 ${record.rowNumber} 行`} checked={selected.has(record.rowNumber)} onChange={(event) => setSelected((previous) => { const next = new Set(previous); if (event.target.checked) next.add(record.rowNumber); else next.delete(record.rowNumber); return next; })} /></td><td className="p-2"><button className="underline" aria-label={`预览第 ${record.rowNumber} 行`} disabled={!selected.has(record.rowNumber)} onClick={() => setPreviewIndex(chosen.findIndex((item) => item.rowNumber === record.rowNumber))}>{record.rowNumber}</button></td>{data?.fields.map((field) => <td key={field} title={record.values[field]} className="max-w-[200px] truncate p-2">{record.values[field] || <span className="text-[var(--mo-text-faint)]">空</span>}</td>)}</tr>)}</tbody></table>{!filtered.length ? <p className="p-4 text-xs">没有符合条件的记录，请调整筛选。</p> : null}</div>
            <div className="flex items-center justify-between text-xs"><Button size="sm" disabled={!tablePage} onClick={() => setTablePage((value) => value - 1)}>上一批记录</Button><span>第 {tablePage + 1} / {Math.max(1, Math.ceil(filtered.length / 25))} 页</span><Button size="sm" disabled={(tablePage + 1) * 25 >= filtered.length} onClick={() => setTablePage((value) => value + 1)}>下一批记录</Button></div>
            <label className="flex items-center gap-2 text-xs"><input type="checkbox" checked={allowEmpty} onChange={(event) => setAllowEmpty(event.target.checked)} />允许空字段，按空白输出（缺列和公式错误仍需修正）</label>
            {issues.length ? <div role="alert" className="rounded-md bg-[var(--mo-surface-2)] p-3 text-xs text-[var(--mo-danger)]"><p>请修正数据或取消选择有问题的记录：</p>{issues.slice(0, 5).map((issue, index) => <p key={index}>{issue}</p>)}{issues.length > 5 ? <p>另有 {issues.length - 5} 项问题。</p> : null}</div> : null}
            {chosen.length > MAX_MERGE_BATCH ? <p role="alert" className="text-xs text-[var(--mo-danger)]">本批超过 {MAX_MERGE_BATCH} 份，请缩小筛选范围或减少勾选。</p> : null}
            <Field label="输出方式"><Select aria-label="合并输出方式" value={output} onChange={(event) => setOutput(event.target.value as OutputKind)}><option value="print">直接批量打印</option><option value="pdf">合并为一个 PDF</option><option value="docx">合并为一个 Word 文档</option><option value="docx-zip">每条记录一个 DOCX，打包 ZIP</option><option value="pdf-zip">每条记录一个 PDF，打包 ZIP</option></Select></Field>
            {(output === 'docx-zip' || output === 'pdf-zip') ? <Field label="文件命名" hint="支持 {序号} 和 {字段名}。自动附加序号，避免重名覆盖。"><Input aria-label="报告文件命名" maxLength={150} value={namePattern} onChange={(event) => setNamePattern(event.target.value)} /></Field> : null}
          </fieldset>
          {busy || progress ? <div className="space-y-2">{busy ? <p role="status" className="text-xs">{status || '处理中…'}</p> : null}<ProgressBar value={progress} /></div> : null}
          {invalidName ? <p role="alert" className="text-xs text-[var(--mo-danger)]">文件命名中的字段不存在：{badNameFields.join("、")}</p> : null}
          {status && !busy ? <p role="status" className="text-xs leading-5">{status}</p> : null}
          {error ? <p role="alert" className="text-xs text-[var(--mo-danger)]">{error}</p> : null}
        </div>
        <div className="min-w-0 space-y-3"><div className="flex items-center justify-between text-xs"><strong>真实版式预览 · 数据第 {current?.rowNumber ?? '—'} 行</strong><span>{chosen.length ? Math.min(previewIndex + 1, chosen.length) : 0} / {chosen.length} 份</span></div><div className="flex justify-between gap-2"><Button size="sm" disabled={busy || previewIndex <= 0} onClick={() => setPreviewIndex((value) => value - 1)}>上一份报告</Button><Select aria-label="报告预览缩放" className="!w-[110px]" value={previewZoom} onChange={(event) => setPreviewZoom(event.target.value)}><option value="fit">适合宽度</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option></Select><Button size="sm" disabled={busy || previewIndex + 1 >= chosen.length} onClick={() => setPreviewIndex((value) => value + 1)}>下一份报告</Button></div><div ref={previewRef} className="max-h-[min(72vh,820px)] min-h-[520px] overflow-auto rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface-2)] p-5">{previewDoc ? <MergeReport doc={previewDoc} renderBlocks={renderBlocks} scale={previewZoom === 'fit' ? Math.max(0.2, Math.min(1, (previewWidth - 40) / doc.page.widthPx)) : Number(previewZoom)} /> : <p className="p-6 text-xs">请选择记录，并检查模板字段是否与数据列对应。</p>}</div><p className="text-[11px] leading-5 text-[var(--mo-text-muted)]">预览使用当前模板和数据快照，每份独立计算页数。长表格按行拆页；超高段落或跨页纵向合并表格会提示调整，避免打印缺失。</p></div>
      </div>
    </Dialog>
  </>;
}
