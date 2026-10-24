import { useEffect, useRef, useState, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { DocBlock, DocModel } from '@mini-office/document-model';
import { paginateBlocks, pageMetrics, type PageSlice } from './pagination';

/** 功能：批量预览和输出共用编辑器的原始段落/表格渲染，保持版式一致；时间：2026-10-24；作者：lq。 */
export type MergeBlockRenderer = (blocks: DocBlock[], doc: DocModel, page: number, total: number) => ReactNode;
export interface MergeReportLayout { pages: HTMLElement[]; slices: PageSlice[] }

/** 功能：等待真实字体、图片与布局后分页，校验超高段落及页眉页脚；时间：2026-10-24；作者：lq。 */
export function MergeReport({ doc, renderBlocks, scale = 1, onReady, onError }: {
  doc: DocModel; renderBlocks: MergeBlockRenderer; scale?: number;
  onReady?: (layout: MergeReportLayout) => void; onError?: (message: string) => void;
}) {
  const measureRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<HTMLDivElement>(null);
  const callbacks = useRef({ onReady, onError });
  callbacks.current = { onReady, onError };
  const [slices, setSlices] = useState<PageSlice[]>([]);
  const [error, setError] = useState('');
  const metrics = pageMetrics(doc.page);
  useEffect(() => {
    let active = true;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    setSlices([]); setError('');
    const run = async () => {
      try {
        const root = measureRef.current;
        if (!root) return;
        await Promise.race([
          Promise.all([document.fonts.ready, ...Array.from(root.querySelectorAll('img')).map((img) => img.decode())]),
          new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error('字体或图片加载超时，请检查模板资源。')), 15000); }),
        ]);
        clearTimeout(timeout);
        await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
        if (!active) return;
        const blocks: DocBlock[] = [];
        const heights = new Map<string, number>();
        let usedOnPage = 0;
        for (const block of doc.blocks) {
          if (block.type === 'pageBreak') { blocks.push(block); usedOnPage = 0; continue; }
          if (block.type === 'paragraph' && block.style.pageBreakBefore) usedOnPage = 0;
          const element = Array.from(root.querySelectorAll<HTMLElement>('[data-merge-measure]')).find((item) => item.dataset.mergeMeasure === block.id)!;
          const height = element.getBoundingClientRect().height;
          if (height <= metrics.contentHeight + 0.5) { if (usedOnPage + height > metrics.contentHeight + 0.5) usedOnPage = 0; blocks.push(block); heights.set(block.id, height); usedOnPage += height; continue; }
          if (block.type !== 'table') throw new Error('有段落超过整页高度，请拆成多个段落或减小字号后再生成。');
          if (block.rows.some((row) => row.cells.some((cell) => cell.rowSpan > 1))) throw new Error('超高表格含纵向合并单元格，请拆分表格或调整页面后生成。');
          // 功能：长表格按实际行高拆页，重复连续表头，避免整表溢出丢行；时间：2026-10-24；作者：lq。
          const grid = element.querySelector<HTMLElement>('[data-table-id]')!;
          const rowHeights = getComputedStyle(grid).gridTemplateRows.split(' ').map(parseFloat);
          let headerCount = 0;
          while (block.rows[headerCount]?.isHeader) headerCount++;
          const headers = block.rows.slice(0, headerCount);
          const headerHeight = rowHeights.slice(0, headerCount).reduce((sum, value) => sum + value, 0);
          if (headerCount === block.rows.length || headerHeight >= metrics.contentHeight) throw new Error('表头超过整页高度，请减少重复标题行。');
          let budget = metrics.contentHeight - usedOnPage;
          if (headerHeight + rowHeights[headerCount] > budget) { blocks.push({ type: 'pageBreak', id: `${block.id}-before-merge` }); budget = metrics.contentHeight; usedOnPage = 0; }
          let rows = [...headers], used = headerHeight, part = 0;
          const flush = () => { const id = `${block.id}-merge-part-${part++}`; blocks.push({ ...block, id, rows }); heights.set(id, used); usedOnPage = part === 1 ? usedOnPage + used : used; rows = [...headers]; used = headerHeight; budget = metrics.contentHeight; };
          for (let i = headerCount; i < block.rows.length; i++) {
            const rowHeight = rowHeights[i];
            if (!Number.isFinite(rowHeight) || headerHeight + rowHeight > metrics.contentHeight) throw new Error('表格单行或表头超过整页高度，请调整模板。');
            if (used + rowHeight > budget && rows.length > headers.length) flush();
            rows.push(block.rows[i]); used += rowHeight;
          }
          if (rows.length > headers.length) flush();
        }
        for (const kind of ['header', 'footer'] as const) {
          const region = root.querySelector<HTMLElement>(`[data-merge-${kind}]`);
          const edge = kind === 'header' ? 'top' : 'bottom';
          if (region && doc[kind] && region.getBoundingClientRect().height > doc.page.margins[edge] - doc.page.margins[kind] + 0.5) throw new Error(`${kind === 'header' ? '页眉' : '页脚'}内容超过预留空间，请增大正文页边距。`);
        }
        setSlices(paginateBlocks(blocks, heights, metrics.contentHeight));
      } catch (cause) {
        if (!active) return;
        const message = cause instanceof Error ? cause.message : String(cause);
        setError(message); callbacks.current.onError?.(message);
      }
    };
    void run();
    return () => { active = false; clearTimeout(timeout); };
  }, [doc, renderBlocks, metrics.contentHeight]);
  useEffect(() => {
    if (slices.length && pagesRef.current) callbacks.current.onReady?.({ pages: Array.from(pagesRef.current.querySelectorAll<HTMLElement>('[data-merge-page]')), slices });
  }, [slices]);
  const margins = doc.page.margins;
  return <div>
    <div ref={measureRef} className="mo-merge-measure" aria-hidden style={{ position: 'fixed', left: -100000, top: 0, width: metrics.contentWidth, background: 'white', color: 'black' }}>
      {doc.blocks.map((block) => <div key={block.id} data-merge-measure={block.id}>{renderBlocks([block], doc, 1, 1)}</div>)}
      {doc.header ? <div data-merge-header>{renderBlocks(doc.header.blocks, doc, 1, 1)}</div> : null}
      {doc.footer ? <div data-merge-footer>{renderBlocks(doc.footer.blocks, doc, 1, 1)}</div> : null}
    </div>
    {error ? <p role="alert" className="p-4 text-sm text-[var(--mo-danger)]">{error}</p> : !slices.length ? <p className="p-4 text-sm">正在计算报告分页…</p> : null}
    <div ref={pagesRef} className="flex flex-col items-center gap-4">{slices.map((page, index) => <div key={index} style={{ width: doc.page.widthPx * scale, height: doc.page.heightPx * scale }}>
      <div data-merge-page={index + 1} style={{ position: 'relative', width: doc.page.widthPx, height: doc.page.heightPx, transform: `scale(${scale})`, transformOrigin: 'top left', overflow: 'hidden', background: 'white', color: 'black' }}>
        {doc.header ? <div style={{ position: 'absolute', left: margins.left, right: margins.right, top: margins.header }}>{renderBlocks(doc.header.blocks, doc, index + (doc.page.pageNumberStart ?? 1), slices.length)}</div> : null}
        <div style={{ position: 'absolute', left: margins.left, top: margins.top, width: metrics.contentWidth }}>{renderBlocks(page.blocks, doc, index + (doc.page.pageNumberStart ?? 1), slices.length)}</div>
        {doc.footer ? <div style={{ position: 'absolute', left: margins.left, right: margins.right, bottom: margins.footer }}>{renderBlocks(doc.footer.blocks, doc, index + (doc.page.pageNumberStart ?? 1), slices.length)}</div> : null}
      </div>
    </div>)}</div>
  </div>;
}

/** 功能：为 PDF、批量打印和分页校验挂载同一份真实页面，调用方负责清理；时间：2026-10-24；作者：lq。 */
export function mountMergeReport(doc: DocModel, renderBlocks: MergeBlockRenderer, parent: HTMLElement): { ready: Promise<MergeReportLayout>; dispose: () => void } {
  const host = document.createElement('div'); parent.appendChild(host);
  const root = createRoot(host);
  const ready = new Promise<MergeReportLayout>((resolve, reject) => root.render(<MergeReport doc={doc} renderBlocks={renderBlocks} onReady={resolve} onError={(message) => reject(new Error(message))} />));
  return { ready, dispose: () => { root.unmount(); host.remove(); } };
}
