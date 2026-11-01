import type { DocBlock, DocModel, DocSection, DocTable, DocTableRow } from '@mini-office/document-model';
import { tableCellPlacements } from './doc-utils';
import { documentSections } from './sections';

export interface PageSlice {
  index: number;
  blocks: DocBlock[];
  /**
   * 功能：本页中的表格跨页片段 id；这些 block 是切片而不是原表，渲染时不能被原表映射覆盖；
   * 时间：2026-10-29；作者：lq。
   */
  fragments?: string[];
  /**
   * 功能：续页片段 id；同一张表只在首片就地编辑，续页只读以免编辑落到错误的行上；
   * 时间：2026-10-29；作者：lq。
   */
  locked?: string[];
  /** 功能：批量报告的节内页序与节总页数；时间：2026-10-24；作者：lq。 */
  section?: DocSection;
  sectionPage?: number;
  sectionTotal?: number;
}

/** 功能：表格实测布局，用于按行切页并重复表头；时间：2026-10-29；作者：lq。 */
export interface TableLayout {
  /** 每行渲染高度（px），长度应与 table.rows 一致。 */
  rowHeights: number[];
  /** 连续表头行数，跨页时重复。 */
  headerRows: number;
}

/** 功能：每节独立分页并恢复编号，合并报告重新打开后仍保持各自页码；时间：2026-10-24；作者：lq。 */
export function paginateDocument(
  doc: DocModel,
  heights: Map<string, number>,
  tables?: Map<string, TableLayout>,
): PageSlice[] {
  const sections = documentSections(doc);
  const pages: PageSlice[] = [];
  let nextNumber = 1;
  for (const section of sections) {
    const slices = paginateBlocks(section.blocks, heights, pageMetrics(section.page).contentHeight, tables);
    const start = section.page.pageNumberStart ?? nextNumber;
    slices.forEach((slice, index) => pages.push({ ...slice, index: pages.length, section: { page: { ...section.page, pageNumberStart: start }, header: section.header, footer: section.footer }, sectionPage: index, sectionTotal: slices.length }));
    nextNumber = start + slices.length;
  }
  return pages;
}

/**
 * 功能：按实测高度把内容装进定高页面；长表格按行切分并重复表头，避免整表被裁掉丢行；
 * 时间：2026-10-29；作者：lq。
 */
export function paginateBlocks(
  blocks: DocBlock[],
  heights: Map<string, number>,
  contentHeight: number,
  tables?: Map<string, TableLayout>,
): PageSlice[] {
  const pages: PageSlice[] = [];
  let current: DocBlock[] = [];
  let fragments: string[] = [];
  let locked: string[] = [];
  let used = 0;

  const flush = () => {
    if (current.length === 0) return;
    pages.push({
      index: pages.length,
      blocks: current,
      fragments: fragments.length ? fragments : undefined,
      locked: locked.length ? locked : undefined,
    });
    current = [];
    fragments = [];
    locked = [];
    used = 0;
  };

  for (const block of blocks) {
    if (block.type === 'pageBreak') {
      // 连续/段首硬分页有实际空白页语义，不能与一个分页符折叠。
      if (current.length === 0) pages.push({ index: pages.length, blocks: [] });
      else flush();
      continue;
    }
    if (block.type === 'paragraph' && block.style.pageBreakBefore && current.length > 0) {
      flush();
    }
    const height = Math.max(0, heights.get(block.id) ?? 0);
    const table = block.type === 'table' ? block : undefined;
    const layout = table ? tables?.get(table.id) : undefined;
    // 功能：整表放得下时保持原样，避免无意义的切片；时间：2026-10-29；作者：lq。
    if (!table || !layout || used + height <= contentHeight + 0.5) {
      if (current.length > 0 && used + height > contentHeight + 0.5) flush();
      current.push(block);
      used += height;
      if (used > contentHeight + 0.5) flush();
      continue;
    }
    // 功能：长表格逐行装页，续页重复表头，行不会被裁掉；时间：2026-10-29；作者：lq。
    let part = 0;
    let start = 0;
    let continuation = false;
    while (start < table.rows.length) {
      const headerRows = continuation ? layout.headerRows : 0;
      const headerHeight = sumHeights(layout.rowHeights, 0, headerRows);
      let budget = contentHeight - used - headerHeight;
      if (budget <= 0) {
        if (current.length > 0) { flush(); continue; }
        // 功能：表头本身高于整页时仍允许输出，避免死循环；时间：2026-10-29；作者：lq。
        budget = Math.max(1, contentHeight - headerHeight);
      }
      let end = start;
      let acc = 0;
      while (end < table.rows.length) {
        const rowHeight = layout.rowHeights[end] ?? 0;
        if (end > start && acc + rowHeight > budget + 0.5) break;
        acc += rowHeight;
        end += 1;
      }
      // 功能：单行超过整页高度时仍独占一页，宁可轻微溢出也不丢行；时间：2026-10-29；作者：lq。
      if (end === start) end = start + 1;
      // 功能：首片沿用原表 id，保证点击表格仍能编辑原表；续页用独立 id 并锁定编辑；时间：2026-10-29；作者：lq。
      const fragmentId = part === 0 ? table.id : `${table.id}-part-${part}`;
      const fragment = sliceTable(table, start, end - 1, headerRows);
      current.push({ ...fragment, id: fragmentId });
      fragments.push(fragmentId);
      if (part > 0) locked.push(fragmentId);
      part += 1;
      used += headerHeight + sumHeights(layout.rowHeights, start, end);
      start = end;
      continuation = true;
      if (start < table.rows.length) flush();
    }
  }
  flush();
  if (pages.length === 0) pages.push({ index: 0, blocks: [] });
  return pages;
}

/** 功能：累加区间行高；时间：2026-10-29；作者：lq。 */
function sumHeights(heights: number[], start: number, end: number): number {
  let total = 0;
  for (let index = start; index < end; index += 1) total += heights[index] ?? 0;
  return total;
}

/**
 * 功能：截取表格的行区间，续页重复表头，跨页合并单元格按页裁剪并保留边框底纹；
 * 时间：2026-10-29；作者：lq。
 */
export function sliceTable(table: DocTable, start: number, end: number, repeatHeader: number): DocTable {
  const placements = tableCellPlacements(table);
  const byRow = new Map<number, Array<{ col: number; cell: DocTableRow['cells'][number] }>>();
  for (let row = start; row <= end; row += 1) byRow.set(row, []);
  for (const placement of placements) {
    const placementEnd = placement.row + placement.cell.rowSpan - 1;
    if (placement.row > end || placementEnd < start) continue;
    const rowStart = Math.max(placement.row, start);
    const rowEnd = Math.min(placementEnd, end);
    const cell = { ...placement.cell, rowSpan: rowEnd - rowStart + 1 };
    // 功能：从上一页延续下来的合并单元格只保留边框与底纹，正文留在起始页；时间：2026-10-29；作者：lq。
    if (placement.row < start) cell.blocks = [];
    byRow.get(rowStart)!.push({ col: placement.col, cell });
  }
  const rows: DocTableRow[] = [];
  for (let row = start; row <= end; row += 1) {
    const source = table.rows[row];
    const cells = (byRow.get(row) ?? []).sort((a, b) => a.col - b.col).map((item) => item.cell);
    rows.push({ ...source, cells });
  }
  const headers = repeatHeader > 0
    ? table.rows.slice(0, repeatHeader).map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell })) }))
    : [];
  return { ...table, rows: [...headers, ...rows] };
}

export interface PageMetrics {
  contentWidth: number;
  contentHeight: number;
}

export function pageMetrics(page: {
  widthPx: number;
  heightPx: number;
  margins: { top: number; right: number; bottom: number; left: number };
}): PageMetrics {
  return {
    contentWidth: Math.max(120, page.widthPx - page.margins.left - page.margins.right),
    contentHeight: Math.max(80, page.heightPx - page.margins.top - page.margins.bottom),
  };
}
