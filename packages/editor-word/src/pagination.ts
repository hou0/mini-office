import type { DocBlock } from '@mini-office/document-model';

export interface PageSlice {
  index: number;
  blocks: DocBlock[];
}

/**
 * Pack measured blocks into fixed-height pages. Block heights must include
 * their vertical spacing. Tables taller than a page are kept intact and allowed
 * to overflow (the UI reports this as a V1 limitation rather than dropping rows).
 */
export function paginateBlocks(
  blocks: DocBlock[],
  heights: Map<string, number>,
  contentHeight: number,
): PageSlice[] {
  const pages: PageSlice[] = [];
  let current: DocBlock[] = [];
  let used = 0;

  const flush = () => {
    if (current.length === 0) return;
    pages.push({ index: pages.length, blocks: current });
    current = [];
    used = 0;
  };

  for (const block of blocks) {
    if (block.type === 'pageBreak') {
      flush();
      continue;
    }
    if (block.type === 'paragraph' && block.style.pageBreakBefore && current.length > 0) {
      flush();
    }
    const height = Math.max(0, heights.get(block.id) ?? 0);
    if (current.length > 0 && used + height > contentHeight + 0.5) {
      flush();
    }
    current.push(block);
    used += height;
    if (used > contentHeight + 0.5) {
      // Oversized single block: keep it on this page and start a new page next.
      flush();
    }
  }
  flush();
  if (pages.length === 0) pages.push({ index: 0, blocks: [] });
  return pages;
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
