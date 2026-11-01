/**
 * 功能：目录（TOC）条目采集与页码计算，渲染、保存与跳转共用同一套判定；时间：2026-10-30；作者：lq。
 *
 * 设计要点：
 * 1. 标题判定与 Word 一致，只看段落的大纲级别（w:outlineLvl）与内置标题样式 ID，
 *    不依赖用户自定义样式的显示名，避免改名后目录突然变空。
 * 2. 条目文本取段落的纯文本（制表符转空格），图片与域占位不计入，保证目录行不会被撑开。
 * 3. 页码由分页结果反查，不写进模型；这样增删内容后目录页码自动跟随，不需要「更新域」。
 */
import type { DocBlock, DocModel, DocParagraph } from '@mini-office/document-model';
import type { PageSlice } from './pagination';

export interface TocEntry {
  /** 从 1 开始的标题级别。 */
  level: number;
  text: string;
  /** 页码（已含节的起始页码偏移）。 */
  page: number;
  paragraphId: string;
}

export interface TocHeading {
  level: number;
  text: string;
  paragraphId: string;
}

/** 功能：段落作为标题时的级别，非标题返回 0；时间：2026-10-30；作者：lq。 */
export function headingLevelOf(paragraph: DocParagraph): number {
  const outline = paragraph.style.outlineLevel;
  if (outline !== undefined && outline >= 0 && outline < 9) return outline + 1;
  const styleId = paragraph.styleId ?? '';
  const heading = /^heading([1-9])$/i.exec(styleId);
  if (heading) return Number(heading[1]);
  const named = /^标题\s*([1-9])$/.exec(styleId);
  if (named) return Number(named[1]);
  return 0;
}

/** 功能：目录条目显示文本，制表符与连续空白归一为单个空格；时间：2026-10-30；作者：lq。 */
export function paragraphOutlineText(paragraph: DocParagraph): string {
  const text = paragraph.runs
    .map((run) => (run.image || run.mergeField ? '' : run.tab ? ' ' : run.text ?? ''))
    .join('');
  return text.replace(/[\s\u00a0\u3000]+/g, ' ').trim();
}

/** 功能：按正文顺序采集 1..maxLevel 级标题，忽略目录段落自身与表格内的标题；时间：2026-10-30；作者：lq。 */
export function collectTocHeadings(blocks: DocBlock[], maxLevel = 3): TocHeading[] {
  const headings: TocHeading[] = [];
  for (const block of blocks) {
    if (block.type !== 'paragraph' || block.toc) continue;
    const level = headingLevelOf(block);
    if (level < 1 || level > maxLevel) continue;
    const text = paragraphOutlineText(block);
    if (!text) continue;
    headings.push({ level, text, paragraphId: block.id });
  }
  return headings;
}

/** 功能：为标题补齐页码，供目录渲染与保存使用；页码缺失时回落到 1；时间：2026-10-30；作者：lq。 */
export function buildTocEntries(doc: DocModel, pages: PageSlice[], maxLevel = 3): TocEntry[] {
  const pageByParagraph = new Map<string, number>();
  for (const page of pages) {
    const base = (page.sectionPage ?? page.index) + (page.section?.page.pageNumberStart ?? doc.page.pageNumberStart ?? 1);
    for (const block of page.blocks) {
      if (block.type === 'paragraph' && !pageByParagraph.has(block.id)) {
        pageByParagraph.set(block.id, base);
      }
    }
  }
  return collectTocHeadings(doc.blocks, maxLevel).map((heading) => ({
    ...heading,
    page: pageByParagraph.get(heading.paragraphId) ?? 1,
  }));
}
