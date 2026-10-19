import type { DocImage, DocParagraph, DocRun, RunStyle } from '@mini-office/document-model';

const NAMED_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#008000',
  blue: '#0000FF',
  yellow: '#FFFF00',
  gray: '#808080',
  grey: '#808080',
  orange: '#FFA500',
  purple: '#800080',
  transparent: '',
};

export function cssColorToHex(value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  const color = value.trim().toLowerCase();
  if (!color || color === 'transparent' || color === 'rgba(0, 0, 0, 0)') return undefined;
  if (color.startsWith('#')) {
    if (/^#[0-9a-f]{3}$/.test(color)) {
      return `#${color
        .slice(1)
        .split('')
        .map((char) => char + char)
        .join('')
        .toUpperCase()}`;
    }
    if (/^#[0-9a-f]{6}$/.test(color)) return color.toUpperCase();
    return undefined;
  }
  const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(color);
  if (rgb) {
    return `#${[rgb[1], rgb[2], rgb[3]]
      .map((part) => Number(part).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()}`;
  }
  return NAMED_COLORS[color] || undefined;
}

export function parseFontSizePt(value: string | undefined | null): number | undefined {
  if (!value) return undefined;
  const match = /^([\d.]+)\s*(pt|px|em|rem)?$/i.exec(value.trim());
  if (!match) return undefined;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return undefined;
  const unit = (match[2] ?? 'px').toLowerCase();
  if (unit === 'pt') return number;
  if (unit === 'px') return number * 0.75;
  if (unit === 'em' || unit === 'rem') return number * 12;
  return number;
}

function firstFontFamily(value: string): string | undefined {
  const first = value.split(',')[0]?.trim().replace(/^['"]|['"]$/g, '');
  if (!first) return undefined;
  if (/^(serif|sans-serif|monospace|system-ui|inherit|initial|unset)$/i.test(first)) return undefined;
  return first;
}

function styleIsEqual(a: RunStyle, b: RunStyle): boolean {
  return (
    a.fontFamily === b.fontFamily &&
    a.fontSizePt === b.fontSizePt &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.color === b.color &&
    a.highlight === b.highlight &&
    a.verticalAlign === b.verticalAlign &&
    a.letterSpacingPx === b.letterSpacingPx &&
    a.allCaps === b.allCaps &&
    a.smallCaps === b.smallCaps
  );
}

function mergeTextRuns(runs: DocRun[]): DocRun[] {
  const result: DocRun[] = [];
  for (const run of runs) {
    const previous = result.at(-1);
    if (
      previous &&
      !previous.image &&
      !previous.tab &&
      !previous.lineBreak &&
      !previous.field &&
      !run.image &&
      !run.tab &&
      !run.lineBreak &&
      !run.field &&
      styleIsEqual(previous.style, run.style)
    ) {
      previous.text = (previous.text ?? '') + (run.text ?? '');
    } else {
      result.push({ ...run });
    }
  }
  return result;
}

export function serializeDomRuns(
  root: HTMLElement,
  fallback: RunStyle,
  imagesById: Map<string, DocImage>,
): DocRun[] {
  const runs: DocRun[] = [];

  const walk = (node: Node, style: RunStyle) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? '';
      if (text) runs.push({ text, style: { ...style } });
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.getAttribute('data-list-marker') === '1') return;
    if (node.hasAttribute('data-tab')) {
      runs.push({ text: '', style: { ...style }, tab: true });
      return;
    }
    const tag = node.tagName;
    if (tag === 'BR') {
      runs.push({ text: '', style: { ...style }, lineBreak: true });
      return;
    }
    if (tag === 'IMG') {
      const id = node.getAttribute('data-image-id') ?? '';
      const stored = imagesById.get(id);
      if (stored) {
        runs.push({ text: '', style: { ...style }, image: stored });
      } else {
        const src = (node as HTMLImageElement).src;
        if (src) {
          runs.push({
            text: '',
            style: { ...style },
            image: {
              id,
              src,
              widthPx: (node as HTMLImageElement).width || 160,
              heightPx: (node as HTMLImageElement).height || 120,
            },
          });
        }
      }
      return;
    }

    const next: RunStyle = { ...style };
    if (tag === 'B' || tag === 'STRONG') next.bold = true;
    if (tag === 'I' || tag === 'EM') next.italic = true;
    if (tag === 'U') next.underline = true;
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') next.strike = true;
    if (tag === 'SUP') next.verticalAlign = 'super';
    if (tag === 'SUB') next.verticalAlign = 'sub';
    if (tag === 'FONT') {
      const element = node as HTMLFontElement;
      if (element.face) next.fontFamily = element.face;
      const color = cssColorToHex(element.color);
      if (color) next.color = color;
      const sizeMap: Record<string, number> = {
        '1': 8,
        '2': 10,
        '3': 12,
        '4': 14,
        '5': 18,
        '6': 24,
        '7': 36,
      };
      if (element.size && sizeMap[element.size]) next.fontSizePt = sizeMap[element.size];
    }

    if (node instanceof HTMLElement) {
      const inline = node.style;
      if (inline.fontWeight) {
        const weight = inline.fontWeight;
        next.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
      }
      if (inline.fontStyle === 'italic' || inline.fontStyle === 'oblique') next.italic = true;
      if (inline.textDecorationLine?.includes('underline') || inline.textDecoration?.includes('underline')) {
        next.underline = true;
      }
      if (inline.textDecorationLine?.includes('line-through') || inline.textDecoration?.includes('line-through')) {
        next.strike = true;
      }
      const color = cssColorToHex(inline.color);
      if (color) next.color = color;
      const background = cssColorToHex(inline.backgroundColor);
      if (background) next.highlight = background;
      const fontSize = parseFontSizePt(inline.fontSize);
      if (fontSize) next.fontSizePt = fontSize;
      if (inline.fontFamily) {
        const family = firstFontFamily(inline.fontFamily);
        if (family) next.fontFamily = family;
      }
      if (inline.letterSpacing) {
        const spacing = parseFloat(inline.letterSpacing);
        if (Number.isFinite(spacing)) next.letterSpacingPx = spacing;
      }
      if (inline.textTransform === 'uppercase') next.allCaps = true;
      if (inline.fontVariant?.includes('small-caps')) next.smallCaps = true;
    }

    for (const child of Array.from(node.childNodes)) walk(child, next);
  };

  for (const child of Array.from(root.childNodes)) walk(child, fallback);
  const merged = mergeTextRuns(runs);
  if (merged.length === 0) merged.push({ text: '', style: { ...fallback } });
  return merged;
}

function isAtomicElement(element: Element): boolean {
  return (
    element.tagName === 'IMG' ||
    element.tagName === 'BR' ||
    element.getAttribute('data-tab') === '1'
  );
}

/** Caret offset in atomic text units (images, tabs and breaks count as 1). */
function offsetAt(root: HTMLElement, container: Node, offset: number): number {
  const measure = document.createRange();
  measure.selectNodeContents(root);
  measure.setEnd(container, offset);
  const fragment = measure.cloneContents();
  let count = 0;
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      count += (node.textContent ?? '').length;
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.getAttribute('data-list-marker') === '1') return;
    if (isAtomicElement(node)) {
      count += 1;
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  for (const child of Array.from(fragment.childNodes)) walk(child);
  return count;
}

export function getCaretOffset(root: HTMLElement): number {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return 0;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.endContainer)) return 0;
  return offsetAt(root, range.endContainer, range.endOffset);
}

/** Capture both ends of the current selection, even if focus moves to a toolbar select. */
export function getSelectionOffsets(root: HTMLElement): { start: number; end: number } | null {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
  const start = offsetAt(root, range.startContainer, range.startOffset);
  const end = offsetAt(root, range.endContainer, range.endOffset);
  return { start: Math.min(start, end), end: Math.max(start, end) };
}

/** Place the caret at an atomic unit offset. */
export function setCaretAtOffset(root: HTMLElement, target: number): void {
  const selection = window.getSelection();
  if (!selection) return;
  let remaining = Math.max(0, target);
  let placed = false;

  const place = (node: Node, offset: number) => {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    placed = true;
  };

  const walk = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = (node.textContent ?? '').length;
      if (remaining <= length) {
        place(node, remaining);
        return true;
      }
      remaining -= length;
      return false;
    }
    if (!(node instanceof Element)) return false;
    if (node.getAttribute('data-list-marker') === '1') return false;
    if (node.tagName === 'IMG' || node.tagName === 'BR' || node.getAttribute('data-tab') === '1') {
      if (remaining <= 0) {
        place(node.parentNode ?? root, Array.from((node.parentNode ?? root).childNodes).indexOf(node));
        return true;
      }
      remaining -= 1;
      return false;
    }
    for (const child of Array.from(node.childNodes)) {
      if (walk(child)) return true;
    }
    return false;
  };

  for (const child of Array.from(root.childNodes)) {
    if (walk(child)) break;
  }
  if (!placed) {
    const range = document.createRange();
    range.selectNodeContents(root);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

export function getParagraphPlainTextFromDom(root: HTMLElement): string {
  let text = '';
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      text += node.textContent ?? '';
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.getAttribute('data-list-marker') === '1') return;
    if (node.tagName === 'BR') {
      text += '\n';
      return;
    }
    if (node.tagName === 'IMG') {
      text += '￼';
      return;
    }
    if (node.getAttribute('data-tab') === '1') {
      text += '\t';
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  for (const child of Array.from(root.childNodes)) walk(child);
  return text;
}

export function runsPlainText(runs: DocRun[]): string {
  return runs
    .map((run) => {
      if (run.image) return '￼';
      if (run.tab) return '\t';
      if (run.lineBreak) return '\n';
      return run.text ?? '';
    })
    .join('');
}

export interface FindHit {
  paragraphId: string;
  index: number;
  length: number;
  pageIndex?: number;
}

export function findInParagraph(
  paragraph: DocParagraph,
  query: string,
  caseSensitive = false,
): Array<{ index: number; length: number }> {
  if (!query) return [];
  const text = runsPlainText(paragraph.runs);
  const hits: Array<{ index: number; length: number }> = [];
  if (caseSensitive) {
    let from = 0;
    while (from <= text.length) {
      const index = text.indexOf(query, from);
      if (index < 0) break;
      hits.push({ index, length: query.length });
      from = index + Math.max(1, query.length);
    }
  } else {
    const haystack = text.toLowerCase();
    const needle = query.toLowerCase();
    let from = 0;
    while (from <= haystack.length) {
      const index = haystack.indexOf(needle, from);
      if (index < 0) break;
      hits.push({ index, length: query.length });
      from = index + Math.max(1, query.length);
    }
  }
  return hits;
}

export function replaceInRuns(
  runs: DocRun[],
  query: string,
  replacement: string,
  limit: number,
  caseSensitive = false,
): { runs: DocRun[]; count: number } {
  if (!query) return { runs, count: 0 };
  let remaining = limit;
  let count = 0;
  const next = runs.map((run) => {
    if (remaining <= 0 || !run.text || run.image || run.tab) return run;
    let text = run.text;
    let changed = true;
    while (changed && remaining > 0) {
      changed = false;
      const haystack = caseSensitive ? text : text.toLowerCase();
      const needle = caseSensitive ? query : query.toLowerCase();
      const index = haystack.indexOf(needle);
      if (index >= 0) {
        text = text.slice(0, index) + replacement + text.slice(index + query.length);
        remaining -= 1;
        count += 1;
        changed = true;
      }
    }
    return text === run.text ? run : { ...run, text };
  });
  return { runs: next, count };
}

export function stripRunsForSearch(runs: DocRun[]): string {
  return runsPlainText(runs);
}

interface DomPosition {
  node: Node;
  offset: number;
}

function locateOffset(root: HTMLElement, target: number, bias: 'start' | 'end'): DomPosition | null {
  let remaining = Math.max(0, target);
  let result: DomPosition | null = null;

  const visit = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = (node.textContent ?? '').length;
      if (remaining <= length) {
        result = { node, offset: remaining };
        return true;
      }
      remaining -= length;
      return false;
    }
    if (!(node instanceof Element)) return false;
    if (node.getAttribute('data-list-marker') === '1') return false;
    if (isAtomicElement(node)) {
      if (remaining <= 0) {
        const parent = node.parentNode ?? root;
        const index = Array.from(parent.childNodes).indexOf(node);
        result = { node: parent, offset: bias === 'end' ? index + 1 : index };
        return true;
      }
      remaining -= 1;
      return false;
    }
    for (const child of Array.from(node.childNodes)) {
      if (visit(child)) return true;
    }
    return false;
  };

  for (const child of Array.from(root.childNodes)) {
    if (visit(child)) break;
  }
  if (!result) {
    result = { node: root, offset: root.childNodes.length };
  }
  return result;
}

export function selectRangeByOffset(
  root: HTMLElement,
  start: number,
  end: number,
  focus = true,
): void {
  const startPosition = locateOffset(root, start, 'start');
  const endPosition = locateOffset(root, end, 'end');
  if (!startPosition || !endPosition) return;
  const selection = window.getSelection();
  if (!selection) return;
  try {
    const range = document.createRange();
    range.setStart(startPosition.node, startPosition.offset);
    range.setEnd(endPosition.node, endPosition.offset);
    selection.removeAllRanges();
    selection.addRange(range);
    if (focus) root.focus({ preventScroll: true });
  } catch {
    // The DOM may have changed between lookup and selection.
  }
}


/** 功能：将跨段落 DOM 选区转换为各段原子偏移，不包含页眉页脚；时间：2026-10-18；作者：lq。 */
export function selectedParagraphSlices(root: HTMLElement, range: Range): Array<{ paragraphId: string; start: number; end: number }> {
  if (range.collapsed || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return [];
  const slices: Array<{ paragraphId: string; start: number; end: number }> = [];
  root.querySelectorAll<HTMLElement>('[data-paragraph-id]').forEach((element) => {
    if (!range.intersectsNode(element)) return;
    const start = element.contains(range.startContainer) ? offsetAt(element, range.startContainer, range.startOffset) : 0;
    const end = element.contains(range.endContainer) ? offsetAt(element, range.endContainer, range.endOffset)
      : offsetAt(element, element, element.childNodes.length);
    if (end > start || !element.textContent) slices.push({ paragraphId: element.dataset.paragraphId!, start, end });
  });
  return slices;
}
