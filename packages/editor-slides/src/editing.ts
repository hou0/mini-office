// 功能：PPT 编辑的文字保真、对象操作、页面复制与撤销历史基础；时间：2026-10-24；作者：lq。
import type { SlideElement, SlideModel, SlideTextParagraph, SlideTextRun } from '@mini-office/document-model';
import { uid } from '@mini-office/shared';

export const paragraphsText = (paragraphs: SlideTextParagraph[]) => paragraphs.map((p) => p.runs.map((r) => r.lineBreak ? '\n' : r.text).join('')).join('\n');

// 功能：仅替换实际变更的文字区间，保留未改变文字的不同字号、颜色和粗体；时间：2026-10-24；作者：lq。
export function replaceParagraphText(previous: SlideTextParagraph[], value: string): SlideTextParagraph[] {
  const old = paragraphsText(previous);
  if (old === value) return previous;
  const styles = previous.flatMap((p, i) => [...p.runs.flatMap((r) => Array.from({ length: r.lineBreak ? 1 : r.text.length }, () => r.style)), ...(i < previous.length - 1 ? [p.runs.at(-1)?.style ?? {}] : [])]);
  let start = 0; let end = 0;
  while (start < Math.min(old.length, value.length) && old[start] === value[start]) start++;
  while (end < Math.min(old.length, value.length) - start && old[old.length - 1 - end] === value[value.length - 1 - end]) end++;
  const insertStyle = styles[Math.max(0, start - 1)] ?? previous[0]?.runs[0]?.style ?? { fontSizePt: 24, color: '#1D2939' };
  let offset = 0;
  return value.split('\n').map((line, index) => {
    const runs: SlideTextRun[] = [];
    for (let i = 0; i < line.length; i++) {
      const pos = offset + i;
      const style = pos < start ? styles[pos] ?? insertStyle : pos >= value.length - end ? styles[old.length - value.length + pos] ?? insertStyle : insertStyle;
      const last = runs.at(-1);
      if (last && JSON.stringify(last.style) === JSON.stringify(style)) last.text += line[i];
      else runs.push({ text: line[i], style: { ...style } });
    }
    offset += line.length + 1;
    if (!runs.length) runs.push({ text: '', style: { ...insertStyle } });
    return { ...(previous[index] ?? previous.at(-1) ?? {}), runs };
  });
}

export function findSlideElement(elements: SlideElement[], id: string | null): SlideElement | undefined {
  for (const element of elements) {
    if (element.id === id) return element;
    if (element.type === 'group') { const found = findSlideElement(element.children, id); if (found) return found; }
  }
}
export function updateElementTree(elements: SlideElement[], id: string, change: (element: SlideElement) => SlideElement | null): SlideElement[] {
  return elements.flatMap((element) => {
    if (element.id === id) { const next = change(element); return next ? [next] : []; }
    return [element.type === 'group' ? { ...element, children: updateElementTree(element.children, id, change) } : element];
  });
}

// 功能：复制页保留源对象身份以复用图片/主题，生成独立页面部件与备注；时间：2026-10-24；作者：lq。
export function copySlide(slide: SlideModel): SlideModel {
  const copy = structuredClone(slide);
  const remap = (element: SlideElement): SlideElement => ({ ...element, id: uid('el'), ...(element.type === 'group' ? { children: element.children.map(remap) } : {}) } as SlideElement);
  return { ...copy, id: uid('copy'), sourcePart: slide.sourcePart, sourceNotes: slide.sourceNotes, elements: copy.elements.map(remap) };
}
export function slideTitle(slide: SlideModel): string {
  const text = slide.elements.filter((el) => !el.inherited && el.type === 'text');
  const title = text.find((el) => el.type === 'text' && ['title', 'ctrTitle'].includes(el.placeholderType ?? '')) ?? text[0];
  return title?.type === 'text' ? paragraphsText(title.paragraphs).slice(0, 80) || '空白幻灯片' : slide.title ?? '空白幻灯片';
}
