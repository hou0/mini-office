import type { DocParagraph, DocRun, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { DOC_FONT_STACK, fontStackFor } from '@mini-office/design-system';
import { escapeHtml } from '@mini-office/shared';

export interface RunRenderContext {
  pageNumber?: number;
  totalPages?: number;
}

function runStyleCss(style: RunStyle): string {
  const css: string[] = [];
  if (style.fontFamily) {
    css.push(`font-family:${fontStackFor(style.fontFamily).replace(/"/g, "'")}`);
  }
  if (style.fontSizePt) css.push(`font-size:${style.fontSizePt}pt`);
  if (style.bold) css.push('font-weight:700');
  if (style.italic) css.push('font-style:italic');
  const decorations: string[] = [];
  if (style.underline) decorations.push('underline');
  if (style.strike) decorations.push('line-through');
  if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`);
  if (style.color) css.push(`color:${style.color}`);
  if (style.highlight) css.push(`background-color:${style.highlight}`);
  if (style.letterSpacingPx) css.push(`letter-spacing:${style.letterSpacingPx}px`);
  if (style.allCaps) css.push('text-transform:uppercase');
  if (style.smallCaps) css.push('font-variant:small-caps');
  return css.join(';');
}

function imageHtml(run: DocRun): string {
  const image = run.image;
  if (!image) return '';
  const alt = escapeHtml(image.alt ?? image.name ?? '');
  const style = `width:${image.widthPx}px;height:${image.heightPx}px;max-width:100%;object-fit:contain;vertical-align:baseline;`;
  return `<img src="${escapeHtml(image.src)}" alt="${alt}" data-image-id="${escapeHtml(
    image.id ?? image.relId ?? '',
  )}" contenteditable="false" draggable="false" style="${style}">`;
}

export function renderRunsHtml(runs: DocRun[], context: RunRenderContext = {}): string {
  const parts: string[] = [];
  for (const run of runs) {
    if (run.image) {
      parts.push(imageHtml(run));
      continue;
    }
    if (run.lineBreak) {
      parts.push('<br>');
      continue;
    }
    if (run.tab) {
      parts.push('<span data-tab="1" contenteditable="false">\u2003\u2003</span>');
      continue;
    }
    let text = run.text ?? '';
    if (run.field === 'PAGE' && context.pageNumber !== undefined) {
      text = String(context.pageNumber);
    } else if (run.field === 'NUMPAGES' && context.totalPages !== undefined) {
      text = String(context.totalPages);
    }
    if (!text) continue;
    const escaped = escapeHtml(text).replace(/\n/g, '<br>');
    const css = runStyleCss(run.style);
    let html = css ? `<span style="${css}">${escaped}</span>` : escaped;
    if (run.style.verticalAlign === 'super') html = `<sup>${html}</sup>`;
    else if (run.style.verticalAlign === 'sub') html = `<sub>${html}</sub>`;
    parts.push(html);
  }
  return parts.join('');
}

export function renderParagraphHtml(
  paragraph: DocParagraph,
  context: RunRenderContext = {},
): string {
  let marker = '';
  if (paragraph.list) {
    marker = `<span data-list-marker="1" contenteditable="false" style="display:inline-block;min-width:1.5em;white-space:nowrap;">${escapeHtml(
      paragraph.list.marker,
    )}</span>`;
  }
  const content = renderRunsHtml(paragraph.runs, context);
  return `${marker}${content || '<br>'}`;
}

export function paragraphWrapperStyle(paragraph: DocParagraph): React.CSSProperties {
  const style = paragraph.style;
  return {
    paddingTop: style.spaceBeforePx ?? 0,
    paddingBottom: style.spaceAfterPx ?? 0,
    marginLeft: style.indentLeftPx ?? 0,
    marginRight: style.indentRightPx ?? 0,
    textAlign: style.alignment ?? 'left',
    backgroundColor: style.backgroundColor,
  };
}

export function paragraphTextStyle(paragraph: DocParagraph, defaultFont: string, defaultFontSizePt: number): React.CSSProperties {
  const style: ParagraphStyle = paragraph.style;
  return {
    fontFamily: fontStackFor(defaultFont) || DOC_FONT_STACK,
    fontSize: `${defaultFontSizePt}pt`,
    lineHeight: style.lineHeightExactPx ? `${style.lineHeightExactPx}px` : style.lineHeight ?? 1.5,
    textIndent: style.firstLineIndentPx ? `${style.firstLineIndentPx}px` : undefined,
    whiteSpace: 'pre-wrap',
    overflowWrap: 'break-word',
    wordBreak: 'break-word',
  };
}

export function paragraphPlainText(paragraph: DocParagraph, context: RunRenderContext = {}): string {
  return paragraph.runs
    .map((run) => {
      if (run.image) return '[图片]';
      if (run.tab) return '\t';
      if (run.lineBreak) return '\n';
      if (run.field === 'PAGE') return String(context.pageNumber ?? run.text ?? '');
      if (run.field === 'NUMPAGES') return String(context.totalPages ?? run.text ?? '');
      return run.text ?? '';
    })
    .join('');
}
