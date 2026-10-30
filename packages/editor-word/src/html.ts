import type { DocParagraph, DocRun, ParagraphStyle, RunStyle, WordTabStop } from '@mini-office/document-model';
import { DOC_FONT_STACK, fontStackFor } from '@mini-office/design-system';
import { escapeHtml } from '@mini-office/shared';

export interface RunRenderContext {
  pageNumber?: number;
  totalPages?: number;
  /** 功能：SECTIONPAGES 使用当前报告节的总页数；时间：2026-10-24；作者：lq。 */
  sectionPages?: number;
  /** 功能：段落制表位，供 tab 运行按左右/居中对齐跳转；时间：2026-10-30；作者：hou0。 */
  tabStops?: WordTabStop[];
}

/** 功能：Word 倍数行距基于字体自然行高（中文字体约 1.2×字号），CSS line-height 单位数是基于字号，需乘此系数对齐 Word 表现；时间：2026-10-30；作者：hou0。 */
const FONT_LINE_HEIGHT_RATIO = 1.2;

function runStyleCss(style: RunStyle): string {
  const css: string[] = [];
  if (style.fontFamily) {
    // 功能：中西文分开组栈——西文用主字体，中文插 eastAsia 字体（Word 的 ascii/eastAsia 双槽模型）；时间：2026-10-30；作者：hou0。
    let stack = fontStackFor(style.fontFamily).replace(/"/g, "'");
    if (style.eastAsiaFont && style.eastAsiaFont !== style.fontFamily) {
      const eastAsiaStack = fontStackFor(style.eastAsiaFont).replace(/"/g, "'");
      const firstComma = stack.indexOf(',');
      stack =
        firstComma === -1
          ? `${stack}, ${eastAsiaStack}`
          : `${stack.slice(0, firstComma)}, ${eastAsiaStack}${stack.slice(firstComma)}`;
    }
    css.push(`font-family:${stack}`);
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
  const imageId = escapeHtml(image.id ?? image.relId ?? '');
  const wrap = image.wrap ?? 'inline';

  // 功能：VML rect 绝对定位浮层（页面坐标系），按 margin-left/margin-top 定位，不占行内空间；时间：2026-10-30；作者：hou0。
  if (image.vmlRect) {
    const { left, top, width, height } = image.vmlRect;
    return `<div data-vml-rect="1" contenteditable="false" style="position:absolute;left:${left}px;top:${top}px;width:${width}px;height:${height}px;line-height:0;"><span style="font-size:${Math.min(width, height)}px;line-height:1;">□</span></div>`;
  }

  const layout = wrap === 'square-left'
    ? 'float:left;margin:0 12px 8px 0;'
    : wrap === 'square-right'
      ? 'float:right;margin:0 0 8px 12px;'
      : wrap === 'top-bottom'
        ? 'display:block;margin:8px auto;'
        : 'display:inline-block;vertical-align:baseline;';
  const frameStyle = `position:relative;line-height:0;${layout}`;
  const imageStyle = `display:block;width:${image.widthPx}px;height:${image.heightPx}px;max-width:100%;object-fit:contain;`;
  // 功能：为 Word 图片输出可直接拖动的八方向尺寸手柄；时间：2026-10-24；作者：lq。
  const handles = [
    ['nw', 'nwse-resize', 'left:-5px;top:-5px;'],
    ['n', 'ns-resize', 'left:calc(50% - 4px);top:-5px;'],
    ['ne', 'nesw-resize', 'right:-5px;top:-5px;'],
    ['e', 'ew-resize', 'right:-5px;top:calc(50% - 4px);'],
    ['se', 'nwse-resize', 'right:-5px;bottom:-5px;'],
    ['s', 'ns-resize', 'left:calc(50% - 4px);bottom:-5px;'],
    ['sw', 'nesw-resize', 'left:-5px;bottom:-5px;'],
    ['w', 'ew-resize', 'left:-5px;top:calc(50% - 4px);'],
  ].map(([name, cursor, position]) => `<span data-image-handle="${name}" contenteditable="false" aria-hidden="true" style="position:absolute;z-index:3;width:8px;height:8px;box-sizing:border-box;border:1px solid #3478f6;background:#fff;border-radius:2px;opacity:0;pointer-events:none;cursor:${cursor};${position}"></span>`).join('');
  return `<span data-image-frame data-image-id="${imageId}" contenteditable="false" style="${frameStyle}"><img src="${escapeHtml(image.src)}" alt="${alt}" data-image-id="${imageId}" contenteditable="false" draggable="false" style="${imageStyle}">${handles}</span>`;
}

export function renderRunsHtml(runs: DocRun[], context: RunRenderContext = {}): string {
  const parts: string[] = [];

  // 缓冲区：合并连续的文本 run（逐字母模式 + 负字距），避免多栏布局下单词被拆分；
  // 时间：2026-10-30；作者：hou0。
  let bufText = '';
  let bufCss = '';

  const flushBuf = () => {
    if (!bufText) return;
    const escaped = escapeHtml(bufText).replace(/\n/g, '<br>');
    parts.push(bufCss ? `<span style="${bufCss} white-space:nowrap;">${escaped}</span>` : escaped);
    bufText = '';
    bufCss = '';
  };

  // 制表位渲染状态：currentX 为已渲染内容的估计横坐标（CSS px）；
  // pendingAlign 为待渲染段落的对齐方式（center/right 由制表位决定）；
  // pendingTargetX 为该段落应对齐的目标位置。
  let currentX = 0;
  let pendingAlign: 'center' | 'right' | null = null;
  let pendingTargetX = 0;
  let pendingParts: string[] = [];

  const charWidthPx = (run: DocRun): number => {
    const sizePt = run.style.fontSizePt ?? 10.5;
    // 估算平均字符宽度 ≈ 字号 × 0.55（中英文混合近似），用于制表位位置估算；
    // pt → px ≈ ×0.75；时间：2026-10-30；作者：hou0。
    return sizePt * 0.55 * 0.75;
  };

  const flushPending = () => {
    if (!pendingParts.length || !pendingAlign) return;
    const width = Math.max(0, pendingTargetX - currentX);
    parts.push(`<span style="display:inline-block;width:${width}px;text-align:${pendingAlign}">${pendingParts.join('')}</span>`);
    currentX = pendingTargetX;
    pendingAlign = null;
    pendingTargetX = 0;
    pendingParts = [];
  };

  const tabStops = context.tabStops ?? [];
  const usedTabStops = new Set<number>();

  const findNextTabStop = (afterX: number): WordTabStop | undefined => {
    return tabStops.find((t, i) => t.posPx > afterX && !usedTabStops.has(i));
  };

  const markTabStopUsed = (t: WordTabStop) => {
    const idx = tabStops.findIndex((s) => s.posPx === t.posPx && s.kind === t.kind);
    if (idx >= 0) usedTabStops.add(idx);
  };

  for (const run of runs) {
    if (run.image) {
      flushPending();
      flushBuf();
      parts.push(imageHtml(run));
      continue;
    }
    if (run.lineBreak) {
      flushPending();
      flushBuf();
      parts.push('<br>');
      continue;
    }
    if (run.pageBreak) {
      flushPending();
      flushBuf();
      parts.push('<span data-page-break="1" contenteditable="false" style="display:inline-block;width:0;height:0;overflow:hidden;"></span>');
      continue;
    }
    if (run.tab) {
      flushPending();
      flushBuf();
      const tabStop = findNextTabStop(currentX);
      if (!tabStop) {
        // 无可用制表位：回退为固定宽度；
        // 时间：2026-10-30；作者：hou0。
        parts.push('<span data-tab="1" contenteditable="false">\u2003\u2003</span>');
        currentX += charWidthPx(run) * 2;
        continue;
      }
      markTabStopUsed(tabStop);
      const gap = Math.max(0, tabStop.posPx - currentX);
      if (tabStop.kind === 'left' || tabStop.kind === 'decimal') {
        parts.push(`<span style="display:inline-block;width:${gap}px"></span>`);
        currentX = tabStop.posPx;
      } else {
        // center / right：将后续文本缓冲，在目标位置内按 text-align 对齐；
        pendingAlign = tabStop.kind;
        pendingTargetX = tabStop.posPx;
      }
      continue;
    }
    if (run.mergeField) {
      flushPending();
      flushBuf();
      parts.push(`<span data-merge-field="${escapeHtml(run.mergeField)}" contenteditable="false" style="${runStyleCss(run.style)}"><span style="background:#e8f1ff;border-radius:3px;color:#245694;">«${escapeHtml(run.mergeField)}»</span></span>`);
      continue;
    }
    let text = run.text ?? '';
    if (run.field === 'PAGE' && context.pageNumber !== undefined) {
      text = String(context.pageNumber);
    } else if (run.field === 'SECTIONPAGES' && (context.sectionPages ?? context.totalPages) !== undefined) {
      text = String(context.sectionPages ?? context.totalPages);
    } else if (run.field === 'NUMPAGES' && context.totalPages !== undefined) {
      text = String(context.totalPages);
    }
    if (!text) continue;

    const escaped = escapeHtml(text).replace(/\n/g, '<br>');
    const css = runStyleCss(run.style);

    // 逐字母模式：连续文本 run 且有 letterSpacing，合并到 buffer，包在 nowrap span 里；
    // 普通文本（无 letterSpacing）直接渲染，遇到 buffer 非空时先 flush。
    const isPlainText = !run.field && !run.anchor && !run.href && run.style.verticalAlign !== 'super' && run.style.verticalAlign !== 'sub';
    if (isPlainText && css) {
      if (!bufCss) bufCss = css;
      bufText += text;
      currentX += text.length * charWidthPx(run);
    } else {
      flushBuf();
      if (pendingAlign) {
        // center / right 制表位下：缓冲到 pendingParts，等 flushPending 一并对齐；
        pendingParts.push(css ? `<span style="${css}">${escaped}</span>` : escaped);
        currentX += text.length * charWidthPx(run);
      } else {
        let html = css ? `<span style="${css}">${escaped}</span>` : escaped;
        if (run.style.verticalAlign === 'super') html = `<sup>${html}</sup>`;
        else if (run.style.verticalAlign === 'sub') html = `<sub>${html}</sub>`;
        if (run.anchor) {
          html = `<a data-anchor="${escapeHtml(run.anchor)}" style="color:inherit;text-decoration:none;cursor:pointer;">${html}</a>`;
        } else if (run.href) {
          html = `<a href="${escapeHtml(run.href)}" style="color:#0563c1;text-decoration:underline;cursor:pointer;">${html}</a>`;
        }
        parts.push(html);
        currentX += text.length * charWidthPx(run);
      }
    }
  }

  flushBuf();
  flushPending();
  return parts.join('');
}

/** 功能：判断是否为 Word 自动目录段落（TOC1~TOC9 样式），作为制表位信息缺失时的兜底；时间：2026-10-30；作者：hou0。 */
function isTocParagraph(paragraph: DocParagraph): boolean {
  return !!paragraph.styleId && /^TOC\s*\d+$/i.test(paragraph.styleId);
}

/** 功能：段落含「右对齐/居中制表位 + 前导符」，应渲染为「标题 + 前导点 + 页码右对齐」；时间：2026-10-30；作者：hou0。 */
function hasLeaderTab(paragraph: DocParagraph): boolean {
  return (
    paragraph.style.tabStops?.some(
      (t) => (t.kind === 'right' || t.kind === 'center') && t.leader && t.leader !== 'none',
    ) ?? false
  );
}

/** 功能：是否走前导点布局（标准 TOC，或带前导符制表位的手动目录如样式"7"）；时间：2026-10-30；作者：hou0。 */
function useLeaderLayout(paragraph: DocParagraph): boolean {
  return hasLeaderTab(paragraph) || isTocParagraph(paragraph);
}

/**
 * 功能：前导符字符——用真实字符绘制点线，密度与 Word 一致（border 虚线过于稀疏，且打印会丢）；
 * 时间：2026-10-30；作者：hou0。
 */
const LEADER_GLYPH: Record<string, string> = {
  dot: '.',
  middleDot: '·',
  hyphen: '-',
  underline: '_',
};

/**
 * 功能：把目录类段落渲染成「标题 + 前导点 + 页码右对齐」的 flex 结构，对齐 Word/WPS 目录外观；
 * 前导符按制表位 leader 类型选择字符（dot 用点、hyphen 用连字符）；时间：2026-10-30；作者：hou0。
 */
function renderLeaderTabHtml(paragraph: DocParagraph, context: RunRenderContext): string {
  const runs = paragraph.runs;
  const tabIndex = runs.findIndex((run) => run.tab);
  if (tabIndex === -1) return renderRunsHtml(runs, context) || '<br>';
  const left = runs.slice(0, tabIndex);
  const right = runs.slice(tabIndex + 1);
  const leftHtml = renderRunsHtml(left, context) || '';
  const rightHtml = renderRunsHtml(right, context) || '';
  const tabStop = paragraph.style.tabStops?.find(
    (t) => (t.kind === 'right' || t.kind === 'center') && t.leader && t.leader !== 'none',
  );
  const leaderChar = LEADER_GLYPH[tabStop?.leader ?? 'dot'] ?? '.';
  // 功能：前导点用足够多的真实字符铺满，再由 overflow 裁剪到弹性宽度，点的疏密与 Word 接近；时间：2026-10-30；作者：hou0。
  const leaderHtml = leaderChar.repeat(240);
  return (
    `<span class="mo-toc-title" style="flex:0 0 auto;white-space:pre-wrap;">${leftHtml}</span>` +
    `<span class="mo-toc-leader" contenteditable="false" style="flex:1 1 auto;overflow:hidden;white-space:nowrap;letter-spacing:0.22em;margin:0 4px;min-width:10px;">${leaderHtml}</span>` +
    `<span class="mo-toc-page" style="flex:0 0 auto;white-space:pre-wrap;">${rightHtml}</span>`
  );
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
  const content = useLeaderLayout(paragraph)
    ? renderLeaderTabHtml(paragraph, context)
    : renderRunsHtml(paragraph.runs, { ...context, tabStops: paragraph.style.tabStops });
  return `${marker}${content || '<br>'}`;
}

export function paragraphWrapperStyle(paragraph: DocParagraph): React.CSSProperties {
  const style = paragraph.style;
  const borderCss = (border: NonNullable<ParagraphStyle['borders']>['top']) => border && border.style !== 'none' ? `${Math.max(0.5, border.widthPx)}px ${border.style} ${border.color}` : undefined;
  return {
    paddingTop: style.spaceBeforePx ?? 0,
    paddingBottom: style.spaceAfterPx ?? 0,
    marginLeft: style.indentLeftPx ?? 0,
    marginRight: style.indentRightPx ?? 0,
    textAlign: style.alignment ?? 'left',
    backgroundColor: style.backgroundColor,
    borderTop: borderCss(style.borders?.top),
    borderRight: borderCss(style.borders?.right),
    borderBottom: borderCss(style.borders?.bottom),
    borderLeft: borderCss(style.borders?.left),
  };
}

export function paragraphTextStyle(paragraph: DocParagraph, defaultFont: string, defaultFontSizePt: number, linePitchPx?: number): React.CSSProperties {
  const style: ParagraphStyle = paragraph.style;
  // 功能：计算段落自然行高（CSS px）；时间：2026-10-30；作者：hou0。
  const naturalLineHeightPx = style.lineHeightExactPx || style.lineHeightAtLeastPx
    ? style.lineHeightExactPx ?? style.lineHeightAtLeastPx
    : (defaultFontSizePt * (style.lineHeight ?? 1.5) * FONT_LINE_HEIGHT_RATIO) / 0.75;
  // 功能：若文档定义了 w:docGrid w:linePitch，段落行高向上取整到网格行高的整数倍，解决封面空行偏矮、整体上移问题；时间：2026-10-30；作者：hou0。
  const lineHeightValue = linePitchPx && naturalLineHeightPx
    ? Math.ceil(naturalLineHeightPx / linePitchPx) * linePitchPx
    : naturalLineHeightPx;
  const base: React.CSSProperties = {
    fontFamily: fontStackFor(defaultFont) || DOC_FONT_STACK,
    fontSize: `${defaultFontSizePt}pt`,
    // 功能：固定值与最小值行距都用 px 表达，CSS 无“至少”语义，取相同高度即可保证分页一致；时间：2026-10-29；作者：lq。
    lineHeight: `${lineHeightValue}px`,
    textIndent: style.firstLineIndentPx ? `${style.firstLineIndentPx}px` : undefined,
    whiteSpace: 'pre-wrap',
    overflowWrap: 'break-word',
    wordBreak: 'break-word',
  };
  // 功能：目录类段落改用 flex，让前导点与页码右对齐；时间：2026-10-30；作者：hou0。
  if (useLeaderLayout(paragraph)) {
    return { ...base, display: 'flex', alignItems: 'baseline' };
  }
  return base;
}

export function paragraphPlainText(paragraph: DocParagraph, context: RunRenderContext = {}): string {
  return paragraph.runs
    .map((run) => {
      if (run.image) return '[图片]';
      if (run.tab) return '\t';
      if (run.lineBreak) return '\n';
      if (run.field === 'PAGE') return String(context.pageNumber ?? run.text ?? '');
      if (run.field === 'SECTIONPAGES') return String(context.sectionPages ?? context.totalPages ?? run.text ?? '');
      if (run.field === 'NUMPAGES') return String(context.totalPages ?? run.text ?? '');
      return run.text ?? '';
    })
    .join('');
}
