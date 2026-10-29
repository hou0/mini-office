/**
 * 功能：Word 排版基础度量与中文字号映射，供工具栏、段落对话框与格式刷共用；时间：2026-10-29；作者：lq。
 *
 * 设计要点：
 * 1. 中文字号（初号…八号）只是磅值的别名，模型仍只存 `fontSizePt`，保存回 DOCX 不需要额外字段。
 * 2. 缩进与行距在模型里统一用 CSS px，本模块集中提供 pt / 字符 / px 的换算，避免各处各写一套系数。
 * 3. 格式刷的“复制—套用”以纯函数实现，便于单元测试，也保证一次套用只产生一条撤销记录。
 */
import type { DocRun, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { splitRunsAtOffset } from './doc-utils';

/** 中文字号到磅值的映射，取值与 Word / WPS 一致。 */
export const CN_FONT_SIZES: ReadonlyArray<{ name: string; pt: number }> = [
  { name: '初号', pt: 42 },
  { name: '小初', pt: 36 },
  { name: '一号', pt: 26 },
  { name: '小一', pt: 24 },
  { name: '二号', pt: 22 },
  { name: '小二', pt: 18 },
  { name: '三号', pt: 16 },
  { name: '小三', pt: 15 },
  { name: '四号', pt: 14 },
  { name: '小四', pt: 12 },
  { name: '五号', pt: 10.5 },
  { name: '小五', pt: 9 },
  { name: '六号', pt: 7.5 },
  { name: '小六', pt: 6.5 },
  { name: '七号', pt: 5.5 },
  { name: '八号', pt: 5 },
];

/** 数字磅值预设，与 Office 下拉列表常见项保持一致。 */
export const PT_FONT_SIZES: ReadonlyArray<number> = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];

/** 1pt = 4/3 px（CSS 96dpi 约定）。 */
export const PX_PER_PT = 4 / 3;

/** 一个汉字的宽度近似值：字号磅值 × 1.333，与既有的首行缩进换算保持一致。 */
export const PX_PER_CHAR_RATIO = 1.333;

export function ptToPx(pt: number): number {
  return pt * PX_PER_PT;
}

export function pxToPt(px: number): number {
  return px / PX_PER_PT;
}

export function charWidthPx(fontSizePt = 12): number {
  return Math.max(1, fontSizePt * PX_PER_CHAR_RATIO);
}

export function pxToChars(px: number, fontSizePt = 12): number {
  return px / charWidthPx(fontSizePt);
}

/** 功能：磅值对应的中文字号名，未命中返回 undefined；时间：2026-10-29；作者：lq。 */
export function cnFontSizeName(pt: number | undefined): string | undefined {
  if (!pt) return undefined;
  return CN_FONT_SIZES.find((item) => Math.abs(item.pt - pt) < 0.01)?.name;
}

/** 功能：字号在下拉框与工具栏中的显示文本，优先显示中文字号；时间：2026-10-29；作者：lq。 */
export function fontSizeLabel(pt: number | undefined): string {
  if (!pt) return '默认';
  const name = cnFontSizeName(pt);
  return name ? `${name} ${formatPt(pt)}` : formatPt(pt);
}

export function formatPt(pt: number): string {
  return Number.isInteger(pt) ? String(pt) : String(Number(pt.toFixed(1)));
}

/**
 * 功能：解析用户输入的字号，支持“三号”“16”“16磅”“16pt”与小数磅值；时间：2026-10-29；作者：lq。
 * 返回 undefined 表示输入无法识别，调用方应保持原值而不是写入非法字号。
 */
export function parseFontSizeInput(raw: string | undefined | null): number | undefined {
  if (!raw) return undefined;
  const text = raw.trim().replace(/\s+/g, '');
  if (!text) return undefined;
  const named = CN_FONT_SIZES.find((item) => item.name === text);
  if (named) return named.pt;
  const numeric = text.replace(/(磅|点|pt|px)$/i, '');
  if (!/^\d+(\.\d+)?$/.test(numeric)) return undefined;
  const value = Number(numeric);
  if (!Number.isFinite(value) || value <= 0) return undefined;
  // 与 Word 一致：字号下限 1 磅、上限 1638 磅。
  return Math.min(1638, Math.max(1, Math.round(value * 10) / 10));
}

export type LineSpacingMode = 'multiple' | 'exact' | 'atLeast';

export interface LineSpacingState {
  mode: LineSpacingMode;
  /** multiple 时为倍数，exact / atLeast 时为磅值。 */
  value: number;
}

/** 功能：从段落样式还原行距状态，兼容历史数据里只有倍数或只有固定值的情况；时间：2026-10-29；作者：lq。 */
export function readLineSpacing(style: ParagraphStyle | undefined, fontSizePt = 12): LineSpacingState {
  if (style?.lineHeightExactPx) {
    return { mode: 'exact', value: Math.round(pxToPt(style.lineHeightExactPx) * 10) / 10 };
  }
  if (style?.lineHeightAtLeastPx) {
    return { mode: 'atLeast', value: Math.round(pxToPt(style.lineHeightAtLeastPx) * 10) / 10 };
  }
  const multiple = style?.lineHeight ?? 1.5;
  void fontSizePt;
  return { mode: 'multiple', value: Math.round(multiple * 100) / 100 };
}

/** 功能：把行距状态写回段落样式补丁，切换模式时清理另外两种字段；时间：2026-10-29；作者：lq。 */
export function writeLineSpacing(state: LineSpacingState): Partial<ParagraphStyle> {
  const value = Math.max(0.5, Math.min(200, state.value || 0));
  if (state.mode === 'exact') {
    return { lineHeight: undefined, lineHeightAtLeastPx: undefined, lineHeightExactPx: ptToPx(value) };
  }
  if (state.mode === 'atLeast') {
    return { lineHeight: undefined, lineHeightExactPx: undefined, lineHeightAtLeastPx: ptToPx(value) };
  }
  return { lineHeightExactPx: undefined, lineHeightAtLeastPx: undefined, lineHeight: Math.max(0.5, value) };
}

/** 功能：行距在工具栏上的紧凑显示，固定值与最小值带单位便于区分；时间：2026-10-29；作者：lq。 */
export function lineSpacingLabel(state: LineSpacingState): string {
  if (state.mode === 'exact') return `固定 ${formatPt(state.value)} 磅`;
  if (state.mode === 'atLeast') return `最小 ${formatPt(state.value)} 磅`;
  return `${formatPt(state.value)} 倍`;
}

/** 功能：段落特殊缩进（首行缩进 / 悬挂缩进）在模型中以 firstLineIndentPx 正负表示；时间：2026-10-29；作者：lq。 */
export type SpecialIndentMode = 'none' | 'firstLine' | 'hanging';

export function readSpecialIndent(style: ParagraphStyle | undefined, fontSizePt = 12): { mode: SpecialIndentMode; chars: number } {
  const px = style?.firstLineIndentPx ?? 0;
  if (!px) return { mode: 'none', chars: 0 };
  const chars = Math.round(pxToChars(Math.abs(px), fontSizePt) * 10) / 10;
  return { mode: px > 0 ? 'firstLine' : 'hanging', chars };
}

export function writeSpecialIndent(mode: SpecialIndentMode, chars: number, fontSizePt = 12): Partial<ParagraphStyle> {
  if (mode === 'none') return { firstLineIndentPx: 0 };
  const px = Math.round(charWidthPx(fontSizePt) * Math.max(0, chars) * 100) / 100;
  return { firstLineIndentPx: mode === 'hanging' ? -px : px };
}

/** 格式刷复制的段落属性集合：只包含排版相关字段，不携带边框、底纹等内容型格式。 */
export const PAINTED_PARAGRAPH_KEYS = [
  'alignment',
  'indentLeftPx',
  'indentRightPx',
  'firstLineIndentPx',
  'spaceBeforePx',
  'spaceAfterPx',
  'lineHeight',
  'lineHeightExactPx',
  'lineHeightAtLeastPx',
  'keepNext',
  'pageBreakBefore',
] as const satisfies ReadonlyArray<keyof ParagraphStyle>;

/** 功能：提取可被格式刷复制的段落样式；时间：2026-10-29；作者：lq。 */
export function pickParagraphStyle(style: ParagraphStyle | undefined): ParagraphStyle {
  const picked: ParagraphStyle = {};
  if (!style) return picked;
  for (const key of PAINTED_PARAGRAPH_KEYS) {
    if (style[key] !== undefined) {
      (picked as Record<string, unknown>)[key] = style[key];
    }
  }
  return picked;
}

/** 功能：把复制的段落样式覆盖到目标段落，未复制的字段保留目标原值；时间：2026-10-29；作者：lq。 */
export function mergeParagraphStyle(target: ParagraphStyle, painted: ParagraphStyle): ParagraphStyle {
  const next: ParagraphStyle = { ...target };
  for (const key of PAINTED_PARAGRAPH_KEYS) {
    (next as Record<string, unknown>)[key] = painted[key];
  }
  return next;
}

function mergeRunStyle(base: RunStyle, painted: RunStyle): RunStyle {
  // 用复制到的样式整体覆盖字体相关字段，未复制的字段（如上下标）保持目标原值。
  return {
    ...base,
    fontFamily: painted.fontFamily,
    fontSizePt: painted.fontSizePt,
    bold: painted.bold,
    italic: painted.italic,
    underline: painted.underline,
    strike: painted.strike,
    color: painted.color,
    highlight: painted.highlight,
  };
}

/**
 * 功能：对段落内 [start, end) 区间的文字套用字体格式，返回新的 runs 数组；时间：2026-10-29；作者：lq。
 * 说明：区间按原子偏移（图片、制表符、域各计 1）切分，因此跨图片选区不会破坏对象本身。
 */
export function applyRunStyleToRuns(
  runs: DocRun[],
  start: number,
  end: number,
  painted: RunStyle,
  fallback: RunStyle = {},
): DocRun[] {
  if (end <= start) return runs;
  const [head, tail] = splitRunsAtOffset(runs, end, fallback);
  const [before, middle] = splitRunsAtOffset(head, start, fallback);
  const styled = middle.map((run) => ({ ...run, style: mergeRunStyle(run.style, painted) }));
  const result = [...before, ...styled, ...tail];
  // 去掉切分产生的空 run，避免保存时出现多余的空 span。
  return result.filter((run, index) => run.text !== '' || run.image || run.tab || run.lineBreak || run.field || run.mergeField || index === result.length - 1);
}
