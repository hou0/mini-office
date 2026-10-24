import {
  type DocBlock,
  type DocImage,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocTable,
  type DocTableCell,
  type DocTableRow,
  type HeaderFooter,
  type ListInfo,
  type PageSetup,
  type ParagraphStyle,
  type RunStyle,
  type TableBorderSide,
  type TableBorders,
} from '@mini-office/document-model';
import {
  cssColor,
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  loadOoxmlPackage,
  mimeForPath,
  numericAttr,
  parseXmlDom,
  readThemeColors,
  readThemeFonts,
  relationshipMap,
  toggleValue,
  twipsToPx,
  emuToPx,
  type OoxmlPackage,
  type Relationship,
} from '@mini-office/ooxml';
import { countWords, uid } from '@mini-office/shared';
// 功能：解析 Word 原生合并域；时间：2026-10-24；作者：lq。
import { parseMergeField } from './merge-fields';

const HIGHLIGHT_COLORS: Record<string, string | undefined> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGray: '#808080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  green: '#00FF00',
  lightGray: '#C0C0C0',
  magenta: '#FF00FF',
  none: undefined,
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

interface NumberingLevel {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
  indentLeftPx?: number;
  firstLineIndentPx?: number;
}

interface ResolvedStyle {
  run: RunStyle;
  para: ParagraphStyle;
  numId?: string;
  ilvl?: number;
}

interface RawStyle {
  id: string;
  type: string;
  basedOn?: string;
  isDefault?: boolean;
  run: RunStyle;
  para: ParagraphStyle;
  numId?: string;
  ilvl?: number;
}

interface StyleSheet {
  docDefaults: { run: RunStyle; para: ParagraphStyle };
  styles: Map<string, ResolvedStyle>;
  defaultParagraph?: ResolvedStyle;
}

interface ParseContext {
  pkg: OoxmlPackage;
  rels: Map<string, Relationship>;
  styles: StyleSheet;
  numbering: Map<string, NumberingLevel[]>;
  counters: Map<string, Map<number, number>>;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  defaultFont: string;
  defaultSizePt: number;
  imageUrls: Map<string, string>;
  usedImages: Map<string, { src: string; part?: string }>;
  pageWidthPx: number;
  blockCounter: number;
  hasTrackedChanges: boolean;
  hasContentControls: boolean;
}

function mergeRun(base: RunStyle, override: RunStyle | undefined): RunStyle {
  if (!override) return { ...base };
  const result = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (result as Record<string, unknown>)[key] = value;
  }
  return result;
}

function mergeParagraph(base: ParagraphStyle, override: ParagraphStyle | undefined): ParagraphStyle {
  if (!override) return { ...base };
  const result = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (result as Record<string, unknown>)[key] = value;
  }
  return result;
}

function attrValue(element: Element | null | undefined, name: string): string | undefined {
  return getAttr(element, name);
}

function colorFromElement(element: Element | null, theme: Record<string, string>): string | undefined {
  if (!element) return undefined;
  const value = attrValue(element, 'w:val');
  if (value && value.toLowerCase() !== 'auto') {
    const hex = value.replace(/^#/, '');
    if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toUpperCase()}`;
    if (/^[0-9a-fA-F]{8}$/.test(hex)) return `#${hex.slice(0, 6).toUpperCase()}`;
  }
  const themeColor = attrValue(element, 'w:themeColor');
  if (!themeColor) return undefined;
  const schemeMap: Record<string, string> = {
    text1: 'tx1',
    background1: 'bg1',
    text2: 'tx2',
    background2: 'bg2',
    dark1: 'dk1',
    light1: 'lt1',
    dark2: 'dk2',
    light2: 'lt2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hyperlink: 'hlink',
    followedHyperlink: 'folHlink',
  };
  const base = theme[schemeMap[themeColor] ?? themeColor];
  if (!base) return undefined;
  const tintHex = attrValue(element, 'w:themeTint');
  const shadeHex = attrValue(element, 'w:themeShade');
  const toPercent = (hex?: string) => {
    if (!hex) return undefined;
    const parsed = parseInt(hex, 16);
    return Number.isFinite(parsed) ? (parsed / 255) * 100000 : undefined;
  };
  return cssColor({ hex: base, tint: toPercent(tintHex), shade: toPercent(shadeHex) }, theme);
}

function applyRunStyle(base: RunStyle, rPr: Element | null, ctx: ParseContext): RunStyle {
  if (!rPr) return { ...base };
  const style: RunStyle = { ...base };
  const fonts = firstChild(rPr, 'w:rFonts');
  const font =
    attrValue(fonts, 'w:eastAsia') ??
    attrValue(fonts, 'w:ascii') ??
    attrValue(fonts, 'w:hAnsi') ??
    attrValue(fonts, 'w:cs');
  if (font) style.fontFamily = font;

  const size = numericAttr(firstChild(rPr, 'w:sz'), 'w:val');
  const sizeCs = numericAttr(firstChild(rPr, 'w:szCs'), 'w:val');
  if (size) style.fontSizePt = size / 2;
  else if (sizeCs) style.fontSizePt = sizeCs / 2;

  const bold = firstChild(rPr, 'w:b');
  if (bold) style.bold = toggleValue(bold);
  const italic = firstChild(rPr, 'w:i');
  if (italic) style.italic = toggleValue(italic);
  const underline = firstChild(rPr, 'w:u');
  if (underline) {
    const value = attrValue(underline, 'w:val');
    style.underline = value ? !['none', '0', 'false'].includes(value.toLowerCase()) : true;
  }
  const strike = firstChild(rPr, 'w:strike');
  if (strike) style.strike = toggleValue(strike);

  const color = colorFromElement(firstChild(rPr, 'w:color'), ctx.themeColors);
  if (color) style.color = color;
  const highlight = firstChild(rPr, 'w:highlight');
  if (highlight) {
    const value = attrValue(highlight, 'w:val');
    if (value && HIGHLIGHT_COLORS[value]) style.highlight = HIGHLIGHT_COLORS[value];
    else if (value === 'none') style.highlight = undefined;
  }
  const shading = firstChild(rPr, 'w:shd');
  const fill = attrValue(shading, 'w:fill');
  if (fill && fill.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)) {
    style.highlight = `#${fill.toUpperCase()}`;
  }

  const vertAlign = attrValue(firstChild(rPr, 'w:vertAlign'), 'w:val');
  if (vertAlign === 'superscript') style.verticalAlign = 'super';
  else if (vertAlign === 'subscript') style.verticalAlign = 'sub';
  else if (vertAlign === 'baseline') style.verticalAlign = 'baseline';

  const spacing = numericAttr(firstChild(rPr, 'w:spacing'), 'w:val');
  if (spacing !== undefined) style.letterSpacingPx = twipsToPx(spacing);
  if (firstChild(rPr, 'w:caps')) style.allCaps = toggleValue(firstChild(rPr, 'w:caps'));
  if (firstChild(rPr, 'w:smallCaps')) style.smallCaps = toggleValue(firstChild(rPr, 'w:smallCaps'));
  return style;
}

function applyParagraphStyle(base: ParagraphStyle, pPr: Element | null, ctx: ParseContext): ParagraphStyle {
  if (!pPr) return { ...base };
  const style: ParagraphStyle = { ...base };
  const alignment = attrValue(firstChild(pPr, 'w:jc'), 'w:val');
  if (alignment) {
    const map: Record<string, ParagraphStyle['alignment']> = {
      left: 'left',
      start: 'left',
      right: 'right',
      end: 'right',
      center: 'center',
      both: 'justify',
      justify: 'justify',
      distribute: 'justify',
    };
    style.alignment = map[alignment] ?? base.alignment;
  }

  const indent = firstChild(pPr, 'w:ind');
  if (indent) {
    const left = numericAttr(indent, 'w:left') ?? numericAttr(indent, 'w:start');
    const right = numericAttr(indent, 'w:right') ?? numericAttr(indent, 'w:end');
    const firstLine = numericAttr(indent, 'w:firstLine');
    const hanging = numericAttr(indent, 'w:hanging');
    if (left !== undefined) style.indentLeftPx = twipsToPx(left);
    if (right !== undefined) style.indentRightPx = twipsToPx(right);
    if (firstLine !== undefined) style.firstLineIndentPx = twipsToPx(firstLine);
    else if (hanging !== undefined) style.firstLineIndentPx = -twipsToPx(hanging);
  }

  const spacing = firstChild(pPr, 'w:spacing');
  if (spacing) {
    const before = numericAttr(spacing, 'w:before');
    const after = numericAttr(spacing, 'w:after');
    const line = numericAttr(spacing, 'w:line');
    const rule = attrValue(spacing, 'w:lineRule') ?? 'auto';
    if (before !== undefined) style.spaceBeforePx = twipsToPx(before);
    if (after !== undefined) style.spaceAfterPx = twipsToPx(after);
    if (line) {
      if (rule === 'auto') style.lineHeight = line / 240;
      else style.lineHeightExactPx = twipsToPx(line);
    }
  }

  const keepNext = firstChild(pPr, 'w:keepNext');
  if (keepNext) style.keepNext = toggleValue(keepNext);
  const pageBreakBefore = firstChild(pPr, 'w:pageBreakBefore');
  if (pageBreakBefore) style.pageBreakBefore = toggleValue(pageBreakBefore);
  const outline = numericAttr(firstChild(pPr, 'w:outlineLvl'), 'w:val');
  if (outline !== undefined) style.outlineLevel = outline;

  const shading = firstChild(pPr, 'w:shd');
  const fill = attrValue(shading, 'w:fill');
  if (fill && /^[0-9a-fA-F]{6}$/.test(fill)) style.backgroundColor = `#${fill.toUpperCase()}`;
  return style;
}

function parseProjectedParagraphNumbering(pPr: Element | null): { numId?: string; ilvl?: number } {
  const numPr = firstChild(pPr, 'w:numPr');
  if (!numPr) return {};
  return {
    numId: attrValue(firstChild(numPr, 'w:numId'), 'w:val'),
    ilvl: numericAttr(firstChild(numPr, 'w:ilvl'), 'w:val'),
  };
}

function parseStyleSheet(stylesXml: string, ctx: ParseContext): StyleSheet {
  const docDefaults: { run: RunStyle; para: ParagraphStyle } = { run: {}, para: {} };
  const rawStyles = new Map<string, RawStyle>();
  let defaultParagraphId: string | undefined;

  if (stylesXml) {
    const doc = parseXmlDom(stylesXml);
    const stylesRoot = firstDescendant(doc, 'w:styles') ?? doc.documentElement;
    const defaults = firstChild(stylesRoot, 'w:docDefaults');
    if (defaults) {
      const rPrDefault = firstChild(firstChild(defaults, 'w:rPrDefault') ?? defaults, 'w:rPr');
      const pPrDefault = firstChild(firstChild(defaults, 'w:pPrDefault') ?? defaults, 'w:pPr');
      docDefaults.run = applyRunStyle({}, rPrDefault, ctx);
      docDefaults.para = applyParagraphStyle({}, pPrDefault, ctx);
    }
    for (const styleElement of Array.from(stylesRoot.childNodes)) {
      if (!(styleElement instanceof Element) || styleElement.tagName !== 'w:style') continue;
      const id = attrValue(styleElement, 'w:styleId');
      if (!id) continue;
      const type = attrValue(styleElement, 'w:type') ?? 'paragraph';
      const basedOn = attrValue(firstChild(styleElement, 'w:basedOn'), 'w:val');
      const isDefault = attrValue(styleElement, 'w:default') === '1';
      const run = applyRunStyle({}, firstChild(styleElement, 'w:rPr'), ctx);
      const para = applyParagraphStyle({}, firstChild(styleElement, 'w:pPr'), ctx);
      const num = parseProjectedParagraphNumbering(firstChild(styleElement, 'w:pPr'));
      rawStyles.set(id, { id, type, basedOn, isDefault, run, para, ...num });
      if (type === 'paragraph' && isDefault) defaultParagraphId = id;
    }
  }

  const resolved = new Map<string, ResolvedStyle>();
  const resolving = new Set<string>();

  const resolve = (id: string, stack: string[] = []): ResolvedStyle => {
    const cached = resolved.get(id);
    if (cached) return cached;
    if (resolving.has(id) || stack.length > 30) {
      return { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    }
    resolving.add(id);
    const raw = rawStyles.get(id);
    if (!raw) {
      resolving.delete(id);
      return { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    }
    const base = raw.basedOn ? resolve(raw.basedOn, [...stack, id]) : { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    const result: ResolvedStyle = {
      run: mergeRun(base.run, raw.run),
      para: mergeParagraph(base.para, raw.para),
      numId: raw.numId ?? base.numId,
      ilvl: raw.ilvl ?? base.ilvl,
    };
    resolved.set(id, result);
    resolving.delete(id);
    return result;
  };

  rawStyles.forEach((_style, id) => resolve(id));
  const defaultParagraph = defaultParagraphId
    ? resolved.get(defaultParagraphId)
    : resolved.get('Normal');
  return { docDefaults, styles: resolved, defaultParagraph };
}

function parseNumbering(numberingXml: string): Map<string, NumberingLevel[]> {
  const result = new Map<string, NumberingLevel[]>();
  if (!numberingXml) return result;
  const doc = parseXmlDom(numberingXml);
  const root = firstDescendant(doc, 'w:numbering') ?? doc.documentElement;
  const abstractLevels = new Map<string, NumberingLevel[]>();
  for (const abstractNum of Array.from(root.childNodes)) {
    if (!(abstractNum instanceof Element) || abstractNum.tagName !== 'w:abstractNum') continue;
    const id = attrValue(abstractNum, 'w:abstractNumId');
    if (!id) continue;
    const levels: NumberingLevel[] = [];
    for (const lvl of Array.from(abstractNum.childNodes)) {
      if (!(lvl instanceof Element) || lvl.tagName !== 'w:lvl') continue;
      const ilvl = numericAttr(lvl, 'w:ilvl') ?? levels.length;
      const pPr = firstChild(lvl, 'w:pPr');
      const ind = firstChild(pPr, 'w:ind');
      const left = numericAttr(ind, 'w:left') ?? numericAttr(ind, 'w:start');
      const hanging = numericAttr(ind, 'w:hanging');
      const firstLine = numericAttr(ind, 'w:firstLine');
      levels.push({
        ilvl,
        numFmt: attrValue(firstChild(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
        lvlText: attrValue(firstChild(lvl, 'w:lvlText'), 'w:val') ?? '',
        start: numericAttr(firstChild(lvl, 'w:start'), 'w:val') ?? 1,
        indentLeftPx: left !== undefined ? twipsToPx(left) : undefined,
        firstLineIndentPx:
          firstLine !== undefined
            ? twipsToPx(firstLine)
            : hanging !== undefined
              ? -twipsToPx(hanging)
              : undefined,
      });
    }
    abstractLevels.set(id, levels);
  }
  for (const num of Array.from(root.childNodes)) {
    if (!(num instanceof Element) || num.tagName !== 'w:num') continue;
    const numId = attrValue(num, 'w:numId');
    const abstractId = attrValue(firstChild(num, 'w:abstractNumId'), 'w:val');
    if (!numId || !abstractId) continue;
    result.set(numId, abstractLevels.get(abstractId) ?? []);
  }
  return result;
}

function toChineseNumber(value: number): string {
  if (value <= 0) return String(value);
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const units = ['', '十', '百', '千'];
  if (value < 10) return digits[value];
  if (value < 20) return value === 10 ? '十' : `十${digits[value % 10]}`;
  if (value < 10000) {
    const text = String(value);
    let result = '';
    let zero = false;
    for (let index = 0; index < text.length; index += 1) {
      const digit = Number(text[index]);
      const unit = units[text.length - index - 1];
      if (digit === 0) {
        zero = true;
        continue;
      }
      if (zero && result) result += '零';
      zero = false;
      result += `${digits[digit]}${unit}`;
    }
    return result;
  }
  return String(value);
}

function toRoman(value: number, upper: boolean): string {
  const numerals: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.max(1, Math.min(3999, Math.floor(value)));
  let text = '';
  for (const [number, symbol] of numerals) {
    while (remaining >= number) {
      text += symbol;
      remaining -= number;
    }
  }
  return upper ? text : text.toLowerCase();
}

function toLetters(value: number, upper: boolean): string {
  let n = Math.max(1, Math.floor(value));
  let text = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    text = String.fromCharCode((upper ? 65 : 97) + rem) + text;
    n = Math.floor((n - 1) / 26);
  }
  return text;
}

function formatListNumber(value: number, format: string): string {
  switch (format) {
    case 'lowerLetter':
      return toLetters(value, false);
    case 'upperLetter':
      return toLetters(value, true);
    case 'lowerRoman':
      return toRoman(value, false);
    case 'upperRoman':
      return toRoman(value, true);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
      return toChineseNumber(value);
    case 'decimalZero':
      return String(value).padStart(2, '0');
    case 'decimal':
    default:
      return String(value);
  }
}

function listInfoFor(ctx: ParseContext, numId: string, ilvl: number): ListInfo | undefined {
  const levels = ctx.numbering.get(numId);
  if (!levels || levels.length === 0) return undefined;
  let counters = ctx.counters.get(numId);
  if (!counters) {
    counters = new Map();
    ctx.counters.set(numId, counters);
  }
  const level = levels.find((item) => item.ilvl === ilvl) ?? levels[0];
  const current = (counters.get(ilvl) ?? level.start - 1) + 1;
  counters.set(ilvl, current);
  for (const key of [...counters.keys()]) {
    if (key > ilvl) counters.delete(key);
  }
  const values = new Map<number, number>();
  for (const item of levels) {
    values.set(item.ilvl, counters.get(item.ilvl) ?? item.start);
  }
  values.set(ilvl, current);
  let marker = level.lvlText;
  if (level.numFmt === 'bullet') {
    marker = marker || '•';
  } else {
    marker = marker.replace(/%([1-9])/g, (_match, index: string) => {
      const target = Number(index) - 1;
      const targetLevel = levels.find((item) => item.ilvl === target);
      return formatListNumber(values.get(target) ?? targetLevel?.start ?? 1, targetLevel?.numFmt ?? 'decimal');
    });
    if (!marker) marker = `${formatListNumber(current, level.numFmt)}.`;
  }
  return { numId, level: ilvl, marker, ordered: level.numFmt !== 'bullet' && level.numFmt !== 'none' };
}

function parseTableBorder(element: Element | null): TableBorderSide | undefined {
  if (!element) return undefined;
  const value = attrValue(element, 'w:val') ?? 'single';
  if (['nil', 'none'].includes(value)) {
    return { style: 'none', widthPx: 0, color: 'transparent' };
  }
  const size = numericAttr(element, 'w:sz', 4) ?? 4;
  const colorRaw = attrValue(element, 'w:color');
  const color =
    colorRaw && colorRaw.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(colorRaw)
      ? `#${colorRaw.toUpperCase()}`
      : '#000000';
  const styleMap: Record<string, TableBorderSide['style']> = {
    single: 'single',
    double: 'double',
    dashed: 'dashed',
    dotted: 'dotted',
    thick: 'single',
  };
  return {
    style: styleMap[value] ?? 'single',
    widthPx: Math.max(0.5, size / 6),
    color,
  };
}

function parseTableBorders(element: Element | null): TableBorders {
  if (!element) return {};
  return {
    top: parseTableBorder(firstChild(element, 'w:top')),
    right: parseTableBorder(firstChild(element, 'w:right')),
    bottom: parseTableBorder(firstChild(element, 'w:bottom')),
    left: parseTableBorder(firstChild(element, 'w:left')),
    insideH: parseTableBorder(firstChild(element, 'w:insideH')),
    insideV: parseTableBorder(firstChild(element, 'w:insideV')),
  };
}

async function loadImageUrl(ctx: ParseContext, part: string): Promise<string | undefined> {
  const cached = ctx.imageUrls.get(part);
  if (cached) return cached;
  if (!ctx.pkg.has(part)) return undefined;
  try {
    const bytes = await ctx.pkg.bytes(part);
    const buffer = new Uint8Array(bytes.byteLength);
    buffer.set(bytes);
    const url = URL.createObjectURL(new Blob([buffer], { type: mimeForPath(part) }));
    ctx.imageUrls.set(part, url);
    return url;
  } catch {
    return undefined;
  }
}

function vmlStyleSize(style: string | undefined): { width: number; height: number } | undefined {
  if (!style) return undefined;
  const width = /(?:^|;)\s*width\s*:\s*([\d.]+)(pt|px|in|cm)/i.exec(style);
  const height = /(?:^|;)\s*height\s*:\s*([\d.]+)(pt|px|in|cm)/i.exec(style);
  if (!width || !height) return undefined;
  const convert = (match: RegExpExecArray) => {
    const value = Number(match[1]);
    switch (match[2].toLowerCase()) {
      case 'pt':
        return value * (4 / 3);
      case 'in':
        return value * 96;
      case 'cm':
        return (value / 2.54) * 96;
      default:
        return value;
    }
  };
  return { width: convert(width), height: convert(height) };
}

async function parseDrawing(drawing: Element, ctx: ParseContext): Promise<DocRun[]> {
  const blip = firstDescendant(drawing, 'a:blip');
  const rid = attrValue(blip, 'r:embed') ?? attrValue(blip, 'r:link');
  if (!rid) return [];
  const rel = ctx.rels.get(rid);
  const part = rel?.resolved;
  if (!part) return [];
  const src = await loadImageUrl(ctx, part);
  if (!src) return [];
  const extent = firstDescendant(drawing, 'wp:extent');
  const widthPx = Math.max(8, emuToPx(numericAttr(extent, 'cx', 1_828_800) ?? 1_828_800));
  const heightPx = Math.max(8, emuToPx(numericAttr(extent, 'cy', 1_371_600) ?? 1_371_600));
  const docPr = firstDescendant(drawing, 'wp:docPr');
  const alt = attrValue(docPr, 'descr') ?? attrValue(docPr, 'title');
  const name = attrValue(docPr, 'name');
  const image: DocImage = {
    id: uid('img'),
    src,
    widthPx,
    heightPx,
    alt,
    name,
    relId: rid,
    part,
    // 功能：识别 Word 图片的内嵌/浮动形态，重新打开后保持文字环绕；时间：2026-10-24；作者：lq。
    wrap: drawing.getElementsByTagName('wp:anchor').length
      ? (firstDescendant(drawing, 'wp:wrapTopAndBottom') ? 'top-bottom' : attrValue(firstDescendant(drawing, 'wp:wrapSquare'), 'wrapText') === 'right' ? 'square-right' : 'square-left')
      : 'inline',
  };
  ctx.usedImages.set(rid, { src, part });
  return [{ text: '', style: {}, image }];
}

async function parseVmlPicture(pict: Element, ctx: ParseContext): Promise<DocRun[]> {
  const imageData = firstDescendant(pict, 'v:imagedata');
  const rid = attrValue(imageData, 'r:id') ?? attrValue(imageData, 'r:embed');
  if (!rid) return [];
  const rel = ctx.rels.get(rid);
  const part = rel?.resolved;
  if (!part) return [];
  const src = await loadImageUrl(ctx, part);
  if (!src) return [];
  const shape = firstDescendant(pict, 'v:shape');
  const size = vmlStyleSize(attrValue(shape, 'style'));
  const image: DocImage = {
    id: uid('img'),
    src,
    widthPx: size?.width ?? 160,
    heightPx: size?.height ?? 120,
    part,
    relId: rid,
    name: attrValue(shape, 'alt') ?? undefined,
  };
  ctx.usedImages.set(rid, { src, part });
  return [{ text: '', style: {}, image }];
}

interface ParseInlineState {
  fieldStack: Array<{ instr: string; phase: 'instr' | 'result'; mergeEmitted?: boolean }>;
}

async function parseRunElement(
  element: Element,
  base: RunStyle,
  ctx: ParseContext,
  state: ParseInlineState,
): Promise<DocRun[]> {
  const rPr = firstChild(element, 'w:rPr');
  const charStyleId = attrValue(firstChild(rPr, 'w:rStyle'), 'w:val');
  const charStyle = charStyleId ? ctx.styles.styles.get(charStyleId)?.run : undefined;
  const style = applyRunStyle(mergeRun(base, charStyle), rPr, ctx);
  const runs: DocRun[] = [];
  const topField = () => state.fieldStack.at(-1);

  for (const child of Array.from(element.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (child.tagName) {
      case 'w:t': {
        const text = child.textContent ?? '';
        if (text) {
          const run: DocRun = { text, style };
          const field = topField();
          if (field && field.phase === 'result') {
            // 功能：跨多个文本 run 的字段只生成一个原子占位符；时间：2026-10-24；作者：lq。
            const name = parseMergeField(field.instr);
            if (name) {
              if (!field.mergeEmitted) runs.push({ text: `«${name}»`, mergeField: name, style });
              field.mergeEmitted = true;
              break;
            }
            if (/\bPAGE\b/i.test(field.instr) && !/NUMPAGES/i.test(field.instr)) run.field = 'PAGE';
            else if (/NUMPAGES/i.test(field.instr)) run.field = 'NUMPAGES';
            else if (/SECTIONPAGES/i.test(field.instr)) run.field = 'SECTIONPAGES';
          }
          runs.push(run);
        }
        break;
      }
      case 'w:tab':
        runs.push({ text: '', style, tab: true });
        break;
      case 'w:br': {
        const type = attrValue(child, 'w:type');
        runs.push({ text: '', style, field: type === 'page' ? '__PAGE_BREAK__' : undefined, lineBreak: type !== 'page' });
        break;
      }
      case 'w:cr':
        runs.push({ text: '', style, lineBreak: true });
        break;
      case 'w:noBreakHyphen':
        runs.push({ text: '‑', style });
        break;
      case 'w:softHyphen':
        runs.push({ text: '\u00AD', style });
        break;
      case 'w:sym': {
        const charCode = attrValue(child, 'w:char');
        if (charCode) {
          const code = parseInt(charCode, 16);
          if (Number.isFinite(code)) runs.push({ text: String.fromCodePoint(code), style });
        }
        break;
      }
      case 'w:drawing': {
        runs.push(...(await parseDrawing(child, ctx)));
        break;
      }
      case 'w:pict': {
        runs.push(...(await parseVmlPicture(child, ctx)));
        break;
      }
      case 'w:fldChar': {
        const type = attrValue(child, 'w:fldCharType');
        if (type === 'begin') state.fieldStack.push({ instr: '', phase: 'instr' });
        else if (type === 'separate' && topField()) topField()!.phase = 'result';
        else if (type === 'end') {
          const field = topField();
          const name = field && parseMergeField(field.instr);
          if (name && !field?.mergeEmitted) runs.push({ text: `«${name}»`, mergeField: name, style });
          state.fieldStack.pop();
        }
        break;
      }
      case 'w:instrText': {
        const field = topField();
        if (field && field.phase === 'instr') field.instr += child.textContent ?? '';
        break;
      }
      case 'w:delText':
        break;
      default:
        break;
    }
  }
  return runs;
}

async function parseInlineChildren(
  parent: Element,
  base: RunStyle,
  ctx: ParseContext,
  state: ParseInlineState,
): Promise<DocRun[]> {
  const runs: DocRun[] = [];
  for (const child of Array.from(parent.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (child.tagName) {
      case 'w:r':
        runs.push(...(await parseRunElement(child, base, ctx, state)));
        break;
      case 'w:hyperlink':
      case 'w:ins':
      case 'w:smartTag':
      case 'w:customXml':
      case 'w:sdt': {
        const container =
          child.tagName === 'w:sdt' ? firstChild(child, 'w:sdtContent') ?? child : child;
        if (child.tagName === 'w:ins') ctx.hasTrackedChanges = true;
        if (child.tagName === 'w:sdt') ctx.hasContentControls = true;
        runs.push(...(await parseInlineChildren(container, base, ctx, state)));
        break;
      }
      case 'w:fldSimple': {
        const instr = attrValue(child, 'w:instr') ?? '';
        // 功能：兼容简式 MERGEFIELD，保持第一个结果字符的样式；时间：2026-10-24；作者：lq。
        const name = parseMergeField(instr);
        if (name) {
          const result = await parseInlineChildren(child, base, ctx, state);
          runs.push({ text: `«${name}»`, mergeField: name, style: result[0]?.style ?? base });
          break;
        }
        const before = runs.length;
        runs.push(...(await parseInlineChildren(child, base, ctx, state)));
        if (/\bPAGE\b/i.test(instr) && !/NUMPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'PAGE';
        } else if (/NUMPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'NUMPAGES';
        } else if (/SECTIONPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'SECTIONPAGES';
        }
        break;
      }
      case 'w:del':
        ctx.hasTrackedChanges = true;
        break;
      case 'w:bookmarkStart':
      case 'w:bookmarkEnd':
      case 'w:commentRangeStart':
      case 'w:commentRangeEnd':
      case 'w:proofErr':
        break;
      case 'w:br':
        runs.push({ text: '', style: base, lineBreak: getAttr(child, 'w:type') !== 'page' });
        break;
      default:
        break;
    }
  }
  return runs;
}

async function parseParagraph(
  element: Element,
  ctx: ParseContext,
  baseStyle?: { run: RunStyle; para: ParagraphStyle },
): Promise<{ paragraph: DocParagraph; pageBreakAfter: boolean }> {
  const pPr = firstChild(element, 'w:pPr');
  const styleId = attrValue(firstChild(pPr, 'w:pStyle'), 'w:val');
  const styleFromId = styleId ? ctx.styles.styles.get(styleId) : undefined;
  const styleBase = styleFromId ?? ctx.styles.defaultParagraph ?? baseStyle ?? {
    run: ctx.styles.docDefaults.run,
    para: ctx.styles.docDefaults.para,
  };
  let paragraphStyle = applyParagraphStyle(mergeParagraph(styleBase.para, baseStyle?.para), pPr, ctx);
  const numbering = parseProjectedParagraphNumbering(pPr);
  const numId = numbering.numId ?? styleFromId?.numId;
  const ilvl = numbering.ilvl ?? styleFromId?.ilvl ?? 0;
  let list: ListInfo | undefined;
  if (numId) {
    list = listInfoFor(ctx, numId, ilvl);
    const level = ctx.numbering.get(numId)?.find((item) => item.ilvl === ilvl);
    if (level) {
      if (paragraphStyle.indentLeftPx === undefined && level.indentLeftPx !== undefined) {
        paragraphStyle.indentLeftPx = level.indentLeftPx;
      }
      if (
        paragraphStyle.firstLineIndentPx === undefined &&
        level.firstLineIndentPx !== undefined
      ) {
        paragraphStyle.firstLineIndentPx = level.firstLineIndentPx;
      }
    }
  }

  const runBase = mergeRun(mergeRun(ctx.styles.docDefaults.run, styleBase.run), baseStyle?.run);
  const state: ParseInlineState = { fieldStack: [] };
  let runs = await parseInlineChildren(element, runBase, ctx, state);
  const pageBreakAfter = runs.some((run) => run.field === '__PAGE_BREAK__');
  runs = runs.filter((run) => run.field !== '__PAGE_BREAK__');
  if (runs.length === 0) runs = [{ text: '', style: runBase }];

  const paragraph: DocParagraph = {
    type: 'paragraph',
    id: uid('p'),
    runs,
    style: paragraphStyle,
    list,
    styleId,
  };
  // 功能：保留段尾分节以及各节独立的页眉页脚与编号；时间：2026-10-24；作者：lq。
  const section = firstChild(pPr, 'w:sectPr');
  if (section) {
    const page = parsePageSetup(section, ctx);
    paragraph.sectionEnd = {
      page,
      header: await parseHeaderFooter(section, ctx, 'header', page.margins.header),
      footer: await parseHeaderFooter(section, ctx, 'footer', page.margins.footer),
    };
  }
  return { paragraph, pageBreakAfter };
}

async function parseTable(element: Element, ctx: ParseContext): Promise<DocTable> {
  const tblPr = firstChild(element, 'w:tblPr');
  const grid = firstChild(element, 'w:tblGrid');
  let columnWidths = Array.from(grid?.childNodes ?? [])
    .filter((node): node is Element => node instanceof Element && node.tagName === 'w:gridCol')
    .map((col) => twipsToPx(numericAttr(col, 'w:w', 0) ?? 0))
    .filter((width) => width > 0);

  const styleId = attrValue(firstChild(tblPr, 'w:tblStyle'), 'w:val');
  const borders = parseTableBorders(firstChild(tblPr, 'w:tblBorders'));
  const hasAnyBorder = Object.values(borders).some((border) => border && border.style !== 'none');
  if (!hasAnyBorder && styleId && /grid/i.test(styleId)) {
    const single: TableBorderSide = { style: 'single', widthPx: 0.75, color: '#000000' };
    borders.top = single;
    borders.bottom = single;
    borders.left = single;
    borders.right = single;
    borders.insideH = single;
    borders.insideV = single;
  }

  const alignmentValue = attrValue(firstChild(tblPr, 'w:jc'), 'w:val');
  const alignmentMap: Record<string, DocTable['alignment']> = {
    center: 'center',
    right: 'right',
    end: 'right',
    left: 'left',
    start: 'left',
  };
  const tableWidth = firstChild(tblPr, 'w:tblW');
  const widthType = attrValue(tableWidth, 'w:type');
  const widthValue = numericAttr(tableWidth, 'w:w');
  const widthPct =
    widthType === 'pct' && widthValue !== undefined ? widthValue / 50 : undefined;

  interface RawCell {
    cell: DocTableCell;
    vMerge?: 'restart' | 'continue';
    colSpan: number;
  }

  const rawRows: Array<{ cells: RawCell[]; heightPx?: number; isHeader: boolean }> = [];
  let colCount = columnWidths.length;

  for (const rowElement of Array.from(element.childNodes)) {
    if (!(rowElement instanceof Element) || rowElement.tagName !== 'w:tr') continue;
    const trPr = firstChild(rowElement, 'w:trPr');
    const height = numericAttr(firstChild(trPr, 'w:trHeight'), 'w:val');
    const isHeader = toggleValue(firstChild(trPr, 'w:tblHeader'));
    const rawCells: RawCell[] = [];
    for (const cellElement of Array.from(rowElement.childNodes)) {
      if (!(cellElement instanceof Element) || cellElement.tagName !== 'w:tc') continue;
      const tcPr = firstChild(cellElement, 'w:tcPr');
      const gridSpan = numericAttr(firstChild(tcPr, 'w:gridSpan'), 'w:val') ?? 1;
      const vMergeElement = firstChild(tcPr, 'w:vMerge');
      const vMergeValue = vMergeElement ? attrValue(vMergeElement, 'w:val') : undefined;
      const vMerge: RawCell['vMerge'] = vMergeElement
        ? vMergeValue === 'restart'
          ? 'restart'
          : 'continue'
        : undefined;
      const cellWidthTwips = numericAttr(firstChild(tcPr, 'w:tcW'), 'w:w');
      const cellWidth = cellWidthTwips ? twipsToPx(cellWidthTwips) : undefined;
      const backgroundRaw = attrValue(firstChild(tcPr, 'w:shd'), 'w:fill');
      const backgroundColor =
        backgroundRaw && /^[0-9a-fA-F]{6}$/.test(backgroundRaw)
          ? `#${backgroundRaw.toUpperCase()}`
          : undefined;
      const vAlignValue = attrValue(firstChild(tcPr, 'w:vAlign'), 'w:val');
      const verticalAlign =
        vAlignValue === 'center' ? 'middle' : vAlignValue === 'bottom' ? 'bottom' : 'top';
      const cellBorders = parseTableBorders(firstChild(tcPr, 'w:tcBorders'));
      const mergedBorders: TableBorders = { ...borders, ...cellBorders };
      const blocks =
        vMerge === 'continue' ? [] : await parseBlockContainer(cellElement, ctx);
      if (vMerge !== 'continue' && blocks.length === 0) {
        blocks.push({
          type: 'paragraph',
          id: uid('p'),
          runs: [{ text: '', style: {} }],
          style: {},
        });
      }
      rawCells.push({
        colSpan: gridSpan,
        vMerge,
        cell: {
          id: uid('cell'),
          blocks,
          colSpan: gridSpan,
          rowSpan: 1,
          widthPx: cellWidth,
          backgroundColor,
          borders: mergedBorders,
          verticalAlign,
        },
      });
      colCount = Math.max(
        colCount,
        rawCells.reduce((sum, cell) => sum + cell.colSpan, 0),
      );
    }
    rawRows.push({
      cells: rawCells,
      heightPx: height ? twipsToPx(height) : undefined,
      isHeader,
    });
  }

  if (columnWidths.length === 0) {
    const available = Math.max(320, ctx.pageWidthPx - 192);
    columnWidths = Array.from({ length: Math.max(1, colCount) }, () => available / Math.max(1, colCount));
  }

  const columnTops: Array<DocTableCell | undefined> = [];
  const rows: DocTableRow[] = [];
  for (const rawRow of rawRows) {
    const cells: DocTableCell[] = [];
    let column = 0;
    for (const raw of rawRow.cells) {
      const span = Math.max(1, raw.colSpan);
      if (raw.vMerge === 'continue') {
        const top = columnTops[column];
        if (top) {
          top.rowSpan += 1;
          for (let offset = 0; offset < span; offset += 1) columnTops[column + offset] = top;
        }
      } else {
        raw.cell.rowSpan = 1;
        cells.push(raw.cell);
        for (let offset = 0; offset < span; offset += 1) columnTops[column + offset] = raw.cell;
      }
      column += span;
    }
    // 功能：保留完全被纵向合并覆盖的行，防止保存重开后表格坍缩；时间：2026-10-19；作者：lq。
    {
      rows.push({
        id: uid('row'),
        cells,
        heightPx: rawRow.heightPx,
        isHeader: rawRow.isHeader,
      });
    }
  }

  return {
    type: 'table',
    id: uid('tbl'),
    rows,
    columnWidthsPx: columnWidths,
    widthPct,
    alignment: alignmentValue ? alignmentMap[alignmentValue] ?? 'left' : 'left',
    borders,
    styleId,
  };
}

async function parseBlockContainer(element: Element, ctx: ParseContext): Promise<DocBlock[]> {
  const blocks: DocBlock[] = [];
  for (const child of Array.from(element.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (child.tagName) {
      case 'w:p': {
        const { paragraph, pageBreakAfter } = await parseParagraph(child, ctx);
        // 功能：纯分页标记不额外占用一行，避免长表格重开后插入空白页；时间：2026-10-24；作者：lq。
        const breakOnly = pageBreakAfter && !paragraph.sectionEnd && paragraph.runs.every((run) => !run.text && !run.image && !run.tab && !run.lineBreak);
        if (!breakOnly) blocks.push(paragraph);
        if (pageBreakAfter) blocks.push({ type: 'pageBreak', id: uid('break') });
        break;
      }
      case 'w:tbl':
        blocks.push(await parseTable(child, ctx));
        break;
      case 'w:sdt': {
        ctx.hasContentControls = true;
        const content = firstChild(child, 'w:sdtContent');
        if (content) blocks.push(...(await parseBlockContainer(content, ctx)));
        break;
      }
      case 'w:customXml': {
        blocks.push(...(await parseBlockContainer(child, ctx)));
        break;
      }
      case 'w:altChunk':
      case 'w:sectPr':
        break;
      default:
        break;
    }
  }
  return blocks;
}

function parsePageSetup(sectPr: Element | null, ctx: ParseContext): PageSetup {
  const pageSize = firstChild(sectPr, 'w:pgSz');
  const widthTwips = numericAttr(pageSize, 'w:w', 11_906) ?? 11_906;
  const heightTwips = numericAttr(pageSize, 'w:h', 16_838) ?? 16_838;
  const orientation = attrValue(pageSize, 'w:orient') === 'landscape' ? 'landscape' : 'portrait';
  const marginsElement = firstChild(sectPr, 'w:pgMar');
  const margins = {
    top: twipsToPx(numericAttr(marginsElement, 'w:top', 1440) ?? 1440),
    right: twipsToPx(numericAttr(marginsElement, 'w:right', 1440) ?? 1440),
    bottom: twipsToPx(numericAttr(marginsElement, 'w:bottom', 1440) ?? 1440),
    left: twipsToPx(numericAttr(marginsElement, 'w:left', 1440) ?? 1440),
    header: twipsToPx(numericAttr(marginsElement, 'w:header', 720) ?? 720),
    footer: twipsToPx(numericAttr(marginsElement, 'w:footer', 720) ?? 720),
    gutter: twipsToPx(numericAttr(marginsElement, 'w:gutter', 0) ?? 0),
  };
  const sizeMap: Array<{ name: string; w: number; h: number }> = [
    { name: 'A4', w: 11_906, h: 16_838 },
    { name: 'A3', w: 16_838, h: 23_811 },
    { name: 'A5', w: 8_391, h: 11_906 },
    { name: 'Letter', w: 12_240, h: 15_840 },
    { name: 'Legal', w: 12_240, h: 20_160 },
    { name: 'B5', w: 9_976, h: 14_170 },
  ];
  const swapped = Math.min(widthTwips, heightTwips) === Math.min(widthTwips, heightTwips);
  const matched = sizeMap.find(
    (size) =>
      (Math.abs(size.w - widthTwips) < 80 && Math.abs(size.h - heightTwips) < 80) ||
      (Math.abs(size.w - heightTwips) < 80 && Math.abs(size.h - widthTwips) < 80),
  );
  const columns = numericAttr(firstChild(sectPr, 'w:cols'), 'w:num');
  return {
    widthPx: twipsToPx(widthTwips),
    heightPx: twipsToPx(heightTwips),
    orientation,
    sizeName: matched?.name ?? (swapped ? '自定义' : '自定义'),
    margins,
    columns: columns && columns > 1 ? columns : undefined,
    // 功能：恢复 DOCX 起始页码；时间：2026-10-24；作者：lq。
    pageNumberStart: numericAttr(firstChild(sectPr, 'w:pgNumType'), 'w:start'),
  };
}

async function parseHeaderFooter(
  sectPr: Element | null,
  ctx: ParseContext,
  kind: 'header' | 'footer',
  distancePx: number,
): Promise<HeaderFooter | undefined> {
  if (!sectPr) return undefined;
  const references = Array.from(sectPr.childNodes).filter(
    (node): node is Element =>
      node instanceof Element && node.tagName === `w:${kind}Reference`,
  );
  if (references.length === 0) return undefined;
  const preferred =
    references.find((reference) => attrValue(reference, 'w:type') === 'default') ?? references[0];
  const relId = preferred.getAttribute('r:id') ?? getAttrLocal(preferred, 'id');
  if (!relId) return undefined;
  const rel = ctx.rels.get(relId);
  if (!rel?.resolved) return undefined;
  const xml = await ctx.pkg.text(rel.resolved);
  const doc = parseXmlDom(xml);
  const root = firstDescendant(doc, kind === 'header' ? 'w:hdr' : 'w:ftr') ?? doc.documentElement;
  // 功能：页眉页脚图片关系按自身部件解析；时间：2026-10-24；作者：lq。
  const blocks = await parseBlockContainer(root, { ...ctx, rels: await relationshipMap(ctx.pkg, rel.resolved) });
  const pageNumbers = blocks.some((block) =>
    block.type === 'paragraph' ? block.runs.some((run) => run.field === 'PAGE') : false,
  );
  return { blocks, distancePx, pageNumbers, sourcePart: rel.resolved };
}

export interface ParseDocxOptions {
  fileName?: string;
}

export async function parseDocxPackage(
  pkg: OoxmlPackage,
  options: ParseDocxOptions = {},
): Promise<DocModel> {
  const documentXml = await pkg.text('word/document.xml');
  const stylesXml = pkg.has('word/styles.xml') ? await pkg.text('word/styles.xml') : '';
  const numberingXml = pkg.has('word/numbering.xml') ? await pkg.text('word/numbering.xml') : '';
  const themeColors = await readThemeColors(pkg, 'word/theme/theme1.xml');
  const themeFonts = await readThemeFonts(pkg, 'word/theme/theme1.xml');
  const rels = await relationshipMap(pkg, 'word/document.xml');

  const provisional: ParseContext = {
    pkg,
    rels,
    styles: {
      docDefaults: { run: {}, para: {} },
      styles: new Map(),
    },
    numbering: parseNumbering(numberingXml),
    counters: new Map(),
    themeColors,
    themeFonts,
    defaultFont: themeFonts.minor || '宋体',
    defaultSizePt: 10.5,
    imageUrls: new Map(),
    usedImages: new Map(),
    pageWidthPx: 793.7,
    blockCounter: 0,
    hasTrackedChanges: false,
    hasContentControls: false,
  };

  provisional.styles = parseStyleSheet(stylesXml, provisional);
  const defaultRun = provisional.styles.docDefaults.run;
  const defaultFont =
    defaultRun.fontFamily ??
    themeFonts.minor ??
    themeFonts.major ??
    provisional.defaultFont;
  provisional.defaultFont = defaultFont;
  provisional.defaultSizePt = defaultRun.fontSizePt ?? 10.5;

  const doc = parseXmlDom(documentXml);
  const body = firstDescendant(doc, 'w:body') ?? doc.documentElement;
  const sectionProperties =
    Array.from(body.childNodes).find(
      (node): node is Element => node instanceof Element && node.tagName === 'w:sectPr',
    ) ??
    firstDescendant(doc, 'w:sectPr');
  const page = parsePageSetup(sectionProperties, provisional);
  provisional.pageWidthPx = page.widthPx;

  const blocks = await parseBlockContainer(body, provisional);

  const header = await parseHeaderFooter(
    sectionProperties,
    provisional,
    'header',
    page.margins.header,
  );
  const footer = await parseHeaderFooter(
    sectionProperties,
    provisional,
    'footer',
    page.margins.footer,
  );

  let words = 0;
  let characters = 0;
  const visitParagraph = (paragraph: DocParagraph) => {
    const text = paragraph.runs.map((run) => run.text ?? '').join('');
    words += countWords(text);
    characters += text.replace(/\s/g, '').length;
  };
  const visitBlocks = (items: DocBlock[]) => {
    for (const block of items) {
      if (block.type === 'paragraph') visitParagraph(block);
      else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
      }
    }
  };
  visitBlocks(blocks);
  if (header) visitBlocks(header.blocks);
  if (footer) visitBlocks(footer.blocks);

  const meta = await parseDocMeta(pkg);

  return {
    version: 1,
    page,
    blocks,
    header,
    footer,
    defaultFont,
    defaultFontSizePt: provisional.defaultSizePt,
    themeColors,
    themeFonts,
    images: Object.fromEntries(provisional.usedImages),
    meta,
    hasTrackedChanges: provisional.hasTrackedChanges,
    hasComments: pkg.has('word/comments.xml'),
    hasContentControls: provisional.hasContentControls,
    stats: { words, characters },
  };
}

async function parseDocMeta(pkg: OoxmlPackage): Promise<DocModel['meta']> {
  if (!pkg.has('docProps/core.xml')) return {};
  try {
    const doc = parseXmlDom(await pkg.text('docProps/core.xml'));
    const val = (names: string[]): string | undefined => {
      for (const name of names) {
        const element =
          doc.getElementsByTagName(name)[0] ?? doc.getElementsByTagName(`dc:${name}`)[0];
        if (element?.textContent) return element.textContent;
      }
      return undefined;
    };
    return {
      title: val(['title', 'dc:title']),
      author: val(['creator', 'dc:creator']),
      subject: val(['subject', 'dc:subject']),
      created: val(['created', 'dcterms:created']),
      modified: val(['modified', 'dcterms:modified']),
      lastModifiedBy: val(['lastModifiedBy', 'cp:lastModifiedBy']),
    };
  } catch {
    return {};
  }
}

export async function parseDocx(bytes: Uint8Array, options: ParseDocxOptions = {}): Promise<DocModel> {
  const pkg = await loadOoxmlPackage(bytes);
  try {
    return await parseDocxPackage(pkg, options);
  } finally {
    pkg.dispose();
  }
}
