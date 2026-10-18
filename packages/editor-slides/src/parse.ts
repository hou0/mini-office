import type {
  SlideDeck,
  SlideElement,
  SlideFill,
  SlideGroupElement,
  SlideImageElement,
  SlideLine,
  SlideModel,
  SlideShapeElement,
  SlideTableElement,
  SlideTextElement,
  SlideTextParagraph,
  SlideTextRun,
  TextRunStyle,
} from '@mini-office/document-model';
import {
  cssColor,
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  loadOoxmlPackage,
  mimeForPath,
  parseXmlDom,
  readThemeColors,
  relationshipMap,
  emuToPx,
  type OoxmlPackage,
  type Relationship,
} from '@mini-office/ooxml';
import { stripExtension, uid, countWords } from '@mini-office/shared';

interface PlaceholderStyle {
  x: number;
  y: number;
  w: number;
  h: number;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: SlideTextParagraph['align'];
  verticalAlign?: SlideTextElement['verticalAlign'];
  fontFamily?: string;
}

interface ParsedContainer {
  elements: SlideElement[];
  placeholders: Map<string, PlaceholderStyle>;
  background?: SlideFill;
}

interface ParseContext {
  pkg: OoxmlPackage;
  theme: Record<string, string>;
  themeFonts: { major: string; minor: string };
  imageCache: Map<string, string>;
  relCache: Map<string, Map<string, Relationship>>;
  containerCache: Map<string, ParsedContainer>;
  slideWidth: number;
  slideHeight: number;
}

function childrenNamed(parent: Element | Document | null | undefined, tagName: string): Element[] {
  if (!parent) return [];
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node instanceof Element && node.tagName === tagName,
  );
}

function firstNamed(parent: Element | Document | null | undefined, tagName: string): Element | null {
  if (!parent) return null;
  return Array.from(parent.childNodes).find(
    (node): node is Element => node instanceof Element && node.tagName === tagName,
  ) ?? null;
}

function numberAttr(element: Element | null | undefined, name: string, fallback = 0): number {
  const raw = getAttr(element, name);
  if (raw === undefined || raw === '') return fallback;
  const value = Number(raw);
  return Number.isFinite(value) ? value : fallback;
}

function boolAttr(element: Element | null | undefined, name: string): boolean {
  const value = getAttr(element, name);
  return value === '1' || value === 'true';
}

function percentValue(element: Element | null | undefined): number | undefined {
  const raw = getAttr(element, 'val');
  if (raw === undefined || raw === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value / 100_000 : undefined;
}

function resolveColor(node: Element | null, ctx: ParseContext, fallback?: string): string | undefined {
  if (!node) return fallback;
  const scheme = firstNamed(node, 'a:schemeClr');
  if (scheme) {
    const value = getAttr(scheme, 'val') ?? 'phClr';
    const base = value === 'phClr' ? fallback ?? ctx.theme.accent1 : ctx.theme[value];
    if (!base) return fallback;
    return cssColor(
      {
        hex: base,
        tint: numberAttr(firstNamed(scheme, 'a:tint'), 'val', 100_000),
        shade: numberAttr(firstNamed(scheme, 'a:shade'), 'val', 100_000),
        lumMod: numberAttr(firstNamed(scheme, 'a:lumMod'), 'val', 100_000),
        lumOff: numberAttr(firstNamed(scheme, 'a:lumOff'), 'val', 0),
        alpha: numberAttr(firstNamed(scheme, 'a:alpha'), 'val', 100_000),
      },
      ctx.theme,
    );
  }
  const srgb = firstNamed(node, 'a:srgbClr');
  if (srgb) {
    const value = getAttr(srgb, 'val');
    if (!value) return fallback;
    return cssColor(
      {
        hex: `#${value}`,
        alpha: numberAttr(firstNamed(srgb, 'a:alpha'), 'val', 100_000),
      },
      ctx.theme,
    );
  }
  const preset = firstNamed(node, 'a:prstClr');
  if (preset) {
    const value = getAttr(preset, 'val');
    if (value === 'black') return '#000000';
    if (value === 'white') return '#FFFFFF';
  }
  const sys = firstNamed(node, 'a:sysClr');
  if (sys) {
    const value = getAttr(sys, 'lastClr');
    if (value) return `#${value.slice(-6)}`;
  }
  return fallback;
}

function parseFillElement(element: Element | null, ctx: ParseContext, fallback?: string): SlideFill | undefined {
  if (!element) return undefined;
  const solid = firstNamed(element, 'a:solidFill');
  if (solid) {
    const color = resolveColor(solid, ctx, fallback);
    if (color) return { type: 'solid', color };
  }
  const noFill = firstNamed(element, 'a:noFill');
  if (noFill) return { type: 'none' };
  const gradient = firstNamed(element, 'a:gradFill');
  if (gradient) {
    const stops: Array<{ offset: number; color: string }> = [];
    for (const gs of childrenNamed(firstNamed(gradient, 'a:gsLst'), 'a:gs')) {
      const position = numberAttr(gs, 'pos', 0) / 100_000;
      const color = resolveColor(gs, ctx);
      if (color) stops.push({ offset: position, color });
    }
    if (stops.length > 0) {
      const linear = firstNamed(gradient, 'a:lin');
      return {
        type: 'gradient',
        stops,
        angle: linear ? numberAttr(linear, 'ang', 0) / 60_000 : 0,
      };
    }
  }
  const blip = firstNamed(element, 'a:blipFill');
  if (blip) {
    const embed = getAttr(firstDescendant(blip, 'a:blip'), 'r:embed');
    if (embed) {
      const part = ctx.relCache.get('__slide__')?.get(embed)?.resolved;
      if (part) {
        const src = ctx.imageCache.get(part);
        if (src) return { type: 'image', src };
      }
    }
  }
  return undefined;
}

function parseLine(element: Element | null, ctx: ParseContext): SlideLine | undefined {
  const ln = firstNamed(element, 'a:ln');
  if (!ln) return undefined;
  const widthEmu = numberAttr(ln, 'w', 0);
  const color = resolveColor(firstNamed(ln, 'a:solidFill'), ctx);
  const dashElement = firstNamed(ln, 'a:prstDash');
  const dashRaw = getAttr(dashElement, 'val');
  const dashMap: Record<string, SlideLine['dash']> = {
    solid: 'solid',
    dash: 'dash',
    sysDash: 'dash',
    dot: 'dot',
    sysDot: 'dot',
    dashDot: 'dashDot',
    lgDash: 'longDash',
  };
  if (!color && widthEmu === 0) return undefined;
  return {
    color,
    widthPx: widthEmu > 0 ? emuToPx(widthEmu) : 1,
    dash: dashRaw ? dashMap[dashRaw] ?? 'solid' : 'solid',
  };
}

function parseStyleRefs(style: Element | null, ctx: ParseContext): { fill?: SlideFill; line?: SlideLine } {
  if (!style) return {};
  const fillRef = firstNamed(style, 'a:fillRef');
  const fillIndex = numberAttr(fillRef, 'idx', 0);
  const accentColors = [
    ctx.theme.accent1,
    ctx.theme.accent2,
    ctx.theme.accent3,
    ctx.theme.accent4,
    ctx.theme.accent5,
    ctx.theme.accent6,
  ];
  const fillColor =
    fillIndex > 0
      ? resolveColor(fillRef, ctx, accentColors[(fillIndex - 1) % 6] ?? ctx.theme.accent1) ??
        accentColors[(fillIndex - 1) % 6]
      : undefined;
  const fill = fillColor ? ({ type: 'solid', color: fillColor } as SlideFill) : undefined;
  const lnRef = firstNamed(style, 'a:lnRef');
  const lineIndex = numberAttr(lnRef, 'idx', 0);
  const lineColor =
    lineIndex > 0 ? resolveColor(lnRef, ctx, accentColors[(lineIndex - 1) % 6]) : undefined;
  const line: SlideLine | undefined = lineColor
    ? { color: lineColor, widthPx: 1, dash: 'solid' }
    : undefined;
  return { fill, line };
}

function parseTextRunStyle(
  rPr: Element | null,
  ctx: ParseContext,
  fallback?: TextRunStyle,
): TextRunStyle {
  const style: TextRunStyle = { ...fallback };
  if (!rPr) return style;
  const size = numberAttr(rPr, 'sz', 0);
  if (size > 0) style.fontSizePt = size / 100;
  if (getAttr(rPr, 'b') !== undefined) style.bold = boolAttr(rPr, 'b');
  if (getAttr(rPr, 'i') !== undefined) style.italic = boolAttr(rPr, 'i');
  const underline = getAttr(rPr, 'u');
  if (underline) style.underline = underline !== 'none';
  if (getAttr(rPr, 'strike') !== undefined) style.strike = getAttr(rPr, 'strike') !== 'noStrike';
  const color = resolveColor(firstNamed(rPr, 'a:solidFill'), ctx);
  if (color) style.color = color;
  const highlight = resolveColor(firstNamed(rPr, 'a:highlight'), ctx);
  if (highlight) style.highlight = highlight;
  const latin = firstNamed(rPr, 'a:latin');
  const ea = firstNamed(rPr, 'a:ea');
  const fontFamily = getAttr(ea, 'typeface') ?? getAttr(latin, 'typeface');
  if (fontFamily && !fontFamily.startsWith('+')) style.fontFamily = fontFamily;
  const spacing = numberAttr(rPr, 'spc', 0);
  if (spacing) style.letterSpacingPx = (spacing / 100) * (4 / 3);
  return style;
}

function defaultPlaceholderRunStyle(
  placeholder: PlaceholderStyle | undefined,
  ctx: ParseContext,
): TextRunStyle {
  return {
    fontFamily: placeholder?.fontFamily ?? ctx.themeFonts.minor ?? '等线',
    fontSizePt: placeholder?.fontSizePt ?? 18,
    bold: placeholder?.bold,
    italic: placeholder?.italic,
    color: placeholder?.color ?? ctx.theme.dk1 ?? '#1D2939',
  };
}

function parseParagraphs(
  txBody: Element | null,
  ctx: ParseContext,
  fallbackRun: TextRunStyle,
  fallbackAlign?: SlideTextParagraph['align'],
): SlideTextParagraph[] {
  if (!txBody) return [];
  const paragraphs: SlideTextParagraph[] = [];
  for (const p of childrenNamed(txBody, 'a:p')) {
    const pPr = firstNamed(p, 'a:pPr');
    const level = numberAttr(pPr, 'lvl', 0);
    const alignRaw = getAttr(pPr, 'algn');
    const alignMap: Record<string, SlideTextParagraph['align']> = {
      l: 'left',
      ctr: 'center',
      r: 'right',
      just: 'justify',
    };
    const bulletNone = firstNamed(pPr, 'a:buNone');
    const bulletChar = getAttr(firstNamed(pPr, 'a:buChar'), 'char');
    const bulletAuto = firstNamed(pPr, 'a:buAutoNum');
    const spaceBefore = firstNamed(pPr, 'a:spcBef');
    const spaceAfter = firstNamed(pPr, 'a:spcAft');
    const lineSpacingPercent = percentValue(firstNamed(pPr, 'a:lnSpc'));
    const defaultRun = parseTextRunStyle(
      firstNamed(pPr, 'a:defRPr') ?? firstNamed(p, 'a:endParaRPr'),
      ctx,
      fallbackRun,
    );
    const runs: SlideTextRun[] = [];
    for (const child of Array.from(p.childNodes)) {
      if (!(child instanceof Element)) continue;
      if (child.tagName === 'a:r') {
        const rPr = firstNamed(child, 'a:rPr');
        runs.push({
          text: firstNamed(child, 'a:t')?.textContent ?? '',
          style: parseTextRunStyle(rPr, ctx, defaultRun),
        });
      } else if (child.tagName === 'a:br') {
        runs.push({ text: '', style: defaultRun, lineBreak: true });
      } else if (child.tagName === 'a:fld') {
        runs.push({
          text: firstNamed(child, 'a:t')?.textContent ?? '',
          style: parseTextRunStyle(firstNamed(child, 'a:rPr'), ctx, defaultRun),
        });
      }
    }
    const spacePointValue = (element: Element | null) => {
      if (!element) return undefined;
      const points = firstNamed(element, 'a:spcPts');
      const value = numberAttr(points, 'val', -1);
      return value >= 0 ? value / 100 : undefined;
    };
    paragraphs.push({
      runs,
      align: alignMap[alignRaw ?? ''] ?? fallbackAlign,
      level,
      bullet: bulletChar || bulletAuto ? true : bulletNone ? false : undefined,
      bulletChar: bulletChar ?? (bulletAuto ? '•' : undefined),
      spaceBeforePt: spacePointValue(spaceBefore),
      spaceAfterPt: spacePointValue(spaceAfter),
      lineSpacing: lineSpacingPercent,
    });
  }
  return paragraphs;
}

function plainText(paragraphs: SlideTextParagraph[]): string {
  return paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n')
    .trim();
}

function placeholderKey(ph: Element | null): string | undefined {
  if (!ph) return undefined;
  const type = getAttr(ph, 'type') ?? 'body';
  const idx = getAttr(ph, 'idx');
  return `${type}:${idx ?? ''}`;
}

function parseShape(
  shape: Element,
  ctx: ParseContext,
  placeholderStyles: Map<string, PlaceholderStyle>,
  fallbackRun: TextRunStyle,
): { element?: SlideTextElement | SlideShapeElement; placeholder?: { key: string; style: PlaceholderStyle } } {
  const spPr = firstNamed(shape, 'p:spPr');
  const xfrm = firstNamed(spPr, 'a:xfrm');
  const off = firstNamed(xfrm, 'a:off');
  const ext = firstNamed(xfrm, 'a:ext');
  const ph = firstDescendant(shape, 'p:ph');
  const key = placeholderKey(ph);
  const inherited = key ? placeholderStyles.get(key) : undefined;
  let x = emuToPx(numberAttr(off, 'x', 0));
  let y = emuToPx(numberAttr(off, 'y', 0));
  let w = emuToPx(numberAttr(ext, 'cx', 0));
  let h = emuToPx(numberAttr(ext, 'cy', 0));
  if ((!off || !ext) && inherited) {
    x = inherited.x;
    y = inherited.y;
    w = inherited.w;
    h = inherited.h;
  }
  if (!off && !inherited) return {};
  const rotation = numberAttr(xfrm, 'rot', 0) / 60_000;
  const flipH = boolAttr(xfrm, 'flipH');
  const flipV = boolAttr(xfrm, 'flipV');
  const txBody = firstNamed(shape, 'p:txBody');
  const bodyPr = firstNamed(txBody, 'a:bodyPr');
  const anchor = getAttr(bodyPr, 'anchor');
  const verticalAlign: SlideTextElement['verticalAlign'] =
    anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : inherited?.verticalAlign ?? 'top';
  const margins = {
    top: emuToPx(numberAttr(bodyPr, 'tIns', 45_720)),
    right: emuToPx(numberAttr(bodyPr, 'rIns', 91_440)),
    bottom: emuToPx(numberAttr(bodyPr, 'bIns', 45_720)),
    left: emuToPx(numberAttr(bodyPr, 'lIns', 91_440)),
  };
  const effectiveFallback: TextRunStyle = {
    ...fallbackRun,
    ...(inherited
      ? {
          fontSizePt: inherited.fontSizePt ?? fallbackRun.fontSizePt,
          bold: inherited.bold ?? fallbackRun.bold,
          italic: inherited.italic ?? fallbackRun.italic,
          color: inherited.color ?? fallbackRun.color,
          fontFamily: inherited.fontFamily ?? fallbackRun.fontFamily,
        }
      : {}),
  };
  const paragraphs = parseParagraphs(txBody, ctx, effectiveFallback, inherited?.align);
  const geometry = firstNamed(spPr, 'a:prstGeom');
  const preset = getAttr(geometry, 'prst') ?? 'rect';
  const styleRefs = parseStyleRefs(firstNamed(shape, 'p:style'), ctx);
  const fill = parseFillElement(spPr, ctx) ?? styleRefs.fill;
  const line = parseLine(spPr, ctx) ?? styleRefs.line;
  const name = getAttr(firstDescendant(shape, 'p:cNvPr'), 'name') ?? undefined;
  const base = {
    id: uid('el'),
    name,
    x,
    y,
    w,
    h,
    rotation: rotation || undefined,
    flipH: flipH || undefined,
    flipV: flipV || undefined,
  };
  if (paragraphs.length > 0 || !geometry) {
    const element: SlideTextElement = {
      ...base,
      type: 'text',
      paragraphs: paragraphs.length
        ? paragraphs
        : [{ runs: [{ text: '', style: effectiveFallback }] }],
      fill,
      line,
      verticalAlign,
      margins,
      placeholderType: getAttr(ph, 'type') ?? undefined,
      shape: preset !== 'rect' ? preset : undefined,
    };
    return {
      element,
      placeholder: key
        ? {
            key,
            style: {
              x,
              y,
              w,
              h,
              fontSizePt: paragraphs[0]?.runs[0]?.style.fontSizePt ?? effectiveFallback.fontSizePt,
              bold: paragraphs[0]?.runs[0]?.style.bold,
              color: paragraphs[0]?.runs[0]?.style.color,
              align: paragraphs[0]?.align,
              verticalAlign,
              fontFamily: paragraphs[0]?.runs[0]?.style.fontFamily,
            },
          }
        : undefined,
    };
  }
  const element: SlideShapeElement = {
    ...base,
    type: 'shape',
    shape: preset,
    fill,
    line,
    paragraphs,
    verticalAlign,
  };
  return { element };
}

async function loadImage(ctx: ParseContext, part: string): Promise<string | undefined> {
  const cached = ctx.imageCache.get(part);
  if (cached) return cached;
  if (!ctx.pkg.has(part)) return undefined;
  try {
    const bytes = await ctx.pkg.bytes(part);
    const buffer = new Uint8Array(bytes.byteLength);
    buffer.set(bytes);
    const url = URL.createObjectURL(new Blob([buffer], { type: mimeForPath(part) }));
    ctx.imageCache.set(part, url);
    return url;
  } catch {
    return undefined;
  }
}

async function parsePicture(
  pic: Element,
  ctx: ParseContext,
  rels: Map<string, Relationship>,
): Promise<SlideImageElement | undefined> {
  const spPr = firstNamed(pic, 'p:spPr');
  const xfrm = firstNamed(spPr, 'a:xfrm');
  const off = firstNamed(xfrm, 'a:off');
  const ext = firstNamed(xfrm, 'a:ext');
  if (!off || !ext) return undefined;
  const blipFill = firstNamed(pic, 'p:blipFill');
  const blip = firstDescendant(blipFill, 'a:blip');
  const embed = getAttr(blip, 'r:embed') ?? getAttr(blip, 'r:link');
  if (!embed) return undefined;
  const part = rels.get(embed)?.resolved;
  if (!part) return undefined;
  const src = await loadImage(ctx, part);
  if (!src) return undefined;
  const srcRect = firstNamed(blipFill, 'a:srcRect');
  const crop = srcRect
    ? {
        left: numberAttr(srcRect, 'l', 0) / 100_000,
        right: numberAttr(srcRect, 'r', 0) / 100_000,
        top: numberAttr(srcRect, 't', 0) / 100_000,
        bottom: numberAttr(srcRect, 'b', 0) / 100_000,
      }
    : undefined;
  return {
    id: uid('el'),
    type: 'image',
    src,
    part,
    name: getAttr(firstDescendant(pic, 'p:cNvPr'), 'name') ?? undefined,
    x: emuToPx(numberAttr(off, 'x')),
    y: emuToPx(numberAttr(off, 'y')),
    w: emuToPx(numberAttr(ext, 'cx')),
    h: emuToPx(numberAttr(ext, 'cy')),
    rotation: numberAttr(xfrm, 'rot') / 60_000 || undefined,
    flipH: boolAttr(xfrm, 'flipH') || undefined,
    flipV: boolAttr(xfrm, 'flipV') || undefined,
    crop,
  };
}

async function parseTable(
  frame: Element,
  ctx: ParseContext,
): Promise<SlideTableElement | undefined> {
  const xfrm = firstDescendant(frame, 'p:xfrm');
  const off = firstNamed(xfrm, 'a:off');
  const ext = firstNamed(xfrm, 'a:ext');
  const tbl = firstDescendant(frame, 'a:tbl');
  if (!tbl || !off || !ext) return undefined;
  const grid = firstNamed(tbl, 'a:tblGrid');
  const columnWidths = childrenNamed(grid, 'a:gridCol').map((col) => emuToPx(numberAttr(col, 'w')));
  const rows: SlideTableElement['rows'] = [];
  const rowHeights: number[] = [];
  for (const tr of childrenNamed(tbl, 'a:tr')) {
    rowHeights.push(emuToPx(numberAttr(tr, 'h', 0)));
    const cells: SlideTableElement['rows'][number]['cells'] = [];
    for (const tc of childrenNamed(tr, 'a:tc')) {
      const tcPr = firstNamed(tc, 'a:tcPr');
      const fill = parseFillElement(tcPr, ctx);
      const anchor = getAttr(tcPr, 'anchor');
      const paragraphs = parseParagraphs(
        firstNamed(tc, 'a:txBody'),
        ctx,
        { fontSizePt: 14, color: ctx.theme.dk1 ?? '#1D2939' },
      );
      const borderColor = resolveColor(firstNamed(tcPr, 'a:lnL'), ctx);
      const cell: SlideTableElement['rows'][number]['cells'][number] = {
        text: paragraphs,
        fill,
        borders: borderColor ? { color: borderColor, widthPx: 0.75 } : undefined,
        verticalAlign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
      };
      const hMerge = getAttr(tcPr, 'hMerge') === '1';
      const vMerge = vMergeCell(tcPr);
      const gridSpan = numberAttr(tcPr, 'gridSpan', 1);
      const rowSpan = numberAttr(tcPr, 'rowSpan', 1);
      if (hMerge) {
        const previous = cells.at(-1);
        if (previous) previous.colSpan = (previous.colSpan ?? 1) + 1;
      } else if (vMerge) {
        const above = rows.at(-1)?.cells[cells.length];
        if (above) above.rowSpan = (above.rowSpan ?? 1) + 1;
      } else {
        cell.colSpan = gridSpan > 1 ? gridSpan : undefined;
        cell.rowSpan = rowSpan > 1 ? rowSpan : undefined;
        cells.push(cell);
      }
    }
    rows.push({ cells });
  }
  return {
    id: uid('el'),
    type: 'table',
    x: emuToPx(numberAttr(off, 'x')),
    y: emuToPx(numberAttr(off, 'y')),
    w: emuToPx(numberAttr(ext, 'cx')),
    h: emuToPx(numberAttr(ext, 'cy')),
    columnWidths: columnWidths.length ? columnWidths : [],
    rowHeights,
    rows,
  };
}

function vMergeCell(tcPr: Element | null): boolean {
  return getAttr(tcPr, 'vMerge') === '1';
}

async function parseGroup(
  group: Element,
  ctx: ParseContext,
  rels: Map<string, Relationship>,
): Promise<SlideGroupElement | undefined> {
  const grpSpPr = firstNamed(group, 'p:grpSpPr');
  const xfrm = firstNamed(grpSpPr, 'a:xfrm');
  const off = firstNamed(xfrm, 'a:off');
  const ext = firstNamed(xfrm, 'a:ext');
  const chOff = firstNamed(xfrm, 'a:chOff');
  const chExt = firstNamed(xfrm, 'a:chExt');
  if (!off || !ext) return undefined;
  const x = emuToPx(numberAttr(off, 'x'));
  const y = emuToPx(numberAttr(off, 'y'));
  const w = emuToPx(numberAttr(ext, 'cx'));
  const h = emuToPx(numberAttr(ext, 'cy'));
  const childX = emuToPx(numberAttr(chOff, 'x', 0));
  const childY = emuToPx(numberAttr(chOff, 'y', 0));
  const childW = emuToPx(numberAttr(chExt, 'cx', numberAttr(ext, 'cx', 1)));
  const childH = emuToPx(numberAttr(chExt, 'cy', numberAttr(ext, 'cy', 1)));
  const scaleX = childW ? w / childW : 1;
  const scaleY = childH ? h / childH : 1;
  const children: SlideElement[] = [];
  for (const child of Array.from(group.childNodes)) {
    if (!(child instanceof Element)) continue;
    const parsed = await parseElement(child, ctx, new Map(), rels);
    for (const element of parsed) {
      element.x = (element.x - childX) * scaleX;
      element.y = (element.y - childY) * scaleY;
      element.w *= scaleX;
      element.h *= scaleY;
      children.push(element);
    }
  }
  return { id: uid('el'), type: 'group', x, y, w, h, children };
}

async function parseElement(
  element: Element,
  ctx: ParseContext,
  placeholders: Map<string, PlaceholderStyle>,
  rels: Map<string, Relationship>,
): Promise<SlideElement[]> {
  if (element.tagName === 'p:sp') {
    const parsed = parseShape(element, ctx, placeholders, defaultPlaceholderRunStyle(undefined, ctx));
    return parsed.element ? [parsed.element] : [];
  }
  if (element.tagName === 'p:pic') {
    const picture = await parsePicture(element, ctx, rels);
    return picture ? [picture] : [];
  }
  if (element.tagName === 'p:graphicFrame') {
    const table = await parseTable(element, ctx);
    return table ? [table] : [];
  }
  if (element.tagName === 'p:grpSp') {
    const group = await parseGroup(element, ctx, rels);
    return group ? [group] : [];
  }
  if (element.tagName === 'p:cxnSp') {
    const parsed = parseShape(element, ctx, placeholders, {});
    if (parsed.element && parsed.element.type === 'shape') {
      parsed.element.shape = 'line';
      parsed.element.fill = { type: 'none' };
      return [parsed.element];
    }
    return parsed.element ? [parsed.element] : [];
  }
  return [];
}

async function parseContainer(
  ctx: ParseContext,
  part: string,
): Promise<ParsedContainer> {
  const cached = ctx.containerCache.get(part);
  if (cached) return cached;
  const xml = await ctx.pkg.text(part);
  const doc = parseXmlDom(xml);
  const rels = await relationshipMap(ctx.pkg, part);
  ctx.relCache.set(part, rels);
  const root = firstDescendant(doc, part.includes('slideLayout') ? 'p:sldLayout' : part.includes('slideMaster') ? 'p:sldMaster' : 'p:sld') ?? doc.documentElement;
  const cSld = firstDescendant(root, 'p:cSld');
  const spTree = firstDescendant(cSld, 'p:spTree');
  const elements: SlideElement[] = [];
  const placeholders = new Map<string, PlaceholderStyle>();
  for (const child of Array.from(spTree?.childNodes ?? [])) {
    if (!(child instanceof Element)) continue;
    if (child.tagName === 'p:sp') {
      const parsed = parseShape(child, ctx, placeholders, {});
      if (parsed.element) elements.push(parsed.element);
      if (parsed.placeholder) placeholders.set(parsed.placeholder.key, parsed.placeholder.style);
    } else {
      const parsed = await parseElement(child, ctx, placeholders, rels);
      elements.push(...parsed);
    }
  }
  const bg = firstDescendant(cSld, 'p:bg');
  const background = parseFillElement(bg, ctx);
  const result: ParsedContainer = { elements, placeholders, background };
  ctx.containerCache.set(part, result);
  return result;
}

function isPlaceholderElement(element: SlideElement): boolean {
  return element.type === 'text' && !!element.placeholderType;
}

async function parseSlide(
  ctx: ParseContext,
  part: string,
  index: number,
  fallbackTitle: string,
): Promise<SlideModel> {
  const doc = parseXmlDom(await ctx.pkg.text(part));
  const rels = await relationshipMap(ctx.pkg, part);
  const root = firstDescendant(doc, 'p:sld') ?? doc.documentElement;
  const layoutRel = Array.from(rels.values()).find((rel) => rel.type.endsWith('/slideLayout'));
  const layout = layoutRel?.resolved ? await parseContainer(ctx, layoutRel.resolved) : undefined;
  const layoutRels = layoutRel?.resolved ? await relationshipMap(ctx.pkg, layoutRel.resolved) : new Map<string, Relationship>();
  const masterRel = Array.from(layoutRels.values()).find((rel) => rel.type.endsWith('/slideMaster'));
  const master = masterRel?.resolved ? await parseContainer(ctx, masterRel.resolved) : undefined;

  const inheritedPlaceholders = new Map(master?.placeholders ?? []);
  for (const [key, style] of layout?.placeholders ?? []) inheritedPlaceholders.set(key, style);

  const elements: SlideElement[] = [];
  const background =
    parseFillElement(firstDescendant(firstDescendant(root, 'p:cSld') ?? root, 'p:bg'), ctx) ??
    layout?.background ??
    master?.background ??
    ({ type: 'solid', color: '#FFFFFF' } as SlideFill);

  // Layout/master decorations that are not placeholders are shown behind the slide.
  for (const decoration of [...(master?.elements ?? []), ...(layout?.elements ?? [])]) {
    if (!isPlaceholderElement(decoration)) elements.push({ ...decoration });
  }

  const cSld = firstDescendant(root, 'p:cSld');
  const spTree = firstDescendant(cSld, 'p:spTree');
  const decorationCount = elements.length;
  let title = '';
  for (const child of Array.from(spTree?.childNodes ?? [])) {
    if (!(child instanceof Element)) continue;
    const parsed = await parseElement(child, ctx, inheritedPlaceholders, rels);
    for (const element of parsed) {
      elements.push(element);
      if (element.type === 'text' && (element.placeholderType === 'title' || element.placeholderType === 'ctrTitle')) {
        title = plainText(element.paragraphs);
      }
    }
  }

  if (!title) {
    let bestSize = 0;
    for (const element of elements.slice(decorationCount)) {
      if (element.type !== 'text') continue;
      for (const paragraph of element.paragraphs) {
        const text = paragraph.runs.map((run) => run.text).join('').trim();
        for (const run of paragraph.runs) {
          const size = run.style.fontSizePt ?? 0;
          if (text && size >= bestSize) {
            bestSize = size;
            title = text;
          }
        }
      }
    }
  }

  let notes: string | undefined;
  const notesRel = Array.from(rels.values()).find((rel) => rel.type.endsWith('/notesSlide'));
  if (notesRel?.resolved) {
    try {
      const notesDoc = parseXmlDom(await ctx.pkg.text(notesRel.resolved));
      const notesSpTree = firstDescendant(notesDoc, 'p:spTree');
      const lines: string[] = [];
      for (const sp of childrenNamed(notesSpTree, 'p:sp')) {
        const ph = firstDescendant(sp, 'p:ph');
        const type = getAttr(ph, 'type');
        if (type === 'sldNum' || type === 'hdr' || type === 'ftr' || type === 'dt') continue;
        const paragraphs = parseParagraphs(firstNamed(sp, 'p:txBody'), ctx, { fontSizePt: 14 });
        const text = plainText(paragraphs);
        if (text) lines.push(text);
      }
      if (lines.length) notes = lines.join('\n');
    } catch {
      notes = undefined;
    }
  }

  const hasAnimations = /<p:timing\b/.test(await ctx.pkg.text(part));
  const hasMedia = /<p:video\b|<p:audio\b/.test(await ctx.pkg.text(part));
  return {
    id: uid('slide'),
    index,
    title: title || `${fallbackTitle} ${index + 1}`,
    background,
    elements,
    notes,
    hasAnimations,
    hasMedia,
    layoutName: layoutRel?.target,
  };
}

export interface ParsePptxOptions {
  fileName?: string;
}

export async function parsePptxPackage(
  pkg: OoxmlPackage,
  options: ParsePptxOptions = {},
): Promise<SlideDeck> {
  const presentationXml = await pkg.text('ppt/presentation.xml');
  const presentation = parseXmlDom(presentationXml);
  const root = firstDescendant(presentation, 'p:presentation') ?? presentation.documentElement;
  const size = firstNamed(root, 'p:sldSz');
  const widthEmu = numberAttr(size, 'cx', 12_192_000);
  const heightEmu = numberAttr(size, 'cy', 6_858_000);
  const widthPx = emuToPx(widthEmu);
  const heightPx = emuToPx(heightEmu);
  const ratio = widthEmu / heightEmu;
  const aspect: SlideDeck['aspect'] =
    Math.abs(ratio - 16 / 9) < 0.03 ? '16:9' : Math.abs(ratio - 4 / 3) < 0.03 ? '4:3' : 'custom';

  const presentationRels = await relationshipMap(pkg, 'ppt/presentation.xml');
  const themeRel = Array.from(presentationRels.values()).find((rel) => rel.type.endsWith('/theme'));
  const themePart = themeRel?.resolved ?? pkg.names().find((name) => /^ppt\/theme\/theme\d+\.xml$/.test(name)) ?? '';
  const theme = themePart ? await readThemeColors(pkg, themePart) : {};
  const themeFonts = { major: '等线 Light', minor: '等线' };
  if (themePart) {
    try {
      const themeDoc = parseXmlDom(await pkg.text(themePart));
      const scheme = firstDescendant(themeDoc, 'a:fontScheme');
      const major = firstDescendant(firstNamed(scheme, 'a:majorFont'), 'a:ea') ?? firstDescendant(firstNamed(scheme, 'a:majorFont'), 'a:latin');
      const minor = firstDescendant(firstNamed(scheme, 'a:minorFont'), 'a:ea') ?? firstDescendant(firstNamed(scheme, 'a:minorFont'), 'a:latin');
      themeFonts.major = getAttr(major, 'typeface') || themeFonts.major;
      themeFonts.minor = getAttr(minor, 'typeface') || themeFonts.minor;
    } catch {
      // keep defaults
    }
  }

  const ctx: ParseContext = {
    pkg,
    theme,
    themeFonts,
    imageCache: new Map(),
    relCache: new Map(),
    containerCache: new Map(),
    slideWidth: widthPx,
    slideHeight: heightPx,
  };

  const slideIdList = firstNamed(root, 'p:sldIdLst');
  const slideParts: Array<{ part: string; id: string }> = [];
  for (const slideId of childrenNamed(slideIdList, 'p:sldId')) {
    const rid = getAttr(slideId, 'r:id');
    const part = rid ? presentationRels.get(rid)?.resolved : undefined;
    if (part) slideParts.push({ part, id: getAttr(slideId, 'id') ?? String(slideParts.length + 1) });
  }

  const slides: SlideModel[] = [];
  for (let index = 0; index < slideParts.length; index += 1) {
    slides.push(await parseSlide(ctx, slideParts[index].part, index, '幻灯片'));
  }

  const meta = await parsePresentationMeta(pkg);
  return {
    version: 1,
    widthEmu,
    heightEmu,
    widthPx,
    heightPx,
    aspect,
    slides,
    themeColors: theme,
    themeFonts,
    meta: { ...meta, title: meta.title ?? stripExtension(options.fileName ?? '演示文稿') },
    hasAnimations: slides.some((slide) => slide.hasAnimations),
    hasMedia: slides.some((slide) => slide.hasMedia),
    hasNotes: slides.some((slide) => !!slide.notes),
    stats: { slides: slides.length },
  };
}

async function parsePresentationMeta(pkg: OoxmlPackage): Promise<SlideDeck['meta']> {
  if (!pkg.has('docProps/core.xml')) return {};
  try {
    const doc = parseXmlDom(await pkg.text('docProps/core.xml'));
    const read = (tag: string): string | undefined => doc.getElementsByTagName(tag)[0]?.textContent ?? undefined;
    return {
      title: read('dc:title'),
      author: read('dc:creator'),
      created: read('dcterms:created'),
      modified: read('dcterms:modified'),
    };
  } catch {
    return {};
  }
}

export async function parsePptx(
  bytes: Uint8Array,
  options: ParsePptxOptions = {},
): Promise<SlideDeck> {
  const pkg = await loadOoxmlPackage(bytes);
  try {
    return await parsePptxPackage(pkg, options);
  } finally {
    pkg.dispose();
  }
}

export function deckStats(deck: SlideDeck): { slides: number; words: number } {
  let words = 0;
  for (const slide of deck.slides) {
    for (const element of slide.elements) {
      if (element.type === 'text') {
        words += countWords(plainText(element.paragraphs));
      } else if (element.type === 'shape' && element.paragraphs) {
        words += countWords(plainText(element.paragraphs));
      }
    }
    if (slide.notes) words += countWords(slide.notes);
  }
  return { slides: deck.slides.length, words };
}
