import JSZip from 'jszip';
import {
  type DocBlock,
  type DocImage,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocTable,
  type DocTableCell,
  type PageSetup,
  type ParagraphStyle,
  type RunStyle,
  type TableBorderSide,
  type TableBorders,
} from '@mini-office/document-model';
import {
  escapeXml,
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  parseXmlDom,
  pxToTwips,
  resolvePartPath,
  xmlDeclaration,
  type OoxmlPackage,
} from '@mini-office/ooxml';
import { extensionOf, uid } from '@mini-office/shared';
import { documentSections } from './sections';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

interface RelationshipInfo {
  id: string;
  type: string;
  target: string;
}

interface SerializeState {
  zip: JSZip;
  sourceExists: boolean;
  documentRels: RelationshipInfo[];
  relsChanged: boolean;
  contentTypesXml?: string;
  contentTypesChanged: boolean;
  addedImages: Map<string, string>;
  imageCounter: number;
  hasNumbering: boolean;
  sectPrXml?: string;
}

const IMAGE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
};

function attr(value: string | number | undefined, name: string): string {
  if (value === undefined || value === '') return '';
  return ` ${name}="${escapeXml(String(value))}"`;
}

function boolTag(tag: string, value: boolean | undefined, valueAttribute = true): string {
  if (value === undefined) return '';
  if (!value) return valueAttribute ? `<w:${tag} w:val="0"/>` : '';
  return valueAttribute ? `<w:${tag}/>` : `<w:${tag}/>`;
}

function runProperties(style: RunStyle): string {
  const parts: string[] = [];
  if (style.fontFamily) {
    parts.push(
      `<w:rFonts w:ascii="${escapeXml(style.fontFamily)}" w:hAnsi="${escapeXml(style.fontFamily)}" w:eastAsia="${escapeXml(style.fontFamily)}" w:cs="${escapeXml(style.fontFamily)}"/>`,
    );
  }
  if (style.bold !== undefined) parts.push(`<w:b${style.bold ? '' : ' w:val="0"'}/>`);
  if (style.italic !== undefined) parts.push(`<w:i${style.italic ? '' : ' w:val="0"'}/>`);
  if (style.underline !== undefined) {
    parts.push(`<w:u w:val="${style.underline ? 'single' : 'none'}"/>`);
  }
  if (style.strike !== undefined) parts.push(`<w:strike${style.strike ? '' : ' w:val="0"'}/>`);
  if (style.color) parts.push(`<w:color w:val="${style.color.replace('#', '')}"/>`);
  if (style.highlight) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${style.highlight.replace('#', '')}"/>`);
  }
  if (style.fontSizePt) {
    const halfPoints = Math.max(2, Math.round(style.fontSizePt * 2));
    parts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  if (style.verticalAlign === 'super') parts.push('<w:vertAlign w:val="superscript"/>');
  if (style.verticalAlign === 'sub') parts.push('<w:vertAlign w:val="subscript"/>');
  if (style.letterSpacingPx) {
    parts.push(`<w:spacing w:val="${Math.round(style.letterSpacingPx * 15)}"/>`);
  }
  if (style.allCaps) parts.push('<w:caps/>');
  if (style.smallCaps) parts.push('<w:smallCaps/>');
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

function paragraphProperties(
  paragraph: DocParagraph,
  state: SerializeState,
  includeNumbering = true,
): string {
  const style = paragraph.style;
  const parts: string[] = [];
  if (paragraph.styleId) parts.push(`<w:pStyle w:val="${escapeXml(paragraph.styleId)}"/>`);
  if (includeNumbering && paragraph.list && state.hasNumbering && paragraph.list.numId) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${paragraph.list.level}"/><w:numId w:val="${escapeXml(paragraph.list.numId)}"/></w:numPr>`,
    );
  }
  if (style.keepNext) parts.push('<w:keepNext/>');
  if (style.pageBreakBefore) parts.push('<w:pageBreakBefore/>');
  if (style.alignment) {
    const map: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    parts.push(`<w:jc w:val="${map[style.alignment] ?? 'left'}"/>`);
  }
  const indentParts: string[] = [];
  if (style.indentLeftPx) indentParts.push(`w:left="${Math.round(style.indentLeftPx * 15)}"`);
  if (style.indentRightPx) indentParts.push(`w:right="${Math.round(style.indentRightPx * 15)}"`);
  if (style.firstLineIndentPx !== undefined) {
    if (style.firstLineIndentPx >= 0) {
      indentParts.push(`w:firstLine="${Math.round(style.firstLineIndentPx * 15)}"`);
    } else {
      indentParts.push(`w:hanging="${Math.round(Math.abs(style.firstLineIndentPx) * 15)}"`);
    }
  }
  if (indentParts.length) parts.push(`<w:ind ${indentParts.join(' ')}/>`);
  const spacingParts: string[] = [];
  if (style.spaceBeforePx !== undefined) spacingParts.push(`w:before="${Math.round(style.spaceBeforePx * 15)}"`);
  if (style.spaceAfterPx !== undefined) spacingParts.push(`w:after="${Math.round(style.spaceAfterPx * 15)}"`);
  if (style.lineHeightExactPx !== undefined) {
    spacingParts.push(`w:line="${Math.round(style.lineHeightExactPx * 15)}"`);
    spacingParts.push('w:lineRule="exact"');
  } else if (style.lineHeight !== undefined) {
    spacingParts.push(`w:line="${Math.round(style.lineHeight * 240)}"`);
    spacingParts.push('w:lineRule="auto"');
  }
  if (spacingParts.length) parts.push(`<w:spacing ${spacingParts.join(' ')}/>`);
  if (style.backgroundColor) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${style.backgroundColor.replace('#', '')}"/>`);
  }
  if (style.outlineLevel !== undefined) parts.push(`<w:outlineLvl w:val="${style.outlineLevel}"/>`);
  return parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : '';
}

function borderXml(tag: string, border: TableBorderSide | undefined): string {
  if (!border) return '';
  if (border.style === 'none' || border.widthPx === 0) return `<w:${tag} w:val="nil"/>`;
  const top = Math.max(2, Math.round(border.widthPx * 6));
  const color = (border.color ?? '#000000').replace('#', '');
  return `<w:${tag} w:val="${border.style === 'double' ? 'double' : border.style}" w:sz="${top}" w:space="0" w:color="${color}"/>`;
}

function tableBordersXml(borders: TableBorders | undefined): string {
  if (!borders) return '';
  const content = [
    borderXml('top', borders.top),
    borderXml('left', borders.left),
    borderXml('bottom', borders.bottom),
    borderXml('right', borders.right),
    borderXml('insideH', borders.insideH),
    borderXml('insideV', borders.insideV),
  ].join('');
  return content ? `<w:tblBorders>${content}</w:tblBorders>` : '';
}

function cellBordersXml(borders: TableBorders | undefined): string {
  if (!borders) return '';
  const content = [
    borderXml('top', borders.top),
    borderXml('left', borders.left),
    borderXml('bottom', borders.bottom),
    borderXml('right', borders.right),
  ].join('');
  return content ? `<w:tcBorders>${content}</w:tcBorders>` : '';
}

function dataUrlToBytes(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(dataUrl);
  if (!match) return null;
  const mime = match[1] || 'application/octet-stream';
  const isBase64 = !!match[2];
  try {
    if (isBase64) {
      const binary = atob(match[3]);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
      return { bytes, mime };
    }
    const decoded = decodeURIComponent(match[3]);
    const bytes = new TextEncoder().encode(decoded);
    return { bytes, mime };
  } catch {
    return null;
  }
}

async function blobUrlToBytes(url: string): Promise<{ bytes: Uint8Array; mime: string } | null> {
  try {
    const response = await fetch(url);
    const blob = await response.blob();
    const buffer = await blob.arrayBuffer();
    return { bytes: new Uint8Array(buffer), mime: blob.type || 'image/png' };
  } catch {
    return null;
  }
}

function existingRels(xml: string | undefined): RelationshipInfo[] {
  if (!xml) return [];
  try {
    const doc = parseXmlDom(xml);
    return Array.from(doc.getElementsByTagName('Relationship')).map((element) => ({
      id: getAttrLocal(element, 'Id') ?? '',
      type: getAttrLocal(element, 'Type') ?? '',
      target: getAttrLocal(element, 'Target') ?? '',
    }));
  } catch {
    return [];
  }
}

async function ensureImageRelationship(
  state: SerializeState,
  image: DocImage,
): Promise<string | undefined> {
  const cacheKey = image.id ?? image.relId ?? image.part ?? image.dataUrl?.slice(0, 80) ?? image.src.slice(0, 120);
  const cached = state.addedImages.get(cacheKey);
  if (cached) return cached;

  if (image.relId) {
    const existing = state.documentRels.find((rel) => rel.id === image.relId);
    if (existing) {
      state.addedImages.set(cacheKey, existing.id);
      return existing.id;
    }
  }

  if (image.part) {
    const target = resolvePartPath('word/document.xml', image.part) ?? image.part;
    const match = state.documentRels.find((rel) => {
      const resolved = resolvePartPath('word/document.xml', rel.target) ?? rel.target;
      return resolved === target;
    });
    if (match) {
      state.addedImages.set(cacheKey, match.id);
      return match.id;
    }
  }

  let payload: { bytes: Uint8Array; mime: string } | null = null;
  if (image.dataUrl) payload = dataUrlToBytes(image.dataUrl);
  else if (image.src.startsWith('blob:') || image.src.startsWith('http')) {
    payload = await blobUrlToBytes(image.src);
  }
  if (!payload) return undefined;

  const ext =
    IMAGE_MIME_EXT[payload.mime] ??
    (extensionOf(image.part ?? '').replace('.', '') || 'png');
  state.imageCounter += 1;
  const fileName = `mini-office-${state.imageCounter}.${ext}`;
  const mediaPath = `word/media/${fileName}`;
  state.zip.file(mediaPath, payload.bytes);
  const relId = `rIdMo${state.imageCounter}`;
  state.documentRels.push({
    id: relId,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    target: `media/${fileName}`,
  });
  state.relsChanged = true;
  state.contentTypesChanged = true;
  state.addedImages.set(cacheKey, relId);
  return relId;
}

function drawingXml(image: DocImage, relId: string, index: number): string {
  const widthEmu = Math.round(image.widthPx * 9525);
  const heightEmu = Math.round(image.heightPx * 9525);
  const name = escapeXml(image.name ?? `图片 ${index + 1}`);
  const descr = escapeXml(image.alt ?? image.name ?? '');
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${index + 1}" name="${name}" descr="${descr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${A_NS}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${PIC_NS}"><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="${name}" descr="${descr}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

async function runXml(run: DocRun, state: SerializeState): Promise<string> {
  if (run.image) {
    const relId = await ensureImageRelationship(state, run.image);
    if (!relId) return '';
    const properties = runProperties(run.style);
    return `<w:r>${properties}${drawingXml(run.image, relId, state.imageCounter)}</w:r>`;
  }
  const properties = runProperties(run.style);
  let content = '';
  if (run.tab) content += '<w:tab/>';
  if (run.lineBreak) content += '<w:br/>';
  if (run.text) content += `<w:t xml:space="preserve">${escapeXml(run.text)}</w:t>`;
  if (!content && !properties) content = '<w:t xml:space="preserve"></w:t>';
  const core = `<w:r>${properties}${content}</w:r>`;
  // 功能：模板字段保存为原生 MERGEFIELD，字段名安全转义；时间：2026-10-24；作者：lq。
  if (run.mergeField) {
    const name = run.mergeField.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    return `<w:fldSimple w:instr="${escapeXml(` MERGEFIELD "${name}" \\* MERGEFORMAT `)}">${core}</w:fldSimple>`;
  }
  if (run.field === 'PAGE' || run.field === 'NUMPAGES' || run.field === 'SECTIONPAGES') {
    const instruction = ` ${run.field} \\* MERGEFORMAT `;
    return `<w:fldSimple w:instr="${instruction}">${core}</w:fldSimple>`;
  }
  return core;
}

async function paragraphXml(
  paragraph: DocParagraph,
  state: SerializeState,
  options: { isInTable?: boolean; forceNumberingAsText?: boolean } = {},
): Promise<string> {
  const includeNumbering = !options.forceNumberingAsText;
  const pPr = paragraphProperties(paragraph, state, includeNumbering);
  let prefix = '';
  if (
    paragraph.list &&
    (!state.hasNumbering || !paragraph.list.numId || options.forceNumberingAsText)
  ) {
    prefix = `<w:r><w:t xml:space="preserve">${escapeXml(paragraph.list.marker)} </w:t></w:r>`;
  }
  const runs = await Promise.all(paragraph.runs.map((run) => runXml(run, state)));
  return `<w:p>${pPr}${prefix}${runs.join('')}</w:p>`;
}

async function cellXml(
  cell: DocTableCell,
  state: SerializeState,
  columnWidthPx: number | undefined,
  continuation = false,
): Promise<string> {
  const props: string[] = [];
  // 功能：以网格列宽优先保存合并区域总宽度，避免旧 tcW 覆盖尺寸调整；时间：2026-10-19；作者：lq。
  const width = columnWidthPx ?? cell.widthPx;
  if (width) props.push(`<w:tcW w:w="${Math.round(width * 15)}" w:type="dxa"/>`);
  if (cell.colSpan > 1) props.push(`<w:gridSpan w:val="${cell.colSpan}"/>`);
  if (continuation) props.push('<w:vMerge/>');
  else if (cell.rowSpan > 1) props.push('<w:vMerge w:val="restart"/>');
  if (cell.backgroundColor) {
    props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cell.backgroundColor.replace('#', '')}"/>`);
  }
  if (cell.verticalAlign) {
    const map = { top: 'top', middle: 'center', bottom: 'bottom' } as const;
    props.push(`<w:vAlign w:val="${map[cell.verticalAlign]}"/>`);
  }
  if (cell.borders) {
    const tcBorders = cellBordersXml(cell.borders);
    if (tcBorders) props.push(tcBorders);
  }
  const content = continuation
    ? ''
    : (await Promise.all(cell.blocks.map((block) => blockXml(block, state)))).join('');
  return `<w:tc><w:tcPr>${props.join('')}</w:tcPr>${content || '<w:p/>'}</w:tc>`;
}

async function tableXml(table: DocTable, state: SerializeState): Promise<string> {
  const props: string[] = [];
  if (table.styleId) props.push(`<w:tblStyle w:val="${escapeXml(table.styleId)}"/>`);
  if (table.widthPct) {
    props.push(`<w:tblW w:w="${Math.round(table.widthPct * 50)}" w:type="pct"/>`);
  } else {
    props.push('<w:tblW w:w="0" w:type="auto"/>');
  }
  if (table.alignment) {
    const map = { left: 'left', center: 'center', right: 'right', justify: 'left' } as const;
    props.push(`<w:jc w:val="${map[table.alignment]}"/>`);
  }
  // 功能：固定网格布局确保 Office 按用户设置的列宽打开表格；时间：2026-10-19；作者：lq。
  props.push('<w:tblLayout w:type="fixed"/>');
  const borders = tableBordersXml(table.borders);
  if (borders) props.push(borders);
  const grid = table.columnWidthsPx
    .map((width) => `<w:gridCol w:w="${Math.round(width * 15)}"/>`)
    .join('');
  const columnCount = Math.max(
    table.columnWidthsPx.length,
    ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + cell.colSpan, 0)),
  );
  const active: Array<{ start: number; span: number; remaining: number }> = [];
  const rows: string[] = [];
  for (const row of table.rows) {
    const trProps: string[] = [];
    if (row.heightPx) trProps.push(`<w:trHeight w:val="${Math.round(row.heightPx * 15)}" w:hRule="atLeast"/>`);
    if (row.isHeader) trProps.push('<w:tblHeader/>');
    const cellsXml: string[] = [];
    let column = 0;
    let cellIndex = 0;
    while (column < columnCount || cellIndex < row.cells.length) {
      const spanning = active.find((item) => item.start === column && item.remaining > 0);
      if (spanning) {
        cellsXml.push(
          await cellXml(
            {
              id: `continue-${column}`,
              blocks: [],
              colSpan: spanning.span,
              rowSpan: 1,
            },
            state,
            table.columnWidthsPx.slice(column, column + spanning.span).reduce((sum, width) => sum + width, 0),
            true,
          ),
        );
        spanning.remaining -= 1;
        column += spanning.span;
        continue;
      }
      const cell = row.cells[cellIndex];
      if (!cell) break;
      cellsXml.push(await cellXml(cell, state, table.columnWidthsPx.slice(column, column + cell.colSpan).reduce((sum, width) => sum + width, 0)));
      if (cell.rowSpan > 1) {
        active.push({ start: column, span: cell.colSpan, remaining: cell.rowSpan - 1 });
      }
      column += cell.colSpan;
      cellIndex += 1;
    }
    for (const item of active) {
      if (item.remaining <= 0) item.remaining = 0;
    }
    rows.push(`<w:tr>${trProps.length ? `<w:trPr>${trProps.join('')}</w:trPr>` : ''}${cellsXml.join('')}</w:tr>`);
  }
  return `<w:tbl><w:tblPr>${props.join('')}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>`;
}

async function blockXml(block: DocBlock, state: SerializeState): Promise<string> {
  if (block.type === 'paragraph') return paragraphXml(block, state);
  if (block.type === 'table') return tableXml(block, state);
  // 功能：分页标记只占最小行高，保持长表分页与预览一致；时间：2026-10-24；作者：lq。
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>';
}

function buildSectPr(page: PageSetup, state: SerializeState): string {
  const width = Math.round(pxToTwips(page.widthPx));
  const height = Math.round(pxToTwips(page.heightPx));
  const margin = page.margins;
  const pgSz = `<w:pgSz w:w="${width}" w:h="${height}"${page.orientation === 'landscape' ? ' w:orient="landscape"' : ''}/>`;
  const pgMar = `<w:pgMar w:top="${Math.round(pxToTwips(margin.top))}" w:right="${Math.round(pxToTwips(margin.right))}" w:bottom="${Math.round(pxToTwips(margin.bottom))}" w:left="${Math.round(pxToTwips(margin.left))}" w:header="${Math.round(pxToTwips(margin.header))}" w:footer="${Math.round(pxToTwips(margin.footer))}" w:gutter="${Math.round(pxToTwips(margin.gutter ?? 0))}"/>`;
  if (state.sectPrXml) {
    try {
      const doc = parseXmlDom(
        `<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${state.sectPrXml}</w:root>`,
      );
      const sectPr = firstDescendant(doc, 'w:sectPr');
      if (sectPr) {
        const size = firstChild(sectPr, 'w:pgSz') ?? doc.createElementNS(W_NS, 'w:pgSz');
        size.setAttribute('w:w', String(width));
        size.setAttribute('w:h', String(height));
        if (page.orientation === 'landscape') size.setAttribute('w:orient', 'landscape');
        else size.removeAttribute('w:orient');
        if (!size.parentNode) sectPr.insertBefore(size, sectPr.firstChild);
        const margins = firstChild(sectPr, 'w:pgMar') ?? doc.createElementNS(W_NS, 'w:pgMar');
        margins.setAttribute('w:top', String(Math.round(pxToTwips(margin.top))));
        margins.setAttribute('w:right', String(Math.round(pxToTwips(margin.right))));
        margins.setAttribute('w:bottom', String(Math.round(pxToTwips(margin.bottom))));
        margins.setAttribute('w:left', String(Math.round(pxToTwips(margin.left))));
        margins.setAttribute('w:header', String(Math.round(pxToTwips(margin.header))));
        margins.setAttribute('w:footer', String(Math.round(pxToTwips(margin.footer))));
        margins.setAttribute('w:gutter', String(Math.round(pxToTwips(margin.gutter ?? 0))));
        if (!margins.parentNode) sectPr.appendChild(margins);
        // 功能：写回起始页码并保留其他页码格式属性；时间：2026-10-24；作者：lq。
        if (page.pageNumberStart !== undefined) {
          const numbering = firstChild(sectPr, 'w:pgNumType') ?? doc.createElementNS(W_NS, 'w:pgNumType');
          numbering.setAttribute('w:start', String(page.pageNumberStart));
          if (!numbering.parentNode) sectPr.appendChild(numbering);
        }
        return new XMLSerializer().serializeToString(sectPr);
      }
    } catch {
      // Fall back to a generated section.
    }
  }
  return `<w:sectPr>${pgSz}${pgMar}${page.pageNumberStart !== undefined ? `<w:pgNumType w:start="${page.pageNumberStart}"/>` : ''}${page.columns && page.columns > 1 ? `<w:cols w:num="${page.columns}"/>` : ''}</w:sectPr>`;
}

/** 功能：保存新增/编辑后的默认页眉页脚，写入独立部件、关系和内容类型；时间：2026-10-24；作者：lq。 */
async function writeHeaderFooters(model: DocModel, state: SerializeState, suffix = ''): Promise<void> {
  const sectionDoc = parseXmlDom(`<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${state.sectPrXml ?? buildSectPr(model.page, state)}</w:root>`);
  const section = firstDescendant(sectionDoc, 'w:sectPr')!;
  for (const kind of ['header', 'footer'] as const) {
    const value = model[kind];
    if (!value || (state.sourceExists && !value.edited)) continue;
    // 每个部件拥有独立关系作用域；新建文本内容不复用原图片关系；时间：2026-10-24；作者：lq。
    const target = `mini-${kind}${suffix}.xml`;
    const part = `word/${target}`;
    const sourceRelPath = value.sourcePart ? `word/_rels/${value.sourcePart.split('/').at(-1)}.rels` : '';
    const partRels = sourceRelPath ? await state.zip.file(sourceRelPath)?.async('text') : undefined;
    const partState: SerializeState = { ...state, documentRels: existingRels(partRels), relsChanged: false, addedImages: new Map() };
    let body = '';
    for (const block of value.blocks) body += await blockXml(block, partState);
    state.imageCounter = partState.imageCounter;
    const root = kind === 'header' ? 'hdr' : 'ftr';
    state.zip.file(part, `${xmlDeclaration()}<w:${root} xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">${body || '<w:p/>'}</w:${root}>`);
    if (partState.documentRels.length) state.zip.file(`word/_rels/${target}.rels`, documentRelsXml(partState.documentRels));
    let rel = state.documentRels.find((item) => item.target === target && item.type.endsWith(`/${kind}`));
    if (!rel) {
      rel = { id: uid(`rId${kind}`), type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, target };
      state.documentRels.push(rel);
    }
    state.relsChanged = true;
    const references = Array.from(section.children).filter((element) => element.tagName === `w:${kind}Reference` && element.getAttribute('w:type') === 'default');
    references.forEach((element) => element.remove());
    const reference = sectionDoc.createElementNS(W_NS, `w:${kind}Reference`);
    reference.setAttribute('w:type', 'default');
    reference.setAttributeNS(R_NS, 'r:id', rel.id);
    section.insertBefore(reference, section.firstChild);
    const ct = parseXmlDom(state.contentTypesXml ?? contentTypesXml(false, false, false));
    const types = ct.documentElement;
    if (!Array.from(types.children).some((element) => element.getAttribute('PartName') === `/${part}`)) {
      const override = ct.createElementNS(types.namespaceURI, 'Override');
      override.setAttribute('PartName', `/${part}`);
      override.setAttribute('ContentType', `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
      types.appendChild(override);
    }
    state.contentTypesXml = updateContentTypes(new XMLSerializer().serializeToString(ct), new Set(partState.documentRels.filter((item) => item.type.endsWith('/image')).map((item) => extensionOf(item.target).replace('.', ''))));
    state.contentTypesChanged = true;
  }
  state.sectPrXml = new XMLSerializer().serializeToString(section);
}

function buildDocumentXml(body: string, page: PageSetup, state: SerializeState): string {
  return `${xmlDeclaration()}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" mc:Ignorable="w14 wp14"><w:body>${body}${buildSectPr(page, state)}</w:body></w:document>`;
}

function defaultStylesXml(doc: DocModel): string {
  const font = escapeXml(doc.defaultFont);
  const halfPoints = Math.max(2, Math.round(doc.defaultFontSizePt * 2));
  return `${xmlDeclaration()}<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
}

function corePropertiesXml(doc: DocModel): string {
  const now = new Date().toISOString();
  const title = escapeXml(doc.meta.title ?? doc.meta.author ?? 'Mini Office 文档');
  const author = escapeXml(doc.meta.author ?? 'Mini Office');
  return `${xmlDeclaration()}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${title}</dc:title><dc:creator>${author}</dc:creator><cp:lastModifiedBy>Mini Office</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${doc.meta.created ?? now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${doc.meta.modified ?? now}</dcterms:modified></cp:coreProperties>`;
}

function appPropertiesXml(): string {
  return `${xmlDeclaration()}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Mini Office</Application><AppVersion>0.1</AppVersion></Properties>`;
}

function contentTypesXml(hasJpeg: boolean, hasPng: boolean, hasGif: boolean): string {
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    hasPng ? '<Default Extension="png" ContentType="image/png"/>' : '',
    hasJpeg ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : '',
    hasGif ? '<Default Extension="gif" ContentType="image/gif"/>' : '',
  ].join('');
  return `${xmlDeclaration()}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
}

function packageRelsXml(): string {
  return `${xmlDeclaration()}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

function documentRelsXml(rels: RelationshipInfo[]): string {
  const items = rels
    .map(
      (rel) =>
        `<Relationship Id="${escapeXml(rel.id)}" Type="${escapeXml(rel.type)}" Target="${escapeXml(rel.target)}"/>`,
    )
    .join('');
  return `${xmlDeclaration()}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

function settingsXml(): string {
  return `${xmlDeclaration()}<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="420"/><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
}

function updateContentTypes(xml: string, extensions: Set<string>): string {
  const typeMap: Record<string, string> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    jpg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
  };
  let next = xml;
  const appended: string[] = [];
  for (const extension of extensions) {
    const lower = extension.toLowerCase();
    if (new RegExp(`<Default[^>]*Extension="${lower}"`, 'i').test(next)) continue;
    const type = typeMap[lower];
    if (!type) continue;
    appended.push(`<Default Extension="${lower}" ContentType="${type}"/>`);
  }
  if (appended.length === 0) return xml;
  if (next.includes('</Types>')) next = next.replace('</Types>', `${appended.join('')}</Types>`);
  return next;
}

function collectSectPrXml(documentXml: string | undefined): string | undefined {
  if (!documentXml) return undefined;
  try {
    const doc = parseXmlDom(documentXml);
    const sections = doc.getElementsByTagName('w:sectPr');
    const section = sections[sections.length - 1];
    return section ? new XMLSerializer().serializeToString(section) : undefined;
  } catch {
    return undefined;
  }
}

function decimalToHexByte(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0');
}

function ensurePngPlaceholder(): string {
  // 1×1 transparent PNG for pathological empty image references.
  return `data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==`;
}

export interface SerializeDocxOptions {
  sourceBytes?: Uint8Array;
  name: string;
  /** 功能：批量报告按节合并，生成结果移除数据源连接；时间：2026-10-24；作者：lq。 */
  sections?: DocModel[];
  materialized?: boolean;
}

export async function serializeDocx(
  doc: DocModel,
  options: SerializeDocxOptions,
): Promise<Uint8Array> {
  const zip = options.sourceBytes
    ? await JSZip.loadAsync(options.sourceBytes)
    : new JSZip();
  const sourceExists = !!options.sourceBytes;
  const documentXml = sourceExists ? await zip.file('word/document.xml')?.async('text') : undefined;
  const relsXml = sourceExists ? await zip.file('word/_rels/document.xml.rels')?.async('text') : undefined;
  const ctXml = sourceExists ? await zip.file('[Content_Types].xml')?.async('text') : undefined;

  const state: SerializeState = {
    zip,
    sourceExists,
    documentRels: existingRels(relsXml),
    relsChanged: !sourceExists,
    contentTypesXml: ctXml,
    contentTypesChanged: !sourceExists,
    addedImages: new Map(),
    imageCounter: 0,
    hasNumbering: sourceExists && zip.file('word/numbering.xml') !== null,
    sectPrXml: collectSectPrXml(documentXml),
  };

  // 功能：每条记录对应独立节，分别设置页眉页脚及起始页码；时间：2026-10-24；作者：lq。
  const sections = options.sections ?? documentSections(doc);
  const originalSection = state.sectPrXml;
  let body = '';
  for (let index = 0; index < sections.length; index++) {
    const current = sections[index];
    state.sectPrXml = originalSection;
    const sectionDoc = sections.length > 1 ? { ...current, header: current.header ? { ...current.header, edited: true } : undefined, footer: current.footer ? { ...current.footer, edited: true } : undefined } : current;
    await writeHeaderFooters(sectionDoc, state, options.sections || sections.length > 1 ? `-${index + 1}` : '');
    let content = '';
    const blocks = [...current.blocks];
    while (blocks.at(-1)?.type === 'pageBreak') blocks.pop();
    for (const block of blocks) content += await blockXml(block, state);
    if (index < sections.length - 1) {
      const sectionDom = parseXmlDom(`<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${buildSectPr(current.page, state)}</w:root>`);
      const section = firstDescendant(sectionDom, 'w:sectPr')!;
      const type = firstChild(section, 'w:type') ?? sectionDom.createElementNS(W_NS, 'w:type');
      type.setAttribute('w:val', 'nextPage');
      if (!type.parentNode) section.insertBefore(type, firstChild(section, 'w:pgSz'));
      const sectionXml = new XMLSerializer().serializeToString(section);
      // 在正文段落内附加节属性；末块为表格时使用零高度节标记，避免额外空白页。
      if (blocks.at(-1)?.type === 'paragraph') {
        const lastStart = content.lastIndexOf('<w:p>');
        const last = content.slice(lastStart);
        const withSection = last.includes('</w:pPr>') ? last.replace('</w:pPr>', `${sectionXml}</w:pPr>`) : last.replace('<w:p>', `<w:p><w:pPr>${sectionXml}</w:pPr>`);
        content = content.slice(0, lastStart) + withSection;
      } else content += `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>${sectionXml}</w:pPr></w:p>`;
    }
    body += content;
  }
  const documentContent = buildDocumentXml(body, sections.at(-1)?.page ?? doc.page, state);
  zip.file('word/document.xml', documentContent);

  if (!sourceExists) {
    zip.file('[Content_Types].xml', contentTypesXml(false, false, false));
    zip.file('_rels/.rels', packageRelsXml());
    zip.file('word/styles.xml', defaultStylesXml(doc));
    zip.file('word/settings.xml', settingsXml());
    zip.file('docProps/core.xml', corePropertiesXml(doc));
    zip.file('docProps/app.xml', appPropertiesXml());
    if (!state.documentRels.some((rel) => rel.type.endsWith('/styles'))) {
      state.documentRels.push({
        id: 'rIdStyles',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
        target: 'styles.xml',
      });
    }
    if (!state.documentRels.some((rel) => rel.type.endsWith('/settings'))) {
      state.documentRels.push({
        id: 'rIdSettings',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
        target: 'settings.xml',
      });
    }
    state.relsChanged = true;
  }

  // 功能：成品不再连接外部合并数据源，打开时刷新分页域；时间：2026-10-24；作者：lq。
  if (options.materialized || options.sections) {
    const settings = parseXmlDom(await zip.file('word/settings.xml')?.async('text') ?? settingsXml());
    Array.from(settings.getElementsByTagName('w:mailMerge')).forEach((element) => element.remove());
    const update = firstChild(settings.documentElement, 'w:updateFields') ?? settings.createElementNS(W_NS, 'w:updateFields');
    update.setAttribute('w:val', 'true');
    if (!update.parentNode) settings.documentElement.appendChild(update);
    zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
  }

  if (state.relsChanged || !zip.file('word/_rels/document.xml.rels')) {
    zip.file('word/_rels/document.xml.rels', documentRelsXml(state.documentRels));
  }
  if (state.contentTypesChanged) {
    const extensions = new Set<string>();
    for (const rel of state.documentRels) {
      if (!rel.type.endsWith('/image')) continue;
      const extension = extensionOf(rel.target).replace('.', '').toLowerCase();
      if (extension) extensions.add(extension);
    }
    const base = state.contentTypesXml ?? contentTypesXml(false, false, false);
    zip.file('[Content_Types].xml', updateContentTypes(base, extensions));
  }

  if (!sourceExists) {
    zip.file(
      '[Content_Types].xml',
      updateContentTypes(
        await zip.file('[Content_Types].xml')!.async('text'),
        new Set(['png', 'jpeg', 'gif']),
      ),
    );
  }

  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

export { ensurePngPlaceholder, decimalToHexByte };
