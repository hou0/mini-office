// 功能：读写 XLSX 标准 DrawingML 图片、关系与媒体文件；时间：2026-10-24；作者：lq。
import type JSZip from 'jszip';
import type { SheetImage, SheetModel } from '@mini-office/document-model';
import { escapeXml, getAttrLocal, mimeForPath, parseXmlDom, relationshipMap, type OoxmlPackage } from '@mini-office/ooxml';
const EMU = 9525;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';

// 功能：按 XML 本地名称匹配 DrawingML 节点，兼容 xdr/a 及其他命名空间前缀；时间：2026-10-24；作者：lq。
function firstChild(parent: Element | null, name: string): Element | null {
  return Array.from(parent?.children ?? []).find((item) => item.localName === name) ?? null;
}
function firstDescendant(parent: Element | null, name: string): Element | null {
  return Array.from(parent?.getElementsByTagName('*') ?? []).find((item) => item.localName === name) ?? null;
}

// 功能：恢复单单元格、双单元格和绝对坐标锚定的内嵌图片；时间：2026-10-24；作者：lq。
export async function parseSheetImages(pkg: OoxmlPackage, worksheetPart: string, worksheet: Element, sheet: SheetModel): Promise<SheetImage[]> {
  const drawingNode = firstChild(worksheet, 'drawing');
  if (!drawingNode) return [];
  const rels = await relationshipMap(pkg, worksheetPart);
  const part = rels.get(getAttrLocal(drawingNode, 'id') ?? '')?.resolved;
  if (!part || !pkg.has(part)) return [];
  const drawing = parseXmlDom(await pkg.text(part));
  const drawingRels = await relationshipMap(pkg, part);
  const images: SheetImage[] = [];
  const number = (node: Element | null, name: string) => Number(firstChild(node, name)?.textContent ?? 0);
  for (const anchor of Array.from(drawing.documentElement.children)) {
    const pic = firstChild(anchor, 'pic');
    const blip = firstDescendant(pic, 'blip');
    const media = drawingRels.get(getAttrLocal(blip, 'embed') ?? '')?.resolved;
    if (!pic || !media || !pkg.has(media)) continue;
    const from = firstChild(anchor, 'from');
    const to = firstChild(anchor, 'to');
    const pos = firstChild(anchor, 'pos');
    const col = Math.max(0, number(from, 'col'));
    const row = Math.max(0, number(from, 'row'));
    const offsetX = from ? number(from, 'colOff') / EMU : Number(getAttrLocal(pos, 'x') ?? 0) / EMU;
    const offsetY = from ? number(from, 'rowOff') / EMU : Number(getAttrLocal(pos, 'y') ?? 0) / EMU;
    const ext = firstChild(anchor, 'ext') ?? firstDescendant(firstDescendant(pic, 'xfrm'), 'ext');
    let width = Number(getAttrLocal(ext, 'cx') ?? 0) / EMU;
    let height = Number(getAttrLocal(ext, 'cy') ?? 0) / EMU;
    if (to) {
      width = number(to, 'colOff') / EMU - offsetX;
      height = number(to, 'rowOff') / EMU - offsetY;
      for (let c = col; c < number(to, 'col'); c++) width += sheet.colWidths[c] ?? sheet.defaultColWidth;
      for (let r = row; r < number(to, 'row'); r++) height += sheet.rowHeights[r] ?? sheet.defaultRowHeight;
    }
    const name = getAttrLocal(firstDescendant(pic, 'cNvPr'), 'name') ?? `图片 ${images.length + 1}`;
    const base64 = await pkg.zip.file(media)!.async('base64');
    images.push({ id: `image-${images.length + 1}`, name, src: `data:${mimeForPath(media)};base64,${base64}`, row, col, offsetX, offsetY, width: Math.max(8, width || 160), height: Math.max(8, height || 120) });
  }
  return images;
}

// 功能：为图片生成可被 Excel 识别的锚点、资源关系及内容类型；时间：2026-10-24；作者：lq。
export function writeSheetImages(zip: JSZip, sheet: SheetModel, index: number): void {
  if (!sheet.images?.length) return;
  const drawingName = `mini-office-sheet-${index + 1}.xml`;
  const drawingPart = `xl/drawings/${drawingName}`;
  const mediaRels: string[] = [];
  const anchors = sheet.images.map((image, i) => {
    const match = /^data:(image\/[\w.+-]+);base64,([\s\S]+)$/.exec(image.src);
    if (!match) throw new Error(`图片“${image.name}”无法保存，请重新插入。`);
    const mime = match[1];
    const ext = mime === 'image/jpeg' ? 'jpg' : mime === 'image/svg+xml' ? 'svg' : mime.split('/')[1];
    const filename = `mini-office-${index + 1}-${i + 1}.${ext}`;
    zip.file(`xl/media/${filename}`, match[2], { base64: true });
    mediaRels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/image" Target="../media/${filename}"/>`);
    const width = Math.round(image.width * EMU); const height = Math.round(image.height * EMU);
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${image.col}</xdr:col><xdr:colOff>${Math.round(image.offsetX * EMU)}</xdr:colOff><xdr:row>${image.row}</xdr:row><xdr:rowOff>${Math.round(image.offsetY * EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${width}" cy="${height}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="${escapeXml(image.name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId${i + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join('');
  zip.file(drawingPart, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}">${anchors}</xdr:wsDr>`);
  zip.file(`xl/drawings/_rels/${drawingName}.rels`, `<Relationships xmlns="${PKG_REL}">${mediaRels.join('')}</Relationships>`);
  zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, `<Relationships xmlns="${PKG_REL}"><Relationship Id="rIdMiniOfficeDrawing" Type="${REL}/drawing" Target="../drawings/${drawingName}"/></Relationships>`);

}

// 功能：补齐图片 MIME 与 DrawingML 部件声明；时间：2026-10-24；作者：lq。
export async function writeImageContentTypes(zip: JSZip, sheets: SheetModel[]): Promise<void> {
  let xml = await zip.file('[Content_Types].xml')!.async('text');
  sheets.forEach((sheet, index) => {
    if (!sheet.images?.length) return;
    const path = `/xl/drawings/mini-office-sheet-${index + 1}.xml`;
    if (!xml.includes(`PartName="${path}"`)) xml = xml.replace('</Types>', `<Override PartName="${path}" ContentType="${DRAWING_TYPE}"/></Types>`);
    sheet.images.forEach((image) => {
      const mime = /^data:([^;]+);/.exec(image.src)?.[1] ?? 'image/png';
      const ext = mime === 'image/jpeg' ? 'jpg' : mime === 'image/svg+xml' ? 'svg' : mime.split('/')[1];
      if (!xml.includes(`Extension="${ext}"`)) xml = xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
    });
  });
  zip.file('[Content_Types].xml', xml);
}
