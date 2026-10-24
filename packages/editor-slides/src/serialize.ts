// 功能：按页面和对象身份增量回写 PPTX，保留未编辑的 OOXML 部件；时间：2026-10-24；作者：lq。
import JSZip from 'jszip';
import type { SlideDeck, SlideElement, SlideFill, SlideModel, SlideTextParagraph } from '@mini-office/document-model';
import { firstChild, firstDescendant, parseXmlDom } from '@mini-office/ooxml';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS = `xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"`;
const EMU = 9525;
const xml = (value: unknown) => String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[char]!);
const serialize = (node: Node) => new XMLSerializer().serializeToString(node);
const children = (node: Element, name?: string) => Array.from(node.children).filter((item) => !name || item.tagName === name);
const relPath = (part: string) => part.replace(/([^/]+)$/, '_rels/$1.rels');
const color = (value: string | undefined, fallback = '1D2939') => /^#[0-9a-f]{6}$/i.test(value ?? '') ? value!.slice(1) : fallback;
const emu = (value: number) => Math.round(value * EMU);
const objectId = (node: Element) => firstDescendant(node, 'p:cNvPr')?.getAttribute('id');

// 功能：生成有命名空间的节点，避免不同 XML 前缀污染原始文档；时间：2026-10-24；作者：lq。
function make(doc: Document, name: string, attrs: Record<string, string | number> = {}): Element {
  const node = doc.createElementNS(name.startsWith('p:') ? P : name.startsWith('a:') ? A : R, name);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, String(value));
  return node;
}
function fragment(doc: Document, text: string): Element {
  return doc.importNode(parseXmlDom(`<root ${NS}>${text}</root>`).documentElement.firstElementChild!, true) as Element;
}
function replaceChild(parent: Element, name: string, next: Element): void {
  const old = firstChild(parent, name);
  if (old) parent.replaceChild(next, old);
  else parent.appendChild(next);
}

// 功能：记录已保存的文字状态，支持保存后撤销再保存仍正确回写；时间：2026-10-24；作者：lq。
function changedText(node: Element, value: unknown, source?: string): boolean {
  const marker = Array.from(node.childNodes).find((child) => child.nodeType === 8 && child.nodeValue?.startsWith('mo-text:'));
  const text = JSON.stringify(value);
  const previous = marker?.nodeValue?.slice(8) ?? source;
  if (text === previous) return false;
  marker?.remove();
  node.appendChild(node.ownerDocument.createComment(`mo-text:${text.replace(/--/g, '\\u002d\\u002d')}`));
  return true;
}

// 功能：保留段落、富文本样式、换行、项目符号和中文字型；时间：2026-10-24；作者：lq。
function paragraphXml(paragraphs: SlideTextParagraph[]): string {
  return (paragraphs.length ? paragraphs : [{ runs: [] }]).map((paragraph) => {
    const align = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[paragraph.align ?? 'left'];
    const spacing = `${paragraph.lineSpacing ? `<a:lnSpc><a:spcPct val="${Math.round(paragraph.lineSpacing * 100000)}"/></a:lnSpc>` : ''}${paragraph.spaceBeforePt != null ? `<a:spcBef><a:spcPts val="${Math.round(paragraph.spaceBeforePt * 100)}"/></a:spcBef>` : ''}${paragraph.spaceAfterPt != null ? `<a:spcAft><a:spcPts val="${Math.round(paragraph.spaceAfterPt * 100)}"/></a:spcAft>` : ''}`;
    const bullet = paragraph.bullet ? `<a:buChar char="${xml(paragraph.bulletChar || '•')}"/>` : '<a:buNone/>';
    const runs = paragraph.runs.map((run) => {
      const s = run.style;
      const attrs = ` lang="zh-CN"${s.fontSizePt ? ` sz="${Math.round(s.fontSizePt * 100)}"` : ''} b="${s.bold ? 1 : 0}" i="${s.italic ? 1 : 0}" u="${s.underline ? 'sng' : 'none'}" strike="${s.strike ? 'sngStrike' : 'noStrike'}"${s.letterSpacingPx ? ` spc="${Math.round(s.letterSpacingPx * 75)}"` : ''}`;
      const style = `<a:rPr${attrs}><a:solidFill><a:srgbClr val="${color(s.color)}"/></a:solidFill>${s.highlight ? `<a:highlight><a:srgbClr val="${color(s.highlight)}"/></a:highlight>` : ''}${s.fontFamily ? `<a:latin typeface="${xml(s.fontFamily)}"/><a:ea typeface="${xml(s.fontFamily)}"/>` : ''}</a:rPr>`;
      return run.lineBreak ? `<a:br/>` : `<a:r>${style}<a:t xml:space="preserve">${xml(run.text)}</a:t></a:r>`;
    }).join('');
    return `<a:p><a:pPr algn="${align}" lvl="${paragraph.level ?? 0}">${spacing}${bullet}</a:pPr>${runs}<a:endParaRPr lang="zh-CN"/></a:p>`;
  }).join('');
}
function writeText(body: Element, paragraphs: SlideTextParagraph[]): void {
  for (const node of children(body, 'a:p')) node.remove();
  const container = parseXmlDom(`<root ${NS}>${paragraphXml(paragraphs)}</root>`).documentElement;
  for (const node of Array.from(container.children)) body.appendChild(body.ownerDocument.importNode(node, true));
}
function fillXml(fill?: SlideFill): string {
  if (!fill || fill.type === 'none') return '<a:noFill/>';
  if (fill.type === 'solid') return `<a:solidFill><a:srgbClr val="${color(fill.color, 'FFFFFF')}"/></a:solidFill>`;
  if (fill.type === 'gradient') return `<a:gradFill><a:gsLst>${fill.stops.map((stop) => `<a:gs pos="${Math.round(stop.offset * 100000)}"><a:srgbClr val="${color(stop.color)}"/></a:gs>`).join('')}</a:gsLst><a:lin ang="${Math.round((fill.angle ?? 0) * 60000)}" scaled="1"/></a:gradFill>`;
  return '';
}

// 功能：组内坐标逆变换，移动/缩放后保持组合图形及未知子对象的坐标系；时间：2026-10-24；作者：lq。
interface CoordinateSpace { x: number; y: number; sx: number; sy: number }
function writeFrame(node: Element, element: SlideElement, space?: CoordinateSpace): void {
  const doc = node.ownerDocument;
  const parent = element.type === 'group' ? firstChild(node, 'p:grpSpPr') : element.type === 'table' ? node : firstChild(node, 'p:spPr');
  if (!parent) return;
  const tag = element.type === 'table' ? 'p:xfrm' : 'a:xfrm';
  const xfrm = firstChild(parent, tag) ?? make(doc, tag);
  if (!xfrm.parentNode) parent.insertBefore(xfrm, parent.firstChild);
  xfrm.setAttribute('rot', String(Math.round((element.rotation ?? 0) * 60000)));
  xfrm.setAttribute('flipH', element.flipH ? '1' : '0');
  xfrm.setAttribute('flipV', element.flipV ? '1' : '0');
  replaceChild(xfrm, 'a:off', make(doc, 'a:off', { x: emu(space ? element.x / space.sx + space.x : element.x), y: emu(space ? element.y / space.sy + space.y : element.y) }));
  replaceChild(xfrm, 'a:ext', make(doc, 'a:ext', { cx: emu(space ? element.w / space.sx : element.w), cy: emu(space ? element.h / space.sy : element.h) }));
}
function textBodyXml(element: SlideElement): string {
  if (element.type !== 'text' && element.type !== 'shape') return '';
  const margins = element.type === 'text' ? element.margins : undefined;
  return `<p:txBody><a:bodyPr anchor="${element.verticalAlign === 'middle' ? 'ctr' : element.verticalAlign === 'bottom' ? 'b' : 't'}"${margins ? ` tIns="${emu(margins.top)}" rIns="${emu(margins.right)}" bIns="${emu(margins.bottom)}" lIns="${emu(margins.left)}"` : ''}/><a:lstStyle/>${paragraphXml(element.paragraphs ?? [])}</p:txBody>`;
}

// 功能：新增文本/形状/图片生成原生可编辑 OOXML 对象，图片写入压缩包；时间：2026-10-24；作者：lq。
async function createObject(doc: Document, element: SlideElement, id: number, zip: JSZip, rels: Document): Promise<Element> {
  const nv = `<p:cNvPr id="${id}" name="mo:${xml(element.id)}"/>`;
  if (element.type === 'image') {
    const match = /^data:(image\/[\w+.-]+);base64,([\s\S]+)$/.exec(element.src);
    let target = element.part ? `/${element.part}` : '';
    if (match) {
      const ext = match[1] === 'image/jpeg' ? 'jpg' : match[1] === 'image/svg+xml' ? 'svg' : match[1].split('/')[1];
      const part = `ppt/media/mo-${element.id}.${ext}`;
      zip.file(part, match[2], { base64: true });
      target = `/${part}`;
    }
    if (!target) throw new Error('图片数据不可用，请重新插入图片。');
    const rid = addRel(rels, `${R}/image`, target);
    return fragment(doc, `<p:pic><p:nvPicPr>${nv}<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
  }
  if (element.type === 'group' || element.type === 'table') throw new Error('该对象缺少原始部件，无法安全保存。');
  const shape = element.shape ?? 'rect';
  return fragment(doc, `<p:sp><p:nvSpPr>${nv}<p:cNvSpPr${element.type === 'text' ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="${xml(shape)}"><a:avLst/></a:prstGeom>${fillXml(element.fill)}${element.line?.color ? `<a:ln w="${emu(element.line.widthPx ?? 1)}"><a:solidFill><a:srgbClr val="${color(element.line.color)}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>'}</p:spPr>${textBodyXml(element)}</p:sp>`);
}
function addRel(doc: Document, type: string, target: string): string {
  let count = 1;
  const used = new Set(children(doc.documentElement).map((item) => item.getAttribute('Id')));
  while (used.has(`rIdMO${count}`)) count++;
  const id = `rIdMO${count}`;
  const node = doc.createElementNS(REL, 'Relationship');
  node.setAttribute('Id', id); node.setAttribute('Type', type); node.setAttribute('Target', target);
  doc.documentElement.appendChild(node);
  return id;
}
async function readRels(zip: JSZip, part: string): Promise<Document> {
  return parseXmlDom(await zip.file(relPath(part))?.async('text') ?? `<Relationships xmlns="${REL}"/>`);
}
function resolvePart(from: string, target: string): string {
  const result: string[] = target.startsWith('/') ? [] : from.split('/').slice(0, -1);
  for (const part of target.split('/')) { if (part === '..') result.pop(); else if (part && part !== '.') result.push(part); }
  return result.join('/');
}
function outputPart(slide: SlideModel): string {
  return slide.sourcePart && !slide.id.startsWith('copy-') ? slide.sourcePart : `ppt/slides/mo-${slide.id.replace(/[^\w-]/g, '')}.xml`;
}

// 功能：只删除已解析且被用户删除的对象，保留未支持的图表/媒体节点；时间：2026-10-24；作者：lq。
async function writeObjects(parent: Element, elements: SlideElement[], originalIds: string[], zip: JSZip, rels: Document, nextId: { value: number }, space?: CoordinateSpace): Promise<void> {
  const own = elements.filter((item) => !item.inherited);
  const wanted = new Set(own.map((item) => item.sourceId));
  const originals = new Map(children(parent).flatMap((node) => { const id = objectId(node); return id ? [[id, node] as const] : []; }));
  for (const node of children(parent)) {
    const name = firstDescendant(node, 'p:cNvPr')?.getAttribute('name') ?? '';
    if (name.startsWith('mo:') || (originalIds.includes(objectId(node) ?? '') && !wanted.has(objectId(node) ?? ''))) node.remove();
  }
  for (const element of own) {
    let node = element.sourceId ? originals.get(element.sourceId) : undefined;
    if (!node || !node.parentNode) {
      node = await createObject(parent.ownerDocument, element, nextId.value++, zip, rels);
      parent.appendChild(node);
    }
    writeFrame(node, element, space);
    if (element.type === 'text' || element.type === 'shape') {
      if (changedText(node, element.paragraphs ?? [], element.sourceText)) {
        let body = firstChild(node, 'p:txBody');
        if (!body) { body = fragment(node.ownerDocument, textBodyXml(element)); node.appendChild(body); }
        else writeText(body, element.paragraphs ?? []);
      }
    } else if (element.type === 'table' && changedText(node, element.rows, element.sourceText)) {
      const table = firstDescendant(node, 'a:tbl');
      if (table) children(table, 'a:tr').forEach((row, rowIndex) => {
        const cells = children(row, 'a:tc');
        element.rows[rowIndex]?.cells.forEach((cell, colIndex) => {
          const body = firstChild(cells[cell.sourceIndex ?? colIndex], 'a:txBody');
          if (body) writeText(body, cell.text);
        });
      });
    } else if (element.type === 'group') {
      const transform = firstDescendant(firstChild(node, 'p:grpSpPr'), 'a:xfrm');
      const chOff = firstChild(transform, 'a:chOff'); const chExt = firstChild(transform, 'a:chExt');
      const cw = Number(chExt?.getAttribute('cx') ?? emu(element.w)) / EMU;
      const ch = Number(chExt?.getAttribute('cy') ?? emu(element.h)) / EMU;
      await writeObjects(node, element.children, element.sourceChildren ?? [], zip, rels, nextId, {
        x: Number(chOff?.getAttribute('x') ?? 0) / EMU, y: Number(chOff?.getAttribute('y') ?? 0) / EMU,
        sx: cw ? element.w / cw : 1, sy: ch ? element.h / ch : 1,
      });
    }
  }
}
function contentType(doc: Document, part: string, type: string): void {
  if (children(doc.documentElement).some((item) => item.getAttribute('PartName') === `/${part}`)) return;
  const node = doc.createElementNS(CT, 'Override'); node.setAttribute('PartName', `/${part}`); node.setAttribute('ContentType', type); doc.documentElement.appendChild(node);
}

// 功能：备注新建与复制时使用独立部件，清空备注不会恢复旧内容；时间：2026-10-24；作者：lq。
async function writeNotes(zip: JSZip, slide: SlideModel, part: string, rels: Document, types: Document): Promise<void> {
  const old = children(rels.documentElement).find((node) => node.getAttribute('Type')?.endsWith('/notesSlide'));
  if (!old && !slide.notes) return;
  const copying = part !== slide.sourcePart;
  const oldPart = old ? resolvePart(slide.sourcePart ?? part, old.getAttribute('Target') ?? '') : '';
  const notesPart = copying || !oldPart ? `ppt/notesSlides/mo-${slide.id.replace(/[^\w-]/g, '')}.xml` : oldPart;
  const notesXml = oldPart ? await zip.file(oldPart)?.async('text') : undefined;
  const notes = parseXmlDom(notesXml ?? `<p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
  const tree = firstDescendant(notes, 'p:spTree')!;
  const shapes = children(tree, 'p:sp').filter((node) => {
    const type = firstDescendant(node, 'p:ph')?.getAttribute('type');
    return type === 'body' || !type;
  });
  let body = shapes[0] ? firstChild(shapes[0], 'p:txBody') : null;
  if (!body) {
    const shape = fragment(notes, `<p:sp><p:nvSpPr><p:cNvPr id="1000" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`);
    tree.appendChild(shape); body = firstChild(shape, 'p:txBody');
  }
  writeText(body!, (slide.notes ?? '').split('\n').map((text) => ({ runs: [{ text, style: { fontSizePt: 14 } }] })));
  for (const shape of shapes.slice(1)) { const extra = firstChild(shape, 'p:txBody'); if (extra) writeText(extra, []); }
  zip.file(notesPart, serialize(notes));
  const notesRels = oldPart ? await readRels(zip, oldPart) : parseXmlDom(`<Relationships xmlns="${REL}"/>`);
  for (const node of children(notesRels.documentElement)) {
    if (node.getAttribute('Type')?.endsWith('/slide')) node.remove();
    else if (oldPart && node.getAttribute('TargetMode') !== 'External') node.setAttribute('Target', `/${resolvePart(oldPart, node.getAttribute('Target') ?? '')}`);
  }
  addRel(notesRels, `${R}/slide`, `/${part}`);
  zip.file(relPath(notesPart), serialize(notesRels));
  if (old) old.setAttribute('Target', `/${notesPart}`); else addRel(rels, `${R}/notesSlide`, `/${notesPart}`);
  contentType(types, notesPart, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
}

// 功能：创建 PowerPoint 可直接打开并继续编辑的最小主题/母版/空白版式；时间：2026-10-24；作者：lq。
function seedPackage(zip: JSZip, deck: SlideDeck): void {
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst><p:sldIdLst/><p:sldSz cx="${deck.widthEmu}" cy="${deck.heightEmu}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="${REL}"><Relationship Id="rIdMaster" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`);
  const tree = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="空白">${tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${NS}><p:cSld>${tree}</p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="../theme/theme1.xml"/></Relationships>`);
  const themeColors = { dk1: '1D2939', lt1: 'FFFFFF', dk2: '344054', lt2: 'F2F4F7', accent1: '2E7D72', accent2: 'C43E1C', accent3: '2B579A', accent4: '8064A2', accent5: '4BACC6', accent6: 'F79646', hlink: '0563C1', folHlink: '954F72' };
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="${A}" name="Mini Office"><a:themeElements><a:clrScheme name="Mini Office">${Object.entries(themeColors).map(([key, value]) => `<a:${key}><a:srgbClr val="${value}"/></a:${key}>`).join('')}</a:clrScheme><a:fontScheme name="Mini Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface="等线"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="等线"/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Mini Office"><a:fillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${'<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'.repeat(3)}</a:lnStyleLst><a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`);
}

/** 功能：保存编辑、新建、复制、排序、备注及图片，原始文件由调用方另存保护；时间：2026-10-24；作者：lq。 */
export async function serializePptx(deck: SlideDeck, sourceBytes?: Uint8Array): Promise<Uint8Array> {
  if (!deck.slides.length) throw new Error('演示文稿至少需要一页幻灯片。');
  const zip = sourceBytes ? await JSZip.loadAsync(sourceBytes) : new JSZip();
  if (!sourceBytes) seedPackage(zip, deck);
  const types = parseXmlDom(await zip.file('[Content_Types].xml')!.async('text'));
  for (const [ext, mime] of Object.entries({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' })) {
    if (!children(types.documentElement).some((item) => item.getAttribute('Extension') === ext)) {
      const node = types.createElementNS(CT, 'Default'); node.setAttribute('Extension', ext); node.setAttribute('ContentType', mime); types.documentElement.insertBefore(node, types.documentElement.firstChild);
    }
  }
  const presentation = parseXmlDom(await zip.file('ppt/presentation.xml')!.async('text'));
  const rels = await readRels(zip, 'ppt/presentation.xml');
  for (const node of children(rels.documentElement)) if (node.getAttribute('Type') === `${R}/slide`) node.remove();
  const list = firstDescendant(presentation, 'p:sldIdLst') ?? make(presentation, 'p:sldIdLst');
  if (!list.parentNode) presentation.documentElement.insertBefore(list, firstDescendant(presentation, 'p:sldSz'));
  list.replaceChildren();
  // 功能：先缓存模板，防止复制页受同一轮保存中原页修改影响；时间：2026-10-24；作者：lq。
  const templates = new Map<string, string>();
  for (const slide of deck.slides) if (slide.sourcePart && !templates.has(slide.sourcePart)) {
    const text = await zip.file(slide.sourcePart)?.async('text'); if (text) templates.set(slide.sourcePart, text);
  }
  for (const [index, slide] of deck.slides.entries()) {
    const part = outputPart(slide);
    const template = slide.sourcePart ? templates.get(slide.sourcePart) : undefined;
    const doc = parseXmlDom(template ?? `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
    const slideRels = slide.sourcePart ? await readRels(zip, slide.sourcePart) : parseXmlDom(`<Relationships xmlns="${REL}"/>`);
    if (!slide.sourcePart) {
      const layout = Object.keys(zip.files).find((name) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(name));
      if (layout) addRel(slideRels, `${R}/slideLayout`, `/${layout}`);
      const bg = fragment(doc, `<p:bg><p:bgPr>${fillXml(slide.background ?? { type: 'solid', color: '#FFFFFF' })}<a:effectLst/></p:bgPr></p:bg>`);
      firstDescendant(doc, 'p:cSld')!.insertBefore(bg, firstDescendant(doc, 'p:spTree'));
    } else if (part !== slide.sourcePart) {
      for (const node of children(slideRels.documentElement)) if (node.getAttribute('TargetMode') !== 'External') node.setAttribute('Target', `/${resolvePart(slide.sourcePart, node.getAttribute('Target') ?? '')}`);
    }
    const maxId = Math.max(1, ...Array.from(doc.getElementsByTagName('p:cNvPr')).map((node) => Number(node.getAttribute('id')) || 1));
    await writeObjects(firstDescendant(doc, 'p:spTree')!, slide.elements, slide.sourceElementIds ?? [], zip, slideRels, { value: maxId + 1 });
    await writeNotes(zip, slide, part, slideRels, types);
    zip.file(part, serialize(doc)); zip.file(relPath(part), serialize(slideRels));
    contentType(types, part, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    const rid = addRel(rels, `${R}/slide`, `/${part}`);
    const id = make(presentation, 'p:sldId', { id: 256 + index }); id.setAttributeNS(R, 'r:id', rid); list.appendChild(id);
  }
  // 功能：补全新建文件的部件类型，已有文件保留未知关系和内容类型；时间：2026-10-24；作者：lq。
  contentType(types, 'ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');
  for (const part of Object.keys(zip.files)) {
    if (/^ppt\/slideMasters\/[^/]+\.xml$/.test(part)) contentType(types, part, 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml');
    if (/^ppt\/slideLayouts\/[^/]+\.xml$/.test(part)) contentType(types, part, 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
    if (/^ppt\/theme\/[^/]+\.xml$/.test(part)) contentType(types, part, 'application/vnd.openxmlformats-officedocument.theme+xml');
  }
  zip.file('ppt/presentation.xml', serialize(presentation)); zip.file(relPath('ppt/presentation.xml'), serialize(rels)); zip.file('[Content_Types].xml', serialize(types));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
