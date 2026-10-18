import type { OoxmlPackage } from './package';
import { asArray, attr, node, parseXml, type XmlObject } from './xml';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'External' | 'Internal';
  /** Absolute part name inside the package (undefined for external targets). */
  resolved?: string;
}

function relsPathFor(part: string): string {
  const normalized = part.replace(/^\//, '');
  const slash = normalized.lastIndexOf('/');
  const dir = slash >= 0 ? normalized.slice(0, slash + 1) : '';
  const file = slash >= 0 ? normalized.slice(slash + 1) : normalized;
  return `${dir}_rels/${file}.rels`;
}

export function resolvePartPath(sourcePart: string, target: string): string | undefined {
  if (!target || /^[a-z]+:/i.test(target) || target.startsWith('#')) return undefined;
  if (target.startsWith('/')) return target.replace(/^\//, '');
  const normalized = sourcePart.replace(/^\//, '');
  const slash = normalized.lastIndexOf('/');
  const base = slash >= 0 ? normalized.slice(0, slash + 1) : '';
  const parts = (base + target).split('/');
  const stack: string[] = [];
  for (const part of parts) {
    if (!part || part === '.') continue;
    if (part === '..') stack.pop();
    else stack.push(part);
  }
  return stack.join('/');
}

export async function readRelationships(
  pkg: OoxmlPackage,
  sourcePart: string,
): Promise<Relationship[]> {
  const path = relsPathFor(sourcePart);
  if (!pkg.has(path)) return [];
  const doc = parseXml(await pkg.text(path));
  const root = node(doc.Relationships) ?? doc;
  return asArray(root.Relationship).map((rel) => {
    const target = attr(rel, 'Target') ?? '';
    const targetMode = attr(rel, 'TargetMode') as 'External' | 'Internal' | undefined;
    return {
      id: attr(rel, 'Id') ?? '',
      type: attr(rel, 'Type') ?? '',
      target,
      targetMode,
      resolved: targetMode === 'External' ? undefined : resolvePartPath(sourcePart, target),
    };
  });
}

export async function relationshipMap(
  pkg: OoxmlPackage,
  sourcePart: string,
): Promise<Map<string, Relationship>> {
  const list = await readRelationships(pkg, sourcePart);
  return new Map(list.map((rel) => [rel.id, rel]));
}

export async function findRelationshipByType(
  pkg: OoxmlPackage,
  sourcePart: string,
  typeSuffix: string,
): Promise<Relationship | undefined> {
  const rels = await readRelationships(pkg, sourcePart);
  return rels.find((rel) => rel.type.endsWith(typeSuffix));
}

/** The package's primary document part, resolved through `_rels/.rels`. */
export async function mainDocumentPart(pkg: OoxmlPackage): Promise<string | null> {
  const rels = await readRelationships(pkg, '');
  const officeDoc = rels.find((rel) => rel.type.endsWith('/officeDocument'));
  return officeDoc?.resolved ?? null;
}

export async function contentTypes(pkg: OoxmlPackage): Promise<{
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}> {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (!pkg.has('[Content_Types].xml')) return { defaults, overrides };
  const normalizedRoot = (await pkg.xml('[Content_Types].xml')).Types as XmlObject | undefined;
  const root = node(normalizedRoot);
  if (!root) return { defaults, overrides };
  for (const item of asArray(root.Default)) {
    const ext = attr(item, 'Extension');
    const type = attr(item, 'ContentType');
    if (ext && type) defaults.set(ext.toLowerCase(), type);
  }
  for (const item of asArray(root.Override)) {
    const part = attr(item, 'PartName');
    const type = attr(item, 'ContentType');
    if (part && type) overrides.set(part.replace(/^\//, ''), type);
  }
  return { defaults, overrides };
}

export async function contentTypeOf(pkg: OoxmlPackage, part: string): Promise<string | undefined> {
  const { defaults, overrides } = await contentTypes(pkg);
  const normalized = part.replace(/^\//, '');
  const override = overrides.get(normalized);
  if (override) return override;
  const ext = normalized.split('.').pop()?.toLowerCase();
  return ext ? defaults.get(ext) : undefined;
}

/** Resolve the deck's slide parts in presentation order. */
export async function presentationSlideParts(
  pkg: OoxmlPackage,
  presentationPart = 'ppt/presentation.xml',
): Promise<string[]> {
  const rels = await relationshipMap(pkg, presentationPart);
  const xml = await pkg.xml(presentationPart);
  const presentation = node(xml.presentation);
  const slideIdList = node(presentation?.sldIdLst);
  const ids = asArray(slideIdList?.sldId);
  return ids
    .map((item) => {
      const rid = attr(item, 'r:id') ?? attr(item, 'id');
      const rel = rid ? rels.get(rid) : undefined;
      return rel?.resolved;
    })
    .filter((part): part is string => !!part);
}

export async function workbookSheetParts(
  pkg: OoxmlPackage,
  workbookPart = 'xl/workbook.xml',
): Promise<Array<{ name: string; sheetId: string; state: string; part: string }>> {
  const rels = await relationshipMap(pkg, workbookPart);
  const xml = await pkg.xml(workbookPart);
  const workbook = node(xml.workbook);
  const sheets = asArray(node(workbook?.sheets)?.sheet);
  return sheets
    .map((sheet) => {
      const rid = attr(sheet, 'r:id') ?? '';
      return {
        name: attr(sheet, 'name') ?? 'Sheet',
        sheetId: attr(sheet, 'sheetId') ?? '',
        state: attr(sheet, 'state') ?? 'visible',
        part: rels.get(rid)?.resolved ?? '',
      };
    })
    .filter((sheet) => !!sheet.part);
}
