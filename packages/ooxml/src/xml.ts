import { XMLParser } from 'fast-xml-parser';

export type XmlNode = Record<string, unknown> | string | number | boolean | null | undefined;
export type XmlObject = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  allowBooleanAttributes: true,
  processEntities: true,
  htmlEntities: false,
});

export function parseXml(xml: string): XmlObject {
  const result = parser.parse(xml) as XmlObject;
  return result ?? {};
}

/** Parse XML and return the document root, ignoring the XML declaration. */
export function parseXmlRoot(xml: string): XmlObject {
  const parsed = parseXml(xml);
  const keys = Object.keys(parsed).filter((key) => key !== '?xml' && !key.startsWith('@_'));
  if (keys.length === 0) return {};
  return (parsed[keys[0]] as XmlObject) ?? {};
}

export function asArray<T = XmlObject>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

export function node(value: unknown): XmlObject | undefined {
  if (value && typeof value === 'object' && !Array.isArray(value)) return value as XmlObject;
  return undefined;
}

export function attr(obj: unknown, name: string): string | undefined {
  const n = node(obj);
  if (!n) return undefined;
  const value = n[`@_${name}`];
  if (value === undefined || value === null) return undefined;
  return String(value);
}

export function numAttr(obj: unknown, name: string, fallback?: number): number | undefined {
  const value = attr(obj, name);
  if (value === undefined || value === '') return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

export function boolAttr(obj: unknown, name: string, fallback = false): boolean {
  const value = attr(obj, name);
  if (value === undefined) return fallback;
  return value === '1' || value.toLowerCase() === 'true' || value.toLowerCase() === 'on';
}

export function textOf(value: unknown): string {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  const n = node(value);
  if (!n) return '';
  const text = n['#text'];
  if (typeof text === 'string' || typeof text === 'number') return String(text);
  return '';
}

export function attrOrText(value: unknown): string {
  return attr(value, 'val') ?? textOf(value);
}

export function child(obj: unknown, name: string): XmlObject | undefined {
  const n = node(obj);
  if (!n) return undefined;
  return node(n[name]);
}

export function children(obj: unknown, name: string): XmlObject[] {
  const n = node(obj);
  if (!n) return [];
  return asArray(n[name]).map((item) => node(item)).filter((item): item is XmlObject => !!item);
}

export function childText(obj: unknown, name: string): string {
  return textOf(child(obj, name));
}

export function childAttr(obj: unknown, name: string, attribute: string): string | undefined {
  return attr(child(obj, name), attribute);
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function escapeXmlAttr(value: string): string {
  return escapeXml(value).replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
}

/** Remove whitespace-only text nodes and XML declaration for compact serialization. */
export function compactXml(xml: string): string {
  return xml.replace(/>\s+</g, '><').trim();
}

export function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
}
