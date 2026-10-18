import { OfficeError, ErrorCodes } from '@mini-office/shared';

/** Parse an XML part into a DOM document (browser / jsdom). */
export function parseXmlDom(xml: string): Document {
  if (typeof DOMParser === 'undefined') {
    throw new OfficeError(ErrorCodes.PARSE_FAILED, '当前环境缺少 XML DOM 解析能力。');
  }
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const error = doc.getElementsByTagName('parsererror')[0];
  if (error) {
    throw new OfficeError(
      ErrorCodes.PARSE_FAILED,
      'OOXML 部件包含无效 XML，无法继续解析。',
      error.textContent ?? undefined,
    );
  }
  if (!doc.documentElement) {
    throw new OfficeError(ErrorCodes.PARSE_FAILED, 'OOXML 部件为空。');
  }
  return doc;
}

export function rootElement(doc: Document): Element {
  return doc.documentElement;
}

export function childElements(
  parent: Element | Document | null | undefined,
  tagName?: string,
): Element[] {
  if (!parent) return [];
  const children: Element[] = [];
  const source = parent.childNodes;
  for (let index = 0; index < source.length; index += 1) {
    const item = source.item(index);
    if (item && item.nodeType === 1) {
      const element = item as Element;
      if (!tagName || element.tagName === tagName) children.push(element);
    }
  }
  return children;
}

export function firstChild(
  parent: Element | Document | null | undefined,
  tagName: string,
): Element | null {
  if (!parent) return null;
  const source = parent.childNodes;
  for (let index = 0; index < source.length; index += 1) {
    const item = source.item(index);
    if (item && item.nodeType === 1 && (item as Element).tagName === tagName) {
      return item as Element;
    }
  }
  return null;
}

export function descendants(
  parent: Element | Document | null | undefined,
  tagName: string,
): Element[] {
  if (!parent) return [];
  return Array.from(parent.getElementsByTagName(tagName));
}

export function firstDescendant(
  parent: Element | Document | null | undefined,
  tagName: string,
): Element | null {
  if (!parent) return null;
  return parent.getElementsByTagName(tagName)[0] ?? null;
}

/** Attribute lookup by local name, tolerating different namespace prefixes. */
export function getAttrLocal(element: Element | null | undefined, localName: string): string | undefined {
  if (!element) return undefined;
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.localName === localName) return attribute.value;
  }
  return undefined;
}

export function getAttr(element: Element | null | undefined, name: string): string | undefined {
  if (!element) return undefined;
  const value = element.getAttribute(name);
  if (value !== null) return value;
  if (!name.includes(':')) return getAttrLocal(element, name);
  return undefined;
}

export function hasAttr(element: Element | null | undefined, name: string): boolean {
  if (!element) return false;
  return element.hasAttribute(name) || getAttrLocal(element, name) !== undefined;
}

export function numericAttr(element: Element | null | undefined, name: string, fallback?: number): number | undefined {
  const raw = getAttr(element, name);
  if (raw === undefined || raw === '') return fallback;
  const value = Number(raw);
  return Number.isFinite(value) ? value : fallback;
}

/** OOXML toggle properties: absent `val` means true. */
export function toggleValue(element: Element | null | undefined): boolean {
  if (!element) return false;
  const value = getAttr(element, 'w:val') ?? getAttrLocal(element, 'val');
  if (value === undefined) return true;
  return !['0', 'false', 'off', 'none'].includes(value.toLowerCase());
}

export function elementText(element: Element | null | undefined): string {
  return element?.textContent ?? '';
}

export function describeElement(element: Element): string {
  return `<${element.tagName}${Array.from(element.attributes)
    .map((attribute) => ` ${attribute.name}="${attribute.value}"`)
    .join('')}>`;
}
