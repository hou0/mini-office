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

// 功能：按 XML 本地名称匹配 OOXML 节点，兼容不同软件生成的命名空间前缀；时间：2026-10-28；作者：lq。
function matchesTag(element: Element, tagName: string): boolean {
  const expected = tagName.includes(':') ? tagName.slice(tagName.indexOf(':') + 1) : tagName;
  return localName(element) === expected;
}

/** 功能：统一取得 XML 元素本地名称，供需要 switch 的解析器兼容任意命名空间前缀；时间：2026-10-28；作者：lq。 */
export function localName(element: Element): string {
  // 功能：避免 split(':') 每次分配数组，该函数在大部件解析中被高频调用；时间：2026-10-29；作者：lq。
  const own = element.localName;
  if (own) return own;
  const tag = element.tagName;
  const index = tag.indexOf(':');
  return index === -1 ? tag : tag.slice(index + 1);
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
      if (!tagName || matchesTag(element, tagName)) children.push(element);
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
    if (item && item.nodeType === 1 && matchesTag(item as Element, tagName)) {
      return item as Element;
    }
  }
  return null;
}

/** 功能：惰性遍历所有元素后代，避免 getElementsByTagName('*') 一次性物化整棵树的巨额开销；时间：2026-10-29；作者：lq。 */
function* iterateDescendants(parent: Element | Document): Generator<Element> {
  const children = parent.children;
  for (let index = 0; index < children.length; index += 1) {
    const child = children[index];
    yield child;
    yield* iterateDescendants(child);
  }
}

export function descendants(
  parent: Element | Document | null | undefined,
  tagName: string,
): Element[] {
  if (!parent) return [];
  // 功能：优先用原生按名查找（无命名空间前缀的 OOXML 部件直接命中），
  // 再惰性扫描补齐带前缀的同名节点，兼顾速度与前缀兼容；时间：2026-10-29；作者：lq。
  const result: Element[] = [];
  const seen = new Set<Element>();
  const exact = parent.getElementsByTagName(tagName);
  for (let index = 0; index < exact.length; index += 1) {
    const element = exact[index];
    if (!seen.has(element)) { seen.add(element); result.push(element); }
  }
  for (const element of iterateDescendants(parent)) {
    if (seen.has(element)) continue;
    if (matchesTag(element, tagName)) { seen.add(element); result.push(element); }
  }
  return result;
}

export function firstDescendant(
  parent: Element | Document | null | undefined,
  tagName: string,
): Element | null {
  if (!parent) return null;
  // 功能：优先用原生按名查找，未命中再惰性扫描（带命名空间前缀时）并提前结束；时间：2026-10-29；作者：lq。
  const exact = parent.getElementsByTagName(tagName);
  if (exact.length > 0) return exact[0];
  for (const element of iterateDescendants(parent)) {
    if (matchesTag(element, tagName)) return element;
  }
  return null;
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
  // 功能：属性前缀不固定时仍读取 w:val、r:id 等 OOXML 属性；时间：2026-10-28；作者：lq。
  return getAttrLocal(element, name.includes(':') ? name.slice(name.indexOf(':') + 1) : name);
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
