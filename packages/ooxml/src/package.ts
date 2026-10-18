import JSZip from 'jszip';
import { OfficeError, ErrorCodes, extensionOf } from '@mini-office/shared';
import { parseXml, type XmlObject } from './xml';

const MAX_ENTRIES = 6000;
const MAX_UNCOMPRESSED_BYTES = 600 * 1024 * 1024;
const MAX_SINGLE_ENTRY_BYTES = 250 * 1024 * 1024;

const MIME_BY_EXTENSION: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
  mp4: 'video/mp4',
  m4a: 'audio/mp4',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  bin: 'application/octet-stream',
};

export function mimeForPath(path: string): string {
  const ext = extensionOf(path).replace('.', '');
  return MIME_BY_EXTENSION[ext] ?? 'application/octet-stream';
}

export interface PackageEntry {
  name: string;
  size: number;
  dir: boolean;
}

/**
 * Thin, safe wrapper around a JSZip OOXML package. Unknown parts are never
 * discarded here: the wrapper is read-only and serializers decide what to copy.
 */
export class OoxmlPackage {
  readonly zip: JSZip;
  readonly originalBytes: Uint8Array;
  private textCache = new Map<string, string>();
  private blobCache = new Map<string, string>();
  private entrySet: Set<string>;

  constructor(zip: JSZip, originalBytes: Uint8Array) {
    this.zip = zip;
    this.originalBytes = originalBytes;
    this.entrySet = new Set(
      Object.values(zip.files)
        .filter((file) => !file.dir)
        .map((file) => file.name),
    );
  }

  names(): string[] {
    return [...this.entrySet];
  }

  entries(): PackageEntry[] {
    return Object.values(this.zip.files).map((file) => ({
      name: file.name,
      size: (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0,
      dir: file.dir,
    }));
  }

  has(name: string): boolean {
    return this.entrySet.has(name.replace(/^\//, ''));
  }

  find(pattern: RegExp): string[] {
    return this.names().filter((name) => pattern.test(name));
  }

  async text(name: string): Promise<string> {
    const path = name.replace(/^\//, '');
    const cached = this.textCache.get(path);
    if (cached !== undefined) return cached;
    const entry = this.zip.file(path);
    if (!entry) throw new OfficeError(ErrorCodes.CORRUPT_PACKAGE, `缺少 OOXML 部件：${path}`);
    const text = await entry.async('text');
    this.textCache.set(path, text);
    return text;
  }

  async xml(name: string): Promise<XmlObject> {
    return parseXml(await this.text(name));
  }

  async bytes(name: string): Promise<Uint8Array> {
    const path = name.replace(/^\//, '');
    const entry = this.zip.file(path);
    if (!entry) throw new OfficeError(ErrorCodes.CORRUPT_PACKAGE, `缺少 OOXML 部件：${path}`);
    return entry.async('uint8array');
  }

  /** Lazily create (and cache) a blob URL for an embedded media part. */
  async blobUrl(name: string): Promise<string | null> {
    const path = name.replace(/^\//, '');
    const cached = this.blobCache.get(path);
    if (cached) return cached;
    if (!this.has(path)) return null;
    const bytes = await this.bytes(path);
    const buffer = new ArrayBuffer(bytes.byteLength);
    new Uint8Array(buffer).set(bytes);
    const url = URL.createObjectURL(new Blob([buffer], { type: mimeForPath(path) }));
    this.blobCache.set(path, url);
    return url;
  }

  dispose(): void {
    for (const url of this.blobCache.values()) {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // ignore
      }
    }
    this.blobCache.clear();
    this.textCache.clear();
  }
}

export function isZipLike(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export function isLegacyOle(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  );
}

export async function loadOoxmlPackage(bytes: Uint8Array): Promise<OoxmlPackage> {
  if (isLegacyOle(bytes)) {
    throw new OfficeError(
      ErrorCodes.LEGACY_FORMAT,
      '这是旧版二进制 Office 格式（.doc/.xls/.ppt），V1 暂不支持直接打开，请先另存为 OOXML 格式。',
    );
  }
  if (!isZipLike(bytes)) {
    throw new OfficeError(ErrorCodes.UNSUPPORTED_FORMAT, '文件不是有效的 OOXML 压缩包。');
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  } catch (error) {
    throw new OfficeError(
      ErrorCodes.CORRUPT_PACKAGE,
      '文件无法解压，可能已损坏或被加密。',
      error instanceof Error ? error.message : String(error),
    );
  }
  const files = Object.values(zip.files).filter((file) => !file.dir);
  if (files.length > MAX_ENTRIES) {
    throw new OfficeError(ErrorCodes.TOO_LARGE, `文件包含过多部件（${files.length}），已超过安全上限。`);
  }
  let total = 0;
  for (const file of files) {
    const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
    total += size;
    if (size > MAX_SINGLE_ENTRY_BYTES) {
      throw new OfficeError(ErrorCodes.TOO_LARGE, `部件 ${file.name} 解压后过大，已阻止打开。`);
    }
  }
  if (total > MAX_UNCOMPRESSED_BYTES) {
    throw new OfficeError(ErrorCodes.TOO_LARGE, '文件解压后体积超过安全上限，已阻止打开。');
  }
  return new OoxmlPackage(zip, bytes);
}
