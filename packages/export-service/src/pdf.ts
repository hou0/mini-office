import { ErrorCodes, OfficeError, type ExportProgress } from '@mini-office/shared';

export interface ElementPdfOptions {
  pages: HTMLElement[];
  fileName: string;
  title?: string;
  pageRange?: string;
  scale?: number;
  pageWidthPx?: number;
  pageHeightPx?: number;
  onProgress?: (progress: ExportProgress) => void;
  isCancelled?: () => boolean;
}

export interface ElementImagesOptions {
  pages: HTMLElement[];
  scale?: number;
  onProgress?: (progress: ExportProgress) => void;
  isCancelled?: () => boolean;
}

/** Parse "1-3,5" into zero-based page indices (clamped to the document). */
export function parsePageRange(range: string | undefined, total: number): number[] {
  if (!range || !range.trim()) return Array.from({ length: total }, (_, index) => index);
  const indices = new Set<number>();
  for (const part of range.split(',')) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const match = /^(\d+)(?:\s*[-–]\s*(\d+))?$/.exec(trimmed);
    if (!match) continue;
    const start = Math.max(1, Number(match[1]));
    const end = Math.min(total, Number(match[2] ?? match[1]));
    for (let page = start; page <= end; page += 1) indices.add(page - 1);
  }
  return [...indices].sort((a, b) => a - b);
}

async function waitForLayout(): Promise<void> {
  if (typeof document !== 'undefined' && 'fonts' in document) {
    try {
      await (document as Document & { fonts: FontFaceSet }).fonts.ready;
    } catch {
      // Font loading may be unsupported.
    }
  }
  await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
}

async function capturePage(
  element: HTMLElement,
  scale: number,
): Promise<{ dataUrl: string; width: number; height: number }> {
  const { default: html2canvas } = await import('html2canvas');
  const width = element.offsetWidth || element.getBoundingClientRect().width;
  const height = element.offsetHeight || element.getBoundingClientRect().height;
  const canvas = await html2canvas(element, {
    scale,
    backgroundColor: '#FFFFFF',
    useCORS: true,
    allowTaint: false,
    logging: false,
    width,
    height,
    windowWidth: width,
    windowHeight: height,
    scrollX: 0,
    scrollY: 0,
  });
  return { dataUrl: canvas.toDataURL('image/jpeg', 0.94), width, height };
}

export async function exportElementsToPdf(options: ElementPdfOptions): Promise<Uint8Array> {
  const { pages, fileName, pageRange, scale = 2, onProgress, isCancelled } = options;
  if (pages.length === 0) {
    throw new OfficeError(ErrorCodes.EXPORT_FAILED, '没有可导出的页面。');
  }
  const indices = parsePageRange(pageRange, pages.length);
  const selected = indices.map((index) => pages[index]).filter(Boolean);
  if (selected.length === 0) {
    throw new OfficeError(ErrorCodes.EXPORT_FAILED, '页码范围无效。');
  }
  onProgress?.({ phase: 'preparing', current: 0, total: selected.length, message: '正在准备字体与版面…' });
  await waitForLayout();

  const { jsPDF } = await import('jspdf');
  let pdf: InstanceType<typeof jsPDF> | null = null;
  for (let index = 0; index < selected.length; index += 1) {
    if (isCancelled?.()) throw new OfficeError(ErrorCodes.CANCELLED, '导出已取消');
    const page = selected[index];
    const width = options.pageWidthPx ?? page.offsetWidth;
    const height = options.pageHeightPx ?? page.offsetHeight;
    onProgress?.({
      phase: 'rendering',
      current: index,
      total: selected.length,
      message: `正在渲染第 ${indices[index] + 1} 页…`,
    });
    const capture = await capturePage(page, scale);
    if (!pdf) {
      pdf = new jsPDF({
        orientation: width >= height ? 'landscape' : 'portrait',
        unit: 'px',
        format: [width, height],
        hotfixes: ['px_scaling'],
        compress: true,
      });
      pdf.setProperties({ title: options.title ?? fileName, creator: 'Mini Office' });
    } else {
      pdf.addPage([width, height], width >= height ? 'landscape' : 'portrait');
    }
    pdf.addImage(capture.dataUrl, 'JPEG', 0, 0, width, height, undefined, 'FAST');
  }
  if (!pdf) throw new OfficeError(ErrorCodes.EXPORT_FAILED, 'PDF 生成失败。');
  onProgress?.({ phase: 'assembling', current: selected.length, total: selected.length, message: '正在写入 PDF…' });
  const output = pdf.output('arraybuffer');
  onProgress?.({ phase: 'done', current: selected.length, total: selected.length, message: '导出完成' });
  return new Uint8Array(output);
}

export async function exportElementsToZip(options: ElementImagesOptions): Promise<Uint8Array> {
  const { pages, scale = 2, onProgress, isCancelled } = options;
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();
  await waitForLayout();
  for (let index = 0; index < pages.length; index += 1) {
    if (isCancelled?.()) throw new OfficeError(ErrorCodes.CANCELLED, '导出已取消');
    onProgress?.({
      phase: 'rendering',
      current: index,
      total: pages.length,
      message: `正在渲染第 ${index + 1} 页…`,
    });
    const capture = await capturePage(pages[index], scale);
    const base64 = capture.dataUrl.split(',')[1] ?? '';
    zip.file(`page-${String(index + 1).padStart(2, '0')}.jpg`, base64, { base64: true });
  }
  onProgress?.({ phase: 'assembling', current: pages.length, total: pages.length, message: '正在打包图片…' });
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  onProgress?.({ phase: 'done', current: pages.length, total: pages.length, message: '导出完成' });
  return bytes;
}

export type { ExportProgress };
