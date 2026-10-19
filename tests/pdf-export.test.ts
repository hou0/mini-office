/** 功能：回归验证 PDF 字体基线、无损画质与画布释放，保留 JPEG 图片导出兼容性；时间：2026-10-19；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { exportElementsToPdf, exportElementsToZip, PDF_CAPTURE_SCALE } from '../packages/export-service/src/pdf';

const mocks = vi.hoisted(() => ({ capture: vi.fn(), addImage: vi.fn(), output: vi.fn() }));
vi.mock('../packages/export-service/node_modules/html2canvas', () => ({ default: mocks.capture }));
vi.mock('../packages/export-service/node_modules/jspdf', () => ({
  jsPDF: class {
    setProperties() {}
    addPage() {}
    addImage = mocks.addImage;
    output = mocks.output;
  },
}));

let page: HTMLDivElement;
let canvas: { width: number; height: number; toDataURL: ReturnType<typeof vi.fn> };

beforeEach(() => {
  vi.clearAllMocks();
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => { callback(0); return 1; });
  page = document.createElement('div');
  document.body.append(page);
  Object.defineProperties(page, { offsetWidth: { value: 794 }, offsetHeight: { value: 1123 } });
  canvas = { width: 2481, height: 3509, toDataURL: vi.fn((mime) => `data:${mime};base64,AA==`) };
  mocks.capture.mockResolvedValue(canvas);
  mocks.output.mockReturnValue(new ArrayBuffer(8));
});
afterEach(() => {
  page.remove();
  vi.unstubAllGlobals();
});

describe('PDF 页面渲染', () => {
  it('默认按 300 DPI 使用无损 PNG，写入原始页面尺寸并释放画布', async () => {
    await exportElementsToPdf({ pages: [page], fileName: 'test.pdf' });
    expect(PDF_CAPTURE_SCALE).toBe(300 / 96);
    expect(mocks.capture.mock.calls[0][1]).toMatchObject({ scale: 300 / 96, width: 794, height: 1123 });
    expect(canvas.toDataURL).toHaveBeenCalledWith('image/png', undefined);
    expect(mocks.addImage).toHaveBeenCalledWith('data:image/png;base64,AA==', 'PNG', 0, 0, 794, 1123, undefined, 'FAST');
    expect(canvas.width).toBe(0); expect(canvas.height).toBe(0);
  });

  it('在原始页面修正截图库的隐藏基线探针，保持真实图片布局并清理样式', async () => {
    const reset = document.createElement('style');
    reset.textContent = 'img { display: block; }';
    document.head.append(reset);
    const realImage = document.createElement('img');
    page.append(realImage);
    const probeContainer = document.createElement('div');
    probeContainer.style.visibility = 'hidden';
    const probe = document.createElement('img');
    probe.width = 1; probe.height = 1;
    probeContainer.append(probe);
    document.body.append(probeContainer);
    try {
      expect(getComputedStyle(probe).display).toBe('block');
      mocks.capture.mockImplementationOnce(async () => {
        expect(getComputedStyle(probe).display).toBe('inline-block');
        expect(getComputedStyle(realImage).display).toBe('block');
        return canvas;
      });
      await exportElementsToPdf({ pages: [page], fileName: 'test.pdf' });
      expect(getComputedStyle(probe).display).toBe('block');
      expect(document.querySelector('[data-pdf-font-metrics]')).toBeNull();
    } finally { reset.remove(); probeContainer.remove(); }
  });

  it('编码失败时仍释放大画布', async () => {
    canvas.toDataURL.mockImplementation(() => { throw new Error('encoding failed'); });
    await expect(exportElementsToPdf({ pages: [page], fileName: 'test.pdf' })).rejects.toThrow('encoding failed');
    expect(canvas.width).toBe(0); expect(canvas.height).toBe(0);
    expect(document.querySelector('[data-pdf-font-metrics]')).toBeNull();
  });

  it('图片 ZIP 保留 JPEG 格式与原有倍率', async () => {
    const bytes = await exportElementsToZip({ pages: [page] });
    expect(mocks.capture.mock.calls[0][1].scale).toBe(2);
    expect(canvas.toDataURL).toHaveBeenCalledWith('image/jpeg', 0.94);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file('page-01.jpg')).not.toBeNull();
  });
});
