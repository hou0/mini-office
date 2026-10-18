import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { analyzeCapabilities, detectKind, loadOoxmlPackage } from '@mini-office/ooxml';
import { parsePptxPackage } from '../packages/editor-slides/src/parse';
import type { SlideElement, SlideTableElement, SlideTextElement, SlideImageElement } from '@mini-office/document-model';

function fixture(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}

function textOf(element: SlideElement): string {
  if (element.type === 'text') {
    return element.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
  }
  if (element.type === 'shape') {
    return (element.paragraphs ?? []).map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
  }
  if (element.type === 'group') {
    return element.children.map(textOf).join('\n');
  }
  return '';
}

describe('PPTX fixture compatibility', () => {
  it('parses slides, shapes, images, tables and notes', async () => {
    if (typeof URL.createObjectURL !== 'function') {
      (URL as unknown as { createObjectURL: (blob: Blob) => string }).createObjectURL = () => 'blob:test';
      (URL as unknown as { revokeObjectURL: (url: string) => void }).revokeObjectURL = () => {};
    }
    const bytes = fixture('tests/fixtures/pptx/deck.pptx');
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const detection = await detectKind(pkg, 'deck.pptx');
      expect(detection.kind).toBe('pptx');
      const capability = await analyzeCapabilities(pkg, 'pptx', { fileName: 'deck.pptx' });
      expect(capability.level).toBe(1);
      const deck = await parsePptxPackage(pkg, { fileName: 'deck.pptx' });
      expect(deck.slides).toHaveLength(3);
      expect(deck.aspect).toBe('16:9');
      expect(Math.round(deck.widthPx)).toBe(1280);
      expect(Math.round(deck.heightPx)).toBe(720);
      expect(deck.slides[0].title).toContain('Mini Office');
      const firstText = deck.slides[0].elements.map(textOf).join('\n');
      expect(firstText).toContain('轻量、离线优先的跨平台办公套件');
      expect(firstText).toContain('统一打开 Word / Excel / PPT');
      const image = deck.slides[1].elements.find(
        (element): element is SlideImageElement => element.type === 'image',
      );
      expect(image).toBeTruthy();
      expect(image!.src).toContain('blob:');
      expect(image!.w).toBeGreaterThan(100);
      const card = deck.slides[1].elements.find(
        (element): element is SlideTextElement => element.type === 'text' && textOf(element).includes('静态高保真预览'),
      );
      expect(card).toBeTruthy();
      const table = deck.slides[2].elements.find(
        (element): element is SlideTableElement => element.type === 'table',
      );
      expect(table).toBeTruthy();
      expect(table!.rows.length).toBe(3);
      expect(table!.rows[1].cells[0].text[0].runs[0].text).toBe('P1');
      expect(deck.slides[2].notes).toContain('演讲提示');
      expect(deck.themeColors.accent1).toBeTruthy();
      expect(deck.slides[2].title).toContain('表格与备注');
    } finally {
      pkg.dispose();
    }
  });
});
