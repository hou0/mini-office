// 功能：验证 PPTX 基础编辑、页面复制、备注和空白文稿往返保存；时间：2026-10-24；作者：lq。
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { parsePptx, parsePptxPackage } from '../packages/editor-slides/src/parse';
import { serializePptx } from '../packages/editor-slides/src/serialize';
import { copySlide, replaceParagraphText } from '../packages/editor-slides/src/editing';
import { createBlankDeck } from '@mini-office/document-model';
import JSZip from 'jszip';

function textOf(element: { type: string; paragraphs?: Array<{ runs: Array<{ text: string }> }> }): string {
  return element.paragraphs?.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n') ?? '';
}

describe('PPT 编辑与安全回写', () => {
  it('可以从空白演示文稿生成可打开的 PPTX', async () => {
    const bytes = await serializePptx(createBlankDeck());
    const parsed = await parsePptx(bytes, { fileName: 'new.pptx' });
    expect(parsed.slides).toHaveLength(1);
    expect(textOf(parsed.slides[0].elements.find((element) => element.type === 'text') as never)).toContain('单击添加标题');
  });

  it('回写文字和备注后可再次解析，复制页拥有独立部件', async () => {
    if (typeof URL.createObjectURL !== 'function') {
      (URL as unknown as { createObjectURL: (blob: Blob) => string }).createObjectURL = () => 'blob:test';
      (URL as unknown as { revokeObjectURL: (url: string) => void }).revokeObjectURL = () => {};
    }
    const source = new Uint8Array(readFileSync('tests/fixtures/pptx/deck.pptx'));
    const deck = await parsePptx(source, { fileName: 'deck.pptx' });
    const firstText = deck.slides[0].elements.find((element) => element.type === 'text');
    expect(firstText?.type).toBe('text');
    if (!firstText || firstText.type !== 'text') throw new Error('fixture text missing');
    firstText.paragraphs = replaceParagraphText(firstText.paragraphs, '已编辑的标题');
    deck.slides[0].notes = '已编辑备注';
    const duplicate = copySlide(deck.slides[0]);
    deck.slides.splice(1, 0, duplicate);
    deck.slides.forEach((slide, index) => { slide.index = index; });
    const bytes = await serializePptx(deck, source);
    const parsed = await parsePptx(bytes, { fileName: 'edited.pptx' });
    expect(parsed.slides).toHaveLength(4);
    expect(textOf(parsed.slides[0].elements.find((element) => element.type === 'text') as never)).toContain('已编辑的标题');
    expect(parsed.slides[0].notes).toBe('已编辑备注');
    const secondSave = await serializePptx(parsed, bytes);
    expect((await parsePptx(secondSave, { fileName: 'edited-again.pptx' })).slides).toHaveLength(4);
    const pkg = await loadOoxmlPackage(bytes);
    try {
      expect(pkg.has('ppt/slides/mo-copy')).toBe(false);
      expect(pkg.names().filter((name) => /^ppt\/slides\/[^/]+\.xml$/.test(name)).length).toBe(4);
    } finally {
      pkg.dispose();
    }
  });

  // 功能：回归验证 PPTX 编辑保存不丢失未知部件；时间：2026-10-28；作者：lq。
  it('保存 PPTX 时保留未知部件', async () => {
    const source = new Uint8Array(readFileSync('tests/fixtures/pptx/deck.pptx'));
    const zip = await JSZip.loadAsync(source);
    zip.file('custom/unknown.bin', new Uint8Array([1, 3, 3, 7]));
    const withUnknown = await zip.generateAsync({ type: 'uint8array' });
    const deck = await parsePptx(withUnknown, { fileName: 'deck.pptx' });
    const saved = await serializePptx(deck, withUnknown);
    const output = await JSZip.loadAsync(saved);
    expect(await output.file('custom/unknown.bin')?.async('uint8array')).toEqual(new Uint8Array([1, 3, 3, 7]));
  });

  // 功能：回归验证 PPTX 形状填充和边框颜色可保存并重开；时间：2026-10-28；作者：lq。
  it('保存形状填充和边框颜色', async () => {
    const source = new Uint8Array(readFileSync('tests/fixtures/pptx/deck.pptx'));
    const deck = await parsePptx(source, { fileName: 'deck.pptx' });
    const shape = deck.slides.flatMap((slide) => slide.elements).find((element) => element.type === 'shape' || element.type === 'text');
    if (!shape || (shape.type !== 'shape' && shape.type !== 'text')) throw new Error('fixture drawable missing');
    shape.fill = { type: 'solid', color: '#123456' };
    shape.line = { color: '#654321', widthPx: 2, dash: 'solid' };
    const saved = await serializePptx(deck, source);
    const reparsed = await parsePptx(saved, { fileName: 'styled.pptx' });
    const updated = reparsed.slides.flatMap((slide) => slide.elements).find((element) => element.sourceId === shape.sourceId);
    if (!updated || (updated.type !== 'shape' && updated.type !== 'text')) throw new Error('updated drawable missing');
    expect(updated.fill).toEqual({ type: 'solid', color: '#123456' });
    expect(updated.line?.color).toBe('#654321');
  });
});
