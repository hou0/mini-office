import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createBlankDoc, createBlankWorkbook, createSheet } from '@mini-office/document-model';
import { analyzeCapabilities, loadOoxmlPackage, detectKind } from '@mini-office/ooxml';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { paginateBlocks } from '../packages/editor-word/src/pagination';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';
import { parseWorkbookPackage } from '../packages/editor-sheet/src/parse';
import { parseCsv, serializedCsvText } from './helpers/csv';
import { recalculateWorkbook } from '../packages/editor-sheet/src/formula/engine';
import { applyFilter, insertRows, sortRange } from '../packages/editor-sheet/src/sheet-utils';

async function macroPackage(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    'word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>宏文档</w:t></w:r></w:p></w:body></w:document>',
  );
  zip.file('word/vbaProject.bin', new Uint8Array([1, 2, 3, 4]));
  return zip.generateAsync({ type: 'uint8array' });
}

describe('OOXML package handling', () => {
  it('serializes and reparses a blank DOCX', async () => {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'paragraph',
        id: 'p1',
        runs: [{ text: 'Mini Office 空白文档', style: { fontSizePt: 12, bold: true } }],
        style: { alignment: 'center' },
      },
    ];
    const bytes = await serializeDocx(doc, { name: 'blank.docx' });
    const pkg = await loadOoxmlPackage(bytes);
    try {
      expect((await detectKind(pkg, 'blank.docx')).kind).toBe('docx');
      const parsed = await parseDocxPackage(pkg, { fileName: 'blank.docx' });
      const text = parsed.blocks.map((block) => (block.type === 'paragraph' ? block.runs.map((run) => run.text ?? '').join('') : '')).join('');
      expect(text).toContain('Mini Office 空白文档');
      expect(parsed.page.sizeName).toBe('A4');
    } finally {
      pkg.dispose();
    }
  });

  // 功能：回归验证 DOCX 编辑保存不丢失未知部件；时间：2026-10-28；作者：lq。
  it('preserves unknown DOCX parts during an edit save', async () => {
    const base = await serializeDocx(createBlankDoc(), { name: 'base.docx' });
    const source = await JSZip.loadAsync(base);
    source.file('custom/unknown.bin', new Uint8Array([9, 8, 7, 6]));
    const withUnknown = await source.generateAsync({ type: 'uint8array' });
    const pkg = await loadOoxmlPackage(withUnknown);
    const doc = await parseDocxPackage(pkg, { fileName: 'base.docx' });
    pkg.dispose();
    const saved = await serializeDocx(doc, { sourceBytes: withUnknown, name: 'edited.docx' });
    const output = await JSZip.loadAsync(saved);
    expect(await output.file('custom/unknown.bin')?.async('uint8array')).toEqual(new Uint8Array([9, 8, 7, 6]));
  });

  it('serializes and reparses a blank XLSX', async () => {
    const workbook = createBlankWorkbook();
    workbook.sheets[0].cells = {
      '0:0': { value: 'Hello' },
      '0:1': { value: 42 },
    };
    const bytes = await serializeXlsx(recalculateWorkbook(workbook), { name: 'blank.xlsx' });
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const parsed = await parseWorkbookPackage(pkg, 'blank.xlsx');
      expect(parsed.sheets).toHaveLength(1);
      expect(parsed.sheets[0].cells['0:0']?.value).toBe('Hello');
      expect(parsed.sheets[0].cells['0:1']?.value).toBe(42);
    } finally {
      pkg.dispose();
    }
  });

  it('classifies macro documents as read-only L1', async () => {
    const bytes = await macroPackage();
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const capability = await analyzeCapabilities(pkg, 'docx', { fileName: 'macro.docm' });
      expect(capability.level).toBe(1);
      expect(capability.warnings.some((warning) => warning.code === 'docx.macro')).toBe(true);
    } finally {
      pkg.dispose();
    }
  });

  it('paginates blocks with explicit and measured breaks', () => {
    const blocks = [
      { type: 'paragraph' as const, id: 'a', runs: [{ text: 'a', style: {} }], style: {} },
      { type: 'paragraph' as const, id: 'b', runs: [{ text: 'b', style: {} }], style: { pageBreakBefore: true } },
      { type: 'paragraph' as const, id: 'c', runs: [{ text: 'c', style: {} }], style: {} },
    ];
    const heights = new Map([
      ['a', 40],
      ['b', 40],
      ['c', 40],
    ]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['a']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['b', 'c']);
  });
});

describe('CSV and sheet utilities', () => {
  it('parses quoted CSV and preserves values on save', () => {
    const csv = '名称,备注\n合同A,"含,逗号"\n合同B,"引号""测试"\n';
    const workbook = parseCsv(csv, 'test.csv');
    expect(workbook.sheets[0].cells['1:1']?.value).toBe('含,逗号');
    expect(workbook.sheets[0].cells['2:1']?.value).toBe('引号"测试');
    const text = serializedCsvText(workbook);
    expect(text).toContain('"含,逗号"');
    expect(text).toContain('"引号""测试"');
  });

  it('inserts rows and shifts cells', () => {
    const base = createBlankWorkbook();
    const sheet = base.sheets[0];
    sheet.cells = { '0:0': { value: 'A' }, '1:0': { value: 'B' } };
    const next = insertRows(base, sheet.id, 0, 1);
    const nextSheet = next.sheets[0];
    expect(nextSheet.cells['1:0']?.value).toBe('A');
    expect(nextSheet.cells['2:0']?.value).toBe('B');
  });

  it('sorts data rows and filters by values', () => {
    const base = createBlankWorkbook();
    const sheet = base.sheets[0];
    sheet.cells = {
      '0:0': { value: '名称' },
      '1:0': { value: 'B' },
      '2:0': { value: 'A' },
      '3:0': { value: 'C' },
    };
    sheet.autoFilter = { r1: 0, c1: 0, r2: 3, c2: 0 };
    const sorted = sortRange(base, sheet, { r1: 1, c1: 0, r2: 3, c2: 0 }, 0, true);
    expect(sorted.sheets[0].cells['1:0']?.value).toBe('A');
    const filtered = applyFilter(sorted, sorted.sheets[0], 0, new Set(['B']), { r1: 1, c1: 0, r2: 3, c2: 0 });
    expect(filtered.sheets[0].hiddenRows[1]).toBe(true);
    expect(filtered.sheets[0].hiddenRows[2]).toBeUndefined();
    expect(filtered.sheets[0].hiddenRows[3]).toBe(true);
  });

  it('keeps sheet rows independent', () => {
    const sheet = createSheet('A', 0);
    const other = createSheet('B', 1);
    expect(sheet.id).not.toBe(other.id);
  });
});
