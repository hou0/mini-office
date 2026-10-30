import { describe, it, expect } from 'vitest';
import { createBlankDoc, type DocRun } from '@mini-office/document-model';
import { parseMergeDirective, parseMergeField } from '../packages/editor-word/src/merge-fields';

describe('DocRun merge fields', () => {
  it('DocRun supports mergeNext / mergeNextIf / mergeSkipIf / mergeRecord / mergeSeq / mergeFormat', () => {
    const run: DocRun = {
      text: '',
      style: {},
      mergeNext: true,
      mergeNextIf: '货品代码 <> ""',
      mergeSkipIf: '批号 = ""',
      mergeRecord: true,
      mergeSeq: true,
      mergeFormat: { num: '#,##0.00', case: 'upper' },
    };
    expect(run.mergeNext).toBe(true);
    expect(run.mergeNextIf).toBe('货品代码 <> ""');
    expect(run.mergeSkipIf).toBe('批号 = ""');
    expect(run.mergeRecord).toBe(true);
    expect(run.mergeSeq).toBe(true);
    expect(run.mergeFormat?.num).toBe('#,##0.00');
    expect(run.mergeFormat?.case).toBe('upper');
  });
});

describe('parseMergeDirective', () => {
  it('识别 NEXT', () => {
    expect(parseMergeDirective(' NEXT ')).toEqual({ kind: 'next' });
  });
  it('识别 NEXTIF 条件', () => {
    expect(parseMergeDirective('NEXTIF 货品代码 <> ""')).toEqual({ kind: 'nextIf', condition: '货品代码 <> ""' });
  });
  it('识别 SKIPIF 条件', () => {
    expect(parseMergeDirective('SKIPIF 批号 = ""')).toEqual({ kind: 'skipIf', condition: '批号 = ""' });
  });
  it('识别 MERGEREC / MERGESEQ', () => {
    expect(parseMergeDirective(' MERGEREC ')).toEqual({ kind: 'record' });
    expect(parseMergeDirective(' MERGESEQ ')).toEqual({ kind: 'seq' });
  });
  it('识别 MERGEFIELD 字段名与开关', () => {
    expect(parseMergeDirective('MERGEFIELD 货品代码 \\* Upper \\# "#,##0.00" \\@ "yyyy" \\b "[" \\f "]"')).toEqual({
      kind: 'mergeField',
      name: '货品代码',
      format: { num: '#,##0.00', date: 'yyyy', case: 'upper', before: '[', after: ']' },
    });
  });
  it('非合并域返回 undefined', () => {
    expect(parseMergeDirective('PAGE \\* MERGEFORMAT')).toBeUndefined();
    expect(parseMergeDirective('DATE \\@ "d"')).toBeUndefined();
  });
});

describe('parseMergeField 兼容', () => {
  it('提取字段名，忽略开关', () => {
    expect(parseMergeField('MERGEFIELD 货品代码 \\* Upper')).toBe('货品代码');
    expect(parseMergeField(' MERGEFIELD 批号 ')).toBe('批号');
  });
});

describe('parse NEXT domain', () => {
  function insertNextField(xml: string): string {
    const field = `<w:p><w:fldSimple w:instr=" NEXT \\* MERGEFORMAT "><w:r><w:rPr/><w:t xml:space="preserve">X</w:t></w:r></w:fldSimple></w:p>`;
    return xml.replace(/<w:body[^>]*>/, (m) => m + field);
  }

  it('将 fldSimple NEXT 解析为 run.mergeNext=true', async () => {
    const { createBlankDoc } = await import('@mini-office/document-model');
    const { serializeDocx } = await import('../packages/editor-word/src/serialize');
    const { parseDocxPackage } = await import('../packages/editor-word/src/parse');
    const { OoxmlPackage } = await import('@mini-office/ooxml');
    const JSZip = (await import('jszip')).default;

    const blank = await serializeDocx(createBlankDoc(), { name: 'blank.docx' });
    const zip = await JSZip.loadAsync(blank);
    let docXml = await zip.file('word/document.xml')!.async('text');
    docXml = insertNextField(docXml);
    zip.file('word/document.xml', docXml);
    const pkg = new OoxmlPackage(zip, blank);
    const doc = await parseDocxPackage(pkg);
    const found = doc.blocks.some(
      (b) => b.type === 'paragraph' && b.runs.some((r) => r.mergeNext === true),
    );
    expect(found).toBe(true);
  });
});

describe('serialize round-trip NEXT', () => {
  it('NEXT 与 MERGEFIELD 格式开关重新打开后保留', async () => {
    const { createBlankDoc } = await import('@mini-office/document-model');
    const { serializeDocx } = await import('../packages/editor-word/src/serialize');
    const { parseDocxPackage } = await import('../packages/editor-word/src/parse');
    const { OoxmlPackage } = await import('@mini-office/ooxml');
    const JSZip = (await import('jszip')).default;

    const doc = createBlankDoc();
    const para = doc.blocks[0] as { type: string; runs: unknown[] };
    para.runs = [
      { text: 'X', style: {}, mergeNext: true },
      { text: 'Y', style: {}, mergeField: '货品代码', mergeFormat: { num: '#,##0.00', case: 'upper', before: '[', after: ']' } },
    ];
    const bytes = await serializeDocx(doc, { name: 'rt.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const pkg = new OoxmlPackage(zip, bytes);
    const reopened = await parseDocxPackage(pkg);
    const runs = reopened.blocks.flatMap((b) => b.type === 'paragraph' ? b.runs : []);
    const nextRun = runs.find((r) => r.mergeNext === true);
    const fieldRun = runs.find((r) => r.mergeField === '货品代码');
    expect(nextRun).toBeDefined();
    expect(nextRun?.text).toBe('«NEXT»');
    expect(fieldRun).toBeDefined();
    expect(fieldRun?.text).toBe('«货品代码»');
    expect(fieldRun?.mergeFormat?.num).toBe('#,##0.00');
    expect(fieldRun?.mergeFormat?.case).toBe('upper');
    expect(fieldRun?.mergeFormat?.before).toBe('[');
    expect(fieldRun?.mergeFormat?.after).toBe(']');
  });
});

describe('mergeNextDocument', () => {
  function makeRun(text: string, extra: Record<string, unknown> = {}): any {
    return { text, style: {}, ...extra };
  }
  function makeCell(runs: any[]): any {
    return {
      id: `c${Math.random().toString(36).slice(2)}`,
      blocks: [{ type: 'paragraph', id: `p${Math.random().toString(36).slice(2)}`, runs, style: {} }],
      colSpan: 1,
      rowSpan: 1,
    };
  }
  function recordBlock(
    tag: string,
    opts: { skipIf?: boolean; nextIf?: string; seq?: boolean; numFormat?: string } = {},
  ): any[] {
    const extraRuns: any[] = [];
    if (opts.skipIf) extraRuns.push(makeRun('«SKIPIF»', { mergeSkipIf: '批号 = "跳过"' }));
    if (opts.nextIf) extraRuns.push(makeRun('«NEXTIF»', { mergeNextIf: opts.nextIf }));
    if (opts.seq) {
      extraRuns.push(makeRun('«MERGEREC»', { mergeRecord: true }));
      extraRuns.push(makeRun('«MERGESEQ»', { mergeSeq: true }));
    }
    const firstCellRuns = [
      makeRun('«NEXT»', { mergeNext: true }),
      makeRun(`«货品代码${tag}»`, { mergeField: '货品代码', ...(opts.numFormat ? { mergeFormat: { num: opts.numFormat } } : {}) }),
      ...extraRuns,
    ];
    return [
      {
        id: `r${tag}a`,
        cells: [
          makeCell(firstCellRuns),
          makeCell([makeRun(`«批号${tag}»`, { mergeField: '批号' })]),
        ],
      },
      {
        id: `r${tag}b`,
        cells: [
          makeCell([makeRun(`«数量${tag}»`, { mergeField: '数量', ...(opts.numFormat ? { mergeFormat: { num: opts.numFormat } } : {}) })]),
          makeCell([makeRun('', {})]),
        ],
      },
    ];
  }
  function makeTemplate(opts: { skipIf?: boolean; nextIf?: string; seq?: boolean; numFormat?: string } = {}): any {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'table',
        id: 't1',
        columnWidthsPx: [200, 200],
        rows: [
          { id: 'h', cells: [makeCell([makeRun('货品代码', {})]), makeCell([makeRun('批号', {})])] },
          ...recordBlock('1', opts),
          ...recordBlock('2', opts),
          ...recordBlock('3', opts),
          { id: 'tail', cells: [makeCell([makeRun('合计', {})]), makeCell([makeRun('', {})])] },
        ],
      },
    ];
    return doc;
  }
  function makeRecord(rowNumber: number, values: Record<string, string>): any {
    return { rowNumber, values, raw: values, errors: [] };
  }
  function tableRowTexts(doc: any): string[][] {
    // 功能：分页输出后每页一张表，这里跨表收集行文本；时间：2026-10-28；作者：lq。
    return doc.blocks
      .filter((b: any) => b.type === 'table')
      .flatMap((table: any) => table.rows.map((row: any) =>
        row.cells.map((cell: any) =>
          cell.blocks.flatMap((b: any) => (b.type === 'paragraph' ? b.runs.map((r: any) => r.text ?? '') : [])).join(''),
        ),
      ));
  }

  it('1 条记录：块1填值，块2/3空白，汇总区保留', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate(), [makeRecord(1, { 货品代码: 'ABC', 批号: '001', 数量: '10' })]);
    const texts = tableRowTexts(merged);
    expect(texts.length).toBe(8);
    expect(texts[1][0]).toContain('ABC');
    expect(texts[1][1]).toBe('001');
    expect(texts[2][0]).toBe('10');
    expect(texts[3][0]).not.toContain('ABC');
    expect(texts[5][0]).not.toContain('ABC');
    expect(texts[7][0]).toBe('合计');
  });

  it('3 条记录：每块各填一条，一页连续显示', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate(), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
      makeRecord(3, { 货品代码: 'C', 批号: '3', 数量: '3' }),
    ]);
    const texts = tableRowTexts(merged);
    expect(texts[1][0]).toContain('A');
    expect(texts[3][0]).toContain('B');
    expect(texts[5][0]).toContain('C');
    expect(texts[7][0]).toBe('合计');
  });

  it('5 条记录：按每页 3 块分两页输出，第 2 页版式与第 1 页一致', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate(), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
      makeRecord(3, { 货品代码: 'C', 批号: '3', 数量: '3' }),
      makeRecord(4, { 货品代码: 'D', 批号: '4', 数量: '4' }),
      makeRecord(5, { 货品代码: 'E', 批号: '5', 数量: '5' }),
    ]);
    const tables = merged.blocks.filter((b: any) => b.type === 'table');
    expect(tables.length).toBe(2);
    expect(merged.blocks.filter((b: any) => b.type === 'pageBreak').length).toBe(1);
    const texts = tableRowTexts(merged);
    expect(texts.length).toBe(16);
    expect(texts[1][0]).toContain('A');
    expect(texts[3][0]).toContain('B');
    expect(texts[5][0]).toContain('C');
    expect(texts[7][0]).toBe('合计');
    expect(texts[9][0]).toContain('D');
    expect(texts[11][0]).toContain('E');
    expect(texts[13][0]).toBe('');
    expect(texts[15][0]).toBe('合计');
  });

  it('SKIPIF：条件为真时记录被跳过，不占块', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate({ skipIf: true }), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '跳过', 数量: '2' }),
      makeRecord(3, { 货品代码: 'C', 批号: '3', 数量: '3' }),
    ]);
    const texts = tableRowTexts(merged);
    // B 被跳过：块1=A，块2=C，块3空白
    expect(texts[1][0]).toContain('A');
    expect(texts[3][0]).toContain('C');
    expect(texts[5][0]).not.toContain('B');
    expect(texts[5][0]).toBe('');
  });

  it('NEXTIF：条件为真时跳到下一条记录下一块', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate({ nextIf: '货品代码 <> "X"' }), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
      makeRecord(3, { 货品代码: 'C', 批号: '3', 数量: '3' }),
    ]);
    const texts = tableRowTexts(merged);
    expect(texts[1][0]).toContain('A');
    expect(texts[3][0]).toContain('B');
    expect(texts[5][0]).toContain('C');
  });

  it('NEXTIF：条件为假时下一块继续复用当前记录', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate({ nextIf: '货品代码 = "X"' }), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
    ]);
    const texts = tableRowTexts(merged);
    // NEXTIF 恒假：第 1 页块1/2/3 都填 A；第 2 页块1 继续 A、块2 用 B
    expect(texts[1][0]).toContain('A');
    expect(texts[3][0]).toContain('A');
    expect(texts[5][0]).toContain('A');
    expect(texts[9][0]).toContain('A');
    expect(texts[11][0]).toContain('B');
  });

  it('MERGEREC/MERGESEQ：填记录序号与合并顺序号', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate({ seq: true }), [
      makeRecord(7, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(8, { 货品代码: 'B', 批号: '2', 数量: '2' }),
    ]);
    const texts = tableRowTexts(merged);
    // 块1：MERGEREC=7，MERGESEQ=1；块2：MERGEREC=8，MERGESEQ=2；块3空白
    expect(texts[1][0]).toContain('7');
    expect(texts[1][0]).toContain('1');
    expect(texts[3][0]).toContain('8');
    expect(texts[3][0]).toContain('2');
    expect(texts[5][0]).toBe('');
  });

  it('格式开关：\\# 千分位、\\* Upper、\\b/\\f 空值不插前后缀', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');

    const template = makeTemplate({ numFormat: '#,##0.00' });
    const merged = mergeNextDocument(template, [makeRecord(1, { 货品代码: 'ABC', 批号: '1', 数量: '1234567.8' })]);
    const texts = tableRowTexts(merged);
    expect(texts[1][0]).toContain('ABC');
    expect(texts[2][0]).toBe('1,234,567.80');
  });

  it('两个 NEXT 表：按页成对输出并共用同一页的记录窗口', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const doc = createBlankDoc();
    const makeTable = (id: string) => ({
      type: 'table' as const,
      id,
      columnWidthsPx: [200, 200],
      rows: [
        { id: `${id}h`, cells: [makeCell([makeRun('货品代码', {})]), makeCell([makeRun('批号', {})])] },
        ...recordBlock(id, {}),
        { id: `${id}tail`, cells: [makeCell([makeRun('合计', {})]), makeCell([makeRun('', {})])] },
      ],
    });
    doc.blocks = [makeTable('t1'), makeTable('t2')];
    const merged = mergeNextDocument(doc, [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
    ]);
    const blocks = merged.blocks as any[];
    const tables = blocks.filter((b) => b.type === 'table');
    expect(tables.length).toBe(4);
    expect(blocks.filter((b) => b.type === 'pageBreak').length).toBe(1);
    // 第 1 页两张表都填 A，第 2 页两张表都填 B，记录不被重复消耗
    expect(tableRowTexts({ blocks: [tables[0]] })[1][0]).toContain('A');
    expect(tableRowTexts({ blocks: [tables[1]] })[1][0]).toContain('A');
    expect(tableRowTexts({ blocks: [tables[2]] })[1][0]).toContain('B');
    expect(tableRowTexts({ blocks: [tables[3]] })[1][0]).toContain('B');
  });

  it('两个 NEXT 表记录块数不一致时抛错，避免静默丢记录', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const doc = createBlankDoc();
    const makeBlock = (tag: string, i: number): any => ({
      id: `r${tag}${i}`,
      cells: [makeCell([makeRun('«NEXT»', { mergeNext: true }), makeRun(`«货品代码${tag}${i}»`, { mergeField: '货品代码' })])],
    });
    const makeTable = (id: string, blockCount: number): any => ({
      type: 'table' as const,
      id,
      columnWidthsPx: [200, 200],
      rows: [
        { id: `${id}h`, cells: [makeCell([makeRun('货品代码', {})])] },
        ...Array.from({ length: blockCount }, (_, i) => makeBlock(id, i)),
        { id: `${id}tail`, cells: [makeCell([makeRun('合计', {})])] },
      ],
    });
    doc.blocks = [makeTable('t1', 2), makeTable('t2', 4)];
    expect(() => mergeNextDocument(doc, [
      makeRecord(1, { 货品代码: 'A' }),
      makeRecord(2, { 货品代码: 'B' }),
      makeRecord(3, { 货品代码: 'C' }),
      makeRecord(4, { 货品代码: 'D' }),
    ])).toThrow(/块数|NEXT/);
  });

  it('hasNextMergeField：含 NEXT 模板返回 true，普通模板返回 false', async () => {
    const { hasNextMergeField } = await import('../packages/editor-word/src/mail-merge');
    expect(hasNextMergeField(makeTemplate())).toBe(true);
    const plain = createBlankDoc();
    plain.blocks = [{ type: 'table', id: 't2', columnWidthsPx: [100], rows: [{ id: 'r', cells: [makeCell([makeRun('x', {})])] }] }];
    expect(hasNextMergeField(plain)).toBe(false);
  });

  it('多页输出：每页表格与单元格 id 唯一，避免 React 重复键导致内容被丢弃', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const merged = mergeNextDocument(makeTemplate(), [
      makeRecord(1, { 货品代码: 'A', 批号: '1', 数量: '1' }),
      makeRecord(2, { 货品代码: 'B', 批号: '2', 数量: '2' }),
      makeRecord(3, { 货品代码: 'C', 批号: '3', 数量: '3' }),
      makeRecord(4, { 货品代码: 'D', 批号: '4', 数量: '4' }),
    ]);
    const tables = (merged.blocks as any[]).filter((b) => b.type === 'table');
    expect(tables.length).toBe(2);
    // 功能：React 按 key 区分子节点，表格 id 重复会报重复键并可能丢内容；时间：2026-10-28；作者：lq。
    expect(new Set(tables.map((t) => t.id)).size).toBe(tables.length);
    const cellIds = tables.flatMap((t) => t.rows.flatMap((r: any) => r.cells.map((c: any) => c.id)));
    expect(new Set(cellIds).size).toBe(cellIds.length);
    // 第 1 页三个记录块分别填 A/B/C，第 2 页填 D
    const first = tableRowTexts({ blocks: [tables[0]] });
    expect(first[1][0]).toContain('A');
    expect(first[3][0]).toContain('B');
    expect(first[5][0]).toContain('C');
    const second = tableRowTexts({ blocks: [tables[1]] });
    expect(second[1][0]).toContain('D');
  });

  it('多页输出：非 NEXT 块（标题段落）在每一页重复，且各页 id 唯一', async () => {
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const title = { type: 'paragraph' as const, id: 'title', runs: [{ text: '原料检验原始记录', style: {} }], style: {} };
    const block = (tag: string): any => ({
      id: `r${tag}`,
      cells: [makeCell([makeRun('«NEXT»', { mergeNext: true }), makeRun(`«货品代码${tag}»`, { mergeField: '货品代码' })])],
    });
    const table = {
      type: 'table' as const,
      id: 't1',
      columnWidthsPx: [200, 200],
      rows: [{ id: 'h', cells: [makeCell([makeRun('货品代码', {})])] }, block('1'), block('2'), block('3')],
    };
    const doc = createBlankDoc();
    doc.blocks = [title, table];
    const merged = mergeNextDocument(doc, [
      makeRecord(1, { 货品代码: 'A' }),
      makeRecord(2, { 货品代码: 'B' }),
      makeRecord(3, { 货品代码: 'C' }),
      makeRecord(4, { 货品代码: 'D' }),
    ]);
    // 功能：表单标题应在每一页都保留，否则第 2 页起只剩表格；时间：2026-10-29；作者：lq。
    const titles = (merged.blocks as any[]).filter((b) => b.type === 'paragraph' && b.runs.some((r: any) => r.text === '原料检验原始记录'));
    expect(titles.length).toBe(2);
    expect(new Set(titles.map((t) => t.id)).size).toBe(titles.length);
  });
});

describe('merge field 格式串引号转义', () => {
  it('含引号的日期格式串不会被双重转义成 &amp;quot;', async () => {
    const { createBlankDoc } = await import('@mini-office/document-model');
    const { serializeDocx } = await import('../packages/editor-word/src/serialize');
    const JSZip = (await import('jszip')).default;

    const doc = createBlankDoc();
    const para = doc.blocks[0] as { type: string; runs: unknown[] };
    para.runs = [
      { text: 'X', style: {}, mergeField: '日期', mergeFormat: { date: 'yyyy"年"M"月"' } },
    ];
    const bytes = await serializeDocx(doc, { name: 'fmt.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    const instr = (xml.match(/<w:fldSimple[^>]*w:instr="([^"]*)"/) ?? ['', ''])[1];
    // 修复前会把格式串里的引号双重转义成 &amp;quot;
    expect(instr).not.toContain('&amp;quot;');
    // XML 实体解码后，Word 读到的字段指令应保留原始引号
    const decoded = instr
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    expect(decoded).toContain('\\@ "yyyy"年"M"月"');
  });
});
