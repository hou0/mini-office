import { describe, it, expect } from 'vitest';
import { createBlankDoc } from '@mini-office/document-model';

describe('mail merge range', () => {
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
  function recordBlock(tag: string): any[] {
    return [
      {
        id: `r${tag}a`,
        cells: [
          makeCell([
            makeRun('«NEXT»', { mergeNext: true }),
            makeRun(`«货品代码${tag}»`, { mergeField: '货品代码' }),
          ]),
          makeCell([makeRun(`«批号${tag}»`, { mergeField: '批号' })]),
        ],
      },
      {
        id: `r${tag}b`,
        cells: [
          makeCell([makeRun(`«数量${tag}»`, { mergeField: '数量' })]),
          makeCell([makeRun('', {})]),
        ],
      },
    ];
  }
  function makeTemplate(): any {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'table',
        id: 't1',
        columnWidthsPx: [200, 200],
        rows: [
          { id: 'h', cells: [makeCell([makeRun('货品代码', {})]), makeCell([makeRun('批号', {})])] },
          ...recordBlock('1'),
          ...recordBlock('2'),
          ...recordBlock('3'),
          { id: 'tail', cells: [makeCell([makeRun('合计', {})]), makeCell([makeRun('', {})])] },
        ],
      },
    ];
    return doc;
  }
  function makePlainTemplate(): any {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'table',
        id: 't2',
        columnWidthsPx: [200],
        rows: [
          { id: 'h', cells: [makeCell([makeRun('名称', {})])] },
          { id: 'r', cells: [makeCell([makeRun('«名称»', { mergeField: '名称' })])] },
        ],
      },
    ];
    return doc;
  }
  function makeRecord(rowNumber: number, values: Record<string, string>): any {
    return { rowNumber, values, raw: values, errors: [] };
  }
  function records(n: number): any[] {
    return Array.from({ length: n }, (_, i) => makeRecord(i + 1, { 货品代码: `CODE-${i + 1}`, 批号: `${i + 1}`, 数量: '1' }));
  }

  it('countNextBlocks 返回 NEXT 块数，普通模板返回 0', async () => {
    const { countNextBlocks } = await import('../packages/editor-word/src/mail-merge');
    expect(countNextBlocks(makeTemplate())).toBe(3);
    expect(countNextBlocks(makePlainTemplate())).toBe(0);
  });

  describe('pageRangeToRecordRange', () => {
    async function rangeFn() {
      const m = await import('../packages/editor-word/src/mail-merge');
      return m.pageRangeToRecordRange;
    }

    it('每页 1 条时页号等于记录号', async () => {
      const pageRangeToRecordRange = await rangeFn();
      expect(pageRangeToRecordRange(1, 1, 1)).toEqual([1, 1]);
      expect(pageRangeToRecordRange(3, 5, 1)).toEqual([3, 5]);
    });

    it('每页 3 条时页号换算成记录区间', async () => {
      const pageRangeToRecordRange = await rangeFn();
      expect(pageRangeToRecordRange(1, 1, 3)).toEqual([1, 3]);
      expect(pageRangeToRecordRange(2, 3, 3)).toEqual([4, 9]);
    });

    it('起止页颠倒时自动交换', async () => {
      const pageRangeToRecordRange = await rangeFn();
      expect(pageRangeToRecordRange(3, 2, 3)).toEqual([4, 9]);
    });

    it('页号小于 1 时按第 1 页处理', async () => {
      const pageRangeToRecordRange = await rangeFn();
      expect(pageRangeToRecordRange(0, 2, 3)).toEqual([1, 6]);
    });
  });

  describe('recordsByRange', () => {
    async function rangeFn() {
      const m = await import('../packages/editor-word/src/mail-merge');
      return m.recordsByRange;
    }

    it('all 返回全部记录', async () => {
      const recordsByRange = await rangeFn();
      expect(recordsByRange(records(9), 'all', 0)).toHaveLength(9);
    });

    it('currentRecord 只取当前预览的记录', async () => {
      const recordsByRange = await rangeFn();
      const rs = records(9);
      expect(recordsByRange(rs, 'currentRecord', 0).map((r) => r.values['货品代码'])).toEqual(['CODE-1']);
      expect(recordsByRange(rs, 'currentRecord', 5).map((r) => r.values['货品代码'])).toEqual(['CODE-6']);
      expect(recordsByRange(rs, 'currentRecord', 99).map((r) => r.values['货品代码'])).toEqual(['CODE-9']);
    });

    it('recordRange 1-2 返回前两条', async () => {
      const recordsByRange = await rangeFn();
      const out = recordsByRange(records(9), 'recordRange', 0, 1, 2);
      expect(out.map((r) => r.values['货品代码'])).toEqual(['CODE-1', 'CODE-2']);
    });

    it('recordRange 3-5 返回第 3~5 条', async () => {
      const recordsByRange = await rangeFn();
      const out = recordsByRange(records(9), 'recordRange', 0, 3, 5);
      expect(out.map((r) => r.values['货品代码'])).toEqual(['CODE-3', 'CODE-4', 'CODE-5']);
    });

    it('起始大于结束自动交换', async () => {
      const recordsByRange = await rangeFn();
      const out = recordsByRange(records(9), 'recordRange', 0, 5, 2);
      expect(out.map((r) => r.values['货品代码'])).toEqual(['CODE-2', 'CODE-3', 'CODE-4', 'CODE-5']);
    });

    it('起始超出记录数返回空', async () => {
      const recordsByRange = await rangeFn();
      expect(recordsByRange(records(9), 'recordRange', 0, 99, 100)).toHaveLength(0);
    });

    it('结束超出记录数 clamp 到末尾', async () => {
      const recordsByRange = await rangeFn();
      expect(recordsByRange(records(9), 'recordRange', 0, 1, 99)).toHaveLength(9);
    });

    it('空数组返回空', async () => {
      const recordsByRange = await rangeFn();
      expect(recordsByRange([], 'all', 0)).toHaveLength(0);
      expect(recordsByRange([], 'currentRecord', 0)).toHaveLength(0);
      expect(recordsByRange([], 'recordRange', 0, 1, 2)).toHaveLength(0);
    });
  });
});
