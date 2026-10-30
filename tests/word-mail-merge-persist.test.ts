import { describe, expect, it } from 'vitest';
import { createBlankDoc } from '@mini-office/document-model';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { parseDocx } from '../packages/editor-word/src/parse';

const PANEL_STATE = {
  conditions: [{ field: '结果', op: 'eq', value: '合格' }],
  search: '张',
  sort: '编号',
  descending: true,
  allowEmpty: false,
  namePattern: '报告_{序号}',
  output: 'docx-zip',
};

describe('Word 邮件合并面板状态持久化', () => {
  it('普通保存把面板状态写入 settings.xml 文档变量，重新打开自动恢复', async () => {
    const doc = createBlankDoc();
    doc.meta.mailMergeState = JSON.stringify(PANEL_STATE);
    const bytes = await serializeDocx(doc, { name: '模板.docx', mailMergeState: doc.meta.mailMergeState });
    const reopened = await parseDocx(bytes);
    expect(reopened.meta.mailMergeState).toBe(JSON.stringify(PANEL_STATE));
  });

  it('成品（materialized）输出不携带面板状态', async () => {
    const doc = createBlankDoc();
    doc.meta.mailMergeState = JSON.stringify(PANEL_STATE);
    const bytes = await serializeDocx(doc, {
      name: '成品.docx',
      mailMergeState: doc.meta.mailMergeState,
      materialized: true,
    });
    const reopened = await parseDocx(bytes);
    expect(reopened.meta.mailMergeState).toBeUndefined();
  });

  it('无面板状态时 settings.xml 保持干净', async () => {
    const doc = createBlankDoc();
    const bytes = await serializeDocx(doc, { name: '模板.docx' });
    const reopened = await parseDocx(bytes);
    expect(reopened.meta.mailMergeState).toBeUndefined();
  });

  it('面板状态可包含多种条件与转义字符，JSON 往返无损', async () => {
    const state = {
      conditions: [
        { field: '姓名', op: 'contains', value: '张"三' },
        { field: '金额', op: 'gte', value: '1,000.5' },
      ],
      search: 'a&b<c>',
      sort: '',
      descending: false,
      allowEmpty: true,
      namePattern: '报告_{姓名}_{序号}',
      output: 'print',
    };
    const doc = createBlankDoc();
    doc.meta.mailMergeState = JSON.stringify(state);
    const bytes = await serializeDocx(doc, { name: '模板.docx', mailMergeState: doc.meta.mailMergeState });
    const reopened = await parseDocx(bytes);
    expect(reopened.meta.mailMergeState).toBe(JSON.stringify(state));
  });
});
