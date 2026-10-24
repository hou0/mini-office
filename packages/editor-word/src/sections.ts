import type { DocBlock, DocModel } from '@mini-office/document-model';

/** 功能：按真实节边界分割报告，供重新打开、分页与保存共用；时间：2026-10-24；作者：lq。 */
export function documentSections(doc: DocModel): DocModel[] {
  const result: DocModel[] = [];
  let blocks: DocBlock[] = [];
  for (const block of doc.blocks) {
    blocks.push(block);
    if (block.type === 'paragraph' && block.sectionEnd) {
      result.push({ ...doc, ...block.sectionEnd, blocks });
      blocks = [];
    }
  }
  if (blocks.length || !result.length) result.push({ ...doc, blocks });
  return result;
}
