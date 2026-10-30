import type { DocBlock, DocModel } from '@mini-office/document-model';

/** 功能：按真实节边界分割报告，供重新打开、分页与保存共用；时间：2026-10-24；作者：lq。 */
export function documentSections(doc: DocModel): DocModel[] {
  const result: DocModel[] = [];
  let blocks: DocBlock[] = [];
  for (const block of doc.blocks) {
    blocks.push(block);
    if (block.type === 'paragraph' && block.sectionEnd) {
      // continuous 节不强制翻页，块继续留在当前节内；时间：2026-10-30；作者：hou0。
      if (block.sectionEnd.page.type === 'continuous') continue;
      result.push({ ...doc, ...block.sectionEnd, blocks });
      blocks = [];
    }
  }
  if (blocks.length || !result.length) result.push({ ...doc, blocks });
  return result;
}
