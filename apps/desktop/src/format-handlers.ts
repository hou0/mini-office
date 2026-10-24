/** 功能：启动时只注册轻量格式代理，首次打开或保存时才加载对应解析器；时间：2026-10-24；作者：lq。 */
import {
  createBlankDoc,
  createBlankWorkbook,
  createBlankDeck,
  registerFormatHandler,
  type AnyFormatHandler,
} from '@mini-office/document-model';

/** 功能：合并并发模块加载，加载失败后允许下一次操作重试；时间：2026-10-24；作者：lq。 */
function registerDeferredHandler(
  kind: AnyFormatHandler['kind'],
  extensions: string[],
  createBlank: AnyFormatHandler['createBlank'],
  loader: () => Promise<AnyFormatHandler>,
): void {
  let pending: Promise<AnyFormatHandler> | undefined;
  const load = () => pending ??= loader().catch((error) => {
    pending = undefined;
    throw error;
  });
  registerFormatHandler({
    kind,
    extensions,
    createBlank,
    open: async (input) => (await load()).open(input),
    serialize: async (data, context) => (await load()).serialize(data, context),
  });
}

/** 功能：保留新建文档的即时响应，同时将三种格式解析代码移出首页启动路径；时间：2026-10-24；作者：lq。 */
export function registerDeferredFormats(): void {
  registerDeferredHandler('docx', ['.docx', '.docm', '.dotx'], createBlankDoc,
    () => import('../../../packages/editor-word/src/handler').then((module) => module.wordFormatHandler));
  registerDeferredHandler('xlsx', ['.xlsx', '.xlsm', '.xltx', '.csv'], createBlankWorkbook,
    () => import('../../../packages/editor-sheet/src/handler').then((module) => module.sheetFormatHandler));
  registerDeferredHandler('pptx', ['.pptx', '.pptm', '.potx'], createBlankDeck,
    () => import('../../../packages/editor-slides/src/handler').then((module) => module.slidesFormatHandler));
}
