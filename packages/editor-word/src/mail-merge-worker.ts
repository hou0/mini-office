// 功能：在 Web Worker 中执行邮件合并数据导入，避免大 XLSX 解析冻结主线程；时间：2026-10-27；作者：lq。
import { readMergeWorkbook } from './mail-merge';

interface MergeImportRequest { id: number; bytes: Uint8Array; name: string }

self.onmessage = async (event: MessageEvent<MergeImportRequest>) => {
  const { id, bytes, name } = event.data;
  try {
    const workbook = await readMergeWorkbook(bytes, name, (stage, completed, total) => {
      self.postMessage({ id, kind: 'progress', stage, completed, total });
    });
    self.postMessage({ id, kind: 'done', workbook });
  } catch (cause) {
    self.postMessage({ id, kind: 'error', message: cause instanceof Error ? cause.message : String(cause) });
  }
};