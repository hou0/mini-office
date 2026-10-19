/** 功能：逐份处理未保存文档，取消或保存失败时阻止退出；时间：2026-10-18；作者：lq。 */
import type { DocumentSession } from '@mini-office/document-model';

export type CloseChoice = 'save' | 'discard' | 'cancel';

/** 功能：重新检查每一版编辑内容，避免异步保存或多文档确认期间丢失更改；时间：2026-10-18；作者：lq。 */
export async function confirmWorkspaceClose(
  getSessions: () => DocumentSession[],
  ask: (session: DocumentSession) => Promise<CloseChoice>,
  save: (id: string) => Promise<boolean>,
): Promise<boolean> {
  const discarded = new Map<string, DocumentSession['data']>();
  for (;;) {
    const sessions = getSessions();
    // 保存/导出中的文档先保持打开，避免退出时中断磁盘写入；时间：2026-10-18；作者：lq。
    if (sessions.some((session) => ['SAVING', 'EXPORTING', 'LOADING', 'VALIDATING'].includes(session.state))) return false;
    const session = sessions.find((item) => item.dirty && discarded.get(item.id) !== item.data);
    if (!session) return true;
    const data = session.data;
    const choice = await ask(session);
    if (choice === 'cancel') return false;
    if (choice === 'discard') discarded.set(session.id, data);
    else if (!(await save(session.id))) return false;
  }
}
