import {
  ErrorCodes,
  OfficeError,
  baseName,
  type DocumentSnapshot,
  type DocumentSnapshotMeta,
  type FontInfo,
} from '@mini-office/shared';
import { fileTypeFilters, toUint8Array, type OpenedFile, type PlatformAdapter, type SaveOutcome, type SaveTarget } from './types';
import { getSnapshot, listSnapshots, putSnapshot, removeSnapshot } from './snapshots';

async function tauriInvoke<T>(command: string, args?: Record<string, unknown>): Promise<T> {
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<T>(command, args);
}

async function pickOpenPath(extensions: string[]): Promise<string | null> {
  const { open } = await import('@tauri-apps/plugin-dialog');
  const selected = await open({
    multiple: false,
    directory: false,
    filters: fileTypeFilters(extensions),
  });
  return typeof selected === 'string' ? selected : null;
}

async function pickSavePath(suggestedName: string, extensions: string[]): Promise<string | null> {
  const { save } = await import('@tauri-apps/plugin-dialog');
  const selected = await save({
    defaultPath: suggestedName,
    filters: fileTypeFilters(extensions),
  });
  return typeof selected === 'string' ? selected : null;
}

function extensionOfName(name: string): string[] {
  const ext = name.split('.').pop();
  return ext ? [ext] : [];
}

export class TauriPlatformAdapter implements PlatformAdapter {
  readonly kind = 'tauri' as const;

  async openPath(path: string): Promise<OpenedFile | null> {
    try {
      const result = await tauriInvoke<ArrayBuffer | number[]>('read_file', { path });
      return {
        name: baseName(path) || '未命名文档',
        bytes: toUint8Array(result),
        path,
      };
    } catch (error) {
      throw new OfficeError(
        ErrorCodes.PARSE_FAILED,
        `无法读取文件：${baseName(path)}`,
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async openFile(extensions: string[]): Promise<OpenedFile | null> {
    const path = await pickOpenPath(extensions);
    if (!path) return null;
    return this.openPath(path);
  }

  /** 功能：调用 Rust 兼容桥转换 DOC/XLS/PPT，并读取生成的 OOXML 副本；时间：2026-10-27；作者：lq。 */
  async convertLegacy(path: string): Promise<OpenedFile | null> {
    try {
      const converted = await tauriInvoke<string>('convert_legacy_document', { input: path });
      return this.openPath(converted);
    } catch (error) {
      throw new OfficeError(ErrorCodes.PARSE_FAILED, `无法转换旧版文件：${baseName(path)}`, error instanceof Error ? error.message : String(error));
    }
  }

  async saveBytes(bytes: Uint8Array, target: SaveTarget, mime: string): Promise<SaveOutcome> {
    let path = target.path;
    if (!path || target.forcePicker) {
      const picked = await pickSavePath(target.suggestedName, extensionOfName(target.suggestedName));
      if (!picked) throw new OfficeError(ErrorCodes.CANCELLED, '保存已取消');
      path = picked;
    }
    try {
      const result = await tauriInvoke<{ path: string; backupPath?: string | null }>('write_file', {
        path,
        bytes: Array.from(bytes),
      });
      return {
        displayName: baseName(result.path || path),
        path: result.path || path,
        location: target.forcePicker ? 'disk' : 'in-place',
      };
    } catch (error) {
      throw new OfficeError(
        ErrorCodes.SAVE_FAILED,
        `保存失败：${baseName(path)}`,
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async listFonts(): Promise<FontInfo[]> {
    try {
      const fonts = await tauriInvoke<Array<{ family: string; style?: string; path: string }>>('list_fonts');
      return fonts.map((font) => ({
        family: font.family,
        style: font.style,
        source: font.path,
        hasChinese: /[\u3400-\u9fff]/.test(font.family) || /CJK|Hei|Song|Kai|Fang|YaHei|PingFang/.test(font.family),
      }));
    } catch {
      return [];
    }
  }

  async saveSnapshot(snapshot: DocumentSnapshot): Promise<void> {
    await putSnapshot(snapshot);
  }

  async listSnapshots(): Promise<DocumentSnapshotMeta[]> {
    return listSnapshots();
  }

  async loadSnapshot(id: string): Promise<DocumentSnapshot | null> {
    return getSnapshot(id);
  }

  async deleteSnapshot(id: string): Promise<void> {
    await removeSnapshot(id);
  }

  async revealInFolder(path: string): Promise<void> {
    try {
      const { revealItemInDir } = await import('@tauri-apps/plugin-opener');
      await revealItemInDir(path);
    } catch {
      // ignore
    }
  }

  async readDroppedFile(file: File): Promise<OpenedFile> {
    const bytes = new Uint8Array(await file.arrayBuffer());
    return { name: file.name, bytes };
  }
}
