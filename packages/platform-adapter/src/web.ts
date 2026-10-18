import { ErrorCodes, OfficeError, baseName, downloadBlob, type DocumentSnapshot, type DocumentSnapshotMeta, type FontInfo } from '@mini-office/shared';
import { acceptFilters, fileTypeFilters, type OpenedFile, type PlatformAdapter, type SaveOutcome, type SaveTarget } from './types';
import { DEFAULT_AVAILABLE_FONTS, detectWebFonts } from './fonts';
import { getSnapshot, listSnapshots, putSnapshot, removeSnapshot } from './snapshots';

function abortError(): OfficeError {
  return new OfficeError(ErrorCodes.CANCELLED, '操作已取消');
}

async function pickWithInput(extensions: string[]): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = extensions.map((ext) => (ext.startsWith('.') ? ext : `.${ext}`)).join(',');
    input.style.position = 'fixed';
    input.style.left = '-9999px';
    document.body.appendChild(input);
    let settled = false;
    const finish = (file: File | null) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(file);
    };
    input.addEventListener('change', () => finish(input.files?.[0] ?? null));
    input.addEventListener('cancel', () => finish(null));
    window.addEventListener(
      'focus',
      () => {
        setTimeout(() => {
          if (!settled) finish(input.files?.[0] ?? null);
        }, 400);
      },
      { once: true },
    );
    input.click();
  });
}

async function fileToOpened(file: File, handle?: FileSystemFileHandle, path?: string): Promise<OpenedFile> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  return {
    name: file.name || path?.split(/[\\/]/).pop() || '未命名文档',
    bytes,
    path,
    handle,
    lastModified: file.lastModified,
  };
}

export class WebPlatformAdapter implements PlatformAdapter {
  readonly kind = 'web' as const;
  private fontCache: FontInfo[] | null = null;

  async openFile(extensions: string[]): Promise<OpenedFile | null> {
    if (typeof window.showOpenFilePicker === 'function') {
      try {
        const handles = await window.showOpenFilePicker({
          multiple: false,
          types: [
            {
              description: 'Office 文件',
              accept: acceptFilters(extensions) as unknown as Record<string, string[]>,
            },
          ],
          excludeAcceptAllOption: false,
        });
        const handle = handles[0];
        if (!handle) return null;
        const file = await handle.getFile();
        return fileToOpened(file, handle);
      } catch (error) {
        if ((error as DOMException)?.name === 'AbortError') return null;
        // Permission or unsupported configuration: fall back to <input>.
      }
    }
    const file = await pickWithInput(extensions);
    if (!file) return null;
    return fileToOpened(file);
  }

  async openPath(): Promise<OpenedFile | null> {
    return null;
  }

  async saveBytes(bytes: Uint8Array, target: SaveTarget, mime: string): Promise<SaveOutcome> {
    const blob = new Blob([bytes.slice().buffer], { type: mime });

    if (target.handle && !target.forcePicker) {
      try {
        const writable = await target.handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return {
          displayName: target.handle.name || target.suggestedName,
          handle: target.handle,
          location: 'in-place',
        };
      } catch (error) {
        if ((error as DOMException)?.name === 'AbortError') throw abortError();
        throw new OfficeError(
          ErrorCodes.PERMISSION_DENIED,
          '没有写入该文件的权限，请使用“另存为”选择新位置。',
          error instanceof Error ? error.message : String(error),
        );
      }
    }

    if (typeof window.showSaveFilePicker === 'function') {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: target.suggestedName,
          types: [
            {
              description: 'Office 文件',
              accept: acceptFilters([target.suggestedName.split('.').pop() ?? '']),
            },
          ],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return { displayName: handle.name, handle, location: 'disk' };
      } catch (error) {
        if ((error as DOMException)?.name === 'AbortError') throw abortError();
        // fall through to download
      }
    }

    downloadBlob(blob, target.suggestedName);
    return { displayName: target.suggestedName, location: 'download' };
  }

  async listFonts(): Promise<FontInfo[]> {
    if (this.fontCache) return this.fontCache;
    const detected = detectWebFonts();
    const merged = new Map<string, FontInfo>();
    for (const font of [...DEFAULT_AVAILABLE_FONTS, ...detected]) {
      if (!merged.has(font.family)) merged.set(font.family, font);
    }
    this.fontCache = [...merged.values()];
    return this.fontCache;
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

  async revealInFolder(): Promise<void> {
    // Browsers cannot reveal downloads in the file manager.
  }

  async readDroppedFile(file: File): Promise<OpenedFile> {
    return fileToOpened(file);
  }
}

export function suggestedNameFor(path: string | undefined, fallback: string): string {
  return path ? baseName(path) : fallback;
}
