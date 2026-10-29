/**
 * 主题预览缩略图。
 *
 * 优先用插件自带的 `preview` 图片；没有图时按主题令牌画一个迷你窗口，
 * 保证"没有设计资源的主题"也能看出底色与强调色，而不是留一个空图标。
 */
import { useEffect, useState } from 'react';
import { readPluginFile, type ThemePackOption } from '@mini-office/plugin-host';

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  webp: 'image/webp',
  gif: 'image/gif',
  svg: 'image/svg+xml',
};

export function ThemePreview({ pack, className = '' }: { pack: ThemePackOption; className?: string }) {
  const preview = pack.preview;
  const [url, setUrl] = useState<string | null>(null);

  useEffect(() => {
    setUrl(null);
    if (!preview) return;
    let cancelled = false;
    let objectUrl: string | null = null;
    void readPluginFile(pack.pluginId, preview)
      .then((bytes) => {
        if (cancelled || !bytes) return;
        const ext = preview.split('.').pop()?.toLowerCase() ?? '';
        // 复制一份到独立的 ArrayBuffer：IndexedDB 取回的是 ArrayBufferLike，
        // 直接塞进 Blob 会因 SharedArrayBuffer 的可能类型而无法通过类型检查。
        const copy = new Uint8Array(bytes.byteLength);
        copy.set(bytes);
        objectUrl = URL.createObjectURL(new Blob([copy], { type: MIME_BY_EXT[ext] ?? 'image/png' }));
        setUrl(objectUrl);
      })
      .catch(() => {
        // 读不到预览图就回退到令牌色块，不影响主题本身可用。
      });
    return () => {
      cancelled = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [pack.pluginId, preview]);

  if (url) {
    return (
      <img
        src={url}
        alt={`${pack.label} 预览`}
        data-theme-preview={pack.key}
        className={`object-cover ${className}`}
      />
    );
  }

  const tokens = pack.theme.tokens;
  const surface = tokens['--mo-surface'] ?? '#ffffff';
  const accent = tokens['--mo-accent'] ?? '#2b579a';
  const text = tokens['--mo-text-primary'] ?? '#1f2328';

  return (
    <span
      aria-hidden
      data-theme-preview={pack.key}
      className={`flex flex-col overflow-hidden ring-1 ring-black/10 ${className}`}
      style={{ background: surface }}
    >
      <span className="h-[34%] w-full" style={{ background: accent, opacity: 0.85 }} />
      <span className="flex flex-1 flex-col justify-center gap-[2px] px-[3px]">
        <span className="h-[1.5px] w-[72%]" style={{ background: text, opacity: 0.7 }} />
        <span className="h-[1.5px] w-[48%]" style={{ background: text, opacity: 0.45 }} />
      </span>
    </span>
  );
}
