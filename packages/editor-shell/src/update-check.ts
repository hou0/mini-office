import { useEffect, useState } from 'react';
import { APP_VERSION } from '@mini-office/shared';

// 功能：查询 GitHub Releases 中的正式版和预览版并比较当前版本；时间：2026-10-28；作者：lq。
export const RELEASES_API_URL = 'https://api.github.com/repos/lq0910/mini-office/releases?per_page=30';

export type LatestRelease = {
  tagName: string;
  name: string;
  url: string;
  prerelease: boolean;
};

let cachedPromise: Promise<LatestRelease | null> | null = null;

function versionParts(version: string): [number, number, number, string[]] | null {
  const match = version.replace(/^v/, '').match(/^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?$/);
  if (!match) return null;
  return [Number(match[1]), Number(match[2]), Number(match[3]), match[4]?.split('.') ?? []];
}

function compareVersions(left: string, right: string): number {
  const a = versionParts(left);
  const b = versionParts(right);
  if (!a || !b) return 0;
  for (let index = 0; index < 3; index += 1) {
    if (a[index] !== b[index]) return a[index] > b[index] ? 1 : -1;
  }
  if (a[3].length === 0 && b[3].length > 0) return 1;
  if (a[3].length > 0 && b[3].length === 0) return -1;
  for (let index = 0; index < Math.max(a[3].length, b[3].length); index += 1) {
    const av = a[3][index];
    const bv = b[3][index];
    if (av === undefined) return -1;
    if (bv === undefined) return 1;
    if (av === bv) continue;
    const an = /^\d+$/.test(av);
    const bn = /^\d+$/.test(bv);
    if (an && bn) return Number(av) > Number(bv) ? 1 : -1;
    if (an !== bn) return an ? -1 : 1;
    return av > bv ? 1 : -1;
  }
  return 0;
}

export function useLatestRelease(): LatestRelease | null {
  const [release, setRelease] = useState<LatestRelease | null>(null);
  useEffect(() => {
    if (cachedPromise) {
      void cachedPromise.then(setRelease);
      return;
    }
    cachedPromise = fetch(RELEASES_API_URL, { headers: { Accept: 'application/vnd.github+json' } })
      .then(async (response) => {
        if (!response.ok) return null;
        const releases = await response.json() as Array<{ tag_name?: string; name?: string; html_url?: string; prerelease?: boolean; draft?: boolean }>;
        const candidates = releases
          .filter((item) => !item.draft && item.tag_name && item.html_url)
          .sort((left, right) => compareVersions(right.tag_name!, left.tag_name!));
        const latest = candidates[0];
        if (!latest || compareVersions(latest.tag_name!, APP_VERSION) <= 0) return null;
        return { tagName: latest.tag_name!, name: latest.name || latest.tag_name!, url: latest.html_url!, prerelease: !!latest.prerelease };
      })
      .catch(() => null);
    void cachedPromise.then(setRelease);
  }, []);
  return release;
}
