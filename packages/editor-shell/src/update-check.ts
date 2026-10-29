import { create } from 'zustand';
import { APP_VERSION } from '@mini-office/shared';

// 功能：查询 GitHub Releases 并比较当前版本，支持小红点与启动提示；时间：2026-10-29；作者：lq。
export const RELEASES_API_URL = 'https://api.github.com/repos/lq0910/mini-office/releases?per_page=30';

/** 自动检查的最小间隔。手动「检查更新」不受此限制。 */
export const UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000;

const DISMISSED_KEY = 'mini-office.update.dismissed.v1';
const CHECKED_AT_KEY = 'mini-office.update.checked-at.v1';
const LATEST_KEY = 'mini-office.update.latest.v1';

export type LatestRelease = {
  tagName: string;
  name: string;
  url: string;
  prerelease: boolean;
  publishedAt?: string;
};

export type UpdateStatus = 'idle' | 'checking' | 'ready' | 'none' | 'error';

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

/** 当前版本是预览版时，才允许把预览版当作更新提示；正式版用户不会被 beta 打扰。 */
const ALLOW_PRERELEASE = /-/.test(APP_VERSION);

let cache: { at: number; value: LatestRelease | null } | null = null;
let inflight: Promise<LatestRelease | null> | null = null;

/** 拉取最新版本；没有更新时返回 null。网络或解析失败会抛出，交由调用方区分"无更新"与"查不到"。 */
export async function fetchLatestRelease(options: { force?: boolean } = {}): Promise<LatestRelease | null> {
  if (!options.force && cache && Date.now() - cache.at < UPDATE_CHECK_INTERVAL_MS) return cache.value;
  if (inflight) return inflight;

  inflight = (async () => {
    const response = await fetch(RELEASES_API_URL, { headers: { Accept: 'application/vnd.github+json' } });
    if (!response.ok) throw new Error(`发布信息请求失败（HTTP ${response.status}）`);
    const releases = (await response.json()) as Array<{
      tag_name?: string;
      name?: string;
      html_url?: string;
      prerelease?: boolean;
      draft?: boolean;
      published_at?: string;
    }>;
    const candidates = releases
      .filter((item) => !item.draft && item.tag_name && item.html_url)
      .filter((item) => ALLOW_PRERELEASE || !item.prerelease)
      .sort((left, right) => compareVersions(right.tag_name!, left.tag_name!));
    const latest = candidates[0];
    if (!latest || compareVersions(latest.tag_name!, APP_VERSION) <= 0) return null;
    return {
      tagName: latest.tag_name!,
      name: latest.name || latest.tag_name!,
      url: latest.html_url!,
      prerelease: !!latest.prerelease,
      publishedAt: latest.published_at,
    } satisfies LatestRelease;
  })();

  try {
    const value = await inflight;
    cache = { at: Date.now(), value };
    return value;
  } finally {
    inflight = null;
  }
}

function readString(key: string): string | null {
  if (typeof localStorage === 'undefined') return null;
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeString(key: string, value: string | null): void {
  if (typeof localStorage === 'undefined') return;
  try {
    if (value === null) localStorage.removeItem(key);
    else localStorage.setItem(key, value);
  } catch {
    // 持久化失败只影响"下次还会提示"，不影响本次使用。
  }
}

function readCheckedAt(): number | null {
  const raw = readString(CHECKED_AT_KEY);
  const value = raw ? Number(raw) : Number.NaN;
  return Number.isFinite(value) ? value : null;
}

/**
 * 读取上次查到的版本。
 *
 * 必须持久化：检查有 6 小时节流，若不缓存结果，重启后小红点会消失到下次联网检查为止，
 * 用户会以为"更新提示丢了"。应用升级后缓存自动失效，避免提示一个比当前更旧的版本。
 */
function readCachedLatest(): LatestRelease | null {
  const raw = readString(LATEST_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as Partial<LatestRelease>;
    if (!parsed.tagName || !parsed.url) return null;
    if (compareVersions(parsed.tagName, APP_VERSION) <= 0) {
      writeString(LATEST_KEY, null);
      return null;
    }
    return {
      tagName: parsed.tagName,
      name: parsed.name || parsed.tagName,
      url: parsed.url,
      prerelease: !!parsed.prerelease,
      publishedAt: parsed.publishedAt,
    };
  } catch {
    return null;
  }
}

function writeCachedLatest(release: LatestRelease | null): void {
  writeString(LATEST_KEY, release ? JSON.stringify(release) : null);
}

export interface UpdateStoreState {
  status: UpdateStatus;
  latest: LatestRelease | null;
  checkedAt: number | null;
  error: string | null;
  /** 用户点过「稍后提醒」的版本号；同一版本不再弹窗，但小红点保留。 */
  dismissedTag: string | null;

  check(options?: { force?: boolean }): Promise<void>;
  dismiss(): void;
  /** 用户主动打开过更新说明后清掉小红点。 */
  clearBadge(): void;
}

export const useUpdateStore = create<UpdateStoreState>((set, get) => {
  const cachedLatest = readCachedLatest();
  return {
    status: cachedLatest ? 'ready' : 'idle',
    latest: cachedLatest,
    checkedAt: readCheckedAt(),
    error: null,
    dismissedTag: readString(DISMISSED_KEY),

    async check(options = {}) {
      if (get().status === 'checking') return;
      const last = get().checkedAt;
      // 节流：自动检查六小时内只查一次，避免每次启动都打 GitHub API。
      if (!options.force && last && Date.now() - last < UPDATE_CHECK_INTERVAL_MS) return;

      set({ status: 'checking', error: null });
      try {
        const release = await fetchLatestRelease({ force: options.force });
        const checkedAt = Date.now();
        writeString(CHECKED_AT_KEY, String(checkedAt));
        writeCachedLatest(release);
        set({
          status: release ? 'ready' : 'none',
          latest: release,
          checkedAt,
          error: null,
        });
      } catch (error) {
        // 查询失败不打扰用户：小红点保持上一次的结果，关于页可手动重试。
        set({
          status: get().latest ? 'ready' : 'error',
          checkedAt: Date.now(),
          error: error instanceof Error ? error.message : '检查更新失败',
        });
      }
    },

    dismiss() {
      const tag = get().latest?.tagName;
      if (!tag) return;
      writeString(DISMISSED_KEY, tag);
      set({ dismissedTag: tag });
    },

    clearBadge() {
      const tag = get().latest?.tagName;
      if (!tag) return;
      writeString(DISMISSED_KEY, tag);
      set({ dismissedTag: tag });
    },
  };
});

/** 有可用更新（用于小红点）：只要查到新版本就亮，即使用户已经忽略弹窗。 */
export function useUpdateAvailable(): boolean {
  return useUpdateStore((state) => state.latest !== null);
}

/** 兼容旧调用点：关于页与首页只需要最新版本信息。 */
export function useLatestRelease(): LatestRelease | null {
  return useUpdateStore((state) => state.latest);
}

/** 启动时调用：不阻塞首屏，失败静默。 */
export function checkForUpdates(options: { force?: boolean } = {}): Promise<void> {
  return useUpdateStore.getState().check(options);
}
