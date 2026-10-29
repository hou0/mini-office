/**
 * GitHub 插件仓库客户端。
 *
 * 官方插件以「GitHub 仓库 + index.json」的方式分发，替代传统插件商店：
 * 不需要服务器、不需要审核后台、不需要支付与统计，维护成本接近零。
 *
 * 拉取策略：
 * 1. 优先走 GitHub Contents API，带 ETag 条件请求。命中 304 不消耗配额。
 * 2. API 限流或不可用时，回退到 raw.githubusercontent.com 的静态文件。
 * 3. 索引内容按 `schemaVersion` 校验，结构不符直接拒绝，避免脏数据进入 UI。
 */
import type { PluginIndex, PluginIndexEntry, PluginRepositoryRef } from '@mini-office/plugin-api';
import { PLUGIN_MAX_UNCOMPRESSED_BYTES } from '@mini-office/plugin-api';

const CACHE_KEY = 'mini-office.plugin-index-cache.v1';

interface IndexCache {
  repoKey: string;
  etag: string;
  fetchedAt: number;
  index: PluginIndex;
}

export class PluginIndexError extends Error {
  readonly hint: string;
  constructor(message: string, hint: string) {
    super(message);
    this.name = 'PluginIndexError';
    this.hint = hint;
  }
}

function repoKeyOf(repo: PluginRepositoryRef): string {
  return `${repo.owner}/${repo.repo}@${repo.branch}:${repo.indexPath}`;
}

function readCache(): IndexCache | null {
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as IndexCache;
    if (!parsed || typeof parsed !== 'object' || !parsed.index) return null;
    return parsed;
  } catch {
    return null;
  }
}

function writeCache(cache: IndexCache): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
  } catch {
    // 缓存写不进去不影响主流程。
  }
}

/** 校验索引结构。返回 null 表示不合法。 */
export function parsePluginIndex(raw: unknown): PluginIndex | null {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return null;
  const source = raw as Record<string, unknown>;
  if (source.schemaVersion !== 1) return null;
  if (!Array.isArray(source.plugins)) return null;

  const plugins: PluginIndexEntry[] = [];
  for (const item of source.plugins) {
    if (typeof item !== 'object' || item === null) continue;
    const entry = item as Record<string, unknown>;
    const id = typeof entry.id === 'string' ? entry.id.trim() : '';
    const name = typeof entry.name === 'string' ? entry.name.trim() : '';
    const version = typeof entry.version === 'string' ? entry.version.trim() : '';
    const publisher = typeof entry.publisher === 'string' ? entry.publisher.trim() : '';
    const download = typeof entry.download === 'string' ? entry.download.trim() : '';
    const sha256 = typeof entry.sha256 === 'string' ? entry.sha256.trim().toLowerCase() : '';
    const tier = entry.tier;
    if (!id || !name || !version || !publisher || !download) continue;
    if (!/^[0-9a-f]{64}$/.test(sha256)) continue;
    if (tier !== 'declarative' && tier !== 'sandboxed' && tier !== 'trusted') continue;

    const permissions = Array.isArray(entry.permissions)
      ? entry.permissions.filter((value): value is string => typeof value === 'string')
      : [];

    const parsed: PluginIndexEntry = {
      id,
      name,
      version,
      publisher,
      tier,
      permissions: permissions as PluginIndexEntry['permissions'],
      download,
      sha256,
    };
    if (typeof entry.description === 'string') parsed.description = entry.description;
    if (typeof entry.preview === 'string') parsed.preview = entry.preview;
    if (typeof entry.size === 'number') parsed.size = entry.size;
    if (Array.isArray(entry.tags)) parsed.tags = entry.tags.filter((value): value is string => typeof value === 'string');
    if (typeof entry.updatedAt === 'string') parsed.updatedAt = entry.updatedAt;
    plugins.push(parsed);
  }

  const index: PluginIndex = { schemaVersion: 1, plugins };
  if (typeof source.updatedAt === 'string') index.updatedAt = source.updatedAt;
  return index;
}

export interface FetchIndexResult {
  index: PluginIndex;
  /** `network` 为真实拉取，`cache` 为 304 命中，`offline` 为网络失败后的降级。 */
  origin: 'network' | 'cache' | 'offline';
  fetchedAt: number;
}

/** 拉取插件索引。 */
export async function fetchPluginIndex(
  repo: PluginRepositoryRef,
  options: { force?: boolean } = {},
): Promise<FetchIndexResult> {
  const cache = readCache();
  const cacheMatches = cache?.repoKey === repoKeyOf(repo);
  const headers: Record<string, string> = {
    Accept: 'application/vnd.github.raw',
    'X-GitHub-Api-Version': '2022-11-28',
  };
  if (!options.force && cacheMatches && cache) headers['If-None-Match'] = cache.etag;

  const apiUrl = `https://api.github.com/repos/${repo.owner}/${repo.repo}/contents/${encodeURI(
    repo.indexPath,
  )}?ref=${encodeURIComponent(repo.branch)}`;

  try {
    const response = await fetch(apiUrl, { headers });
    if (response.status === 304 && cacheMatches && cache) {
      return { index: cache.index, origin: 'cache', fetchedAt: cache.fetchedAt };
    }
    if (response.ok) {
      const parsed = parsePluginIndex(await response.json());
      if (!parsed) {
        throw new PluginIndexError('插件索引格式不正确', '请检查仓库中的 index.json 是否符合 schemaVersion: 1 规范。');
      }
      const fetchedAt = Date.now();
      writeCache({ repoKey: repoKeyOf(repo), etag: response.headers.get('ETag') ?? '', fetchedAt, index: parsed });
      return { index: parsed, origin: 'network', fetchedAt };
    }
    if (response.status === 404) {
      throw new PluginIndexError(
        '找不到插件索引',
        `仓库 ${repo.owner}/${repo.repo} 的 ${repo.branch} 分支下没有 ${repo.indexPath}。`,
      );
    }
    if (response.status === 403 || response.status === 429) {
      // 未认证的 GitHub API 限流是 60 次/小时，此时回退到静态文件。
      const fallback = await fetchRawIndex(repo);
      if (fallback) return fallback;
      throw new PluginIndexError(
        'GitHub API 请求受限',
        '未认证的 GitHub API 每小时限流 60 次。已尝试改用静态地址但仍失败，请稍后重试。',
      );
    }
    throw new PluginIndexError(`拉取插件索引失败（HTTP ${response.status}）`, '请检查网络连接后重试。');
  } catch (error) {
    if (error instanceof PluginIndexError) throw error;
    const fallback = await fetchRawIndex(repo);
    if (fallback) return fallback;
    if (cacheMatches && cache) {
      return { index: cache.index, origin: 'offline', fetchedAt: cache.fetchedAt };
    }
    throw new PluginIndexError(
      '无法连接 GitHub',
      '请检查网络或代理设置。插件列表需要访问 api.github.com。',
    );
  }
}

async function fetchRawIndex(repo: PluginRepositoryRef): Promise<FetchIndexResult | null> {
  try {
    const url = `https://raw.githubusercontent.com/${repo.owner}/${repo.repo}/${repo.branch}/${repo.indexPath}`;
    const response = await fetch(url, { cache: 'no-cache' });
    if (!response.ok) return null;
    const parsed = parsePluginIndex(await response.json());
    if (!parsed) return null;
    const fetchedAt = Date.now();
    writeCache({ repoKey: repoKeyOf(repo), etag: '', fetchedAt, index: parsed });
    return { index: parsed, origin: 'network', fetchedAt };
  } catch {
    return null;
  }
}

/** 计算 SHA-256 小写十六进制。 */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return '';
  const view = new Uint8Array(bytes.byteLength);
  view.set(bytes);
  const digest = await subtle.digest('SHA-256', view.buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}

export interface DownloadResult {
  bytes: Uint8Array;
  verified: boolean;
}

/** 下载插件包并校验 SHA-256。校验失败直接抛错，绝不安装未通过校验的包。 */
export async function downloadPluginPackage(entry: PluginIndexEntry): Promise<DownloadResult> {
  const response = await fetch(entry.download, { redirect: 'follow' });
  if (!response.ok) {
    throw new PluginIndexError(`下载插件包失败（HTTP ${response.status}）`, '插件包地址可能已失效，请刷新列表。');
  }
  const buffer = await response.arrayBuffer();
  if (buffer.byteLength > PLUGIN_MAX_UNCOMPRESSED_BYTES * 2) {
    throw new PluginIndexError('插件包体积异常', '超过允许的上限，已中止安装。');
  }
  const bytes = new Uint8Array(buffer);
  const actual = await sha256Hex(bytes);
  const verified = actual !== '' && actual === entry.sha256;
  if (!verified) {
    throw new PluginIndexError(
      '插件包校验失败',
      '下载内容与索引记录的 SHA-256 不一致，可能被篡改或下载不完整，已中止安装。',
    );
  }
  return { bytes, verified };
}
