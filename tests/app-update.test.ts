import { beforeEach, describe, expect, it, vi } from 'vitest';

// 功能：把被测版本固定成预览版，用例才能覆盖「预览版用户可见 beta」的分支；时间：2026-10-30；作者：lq。
// 说明：正式版发版后根包 version 不含 "-"，ALLOW_PRERELEASE 会变为 false，
// 这些用例原本会因预览版 tag 被过滤而全部拿到 null，与实际逻辑无关。
vi.mock('@mini-office/shared', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../packages/shared/src/index')>();
  return { ...actual, APP_VERSION: '0.1.5-preview.9' };
});

const { checkForUpdates, fetchLatestRelease, useUpdateStore } = await import(
  '../packages/editor-shell/src/update-check'
);

const LATEST_KEY = 'mini-office.update.latest.v1';
const DISMISSED_KEY = 'mini-office.update.dismissed.v1';

function release(tag: string, extra: Record<string, unknown> = {}) {
  return {
    tag_name: tag,
    name: tag,
    html_url: `https://example.com/${tag}`,
    prerelease: tag.includes('-'),
    draft: false,
    published_at: '2026-10-30T00:00:00Z',
    ...extra,
  };
}

function mockReleases(list: unknown[]) {
  const fetchMock = vi.fn(async () => ({ ok: true, json: async () => list }));
  vi.stubGlobal('fetch', fetchMock);
  return fetchMock;
}

function resetStore() {
  useUpdateStore.setState({
    status: 'idle',
    latest: null,
    checkedAt: null,
    error: null,
    dismissedTag: null,
  });
}

describe('应用版本更新检查', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.unstubAllGlobals();
    resetStore();
  });

  it('远端没有更高版本时视为无更新', async () => {
    mockReleases([release('v0.1.4')]);
    expect(await fetchLatestRelease({ force: true })).toBeNull();
  });

  it('取版本号最高的条目，并跳过草稿', async () => {
    mockReleases([
      release('v0.1.6-preview.1'),
      release('v0.1.7-preview.1'),
      release('v9.9.9', { draft: true }),
      release('v0.1.5-preview.2'),
    ]);
    const latest = await fetchLatestRelease({ force: true });
    expect(latest?.tagName).toBe('v0.1.7-preview.1');
    expect(latest?.prerelease).toBe(true);
  });

  it('缺少必要字段的条目被忽略', async () => {
    mockReleases([{ tag_name: 'v0.1.9', draft: false }, release('v0.1.6-preview.1')]);
    const latest = await fetchLatestRelease({ force: true });
    expect(latest?.tagName).toBe('v0.1.6-preview.1');
  });

  it('查到新版本后写入缓存，节流期内不再发请求', async () => {
    const fetchMock = mockReleases([release('v0.1.7-preview.1')]);
    await checkForUpdates({ force: true });

    expect(useUpdateStore.getState().latest?.tagName).toBe('v0.1.7-preview.1');
    expect(useUpdateStore.getState().status).toBe('ready');
    expect(localStorage.getItem(LATEST_KEY)).toContain('v0.1.7-preview.1');

    const calls = fetchMock.mock.calls.length;
    await checkForUpdates();
    expect(fetchMock.mock.calls.length).toBe(calls);
  });

  it('接口失败时保留上一次结果，只记录错误', async () => {
    useUpdateStore.setState({
      latest: { tagName: 'v0.1.7-preview.1', name: 'x', url: 'https://example.com', prerelease: true },
      status: 'ready',
    });
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => ({ ok: false, status: 403 })),
    );

    await checkForUpdates({ force: true });

    const state = useUpdateStore.getState();
    expect(state.error).toContain('403');
    expect(state.status).toBe('ready');
    expect(state.latest?.tagName).toBe('v0.1.7-preview.1');
  });

  it('稍后提醒写入版本号，用于抑制重复弹窗', async () => {
    mockReleases([release('v0.1.7-preview.1')]);
    await checkForUpdates({ force: true });
    useUpdateStore.getState().dismiss();

    expect(localStorage.getItem(DISMISSED_KEY)).toBe('v0.1.7-preview.1');
    expect(useUpdateStore.getState().dismissedTag).toBe('v0.1.7-preview.1');
    // 小红点仍然依赖 latest，不因忽略而消失。
    expect(useUpdateStore.getState().latest).not.toBeNull();
  });
});
