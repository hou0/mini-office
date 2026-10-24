// 功能：验证自动标签、草稿复用、已有源码保护及双平台创建竞争；时间：2026-10-24；作者：lq。
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { prepareRelease } = require('./prepare-release.cjs');

// 功能：提供独立的 GitHub 与 Git 模拟，禁止测试修改真实远端；时间：2026-10-24；作者：lq。
function fixture({ existingTag = false, existingRelease = false, race = false, version = '0.1.1', packageVersion = version } = {}) {
  const calls = [];
  let output;
  const release = { tag_name: 'v0.1.1', draft: true, html_url: 'https://example.test/draft' };
  return {
    calls, output: () => output,
    options: {
      api: async (method, path, data) => {
        calls.push([method, path, data]);
        if (path.startsWith('/git/ref/')) return existingTag || race ? { object: { sha: 'original' } } : null;
        if (path === '/git/refs') return { object: { sha: 'current' } };
        if (path.startsWith('/releases?')) return existingRelease ? [release] : [];
        if (path === '/releases') return release;
        throw new Error(`Unexpected API: ${path}`);
      },
      git: args => { calls.push(['git', ...args]); return args[0] === 'rev-parse' ? 'original' : ''; },
      read: path => {
        if (path === 'package.json') return JSON.stringify({ version: packageVersion });
        if (path.includes('tauri.conf.json')) return JSON.stringify({ version });
        return 'Release notes';
      },
      output: value => { output = value; },
    },
  };
}

test('首次运行自动建标签和未公开草稿，输出固定源码与版本', async () => {
  const f = fixture();
  await prepareRelease(f.options);
  assert.equal(f.calls.find(c => c[1] === '/git/refs')[2].ref, 'refs/tags/v0.1.1');
  assert.equal(f.calls.find(c => c[1] === '/releases')[2].draft, true);
  assert.match(f.output(), /SOURCE_SHA=original/);
  assert.match(f.output(), /APP_VERSION=0.1.1\nRELEASE_TAG=v0.1.1/);
});

test('已有草稿与标签复用，不重建、不覆盖、切换到标签源码', async () => {
  const f = fixture({ existingTag: true, existingRelease: true });
  await prepareRelease(f.options);
  assert.equal(f.calls.filter(c => c[0] === 'POST').length, 0);
  assert.ok(f.calls.some(c => c[0] === 'git' && c[1] === 'checkout' && c[3] === 'FETCH_HEAD'));
});

test('标签与源码版本不一致时停止，不创建新版本', async () => {
  const f = fixture();
  await assert.rejects(prepareRelease({ ...f.options, requestedTag: 'v0.1.2' }), /新标签必须/);
  assert.equal(f.calls.filter(c => c[0] === 'POST').length, 0);
});

test('包版本不一致时禁止创建草稿或输出发布信息', async () => {
  const f = fixture({ existingTag: true, packageVersion: '0.1.0' });
  await assert.rejects(prepareRelease(f.options), /版本不一致/);
  assert.equal(f.output(), undefined);
});

test('标签创建竞争时重新读取，且不强推标签', async () => {
  const f = fixture();
  const original = f.options.api;
  let reads = 0;
  f.options.api = async (...args) => {
    if (args[1].startsWith('/git/ref/')) return ++reads === 1 ? null : { object: { sha: 'original' } };
    if (args[1] === '/git/refs') return null;
    return original(...args);
  };
  await prepareRelease(f.options);
  assert.equal(reads, 2);
  assert.ok(!f.calls.some(c => c.includes('--force')));
});

test('API 权限错误直接失败，不被当作标签缺失', async () => {
  const f = fixture();
  f.options.api = async () => { throw new Error('GitHub API 403'); };
  await assert.rejects(prepareRelease(f.options), /403/);
  assert.equal(f.output(), undefined);
});

// 功能：覆盖分页中的草稿和首次发布版本错误，防止遗漏草稿及创建错误标签；时间：2026-10-24；作者：lq。
test('草稿位于第二页时仍复用，不创建重复草稿', async () => {
  const f = fixture({ existingTag: true });
  const original = f.options.api;
  f.options.api = async (...args) => {
    if (args[1] === '/releases?per_page=100&page=1') return Array.from({ length: 100 }, () => ({ tag_name: 'old' }));
    if (args[1] === '/releases?per_page=100&page=2') return [{ tag_name: 'v0.1.1', draft: true }];
    return original(...args);
  };
  await prepareRelease(f.options);
  assert.equal(f.calls.filter(c => c[0] === 'POST').length, 0);
});

test('首次发布版本不一致时也不创建标签', async () => {
  const f = fixture({ packageVersion: '0.1.0' });
  await assert.rejects(prepareRelease(f.options), /版本不一致/);
  assert.equal(f.calls.filter(c => c[0] === 'POST').length, 0);
});
