// 功能：自动补建版本标签和 Release 草稿，已有标签始终使用其原始源码；时间：2026-10-24；作者：lq。
const fs = require('node:fs');
const { execFileSync } = require('node:child_process');

// 功能：复用双平台发布准备逻辑，并允许注入依赖验证竞态及失败分支；时间：2026-10-24；作者：lq。
async function prepareRelease({ api, git, read, output, requestedTag = '' }) {
  const initialVersion = JSON.parse(read('apps/desktop/src-tauri/tauri.conf.json')).version;
  const tag = requestedTag.trim() || `v${initialVersion}`;
  if (!/^v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(tag)) {
    throw new Error('版本标签格式应为 v0.1.1。');
  }
  const refPath = `/git/ref/tags/${encodeURIComponent(tag)}`;
  let ref = await api('GET', refPath, undefined, [404]);
  if (!ref) {
    if (tag !== `v${initialVersion}`) throw new Error('新标签必须与所选源码版本一致。');
    if (JSON.parse(read('package.json')).version !== initialVersion) {
      throw new Error('Tauri 和 package.json 版本不一致，禁止创建标签。');
    }
    const sha = git(['rev-parse', 'HEAD']).trim();
    // 功能：两个平台可能同时创建标签，冲突时重新读取，不覆盖已存在标签；时间：2026-10-24；作者：lq。
    ref = await api('POST', '/git/refs', { ref: `refs/tags/${tag}`, sha }, [422]);
    if (!ref) ref = await api('GET', refPath);
  }
  git(['fetch', '--depth=1', 'origin', `refs/tags/${tag}`]);
  git(['checkout', '--detach', 'FETCH_HEAD']);
  const version = JSON.parse(read('apps/desktop/src-tauri/tauri.conf.json')).version;
  const packageVersion = JSON.parse(read('package.json')).version;
  if (tag !== `v${version}` || packageVersion !== version) {
    throw new Error('标签、Tauri 和 package.json 版本不一致，已停止发布。');
  }
  // 功能：按标签查询接口不返回草稿，因此分页查询发行列表；时间：2026-10-24；作者：lq。
  async function findRelease() {
    for (let page = 1; ; page++) {
      const releases = await api('GET', `/releases?per_page=100&page=${page}`);
      const found = releases.find(item => item.tag_name === tag);
      if (found || releases.length < 100) return found;
    }
  }
  let release = await findRelease();
  if (!release) {
    const notesPath = `docs/releases/${tag}.md`;
    let body = `Mini Office ${version}`;
    try { body = read(notesPath); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    release = await api('POST', '/releases', {
      tag_name: tag, target_commitish: git(['rev-parse', 'HEAD']).trim(),
      name: `Mini Office ${tag}`, body, draft: true, prerelease: false,
    }, [422]);
    // 功能：处理双平台同时建草稿的竞争，其他错误直接终止；时间：2026-10-24；作者：lq。
    if (!release) release = await findRelease();
    if (!release) throw new Error('无法创建或找到 Release，请检查仓库发布权限。');
  }
  output(`APP_VERSION=${version}\nRELEASE_TAG=${tag}\nSOURCE_SHA=${git(['rev-parse', 'HEAD']).trim()}\n`);
  console.log(`发布准备完成：${tag}，${release.draft ? '草稿' : '已有公开版本'}，${release.html_url}`);
}

// 功能：使用工作流令牌调用 GitHub API，仅允许明确的可恢复状态码；时间：2026-10-24；作者：lq。
async function main() {
  const { GITHUB_TOKEN, GITHUB_REPOSITORY, GITHUB_ENV, GITHUB_OUTPUT, GITHUB_API_URL, RELEASE_TAG_INPUT } = process.env;
  if (!GITHUB_TOKEN || !GITHUB_REPOSITORY || !GITHUB_ENV) throw new Error('缺少 GitHub 工作流环境变量。');
  await prepareRelease({
    api: async (method, path, data, allowed = []) => {
      const response = await fetch(`${GITHUB_API_URL || 'https://api.github.com'}/repos/${GITHUB_REPOSITORY}${path}`, {
        method, headers: { Authorization: `Bearer ${GITHUB_TOKEN}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json' },
        body: data === undefined ? undefined : JSON.stringify(data),
      });
      if (allowed.includes(response.status)) return null;
      if (!response.ok) throw new Error(`GitHub API ${method} ${path} 失败（${response.status}）：${await response.text()}`);
      return response.json();
    },
    git: args => execFileSync('git', args, { encoding: 'utf8' }),
    read: path => fs.readFileSync(path, 'utf8'),
    output: value => {
      fs.appendFileSync(GITHUB_ENV, value);
      if (GITHUB_OUTPUT) fs.appendFileSync(GITHUB_OUTPUT, value);
    },
    requestedTag: RELEASE_TAG_INPUT,
  });
}

if (require.main === module) main().catch(error => { console.error(error.message); process.exitCode = 1; });
module.exports = { prepareRelease };
