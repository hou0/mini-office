#!/usr/bin/env node
/**
 * 生成插件索引：扫描 plugins/ 下的所有插件，打包成 .moext，
 * 计算 SHA-256 与体积，并写出 index.json。
 *
 * 本地：
 *   node scripts/build-index.mjs --owner lq0910 --repo mini-office-plugins
 *
 * CI（GitHub Actions 自动注入环境变量）：
 *   node scripts/build-index.mjs
 *
 * 产物：
 *   packages/*.moext   （提交进仓库，供 raw.githubusercontent.com 分发）
 *   index.json
 *
 * 为什么不用 GitHub Release 资产：
 * `github.com/.../releases/download/...` 会 302 到 objects.githubusercontent.com，
 * 该域不带 `Access-Control-Allow-Origin`，浏览器与 Tauri WebView 的 fetch 都会被
 * CORS 拦截。`raw.githubusercontent.com` 带 `*`，是唯一能在客户端直接下载的通道。
 * 插件包只有 1KB 量级，提交进仓库的代价可以忽略。
 */
import { createHash } from 'node:crypto';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { join, relative, resolve } from 'node:path';
import JSZip from 'jszip';

const PLUGIN_EXT = 'moext';
const PACKAGES_DIR = 'packages';
const MAX_BYTES = 8 * 1024 * 1024;
/**
 * 固定 zip 条目时间戳，保证同一份源码每次打包得到**完全相同的字节**。
 * 否则每次重新打包 SHA-256 都会变，索引里的校验值与仓库里的包对不上，
 * 用户安装时会报"插件包校验失败"。
 */
const FIXED_DATE = new Date('1980-01-01T00:00:00Z');

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith('--')) continue;
    args[key.slice(2)] = argv[index + 1];
    index += 1;
  }
  return args;
}

function fail(message) {
  console.error(`\n[build-index] ${message}\n`);
  process.exit(1);
}

async function walk(root, current = root, files = []) {
  const entries = await readdir(current, { withFileTypes: true });
  for (const entry of entries) {
    const full = join(current, entry.name);
    if (entry.isDirectory()) await walk(root, full, files);
    else if (entry.isFile()) files.push(full);
  }
  // 必须排序：readdir 的返回顺序由文件系统决定，同一份源码在不同机器（或同一机器
  // 不同时刻）可能给出不同顺序，zip 条目顺序一变字节就变，SHA-256 跟着变。
  // 单文件插件看不出来，多文件插件（带模板/资源）会直接破坏可重现打包。
  return files.sort();
}

const args = parseArgs(process.argv.slice(2));
// 支持三种写法：`--repo owner/name`、`--owner x --repo name`、CI 注入的 GITHUB_REPOSITORY。
const repoArg = args.repo ?? process.env.GITHUB_REPOSITORY ?? 'lq0910/mini-office-plugins';
const ownerArg = args.owner ?? process.env.GITHUB_REPOSITORY_OWNER ?? 'lq0910';
let repoOwner;
let repoName;
if (repoArg.includes('/')) {
  [repoOwner, repoName] = repoArg.split('/');
} else {
  repoOwner = ownerArg;
  repoName = repoArg;
}
if (!repoOwner || !repoName) fail(`无法确定仓库坐标：${repoArg}`);
// raw 通道必须指向**包含 packages/ 的分支**。不要用 tag：CI 是在 tag 检出后才打包的，
// tag 里没有新产物，指向 tag 会 404。索引与包在同一提交里，天然对齐。
const ref = args.ref ?? 'main';
// 默认走 raw 通道；`--mode release` 留给自建镜像或需要 Release 页面的场景。
const mode = args.mode ?? 'raw';
const baseUrl =
  mode === 'release'
    ? (args['base-url'] ?? `https://github.com/${repoOwner}/${repoName}/releases/download/${args.tag ?? 'v1'}`).replace(/\/$/, '')
    : `https://raw.githubusercontent.com/${repoOwner}/${repoName}/${ref}/${PACKAGES_DIR}`;

const root = resolve('plugins');
let pluginDirs;
try {
  pluginDirs = (await readdir(root, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .sort((a, b) => a.name.localeCompare(b.name));
} catch {
  fail('找不到 plugins/ 目录');
}
if (pluginDirs.length === 0) fail('plugins/ 下没有任何插件');

const outDir = resolve(PACKAGES_DIR);
await mkdir(outDir, { recursive: true });

const plugins = [];
for (const dir of pluginDirs) {
  const pluginRoot = join(root, dir.name);
  let manifest;
  try {
    manifest = JSON.parse(await readFile(join(pluginRoot, 'manifest.json'), 'utf8'));
  } catch (error) {
    fail(`${dir.name}: 无法读取 manifest.json（${error.message}）`);
  }

  for (const field of ['id', 'name', 'version', 'publisher', 'tier', 'engines']) {
    if (!manifest[field]) fail(`${dir.name}: manifest.json 缺少 ${field}`);
  }

  const files = await walk(pluginRoot);
  const zip = new JSZip();
  let total = 0;
  for (const file of files) {
    const bytes = await readFile(file);
    total += bytes.byteLength;
    if (total > MAX_BYTES) fail(`${dir.name}: 解压后超过 ${MAX_BYTES / 1024 / 1024} MB`);
    // createFolders: false 是关键——JSZip 默认会为 `templates/` 这类父目录补一个
    // 目录条目，而那个条目用的是**当前时间**，不是 FIXED_DATE。目录条目时间一变，
    // 整包字节就变，SHA-256 跟着变，可重现打包当场失效（单文件插件不会触发，
    // 带子目录的插件才会，所以很容易漏掉）。
    zip.file(relative(pluginRoot, file).split('\\').join('/'), bytes, {
      date: FIXED_DATE,
      createFolders: false,
    });
  }

  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
  const sha256 = createHash('sha256').update(buffer).digest('hex');
  const fileName = `${manifest.id}-${manifest.version}.${PLUGIN_EXT}`;
  await writeFile(join(outDir, fileName), buffer);

  const entry = {
    id: manifest.id,
    name: manifest.name,
    version: manifest.version,
    publisher: manifest.publisher,
    tier: manifest.tier,
    permissions: manifest.permissions ?? [],
    download: `${baseUrl}/${fileName}`,
    sha256,
    size: buffer.byteLength,
  };
  if (manifest.description) entry.description = manifest.description;
  if (manifest.license) entry.license = manifest.license;
  if (manifest.contributes?.themes?.length) {
    entry.tags = ['theme', ...manifest.contributes.themes.map((theme) => theme.id)];
  }
  plugins.push(entry);
  console.log(`[build-index] ${fileName}  ${sha256.slice(0, 12)}…  ${buffer.byteLength} B`);
}

plugins.sort((a, b) => a.id.localeCompare(b.id));
const index = {
  schemaVersion: 1,
  updatedAt: new Date().toISOString(),
  plugins,
};
await writeFile(resolve('index.json'), `${JSON.stringify(index, null, 2)}\n`);

console.log(`\n[build-index] 已写出 index.json，共 ${plugins.length} 个插件`);
console.log(`[build-index] 下载地址前缀：${baseUrl}`);
