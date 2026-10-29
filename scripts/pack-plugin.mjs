#!/usr/bin/env node
/**
 * 打包插件为 .moext，并输出可直接粘贴进 index.json 的条目。
 *
 * 用法：
 *   node scripts/pack-plugin.mjs plugin-repo/plugins/theme-office-2003
 *   node scripts/pack-plugin.mjs /path/to/my-theme --base-url https://github.com/OWNER/REPO/releases/download/v1
 *
 * 整仓打包（生成完整 index.json）请用 plugin-repo/scripts/build-index.mjs。
 *
 * 产物：
 *   dist-plugins/<id>-<version>.moext
 *   stdout 打印索引条目 JSON（sha256 已计算）
 */
import { createHash } from 'node:crypto';
import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import { dirname, join, relative, resolve } from 'node:path';
import JSZip from 'jszip';

const PLUGIN_EXT = 'moext';
const MAX_BYTES = 8 * 1024 * 1024;
/** 固定 zip 条目时间戳，保证同一份源码重复打包得到相同字节与相同 SHA-256。 */
const FIXED_DATE = new Date('1980-01-01T00:00:00Z');

function fail(message) {
  console.error(`\n[pack-plugin] ${message}\n`);
  process.exit(1);
}

function parseArgs(argv) {
  const positional = [];
  let baseUrl = '';
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--base-url') {
      baseUrl = argv[index + 1] ?? '';
      index += 1;
      continue;
    }
    positional.push(arg);
  }
  return { dir: positional[0], baseUrl };
}

async function walk(root, current = root, files = []) {
  const entries = await readdir(current, { withFileTypes: true });
  for (const entry of entries) {
    const full = join(current, entry.name);
    if (entry.isDirectory()) {
      await walk(root, full, files);
    } else if (entry.isFile()) {
      files.push(full);
    }
  }
  // 排序保证 zip 条目顺序稳定；readdir 的顺序由文件系统决定，不排序会破坏可重现打包。
  return files.sort();
}

const { dir, baseUrl } = parseArgs(process.argv.slice(2));
if (!dir) fail('请指定插件目录，例如：node scripts/pack-plugin.mjs plugins/theme-office-2003');

const root = resolve(dir);
let manifestRaw;
try {
  manifestRaw = await readFile(join(root, 'manifest.json'), 'utf8');
} catch {
  fail(`${dir} 下没有 manifest.json`);
}

let manifest;
try {
  manifest = JSON.parse(manifestRaw);
} catch (error) {
  fail(`manifest.json 不是合法 JSON：${error.message}`);
}

for (const field of ['id', 'name', 'version', 'publisher', 'tier', 'engines']) {
  if (!manifest[field]) fail(`manifest.json 缺少必填字段 ${field}`);
}

const files = await walk(root);
if (files.length === 0) fail('插件目录为空');

const zip = new JSZip();
let total = 0;
for (const file of files) {
  const bytes = await readFile(file);
  total += bytes.byteLength;
  if (total > MAX_BYTES) fail(`插件总大小超过 ${MAX_BYTES / 1024 / 1024} MB 上限`);
  // createFolders: false 必须显式给：JSZip 默认会给 `templates/` 这类父目录补一个
  // 目录条目，而那个条目用当前时间而不是 FIXED_DATE，于是每次打包哈希都变。
  zip.file(relative(root, file).split('\\').join('/'), bytes, { date: FIXED_DATE, createFolders: false });
}

const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
const sha256 = createHash('sha256').update(buffer).digest('hex');

const outDir = resolve('dist-plugins');
await mkdir(outDir, { recursive: true });
const fileName = `${manifest.id}-${manifest.version}.${PLUGIN_EXT}`;
const outPath = join(outDir, fileName);
await writeFile(outPath, buffer);

const download = baseUrl ? `${baseUrl.replace(/\/$/, '')}/${fileName}` : `<填写 ${fileName} 的下载地址>`;

const entry = {
  id: manifest.id,
  name: manifest.name,
  version: manifest.version,
  publisher: manifest.publisher,
  description: manifest.description ?? '',
  tier: manifest.tier,
  permissions: manifest.permissions ?? [],
  download,
  sha256,
  size: buffer.byteLength,
};

const size = await stat(outPath);
console.log(`\n[pack-plugin] 已生成 ${relative(process.cwd(), outPath)}（${size.size} 字节）`);
console.log(`[pack-plugin] SHA-256 ${sha256}\n`);
console.log('[pack-plugin] 将下面条目加入仓库的 index.json 的 plugins 数组：\n');
console.log(JSON.stringify(entry, null, 2));
console.log('');
