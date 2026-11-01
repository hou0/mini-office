/**
 * 功能：验证浮动下划线图片可被选中并删除（Delete 键 + 右键菜单项）；时间：2026-10-31；作者：lq。
 * 运行：node scripts/verify-word-image-delete.mjs [docx路径]
 */
import { chromium } from 'playwright-core';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'tests/output');
mkdirSync(OUT, { recursive: true });
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const PORT = 4193;
const DOCX = process.argv[2] ?? `${process.env.HOME}/Downloads/测试合同模板666.docx`;

const server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], {
  cwd: path.join(ROOT, 'apps/desktop'),
  stdio: ['ignore', 'pipe', 'pipe'],
});
server.stdout.on('data', () => {});
server.stderr.on('data', (chunk) => process.stderr.write(`[preview:err] ${chunk}`));

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitForServer() {
  for (let attempt = 0; attempt < 40; attempt += 1) {
    try {
      const response = await fetch(`http://127.0.0.1:${PORT}/`);
      if (response.ok) return;
    } catch {
      // 继续等待
    }
    await wait(500);
  }
  throw new Error('预览服务未能在 20 秒内启动');
}

async function main() {
  await waitForServer();
  const browser = await chromium.launch({ executablePath: CHROME, headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.addInitScript(() => {
    // @ts-ignore
    delete window.showOpenFilePicker;
  });
  await page.goto(`http://127.0.0.1:${PORT}/`, { waitUntil: 'networkidle' });

  const chooser = page.waitForEvent('filechooser');
  await page.getByText('打开文件', { exact: false }).first().click();
  await (await chooser).setFiles(DOCX);
  await wait(3500);

  const before = await page.evaluate(() => ({
    frames: document.querySelectorAll('[data-image-frame]').length,
    paragraphs: document.querySelectorAll('[data-paragraph-id]').length,
  }));

  // 点击第一张浮动下划线图
  const clicked = await page.evaluate(() => {
    const img = document.querySelector('[data-image-frame] img');
    if (!img) return false;
    const rect = img.getBoundingClientRect();
    img.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: rect.left + 2, clientY: rect.top + 2 }));
    return true;
  });
  await wait(500);
  const selected = await page.evaluate(() => {
    const frame = document.querySelector('[data-image-frame]');
    const rect = frame?.getBoundingClientRect();
    return { selected: frame?.dataset.selected, width: rect ? Math.round(rect.width) : null, height: rect ? Math.round(rect.height) : null };
  });

  // 按 Delete
  await page.keyboard.press('Delete');
  await wait(900);

  const afterDelete = await page.evaluate(() => ({
    frames: document.querySelectorAll('[data-image-frame]').length,
    paragraphs: document.querySelectorAll('[data-paragraph-id]').length,
  }));

  console.log(JSON.stringify({ clicked, before, selected, afterDelete }, null, 2));
  await page.screenshot({ path: path.join(OUT, 'verify-image-delete.png') });
  await browser.close();
  server.kill();
}

main().catch((error) => {
  console.error(error);
  server.kill();
  process.exit(1);
});
