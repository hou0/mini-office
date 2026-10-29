/**
 * 功能：用真实浏览器核对 Word 字号/行距/段落/格式刷的界面与交互；时间：2026-10-29；作者：lq。
 * 运行：node scripts/verify-word-typography.mjs
 */
import { chromium } from 'playwright-core';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'tests/output');
mkdirSync(OUT, { recursive: true });
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const PORT = 4187;

const server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], {
  cwd: path.join(ROOT, 'apps/desktop'),
  stdio: ['ignore', 'pipe', 'pipe'],
});
server.stdout.on('data', (chunk) => process.stdout.write(`[preview] ${chunk}`));
server.stderr.on('data', (chunk) => process.stdout.write(`[preview:err] ${chunk}`));

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

/** 功能：在页面内构造真实 DOM 选区（键盘模拟无法在 contentEditable 上建立选区）；时间：2026-10-29；作者：lq。 */
async function selectInParagraph(page, index, length = 6) {
  const text = await page.evaluate(
    ({ index, length }) => {
      const element = document.querySelectorAll('[data-paragraph-id]')[index];
      if (!element) return '';
      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
      let node = walker.nextNode();
      while (node && !node.textContent?.trim()) node = walker.nextNode();
      if (!node || !node.textContent) return '';
      const range = document.createRange();
      range.setStart(node, 0);
      range.setEnd(node, Math.min(length, node.textContent.length));
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      return selection.toString();
    },
    { index, length },
  );
  await wait(250);
  return text;
}

async function main() {
  await waitForServer();
  const browser = await chromium.launch({ executablePath: CHROME, headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });

  // 强制走 input[type=file] 回退，便于自动化选择本地文档。
  await page.addInitScript(() => {
    // @ts-ignore
    delete window.showOpenFilePicker;
  });
  await page.goto(`http://127.0.0.1:${PORT}/`, { waitUntil: 'networkidle' });

  const chooser = page.waitForEvent('filechooser');
  await page.getByText('打开文件', { exact: false }).first().click();
  const fileChooser = await chooser;
  await fileChooser.setFiles(path.join(ROOT, 'tests/fixtures/docx/contract.docx'));
  await page.waitForSelector('[data-word-editor]', { timeout: 20000 });
  await wait(1200);

  // 选中正文一段文字，验证字号与行距控件能作用到选区。
  const paragraph = page.locator('[data-paragraph-id]').nth(2);
  console.log('选中文字：', await selectInParagraph(page, 2, 6));
  await page.screenshot({ path: path.join(OUT, 'typography-01-toolbar.png'), clip: { x: 0, y: 0, width: 1440, height: 300 } });

  const fontSizeTrigger = page.locator('[aria-label="字号"]');
  console.log('字号显示：', (await fontSizeTrigger.textContent())?.trim());
  await fontSizeTrigger.click();
  await page.waitForSelector('[data-testid="font-size-panel"]');
  await page.screenshot({ path: path.join(OUT, 'typography-02-fontsize.png'), clip: { x: 0, y: 0, width: 720, height: 620 } });
  await page.locator('[data-testid="font-size-panel"] button', { hasText: '三号' }).first().click();
  await wait(500);
  console.log('设置三号后显示：', (await fontSizeTrigger.textContent())?.trim());
  console.log('段落含 16pt：', (await paragraph.innerHTML()).includes('16pt'));

  await page.locator('[aria-label="行距"]').click();
  await wait(400);
  await page.screenshot({ path: path.join(OUT, 'typography-03-linespacing.png'), clip: { x: 300, y: 0, width: 1000, height: 560 } });
  const exactInput = page.locator('[aria-label="固定值行距"]');
  await exactInput.fill('26');
  await page.locator('button', { hasText: /^应用$/ }).first().click();
  await wait(600);
  console.log('固定行距后显示：', (await page.locator('[aria-label="行距"]').textContent())?.trim());

  await page.locator('[aria-label="段落"]').click();
  await wait(400);
  await page.screenshot({ path: path.join(OUT, 'typography-04-paragraph-menu.png'), clip: { x: 300, y: 0, width: 1000, height: 460 } });
  await page.locator('button', { hasText: '首行缩进 2 字符' }).first().click();
  await wait(600);
  const indented = await page.evaluate(() => {
    const element = document.querySelectorAll('[data-paragraph-id]')[2];
    return element ? getComputedStyle(element).textIndent : '';
  });
  console.log('首行缩进后 text-indent：', indented);

  await page.locator('[aria-label="段落"]').click();
  await page.locator('button', { hasText: '段落设置' }).first().click();
  await page.waitForSelector('[role="dialog"]');
  await wait(400);
  await page.screenshot({ path: path.join(OUT, 'typography-05-paragraph-dialog.png') });
  await page.locator('[role="dialog"] button', { hasText: '取消' }).click();

  // 格式刷：先复制标题格式，再套用到正文。
  await page.locator('[data-paragraph-id]').first().click();
  await page.locator('[aria-label="格式刷"]').click();
  await wait(300);
  const target = page.locator('[data-paragraph-id]').nth(3);
  const beforePaint = await target.innerHTML();
  const box = await target.boundingBox();
  await page.mouse.move(box.x + 10, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width - 20, box.y + box.height / 2, { steps: 8 });
  await page.mouse.up();
  await wait(800);
  const afterPaint = await target.innerHTML();
  console.log('格式刷生效：', beforePaint !== afterPaint);
  console.log('格式刷按钮状态：', await page.evaluate(() => {
    const button = document.querySelector('[aria-label="格式刷"]');
    return { active: button?.getAttribute('data-active'), bg: getComputedStyle(button).backgroundColor };
  }));
  console.log('  前：', beforePaint.slice(0, 120));
  console.log('  后：', afterPaint.slice(0, 120));
  await page.screenshot({ path: path.join(OUT, 'typography-06-format-painter.png'), clip: { x: 0, y: 0, width: 1440, height: 700 } });

  console.log('页面错误：', errors.length === 0 ? '无' : errors.slice(0, 5));
  await browser.close();
  server.kill();
}

main()
  .catch((error) => {
    console.error(error);
    server.kill();
    process.exit(1);
  })
  .finally(() => setTimeout(() => process.exit(0), 500));
