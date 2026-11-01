/**
 * 功能：核对 Word 表格边框线宽、右键菜单、表格手柄与深色纸张；时间：2026-10-31；作者：lq。
 * 运行：node scripts/verify-word-table-ui.mjs
 */
import { chromium } from 'playwright-core';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'output/playwright');
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const URL = 'http://127.0.0.1:4190/';
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const browser = await chromium.launch({ executablePath: CHROME, headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  await page.addInitScript(() => { delete window.showOpenFilePicker; });
  await page.goto(URL, { waitUntil: 'domcontentloaded' });
  const chooser = page.waitForEvent('filechooser');
  await page.getByText('打开文件', { exact: false }).first().click();
  await (await chooser).setFiles(path.join(process.env.HOME, 'Downloads/测试合同模板666.docx'));
  await page.waitForSelector('[data-table-id]', { timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);
  await wait(1200);

  const table = page.locator('[data-table-id]').first();
  await table.scrollIntoViewIfNeeded();
  await wait(400);
  await table.screenshot({ path: path.join(OUT, 'table-01-borders.png') });

  // 线宽统计：同一条竖线不应出现 2px 与 1px 混杂
  const widths = await page.evaluate(() => {
    const grid = document.querySelector('[data-table-id]');
    const cells = Array.from(grid.querySelectorAll('[data-cell-id]'));
    const tally = (side) => {
      const counts = {};
      for (const cell of cells) {
        const value = getComputedStyle(cell)[side];
        counts[value] = (counts[value] ?? 0) + 1;
      }
      return counts;
    };
    return { left: tally('borderLeftWidth'), right: tally('borderRightWidth'), top: tally('borderTopWidth'), bottom: tally('borderBottomWidth') };
  });
  console.log('线宽分布', JSON.stringify(widths));

  // 右键菜单
  const cell = table.locator('[data-cell-id]').nth(2);
  await cell.click({ button: 'right' });
  await page.waitForSelector('[role="menu"]', { timeout: 5000 });
  const menu = await page.evaluate(() => {
    const list = document.querySelector('[role="menu"]');
    return Array.from(list.querySelectorAll('button, [role="menuitem"]')).map((item) => item.textContent?.trim()).filter(Boolean);
  });
  console.log('右键菜单', JSON.stringify(menu));
  await page.screenshot({ path: path.join(OUT, 'table-02-context-menu.png') });
  await page.keyboard.press('Escape');
  await wait(300);

  // 表格手柄：单击选中整表
  await table.hover();
  await wait(200);
  const handle = page.locator('[data-table-handle]').first();
  console.log('手柄可见', await handle.isVisible());
  await handle.click();
  await wait(500);
  const selected = await page.evaluate(() => {
    const grid = document.querySelector('[data-table-id]');
    const cells = Array.from(grid.querySelectorAll('[data-cell-id]'));
    const tinted = cells.filter((cell) => getComputedStyle(cell).boxShadow !== 'none').length;
    const blueBorder = cells.some((cell) => getComputedStyle(cell).boxShadow.includes('52, 120, 246'));
    return { total: cells.length, tinted, blueBorder };
  });
  console.log('整表选中', JSON.stringify(selected));
  await table.screenshot({ path: path.join(OUT, 'table-03-selected.png') });

  // 深色纸张
  await page.evaluate(() => { document.documentElement.dataset.theme = 'dark'; });
  await wait(600);
  const dark = await page.evaluate(() => {
    const paper = document.querySelector('.mo-word-page');
    const run = document.querySelector('.mo-word-page [data-ink="1"]');
    const root = getComputedStyle(document.documentElement);
    const scroll = document.querySelector('.mo-word-scroll');
    return {
      paper: getComputedStyle(paper).backgroundColor,
      canvas: scroll ? getComputedStyle(scroll).backgroundColor : null,
      appBg: root.getPropertyValue('--mo-bg').trim(),
      ink: run ? getComputedStyle(run).color : null,
      border: getComputedStyle(document.querySelector('[data-cell-id]')).borderLeftColor,
    };
  });
  console.log('深色纸张', JSON.stringify(dark));
  if (dark.canvas && dark.paper !== dark.canvas) console.log(`注意：纸张 ${dark.paper} 与画布 ${dark.canvas} 不一致`);
  await page.screenshot({ path: path.join(OUT, 'table-04-dark-paper.png') });
  console.log('页面错误', JSON.stringify(errors));
} finally {
  await browser.close();
}
