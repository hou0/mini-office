import { chromium } from 'playwright-core';
import path from 'node:path';
import fs from 'node:fs';

const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const URL = 'http://127.0.0.1:4191/';
const OUT = path.join(process.cwd(), 'output/playwright');
fs.mkdirSync(OUT, { recursive: true });

const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const errors = [];
page.on('pageerror', (error) => errors.push(String(error)));
await page.addInitScript(() => { delete window.showOpenFilePicker; });
await page.goto(URL, { waitUntil: 'domcontentloaded' });
await new Promise((r) => setTimeout(r, 2500));

async function openFile(file) {
  const chooser = page.waitForEvent('filechooser');
  await page.getByText('打开文件', { exact: false }).first().click();
  await (await chooser).setFiles(file);
  await page.waitForSelector('.mo-word-page, .mo-sheet', { timeout: 90000 });
  await new Promise((r) => setTimeout(r, 1500));
}

// 1) Word 表格单元格：点击后不应出现灰底
await openFile(path.join(process.env.HOME, 'Downloads/测试合同模板-修复验证.docx'));
await page.waitForSelector('[data-cell-id]', { timeout: 60000 });
const cell = page.locator('[data-cell-id]').first();
await cell.click();
await new Promise((r) => setTimeout(r, 500));
const cellShadow = await page.evaluate(() => {
  const cells = Array.from(document.querySelectorAll('[data-cell-id]'));
  return cells.map((el) => getComputedStyle(el).boxShadow).filter((v) => v && v !== 'none');
});
console.log('单元格灰底数量:', cellShadow.length);

// 2) Word 打印预览布局
await page.keyboard.press('Meta+p');
await page.waitForSelector('[aria-label="文档打印预览"]', { timeout: 30000 });
await new Promise((r) => setTimeout(r, 1200));
const wordPreview = await page.evaluate(() => {
  const host = document.querySelector('[aria-label="文档打印预览"]');
  const panel = host?.closest('div[class*="grid"]');
  return {
    hostBackground: host ? getComputedStyle(host).backgroundColor : null,
    panelBackground: panel ? getComputedStyle(panel).backgroundColor : null,
    frames: host ? host.querySelectorAll(':scope > div > div').length : 0,
    pager: document.body.innerText.match(/\d+ \/ \d+/g)?.slice(0, 3) ?? [],
    cellShadow: Array.from(host?.querySelectorAll('[data-cell-id]') ?? []).map((el) => getComputedStyle(el).boxShadow).filter((v) => v && v !== 'none').length,
  };
});
console.log('Word 预览:', JSON.stringify(wordPreview));
await page.screenshot({ path: path.join(OUT, 'preview-word.png') });
await page.getByRole('button', { name: '取消' }).click();
await new Promise((r) => setTimeout(r, 600));

// 3) Excel 打印预览布局
const xlsx = path.join(process.env.HOME, 'Downloads/444444444.xlsx');
if (fs.existsSync(xlsx)) {
  await openFile(xlsx);
  await page.keyboard.press('Meta+p');
  await page.waitForSelector('[aria-label="工作表打印预览"]', { timeout: 30000 });
  await new Promise((r) => setTimeout(r, 1200));
  const sheetPreview = await page.evaluate(() => {
    const host = document.querySelector('[aria-label="工作表打印预览"]');
    const panel = host?.closest('div[class*="grid"]');
    return {
      hostBackground: host ? getComputedStyle(host).backgroundColor : null,
      panelBackground: panel ? getComputedStyle(panel).backgroundColor : null,
      pager: document.body.innerText.match(/\d+ \/ \d+/g)?.slice(0, 3) ?? [],
    };
  });
  console.log('Excel 预览:', JSON.stringify(sheetPreview));
  await page.screenshot({ path: path.join(OUT, 'preview-sheet.png') });
}

if (errors.length) console.log('页面报错:', errors.slice(0, 3).join(' | '));
await browser.close();
