/**
 * 功能：核对 Word 深色模式下没有任何「浅色块」——文字底纹、段落底纹、表格底纹都要跟随纸张压暗；
 * 同时确认浅色模式下底纹仍然存在（避免把底纹整体删掉冒充修复）；时间：2026-10-31；作者：lq。
 * 运行：node scripts/verify-word-dark-shading.mjs
 */
import { chromium } from 'playwright-core';
import path from 'node:path';
import fs from 'node:fs';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'output/playwright');
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const URL = 'http://127.0.0.1:4190/';
const FILE = process.env.WORD_FILE ?? path.join(process.env.HOME, 'Downloads/测试合同模板-修复验证.docx');
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

fs.mkdirSync(OUT, { recursive: true });

const browser = await chromium.launch({ executablePath: CHROME, headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  await page.addInitScript(() => { delete window.showOpenFilePicker; });
  await page.goto(URL, { waitUntil: 'domcontentloaded' });
  const chooser = page.waitForEvent('filechooser');
  await page.getByText('打开文件', { exact: false }).first().click();
  await (await chooser).setFiles(FILE);
  await page.waitForSelector('.mo-word-page', { timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);
  await wait(1500);

  const audit = () =>
    page.evaluate(() => {
      const parse = (value) => {
        const match = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(value);
        if (!match) return null;
        return { r: +match[1], g: +match[2], b: +match[3], a: match[4] === undefined ? 1 : +match[4] };
      };
      const blocks = [];
      let shaded = 0;
      for (const element of document.querySelectorAll('.mo-word-page, .mo-word-page *')) {
        const rgb = parse(getComputedStyle(element).backgroundColor);
        if (!rgb || rgb.a < 0.1) continue;
        shaded += 1;
        // 浅色块：三个通道都很亮（纸张本身在深色模式下远低于此阈值）
        if (rgb.r >= 190 && rgb.g >= 190 && rgb.b >= 190) {
          blocks.push({ tag: element.tagName, text: (element.textContent ?? '').trim().slice(0, 18), color: `${rgb.r},${rgb.g},${rgb.b}` });
        }
      }
      return { shaded, blocks };
    });

  const light = await audit();
  console.log('浅色模式底纹块数', light.shaded);
  await page.screenshot({ path: path.join(OUT, 'word-dark-01-light.png') });

  await page.evaluate(() => { document.documentElement.dataset.theme = 'dark'; });
  await wait(900);
  const dark = await audit();
  console.log('深色模式底纹块数', dark.shaded, '浅色块', dark.blocks.length);
  if (dark.blocks.length) console.log('浅色块明细', JSON.stringify(dark.blocks.slice(0, 12), null, 2));
  await page.screenshot({ path: path.join(OUT, 'word-dark-02-dark.png') });

  // 签字页（第二页）单独截图，便于肉眼核对
  const signature = page.locator('.mo-word-page').nth(1);
  await signature.scrollIntoViewIfNeeded();
  await wait(400);
  await signature.screenshot({ path: path.join(OUT, 'word-dark-03-signature.png') });

  const problems = [];
  if (dark.blocks.length) problems.push(`深色模式仍有 ${dark.blocks.length} 个浅色块`);
  if (light.shaded === 0) problems.push('浅色模式下没有任何底纹，说明底纹被整体删除而非换算');
  if (dark.shaded !== light.shaded) problems.push(`底纹元素数量变化：浅色 ${light.shaded} / 深色 ${dark.shaded}`);
  if (errors.length) problems.push(`页面报错：${errors.join(' | ')}`);
  console.log(problems.length ? `问题：\n- ${problems.join('\n- ')}` : '全部检查通过');
  process.exitCode = problems.length ? 1 : 0;
} finally {
  await browser.close();
}
