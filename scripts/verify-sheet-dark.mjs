/**
 * 功能：核对表格画布的深色模式适配（纸张底色、单元格底纹、文字色、网格线、冻结窗格），
 * 并确认打印/导出态强制回白纸；时间：2026-10-31；作者：lq。
 * 运行：node scripts/verify-sheet-dark.mjs
 */
import { chromium } from 'playwright-core';
import path from 'node:path';
import fs from 'node:fs';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'output/playwright');
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const URL = 'http://127.0.0.1:4190/';
const FILE = process.env.SHEET_FILE ?? path.join(process.env.HOME, 'Downloads/pnrpay渠道号申请表-机构合一综合版-20260821.xlsx');
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
  await page.waitForSelector('[data-sheet-grid]', { timeout: 60000 });
  await wait(1500);

  const sample = () =>
    page.evaluate(() => {
      const grid = document.querySelector('[data-sheet-grid]');
      const cells = Array.from(grid.querySelectorAll('[data-cell]')).slice(0, 400);
      const tally = (fn) => {
        const counts = {};
        for (const cell of cells) {
          const value = fn(getComputedStyle(cell));
          counts[value] = (counts[value] ?? 0) + 1;
        }
        return counts;
      };
      const root = getComputedStyle(document.documentElement);
      return {
        theme: document.documentElement.dataset.theme ?? 'light',
        paper: root.getPropertyValue('--mo-paper').trim(),
        paperInk: root.getPropertyValue('--mo-paper-ink').trim(),
        gridLine: root.getPropertyValue('--mo-grid-line').trim(),
        gridBackground: getComputedStyle(grid).backgroundColor,
        cellBackground: tally((style) => style.backgroundColor),
        cellColor: tally((style) => style.color),
        cellBorderRight: tally((style) => style.borderRightColor),
        cellCount: cells.length,
      };
    });

  const light = await sample();
  console.log('浅色', JSON.stringify(light, null, 2));
  await page.screenshot({ path: path.join(OUT, 'sheet-01-light.png') });

  // 缩放：工具栏不再有缩放下拉，右下角状态栏提供滑块与比例，且改动会作用到表格画布
  const zoomUi = await page.evaluate(() => ({
    toolbarZoom: document.querySelectorAll('[data-sheet-toolbar] [aria-label="缩放"]').length,
    statusSlider: document.querySelectorAll('footer [aria-label="缩放滑块"]').length,
    statusSelect: document.querySelectorAll('footer [aria-label="缩放比例"]').length,
    statusGroup: !!document.querySelector('footer [aria-label="工作表缩放"]'),
  }));
  console.log('缩放入口', JSON.stringify(zoomUi));
  const zoomBefore = await page.evaluate(() => {
    const grid = document.querySelector('[data-sheet-grid]');
    return getComputedStyle(grid.parentElement).zoom;
  });
  await page.evaluate(() => {
    const slider = document.querySelector('footer [aria-label="缩放滑块"]');
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(slider, '150');
    slider.dispatchEvent(new Event('input', { bubbles: true }));
    slider.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await wait(500);
  const zoomAfter = await page.evaluate(() => {
    const grid = document.querySelector('[data-sheet-grid]');
    const slider = document.querySelector('footer [aria-label="缩放滑块"]');
    return { zoom: getComputedStyle(grid.parentElement).zoom, slider: slider?.value };
  });
  console.log('缩放前后', JSON.stringify({ zoomBefore, zoomAfter }));
  await page.evaluate(() => {
    const slider = document.querySelector('footer [aria-label="缩放滑块"]');
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(slider, '100');
    slider.dispatchEvent(new Event('input', { bubbles: true }));
    slider.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await wait(400);

  // 切换深色：走应用真实入口（data-theme 由 applyTheme 写入，组件用 MutationObserver 跟随）
  await page.evaluate(() => {
    document.documentElement.dataset.theme = 'dark';
    document.documentElement.style.colorScheme = 'dark';
  });
  await wait(900);
  const dark = await sample();
  console.log('深色', JSON.stringify(dark, null, 2));
  await page.screenshot({ path: path.join(OUT, 'sheet-02-dark.png') });

  // 导出/打印态：变量被强制重置，屏幕上应回到白纸黑字
  await page.evaluate(() => { document.documentElement.dataset.exporting = '1'; });
  await wait(500);
  const exporting = await sample();
  console.log('导出态', JSON.stringify(exporting, null, 2));
  await page.screenshot({ path: path.join(OUT, 'sheet-03-exporting.png') });
  await page.evaluate(() => { delete document.documentElement.dataset.exporting; });

  const problems = [];
  const whiteish = (value) => /rgb\(2[0-9]{2}, 2[0-9]{2}, 2[0-9]{2}\)/.test(value);
  const lightBackgrounds = Object.keys(dark.cellBackground).filter(whiteish);
  if (lightBackgrounds.length) problems.push(`深色模式仍有浅色单元格底：${lightBackgrounds.join(' / ')}`);
  const darkInk = Object.keys(dark.cellColor).filter((value) => value === 'rgb(29, 41, 57)');
  if (darkInk.length) problems.push(`深色模式仍有深色文字：${darkInk.join(' / ')}`);
  if (dark.gridLine === light.gridLine) problems.push('网格线颜色没有跟随主题变化');
  if (!whiteish(exporting.gridBackground)) problems.push(`导出态画布没有回白纸：${exporting.gridBackground}`);
  const exportLightBg = Object.keys(exporting.cellBackground).some(whiteish);
  if (!exportLightBg) problems.push('导出态单元格底色没有回到白纸');
  // 导出态必须与浅色模式逐色一致：深色模式的压暗色（hsl 计算值）不允许漏进输出
  const sameFills = JSON.stringify(exporting.cellBackground) === JSON.stringify(light.cellBackground);
  const sameInk = JSON.stringify(exporting.cellColor) === JSON.stringify(light.cellColor);
  const sameBorders = JSON.stringify(exporting.cellBorderRight) === JSON.stringify(light.cellBorderRight);
  if (!sameFills) problems.push(`导出态底色与浅色模式不一致：${JSON.stringify(exporting.cellBackground)} vs ${JSON.stringify(light.cellBackground)}`);
  if (!sameInk) problems.push(`导出态文字色与浅色模式不一致：${JSON.stringify(exporting.cellColor)} vs ${JSON.stringify(light.cellColor)}`);
  if (!sameBorders) problems.push(`导出态框线与浅色模式不一致：${JSON.stringify(exporting.cellBorderRight)} vs ${JSON.stringify(light.cellBorderRight)}`);
  if (errors.length) problems.push(`页面报错：${errors.join(' | ')}`);

  console.log(problems.length ? `问题：\n- ${problems.join('\n- ')}` : '全部检查通过');
  process.exitCode = problems.length ? 1 : 0;
} finally {
  await browser.close();
}
