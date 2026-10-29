/**
 * 功能：端到端验证 Word 样式库——面板展示、批量套用、显示更多、新建样式与样式集保存；
 * 时间：2026-10-29；作者：lq。
 *
 * 运行前提：apps/desktop 已构建，且 `vite preview --port 1420` 正在运行。
 */
import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const out = (name) => resolve('tests/output', name);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1440, height: 940 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(1200);

// 新建 Word 文档
await page.getByRole('button', { name: /新建.*Word|Word 文档/ }).first().click();
await page.waitForTimeout(2500);

// 输入三段文字
const editor = page.locator('[data-word-editor] [data-paragraph-id]').first();
await editor.click();
await page.keyboard.type('第一章 项目概况');
await page.keyboard.press('Enter');
await page.keyboard.type('本项目为轻量离线办公套件的样式能力验证，正文段落用于检查批量排版效果。');
await page.keyboard.press('Enter');
await page.keyboard.type('1.1 建设目标');
await page.waitForTimeout(500);
await page.screenshot({ path: out('styles-01-before.png') });

// 选中前两段后打开样式库
await page.keyboard.press('Meta+A');
await page.waitForTimeout(300);
await page.getByRole('button', { name: '样式库' }).click();
await page.waitForTimeout(600);
const cardCount = await page.locator('button[data-style-card="1"]').count();
console.log('basic cards:', cardCount);
await page.screenshot({ path: out('styles-02-gallery.png') });

// 套用标题 1：只作用于第一段
await page.keyboard.press('Escape');
await page.waitForTimeout(200);
await page.locator('[data-paragraph-id]').first().click();
await page.waitForTimeout(200);
await page.getByRole('button', { name: '样式库' }).click();
await page.waitForTimeout(400);
await page.locator('button[data-style-card="1"]', { hasText: '标题 1' }).first().click();
await page.waitForTimeout(900);
const headingStyle = await page.locator('[data-paragraph-id]').first().locator('span').first().evaluate((el) => {
  const style = getComputedStyle(el);
  return { fontSize: style.fontSize, fontWeight: style.fontWeight };
});
console.log('heading run style:', JSON.stringify(headingStyle));
const currentLabel = await page.getByRole('button', { name: '样式库' }).innerText();
console.log('current style label:', currentLabel.replace(/\s+/g, ' '));
await page.screenshot({ path: out('styles-03-applied.png') });

// 显示更多样式
await page.getByRole('button', { name: /显示更多样式/ }).click();
await page.waitForTimeout(400);
const moreCount = await page.locator('button[data-style-card="1"]').count();
console.log('cards after more:', moreCount);
await page.screenshot({ path: out('styles-04-more.png') });

// 新建样式
await page.getByRole('button', { name: /新建样式/ }).click();
await page.waitForTimeout(500);
await page.getByPlaceholder('例如：合同条款').fill('章标题');
await page.waitForTimeout(200);
await page.screenshot({ path: out('styles-05-new-style.png') });
await page.getByRole('button', { name: '创建' }).click();
await page.waitForTimeout(800);

// 保存为样式集
await page.getByRole('button', { name: '样式库' }).click();
await page.waitForTimeout(400);
const docStylesSection = await page.locator('text=本文档样式').count();
console.log('document styles section:', docStylesSection);
await page.getByRole('button', { name: /保存此文档为我的样式集/ }).click();
await page.waitForTimeout(500);
await page.getByRole('button', { name: '保存', exact: true }).click();
await page.waitForTimeout(800);
await page.getByRole('button', { name: '样式库' }).click();
await page.waitForTimeout(500);
const sets = await page.locator('text=我的样式集').count();
console.log('style sets section:', sets);
await page.screenshot({ path: out('styles-06-style-set.png') });

// 清除格式
await page.locator('[data-paragraph-id]').first().click();
await page.waitForTimeout(200);
await page.getByRole('button', { name: '样式库' }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: /清除格式/ }).click();
await page.waitForTimeout(800);
const cleared = await page.locator('[data-paragraph-id]').first().locator('span').first().evaluate((el) => getComputedStyle(el).fontSize);
console.log('font size after clear:', cleared);
await page.screenshot({ path: out('styles-07-cleared.png') });

console.log('errors:', errors.length ? errors.join(' | ') : 'none');
await browser.close();
