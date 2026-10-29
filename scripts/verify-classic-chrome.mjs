import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const pkg = resolve('plugin-repo/packages/mini-office.theme-office-2003-1.0.0.moext');
const out = (name) => resolve('tests/output', name);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(1200);

// 装主题包
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(700);
await page.getByRole('button', { name: '扩展', exact: true }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: '管理插件' }).click();
await page.waitForTimeout(600);
await page.setInputFiles('input[type=file][accept*=".moext"]', pkg);
await page.waitForTimeout(2200);
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
await page.keyboard.press('Escape');
await page.waitForTimeout(300);

// 应用经典主题
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(600);
await page.getByRole('button', { name: 'Office 2003 经典' }).click();
await page.waitForTimeout(700);
await page.keyboard.press('Escape');
await page.waitForTimeout(500);

const menubarCount = await page.locator('nav[aria-label="菜单栏"] button').count();
console.log('menubar items:', menubarCount);
await page.screenshot({ path: out('plugin-07-classic-home.png') });

// 打开一份新文档，观察"标题栏 + 菜单栏 + 工具栏"三段式
await page.getByRole('button', { name: /新建.*Word|Word 文档/ }).first().click();
await page.waitForTimeout(3500);
const editorMenubar = await page.locator('nav[aria-label="菜单栏"]').isVisible().catch(() => false);
console.log('menubar in editor:', editorMenubar);
await page.screenshot({ path: out('plugin-08-classic-editor.png') });

// 展开"文件"菜单确认可用
await page.locator('nav[aria-label="菜单栏"] button', { hasText: '文件' }).first().click();
await page.waitForTimeout(500);
const menuItems = await page.getByRole('menu').innerText().catch(() => '');
console.log('file menu:', menuItems.split('\n').filter(Boolean).slice(0, 8).join(' | '));
await page.screenshot({ path: out('plugin-09-classic-menu-open.png') });

console.log('errors:', errors.length ? errors.slice(0, 5) : 'none');
await browser.close();
