import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const base = 'http://127.0.0.1:1420/';
const pkg = resolve('plugin-repo/packages/mini-office.theme-office-2003-1.0.0.moext');
const out = (name) => resolve('tests/output', name);

// 复用本机已缓存的 Chrome for Testing，避免因 playwright 版本不匹配而重新下载浏览器。
const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1280, height: 860 } });

const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto(base, { waitUntil: 'networkidle' });
await page.waitForTimeout(1500);
await page.screenshot({ path: out('plugin-01-home.png') });

// 打开设置（默认停在外观页）
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(800);
const settingsVisible = await page.getByText('安全与隐私').isVisible().catch(() => false);
console.log('settings visible:', settingsVisible);
await page.screenshot({ path: out('plugin-02-settings.png') });

// 进入插件管理：设置已改为分页结构，先切到「扩展」
await page.getByRole('button', { name: '扩展', exact: true }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: '管理插件' }).click();
await page.waitForTimeout(800);
const pluginDialog = await page.getByText('插件来源').isVisible().catch(() => false);
console.log('plugin dialog visible:', pluginDialog);
await page.screenshot({ path: out('plugin-03-manager.png') });

// 安装本地插件包
await page.setInputFiles('input[type=file][accept*=".moext"]', pkg);
await page.waitForTimeout(2500);
const installed = await page.getByText('已安装「Office 2003 经典外观」1.0.0').isVisible().catch(() => false);
console.log('installed:', installed);
await page.screenshot({ path: out('plugin-04-installed.png') });

// 关闭对话框，回到设置应用主题
await page.keyboard.press('Escape');
await page.waitForTimeout(400);
await page.keyboard.press('Escape');
await page.waitForTimeout(400);
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(800);
await page.getByRole('button', { name: 'Office 2003 经典' }).click();
await page.waitForTimeout(900);
const themeAttr = await page.evaluate(() => ({
  pack: document.documentElement.dataset.themePack,
  chrome: document.documentElement.getAttribute('data-chrome-control-shape'),
  density: document.documentElement.getAttribute('data-chrome-density'),
  bg: getComputedStyle(document.documentElement).getPropertyValue('--mo-bg').trim(),
}));
console.log('theme applied:', JSON.stringify(themeAttr));
await page.screenshot({ path: out('plugin-05-theme-applied.png') });

console.log('errors:', errors.length ? errors.slice(0, 8) : 'none');
await browser.close();
