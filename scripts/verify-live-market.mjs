import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const out = (name) => resolve('tests/output', name);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1360, height: 920 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(2500);
await page.screenshot({ path: out('plugin-10a-before.png') });
const initialText = await page.locator('body').innerText();
console.log('initial body head:', JSON.stringify(initialText.split('\n').filter(Boolean).slice(0, 8)));

// 进入插件管理
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(900);
const settingsOpen = await page.getByText('安全与隐私').isVisible().catch(() => false);
console.log('settings open:', settingsOpen);
if (!settingsOpen) {
  await page.screenshot({ path: out('plugin-10b-settings-fail.png') });
  await browser.close();
  process.exit(1);
}
await page.getByRole('button', { name: '扩展', exact: true }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: '管理插件' }).click();
await page.waitForTimeout(700);

// 从 GitHub 拉取真实索引
await page.getByRole('button', { name: '刷新列表' }).click();
await page.waitForTimeout(6000);

const bodyAfterRefresh = await page.locator('body').innerText();
const listed = ['公文排版助手', 'Office 2003 经典外观', 'WPS 风格外观', 'Microsoft 365 风格外观'].filter((name) =>
  bodyAfterRefresh.includes(name),
);
console.log('listed from GitHub:', listed.length, JSON.stringify(listed));
await page.screenshot({ path: out('plugin-10-market-live.png') });

// 从 GitHub 安装经典主题（走下载 + SHA-256 校验）
const installButtons = page.getByRole('button', { name: '安装', exact: true });
const count = await installButtons.count();
console.log('install buttons:', count);
if (count > 0) {
  await installButtons.first().click();
  await page.waitForTimeout(6000);
  const afterInstall = await page.locator('body').innerText();
  const ok = /已安装「/.test(afterInstall);
  console.log('install from github:', ok);
  await page.screenshot({ path: out('plugin-11-installed-from-github.png') });
}

console.log('errors:', errors.length ? errors.slice(0, 6) : 'none');
await browser.close();
