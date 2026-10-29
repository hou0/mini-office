import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const pkg = resolve('plugin-repo/packages/mini-office.pack-gov-doc-1.0.0.moext');
const out = (name) => resolve('tests/output', name);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(1800);

// 安装功能插件（本地包）
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(800);
await page.getByRole('button', { name: '扩展', exact: true }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: '管理插件' }).click();
await page.waitForTimeout(700);
await page.setInputFiles('input[type=file][accept*=".moext"]', pkg);
await page.waitForTimeout(2500);
const managerText = await page.locator('body').innerText();
console.log('installed:', managerText.includes('公文排版助手'));
console.log('shows command badge:', /6 个命令/.test(managerText));
console.log('shows template badge:', /1 个模板/.test(managerText));
await page.screenshot({ path: out('macro-01-installed.png') });

// 回首页确认模板出现
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
await page.keyboard.press('Escape');
await page.waitForTimeout(600);
const homeText = await page.locator('body').innerText();
console.log('template listed:', homeText.includes('党政机关公文'));
await page.screenshot({ path: out('macro-02-home-template.png') });

// 从插件模板新建
await page.getByText('党政机关公文', { exact: true }).first().click();
await page.waitForTimeout(3500);
const editorText = await page.locator('body').innerText();
console.log('template opened:', editorText.includes('关于') && editorText.includes('各有关单位'));
await page.screenshot({ path: out('macro-03-template-opened.png') });

// 运行宏：套用公文版式
await page.keyboard.press('Meta+K');
await page.waitForTimeout(600);
await page.keyboard.type('套用公文版式');
await page.waitForTimeout(500);
await page.screenshot({ path: out('macro-04-palette.png') });
await page.keyboard.press('Enter');
await page.waitForTimeout(1500);
const afterMacro = await page.locator('body').innerText();
console.log('macro ran:', afterMacro.includes('已执行「套用公文版式」'));
await page.screenshot({ path: out('macro-05-after-macro.png') });

// 读回文档模型，确认页边距真的变了
const margins = await page.evaluate(() => {
  const host = document.querySelector('[data-page-metrics]');
  return host ? host.getAttribute('data-page-metrics') : null;
});
console.log('page metrics attr:', margins);

console.log('errors:', errors.length ? errors.slice(0, 5) : 'none');
await browser.close();
