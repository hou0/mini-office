/**
 * 版本更新提示验证：启动弹窗、小红点、忽略后不再重复弹、关于页手动检查。
 *
 * 用 route 拦截 GitHub API 返回一个更高的版本，避免为了测试真的去发一个 Release。
 */
import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const out = (name) => resolve('tests/output', name);

const FAKE_RELEASES = [
  {
    tag_name: 'v0.1.6-preview.1',
    name: 'Mini Office 0.1.6 preview',
    html_url: 'https://github.com/lq0910/mini-office/releases/tag/v0.1.6-preview.1',
    prerelease: true,
    draft: false,
    published_at: '2026-10-30T02:00:00Z',
  },
  {
    tag_name: 'v0.1.5-preview.2',
    name: 'Mini Office 0.1.5 preview 2',
    html_url: 'https://github.com/lq0910/mini-office/releases/tag/v0.1.5-preview.2',
    prerelease: true,
    draft: false,
    published_at: '2026-10-29T01:46:27Z',
  },
];

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.route('**/api.github.com/repos/lq0910/mini-office/releases**', (route) =>
  route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(FAKE_RELEASES) }),
);

// 1) 启动即提示
await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(2500);
let text = await page.locator('body').innerText();
console.log('startup dialog:', text.includes('发现新版本') && text.includes('v0.1.6-preview.1'));
console.log('shows version jump:', /0\.1\.5-preview\.2\s*→\s*v0\.1\.6-preview\.1/.test(text));
console.log('prerelease badge:', text.includes('预览版'));
await page.screenshot({ path: out('update-01-dialog.png') });

// 2) 稍后提醒 → 关闭并记录
await page.getByRole('button', { name: /稍后提醒/ }).click();
await page.waitForTimeout(600);
text = await page.locator('body').innerText();
console.log('dialog closed:', !text.includes('0.1.5-preview.2 → v0.1.6-preview.1'));

// 3) 重启（刷新）后不再弹窗，但小红点仍在
await page.reload({ waitUntil: 'networkidle' });
await page.waitForTimeout(2500);
text = await page.locator('body').innerText();
console.log('no repeat dialog:', !text.includes('0.1.5-preview.2 → v0.1.6-preview.1'));
const dots = await page.locator('[data-update-dot]').count();
console.log('red dot:', dots);
await page.screenshot({ path: out('update-02-dot.png') });

// 4) 「更多」菜单里能看到新版本入口
await page.getByRole('button', { name: /更多/ }).click();
await page.waitForTimeout(500);
const menuText = await page.locator('body').innerText();
console.log('menu entry:', menuText.includes('发现新版本 v0.1.6-preview.1'));
await page.screenshot({ path: out('update-03-menu.png') });
await page.keyboard.press('Escape');
await page.waitForTimeout(300);

// 5) 设置 → 关于：状态直接可见，并能手动检查
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(800);
await page.getByRole('button', { name: '关于', exact: true }).click();
await page.waitForTimeout(800);
const aboutText = await page.locator('body').innerText();
console.log('about shows update:', aboutText.includes('发现新版本 v0.1.6-preview.1'));
console.log('about check button:', aboutText.includes('有更新') || aboutText.includes('检查更新'));
await page.screenshot({ path: out('update-04-about.png') });

// 6) 独立关于弹窗里也有检查入口
await page.getByRole('button', { name: '查看详情' }).click();
await page.waitForTimeout(700);
const detailText = await page.locator('body').innerText();
console.log('about dialog row:', detailText.includes('版本更新') && detailText.includes('检查更新'));
await page.screenshot({ path: out('update-05-about-dialog.png') });

console.log('errors:', errors.length ? errors.slice(0, 5) : 'none');
await browser.close();
