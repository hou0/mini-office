// 功能：验证浏览器文件选择接口优先、旧 WebView 兼容和失败可被界面捕获；时间：2026-10-24；作者：lq。
import { afterEach, describe, expect, it, vi } from 'vitest';
import { openSheetImagePicker } from '../packages/editor-sheet/src/image-picker';
let input: HTMLInputElement | null = null;
afterEach(() => { input?.remove(); input = null; });
function fileInput() { input = document.createElement('input'); input.type = 'file'; document.body.append(input); return input; }

describe('Excel 图片文件选择器', () => {
  it('优先同步调用 showPicker，而不是依赖隐藏控件的合成点击', () => {
    const element = fileInput(); const show = vi.fn(); const click = vi.spyOn(element, 'click');
    Object.defineProperty(element, 'showPicker', { value: show });
    openSheetImagePicker(element);
    expect(show).toHaveBeenCalledOnce(); expect(click).not.toHaveBeenCalled();
  });
  it('旧浏览器没有 showPicker 时使用 click 兼容路径', () => {
    const element = fileInput(); Object.defineProperty(element, 'showPicker', { value: undefined });
    const click = vi.spyOn(element, 'click').mockImplementation(() => {});
    openSheetImagePicker(element); expect(click).toHaveBeenCalledOnce();
  });
  it('浏览器拒绝时保留失败信息，让界面可以显示原生文件控件重试', () => {
    const element = fileInput();
    Object.defineProperty(element, 'showPicker', { value: () => { throw new DOMException('需要用户手势', 'NotAllowedError'); } });
    expect(() => openSheetImagePicker(element)).toThrow('需要用户手势');
  });
  it('控件尚未挂载或已卸载时明确报错，而不是无声忽略点击', () => {
    expect(() => openSheetImagePicker(null)).toThrow('尚未准备好');
    const detached = document.createElement('input'); detached.type = 'file';
    expect(() => openSheetImagePicker(detached)).toThrow('尚未准备好');
  });
});
