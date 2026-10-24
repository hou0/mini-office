// 功能：同步调用浏览器原生图片选择器，调用失败交给可见重试界面处理；时间：2026-10-24；作者：lq。
export function openSheetImagePicker(input: HTMLInputElement | null): void {
  if (!input?.isConnected) throw new Error('图片选择器尚未准备好，请重新选择图片。');
  input.value = '';
  // 功能：优先使用明确的文件选择接口，保留旧浏览器和桌面 WebView 的点击兼容路径；时间：2026-10-24；作者：lq。
  if (typeof input.showPicker === 'function') input.showPicker();
  else input.click();
}
