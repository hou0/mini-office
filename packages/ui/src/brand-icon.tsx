/** 功能：统一显示 Mini Office 品牌图标，避免各界面继续使用字母占位；时间：2026-10-18；作者：lq。 */
export function BrandIcon({ size = 32, className = '' }: { size?: number; className?: string }) {
  return <img src="/brand/mini-office.svg" alt="Mini Office" width={size} height={size}
    draggable={false} className={`shrink-0 select-none ${className}`} style={{ width: size, height: size }} />;
}
