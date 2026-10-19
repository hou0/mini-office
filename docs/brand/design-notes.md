# Mini Office 图标设计

> 功能：品牌图标设计与交付记录；时间：2026-10-18；作者：lq。

应用：Mini Office；出品：非人类程序。品牌气质：克制、精致、可靠。沿用青绿色，形成独立办公应用识别。

## 三个方向

| 方向 | 类型与形状 | 含义 | 颜色 | 字体与布局 |
|---|---|---|---|---|
| 折页 M（采用） | 字母图形；两条纸面立柱与一条斜向折页构成 M | Mini 的首字母，同时表达纸页和整理；16px 下仍保留明确轮廓 | 墨青 #12373C、纸白 #FAF9F2、薄荷 #80D6C0 | 系统无衬线；可搭配 Noto Sans SC / Source Han Sans；图标在左、名称在右 |
| 文档叠页 | 抽象符号；三片错位的圆角文档 | 对应文档、表格和演示；以叠放表达统一工作空间 | 墨青 #12373C、纸白 #FAF9F2、浅青 #80D6C0 | 系统无衬线；图标在上、名称在下 |
| 书签印章 | 徽章；带负空间书签的圆角正方形 | 本地存档、可靠保存；适合单色印章 | 墨青 #12373C、纸白 #FAF9F2 | 系统无衬线；独立图标与水平组合 |

选择折页 M：比普通字母占位有更强的图形特征，小尺寸易读，并可保持单色轮廓。主体仅用三种色系，微渐变用于材质层次，识别不依赖渐变。

## 交付

- `mini-office.svg`：可编辑矢量主图；透明外沿 + 圆角底板。
- `apps/desktop/public/brand/mini-office.svg`：界面实际使用的同版图标。
- `apps/desktop/public/favicon.svg`：浏览器图标。
- `apps/desktop/src-tauri/icons/`：由 Tauri 官方图标转换器从同一 SVG 生成 PNG、ICO、ICNS。
- 首页、标题栏、关于及启动页共用 `BrandIcon`，保证一致。

设计方法：Logo Designer 工作流 + 手工 SVG；本轮未调用位图图像生成模型。没有进行商标注册或检索。

## 可复用图像提示词

1. 折页 M：Design an original Mini Office app icon, a sculptural folded-paper M, ivory paper and mint folded plane on a deep teal rounded square, restrained desktop craftsmanship, strong readable silhouette, no text, no extra objects, no existing brand likeness.
2. 文档叠页：Design an original Mini Office symbol made of three offset paper planes, minimal deep teal and ivory with one mint accent, precise rounded geometry, legible at 16 pixels, no text or existing brand likeness.
3. 书签印章：Design an original Mini Office emblem, one rounded square with a bookmark cutout in negative space, deep teal and ivory, quiet archival character, single-color compatible silhouette, no text or existing brand likeness.

后续可衍生：纯单色、深浅底反白、横排与上下排组合。本轮产品直接使用独立图标。
