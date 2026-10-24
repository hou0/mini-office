Mini Office 0.1.0 macOS 本地构建发布，提供安装包和应用运行包。

## 下载与使用

- `Mini-Office_0.1.0_aarch64.dmg`：安装包，打开后将 Mini Office 拖到 Applications（应用程序）。
- `Mini-Office_0.1.0_aarch64.app.zip`：应用运行包，解压后得到 Mini Office.app，可移到应用程序目录后打开。
- `SHA256SUMS.txt`：上述文件的 SHA-256 校验清单。

## 兼容性与首次启动

- 适用于 Apple Silicon（M 系列芯片）Mac；本次未提供 Intel 版本。
- 当前构建未使用 Apple Developer ID 签名，也未完成 Apple 公证。若 macOS 阻止打开，请先尝试右键应用选择“打开”；仍受阻时，可在“系统设置 → 隐私与安全性”中对该应用选择“仍要打开”。
- 安装包与运行包来自同一个现有本地构建（2026-10-19），应用主程序 SHA-256 已确认一致。

这是早期版本；本次发布核对了安装包完整性、应用版本、CPU 架构及文件校验值，未重新执行全套功能验收。
