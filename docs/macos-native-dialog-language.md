# macOS 打印面板语言

日期：2026-10-24；作者：lq。

## 问题与处理

Mini Office 桌面打印通过 Tauri `window.print()` 调用系统原生打印面板。检查已有调试应用包发现 `CFBundleDevelopmentRegion` 为 `English`，且没有 `CFBundleLocalizations` 中文声明。这是打印窗口与应用中文界面不一致的直接相关配置缺口。

新增 `apps/desktop/src-tauri/Info.plist`，将默认开发语言和支持语言声明为 `zh-Hans`。Tauri 构建时会合并该文件，也会在 macOS 开发可执行文件中嵌入相应 plist；不需要修改用户的系统语言或偏好设置。

参考：[Tauri 原生配置合并](https://v2.tauri.app/distribute/macos-application-bundle/#native-configuration)、[Apple 应用语言选择及系统界面语言排查](https://developer.apple.com/library/archive/qa/qa1828/_index.html)。后者为 Apple iOS 文档，解释通用的 Bundle 语言声明原则；macOS 实际效果仍需验收。

## 验证与使用

- `plutil -lint` 检查通过。
- `pnpm --filter @mini-office/desktop exec tauri build --debug --bundles app` 通过。
- 检查实际 `.app/Contents/Info.plist`，默认语言与支持语言均为 `zh-Hans`，原有 4 组文件关联保留。
- 新构建位于 `target/debug/bundle/macos/Mini Office.app`。需退出旧进程，再启动新构建；此修改不能通过网页热更新影响已运行进程的原生语言。
- 本次未替换已安装应用或重发旧 DMG，也未实际操作打印机。仍需在新构建中打开打印窗口、展开纸张/打印机设置确认中文；打印机厂商扩展能否显示中文，还取决于对应驱动的本地化资源。
- 网页模式调用浏览器打印，语言由浏览器控制，本项 macOS Bundle 配置只影响桌面应用。
