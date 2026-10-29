# 默认应用与文件关联

<!-- 功能：补齐默认应用设置与 Windows 冷启动文件打开链路；时间：2026-10-29；作者：lq。 -->

> 相关反馈：Windows 11 用户把文档拖到 Mini Office 图标时，只启动软件，不打开文档。

## 1. 问题根因

反馈表象是“拖拽不打开文档”，实际由三个独立缺陷叠加：

| 缺陷 | 影响 | 处理 |
|---|---|---|
| 打包配置未声明 `fileAssociations` | 系统不知道本应用能打开 DOCX/XLSX/PPTX/CSV，双击与拖拽都不会把路径交给应用 | `tauri.conf.json` 补齐七组关联声明 |
| 桌面壳只处理 macOS 的 `RunEvent::Opened` | Windows/Linux 通过命令行参数传入的文件路径被忽略，表现为“只打开软件” | 启动时解析 `std::env::args()`，命中受支持文件即入队 |
| 未启用单实例 | 应用已在运行时再次拖入文件会启动第二个进程，两个窗口争用同一份恢复库 | 接入 `tauri-plugin-single-instance`（Windows/Linux），把参数转发到已运行窗口 |

## 2. 功能入口

- **首次启动提示**：工作区首帧完成后弹窗一次，列出四类文档的当前默认状态，可一键申请；关闭或选择“以后再说”后写入 `mini-office.default-app-prompt`，不再重复打扰。已全部为默认时静默跳过。
- **设置 → 默认应用**：常驻分区，展示 DOCX / XLSX / PPTX / CSV 的实时状态，提供“设为默认应用”（Windows 上为“去系统设置设为默认”）与“打开系统设置”两个动作。
- **菜单 → 设为默认应用…**、**命令面板 → 设为默认应用**：手动进入同一面板；手动进入不会写入“已提示”标记。
- 窗口重新获得焦点时自动刷新状态，用户从系统设置返回后无需重开面板。

## 3. 平台实现

| 平台 | 登记 | 申请默认 | 状态判断 |
|---|---|---|---|
| Windows | `HKCU\Software\Classes\MiniOffice.<ext>`（ProgID、图标、`shell\open\command`）、`.ext\OpenWithProgids`、`HKCU\Software\MiniOffice\Capabilities` + `RegisteredApplications` | 系统禁止程序静默改写，登记候选后打开 `ms-settings:defaultapps?registeredAppUser=MiniOffice` | 读取 `FileExts\.ext\UserChoice\ProgId`，ProgID 含 `MiniOffice` 或打开命令指向本应用即视为默认 |
| macOS | `Info.plist` 的 `CFBundleDocumentTypes`（打包配置生成，`LSHandlerRank=Alternate`，不抢占系统默认） | `LSSetDefaultRoleHandlerForContentType` 逐类型申请 | `LSCopyDefaultRoleHandlerForContentType` 与包标识 `com.minioffice.desktop` 比对 |
| Linux | 安装包写入 `.desktop` 的 `MimeType` | `xdg-mime default <desktop> <mime>` | `xdg-mime query default <mime>` |

Windows 的注册表登记通过 `reg.exe` 完成（`CREATE_NO_WINDOW`，不闪控制台），失败不影响打开设置页，只在提示文案中说明。安装包（NSIS/MSI）会按 `fileAssociations` 写入同类键值，运行时登记用于便携版与开发版。

## 4. 文件打开链路

1. 冷启动：`setup` 阶段解析启动参数，命中受支持扩展名且真实存在的文件入 `pending_files`。
2. 运行中：macOS 走 `RunEvent::Opened`；Windows/Linux 由单实例插件把第二次启动的参数转发给已运行实例。
3. 前端：`watchDesktopFiles` 先订阅 `mo:open-files` 再原子领取队列，串行打开，重复路径只保留一次，中文与空格路径由 `canonicalize` 归一。

## 5. 验证

- `pnpm typecheck`、`pnpm test`（203 项，含新增 `default-app-bridge`、`default-app-prompt`）、`cargo test --workspace`（39 项）通过。
- `office-platform::default_app` 新增 8 项单元测试：关联目标完整性、扩展名归一、启动参数过滤、`reg query` 解析、ProgID 归属判断、状态聚合、路径前缀归一。
- 面板在 Windows 状态下的排版已用本地预览页与浏览器截图核对（四行类型、状态徽标、说明与按钮）；设置页分区在 Web 模式下显示“仅桌面客户端可用”。
- `tauri build --debug --bundles app` 通过，生成的 `Info.plist` 含四组 `CFBundleDocumentTypes`（DOCX/DOTX、XLSX/CSV、PPTX/POTX、含宏只读），打包后启动冒烟 27 秒无错误输出。
- 未覆盖：Windows 11 真机拖拽到图标与 `ms-settings` 深链、macOS 真机 `LSSetDefaultRoleHandler` 生效、Linux 桌面环境验证，均需实机确认。

## 6. 已知限制

- Windows 8 以后系统不再提供静默设置默认应用的 API，任何“一键设为默认”都只能落到系统设置页由用户确认，因此文案与按钮使用“去系统设置设为默认”而非“已设为默认”。
- macOS 上若应用未放入“应用程序”或未正常注册 LaunchServices，申请可能被系统忽略，此时需按 [macOS 文档关联](./macos-file-associations.md) 中的“显示简介 → 全部更改”手动确认。
- 关联范围限定 DOCX/DOCM/DOTX、XLSX/XLSM/XLTX、PPTX/PPTM/POTX、CSV；旧版 DOC/XLS/PPT 仍只支持在应用内转换打开，不注册为默认程序。
