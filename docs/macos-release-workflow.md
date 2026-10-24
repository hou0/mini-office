# macOS Release 打包流程

<!-- 功能：说明 macOS 通用包的触发方式、下载位置和发布限制；时间：2026-10-24；作者：lq。 -->

工作流文件：`.github/workflows/macos-release.yml`。

1. 将工作流提交并推送到仓库默认分支 `main`。
2. 确认需要打包的源码已经创建标签，并且该标签已有 GitHub Release。标签格式须为 `v` 加上该标签中 `apps/desktop/src-tauri/tauri.conf.json` 的版本号，例如 `v0.1.1`。
3. 打开 GitHub → Actions → **macOS Release** → **Run workflow**，选择 `main`，输入 Release 标签并运行。
4. 完成后，在对应 Release 下载 DMG、ZIP 或 SHA256 校验文件。运行详情页的 Artifacts 也保留一份产物，有效期 14 天。

每次构建生成一个兼容 Apple Silicon 与 Intel 的通用应用：

- `Mini-Office_<版本>_macos-universal.dmg`：安装镜像。
- `Mini-Office_<版本>_macos-universal.zip`：包含完整 `.app` 的压缩包。
- `SHA256SUMS-macos-universal.txt`：上述文件的 SHA256 校验清单。

构建使用输入标签的源码。因此，仅推送新的 `main` 提交不会改变旧标签打出的程序；要发布新代码，应先更新版本号并创建对应的新标签和 Release。

流程使用 macOS 临时签名（ad-hoc），无需配置 Apple 证书 Secret；没有 Developer ID 签名及 Apple 公证，首次下载运行仍可能受到 Gatekeeper 限制。签名方式见 [Tauri 官方说明](https://v2.tauri.app/distribute/sign/macos/)。

上传不会覆盖同名附件。重复运行时，已有同名文件会导致 Release 上传步骤失败，新的产物仍可从该次运行的 Artifacts 下载。工作流不会修改已有 Windows 或其他 macOS 下载文件。
