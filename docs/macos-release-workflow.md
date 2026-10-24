# macOS Release 打包流程

<!-- 功能：说明 macOS 通用包的触发方式、下载位置和发布限制；时间：2026-10-24；作者：lq。 -->

工作流文件：`.github/workflows/macos-release.yml`。

1. 更新项目版本号并将源码推送到 `main`。
2. 打开 GitHub → Actions → **macOS Release** 或 **Windows Release** → **Run workflow**，选择 `main`，标签输入框留空即可。
3. 工作流从所选源码的 `tauri.conf.json` 读取版本号，自动创建缺失的 `v<版本号>` 标签和 GitHub Release 草稿。无需手动创建标签或 Release；两平台的准备阶段串行执行，避免重复创建草稿，之后并行编译。
4. 完成后，在 GitHub Releases 的相应草稿中下载安装包并检查，确认后可自行点击 **Publish release**。运行详情页的 Artifacts 也保留产物 14 天。

标签输入框可填写已有标签来重新构建旧版本。已有标签始终使用原标签源码，不会移动或覆盖；新标签必须匹配所选源码版本号。标签、根 `package.json` 与 Tauri 版本不一致时停止发布。已有 Release 会复用，草稿不会自动公开，公开版本也不会被改成草稿。

每次构建生成一个兼容 Apple Silicon 与 Intel 的通用应用：

- `Mini-Office_<版本>_macos-universal.dmg`：安装镜像。
- `Mini-Office_<版本>_macos-universal.zip`：包含完整 `.app` 的压缩包。
- `SHA256SUMS-macos-universal.txt`：上述文件的 SHA256 校验清单。

构建使用输入标签的源码。因此，仅推送新的 `main` 提交不会改变旧标签打出的程序；要发布新代码，应先更新版本号并推送代码，再运行工作流，标签和草稿由流程自动创建。

流程使用 macOS 临时签名（ad-hoc），无需配置 Apple 证书 Secret；没有 Developer ID 签名及 Apple 公证，首次下载运行仍可能受到 Gatekeeper 限制。签名方式见 [Tauri 官方说明](https://v2.tauri.app/distribute/sign/macos/)。

上传不会覆盖同名附件。重复运行时，已有同名文件会导致 Release 上传步骤失败，新的产物仍可从该次运行的 Artifacts 下载。工作流不会修改已有 Windows 或其他 macOS 下载文件。
