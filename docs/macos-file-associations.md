# macOS 文档关联与安全退出

> 功能说明：补齐文件关联、退出保护及 Excel 结构引用；时间：2026-10-18；作者：lq。
>
> 关联声明已迁移到 `tauri.conf.json` 的 `fileAssociations`，并新增默认应用申请与首次启动提示，参见[默认应用与文件关联](./default-app-associations.md)。

## 本轮实现

- Excel 插入/删除行列同时维护本表、跨表公式和命名区域，删除全部引用目标时输出 `#REF!`。
- 工作表重命名维护跨表引用并验证名称，删除工作表将相关引用置为 `#REF!`。
- 排序平移被移动公式中的相对行引用，保留绝对引用。修正 `$A1` 与 `A$1` 的解析含义。
- 不改写公式中的字符串及运算结构；遇到结构化引用等无法安全调整的语法时，取消此次结构操作并提示。
- 原生关闭窗口、Cmd+Q 和程序坞退出先等待前端逐份确认。取消或保存失败保留窗口及标签；确认期间新增编辑须重新确认；已有写盘/导出任务时暂不退出。
- 阻止重复并发保存，保存期间继续编辑不清除未保存标记。
- macOS 包声明 DOCX/DOTX、XLSX/CSV、PPTX/POTX 和含宏文档关联。PPT 与含宏文档声明为 Viewer，其他为 Editor；采用 Alternate 排名，不主动替换系统默认程序。
- Rust 接收 `RunEvent::Opened`，队列保留冷启动文件；前端初始化后先订阅再领取，串行加载多个文件。重复打开已有路径时激活原标签。

## 无证书使用与设置默认应用

默认文档关联不要求付费 Apple 开发者证书；应用必须能正常启动。下载后的未签名应用仍可能受到 Gatekeeper 首次启动限制。

1. 将新构建的 `Mini Office.app` 放入“应用程序”，先确认可正常启动。
2. 在访达选中一个 `.docx` 文件，打开“显示简介”。
3. 在“打开方式”选择 Mini Office，必要时从“其他”定位到应用。
4. 点击“全部更改”，为该类型的所有文件设置默认应用。
5. 对 `.xlsx`、`.csv` 和 `.pptx` 分别操作。

只有用户在系统中确认后才会变更默认打开方式。代码及打包配置不会自动改写用户的现有默认应用。

官方说明：[Apple 默认打开方式](https://support.apple.com/zh-cn/guide/mac-help/mh35597/mac)、[Apple 首次启动安全提示](https://support.apple.com/zh-cn/102445)、[Tauri 文件关联配置](https://v2.tauri.app/reference/config/#fileassociations)。

## 验证及范围

- 类型检查、62 项前端测试通过（原有 36 项，新增 26 项）。
- Rust workspace 19 项测试通过，桌面代码编译通过。
- `tauri build --debug --bundles app --no-sign` 完成，生成 `target/debug/bundle/macos/Mini Office.app`。
- 新 `.app` 启动冒烟通过：独立测试进程运行 5 秒，无错误输出，随后仅终止本次测试进程。
- 检查生成的 Info.plist 已包含四组 `CFBundleDocumentTypes`、角色及 Alternate 排名。
- 队列及退出保护有自动化逻辑测试；尚未完成访达冷/热启动、真实 Cmd+Q 弹窗、多文件逐份确认的桌面交互验收。
- 当前仍为调试构建，无 Developer ID 签名及公证。用户默认应用设置、证书、钥匙串和系统安全配置未修改。
- Word 超链接等复杂 OOXML 保存保真限制继续存在，本轮未扩展此范围。
