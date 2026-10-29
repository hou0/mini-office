# Mini Office 插件扩展系统设计

> 状态：设计稿（待评审）｜最后更新：2026-10-29｜作者：lq
>
> 关联：ADR-0005 插件扩展系统

## 1. 结论先行

"主体保持 mini" 的正确做法不是"少做功能"，而是**内核只保留不可替代的部分，其余全部可插拔**。
三条主线：

1. **内核收敛为三件事**：文档引擎、宿主外壳、稳定契约。任何新增需求先问"能不能做成插件"。
2. **内置功能也走插件 API**（dogfooding）。邮件合并、主题色、导出目标都以贡献点方式注册，
   保证 API 被真实使用，而不是设计出来没人用。
3. **插件分三级信任模型**。绝大多数客户定制（主题、模板、快捷键）属于最低风险的声明式插件，
   不需要执行代码——这是"轻量"能否守住的成败点。

## 2. 内核边界

### 2.1 判据

> **如果一个功能出错，会不会导致用户文件损坏、内容丢失或安全失守？**
> 会 → 内核；不会 → 插件。

这条判据优先于"功能是否常用"。常用但可失败的功能（如批量导出）应当是插件；
不常用但会毁文件的功能（如未知 OOXML 部件保留）必须在内核。

### 2.2 必须留在内核

| 模块 | 现有位置 | 理由 |
|---|---|---|
| OOXML 包读写、未知部件保留 | `packages/ooxml` | 决定文件是否被静默破坏 |
| 安全扫描、L0–L3 能力分级 | `office-security`、`capabilities.ts` | 决定只读/另存策略 |
| 文档模型与生命周期状态机 | `packages/document-model` | 所有编辑器的唯一状态源 |
| 基础排版：分页测量 / 虚拟网格 / 幻灯片渲染 | `editor-word`、`editor-sheet`、`editor-slides` | 排版结果即文件语义 |
| 原子写入、时间戳备份、自动恢复 | `office-storage`、`platform-adapter` | 数据安全底线 |
| 主题令牌 schema 与外观管线 | `design-system` | 主题的**契约**不能由主题自己定义 |
| 插件宿主（加载/校验/权限/生命周期） | 新增 `plugin-host` | 信任根必须在内核 |

### 2.3 可以插件化

以下能力当前内置，但应当在 P2 之后逐步迁出为插件（或至少以插件 API 实现）：

- 邮件合并与批量报告（`MailMergePanel`）
- 表格高级工具：条件格式、数据透视、分列
- 页眉页脚高级设置、多级列表
- 格式转换：Markdown / ODT / HTML 往返
- 校对、翻译、OCR、AI 助手
- 模板库、印章、条码、公式编辑器

> 迁移原则：**不为了插件化而重写已稳定的功能**。先让新功能走插件 API，
> 老功能在新 API 被验证后（≥3 个真实插件）再迁。

## 3. 三级信任模型

这是整套设计的核心。插件的能力等级决定它的执行环境、审核门槛与失败影响面。

| 级别 | 形态 | 执行环境 | 能力 | 安装门槛 | 典型场景 |
|---|---|---|---|---|---|
| **T1 声明式** | `manifest.json` + `tokens.json` + 可选 `extra.css` | 不执行代码，宿主解析 | 主题、模板、命令别名、快捷键映射、工具栏布局、设置表单（JSON Schema） | 自动可用，无需审核 | Office 2003 主题、WPS 主题、公文模板、快捷键方案 |
| **T2 沙箱** | JS bundle | 独立 Worker（逻辑）+ iframe（UI），仅 RPC | 结构化读写文档、注册命令/面板、导入导出转换、私有存储 | 展示权限，一键确认 | Markdown 导出、格式检查、批量重命名、数据导入 |
| **T3 受信** | JS + Rust 侧能力 | 同进程 + 原生桥 | 文件系统、网络、外部程序、原生打印扩展、AI 调用 | 签名 + 显式信任 + 权限清单 | OCR、AI 写作、电子签章、企业系统对接 |

**关键判断：T1 应覆盖 80% 的定制需求。** 主题、模板、快捷键、菜单调整、工具栏精简——
这些是客户最常提的需求，且完全不需要跑代码。把 T1 做扎实，内核就不必为个别客户膨胀。

T2 的 UI 使用 iframe 而非同进程 React，原因是：插件 UI 一旦与宿主同进程，
就没有任何技术手段阻止它读取 `document`、改写编辑器 DOM、绕过权限。
代价是样式隔离（通过注入 `--mo-*` 令牌解决）与轻微交互延迟，这个代价可以接受。

### 3.1 T3 里的"签名、审计、商店"到底指什么

这三个词容易被当成同一件事，其实是三个独立机制，成本差异很大：

| 名词 | 实际含义 | 解决什么问题 | 成本 |
|---|---|---|---|
| **签名** | 用发布者的私钥（Ed25519）对插件包做密码学签名，宿主用公钥验签 | 证明"这个包确实出自某发布者，且发布后没有被改过一个字节"。类似 macOS 应用的开发者签名 | 需维护密钥对、公钥分发与轮换 |
| **审计** | 记录每个插件声明了什么权限、何时被启用、读过或写过哪些文件 | 出事可追溯：客户报"文档被改坏了"，能定位到是哪个插件在什么时候做的 | 需要本地审计表 + 查看界面 |
| **商店** | 官方审核 + 分发平台：提交、审核、上架、版本灰度、统计、支付 | 让用户能"发现"插件，并让官方对插件质量背书 | 需要服务器、审核人力、运营，**最重** |

**为什么本方案把它们从 P3 降级了**：客户要的是"能装插件"，不是"有一个插件市场"。
用 GitHub 仓库做索引（见 §10），商店的三项成本直接归零；索引里的 SHA-256 校验
承担了签名 90% 的实用价值（防篡改、防下载不完整），只是不能证明"发布者身份"。

因此实际落地顺序调整为：

- **P1 就做**：SHA-256 完整性校验（已在实现中）+ 权限清单展示。这解决"包装坏了/被换了"。
- **P2 再做**：本地审计表。这解决"出问题查不到"。
- **只有企业侧载才需要真签名**：MDM 下发到受管目录时，用企业公钥验签，
  防止用户手动替换成恶意插件。个人用户场景下签名收益不足以抵消密钥管理成本。

> 结论：**"商店"这个词可以从路线图里去掉**，替换为"GitHub 插件仓库"。
> 签名与审计保留，但按需启用，不作为个人用户的前置条件。

## 4. 贡献点（Contribution Points）

插件通过声明式清单注册能力，宿主按贡献点类型装配。所有 id 使用命名空间
`<publisher>.<plugin>.<pointId>`，避免冲突。

| 贡献点 | 内容 | 状态 | 现有落点 |
|---|---|---|---|
| `themes` | 主题包（Token + Chrome） | ✅ 已实现 | `design-system`、`plugin-host/theme.ts` |
| `commands` | 命令、快捷键、关键词；**v1 只支持宏** | ✅ 已实现 | `CommandPalette.tsx` + `plugin-host/macro.ts` |
| `templates` | 新建文档模板 | ✅ 已实现 | `HomePage.tsx` 模板区 + 插件文件存储 |
| `menus` | 菜单项，按锚点 id 插入 | ⬜ | `chrome.tsx`、`ui/menu.tsx` |
| `toolbars` | 工具栏按钮组，按 `word` / `sheet` / `slides` 分区 | ⬜ | `WordToolbar`、`SheetToolbar`、`SlidesToolbar` |
| `panels` | 右侧属性面板、底部面板 | ⬜ | `AppShell` 插槽 |
| `dialogs` | 模态框 | ⬜ | `dialogs.tsx` |
| `documentHooks` | `onOpen` / `onWillSave` / `onDidSave` / `onExport` / `onPrint` | ⬜ | store 生命周期 |
| `formats` | 新文件格式 handler | ⬜ | `registerFormatHandler`（内核已有，插件侧未开放） |
| `exporters` | 新导出目标 | ⬜ | `export-service` |
| `settings` | 设置分区 | ⬜ | 设置页 |
| `statusBar` | 状态栏项 | ⬜ | `AppShell` |

**未实现的贡献点会被拒绝安装，而不是静默忽略。** 清单里写了 `panels` 却什么也没发生，
对插件作者是最难排查的一类问题；`validateManifest` 因此对 `contributes` 的键做了白名单检查。

`commands` 在 v1 是**宏**：JSON 描述的操作序列，由宿主按白名单解释执行，插件不跑代码。
操作清单与 `scope` 语义见[插件宏参考](./plugin-macros.md)。

`formats` 贡献点已有现成实现（`packages/document-model/src/store.ts` 的
`registerFormatHandler` + 懒加载），说明注册表模式在本项目已被验证，
其余贡献点应沿用同一套 `register / dispose` 语义。

## 5. 主题系统设计

### 5.1 现状评估：条件已经具备

- 全部 UI 走 CSS 自定义属性（`packages/design-system/src/tokens.css`）
- 主题切换唯一入口是 `applyTheme` / `applyAccent` / `applyAppearance`
- 通过 `data-theme="light|dark"` 与 `data-accent` 在 `documentElement` 上选择令牌集

这意味着**主题在技术上就是一个令牌字典**，不需要重新渲染任何组件。

### 5.2 但只换颜色做不出 Office 2003

Office 2003 / XP 的辨识度来自**结构**而非配色：直角、渐变标题栏、凸起/凹陷边框、
菜单栏 + 工具栏两层、控件密集、Tahoma 9pt / 宋体。仅替换 `--mo-surface` 只会得到一个
"换了底色的现代界面"，客户不会认。

因此主题包分两层：

#### L1 — Token 层（纯数据，T1）

```json
{
  "id": "office-2003",
  "label": "Office 2003 经典",
  "light": {
    "--mo-bg": "#ECE9D8",
    "--mo-surface": "#FFFFFF",
    "--mo-surface-2": "#F1EFE2",
    "--mo-surface-3": "#DED8C3",
    "--mo-line": "#ACA899",
    "--mo-line-strong": "#716F64",
    "--mo-text-primary": "#000000",
    "--mo-radius-md": "0px",
    "--mo-radius-lg": "0px",
    "--mo-shadow-panel": "none",
    "--mo-font-ui": "Tahoma, SimSun, sans-serif"
  },
  "dark": { "...": "..." },
  "defaultAccent": "classic-blue"
}
```

校验规则（缺失即拒绝激活，不做静默兜底）：

1. **完整性**：所有必填令牌必须给出；`dark` 变体可缺省（回退到 light 派生）。
2. **白名单**：只允许 `--mo-*` 中标记为 `themeable` 的令牌；`--mo-doc-*` 一律禁止（见 5.4）。
3. **对比度**：`text-primary` / `surface`、`accent-ink` / `accent` 需满足 WCAG AA（≥4.5:1），
   不达标时拒绝激活并给出具体令牌对，避免"装上就看不见字"。

#### L2 — Chrome 层（受约束枚举，T1）

```json
"chrome": {
  "toolbar": "classic-two-row",   // modern | classic-two-row | ribbon
  "controlShape": "beveled",      // flat | beveled | rounded
  "density": "compact",           // compact | default | comfortable
  "titlebar": "gradient"          // flat | gradient
}
```

主题只能从枚举中选择，**不能注入 DOM 结构**。前提是宿主需要先把外壳抽成
"插槽 + 变体"骨架：`AppShell` 提供 `titlebar` / `menubar` / `toolbar` / `canvas` / `statusbar`
五个插槽，每个插槽支持 2–3 个渲染变体。

这是本设计里工作量最大、也最容易被低估的部分。没有 L2，主题插件就只是换肤；
有了 L2，才谈得上"Office 风格 / WPS 风格"。

#### L3 — 可选 `extra.css`（受控）

- 作用域强制限定在 `[data-theme-pack="<id>"]` 内
- 禁止 `@import`；`url()` 仅允许 `plugin://` 本地资源，禁止外链
- 体积上限 64 KB；安装期做 CSS AST 校验，只放行声明块与简单选择器
- 用于补齐 L1/L2 无法表达的效果（如 XP 标题栏渐变贴图、按钮凸起边框）

### 5.3 主题包结构

主题令牌**内联在 manifest.json 里**，不单独拆 `tokens.json`。理由是插件包越小越好，
多一个文件就多一次读取与一次"文件缺失"的失败可能；清单本身就是元数据载体，内联最省。

```text
theme-office-2003/
├── manifest.json      # id / 版本 / engines / 权限(空) / contributes.themes（含令牌与 chrome）
├── preview.png        # 设置页缩略图，可选
├── extra.css          # L3，可选（当前实现未启用，见 §11）
└── assets/            # 主题引用的本地资源，可选
```

`L3 extra.css` 在**当前实现中不启用**。原因：它需要 CSS AST 级的作用域重写与
外部 URL 拦截，收益（少数无法用令牌表达的效果）不足以抵消引入的复杂度与风险。
先只交付 L1 + L2；如果三个官方主题都做完仍有表达不出来的效果，再评估是否引入。

### 5.4 必须守住的边界：主题不能影响文档渲染

这是文档编辑器区别于普通应用的关键约束。

`--mo-doc-*`（页面字体、行高、纸张底色、标尺刻度）与 UI 令牌**强制分离**。
主题**只允许改 chrome，不允许改文档画布内的字体度量**。原因：

- 分页依赖字体度量测量（`editor-word/src/pagination.ts`）。若主题把文档字体从
  "PingFang SC" 换成 "Tahoma"，同一份 DOCX 的页数会变，用户看到的页码与打印结果不一致。
- 打印/PDF 导出走独立的样式通道，主题一旦渗入会造成"屏幕所见 ≠ 打印所得"。

允许的最小让步：主题可提供**纸张色调**（`docPaperTint`，十六进制色值），
宿主将其写入 `--mo-doc-paper`，只影响画布背景，不参与任何度量计算。

### 5.5 与现有强调色的关系

`accent` 是与主题正交的维度：主题提供 `defaultAccent`，用户仍可在设置中覆盖。
`applyAppearance` 演进为：

```ts
applyThemePack(themeId: string, opts: { mode: ThemeMode; accent?: AccentId; customAccent?: string })
```

内部顺序：`加载 token → 应用 chrome 变体 → 应用 extra.css → 应用 accent`。
accent 必须最后应用，否则会被主题自带的 accent 覆盖。

现有 5 套强调色（`ACCENT_PRESETS`）保持不变，作为"任意主题可用的强调色"。
P0 阶段可先把它们包装成内置主题包以验证管线，但 accent 维度不必迁移。

## 6. 运行时架构

```text
┌──────────────────────────────────────────────────────────┐
│ Core Shell (React)                                       │
│   AppShell · CommandPalette · Toolbars · Dialogs          │
│   └ ContributionRegistry（单例、版本化、可枚举/可回滚）      │
├──────────────────────────────────────────────────────────┤
│ Plugin Host (TypeScript)                                 │
│   ManifestValidator · PermissionBroker · LifecycleManager │
│   ThemeEngine · CommandBus · T1Loader · T2WorkerBridge    │
├──────────────────────────────────────────────────────────┤
│ @mini-office/plugin-api   ← 唯一对外稳定契约                │
├──────────────────────────────────────────────────────────┤
│ Document Model · OOXML · PlatformAdapter · export-service │
├──────────────────────────────────────────────────────────┤
│ Rust: office-plugin-host（解包 / 签名 / 权限审计）          │
└──────────────────────────────────────────────────────────┘
```

**铁律：插件只能 `import '@mini-office/plugin-api'`，不得引用任何内部包。**
这样内核可以自由重构，插件不会随之碎裂。`plugin-api` 只依赖 `shared` 的基础类型。

## 7. 插件 API 草案

```ts
export interface PluginManifest {
  id: string;                       // com.example.markdown-export
  name: string;
  version: string;                  // semver
  publisher: string;
  engines: { miniOffice: string; api: string };
  tier: 'declarative' | 'sandboxed' | 'trusted';
  permissions: Permission[];
  contributes: Contributions;
}

export type Permission =
  | 'document:read' | 'document:write'
  | 'fs:read' | 'fs:write'
  | 'network' | 'clipboard' | 'print' | 'shell:exec';

export interface PluginApi {
  commands: {
    register(cmd: CommandContribution): Disposable;
    execute(id: string, args?: unknown): Promise<unknown>;
  };
  ui: {
    registerPanel(p: PanelContribution): Disposable;
    registerToolbarButton(p: ToolbarContribution): Disposable;
    showToast(level: 'info' | 'warning' | 'error', title: string, detail?: string): void;
    openDialog(d: DialogDescriptor): Promise<unknown>;
  };
  documents: {
    active(): DocumentHandle | null;
    onDidChangeActive(cb: (h: DocumentHandle | null) => void): Disposable;
    onDidOpen(cb: (h: DocumentHandle) => void): Disposable;
    onWillSave(cb: (h: DocumentHandle) => void | Promise<void>): Disposable;
  };
  storage: {
    get<T>(key: string): Promise<T | undefined>;
    set<T>(key: string, value: T): Promise<void>;
  };
  net?: { fetch(input: string, init?: RequestInit): Promise<Response> }; // 需 network
}
```

### 7.1 文档访问：结构化事务，不是 DOM

```ts
interface DocumentHandle {
  readonly kind: 'word' | 'sheet' | 'slides';
  readonly path?: string;
  read(): Promise<DocumentSnapshot>;                                  // 只读快照
  transaction(fn: (tx: DocTransaction) => void): Promise<void>;       // 原子修改
}
```

`transaction` 是硬性要求：插件对文档的修改必须合并为**一步可撤销操作**。
否则一个批量替换插件会让用户的撤销栈炸开，这是插件体系最常见的体验灾难。

快照按需拉取、事务批量提交，避免逐字符 RPC 造成卡顿。

## 8. 安全模型

### 8.1 权限

| 权限 | 含义 | 默认授予 |
|---|---|---|
| `document:read` | 读取当前文档内容 | T2 是 |
| `document:write` | 修改文档 | 否，需确认 |
| `fs:read` / `fs:write` | 访问文档以外的文件 | 否 |
| `network` | 对外发起请求 | 否 |
| `clipboard` | 读写剪贴板 | 否 |
| `print` | 触发打印 | 否 |
| `shell:exec` | 调用外部程序 | 否，仅 T3 |

### 8.2 关键规则

1. **组合即敏感**：`network` + `document:read` 意味着文档内容可能外发。
   首次触发时二次确认，并在状态栏常驻"外发中"提示。
2. **T1 无权限**：声明式插件不执行代码，权限清单必须为空，否则校验失败。
3. **钩子超时降级**：`onWillSave` 等钩子 3 秒超时后放行并记录告警，
   绝不允许插件卡死导致用户无法保存——这是数据安全红线。
4. **崩溃隔离**：T2 Worker 崩溃只禁用该插件；T3 的每个钩子调用包 try/catch。

### 8.3 Rust 侧 `office-plugin-host`（P3，按需启用）

- 解包 `.moext`（复用 `office-security` 的压缩炸弹与路径穿越检查）
- 注册 `plugin://` 自定义协议，只读映射插件目录，拒绝 `..` 与绝对路径
- Ed25519 签名校验**仅用于企业侧载**：MDM 下发的受管插件用企业公钥验签，
  防止用户手动替换成恶意包。个人用户场景不做签名，避免密钥管理成本
- 权限清单哈希写入 SQLite 审计表，支持"这个插件改过哪些文件"追溯（P2 先做本地版）

## 9. 版本与兼容

- `engines.miniOffice: ">=0.1.5 <0.5.0"`、`engines.api: "^1.0.0"`
- `plugin-api` 遵循 semver；宿主暴露 `apiVersion` 供插件运行时判断
- 弃用流程：标记 `@deprecated` → 保留 2 个 minor 版本 → 移除
- **不兼容时禁用而非删除**：设置页显示原因与"检查更新"入口，用户数据（插件私有存储）保留

## 10. 分发渠道：GitHub 仓库即插件市场

不建商店。官方插件放在一个公开 GitHub 仓库里，客户端通过 GitHub API 拉取索引。
这样服务器成本、审核流程、运营人力全部为零，而用户能拿到的能力完全一样。

### 10.1 渠道一览

| 渠道 | 位置 | 说明 |
|---|---|---|
| 内置 | 随安装包 | 不可卸载，用于官方主题与样板插件 |
| **GitHub 仓库** | `index.json` + `packages/*.moext`（raw 分发） | **主渠道**，客户端自动拉取列表并一键安装 |
| 本地文件 | `*.moext`（zip） | 客户定制、协作者交付、离线环境 |
| 企业侧载 | 受管目录 + MDM | 强制启用、不可禁用、需签名 |
| 插件文件存储 | IndexedDB 虚拟文件系统 | 当前实现；P2 可切换为 `~/.mini-office/plugins/` 以支持热重载 |

### 10.2 官方插件仓库结构

脚手架已内置在本仓库的 **`plugin-repo/`**，作为独立仓库（`lq0910/mini-office-plugins`）推送即可使用。

```text
plugin-repo/                   （推送为独立仓库）
├── index.json                 # 客户端拉取的插件索引（脚本生成）
├── package.json
├── plugins/
│   ├── theme-office-2003/manifest.json
│   ├── theme-wps/manifest.json
│   └── theme-office-365/manifest.json
├── packages/                  # 打包产物，必须提交（见 §10.5 的 CORS 约束）
│   ├── mini-office.theme-office-2003-1.0.0.moext
│   ├── mini-office.theme-office-365-1.0.0.moext
│   └── mini-office.theme-wps-1.0.0.moext
├── scripts/build-index.mjs    # 打包 + SHA-256 + 生成索引
└── .github/workflows/release.yml
```

已上线：`https://github.com/lq0910/mini-office-plugins`。
`.moext` 提交进 `packages/`，客户端通过 raw 地址下载；同时挂一份到 GitHub Release
作为人工下载镜像。

**打包必须可重现**：脚本固定了 zip 条目时间戳，同一份源码重复打包得到相同字节与
相同 SHA-256。否则 CI 每次重新打包都会改变哈希，索引里的校验值与仓库里的包对不上，
用户安装时会报"插件包校验失败"。这一点已在实现中用连续两次打包比对、以及本地与
CI 产物哈希比对双重验证。

### 10.3 `index.json` 规范

```json
{
  "schemaVersion": 1,
  "updatedAt": "2026-10-29T06:00:00Z",
  "plugins": [
    {
      "id": "mini-office.theme-office-2003",
      "name": "Office 2003 经典外观",
      "version": "1.0.0",
      "publisher": "mini-office",
      "description": "还原 Windows XP / Office 2003 的 Luna 风格。",
      "tier": "declarative",
      "permissions": [],
      "download": "https://raw.githubusercontent.com/lq0910/mini-office-plugins/main/packages/mini-office.theme-office-2003-1.0.0.moext",
      "sha256": "63289ec7e4dec3bbfdea121b0c10201e75231384b2fbb09cc36055a460a8da8f",
      "size": 1013,
      "license": "MIT",
      "tags": ["theme", "classic"]
    }
  ]
}
```

约束：

- `sha256` 必填且为 64 位小写十六进制。**缺少或格式不符的条目会被客户端直接丢弃**，
  这样即使索引被篡改，也无法塞入一个没有校验值的包。
- `download` 必须是**浏览器可直接 fetch 的地址**（见 §10.5 的 CORS 约束），
  支持任意 HTTPS 域名，便于企业自建镜像。
- 未知字段被忽略，便于将来扩展而不破坏旧客户端。

### 10.4 拉取策略（已实现）

1. 优先 `GET https://api.github.com/repos/{owner}/{repo}/contents/index.json?ref={branch}`，
   带 `Accept: application/vnd.github.raw` 与 `If-None-Match`（ETag）。
   **304 命中不消耗配额**，正常使用几乎不会碰到 60 次/小时的未认证限流。
2. 遇到 403/429 或网络失败，回退到 `raw.githubusercontent.com` 的静态地址。
3. 仍失败则使用本地缓存并在界面上标注"当前为缓存列表"。
4. 仓库坐标（owner / repo / branch / indexPath）写在 `DEFAULT_PLUGIN_REPOSITORY`，
   可通过 `setRepository` 覆盖，便于客户指向私有仓库镜像。

### 10.5 安装链路与 CORS 约束

```text
本地文件：用户选 .moext → 解析 zip → 校验清单/令牌/路径 → 写入 IndexedDB → 记录到 localStorage
GitHub：拉索引 → 下载包 → 校验 SHA-256 → 同上
```

**插件包必须通过 `raw.githubusercontent.com` 分发，不能用 Release 资产。**
这是实测得出的硬约束：

| 通道 | CORS | 可用性 |
|---|---|---|
| `api.github.com` | `Access-Control-Allow-Origin: *` | 可用 |
| `raw.githubusercontent.com` | `Access-Control-Allow-Origin: *` | 可用 |
| `github.com/.../releases/download/...` | 无（302 到 `objects.githubusercontent.com`，该域不带头） | **浏览器与 Tauri WebView 均被拦截** |

Tauri 的 WebView 与浏览器遵循同一套同源策略，所以这不是"只有 Web 模式才有的问题"。
解决方式是把 `.moext`（约 1 KB）提交进仓库的 `packages/`，索引指向 raw 地址；
Release 仅作为人工下载的镜像保留。

副作用：raw 有几分钟 CDN 缓存，索引更新后可能短暂读到旧版本。由于客户端优先走
`api.github.com`（无缓存），这个窗口基本不会影响真实用户；只有 API 限流回退到 raw 时才可能遇到。

任何一步失败都不落地。安装后插件出现在「设置 → 扩展 → 插件」，主题类插件
同时出现在「外观 → 界面风格」。

## 11. 分阶段落地

> 状态：**P0 与 P1 的核心部分已实现**（见 §11.5 实现清单）。

### P0 — 契约与注册表（零用户可见变化）

- 新建 `packages/plugin-api`（纯类型 + 常量）
- 新建 `packages/plugin-host`（校验 / 主题引擎 / GitHub 客户端 / 存储 / 注册表）
- **验收**：现有测试全绿，UI 像素级无变化

### P1 — T1 声明式插件 + 主题（客户可见）

- `manifest.json` 解析与校验（完整性 / 令牌白名单 / 对比度 / engines）
- `.moext` 包解析：条目数、解压体积、路径穿越、清单合法性
- 主题包加载：L1 令牌 + L2 chrome 变体（`data-chrome-*` 属性 + 外壳适配 CSS）
- 设置页「外观 → 界面风格」网格、「扩展 → 插件」入口
- 插件管理对话框：本地安装、GitHub 列表、启用/禁用、卸载、主题应用
- 内置示例主题 `plugin-repo/plugins/theme-office-2003`
- 打包脚本 `scripts/pack-plugin.mjs`（生成 .moext + SHA-256 + 索引条目）
- **验收**：切到 2003 主题后外壳风格成立；同一份 DOCX 在任意主题下页数与打印结果一致

### P1.5 — 经典外壳变体（已实现）

`chrome.toolbar: 'classic-two-row'` 现在会真正改变 DOM 结构：`AppShell` 在标题栏下
额外渲染一行菜单栏（`MenuBar.tsx`），形成「标题栏 → 菜单栏 → 标签栏 → 工具栏」的三段式，
这是 Office 2003 风格能否立住的关键。

实现方式刻意保持轻量：**没有做完整的插槽化重构**，只加了一个按变体条件渲染的组件。

- `useChromeVariant()` 从插件注册表解析当前主题的 chrome 配置，主题被禁用后自动回退默认值。
- 菜单只放**应用级**命令（文件 / 编辑 / 视图 / 工具 / 帮助）。
  插入表格、字体、加粗这类文档级操作仍由各编辑器自己的工具栏承担，
  不在外壳重复实现——否则两套入口会各自漂移。
- `controlShape` / `density` / `titlebar` 仍走 `data-chrome-*` 属性 + CSS 适配，不触碰组件。

仍未实现：`ribbon` 变体（需要功能区布局，工作量远大于菜单栏，暂不排期）。

### P2 — T2 沙箱 + 文档钩子

- Worker + RPC 桥、`CapabilityToken`
- `documentHooks`（`onOpen` / `onWillSave` / `onExport`）
- `transaction` 写入 + 单步撤销
- 本地审计表：记录插件权限与文档写入事件
- 首个示范插件 `markdown-export`，证明导出类插件可行
- 插件文件存储切换为原生目录（支持热重载）
- **验收**：插件崩溃不影响宿主；撤销一次即回滚整个插件事务

### P3 — T3 受信能力（按需）

- Rust `office-plugin-host`：`plugin://` 协议、企业侧载验签
- 权限确认 UI、外发提示
- AI 类插件（需 `network`）

### 11.5 已实现清单

| 能力 | 位置 |
|---|---|
| 插件契约（清单 / 索引 / 主题 / 权限类型） | `packages/plugin-api/src/index.ts` |
| 清单校验与 semver 范围 | `packages/plugin-host/src/manifest.ts` |
| 主题令牌白名单 / 对比度 / chrome 应用 | `packages/plugin-host/src/theme.ts` |
| GitHub 索引拉取（ETag + raw 回退）与 SHA-256 校验 | `packages/plugin-host/src/github.ts` |
| `.moext` 解析与安装 | `packages/plugin-host/src/installer.ts` |
| IndexedDB 虚拟文件存储 | `packages/plugin-host/src/storage.ts` |
| 已安装 / 索引 / 当前主题的状态源 | `packages/plugin-host/src/registry.ts` |
| 插件管理界面 | `packages/editor-shell/src/PluginManager.tsx` |
| 经典菜单栏与外壳变体解析 | `packages/editor-shell/src/MenuBar.tsx`、`chrome-variant.ts` |
| 设置入口与外观风格网格 | `packages/editor-shell/src/dialogs.tsx` |
| 命令面板「管理插件」 | `packages/editor-shell/src/CommandPalette.tsx` |
| 外壳适配 CSS（beveled / compact / gradient / menubar） | `apps/desktop/src/styles.css` |
| 官方插件仓库脚手架（3 个主题 + 索引 + CI） | `plugin-repo/` |
| 打包脚本（单包） | `scripts/pack-plugin.mjs` |
| 索引生成脚本（整仓 + 可重现打包） | `plugin-repo/scripts/build-index.mjs` |
| 回归测试（27 项） | `tests/plugin-system.test.ts` |


## 12. 风险与取舍

| 风险 | 影响 | 缓解 |
|---|---|---|
| 主题破坏可用性（低对比、控件挤爆） | 高 | 令牌校验 + 对比度检查 + chrome 枚举化 |
| 主题影响文档分页与打印 | 高 | `--mo-doc-*` 与 UI 令牌强制分离，主题无权修改 |
| 插件 API 过早固化 | 中 | v0 标 `unstable`，仅对 T1 承诺稳定；T2/T3 随真实插件演进 |
| T2 RPC 性能 | 中 | 批量事务、快照按需、禁止逐字符调用 |
| 插件生态冷启动 | 中 | 先交付 3 个官方主题 + 1 个导出插件作为样板与文档示例 |
| 插槽抽象导致外壳复杂度上升 | 中 | 变体数量硬限制（每插槽 ≤3），新增变体需 ADR |

### 核心取舍

**不做"万能插件运行时"。** VS Code 式的 Extension Host 需要完整 Node 运行时，
会让安装包增加数十至上百 MB，直接违背"轻量、快速"的产品前提。

本方案用「**声明式优先 + 窄 API 沙箱**」换取体积与安全，代价是插件表达力受限：
插件不能任意改 UI、不能直接操作 DOM、不能随意访问文件系统。
对于办公套件的定制场景（主题、模板、导出、校对），这个表达力已经够用；
真正需要全能力的场景应走 T3 原生插件 + 签名，而不是放开沙箱。

## 13. 与客户反馈的对应关系

| 反馈要点 | 本方案回应 |
|---|---|
| 主体保持 mini | §2 内核边界判据：内核只保留"会毁文件"的部分 |
| 不同使用者需求不一样 | §3 三级模型 + §4 贡献点：T1 覆盖 80% 定制，客户可自建 |
| 参考 deepseek harness 的插件概念 | 方向一致：极简内核 + 声明式能力注入；差异是我们额外引入信任分级与主题结构层 |
| 主题插件（Office / WPS / XP 2003） | §5 双层主题：Token 层解决配色，Chrome 层解决"风格"；`--mo-doc-*` 隔离保证分页与打印不受影响 |
| 软件内有插件扩展入口 | §11.5：设置 → 扩展 → 插件；命令面板「管理插件」 |
| 用户可安装本地插件功能包 | `.moext` 包 + 原生文件选择，全静态校验后落地 |
| 官方插件打包上传 GitHub，通过 API 拉列表 | §10：`index.json` + `packages/*.moext`，ETag 拉取与 raw 回退，SHA-256 校验 |
| 功能插件（不只换肤） | §4 `commands`（宏）+ `templates`，见[插件宏参考](./plugin-macros.md) |
| 客户要求 MCP 支持 | §15：**做成 T3 插件，不进核心**；前提是内核先有文档命令总线 |

## 14. 待决问题

1. **`ribbon` 变体是否要做**：`classic-two-row`（菜单栏）已实现；功能区布局是另一套
   完全不同的控件组织方式，工作量远大于菜单栏，暂不排期。
2. **T2 UI 的 iframe 主题继承**：注入 `--mo-*` 令牌能否覆盖插件自带样式？是否需要强制插件使用宿主提供的组件库？
3. **插件文件存储是否切换到原生目录**：当前用 IndexedDB 虚拟文件系统（Web 与桌面统一、零 Rust 改动），
   代价是开发者无法直接丢文件热重载。P2 是否值得为此加 Rust 命令？
4. **主题是否应带走强调色**：当前强调色由用户设置决定，主题的 `defaultAccent` 不生效
   （`applyAccent` 的内联样式优先级更高）。忠实还原 2003 的蓝色需要跨 store 协调，
   需确认是否值得为"风格完整性"引入这层耦合。
5. **菜单栏是否要扩展到「插入」「格式」**：当前只放应用级命令，文档级操作仍在编辑器工具栏。
   若要补齐九个菜单，需要给编辑器定义一套跨进程命令接口，属于 T2 的范畴。

5. **是否允许插件替换整个编辑器**（如接入第三方表格内核）：当前设计**不允许**，
   因为会破坏 L0–L3 安全模型。若客户有此需求，应作为独立产品分支而非插件。


## 15. MCP 的定位（决策）

客户提出要 MCP 支持。结论：**做成 T3 插件，不进核心**。

### 15.1 为什么不是核心功能

按 §2.1 的内核边界判据：MCP 出错**不会**导致文件损坏、内容丢失或安全失守，
只会让某个外部工具调不通。判据给出的答案是插件。

三条工程理由：

1. **与产品承诺冲突**。核心承诺是"离线优先、轻量快速"。MCP 需要监听或出站网络、
   托管鉴权凭据、校验工具 schema、维持长驻进程。放进核心意味着每个用户都付这份代价，
   而真正需要它的只是少数客户。
2. **协议仍在演进**。传输层从 stdio 到 streamable HTTP 已经变过一次。
   放核心，内核就得跟着协议改；做成插件，升级插件即可。
3. **客户要的 MCP 各不相同**。各家服务器、鉴权方式、工具清单都不同。
   这正是插件该干的活；放核心等于让内核为个别客户做适配。

### 15.2 但内核必须先有"文档命令总线"

MCP 插件需要一个稳定的文档操作接口。所以两件事要分开：

| 归属 | 内容 |
|---|---|
| **内核** | 文档命令总线：RPC 面、权限门禁、本地审计表。这是 T2 的基础设施，MCP 只是它的第一个消费者 |
| **插件** | MCP 适配层：协议解析、工具清单映射、连接管理 |

有了命令总线，MCP server 大概只是一个几百行的适配器；没有它，每个 MCP 插件都要
自己重写一遍文档操作，且各自绕过权限。

### 15.3 方向选择

| 方向 | 含义 | 评价 |
|---|---|---|
| **A. Mini Office 作为 MCP server** | 对外暴露文档操作，让客户自家的 AI 助手或 Claude / Cursor 来读写文档 | **优先做**。不涉及凭据托管，风险低，能力全部来自命令总线 |
| B. Mini Office 作为 MCP client | 连接外部 MCP server 拉数据填进文档 | 更偏数据集成，需要凭据管理与出站白名单，放到 A 之后 |

### 15.4 落地顺序

```text
P1.5  T1 宏 + 模板（已完成）        零运行时，覆盖多数定制
  ↓
P2    文档命令总线 + T2 沙箱        功能插件与 MCP 的共同底座
  ↓
P3    MCP server 插件（T3）         基于命令总线，默认关闭，企业侧载或手动开启
```
