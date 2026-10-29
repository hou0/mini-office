# 插件发布指南（GitHub 仓库即插件市场）

> 状态：可用 ｜ 最后更新：2026-10-29 ｜ 作者：lq
>
> 关联：`docs/plugin-architecture.md` §10、ADR-0005

## 1. 总体流程

```text
写插件（plugins/<名字>/manifest.json）
      ↓
node scripts/build-index.mjs        打包到 packages/ + 算 SHA-256 + 生成 index.json
      ↓
git push origin main                产物随提交上线，raw 立即可下载
      ↓
客户端「刷新列表」→ 用户一键安装
```

打 tag（`git tag v1 && git push origin v1`）是可选动作，只用于生成 Release 镜像，
**不是安装链路的一环**。日常发插件就是一次普通 `push`。

不需要服务器、不需要审核后台、不需要数据库。唯一的日常动作是加一个目录、打一个 tag。

## 2. 仓库

本仓库已内置可用的脚手架：**`plugin-repo/`**。它是一份完整的、可直接推送到独立 GitHub 仓库的目录，
目标坐标 `lq0910/mini-office-plugins`。

```text
plugin-repo/
├── index.json                 # 客户端唯一入口，由脚本生成，不要手改
├── package.json
├── plugins/                   # 源码：一个目录一个插件
│   ├── theme-office-2003/manifest.json
│   ├── theme-wps/manifest.json
│   └── theme-office-365/manifest.json
├── packages/                  # 产物：打包出的 .moext，必须提交
│   ├── mini-office.theme-office-2003-1.0.0.moext
│   ├── mini-office.theme-wps-1.0.0.moext
│   └── mini-office.theme-office-365-1.0.0.moext
├── scripts/build-index.mjs
└── .github/workflows/release.yml
```

初始化：

```bash
cd plugin-repo
git init
git add .
git commit -m "init: Mini Office 官方插件仓库"
git remote add origin git@github.com:lq0910/mini-office-plugins.git
git push -u origin main
```

> `.github/workflows/release.yml` 只有位于仓库根目录才会被 GitHub 执行。
> 把 `plugin-repo/` 里的内容作为独立仓库根目录推送，工作流才会生效。

## 3. 写一个主题插件
`plugins/theme-wps/manifest.json`（节选）：

```json
{
  "id": "mini-office.theme-wps",
  "name": "WPS 风格外观",
  "version": "1.0.0",
  "publisher": "mini-office",
  "tier": "declarative",
  "engines": { "miniOffice": ">=0.1.5 <0.5.0", "api": "^1.0.0" },
  "permissions": [],
  "contributes": {
    "themes": [
      {
        "id": "wps",
        "label": "WPS 风格",
        "tokens": {
          "--mo-bg": "#F4F6F9",
          "--mo-surface": "#FFFFFF",
          "--mo-surface-2": "#EEF1F6",
          "--mo-text-primary": "#1A1A1A",
          "--mo-text-muted": "#5C6470",
          "--mo-line": "#DFE3EA",
          "--mo-radius-md": "4px",
          "--mo-radius-lg": "6px"
        },
        "chrome": { "controlShape": "rounded", "density": "default" }
      }
    ]
  }
}
```

字段定义见 `packages/plugin-api/src/index.ts`。要点：

- `permissions` 对声明式插件**必须为空数组**，否则校验失败。
- `tokens` 只能写 `--mo-*` 白名单内的令牌，写 `--mo-doc-*` 会被直接拒绝
  （会改变文档分页与打印结果）。
- 必填令牌：`--mo-bg`、`--mo-surface`、`--mo-surface-2`、`--mo-text-primary`、
  `--mo-text-muted`、`--mo-line`。缺一个就拒绝激活。
- 正文与背景对比度需 ≥ 4.5:1，次要文字 ≥ 3:1，否则拒绝激活并给出具体令牌对。
- `chrome` 只能从枚举里选：

  | 键 | 取值 | 效果 |
  |---|---|---|
  | `controlShape` | `flat` / `beveled` / `rounded` | `beveled` 为凸起/凹陷的经典边框 |
  | `density` | `compact` / `default` / `comfortable` | 控件高度与内边距 |
  | `titlebar` | `flat` / `gradient` | 标题栏渐变 |
  | `toolbar` | `modern` / `classic-two-row` / `ribbon` | `classic-two-row` 在标题栏下渲染经典菜单栏 |

## 3.1 写一个功能插件（宏 + 模板）

功能插件不写代码，用 JSON 描述操作步骤。完整操作清单与 `scope` 语义见
[插件宏参考](./plugin-macros.md)，这里只给最小骨架：

```json
{
  "id": "mini-office.pack-gov-doc",
  "name": "公文排版助手",
  "version": "1.0.0",
  "publisher": "mini-office",
  "tier": "declarative",
  "engines": { "miniOffice": ">=0.1.5 <0.5.0", "api": "^1.0.0" },
  "permissions": [],
  "contributes": {
    "commands": [
      {
        "id": "apply-gov-format",
        "title": "套用公文版式",
        "keywords": "公文 排版",
        "steps": [
          { "op": "word.page.margins", "top": 132, "bottom": 132, "left": 106, "right": 106 },
          { "op": "word.paragraph.format", "alignment": "justify", "lineHeight": 1.5, "firstLineIndentPx": 32 },
          { "op": "word.font.apply", "fontFamily": "仿宋_GB2312", "fontSizePt": 16 }
        ]
      }
    ],
    "templates": [
      { "id": "gov-letter", "label": "党政机关公文", "kind": "docx", "file": "templates/gov-letter.docx" }
    ]
  }
}
```

要点：

- `permissions` 仍然必须为空——宏由宿主解释执行，插件不跑代码。
- 一条命令的所有步骤必须同属 `docx` 或 `xlsx`，混用会被拒绝安装。
- 段落与字体操作的 `scope` 默认 `body`，会**跳过居中标题与右对齐落款**。
  需要"全选后统一格式"时显式写 `"scope": "all"`。
- 模板文件必须真的在包里，安装期就会校验，不会等到用户点击才失败。

`plugin-repo/plugins/pack-gov-doc/` 是可对照的完整样板。

## 4. 打包与生成索引

```bash
npm install
node scripts/build-index.mjs --owner lq0910 --repo mini-office-plugins
```

输出：

```text
packages/mini-office.theme-office-2003-1.0.0.moext   1013 B
packages/mini-office.theme-office-365-1.0.0.moext     942 B
packages/mini-office.theme-wps-1.0.0.moext            890 B
已写出 index.json，共 3 个插件
下载地址前缀：https://raw.githubusercontent.com/lq0910/mini-office-plugins/main/packages
```

**产物必须提交进仓库**（`packages/` 不在 `.gitignore` 里）。原因见下节。

**打包是可重现的**：脚本固定了 zip 条目时间戳，同一份源码重复打包得到完全相同的
字节与 SHA-256。这不是洁癖——如果每次打包哈希都变，索引里的校验值与仓库里的包
就会对不上，用户安装时会报"插件包校验失败"。

> 这里有一个只会在多文件插件上暴露的坑：JSZip 默认会给 `templates/` 这类父目录
> 自动补一个目录条目，而那个条目用的是**当前时间**，不是固定值。单文件主题插件
> 看不出来，带子目录的插件每次打包哈希都不一样。脚本因此显式传了
> `createFolders: false` 并对文件列表排序，`tests/plugin-system.test.ts` 里有一条
> 断言守着"所有条目时间戳为 1980 且无目录条目"。

## 4.1 为什么不用 GitHub Release 资产

实测结论：`github.com/.../releases/download/...` 会 302 到 `objects.githubusercontent.com`，
该域**不返回 `Access-Control-Allow-Origin`**，浏览器与 Tauri WebView 的 `fetch` 都会被
CORS 拦截，安装直接失败。

| 通道 | CORS | 可用 |
|---|---|---|
| `api.github.com` | `*` | 是 |
| `raw.githubusercontent.com` | `*` | 是 |
| `github.com/.../releases/download/...` | 无 | 否 |

所以索引里的 `download` 一律指向
`https://raw.githubusercontent.com/{owner}/{repo}/main/packages/{file}`。
插件包只有 1KB 量级，进仓库的代价可以忽略；Release 仍会创建，但只作为人工下载镜像。

raw 有几分钟 CDN 缓存，索引更新后可能短暂读到旧版本。客户端优先走 `api.github.com`
（无缓存），所以正常使用感知不到。


## 5. 发布

正常发布（推荐）：

```bash
node scripts/build-index.mjs
git add index.json packages
git commit -m "feat: 新增 xxx 插件"
git push origin main
```

推送完成后 raw 上即可下载，客户端「刷新列表」就能看到。**这是唯一的必需动作。**

可选：打 tag 生成 Release 镜像，方便人工下载存档。

```bash
git tag v1
git push origin v1
```

CI（`.github/workflows/release.yml`）会依次执行：

1. 打包全部插件到 `packages/`；
2. 把 `index.json` 与 `packages/` 提交回 `main`（无变化则跳过）；
3. 创建 Release，把 `packages/*.moext` 作为附件上传。

> 第 3 步只是镜像。客户端**不读** Release 资产——那条链路的 CORS 是断的（见 §4.1）。

> 已安装用户不会自动升级。当前是"装/卸"模型，检查更新与灰度留到 P2。

## 6. 客户端如何拉取

| 步骤 | 行为 |
|---|---|
| 1 | `GET api.github.com/repos/{owner}/{repo}/contents/index.json?ref={branch}`，带 ETag |
| 2 | 304 → 用缓存，不消耗配额 |
| 3 | 403/429（未认证 60 次/小时）或网络失败 → 回退 `raw.githubusercontent.com` |
| 4 | 仍失败 → 用本地缓存，界面标注"当前为缓存列表" |
| 5 | 安装时下载 `.moext`，比对索引里的 SHA-256，不一致直接中止 |

仓库坐标在 `packages/plugin-api/src/index.ts` 的 `DEFAULT_PLUGIN_REPOSITORY`：

```ts
export const DEFAULT_PLUGIN_REPOSITORY = {
  owner: 'lq0910',
  repo: 'mini-office-plugins',
  branch: 'main',
  indexPath: 'index.json',
};
```

换仓库、或企业内网镜像，不需要改代码：

```ts
usePluginRegistry.getState().setRepository({
  owner: 'my-org',
  repo: 'my-plugins',
  branch: 'main',
  indexPath: 'index.json',
});
```

该坐标会持久化到 `localStorage`，下次启动沿用。

## 7. `index.json` 规范

```json
{
  "schemaVersion": 1,
  "updatedAt": "2026-10-29T07:53:29.412Z",
  "plugins": [
    {
      "id": "mini-office.theme-office-2003",
      "name": "Office 2003 经典外观",
      "version": "1.0.0",
      "publisher": "mini-office",
      "tier": "declarative",
      "permissions": [],
      "download": "https://raw.githubusercontent.com/lq0910/mini-office-plugins/main/packages/mini-office.theme-office-2003-1.0.0.moext",
      "sha256": "63289ec7e4dec3bbfdea121b0c10201e75231384b2fbb09cc36055a460a8da8f",
      "size": 1013,
      "description": "…",
      "license": "MIT",
      "tags": ["theme", "office-2003"]
    }
  ]
}
```

约束：

- `schemaVersion` 必须是 `1`，否则客户端拒绝整份索引。
- `sha256` 必填且为 64 位小写十六进制。**缺少或格式不符的条目会被客户端直接丢弃**，
  这样即使索引被篡改，也无法塞入一个没有校验值的包。
- `download` 支持任意 HTTPS 地址，便于企业自建镜像；但必须是**直出**地址
  （响应带 `Access-Control-Allow-Origin`），不要用会 302 到无 CORS 域的中转链接。
- 未知字段被忽略，便于将来扩展而不破坏旧客户端。

## 8. 发布前自检

```bash
npx vitest run tests/plugin-system.test.ts
```

该测试会校验 `plugin-repo/plugins/` 下所有官方主题的清单与令牌，
并检查 `plugin-repo/index.json` 与主题目录是否一一对应。

另外人工确认：

- [ ] 浅色与深色下文字都清晰可读（深色缺省时回退浅色令牌）
- [ ] `.moext` 内 `manifest.json` 在根目录（允许单层目录包裹）
- [ ] `download` 用 HTTPS、能匿名访问，且**响应带 CORS 头**（raw 可以，Release 不行）
- [ ] `packages/` 里的包已提交，且与 `index.json` 的 `sha256` 一致
- [ ] `engines.miniOffice` 覆盖当前发布版本
- [ ] 没有手改 `index.json` 的 `sha256`

## 9. 客户定制插件怎么交付

客户或协作者不需要仓库：

1. 给他们 `scripts/pack-plugin.mjs`（在主仓库）和插件目录；
2. 他们打包出 `.moext`：
   ```bash
   node scripts/pack-plugin.mjs /path/to/my-theme
   ```
3. 在「设置 → 扩展 → 插件 → 安装本地插件」选择文件即可。

如需纳入官方列表，再走第 3–5 步提交。
