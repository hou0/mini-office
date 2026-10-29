# 插件宏参考（功能插件 v1）

> 状态：可用 ｜ 最后更新：2026-10-29 ｜ 作者：lq
>
> 关联：[插件架构](./plugin-architecture.md) §4 贡献点、[插件发布指南](./plugin-publishing.md)

## 1. 宏是什么，为什么不是脚本

客户要的"定制功能"绝大多数是**一串固定的文档操作**：把正文改成仿宋三号、把表格改成三线表、
给表头加底纹。这类需求不需要运行任意代码，只需要一张操作表。

所以 v1 的功能插件是**宏**：

- 插件在 `manifest.json` 里用 JSON 描述步骤；
- 宿主按白名单解释执行，插件代码一行都不跑；
- 因此它仍然属于 T1 声明式插件，安装不需要签名、不需要审核、不需要权限声明。

需要真正执行逻辑的功能（校对、OCR、格式转换、AI）属于 T2 沙箱插件，走 Worker + RPC，不在本文范围。

## 2. 一条命令长什么样

```json
{
  "id": "apply-gov-format",
  "title": "套用公文版式",
  "description": "正文仿宋三号、1.5 倍行距、首行缩进两字",
  "keywords": "公文 排版 格式",
  "shortcut": "Mod+Shift+G",
  "steps": [
    { "op": "word.page.margins", "top": 132, "bottom": 132, "left": 106, "right": 106 },
    { "op": "word.paragraph.format", "alignment": "justify", "lineHeight": 1.5, "firstLineIndentPx": 32 },
    { "op": "word.font.apply", "fontFamily": "仿宋_GB2312", "fontSizePt": 16 }
  ]
}
```

命令出现在命令面板的「插件 · <插件名>」分组里，只在**文档类型匹配**时可点。

## 3. 硬规则

| 规则 | 说明 |
|---|---|
| 先校验后执行 | 任何一步不合法，整条命令都不执行，不会留下改了一半的文档 |
| 单一文档类型 | 一条命令的所有步骤必须同属 `docx` 或 `xlsx`，混用会被拒绝安装 |
| 步骤数上限 | 32 步 |
| 不改入参 | 执行返回新文档对象，失败时原数据未被触碰 |
| 只碰正文段落 | 段落与字体操作不进入表格单元格，避免改动表内文字 |

## 4. `scope`：为什么默认不动标题和落款

`word.paragraph.format` 与 `word.font.apply` 都支持 `scope`：

| 取值 | 行为 |
|---|---|
| `body`（默认） | **跳过居中与右对齐段落** |
| `all` | 字面意义上的全部段落 |

这条默认值是被真实排版问题逼出来的。公文模板里标题居中、落款右对齐，
如果一视同仁地套正文格式：

- 标题会被拉成两端对齐，视觉上偏到一边；
- 落款会多出两字首行缩进，`年 月 日` 直接折行。

所以 `body` 跳过这两类段落，`all` 留给确实需要"全选后统一格式"的场景。

## 5. 操作白名单

### Word（`docx`）

| op | 参数 | 说明 |
|---|---|---|
| `word.page.margins` | `top` `right` `bottom` `left`（px，0–1440） | 缺省字段保持不变 |
| `word.page.orientation` | `orientation`: `portrait` \| `landscape` | 同时交换纸张宽高 |
| `word.paragraph.format` | `scope` `alignment` `lineHeight` `firstLineIndentPx` `spaceBeforePx` `spaceAfterPx` `indentLeftPx` `indentRightPx` | 至少给一个参数 |
| `word.font.apply` | `scope` `fontFamily` `fontSizePt` `bold` `italic` `color` | 作用于段落内所有 run |
| `word.table.borders` | `preset`: `three-line` \| `grid` \| `none` | 作用于文档内所有表格 |
| `word.defaultFont.set` | `fontFamily`（必填） `fontSizePt` | 改文档默认字体，不改已有 run |

`three-line` 的效果：首行上方 1.5px、首行下方 0.75px、末行下方 1.5px，其余无框线。

### Excel（`xlsx`）

| op | 参数 | 说明 |
|---|---|---|
| `sheet.freezeHeader` | `sheet`: `active` \| `all` | 冻结首行 |
| `sheet.headerStyle` | `sheet` `fontFamily` `fontSizePt` `bold` `fill` `color` | 作用于首行已用列，空单元格也会被填充 |

> 数值单位统一用 CSS px：96px ≈ 1 英寸。国标公文页边距（上 37mm、下 35mm、左 28mm、右 26mm）
> 约等于 `top: 140, bottom: 132, left: 106, right: 98`。

## 6. 模板

模板是**放在包里的真实 Office 文件**，用户点一下就以"未保存副本"打开，永远不会覆盖插件文件。

```json
{
  "id": "gov-letter",
  "label": "党政机关公文",
  "kind": "docx",
  "description": "含标题、主送机关、正文层级与落款位置",
  "file": "templates/gov-letter.docx"
}
```

约束：

- `file` 必须是包内相对路径，不能以 `/` 开头、不能含 `\` 或 `..`；
- 扩展名必须与 `kind` 一致；
- **文件必须真的在包里**，否则安装阶段就报"模板文件缺失"，不会等到用户点击才失败。

## 7. 完整示例

仓库里的 [`pack-gov-doc`](../plugin-repo/plugins/pack-gov-doc/manifest.json) 是可直接参考的样板：
6 条命令 + 1 个公文模板，全部声明式，包体 3.6 KB。

模板文件由 [`make-template.mjs`](../plugin-repo/scripts/make-template.mjs) 生成，
脚本固定了 zip 时间戳，重复执行得到完全相同的字节。

## 8. 还没做的

| 能力 | 状态 | 原因 |
|---|---|---|
| 按选区执行 | ⬜ | 选区状态在各编辑器内部，需要 T2 的命令总线统一暴露 |
| 条件判断、循环 | ⬜ | 一旦需要控制流就说明该上 T2 沙箱，而不是把宏做成脚本语言 |
| 文本片段（snippets） | ⬜ | 需要在 contenteditable 里拦截输入，侵入性高于收益，留到 T2 |
| 新增导出目标 | ⬜ | 需要跑代码，属 T2 |
