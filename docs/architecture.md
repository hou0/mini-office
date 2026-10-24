# 架构说明

> 状态：V1 实施完成 ｜ 最后更新：2026-10-18

## 1. 总览

Mini Office 采用 pnpm workspace + Cargo workspace 单仓库，按“UI（React）—文档模型（TypeScript）—
格式适配（TypeScript OOXML）—平台抽象—Rust 核心服务—Tauri 桌面壳”分层。

```text
┌────────────────────────────────────────────────────────────────────┐
│ apps/desktop (React + Vite)                                        │
│  └─ editor-shell：首页 / 标签 / 命令面板 / 设置 / 兼容性对话框     │
├────────────────────────────────────────────────────────────────────┤
│ packages/document-model                                            │
│  ├─ 文档模型：DocModel / WorkbookModel / SlideDeck                 │
│  ├─ 生命周期状态机：VALIDATING → LOADING → VIEW_READY / EDIT_READY │
│  └─ Zustand workspace store：sessions / tabs / settings / recent    │
├────────────────────────────────────────────────────────────────────┤
│ packages/editor-word / editor-sheet / editor-slides                │
│  └─ FormatHandler：open / serialize / createBlank / stats          │
│      ├─ ooxml：JSZip + DOMParser、关系解析、能力扫描               │
│      ├─ word：分页测量、contentEditable 编辑、DOCX 重写             │
│      ├─ sheet：虚拟网格、公式引擎、XLSX 重写                       │
│      └─ slides：母版/版式继承、静态渲染、演示                       │
├────────────────────────────────────────────────────────────────────┤
│ packages/platform-adapter                                          │
│  ├─ WebPlatformAdapter：File System Access / 下载 / IndexedDB       │
│  └─ TauriPlatformAdapter：dialog / invoke / 字体                    │
├────────────────────────────────────────────────────────────────────┤
│ apps/desktop/src-tauri（Tauri 命令）                                │
├────────────────────────────────────────────────────────────────────┤
│ crates/*（Rust：安全扫描 / OOXML / 存储 / 字体 / PDF / 平台）       │
└────────────────────────────────────────────────────────────────────┘
```

### 依赖方向

- UI 组件不直接调用 Tauri API；所有系统能力通过 `PlatformAdapter` 注入。
- 编辑器组件不实现文件写入；保存由 workspace store + `FormatHandler.serialize` + 平台适配器完成。
- TS 侧的 `ooxml` 只做读取、解析、扫描与媒体 URL 管理；写回由各编辑器的 serializer 负责。
- Rust 侧不依赖前端状态，提供无会话的纯函数/服务，便于 CLI、测试和未来移动端复用。

## 2. 工作区与文档生命周期

`packages/document-model/src/store.ts` 是唯一的工作区状态源：

```text
UNOPENED → VALIDATING → LOADING → VIEW_READY ─┐
                                    ├─ READ_ONLY
                                    └─ EDIT_READY → DIRTY → SAVING
                                                         ├─ SAVED
                                                         └─ SAVE_FAILED
VIEW_READY / EDIT_READY → EXPORTING → 回到编辑/预览
```

- `FormatHandler.open` 自行完成“检测格式 → 安全/能力扫描 → 解析”。
- 打开时若能力等级 ≤ L1，会话强制只读；L2 允许编辑但在首次保存时强制另存副本。
- 会话的源字节保存在内存映射中，序列化时用于保留未知 OOXML 部件。
- 自动恢复：Web 与 Tauri 的 WebView 均使用 IndexedDB 保存序列化后的完整文件快照（最多 12 个，
  总量 160 MB）；恢复后未保存的会话重新进入 DIRTY 状态，并要求另存。
- 最近文件由快照元数据派生，不依赖原生路径；Tauri 另有 SQLite 最近文件表供未来原生菜单/文件关联使用。

## 3. OOXML 处理

### 3.1 包读取与安全

`packages/ooxml/src/package.ts`：

- 校验 ZIP 魔数、旧版 OLE 格式与损坏。
- 限制条目数（6000）、单条目解压大小（250 MB）、总解压大小（600 MB），防止压缩炸弹。
- 未知部件不会被丢弃；`OoxmlPackage` 只读，媒体通过惰性 Blob URL 提供。

Rust `office-security` 提供同一套策略的服务端版本，并额外检查：

- 路径穿越（`..`、绝对路径、反斜杠）。
- 压缩比异常。
- `vbaProject.bin`、`_xmlsignatures/`、外部关系。
- 输出扩展名白名单。

### 3.2 DOCX

`packages/editor-word`：

- 使用 `DOMParser` 按文档顺序遍历 `w:body`，保留段落/表格的原始顺序。
- 解析 `styles.xml` 的 `docDefaults`、样式继承链、`numbering.xml`、主题色/字体、页眉页脚关系。
- 解析字段（`fldSimple` 与复杂字段的 `fldChar`）、图片（DrawingML/VML）、表格（`gridSpan`/`vMerge`、
  单元格底纹/边框/垂直对齐）、章节属性（纸张、方向、页边距、页眉页脚距离）。
- 分页：隐藏测量层渲染所有块，测量高度后按内容高度装箱；分页符与 `pageBreakBefore` 强制换页。
  跨页表格不拆行，超出时保持整表并由 UI 提示限制。
- 编辑：每个段落是独立的 `contentEditable` 容器；输入后从 DOM 序列化回 `DocRun[]`（含字体、
  字号、颜色、加粗等），保留原有字段与图片。Enter 拆分段落、Backspace 合并段落、⌘/Ctrl+Z 模型级撤销。
- 保存：复制源 ZIP 的所有未知部件，仅重写 `word/document.xml`；新增图片写入 `word/media` 并补充关系与
  内容类型；`sectPr` 从原文档克隆后只更新纸张/页边距，保留页眉页脚引用。空白新建文档会生成完整的
  最小 DOCX 包。

### 3.3 XLSX / CSV

`packages/editor-sheet`：

- 解析 `workbook.xml`、关系、`sharedStrings.xml`、`styles.xml`（字体/填充/边框/numFmt/对齐）、
  工作表 XML（单元格、共享公式、合并、冻结、自动筛选、列宽行高、隐藏行列、底纹）。
- 公式引擎（见 ADR-0003）在解析后用记忆化递归求值，检测循环引用，生成 `computed` 缓存。
- 网格仅渲染可视区域；行高/列宽使用前缀偏移数组，支持 10 万行级别的虚拟滚动。
- 合并单元格按占用网格计算起点与 span；冻结窗格通过独立覆盖层实现。
- 保存：重建 `workbook.xml`、`workbook.xml.rels`、`styles.xml` 与所有工作表 XML；字符串使用
  `inlineStr`，避免共享字符串索引维护；删除 `calcChain.xml` 并设置 `fullCalcOnLoad="1"`，
  让 Excel/LibreOffice 重新计算；保留主题、媒体和其他未知部件。
- CSV：RFC 4180 子集解析（引号、逗号、分号、Tab 探测），保存时输出 UTF-8 BOM + CRLF，公式保存为
  计算后的显示值。

### 3.4 PPTX

`packages/editor-slides`：

- 解析演示文稿尺寸、主题色/字体、幻灯片顺序、版式与母版关系。
- 解析形状、文本、图片、表格、组合、背景；对占位符应用版式/母版的坐标与文本样式继承。
- 将布局/母版中的非占位符装饰元素渲染在幻灯片底层。
- 高保真渲染：EMU → CSS px，绝对定位；形状使用 CSS border-radius / clip-path；图片支持裁剪；
  表格支持合并与单元格填充；文本支持段落/run 级样式、项目符号、行距与垂直对齐。
- 全屏演示：独立覆盖层、键盘/点击换页、备注（N 键）、Esc 退出；浏览器使用 Fullscreen API，
  Tauri 使用原生 `setFullscreen(true)`，退出演示时同步恢复窗口。
- PPTX 基础编辑：`FormatHandler.serialize` 保留原始压缩包并按页面/对象身份增量回写，复杂未支持部件原样保留；首次保存按 L2 另存副本。

## 4. 保存与导出

### 4.1 保存策略

```text
store.saveSession
  → FormatHandler.serialize（保留未知部件）
  → PlatformAdapter.saveBytes
      Web：File System Access / 下载
      Tauri：write_file → office-storage::atomic_write
  → 原子写入（临时文件 + fsync + rename）
  → 原件时间戳备份
  → 容器校验（OOXML 保存前）
  → 更新快照 + 最近文件
```

- L2 文档首次保存强制弹出“另存为”，避免覆盖原文件。
- 保存失败保留 DIRTY/SAVE_FAILED 状态与错误消息；自动恢复快照仍可恢复。
- 前端保存前不会删除任何未知部件；DOCX 重写只触及主文档，XLSX 重写已知结构。

### 4.2 PDF / 打印

- Word：按分页 DOM 页面逐页 `html2canvas` 渲染，再写入 jsPDF；支持页码范围、进度与取消。
- Excel：把当前工作表数据区域构建为 A4 横向分页表格（重复表头、保留字体/填充/对齐/边框），
  逐页渲染为 PDF；打印使用同一套分页 DOM 与打印样式。
- PPT：导出层在屏幕外以原始尺寸渲染所有幻灯片，逐页生成 PDF；图片 ZIP 输出每页 JPEG。
- 打印：Word 复用分页 DOM；Excel 构建 A4 横向分页 DOM；PPT 渲染打印层。桌面端通过 Tauri
  `print_window` 调用原生打印面板，Web 端回退到 `window.print()`。
- Rust `office-pdf` 提供 LibreOffice Headless 发现与转换，仅作为 PoC 对照后端。

## 5. 外观与交互层

- 主题：浅色/深色/跟随系统；五套内置强调色与自定义强调色通过 CSS 自定义属性落盘，
  自动选择可读的前景 ink 色，并有对比度测试。
- 毛玻璃：macOS 使用透明窗口 + `Window.setEffects` 原生材质；不可用时回退为不透明表面。
- 浮层：Dialog、颜色面板、菜单和 Tooltip 使用 Portal 渲染，避免工具栏 `overflow-x` 裁剪；
  Dialog 固定内容区滚动、长文本换行/截断和键盘焦点管理。
- 桌面窗口：标题栏与标签栏空白区通过 `startDragging()` 拖动，双击最大化；按钮/标签/输入框不误触。

## 6. 性能

- 三大编辑器按需加载；Vite 将 PDF 相关依赖拆分为独立 chunk。
- Word 使用隐藏测量层 + 视觉页缓存；输入时不做分页，失焦/结构变化后重排。
- Excel 只渲染可视单元格；公式按单元格记忆化求值，仅公式单元格参与递归。
- PPT 缩略图复用同一渲染组件；导出时才创建全尺寸屏幕外页面。
- 自动恢复序列化带防抖（默认 1.4 s）和并发保护。

## 7. 安全与隐私

- 文档内容不上传；Web 快照保存在 IndexedDB，Tauri 快照/数据库保存在应用数据目录。
- 只读策略优先于“尽力保存”：宏、数字签名、旧版格式、无法保留的对象均不会静默丢弃。
- 外部链接不会自动加载；图片通过 Blob URL 本地读取。
- Tauri 能力声明按需启用 dialog、opener、window（拖动/最大化/全屏/effects）权限；
  文件读写通过自定义命令完成路径校验与安全写入。

## 8. 测试与质量

- TS：Vitest + jsdom；6 个测试文件 / 36 项；真实 DOCX/XLSX/PPTX 样例回归、公式对照、
  主题强调色对比度、序列化二次往返。
- Rust：能力扫描、路径穿越/压缩炸弹、原子保存、SQLite、字体枚举。
- UI：使用 Playwright + 系统 Chrome 做过手动自动化验证（首页、三编辑器、查找、格式化、
  撤销、保存、PDF/图片导出、自动恢复、深色主题）。
- 视觉回归、10 万行性能基准、三平台字体矩阵尚未纳入 CI，见 `docs/milestones.md`。

## 9. 技术债与后续

1. 将 Tauri 文件 IPC 从 JSON 字节数组改为原始响应/分块流，支持数百 MB 文件。
2. 为 XLSX 增加共享字符串写回、表格/图表/条件格式渲染与更完整的样式保留。
3. DOCX 增加列表编辑、页眉页脚编辑、批注/修订只读面板与更精细的表格跨页。
4. PPTX 增加文本/图片编辑、切页动画、视频播放与备注编辑。
5. Rust `office-ooxml` 的对等能力扫描与 TS 侧统一为同一份测试夹具。
6. PDF 排版从 canvas 截图升级为矢量文本绘制，减少模糊与体积。
7. CI：三平台构建、截图视觉回归、性能基准、安装包签名与自动更新。
