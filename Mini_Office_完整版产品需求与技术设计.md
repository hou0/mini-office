# Mini Office 产品需求与技术落地方案

> 版本：V1.0 ｜ 日期：2026-10-18 ｜ 定位：轻量、离线优先的跨平台办公套件

## 01 项目概述

产品名称：Mini Office（暂定）。产品定位：离线优先、轻量、快速、跨平台的日常办公套件。统一支持 Word、Excel、PPT 三类基本办公场景，以高频功能和高质量文件兼容为核心。

目标用户：日常合同编辑、表格处理、幻灯片阅读与演示的个人、团队及中小企业。产品口号：只保留真正需要的 Office 功能。

核心原则：简单、快速、可靠、中文友好、可扩展。尤其强调：UI 功能精简不代表原始文件可被破坏；无法保证无损编辑时应只读或另存副本。

## 02 平台与技术选型

桌面首发：macOS（Apple Silicon / Intel）、Windows x64、Linux x64。后续预留 Android、iOS 和 HarmonyOS NEXT 扩展。鸿蒙支持需独立评估 ArkUI/Web 容器及 Rust 原生桥接，不能假设 Tauri 桌面代码直接打包运行。

推荐组合：Rust + Tauri 2 + React + TypeScript + Tailwind CSS + Zustand + SQLite。Tauri 负责桌面窗口、系统调用与打包；Rust 负责文件管理、格式接口、任务调度、字体发现、PDF 调用与缓存；React 负责 UI 和编辑器壳。

编辑器：Word 优先评估 Tiptap/ProseMirror，Excel 优先评估 Univer 等现有组件，PPT 优先采用成熟的高保真预览或渲染内核；最终选型以真实 OOXML 样例的兼容性 PoC 为准。禁止在未验证兼容性之前将浏览器 DOM 直接作为 Office 文件真实模型。

注意许可：ONLYOFFICE、Univer 等项目的开源与商业模块许可各异。嵌入、修改、闭源分发、在线转换等使用方式在产品化之前必须单独审核。

## 03 Word：合同与常用文档

V1 必须：DOCX 打开、查看、受支持内容编辑与保存；文字样式（字体、字号、颜色、加粗、下划线）、段落（对齐、缩进、行距）、页面（A4、方向、页边距、分页）、基础表格、图片、页码、页眉页脚、查找替换、打印和 PDF 导出。

合同专项：首行缩进、固定行距、中文标点禁则、分页与跨页表格、签字落款区域、盖章图片位置；提供常见合同模板，打印与 PDF 版面尽量一致。

不在 V1：宏、邮件合并、实时多人协作、复杂批注与修订体系、复杂域和控件。遇到已有修订、数字签名、复杂控件等内容，显示兼容性警示，限制直接覆盖保存。

## 04 Excel：日常表格与公式

V1 必须：XLSX/CSV 打开与保存、单元格编辑、撤销重做、复制粘贴、自动填充、行列增删和宽高调整、单元格格式、合并、冻结窗格、工作表管理、排序、筛选、查找替换、打印区域与 PDF 导出。

首批公式：SUM、AVERAGE、COUNT、MAX、MIN、ROUND、SUMIF、COUNTIF、IF、AND、OR、IFERROR、IFS、VLOOKUP、XLOOKUP、INDEX、MATCH、LEFT、RIGHT、MID、LEN、CONCAT、TRIM、TODAY、NOW、DATE、YEAR、MONTH、DAY；最终目标 50–80 个常用函数。

公式实现不应只看函数名：必须测试相对/绝对引用、跨表引用、循环引用、错误值、日期序列及增量重算。采用虚拟滚动，仅渲染可视单元格。

不在 V1：VBA、Power Query、复杂透视表、外部数据库连接。对于 XLSM 文件默认不执行宏，并避免承诺无损保存。

## 05 PPT：静态高保真预览与演示

V1 必须：PPTX 打开、缩略图列表、16:9/4:3 适配、静态页面高保真显示、全屏演示、上一页/下一页、键鼠控制、备注查看、PDF 与图片导出。

优先展示正确的文字、形状、图片、表格、主题和母版；复杂动画、视频与转场可后置，但需明确展示不支持提示。后续版本再做文本修改、元素调整、多屏演讲者视图及更多动画。

## 06 UI/UX 设计规范

设计方向：macOS 原生应用轻盈感 + Office 高频操作效率 + 浏览器式多文件标签。默认取消高占屏的传统 Ribbon，采用紧凑工具栏、选区浮动工具栏、上下文属性栏与命令搜索。

结构：顶部窗口标题及标签页；其下高频工具栏；中间最大化内容画布；可收起右侧属性区；底部状态栏（页码、字数/表格状态、缩放、保存状态）。Word/Excel/PPT 三类文件可以混合打开并快速切换。

视觉建议：浅色背景 #F6F8FC，内容白色 #FFFFFF，默认强调色为雾青 #2E7D72（另提供远山蓝 / 紫藤 / 松针 / 陶土与自定义取色），深色背景 #1C2028；8–12px 圆角，120–180ms 轻动效。macOS 桌面端支持原生毛玻璃开关。默认跟随系统主题。中文 UI 字体使用系统可用字体，注意各平台字号和可读性。

效率功能：Ctrl/⌘+K 命令面板、Ctrl/⌘+S 保存、Ctrl/⌘+P 打印、全局最近文件、拖拽打开、可恢复标签页、可隐藏工具栏。

## 07 兼容性与安全保存架构

采用双模式文档能力：①保真查看模式（重点保证原貌）；②基础编辑模式（仅对已验证能力开放）。打开文件时分析文档特性，标记字体、特殊公式、动画、嵌入对象及其他不支持的 OOXML 元素。

推荐流程：原始 Office 文件 → 格式解析及能力检测 → 内部文档模型/保真查看模型 → 基础编辑或只读预览 → 兼容性检查 → 安全保存或另存副本 → PDF 输出。

保存原则：尽可能保留未修改的 OOXML 包部件与关系；未知对象不能静默删除；保存至临时文件并验证后原子替换；原件备份和自动恢复快照；签名与宏等特殊内容采用受限操作策略。

注意：保留未知 OOXML 部件是工程能力而非通用解析库的天然属性；需要对关系 ID、内容类型、元数据、数字签名等专项测试。

## 08 PDF 与中文字体

三个模块共用统一 PDF 导出服务，具备进度、取消、异常报告、页码/范围等配置；尽量复用预览布局，避免“屏幕正确、打印错位”。

优先支持宋体、仿宋、黑体、楷体、微软雅黑、等线、思源宋体与思源黑体等常见中文字体；通过系统枚举确认安装情况。字体缺失时提供替换提示，支持合法的字体嵌入或子集化，重点测试跨平台分页。

可将 LibreOffice Headless 作为概念验证阶段的 PDF 转换对照后端，但正式产品需考虑安装包体积、启动性能、许可、打包复杂度与格式差异。

## 09 性能与质量目标

以下仅是验收优化目标，非已经实测的性能承诺：冷启动≤1.5 秒、热启动≤0.5 秒；常见 DOCX/XLSX ≤2 秒可用；20 页 PPTX ≤3 秒出现可用预览；普通编辑交互反馈尽量低于 50ms；Excel 滚动目标 60FPS。

空白首页内存尽量≤150MB，安装包期望 50–100MB（是否可行取决于最终排版引擎）；必须对 Windows/macOS/Linux 分别测量，不使用单平台数据替代。

性能策略：按需加载三大编辑模块、Rust 后台解析、表格虚拟渲染、分阶段加载、公式增量重算、图片缓存上限、异步 PDF 导出与失败隔离。

## 10 建议工程目录

apps/desktop（Tauri）；apps/mobile（预留）；apps/harmony（预留）；packages/ui、design-system、editor-shell、editor-word、editor-sheet、editor-slides、document-model、shared；crates/office-core、office-ooxml、office-pdf、office-font、office-storage、office-formula、office-platform；tests/compatibility、visual-regression、performance。

所有 Web UI 通过 PlatformAdapter 接入文件选择、系统打印、剪贴板、字体及存储。不要在编辑器组件中直接调用 Tauri API，为未来移动端及鸿蒙适配保留空间。

## 11 分阶段实施计划

阶段 1（第 1–3 周）：工程搭建与兼容 PoC。选取真实合同、工作簿、演示文稿，验证打开、保存、PDF、字体与许可可行性；产出明确的引擎决策。

阶段 2（第 4–6 周）：统一桌面 UI、首页、标签页、文件管理、只读预览、快捷键、主题、自动恢复。

阶段 3（第 7–12 周）：Word 合同编辑、表格、页面、中文布局、DOCX 保存、PDF 导出与回归测试。

阶段 4（第 13–18 周）：Excel 基础编辑、公式、筛选排序、虚拟滚动、XLSX/CSV/PDF 处理。

阶段 5（第 19–24 周）：PPT 高保真静态预览、播放、PDF 导出、桌面三平台打包、性能和稳定性优化。

24 周是假定有经验小团队、成熟引擎可复用的粗略排期；单人开发或自研高保真排版引擎会明显延长周期。

## 12 验收测试与发布门槛

测试样本：20 页中文合同（跨页表格、页眉、签章）、含复杂图片定位 DOCX、千行含公式 XLSX、多工作表引用 XLSX、20 页商务 PPTX 和含动画 PPTX。

按“可正确查看 / 可安全编辑保存 / 只读 / 暂不支持”四档评估每个特性，并设立截图视觉回归、文件结构回归、公式对照计算、PDF 版面比对、崩溃恢复与性能基准。

最低发布原则：不可静默删除不支持对象、不能无提示改变重要合同版面、常用公式结果与预期一致、保存失败必须可恢复、演示模式不会因普通文件崩溃。

## 13 风险与决策建议

最大风险：OOXML 高保真编辑与保存复杂度、三平台字体和打印差异、第三方内核许可、鸿蒙生态适配，以及轻量包体与成熟完整渲染引擎之间的权衡。

推荐立项决策：先用 2–3 周完成引擎 PoC，再冻结 V1 需求。优先实现三格式可靠预览，随后 Word 合同编辑、Excel 日常计算，最后完善 PPT 播放。不要为了覆盖全部 Office 功能牺牲速度与文档安全。

---

**版本说明：** 本文为实施规划初稿，具体依赖选型、性能和排期均须经 PoC 验证。

## 13 可落地的完整项目目录与模块职责

采用 pnpm workspace + Cargo workspace 单仓库。以下是建议初始结构，不要求第一天创建所有实现文件；先固化边界接口，随后逐步填充。

```text
mini-office/
├── apps/
│   ├── desktop/                   # 主桌面应用（Tauri 2）
│   │   ├── src/                   # React 页面与前端入口
│   │   │   ├── app/               # Provider、路由、启动流程
│   │   │   ├── pages/             # 首页、编辑器、设置
│   │   │   └── commands/          # 前端命令调度
│   │   ├── src-tauri/
│   │   │   ├── src/main.rs
│   │   │   ├── src/lib.rs
│   │   │   ├── src/commands/      # Tauri 命令处理器
│   │   │   └── capabilities/      # 最小权限声明
│   │   └── tauri.conf.json
│   ├── mobile/                    # Android/iOS 预留入口
│   └── harmony/                   # HarmonyOS NEXT 单独适配
├── packages/
│   ├── design-system/             # Token、图标、字体、主题
│   ├── ui/                        # Button、Dialog、Tabs、Menu
│   ├── editor-shell/              # 标签、状态栏、命令面板
│   ├── editor-word/               # 文档编辑 UI 适配层
│   ├── editor-sheet/              # 单元格交互和表格虚拟化
│   ├── editor-slides/             # 缩略图、画布、放映
│   ├── document-model/            # 文档状态和编辑命令类型
│   ├── platform-adapter/          # 平台接口 + 桌面实现
│   └── shared/                    # 通用类型和工具
├── crates/
│   ├── office-core/               # 核心类型、错误、能力标记
│   ├── office-ooxml/              # OOXML 解析与写回适配
│   ├── office-doc/                # DOCX 模型和分页适配
│   ├── office-sheet/              # 工作簿格式与计算适配
│   ├── office-slides/             # 演示文稿适配
│   ├── office-pdf/                # PDF 导出调度
│   ├── office-font/               # 字体枚举、替代与嵌入
│   ├── office-storage/            # 原子保存、快照、恢复
│   ├── office-security/           # OOXML 安全扫描与限制
│   └── office-platform/           # Rust 系统能力抽象
├── tests/
│   ├── fixtures/
│   │   ├── docx/                  # 合同、图片、表格、修订
│   │   ├── xlsx/                  # 公式、日期、跨表、样式
│   │   └── pptx/                  # 母版、字体、动画、备注
│   ├── compatibility/             # 打开-编辑-保存-重开测试
│   ├── visual-regression/         # 跨平台截图比较
│   ├── performance/               # 启动、内存、导出性能
│   └── security/                  # 恶意 zip、超大嵌入内容
├── docs/
│   ├── prd.md                     # 产品需求
│   ├── architecture.md            # 架构决策
│   ├── ui-design.md               # UI 规范
│   ├── compatibility-matrix.md    # 格式兼容矩阵
│   ├── milestones.md              # 里程碑
│   └── adr/                       # 架构决策记录
├── .github/workflows/            # CI / 发行工作流
├── Cargo.toml                     # Rust workspace
├── pnpm-workspace.yaml
├── package.json
└── README.md
```

### 13.1 模块职责和依赖约束

| 模块 | 输入 | 输出 | 不应该负责 |
|---|---|---|---|
| editor-shell | 最近文件、标签、命令 | 编辑器装载事件 | OOXML 解析 |
| editor-word | 结构化编辑模型 | 编辑事务 | 操作系统文件写入 |
| editor-sheet | 工作簿视图、公式结果 | 单元格变更 | 自行实现存储事务 |
| editor-slides | 演示布局、缩略图 | 放映指令 | 擅自重写未知 PPTX 内容 |
| office-ooxml | Office 压缩包 | 已解析部件、兼容报告 | 界面绘制 |
| office-storage | 临时文件、校验数据 | 安全提交与恢复记录 | 控制编辑器交互 |
| office-pdf | 渲染模型和导出配置 | PDF 文件 / 错误 | 修改原始文件 |

统一调用链：UI → editor-shell → editor adapter → Rust command → office-core → 具体格式/存储/打印服务。未来平台接入通过 PlatformAdapter 完成，禁止在业务组件中散落直接的 Tauri API 调用。

## 14 UI 页面原型与界面规格

以下三幅为**概念性 UI 原型图**，不是已完成的产品截图。Word/Excel/PPT 使用同一导航系统、标签栏、工具栏和状态栏，仅内容画布随格式改变。原型重点传达布局和视觉层级。

### 14.1 Word 文档编辑原型

![Word 编辑器布局原型](ui_word.png)

布局：顶栏 48px、标签页 51px、操作栏 60px、中央 A4 比例页面、底部精简状态栏。文字选中时出现就近浮动工具栏；默认不显示右侧属性栏。真正开发需按页面实际 DPI 和 Zoom 缩放，不可用固定像素高度近似排版尺寸。

### 14.2 Excel 表格编辑原型

![Excel 表格布局原型](ui_excel.png)

布局：表格模式在操作栏下方增加名称框与公式栏，列标题和行号固定；虚拟化滚动区域只绘制可视单元格，选区、冻结窗格和复制边框应跟随滚动正确叠加。表格底部提供工作表标签。

### 14.3 PPT 阅读与演示原型

![PPT 阅读及演示原型](ui_ppt.png)

布局：左侧缩略图列表，右侧大幅 16:9/4:3 画布；F5 或演示按钮进入全屏，Esc 退出，左右方向键换页。V1 重点确保静态内容高保真；动画缺失时在详情提示。

### 14.4 设计 Token

| Token | 浅色 | 深色 | 用途 |
|---|---|---|---|
| background | #F6F8FC | #1C2028 | 应用底色 |
| surface | #FFFFFF | #303641 | 卡片、面板 |
| accent（默认雾青） | #2E7D72 | #6CC9BB | 选中、重点操作；可切换预设或自定义 |
| accent-soft | #E6F2EF | #1B3A36 | 选中背景、标签 |
| accent-ink | #FFFFFF | #10221F | 强调色上的文字/图标 |
| text-primary | #1D2939 | #F2F4F7 | 正文 |
| text-muted | #667085 | #98A2B3 | 次要文本 |
| radius-md | 8px | 8px | 常规控件 |
| radius-lg | 12px | 12px | 弹窗、面板 |
| motion-fast | 120ms | 120ms | 状态反馈 |

交互约束：高频按钮在默认可见工具栏；低频功能转移到“更多”菜单；右键菜单按元素类型变化。保存状态与当前编辑模式必须随时可见。支持键盘导航、辅助功能名称和高对比度主题。

### 14.5 页面与组件拆解

| 页面/组件 | 必须显示 | 主要交互 |
|---|---|---|
| Home | 最近文件、创建、打开、模板 | 拖拽打开、搜索最近文件 |
| Tabs | 混合 Office 标签、未保存点 | 关闭、重排、恢复会话 |
| Toolbar | 文档上下文操作 | 展开更多、快捷键 |
| WordCanvas | A4 页面、页边界与分页 | 输入、选区、表格、缩放 |
| SheetCanvas | 网格、公式栏、工作表 | 编辑、选择、排序、筛选 |
| SlidesCanvas | 缩略图、演示画布 | 选页、放映、备注 |
| ExportDialog | 文件名、路径、范围、进度 | 取消、导出、错误详情 |
| CompatibilityDialog | 风险特性、字体提示 | 只读、编辑副本、取消 |
| Settings | 主题、语言、快捷键、恢复 | 保存设置、重置 |

## 15 接口契约草案

### 15.1 前端到 Rust 的命令

```ts
type OfficeKind = 'docx' | 'xlsx' | 'pptx';
type OpenResult = {
  sessionId: string;
  kind: OfficeKind;
  displayName: string;
  readonly: boolean;
  warnings: Array<{ code: string; message: string; severity: 'info' | 'warning' | 'blocking' }>;
};
interface DesktopOfficeAPI {
  openFile(path: string): Promise<OpenResult>;
  saveSession(sessionId: string): Promise<void>;
  saveSessionAs(sessionId: string, path: string): Promise<void>;
  exportPdf(sessionId: string, outputPath: string, pageRange?: string): Promise<string>;
  closeSession(sessionId: string): Promise<void>;
}
```

路径权限应通过用户明确的文件选择行为授予，不允许网页内容直接传任意路径触发写入。大文件解析和转换采用异步任务，使用 taskId 汇报 progress/cancel/error，不阻塞主线程。

### 15.2 文档生命周期状态机

```text
UNOPENED → VALIDATING → LOADING → VIEW_READY
                                    ├─ READ_ONLY
                                    └─ EDIT_READY → DIRTY → SAVING
                                                         ├─ SAVED
                                                         └─ SAVE_FAILED
VIEW_READY / EDIT_READY → EXPORTING → VIEW_READY / EDIT_READY
```

保存需先写临时文件、校验容器、核实目标路径，再以安全方式提交；若原文档含无法保留的复杂部件，应拒绝直接覆盖，并提供副本导出。

## 16 兼容性等级和验收计划

兼容能力分为 L0 无法打开、L1 可预览、L2 可编辑但需另存、L3 可验证安全往返保存。不要把“能渲染预览”当成“能够往返无损保存”。

| 测试场景 | 初版验收关注点 | 不通过处理 |
|---|---|---|
| 20 页中文合同 | 分页、页码、跨页表格 | 提醒只读/调整后端 |
| 印章图片合同 | 位置与环绕稳定 | 禁止覆盖原件 |
| 多工作表 XLSX | 公式、格式、跨表引用 | 兼容模式 |
| 大型表格（10 万行） | 页面不死锁，虚拟滚动 | 延迟载入/内存限制 |
| 复杂 PPTX | 母版、文字、图片 | 展示降级提示 |
| 含宏工作簿 | 不执行宏，不破坏原始宏部件 | 强制只读或副本 |
| 缺失中文字体 | 发现并说明替代字体 | 显示警告 |
| PDF 输出 | 页数、尺寸、断行、图片 | 禁止静默成功 |

兼容回归采用测试基线文件、视觉快照、结构比对、重新打开验证四层测试。建立 Windows / macOS / Linux 的代表性字体安装矩阵。

## 17 研发排期与交付验收

| 阶段 | 时间 | 核心交付件 | 退出标准 |
|---|---|---|---|
| P0 引擎 PoC | 1–3 周 | 兼容报告与技术 ADR | 样例打开、导出、许可路径可行 |
| P1 应用壳 | 4–6 周 | 三平台窗口、标签、首页 | 启动、打开、恢复可靠 |
| P2 Word | 7–12 周 | 合同编辑和 PDF | 目标合同样本通过 |
| P3 Excel | 13–18 周 | 表格与公式 | 关键公式样本一致 |
| P4 PPT | 19–24 周 | 全屏演示与三平台安装包 | 静态演示稳定，发布检查通过 |

以上为有经验小团队、并复用成熟编辑/转换基础设施的计划假设。若核心排版引擎自研，需另外估算；个人开发时不应把此计划当作承诺。

### 17.1 首批开发任务（可直接建 Issue）

- [ ] INIT-001 单仓库与 Rust/Node 工作区
- [ ] INIT-002 Tauri 三系统构建与 CI
- [ ] ARCH-001 平台调用抽象与权限模型
- [ ] FORMAT-001 OOXML 文件能力扫描器
- [ ] FORMAT-002 真实兼容测试样例集
- [ ] UI-001 Design Token / Light / Dark
- [ ] UI-002 首页、最近文件、标签页
- [ ] UI-003 命令面板与快捷键映射
- [ ] DOC-001 Word 查看、分页基线
- [ ] DOC-002 合同文字编辑与表格
- [ ] DOC-003 DOCX 保存与副本策略
- [ ] SHEET-001 单元格虚拟滚动与选择
- [ ] SHEET-002 公式引擎验证集
- [ ] SLIDES-001 PPTX 预览和缩略图
- [ ] SLIDES-002 全屏演示控制
- [ ] EXPORT-001 统一 PDF 后端
- [ ] STORAGE-001 原子写入及自动恢复
- [ ] QA-001 跨平台格式与视觉回归
- [ ] RELEASE-001 签名、安装、自动更新

## 18 风险、非目标与决策门槛

| 风险 | 可能影响 | 处理策略 |
|---|---|---|
| Office 高保真复杂 | 版式错乱 | 先做引擎 PoC 和 L1–L3 等级 |
| Web 编辑模型丢失对象 | 文件内容损坏 | 能力扫描、只读、另存、保留未知结构 |
| 中文字体差异 | 合同分页改变 | 字体替代警告和跨平台快照 |
| 引擎打包体积过大 | 轻量目标落空 | 统一测量；必要时采用可选组件 |
| 开源许可不兼容 | 产品化受限 | 在引擎选型前明确审核 |
| HarmonyOS 适配不确定 | 无法如期上线 | 单独 PoC、平台接口隔离 |

首版不开发云端协作、VBA、完整 PPT 动画、复杂数据透视或 AI 助手。只有当 Word 合同、Excel 日常计算、PPT 稳定播放三个目标的兼容测试达到可接受水平后，再考虑扩张产品范围。

## 19 参考资料与选型复核

- Tauri 2 官方文档：https://v2.tauri.app/
- ProseMirror：https://prosemirror.net/
- Tiptap：https://tiptap.dev/
- Univer：https://univer.ai/
- ONLYOFFICE Desktop Editors：https://github.com/ONLYOFFICE/DesktopEditors
- LibreOffice：https://www.libreoffice.org/
- OOXML 标准参考：ISO/IEC 29500（以合法获取的标准文本为准）

注：上述为候选技术入口，不代表已完成深度性能/许可证核验。落地之前必须进行真实文档测试和分发许可证审核。
