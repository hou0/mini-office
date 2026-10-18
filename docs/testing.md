# 测试与验证说明

## 1. 自动化测试

### TypeScript / Vitest

```bash
pnpm test
```

环境：Vitest 2 + jsdom。当前基线：**6 个测试文件 / 36 项通过**。覆盖：

| 文件 | 内容 |
|---|---|
| `tests/formula.test.ts` | 算术优先级、引用/区域、逻辑函数、查找函数、文本函数、日期、SUMIF/COUNTIF、跨表、循环引用、错误传播、共享公式平移、数字/日期格式 |
| `tests/xlsx.test.ts` | 样式/合并/冻结/筛选/跨表公式的序列化—再解析往返；二次往返 |
| `tests/office-fixtures.test.ts` | 真实 DOCX 合同、XLSX 工作簿的解析与保存；页眉页脚/图片/表格；工作簿样式与公式缓存 |
| `tests/pptx.test.ts` | 真实 PPTX 的主题、形状、图片、表格、备注与标题推导 |
| `tests/ooxml.test.ts` | 空白 DOCX/XLSX 往返、宏文档 L1、分页、CSV 引号、行列插入、排序/筛选 |
| `tests/theme.test.ts` | 主题色预设与深色变体、自定义色派生、对比度、CSS 变量落盘、默认设置 |

### Rust

```bash
cargo test --workspace
```

覆盖：

- `office-core`：OfficeKind 扩展名、能力等级与警告逻辑、SHA-256。
- `office-security`：路径穿越、宏/签名、输出扩展名白名单、干净包。
- `office-ooxml`：DOCX/XLSX/PPTX 检测、真实样本扫描、PPTX L1、容器验证。
- `office-storage`：原子写入 + 备份、SQLite 最近文件、恢复快照。
- `office-font`：系统字体目录扫描。
- `office-pdf`：LibreOffice 发现与缺失输入错误。
- `office-platform`：平台信息。

### 构建

```bash
pnpm typecheck
pnpm build
pnpm --filter @mini-office/desktop exec tauri build --debug --bundles app
```

## 2. 真实文件夹具

见 [`tests/fixtures/README.md`](../tests/fixtures/README.md)。

- DOCX：中文租赁合同（2 页、表格、图片、页眉页脚、页码域、红色标注/高亮）。
- XLSX：3 个工作表、共享字符串、日期/货币样式、合并、冻结、筛选、公式。
- PPTX：3 页 16:9、主题、形状、图片、表格、备注。

这些文件由 python-docx / openpyxl / python-pptx 生成，不包含真实个人或商业数据。

## 3. 浏览器端验证（手动/自动化）

仓库开发过程中使用 Playwright + 系统 Chrome for Testing 做端到端验证。
当前不把 Playwright 作为 CI 依赖，原因是它需要本机浏览器二进制；可在发布前补齐 CI 浏览器。

已验证场景：

1. 首页渲染、深色主题、命令面板、设置。
2. 拖拽打开 DOCX/XLSX/PPTX/CSV。
3. Word 查找（计数、上下一个）、选中文本加粗、撤销恢复 DOM、编辑后保存。
4. Excel 单元格编辑、公式输入、总计重算、加粗/填充、合并、冻结、排序、筛选、查找。
5. PPT 缩略图、全屏演示、方向键换页、备注、PDF、图片 ZIP。
6. Word/Excel/PPT 导出文件可由 python-docx / openpyxl / PDF 解析器打开。
7. 兼容性报告/保存确认在 750×624 小窗口下无横向溢出；长文件名自动解码并截断。
8. 颜色面板、下拉菜单、悬停提示、Word 字体/字号/颜色/高亮与 Excel 填充色。
9. 浏览器 Fullscreen API 进入/退出 PPT 全屏演示；Tauri 原生全屏与原生打印命令完成构建验证。
10. 编辑后刷新页面，首页显示未保存快照，恢复后文本仍存在。
11. Tauri `.app` 在可写 HOME 下启动并保持运行；DMG 打包受沙箱 `hdiutil` 限制。

## 4. 手工验收建议

### Word

- 打开 `tests/fixtures/docx/contract.docx`，确认标题、页眉、页码、表格边框、印章图片。
- 编辑正文，工具栏设置加粗/颜色/对齐/行距，插入表格/图片，撤销/重做。
- 另存副本，用 Microsoft Word / LibreOffice 打开，检查页眉页脚、表格、图片与分页。
- 用含修订、批注、宏的文档确认只读/另存提示。

### Excel

- 打开 `tests/fixtures/xlsx/workbook.xlsx`，确认三个工作表、合并、冻结、筛选、金额/日期格式。
- 修改 C3 金额，确认 I2 合计、汇总表 SUMIF 更新。
- 输入跨表公式、循环引用、`#DIV/0!`，确认结果与 Excel 一致。
- 应用字体/填充/边框/数字格式，另存副本后由 Excel/LibreOffice 重新打开。

### PowerPoint

- 打开 `tests/fixtures/pptx/deck.pptx`，确认三页缩略图、形状、图片、表格、备注。
- F5 全屏演示，方向键/空格/Home/End/Esc，N 显示备注。
- 导出 PDF 与图片 ZIP，检查页数/尺寸。

### 桌面

- `pnpm tauri:dev` 启动桌面窗口，打开/保存文件，确认原生对话框与原子写入备份文件。
- 检查应用数据目录：`~/Library/Application Support/MiniOffice`（macOS）。
- 断网启动、重启后恢复。

## 5. 已知测试限制

- jsdom 没有真实字体/Canvas 渲染，PDF 与视觉结果需浏览器验证。
- Playwright 验证脚本依赖本机 Chrome 路径，未纳入仓库 CI。
- 原生毛玻璃、Tauri 打印对话框、窗口拖拽需要 macOS 实机视觉/交互验收；当前完成构建与代码级验证。
- 10 万行性能、三平台字体差异、Windows/Linux 实机打包尚未形成自动化基线。
