# 回归测试样例

这些文件是 Mini Office 的兼容性基线，均为**合成数据**，不包含真实合同或个人信息。

| 文件 | 生成方式 | 关键特性 |
|---|---|---|
| `docx/contract.docx` | python-docx 1.2.0 + Pillow | A4、页边距、页眉、页脚 PAGE/NUMPAGES 域、标题/加粗/颜色/高亮、首行缩进、固定行距、3×4 表格（Table Grid、合并单元格、表头底纹）、印章图片、分页符 |
| `docx/stamp.png` | Pillow | 合同印章示意图片 |
| `xlsx/workbook.xlsx` | openpyxl 3.1.5 | 3 个工作表、共享字符串、合并、冻结、自动筛选、列宽行高、金额/日期格式、边框、条件填充、SUM/AVERAGE/MAX/COUNT/SUMIF/VLOOKUP/XLOOKUP/IFERROR/DATE 公式 |
| `pptx/deck.pptx` | python-pptx 1.0.2 | 16:9、主题、标题/正文文本框、装饰矩形、圆角卡片、图片、3×3 表格、演讲者备注 |

## 重新生成

生成脚本使用以下 Python 包（与 Mini Office 的运行时无关）：

```bash
python3 -m pip install python-docx openpyxl python-pptx Pillow
```

生成逻辑与当前文件保持等价；如修改夹具，请同步更新：

- `tests/office-fixtures.test.ts` 中的断言；
- `docs/compatibility-matrix.md` 的样例表；
- Rust `crates/office-ooxml` 的夹具测试。

> 注意：夹具用于兼容性回归，不代表所有 Word/Excel/PowerPoint 特性均已支持。
