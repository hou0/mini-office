# ADR 0003：公式引擎

- 状态：已采纳
- 日期：2026-10-18

## 背景

PRD 要求 Excel 日常计算：SUM、AVERAGE、COUNT、MAX、MIN、ROUND、SUMIF、COUNTIF、IF、AND、OR、
IFERROR、IFS、VLOOKUP、XLOOKUP、INDEX、MATCH、文本、日期等，并覆盖相对/绝对引用、跨表引用、
循环引用、错误值、日期序列与增量重算。最终目标 50–80 个常用函数。

## 决策

自研 TypeScript 公式引擎，分层为：

```text
formula/tokenizer.ts → formula/parser.ts → formula/functions.ts → formula/engine.ts
```

- **Tokenizer**：数字/字符串/布尔/错误字面量、单元格/区域/整列/整行、带引号工作表名、
  运算符（+ - * / ^ & % = <> < > <= >=）、函数调用与括号。
- **Parser**：递归下降 + 优先级；生成可打印 AST，用于复制/填充与共享公式平移。
- **Functions**：注册表 + 惰性求值函数（IF/IFS/IFERROR/SWITCH），区域作为一等值传入，
  支持 criteria、lookup、文本、日期、信息函数。
- **Engine**：
  - 每次修改后对工作簿做一次重算；
  - 以单元格为键记忆化，递归求值，检测循环引用（`#CIRC!`）；
  - 跨表按名称（大小写不敏感）解析；
  - 范围按工作表已用维度裁剪（整列/整行），限制范围单元格数防止内存爆炸；
  - 计算结果缓存为 `computed[sheetId][row:col] = { value, display, error }`；
  - 日期使用 Excel 1900/1904 序列与 1900 闰年兼容规则；
  - 数字显示格式独立于值，支持内置/自定义 numFmt、日期、百分比、货币、千分位。

### 已实现函数（60+）

聚合/统计：SUM、AVERAGE、COUNT、COUNTA、MAX、MIN、PRODUCT、MEDIAN、STDEV(.S/P)、VAR(P)、
SUMPRODUCT、LARGE/SMALL 等扩展。

条件：SUMIF(S)、COUNTIF(S)、AVERAGEIF(S)、MAXIFS/MINIFS。

逻辑：IF、IFS、IFERROR、IFNA、AND、OR、NOT、SWITCH、TRUE/FALSE。

查找：VLOOKUP、HLOOKUP、XLOOKUP、INDEX、MATCH。

文本：LEFT、RIGHT、MID、LEN、CONCAT、CONCATENATE、TRIM、LOWER、UPPER、PROPER、
SUBSTITUTE、REPLACE、FIND、SEARCH、EXACT、TEXTJOIN、REPT、TEXT、VALUE、CHAR、CODE。

日期：TODAY、NOW、DATE、TIME、YEAR、MONTH、DAY、HOUR、MINUTE、SECOND、WEEKDAY、
EDATE、EOMONTH、DATEDIF。

数学：ROUND/ROUNDUP/ROUNDDOWN、INT、TRUNC、ABS、SIGN、MOD、POWER、SQRT、EXP、LN、LOG、
LOG10、FLOOR、CEILING、PI、RAND、RANDBETWEEN、GCD、LCM、FACT。

信息：ISBLANK、ISNUMBER、ISTEXT、ISLOGICAL、ISERROR、ISERR、ISNA、ISEVEN、ISODD、
N、T、NA、ROW、COLUMN、ROWS、COLUMNS。

## 理由

- 公式语法/错误值/日期系统必须能在 Node 中回归，不依赖桌面 Excel。
- 注册表 + AST 使新增函数成本低，复制/填充可以安全平移相对引用。
- 记忆化递归天然支持循环检测；对日常规模工作簿足够快。

## 后果

- 不支持 VBA 自定义函数、动态数组溢出、LAMBDA、Power Query、部分统计/金融/工程函数。
- 整列范围按已用维度裁剪，极端稀疏表可能与 Excel 的全列语义略有差异。
- 排序/行列插入不会自动调整公式引用（V1 限制），需要用户重新核对。
- 时间/日期本地时区处理为近似实现，跨时区文件可能存在细微差异。

## 测试

`tests/formula.test.ts` 覆盖算术、引用、逻辑、查找、文本、日期、条件聚合、跨表、循环、错误传播、
共享公式平移与格式；真实 XLSX 工作簿验证 SUM/SUMIF/VLOOKUP/XLOOKUP/IFERROR/DATE。

## 后续

- 依赖图增量重算（当前为一次记忆化全量求值）。
- 更多函数与数组语义。
- 公式引用在结构变更/排序时的自动平移。
- 与 Excel 对照的自动化黄金数据集。
