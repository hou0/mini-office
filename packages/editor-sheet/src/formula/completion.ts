/**
 * 功能：提供公式函数的中文说明、参数格式、光标上下文匹配和局部补全。
 * 时间：2026-10-24；作者：lq。
 */
import { listFunctions } from './functions';

export interface FunctionSuggestion {
  name: string;
  description: string;
  signature: string;
  maxArgs: number;
}

export interface FormulaCompletion {
  start: number;
  end: number;
  prefix: string;
  suggestions: FunctionSuggestion[];
}

// 功能：为内置函数提供简明中文帮助，中括号表示可选参数；时间：2026-10-24；作者：lq。
const HELP: Record<string, [description: string, parameters: string]> = {
  SUM: ['求和', '数值1, [数值2], …'],
  AVERAGE: ['计算平均值', '数值1, [数值2], …'],
  COUNT: ['统计数字个数', '[值1], [值2], …'],
  COUNTA: ['统计非空值个数', '[值1], [值2], …'],
  MAX: ['返回最大值', '[数值1], [数值2], …'],
  MIN: ['返回最小值', '[数值1], [数值2], …'],
  PRODUCT: ['计算乘积', '[数值1], [数值2], …'],
  MEDIAN: ['计算中位数', '数值1, [数值2], …'],
  STDEV: ['计算样本标准差', '数值1, [数值2], …'],
  'STDEV.S': ['计算样本标准差', '数值1, [数值2], …'],
  STDEVP: ['计算总体标准差', '数值1, [数值2], …'],
  'STDEV.P': ['计算总体标准差', '数值1, [数值2], …'],
  VAR: ['计算样本方差', '数值1, [数值2], …'],
  VARP: ['计算总体方差', '数值1, [数值2], …'],
  ROUND: ['四舍五入到指定位数', '数值, 位数'],
  ROUNDUP: ['向远离零的方向舍入', '数值, 位数'],
  ROUNDDOWN: ['向零的方向舍入', '数值, 位数'],
  INT: ['向下取整', '数值'],
  TRUNC: ['截断小数', '数值, [位数]'],
  ABS: ['返回绝对值', '数值'],
  SIGN: ['返回数值的符号', '数值'],
  MOD: ['返回余数', '数值, 除数'],
  POWER: ['计算乘方', '底数, 指数'],
  SQRT: ['计算平方根', '数值'],
  EXP: ['计算 e 的幂', '指数'],
  LN: ['计算自然对数', '数值'],
  LOG: ['计算指定底数的对数', '数值, [底数]'],
  LOG10: ['计算以 10 为底的对数', '数值'],
  FLOOR: ['向下舍入到指定倍数', '数值, [倍数]'],
  CEILING: ['向上舍入到指定倍数', '数值, [倍数]'],
  PI: ['返回圆周率', ''],
  RAND: ['返回 0 到 1 之间的随机数', ''],
  RANDBETWEEN: ['返回指定范围内的随机整数', '最小值, 最大值'],
  GCD: ['计算最大公约数', '数值1, [数值2], …'],
  LCM: ['计算最小公倍数', '数值1, [数值2], …'],
  FACT: ['计算阶乘', '数值'],
  IF: ['按条件返回不同结果', '条件, 为真时的值, [为假时的值]'],
  IFERROR: ['公式出错时返回指定值', '值, 出错时的值'],
  IFNA: ['出现 #N/A 时返回指定值', '值, 未找到时的值'],
  IFS: ['按多个条件返回首个匹配结果', '条件1, 值1, [条件2, 值2], …'],
  SWITCH: ['按表达式匹配结果', '表达式, 匹配值1, 结果1, …'],
  AND: ['判断条件是否全部成立', '条件1, [条件2], …'],
  OR: ['判断是否至少一个条件成立', '条件1, [条件2], …'],
  NOT: ['对逻辑值取反', '条件'],
  TRUE: ['返回逻辑值 TRUE', ''],
  FALSE: ['返回逻辑值 FALSE', ''],
  COUNTIF: ['按条件统计个数', '范围, 条件'],
  COUNTIFS: ['按多个条件统计个数', '范围1, 条件1, [范围2, 条件2], …'],
  SUMIF: ['按条件求和', '条件范围, 条件, [求和范围]'],
  SUMIFS: ['按多个条件求和', '求和范围, 条件范围1, 条件1, …'],
  AVERAGEIF: ['按条件求平均值', '条件范围, 条件, [平均值范围]'],
  AVERAGEIFS: ['按多个条件求平均值', '平均值范围, 条件范围1, 条件1, …'],
  MAXIFS: ['按多个条件返回最大值', '数值范围, 条件范围1, 条件1, …'],
  MINIFS: ['按多个条件返回最小值', '数值范围, 条件范围1, 条件1, …'],
  SUMPRODUCT: ['计算对应元素乘积之和', '数组1, [数组2], …'],
  VLOOKUP: ['在首列查找并返回指定列的值', '查找值, 范围, 列序号, [近似匹配]'],
  HLOOKUP: ['在首行查找并返回指定行的值', '查找值, 范围, 行序号, [近似匹配]'],
  XLOOKUP: ['查找并返回对应位置的值', '查找值, 查找范围, 返回范围, [未找到时的值], [匹配模式], [搜索模式]'],
  INDEX: ['返回指定位置的值', '范围, 行序号, [列序号]'],
  MATCH: ['返回查找值的相对位置', '查找值, 范围, [匹配类型]'],
  LEFT: ['从左侧提取字符', '文本, [字符数]'],
  RIGHT: ['从右侧提取字符', '文本, [字符数]'],
  MID: ['从指定位置提取字符', '文本, 起始位置, 字符数'],
  LEN: ['返回文本长度', '文本'],
  CONCAT: ['连接文本或区域', '文本1, [文本2], …'],
  CONCATENATE: ['连接多个文本', '文本1, [文本2], …'],
  TRIM: ['删除多余空格', '文本'],
  LOWER: ['转换为小写字母', '文本'],
  UPPER: ['转换为大写字母', '文本'],
  PROPER: ['将每个单词的首字母大写', '文本'],
  SUBSTITUTE: ['替换匹配文本', '文本, 原文本, 新文本, [替换序号]'],
  REPLACE: ['替换指定位置的字符', '文本, 起始位置, 字符数, 新文本'],
  FIND: ['查找文本位置，区分大小写', '查找文本, 文本, [起始位置]'],
  SEARCH: ['查找文本位置，不区分大小写', '查找文本, 文本, [起始位置]'],
  EXACT: ['比较文本是否完全相同', '文本1, 文本2'],
  TEXTJOIN: ['使用分隔符连接文本', '分隔符, 忽略空值, 文本1, …'],
  REPT: ['重复文本', '文本, 次数'],
  TEXT: ['按格式将数值转为文本', '数值, 格式'],
  VALUE: ['将文本转为数值', '文本'],
  CHAR: ['按编码返回字符', '编码'],
  CODE: ['返回首字符的编码', '文本'],
  TODAY: ['返回今天的日期', ''],
  NOW: ['返回当前日期和时间', ''],
  DATE: ['由年月日生成日期', '年, 月, 日'],
  TIME: ['由时分秒生成时间', '时, 分, 秒'],
  YEAR: ['提取年份', '日期'],
  MONTH: ['提取月份', '日期'],
  DAY: ['提取日', '日期'],
  HOUR: ['提取小时', '时间'],
  MINUTE: ['提取分钟', '时间'],
  SECOND: ['提取秒数', '时间'],
  WEEKDAY: ['返回星期序号', '日期, [返回类型]'],
  EDATE: ['计算相隔指定月数的日期', '开始日期, 月数'],
  EOMONTH: ['返回指定月份的最后一天', '开始日期, 月数'],
  DATEDIF: ['计算两个日期的间隔', '开始日期, 结束日期, 单位'],
  ISBLANK: ['判断是否为空', '值'],
  ISNUMBER: ['判断是否为数字', '值'],
  ISTEXT: ['判断是否为文本', '值'],
  ISLOGICAL: ['判断是否为逻辑值', '值'],
  ISERROR: ['判断是否为错误值', '值'],
  ISERR: ['判断是否为 #N/A 以外的错误', '值'],
  ISNA: ['判断是否为 #N/A', '值'],
  ISEVEN: ['判断是否为偶数', '数值'],
  ISODD: ['判断是否为奇数', '数值'],
  N: ['将值转为数值', '值'],
  T: ['保留文本，其他值返回空文本', '值'],
  NA: ['返回 #N/A 错误', ''],
  ROW: ['返回行号', '[引用]'],
  COLUMN: ['返回列号', '[引用]'],
  ROWS: ['返回区域行数', '范围'],
  COLUMNS: ['返回区域列数', '范围'],
};

// 功能：仅在函数位置匹配光标左侧前缀，排除字符串、工作表名和单元格引用；时间：2026-10-24；作者：lq。
export function getFormulaCompletion(text: string, caret: number): FormulaCompletion | null {
  if (!text.startsWith('=') || caret < 1 || caret > text.length) return null;
  const match = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(text.slice(0, caret));
  if (!match) return null;
  const start = match.index;
  const before = text.slice(0, start).trimEnd();
  if (!/[=+\-*/^&<>,;(]$/.test(before)) return null;
  let quote: string | null = null;
  for (let index = 1; index < start; index += 1) {
    const char = text[index];
    if (quote && char === quote && text[index + 1] === quote) { index += 1; continue; }
    if (quote === char) quote = null;
    else if (!quote && (char === '"' || char === "'")) quote = char;
  }
  if (quote) return null;
  let end = caret;
  while (end < text.length && /[A-Za-z0-9_.]/.test(text[end])) end += 1;
  if (/^\s*[!:]/.test(text.slice(end))) return null;
  const prefix = match[0].toUpperCase();
  const suggestions = listFunctions()
    .filter(({ name }) => name.startsWith(prefix))
    .map(({ name, minArgs, maxArgs }) => {
      const [description, parameters] = HELP[name] ?? ['工作表函数', maxArgs === 0 ? '' : minArgs === 0 ? '[参数], …' : '参数, …'];
      return { name, description, signature: `${name}(${parameters})`, maxArgs };
    });
  return suggestions.length ? { start, end, prefix, suggestions } : null;
}

// 功能：只替换当前函数名，保留公式后半段及已有括号，并返回继续输入的位置；时间：2026-10-24；作者：lq。
export function applyFormulaCompletion(text: string, context: FormulaCompletion, suggestion: FunctionSuggestion) {
  const suffix = text.slice(context.end);
  const existingParen = /^\s*\(/.exec(suffix);
  const insertion = suggestion.name + (existingParen ? '' : suggestion.maxArgs === 0 ? '()' : '(');
  return {
    text: text.slice(0, context.start) + insertion + suffix,
    caret: context.start + insertion.length + (existingParen?.[0].length ?? 0),
  };
}
