/** 功能：把邮件合并的筛选条件翻译成 Word 查询并生成标准 w:mailMerge，供 Word / WPS 识别；时间：2026-10-29；作者：lq。 */
import { escapeXml } from '@mini-office/ooxml';
import type { DocMergeCondition } from '@mini-office/document-model';

/** Word 邮件合并数据源的外部关系类型 */
export const MAIL_MERGE_SOURCE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource';
/** settings.xml.rels 中 w:dataSource 使用的关系 ID */
export const MAIL_MERGE_DATA_REL = 'rIdMailMergeSource';
/** settings.xml.rels 中 w:odso/w:src 使用的关系 ID */
export const MAIL_MERGE_ODSO_REL = 'rIdMailMergeOdso';

export interface WordMailMergeInfo {
  /** 数据源在磁盘上的绝对路径 */
  dataSourcePath: string;
  /** 数据源中的工作表名 */
  sheetName: string;
  /** 与面板一致的筛选条件 */
  conditions: DocMergeCondition[];
  /** 数据源的列名，用于 ODSO 字段映射 */
  fields: string[];
}

/** 标识符用反引号包裹，内部反引号按 SQL 约定翻倍 */
function quoteIdentifier(name: string): string {
  return `\`${name.replace(/`/g, '``')}\``;
}

/** SQL 字符串字面量转义：单引号翻倍 */
export function escapeSqlValue(value: string): string {
  return value.replace(/'/g, "''");
}

/** LIKE 模式里的 % _ [ 按 Jet/ACE 约定用方括号转义，避免被当成通配符 */
function escapeLikePattern(value: string): string {
  return value.replace(/[%_[]/g, (char) => `[${char}]`);
}

/** 单个条件转成 SQL 片段；数值比较遇到非数值时返回恒假，与面板的数值比较结果保持一致 */
export function mergeConditionToSql(condition: DocMergeCondition): string {
  const field = quoteIdentifier(condition.field);
  const value = condition.value;
  if (condition.op === 'empty') return `${field} IS NULL`;
  if (condition.op === 'notEmpty') return `${field} IS NOT NULL`;
  if (condition.op === 'contains') return `${field} LIKE '%${escapeLikePattern(escapeSqlValue(value))}%'`;
  if (condition.op === 'notContains') return `${field} NOT LIKE '%${escapeLikePattern(escapeSqlValue(value))}%'`;
  if (condition.op === 'eq') return `${field} = '${escapeSqlValue(value)}'`;
  if (condition.op === 'ne') return `${field} <> '${escapeSqlValue(value)}'`;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return '(1 = 0)';
  const operator = condition.op === 'gt' ? '>' : condition.op === 'gte' ? '>=' : condition.op === 'lt' ? '<' : '<=';
  return `${field} ${operator} ${numeric}`;
}

/**
 * 条件组合成 Word 查询。
 * 与 filterMergeRecords 一致采用从左到右求值（且/或 同级），因此逐层加括号，
 * 避免 SQL 里 AND 优先于 OR 造成的差异。
 */
export function conditionsToWordQuery(conditions: DocMergeCondition[], sheetName: string): string {
  const table = quoteIdentifier(`${sheetName}$`);
  if (!conditions.length) return `SELECT * FROM ${table}`;
  let expression = `(${mergeConditionToSql(conditions[0])})`;
  for (let index = 1; index < conditions.length; index++) {
    const operator = conditions[index].conjunction === 'or' ? 'OR' : 'AND';
    expression = `(${expression} ${operator} (${mergeConditionToSql(conditions[index])}))`;
  }
  return `SELECT * FROM ${table} WHERE ${expression}`;
}

/** Excel 数据源的 OLEDB 连接串：参数与 Word 自身写出的 Jet OLEDB 串保持一致，便于 Word 直接识别 */
export function buildWordConnectString(path: string): string {
  return `Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=${path};Mode=Read;Extended Properties="HDR=YES;IMEX=1;";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=37;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False`;
}

/** 磁盘路径转成外部关系使用的 file URI：Windows 路径规范化为正斜杠，与 Word 自身写法一致 */
export function fileUriTarget(path: string): string {
  const normalized = /^[A-Za-z]:[\\/]/.test(path) ? path.replace(/\\/g, '/') : path;
  return normalized.startsWith('/') ? `file://${normalized}` : `file:///${normalized}`;
}

/** 生成 w:mailMerge 片段；子元素顺序遵循 CT_MailMerge 的 schema sequence */
export function buildWordMailMergeXml(info: WordMailMergeInfo): string {
  const connectString = buildWordConnectString(info.dataSourcePath);
  const query = conditionsToWordQuery(info.conditions, info.sheetName);
  const table = `${info.sheetName}$`;
  const fieldMap = (info.fields ?? []).map((field, index) =>
    `<w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="${escapeXml(field)}"/><w:mappedName w:val="${escapeXml(field)}"/><w:column w:val="${index}"/></w:fieldMapData>`).join('');
  return `<w:mailMerge>`
    + `<w:mainDocumentType w:val="formLetters"/>`
    + `<w:linkToQuery/>`
    + `<w:dataType w:val="native"/>`
    + `<w:connectString w:val="${escapeXml(connectString)}"/>`
    + `<w:query w:val="${escapeXml(query)}"/>`
    + `<w:dataSource r:id="${MAIL_MERGE_DATA_REL}"/>`
    + `<w:viewMergedData/>`
    + `<w:activeRecord w:val="1"/>`
    + `<w:odso>`
    + `<w:udl w:val="${escapeXml(connectString)}"/>`
    + `<w:table w:val="${escapeXml(table)}"/>`
    + `<w:src r:id="${MAIL_MERGE_ODSO_REL}"/>`
    + `<w:colDelim w:val="9"/>`
    + `<w:type w:val="database"/>`
    + `<w:fHdr/>`
    + fieldMap
    + `</w:odso>`
    + `</w:mailMerge>`;
}
