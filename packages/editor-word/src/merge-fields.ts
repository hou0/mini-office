/** 功能：识别 Word 合并域的引号字段名和可选开关；时间：2026-10-24；作者：lq。 */
export function parseMergeField(instruction: string): string | undefined {
  const match = /^\s*MERGEFIELD\s+(?:"((?:[^"\\]|\\.)*)"|([^\s\\]+))/i.exec(instruction);
  return match ? (match[1]?.replace(/\\(["\\])/g, '$1') ?? match[2]) : undefined;
}
