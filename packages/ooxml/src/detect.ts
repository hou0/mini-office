import { extensionOf, OfficeError, ErrorCodes, type OfficeKind } from '@mini-office/shared';
import type { OoxmlPackage } from './package';
import { mainDocumentPart } from './relationships';

export interface DetectionResult {
  kind: OfficeKind;
  extension: string;
  macroEnabled: boolean;
}

const EXTENSION_KIND: Record<string, OfficeKind> = {
  '.docx': 'docx',
  '.docm': 'docx',
  '.dotx': 'docx',
  '.xlsx': 'xlsx',
  '.xlsm': 'xlsx',
  '.xltx': 'xlsx',
  '.pptx': 'pptx',
  '.pptm': 'pptx',
  '.potx': 'pptx',
};

export function kindFromExtension(fileName: string): OfficeKind | null {
  return EXTENSION_KIND[extensionOf(fileName)] ?? null;
}

export function isMacroExtension(fileName: string): boolean {
  return ['.docm', '.xlsm', '.pptm', '.xltm', '.potm'].includes(extensionOf(fileName));
}

export async function detectKind(pkg: OoxmlPackage, fileName = ''): Promise<DetectionResult> {
  const extension = extensionOf(fileName);
  const macroEnabled = isMacroExtension(fileName);

  const main = await mainDocumentPart(pkg);
  if (main?.startsWith('word/')) return { kind: 'docx', extension, macroEnabled };
  if (main?.startsWith('ppt/')) return { kind: 'pptx', extension, macroEnabled };
  if (main?.startsWith('xl/')) return { kind: 'xlsx', extension, macroEnabled };

  if (pkg.has('word/document.xml')) return { kind: 'docx', extension, macroEnabled };
  if (pkg.has('xl/workbook.xml')) return { kind: 'xlsx', extension, macroEnabled };
  if (pkg.has('ppt/presentation.xml')) return { kind: 'pptx', extension, macroEnabled };

  const fromName = kindFromExtension(fileName);
  if (fromName) return { kind: fromName, extension, macroEnabled };

  throw new OfficeError(
    ErrorCodes.UNSUPPORTED_FORMAT,
    '无法识别该 Office 文件格式，V1 支持 DOCX、XLSX、PPTX。',
  );
}
