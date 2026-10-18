import { parseCsv, parseCsvRows } from '../../packages/editor-sheet/src/parse';
import { serializeCsv } from '../../packages/editor-sheet/src/serialize';
import type { WorkbookModel } from '@mini-office/document-model';

export { parseCsv, parseCsvRows };

export function serializedCsvText(workbook: WorkbookModel): string {
  return new TextDecoder('utf-8').decode(serializeCsv(workbook)).replace(/^\uFEFF/, '');
}
