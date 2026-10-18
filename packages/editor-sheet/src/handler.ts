import {
  createBlankWorkbook,
  type FormatHandler,
  type WorkbookModel,
} from '@mini-office/document-model';
import { analyzeCapabilities, detectKind, loadOoxmlPackage } from '@mini-office/ooxml';
import {
  ErrorCodes,
  OfficeError,
  baseName,
  extensionOf,
  type CapabilityReport,
  type DocumentStats,
} from '@mini-office/shared';
import { parseCsv, parseWorkbookPackage } from './parse';
import { serializeCsv, serializeXlsx } from './serialize';

function csvCapability(): CapabilityReport {
  return {
    level: 3,
    warnings: [
      {
        code: 'csv.textOnly',
        message: 'CSV 仅保存文本与数值，不包含字体、颜色、合并单元格等格式。',
        severity: 'info',
      },
    ],
    features: [
      { code: 'csv.values', label: '文本与数值', supported: true, severity: 'info' },
      { code: 'csv.formulas', label: '公式（保存为计算结果）', supported: false, severity: 'info' },
    ],
    missingFonts: [],
    scannedAt: Date.now(),
  };
}

export const sheetFormatHandler: FormatHandler<WorkbookModel> = {
  kind: 'xlsx',
  extensions: ['.xlsx', '.xlsm', '.xltx', '.csv'],

  createBlank(): WorkbookModel {
    return createBlankWorkbook();
  },

  async open(input) {
    if (extensionOf(input.name) === '.csv') {
      const text = new TextDecoder('utf-8').decode(input.bytes);
      return {
        data: parseCsv(text, input.name),
        capability: csvCapability(),
        readonly: false,
        displayName: input.name,
      };
    }
    const pkg = await loadOoxmlPackage(input.bytes);
    try {
      const detection = await detectKind(pkg, input.name);
      if (detection.kind !== 'xlsx') {
        throw new OfficeError(
          ErrorCodes.UNSUPPORTED_FORMAT,
          '该文件不是 Excel 工作簿（XLSX）。',
        );
      }
      const capability = await analyzeCapabilities(pkg, 'xlsx', {
        fileName: input.name,
        knownFonts: input.knownFonts,
      });
      const data = await parseWorkbookPackage(pkg, input.name);
      return {
        data,
        capability,
        readonly: capability.level <= 1,
        displayName: input.name,
      };
    } finally {
      pkg.dispose();
    }
  },

  async serialize(data, context) {
    if (data.csv) return serializeCsv(data);
    return serializeXlsx(data, {
      sourceBytes: context.sourceBytes,
      name: context.name,
    });
  },

  getStats(data): DocumentStats {
    return {
      cells: data.stats?.cells ?? Object.keys(data.sheets[0]?.cells ?? {}).length,
      sheets: data.sheets.length,
    };
  },
};

export default sheetFormatHandler;

export function describeCsvSource(name: string): string {
  return baseName(name);
}
