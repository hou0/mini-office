import {
  createBlankDeck,
  type FormatHandler,
  type SlideDeck,
} from '@mini-office/document-model';
import { analyzeCapabilities, detectKind, loadOoxmlPackage } from '@mini-office/ooxml';
import { ErrorCodes, OfficeError, type DocumentStats } from '@mini-office/shared';
import { deckStats, parsePptxPackage } from './parse';
import { serializePptx } from './serialize';

export const slidesFormatHandler: FormatHandler<SlideDeck> = {
  kind: 'pptx',
  extensions: ['.pptx', '.pptm', '.potx'],

  createBlank(): SlideDeck {
    return createBlankDeck();
  },

  async open(input) {
    const pkg = await loadOoxmlPackage(input.bytes);
    try {
      const detection = await detectKind(pkg, input.name);
      if (detection.kind !== 'pptx') {
        throw new OfficeError(
          ErrorCodes.UNSUPPORTED_FORMAT,
          '该文件不是 PowerPoint 演示文稿（PPTX）。',
        );
      }
      const capability = await analyzeCapabilities(pkg, 'pptx', {
        fileName: input.name,
        knownFonts: input.knownFonts,
      });
      const data = await parsePptxPackage(pkg, { fileName: input.name });
      // 功能：编辑器层将 PPTX 提升为需另存的 L2，同时底层能力扫描仍保留只读基线供兼容性报告使用；时间：2026-10-24；作者：lq。
      const blocked = capability.warnings.some((warning) => warning.severity === 'blocking');
      const editableCapability = {
        ...capability,
        level: blocked ? 1 as const : 2 as const,
        warnings: capability.warnings.filter((warning) => warning.code !== 'pptx.readonly'),
      };
      return {
        data,
        capability: editableCapability,
        readonly: blocked,
        displayName: input.name,
      };
    } finally {
      pkg.dispose();
    }
  },

  async serialize(data, context) {
    try {
      return await serializePptx(data, context.sourceBytes);
    } catch (error) {
      if (error instanceof OfficeError) throw error;
      throw new OfficeError(ErrorCodes.SAVE_FAILED, error instanceof Error ? error.message : String(error));
    }
  },

  getStats(data): DocumentStats {
    const stats = deckStats(data);
    return { slides: stats.slides };
  },
};

export default slidesFormatHandler;
