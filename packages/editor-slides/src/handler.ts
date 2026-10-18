import {
  createBlankDeck,
  type FormatHandler,
  type SlideDeck,
} from '@mini-office/document-model';
import { analyzeCapabilities, detectKind, loadOoxmlPackage } from '@mini-office/ooxml';
import { ErrorCodes, OfficeError, type DocumentStats } from '@mini-office/shared';
import { deckStats, parsePptxPackage } from './parse';

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
      return {
        data,
        capability,
        readonly: true,
        displayName: input.name,
      };
    } finally {
      pkg.dispose();
    }
  },

  async serialize(_data, context) {
    if (context.sourceBytes) return context.sourceBytes;
    throw new OfficeError(
      ErrorCodes.SAVE_FAILED,
      'V1 的 PPTX 仅支持静态预览与演示，不支持保存编辑结果。',
    );
  },

  getStats(data): DocumentStats {
    const stats = deckStats(data);
    return { slides: stats.slides };
  },
};

export default slidesFormatHandler;
