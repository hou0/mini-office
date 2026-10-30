import {
  createBlankDoc,
  type DocModel,
  type FormatHandler,
} from '@mini-office/document-model';
import {
  analyzeCapabilities,
  detectKind,
  loadOoxmlPackage,
} from '@mini-office/ooxml';
import { ErrorCodes, OfficeError, type DocumentStats } from '@mini-office/shared';
import { parseDocxPackage } from './parse';
import { serializeDocx } from './serialize';

export const wordFormatHandler: FormatHandler<DocModel> = {
  kind: 'docx',
  extensions: ['.docx', '.docm', '.dotx'],

  createBlank(): DocModel {
    return createBlankDoc();
  },

  async open(input) {
    const pkg = await loadOoxmlPackage(input.bytes);
    try {
      const detection = await detectKind(pkg, input.name);
      if (detection.kind !== 'docx') {
        throw new OfficeError(
          ErrorCodes.UNSUPPORTED_FORMAT,
          '该文件不是 Word 文档（DOCX）。',
        );
      }
      const capability = await analyzeCapabilities(pkg, 'docx', {
        fileName: input.name,
        knownFonts: input.knownFonts,
      });
      const data = await parseDocxPackage(pkg, { fileName: input.name });
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
    return serializeDocx(data, {
      sourceBytes: context.sourceBytes,
      name: context.name,
      // 功能：模板保存时把面板状态与 Word / WPS 兼容的邮件合并信息一并写出；时间：2026-10-29；作者：lq。
      mailMergeState: data.meta?.mailMergeState,
      mailMerge: data.meta?.mailMergeTemplate,
    });
  },

  getStats(data): DocumentStats {
    return {
      words: data.stats?.words ?? 0,
      characters: data.stats?.characters ?? 0,
    };
  },
};

export default wordFormatHandler;
