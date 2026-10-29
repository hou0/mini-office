/**
 * 生成示例插件的公文模板（.docx）。
 *
 * 为什么用脚本生成而不是提交一个手工文件：模板是二进制，改了没法 review。
 * 脚本跑出来的包是确定性的（固定 zip 时间戳），重新生成不会污染 git 历史。
 *
 * 用法：node scripts/make-template.mjs
 */
import JSZip from 'jszip';
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

const OUTPUT = resolve('plugins/pack-gov-doc/templates/gov-letter.docx');
const FIXED_DATE = new Date('1980-01-01T00:00:00Z');

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/><w:sz w:val="32"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`;

function paragraph(text, options = {}) {
  const { align = 'both', bold = false, sizeHalfPt = 32, font = '仿宋_GB2312', firstLineIndent = 640 } = options;
  const boldXml = bold ? '<w:b/>' : '';
  const indentXml = firstLineIndent > 0 ? `<w:ind w:firstLine="${firstLineIndent}"/>` : '';
  const spacing = '<w:spacing w:line="360" w:lineRule="auto"/>';
  return `<w:p><w:pPr><w:jc w:val="${align}"/>${spacing}${indentXml}<w:rPr><w:rFonts w:eastAsia="${font}"/><w:sz w:val="${sizeHalfPt}"/>${boldXml}</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="${font}"/><w:sz w:val="${sizeHalfPt}"/>${boldXml}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
${paragraph('关于           工作的通知', { align: 'center', bold: true, sizeHalfPt: 44, font: '方正小标宋简体', firstLineIndent: 0 })}
${paragraph('')}
${paragraph('各有关单位：', { firstLineIndent: 0 })}
${paragraph('为做好相关工作，现将有关事项通知如下，请结合实际认真贯彻落实。', {})}
${paragraph('一、总体要求', { bold: true, font: '黑体' })}
${paragraph('按照统一部署，明确责任分工，确保各项任务按期完成。', {})}
${paragraph('二、主要任务', { bold: true, font: '黑体' })}
${paragraph('（一）梳理现状。全面摸清底数，形成问题清单。', {})}
${paragraph('（二）制定方案。结合实际情况，细化实施步骤。', {})}
${paragraph('三、工作要求', { bold: true, font: '黑体' })}
${paragraph('加强组织领导，强化督促检查，及时报送进展情况。', {})}
${paragraph('')}
${paragraph('（此处加盖公章）', { align: 'right', firstLineIndent: 0 })}
${paragraph('年　　月　　日', { align: 'right', firstLineIndent: 0 })}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1588" w:header="851" w:footer="992"/></w:sectPr>
</w:body>
</w:document>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', CONTENT_TYPES, { date: FIXED_DATE, createFolders: false });
zip.file('_rels/.rels', ROOT_RELS, { date: FIXED_DATE, createFolders: false });
zip.file('word/document.xml', DOCUMENT, { date: FIXED_DATE, createFolders: false });
zip.file('word/_rels/document.xml.rels', DOC_RELS, { date: FIXED_DATE, createFolders: false });
zip.file('word/styles.xml', STYLES, { date: FIXED_DATE, createFolders: false });

const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
await mkdir(dirname(OUTPUT), { recursive: true });
await writeFile(OUTPUT, bytes);
console.log(`已生成 ${OUTPUT}（${bytes.byteLength} 字节）`);
