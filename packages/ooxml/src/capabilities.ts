import type {
  CapabilityFeature,
  CapabilityLevel,
  CapabilityReport,
  CompatibilityWarning,
  OfficeKind,
  WarningSeverity,
} from '@mini-office/shared';
import type { OoxmlPackage } from './package';
import { asArray, attr, children, node, parseXml, textOf } from './xml';

export interface CapabilityScanOptions {
  fileName?: string;
  knownFonts?: string[];
}

interface ScanState {
  features: CapabilityFeature[];
  warnings: CompatibilityWarning[];
  fontNames: Set<string>;
}

const CJK_RANGES = /[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]/;

function addFeature(
  state: ScanState,
  code: string,
  label: string,
  supported: boolean,
  severity: WarningSeverity,
  note?: string,
): void {
  if (state.features.some((feature) => feature.code === code)) return;
  state.features.push({ code, label, supported, severity, note });
}

function addWarning(
  state: ScanState,
  code: string,
  message: string,
  severity: WarningSeverity,
  detail?: string,
): void {
  if (state.warnings.some((warning) => warning.code === code)) return;
  state.warnings.push({ code, message, severity, detail });
}

async function readIfPresent(pkg: OoxmlPackage, part: string): Promise<string> {
  return pkg.has(part) ? pkg.text(part) : '';
}

function collectFontNamesFromXml(xml: string, state: ScanState): void {
  if (!xml) return;
  const add = (raw: string | undefined) => {
    const name = raw?.trim();
    if (
      name &&
      !name.startsWith('+') &&
      !/^(minor|major|none|default)$/i.test(name) &&
      name.length < 80
    ) {
      state.fontNames.add(name);
    }
  };
  // WordprocessingML: only actual w:rFonts elements (not w:lang eastAsia attributes).
  for (const tag of xml.match(/<w:rFonts\b[^>]*>/g) ?? []) {
    const attributes = /w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"/g;
    let match: RegExpExecArray | null;
    while ((match = attributes.exec(tag))) add(match[1]);
  }
  // DrawingML run/theme fonts.
  const drawing = /<a:(?:latin|ea|cs)\s+typeface="([^"]+)"/g;
  let drawingMatch: RegExpExecArray | null;
  while ((drawingMatch = drawing.exec(xml))) add(drawingMatch[1]);
  // Excel style fonts (<name val="...">) and Word font table (<w:font w:name="...">).
  const styleName = /<name\s+val="([^"]+)"/g;
  let styleMatch: RegExpExecArray | null;
  while ((styleMatch = styleName.exec(xml))) add(styleMatch[1]);
  const wordFont = /<w:font\b[^>]*w:name="([^"]+)"/g;
  let wordFontMatch: RegExpExecArray | null;
  while ((wordFontMatch = wordFont.exec(xml))) add(wordFontMatch[1]);
}

const FONT_ALIASES: Record<string, string[]> = {
  宋体: ['simsun', 'songti sc', 'stsong', 'noto serif cjk sc', 'source han serif sc', 'serif'],
  新宋体: ['simsun', 'songti sc', 'stsong'],
  仿宋: ['fangsong', 'stfangsong', 'fangsong_gb2312'],
  '仿宋_GB2312': ['fangsong', 'stfangsong', 'fangsong_gb2312'],
  黑体: ['simhei', 'heiti sc', 'stheiti', 'noto sans cjk sc', 'source han sans sc'],
  楷体: ['kaiti', 'stkaiti', 'kaiti_gb2312', 'kaiti sc'],
  '楷体_GB2312': ['kaiti', 'stkaiti', 'kaiti_gb2312', 'kaiti sc'],
  微软雅黑: ['microsoft yahei', 'microsoftyahei'],
  '微软雅黑 Light': ['microsoft yahei', 'microsoftyahei'],
  等线: ['dengxian'],
  '等线 Light': ['dengxian'],
  思源宋体: ['source han serif sc', 'noto serif cjk sc'],
  思源黑体: ['source han sans sc', 'noto sans cjk sc'],
  // 功能：华文/方正等常见商业字体名补上等价关系，避免已能正确渲染的字体仍被报“缺失”；时间：2026-10-31；作者：lq。
  华文宋体: ['stsong', 'songti sc', 'simsun'],
  华文中宋: ['stzhongsong', 'songti sc', 'stsong'],
  方正小标宋简体: ['stzhongsong', 'songti sc', 'stsong'],
  方正书宋简体: ['stsong', 'songti sc'],
  华文仿宋: ['stfangsong', 'fangsong'],
  方正仿宋简体: ['stfangsong', 'fangsong'],
  华文楷体: ['stkaiti', 'kaiti', 'kaiti sc'],
  方正楷体简体: ['stkaiti', 'kaiti', 'kaiti sc'],
  华文细黑: ['stheiti', 'heiti sc'],
  隶书: ['baoli sc', 'stliti'],
  华文隶书: ['stliti', 'baoli sc'],
  幼圆: ['yuanti sc', 'youyuan'],
};

function normalizeFont(name: string): string {
  return name.toLowerCase().replace(/[\s_-]+/g, '');
}

export function findMissingFonts(used: string[], available: string[]): string[] {
  if (available.length === 0) return [];
  const availableSet = new Set(available.map(normalizeFont));
  const missing: string[] = [];
  for (const font of used) {
    const normalized = normalizeFont(font);
    if (availableSet.has(normalized)) continue;
    const aliases = FONT_ALIASES[font] ?? [];
    if (aliases.some((alias) => availableSet.has(normalizeFont(alias)))) continue;
    // Serif/sans-serif generic fallbacks are always considered available.
    if (/^(serif|sans-serif|monospace|系统默认)$/i.test(font)) continue;
    missing.push(font);
  }
  return [...new Set(missing)];
}

async function scanDocx(pkg: OoxmlPackage, state: ScanState): Promise<void> {
  const documentXml = await readIfPresent(pkg, 'word/document.xml');
  const stylesXml = await readIfPresent(pkg, 'word/styles.xml');
  const fontTableXml = await readIfPresent(pkg, 'word/fontTable.xml');
  const settingsXml = await readIfPresent(pkg, 'word/settings.xml');
  collectFontNamesFromXml(documentXml, state);
  collectFontNamesFromXml(stylesXml, state);
  collectFontNamesFromXml(fontTableXml, state);

  const macros = pkg.names().filter((name) => /vbaProject\.bin$/i.test(name));
  if (macros.length > 0) {
    addFeature(state, 'docx.macro', 'VBA 宏', false, 'blocking', 'V1 不执行宏，且不覆盖宏部件。');
    addWarning(
      state,
      'docx.macro',
      '检测到 VBA 宏。V1 不会执行宏，将以只读方式打开，避免破坏原始宏部件。',
      'blocking',
    );
  }
  if (pkg.names().some((name) => name.startsWith('_xmlsignatures/'))) {
    addFeature(state, 'docx.signature', '数字签名', false, 'blocking');
    addWarning(
      state,
      'docx.signature',
      '文档包含数字签名。任何编辑都会使签名失效，因此以只读方式打开。',
      'blocking',
    );
  }
  if (documentXml.includes('<w:ins ') || documentXml.includes('<w:del ')) {
    addFeature(state, 'docx.revisions', '修订记录', false, 'warning');
    addWarning(
      state,
      'docx.revisions',
      '文档包含未接受的修订。V1 可以查看，但不会编辑修订；保存前建议在 Word 中接受或拒绝修订。',
      'warning',
    );
  }
  if (pkg.has('word/comments.xml') || documentXml.includes('commentReference')) {
    addFeature(state, 'docx.comments', '批注', false, 'warning');
    addWarning(state, 'docx.comments', '文档包含批注。批注内容会保留，但 V1 不提供批注编辑。', 'warning');
  }
  if (documentXml.includes('<w:sdt')) {
    addFeature(state, 'docx.contentControls', '内容控件', false, 'warning');
    addWarning(
      state,
      'docx.contentControls',
      '文档包含内容控件（如日期选择器、下拉列表），V1 以普通文本方式显示。',
      'warning',
    );
  }
  if (pkg.names().some((name) => /^word\/embeddings\//.test(name)) || documentXml.includes('<w:object')) {
    addFeature(state, 'docx.embedded', '嵌入对象', false, 'warning');
    addWarning(
      state,
      'docx.embedded',
      '文档包含嵌入对象（如 Excel 图表）。对象会原样保留，但无法在 Mini Office 中编辑。',
      'warning',
    );
  }
  if (documentXml.includes('<w:altChunk')) {
    addFeature(state, 'docx.altChunk', '外部内容块', false, 'warning');
  }
  if (documentXml.includes('<w:drawing') || documentXml.includes('<w:pict')) {
    addFeature(state, 'docx.images', '图片与图形', true, 'info');
  }
  if (documentXml.includes('<w:tbl>')) addFeature(state, 'docx.tables', '表格', true, 'info');
  if (documentXml.includes('<w:headerReference') || pkg.has('word/header1.xml')) {
    addFeature(state, 'docx.headerFooter', '页眉页脚', true, 'info');
  }
  if (settingsXml.includes('<w:trackChanges')) {
    addWarning(state, 'docx.trackChangesOn', '文档开启了“修订跟踪”。', 'warning');
  }
  addFeature(state, 'docx.fields', '域与页码', true, 'info');
}

async function scanXlsx(pkg: OoxmlPackage, state: ScanState): Promise<void> {
  const stylesXml = await readIfPresent(pkg, 'xl/styles.xml');
  collectFontNamesFromXml(stylesXml, state);
  const themeXml = await readIfPresent(pkg, 'xl/theme/theme1.xml');
  collectFontNamesFromXml(themeXml, state);

  if (pkg.names().some((name) => /vbaProject\.bin$/i.test(name))) {
    addFeature(state, 'xlsx.macro', 'VBA 宏', false, 'blocking');
    addWarning(
      state,
      'xlsx.macro',
      '工作簿包含 VBA 宏。V1 不执行宏，将以只读方式打开；请另存为无宏副本后再编辑。',
      'blocking',
    );
  }
  if (pkg.has('xl/externalLinks/externalLink1.xml') || pkg.names().some((n) => n.startsWith('xl/externalLinks/'))) {
    addFeature(state, 'xlsx.externalLinks', '外部工作簿链接', false, 'warning');
    addWarning(
      state,
      'xlsx.externalLinks',
      '工作簿包含外部链接。链接缓存会保留，但不会自动刷新外部数据。',
      'warning',
    );
  }
  if (pkg.names().some((name) => name.startsWith('xl/pivotCache/') || name.startsWith('xl/pivotTables/'))) {
    addFeature(state, 'xlsx.pivot', '数据透视表', false, 'warning');
    addWarning(
      state,
      'xlsx.pivot',
      '工作簿包含数据透视表。透视结果会保留，但 V1 不支持编辑或刷新。',
      'warning',
    );
  }
  if (pkg.names().some((name) => name.startsWith('xl/charts/'))) {
    addFeature(state, 'xlsx.charts', '图表', false, 'warning');
    addWarning(
      state,
      'xlsx.charts',
      '工作簿包含图表。图表部件会保留，但编辑对应工作表后需要重新核对显示。',
      'warning',
    );
  }
  // 功能：内嵌图片可编辑，图表和其他图形继续单独标识兼容范围；时间：2026-10-24；作者：lq。
  const drawingParts = pkg.names().filter((name) => /^xl\/drawings\/[^/]+\.xml$/.test(name));
  let hasPictures = false; let hasOtherDrawings = false;
  for (const part of drawingParts) {
    const xml = await pkg.text(part);
    if (/<(?:\w+:)?pic\b/.test(xml)) hasPictures = true;
    if (/<(?:\w+:)?(?:sp|graphicFrame|grpSp|cxnSp)\b/.test(xml)) hasOtherDrawings = true;
  }
  if (hasPictures) addFeature(state, 'xlsx.images', '内嵌图片', true, 'info');
  if (hasOtherDrawings) addFeature(state, 'xlsx.drawings', '图形与图表对象', false, 'warning');
  if (pkg.names().some((name) => name.startsWith('xl/comments'))) {
    addFeature(state, 'xlsx.comments', '批注', false, 'warning');
  }

  const sheetParts = pkg.names().filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
  for (const part of sheetParts) {
    const xml = await pkg.text(part);
    if (/<conditionalFormatting\b/.test(xml)) {
      addFeature(state, 'xlsx.conditionalFormatting', '条件格式', false, 'warning');
      addWarning(
        state,
        'xlsx.conditionalFormatting',
        '工作表包含条件格式。V1 会保留规则但不会预览条件格式效果。',
        'warning',
      );
      break;
    }
  }
  addFeature(state, 'xlsx.formulas', '公式计算', true, 'info');
  addFeature(state, 'xlsx.sharedStrings', '共享字符串', true, 'info');
}

async function scanPptx(pkg: OoxmlPackage, state: ScanState): Promise<void> {
  if (pkg.names().some((name) => /vbaProject\.bin$/i.test(name))) {
    addFeature(state, 'pptx.macro', 'VBA 宏', false, 'blocking', '不会执行或覆盖宏部件。');
    addWarning(state, 'pptx.macro', '演示文稿包含 VBA 宏，为保护宏和原文件，当前以只读方式打开。', 'blocking');
  }
  const presentationXml = await readIfPresent(pkg, 'ppt/presentation.xml');
  const themeXml = await pkg.names().find((name) => /^ppt\/theme\/theme\d+\.xml$/.test(name));
  if (themeXml) collectFontNamesFromXml(await readIfPresent(pkg, themeXml), state);
  const slideParts = pkg.names().filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  let animations = false;
  let transitions = false;
  let media = false;
  let notes = false;
  for (const part of slideParts) {
    const xml = await pkg.text(part);
    if (/<p:timing\b/.test(xml)) animations = true;
    if (/<p:transition\b/.test(xml)) transitions = true;
    if (/<p:video\b|<p:audio\b/.test(xml)) media = true;
    if (/<a:tbl\b/.test(xml)) addFeature(state, 'pptx.tables', '表格', true, 'info');
    if (/<p:pic\b/.test(xml)) addFeature(state, 'pptx.pictures', '图片', true, 'info');
    if (/<c:chart\b|<dgm:relIds\b/.test(xml)) addFeature(state, 'pptx.graphicFrames', '图表/SmartArt（保留原部件）', false, 'warning');
  }
  notes = pkg.names().some((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name));
  if (animations) {
    addFeature(state, 'pptx.animations', '动画', false, 'warning');
    addWarning(
      state,
      'pptx.animations',
      '演示文稿包含动画。V1 只显示静态页面，播放时不执行动画。',
      'warning',
    );
  }
  if (transitions) {
    addFeature(state, 'pptx.transitions', '页面切换', false, 'info');
  }
  if (media) {
    addFeature(state, 'pptx.media', '视频/音频', false, 'warning');
    addWarning(state, 'pptx.media', '演示文稿包含视频或音频，V1 不会播放媒体。', 'warning');
  }
  if (pkg.names().some((name) => /^ppt\/embeddings\//.test(name))) {
    addFeature(state, 'pptx.embedded', '嵌入对象', false, 'warning');
  }
  if (notes) addFeature(state, 'pptx.notes', '演讲者备注', true, 'info');
  if (presentationXml) addFeature(state, 'pptx.theme', '主题与母版', true, 'info');
  addFeature(state, 'pptx.preview', '静态高保真预览', true, 'info');
  // 功能：PPT 开启基础内容编辑并默认另存副本保护原文件；时间：2026-10-24；作者：lq。
  addFeature(state, 'pptx.editing', '文字、表格、对象布局与备注编辑', true, 'info');
  addWarning(state, 'pptx.editing', '支持基础内容编辑，首次保存为副本。母版装饰、图表、动画和媒体保留原部件，复杂效果请在 PowerPoint 中核对。', 'info');
}

function levelFor(kind: OfficeKind, state: ScanState): CapabilityLevel {
  if (state.warnings.some((warning) => warning.severity === 'blocking')) return 1;
  if (kind === 'pptx') return 1;
  if (
    state.warnings.some(
      (warning) => warning.severity === 'warning' && warning.code !== 'fonts.missing',
    )
  ) {
    return 2;
  }
  return 3;
}

export async function analyzeCapabilities(
  pkg: OoxmlPackage,
  kind: OfficeKind,
  options: CapabilityScanOptions = {},
): Promise<CapabilityReport> {
  const state: ScanState = { features: [], warnings: [], fontNames: new Set() };
  if (kind === 'docx') await scanDocx(pkg, state);
  if (kind === 'xlsx') await scanXlsx(pkg, state);
  if (kind === 'pptx') await scanPptx(pkg, state);

  const missingFonts = findMissingFonts([...state.fontNames], options.knownFonts ?? []);
  if (missingFonts.length > 0) {
    addWarning(
      state,
      'fonts.missing',
      `缺少 ${missingFonts.slice(0, 4).join('、')}${missingFonts.length > 4 ? ' 等' : ''}字体，将使用替代字体显示，版面可能与原件不同。`,
      'warning',
      missingFonts.join(', '),
    );
  }

  return {
    level: levelFor(kind, state),
    warnings: state.warnings,
    features: state.features,
    missingFonts,
    scannedAt: Date.now(),
  };
}

/** Extract theme color scheme names (dk1/lt1/accent1…) from a DrawingML theme part. */
export async function readThemeColors(
  pkg: OoxmlPackage,
  part: string,
): Promise<Record<string, string>> {
  if (!pkg.has(part)) return {};
  const xml = parseXml(await pkg.text(part));
  const scheme = node(node(node(xml['a:theme']) ?? xml.theme)?.['a:themeElements'])?.['a:clrScheme'];
  const result: Record<string, string> = {};
  if (!scheme) return result;
  for (const [rawKey, value] of Object.entries(scheme)) {
    if (rawKey.startsWith('@_') || !value) continue;
    const key = rawKey.replace(/^a:/, '');
    const colorNode = node(value);
    if (!colorNode) continue;
    const srgb = attr(colorNode['a:srgbClr'] as Record<string, unknown>, 'val');
    const sys = attr(colorNode['a:sysClr'] as Record<string, unknown>, 'lastClr');
    const hex = srgb ?? sys;
    if (hex) result[key] = `#${hex}`;
  }
  // Common aliases.
  if (result.dk1) result.tx1 = result.dk1;
  if (result.lt1) result.bg1 = result.lt1;
  if (result.dk2) result.tx2 = result.dk2;
  if (result.lt2) result.bg2 = result.lt2;
  return result;
}

/** Read the default font from a theme's major/minor font scheme. */
export async function readThemeFonts(
  pkg: OoxmlPackage,
  part: string,
): Promise<{ major: string; minor: string }> {
  if (!pkg.has(part)) return { major: '', minor: '' };
  const xml = parseXml(await pkg.text(part));
  const fontScheme = node(node(node(xml['a:theme']) ?? xml.theme)?.['a:themeElements'])?.['a:fontScheme'];
  // 功能：主题字体优先采用东亚字体，避免中文文档错误回退到 Arial 等西文字体；时间：2026-10-28；作者：lq。
  const majorFont = node(node(fontScheme)?.['a:majorFont']);
  const minorFont = node(node(fontScheme)?.['a:minorFont']);
  const major = node(majorFont)?.['a:ea'] ?? node(majorFont)?.['a:latin'];
  const minor = node(minorFont)?.['a:ea'] ?? node(minorFont)?.['a:latin'];
  return {
    major: attr(major, 'typeface') ?? '',
    minor: attr(minor, 'typeface') ?? '',
  };
}

/** Enumerate the font families declared or used in a package (best-effort). */
export async function collectFonts(pkg: OoxmlPackage, kind: OfficeKind): Promise<string[]> {
  const state: ScanState = { features: [], warnings: [], fontNames: new Set() };
  if (kind === 'docx') {
    collectFontNamesFromXml(await readIfPresent(pkg, 'word/document.xml'), state);
    collectFontNamesFromXml(await readIfPresent(pkg, 'word/styles.xml'), state);
    collectFontNamesFromXml(await readIfPresent(pkg, 'word/fontTable.xml'), state);
  } else if (kind === 'xlsx') {
    collectFontNamesFromXml(await readIfPresent(pkg, 'xl/styles.xml'), state);
    collectFontNamesFromXml(await readIfPresent(pkg, 'xl/theme/theme1.xml'), state);
  } else {
    const theme = pkg.names().find((name) => /^ppt\/theme\/theme\d+\.xml$/.test(name));
    if (theme) collectFontNamesFromXml(await readIfPresent(pkg, theme), state);
    const slide = pkg.names().find((name) => /^ppt\/slides\/slide1\.xml$/.test(name));
    if (slide) collectFontNamesFromXml(await readIfPresent(pkg, slide), state);
  }
  return [...state.fontNames];
}

export function hasChineseText(xml: string): boolean {
  return CJK_RANGES.test(xml);
}
