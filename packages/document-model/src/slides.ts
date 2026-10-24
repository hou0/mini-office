export type SlideAspect = '16:9' | '4:3' | 'custom';

export interface SlideColorFill {
  type: 'solid';
  color: string;
}
export interface SlideGradientFill {
  type: 'gradient';
  stops: Array<{ offset: number; color: string }>;
  angle?: number;
}
export interface SlideImageFill {
  type: 'image';
  src: string;
}
export interface SlideNoFill {
  type: 'none';
}
export type SlideFill = SlideColorFill | SlideGradientFill | SlideImageFill | SlideNoFill;

export interface SlideLine {
  color?: string;
  widthPx?: number;
  dash?: 'solid' | 'dash' | 'dot' | 'dashDot' | 'longDash';
}

export interface TextRunStyle {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  letterSpacingPx?: number;
}

export interface SlideTextRun {
  text: string;
  style: TextRunStyle;
  lineBreak?: boolean;
}

export interface SlideTextParagraph {
  runs: SlideTextRun[];
  align?: 'left' | 'center' | 'right' | 'justify';
  level?: number;
  bullet?: boolean;
  bulletChar?: string;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: number;
}

export interface ElementBase {
  id: string;
  name?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  // 功能：定位原始对象并判断文字改动，锁定母版装饰防止误编辑；时间：2026-10-24；作者：lq。
  sourceId?: string;
  sourceText?: string;
  inherited?: boolean;
}

export interface SlideTextElement extends ElementBase {
  type: 'text';
  paragraphs: SlideTextParagraph[];
  fill?: SlideFill;
  line?: SlideLine;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  margins?: { top: number; right: number; bottom: number; left: number };
  placeholderType?: string;
  shape?: string;
}

export interface SlideShapeElement extends ElementBase {
  type: 'shape';
  shape: string;
  fill?: SlideFill;
  line?: SlideLine;
  paragraphs?: SlideTextParagraph[];
  verticalAlign?: 'top' | 'middle' | 'bottom';
  adjustment?: Record<string, number>;
}

export interface SlideImageElement extends ElementBase {
  type: 'image';
  src: string;
  part?: string;
  crop?: { left: number; right: number; top: number; bottom: number };
  alt?: string;
}

export interface SlideTableElement extends ElementBase {
  type: 'table';
  columnWidths: number[];
  rowHeights: number[];
  rows: Array<{
    cells: Array<{
      text: SlideTextParagraph[];
      fill?: SlideFill;
      borders?: { color?: string; widthPx?: number };
      colSpan?: number;
      rowSpan?: number;
      verticalAlign?: 'top' | 'middle' | 'bottom';
      sourceIndex?: number;
    }>;
  }>;
}

export interface SlideGroupElement extends ElementBase {
  type: 'group';
  children: SlideElement[];
  sourceChildren?: string[];
}

export type SlideElement =
  | SlideTextElement
  | SlideShapeElement
  | SlideImageElement
  | SlideTableElement
  | SlideGroupElement;

export interface SlideModel {
  id: string;
  index: number;
  title?: string;
  layoutName?: string;
  masterName?: string;
  background?: SlideFill;
  elements: SlideElement[];
  notes?: string;
  hasAnimations?: boolean;
  hasMedia?: boolean;
  // 功能：按原部件身份保存、复制和删除页面及备注；时间：2026-10-24；作者：lq。
  sourcePart?: string;
  sourceElementIds?: string[];
  sourceNotes?: string;
}

export interface SlideDeck {
  version: 1;
  widthEmu: number;
  heightEmu: number;
  widthPx: number;
  heightPx: number;
  aspect: SlideAspect;
  slides: SlideModel[];
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  meta: { title?: string; author?: string; created?: string; modified?: string };
  hasAnimations?: boolean;
  hasMedia?: boolean;
  hasNotes?: boolean;
  stats?: { slides: number };
}

export function createBlankDeck(): SlideDeck {
  const widthEmu = 12_192_000;
  const heightEmu = 6_858_000;
  return {
    version: 1,
    widthEmu,
    heightEmu,
    widthPx: 1280,
    heightPx: 720,
    aspect: '16:9',
    slides: [
      {
        id: 'slide-1',
        index: 0,
        title: '单击添加标题',
        layoutName: 'Title Slide',
        elements: [
          {
            id: 'el-title',
            type: 'text',
            // 功能：新建幻灯片统一使用画布像素坐标；时间：2026-10-24；作者：lq。
            x: 128,
            y: 230,
            w: 1024,
            h: 130,
            paragraphs: [
              {
                runs: [{ text: '单击添加标题', style: { fontSizePt: 40, bold: true, color: '#1D2939' } }],
                align: 'center',
              },
            ],
            verticalAlign: 'middle',
          },
          {
            id: 'el-subtitle',
            type: 'text',
            x: 192,
            y: 389,
            w: 896,
            h: 86,
            paragraphs: [
              {
                runs: [
                  { text: 'Mini Office 演示文稿', style: { fontSizePt: 20, color: '#667085' } },
                ],
                align: 'center',
              },
            ],
            verticalAlign: 'middle',
          },
        ],
      },
    ],
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    meta: { title: '未命名演示文稿' },
    stats: { slides: 1 },
  };
}
