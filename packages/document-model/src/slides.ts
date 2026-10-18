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
    }>;
  }>;
}

export interface SlideGroupElement extends ElementBase {
  type: 'group';
  children: SlideElement[];
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
            x: 0.1,
            y: 0.32,
            w: 0.8,
            h: 0.18,
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
            x: 0.15,
            y: 0.54,
            w: 0.7,
            h: 0.12,
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
