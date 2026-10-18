import {
  createBlankDoc,
  createBlankWorkbook,
  type CellStyle,
  type DocBlock,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocTable,
  type RunStyle,
  type WorkbookModel,
} from '@mini-office/document-model';
import { uid } from '@mini-office/shared';
import { recalculateWorkbook } from '@mini-office/editor-sheet';

function run(text: string, style: RunStyle = {}): DocRun {
  return { text, style };
}

function paragraph(
  text: string,
  style: DocParagraph['style'] = {},
  runStyle: RunStyle = {},
): DocParagraph {
  return {
    type: 'paragraph',
    id: uid('p'),
    runs: [run(text, runStyle)],
    style: { lineHeight: 1.75, ...style },
  };
}

function heading(text: string): DocParagraph {
  return paragraph(text, { spaceBeforePx: 12, spaceAfterPx: 6, keepNext: true }, {
    fontFamily: '黑体',
    fontSizePt: 12,
    bold: true,
  });
}

function contractTable(): DocTable {
  const makeCell = (text: string, bold = false, fill?: string) => ({
    id: uid('cell'),
    colSpan: 1,
    rowSpan: 1,
    backgroundColor: fill,
    borders: {
      top: { style: 'single' as const, widthPx: 0.75, color: '#1D2939' },
      bottom: { style: 'single' as const, widthPx: 0.75, color: '#1D2939' },
      left: { style: 'single' as const, widthPx: 0.75, color: '#1D2939' },
      right: { style: 'single' as const, widthPx: 0.75, color: '#1D2939' },
    },
    verticalAlign: 'middle' as const,
    blocks: [
      {
        type: 'paragraph' as const,
        id: uid('p'),
        runs: [run(text, { bold })],
        style: { alignment: 'center' as const, lineHeight: 1.4 },
      },
    ],
  });
  return {
    type: 'table',
    id: uid('tbl'),
    columns: undefined,
    columnWidthsPx: [200, 160, 200],
    alignment: 'center',
    rows: [
      {
        id: uid('row'),
        isHeader: true,
        cells: [
          makeCell('项目', true, '#EEF2F9'),
          makeCell('金额（元）', true, '#EEF2F9'),
          makeCell('备注', true, '#EEF2F9'),
        ],
      },
      {
        id: uid('row'),
        cells: [makeCell('押金'), makeCell('36,000'), makeCell('相当于三个月租金')],
      },
      {
        id: uid('row'),
        cells: [makeCell('首期租金'), makeCell('36,000'), makeCell('签约后三日内支付')],
      },
    ],
  } as DocTable;
}

export function contractTemplate(): DocModel {
  const base = createBlankDoc();
  const blocks: DocBlock[] = [
    {
      type: 'paragraph',
      id: uid('p'),
      runs: [run('房屋租赁合同', { fontFamily: '黑体', fontSizePt: 22, bold: true })],
      style: { alignment: 'center', spaceAfterPx: 18 },
    },
    {
      type: 'paragraph',
      id: uid('p'),
      runs: [run('合同编号：MO-2026-______', { fontSizePt: 10.5 })],
      style: { alignment: 'right' },
    },
    paragraph('甲方（出租方）：____________________________', { firstLineIndentPx: 28 }),
    paragraph('乙方（承租方）：____________________________', { firstLineIndentPx: 28 }),
    paragraph(
      '根据《中华人民共和国民法典》及相关法律法规，甲乙双方在平等、自愿、协商一致的基础上，就房屋租赁事宜达成如下协议：',
      { firstLineIndentPx: 28 },
    ),
    heading('第一条　租赁房屋'),
    paragraph(
      '甲方将位于____________________________的房屋出租给乙方使用，建筑面积约为______平方米。房屋用途为办公，未经甲方书面同意，乙方不得擅自改变房屋用途。',
      { firstLineIndentPx: 28 },
    ),
    heading('第二条　租赁期限与租金'),
    paragraph(
      '租赁期限为______年，自______年____月____日起至______年____月____日止。月租金为人民币____________元整（¥________），按______支付。',
      { firstLineIndentPx: 28 },
    ),
    paragraph(
      '乙方应于每期租金到期前五日内支付下一期租金。逾期支付的，每逾期一日按应付金额的万分之五向甲方支付违约金。',
      { firstLineIndentPx: 28 },
    ),
    heading('第三条　押金与费用'),
    contractTable(),
    paragraph(
      '租赁期间产生的水费、电费、物业费、网络费等由乙方承担。租赁期满且乙方无违约行为的，甲方应在房屋交接完成后七日内无息退还押金。',
      { firstLineIndentPx: 28, spaceBeforePx: 8 },
    ),
    heading('第四条　房屋交付与返还'),
    paragraph(
      '甲方应于租赁期开始之日将房屋按约定条件交付乙方。乙方应于租赁期届满之日将房屋恢复原状并返还甲方，双方共同签署交接清单。',
      { firstLineIndentPx: 28 },
    ),
    heading('第五条　违约责任'),
    paragraph(
      '任何一方违反本合同约定，应承担相应的违约责任。因不可抗力导致合同无法履行的，双方互不承担违约责任。',
      { firstLineIndentPx: 28 },
    ),
    heading('第六条　争议解决'),
    paragraph(
      '本合同履行过程中发生争议的，双方应友好协商解决；协商不成的，任何一方可向房屋所在地人民法院提起诉讼。',
      { firstLineIndentPx: 28 },
    ),
    paragraph('', { spaceBeforePx: 24 }),
    paragraph('甲方（盖章）：____________________　　乙方（盖章）：____________________', {
      lineHeight: 2.4,
    }),
    paragraph('授权代表：________________________　　授权代表：________________________', {
      lineHeight: 2.4,
    }),
    paragraph('签订日期：______年____月____日　　　　签订日期：______年____月____日', {
      lineHeight: 2.4,
    }),
  ];
  return {
    ...base,
    blocks,
    meta: { title: '房屋租赁合同模板' },
    stats: { words: 0, characters: 0 },
  };
}

export function expenseTemplate(): WorkbookModel {
  const base = createBlankWorkbook();
  const sheet = base.sheets[0];
  sheet.name = '费用报销';
  sheet.cells = {
    '0:0': { value: '费用报销单', styleId: 'title' },
    '1:0': { value: '日期', styleId: 'header' },
    '1:1': { value: '类别', styleId: 'header' },
    '1:2': { value: '金额', styleId: 'header' },
    '1:3': { value: '备注', styleId: 'header' },
    '2:0': { value: 46313, styleId: 'date' },
    '2:1': { value: '交通费' },
    '2:2': { value: 128.5, styleId: 'money' },
    '2:3': { value: '客户拜访' },
    '3:0': { value: 46314, styleId: 'date' },
    '3:1': { value: '餐饮费' },
    '3:2': { value: 380, styleId: 'money' },
    '3:3': { value: '项目沟通' },
    '4:0': { value: 46315, styleId: 'date' },
    '4:1': { value: '办公用品' },
    '4:2': { value: 259.9, styleId: 'money' },
    '4:3': { value: '打印纸、墨盒' },
    '5:0': { value: 46316, styleId: 'date' },
    '5:1': { value: '差旅费' },
    '5:2': { value: 1200, styleId: 'money' },
    '5:3': { value: '高铁票' },
    '6:0': { value: '合计', styleId: 'total' },
    '6:2': { value: null, formula: '=SUM(C3:C6)', styleId: 'totalMoney' },
    '6:3': { value: '由财务审核后报销', styleId: 'total' },
    '8:0': { value: '说明：请附上原始发票，填写完成后提交财务。' },
  };
  sheet.merges = [{ r1: 0, c1: 0, r2: 0, c2: 3 }];
  sheet.freeze = { rows: 2, cols: 0 };
  sheet.colWidths = { 0: 100, 1: 100, 2: 110, 3: 200 };
  sheet.rowHeights = { 0: 34, 1: 24 };
  sheet.autoFilter = { r1: 1, c1: 0, r2: 5, c2: 3 };
  const styles: Record<string, CellStyle> = {
    default: { fontFamily: '等线', fontSizePt: 11, numFmt: 'General', verticalAlign: 'middle' },
    title: {
      fontFamily: '微软雅黑',
      fontSizePt: 18,
      bold: true,
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      color: '#1D2939',
    },
    header: {
      fontFamily: '等线',
      fontSizePt: 11,
      bold: true,
      fill: '#3478F6',
      color: '#FFFFFF',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      borders: {
        bottom: { style: 'thin', color: '#2F6EE0' },
      },
    },
    date: { fontFamily: '等线', fontSizePt: 11, numFmt: 'yyyy-mm-dd', horizontalAlign: 'center' },
    money: { fontFamily: '等线', fontSizePt: 11, numFmt: '¥#,##0.00', horizontalAlign: 'right' },
    total: {
      fontFamily: '等线',
      fontSizePt: 11,
      bold: true,
      fill: '#EEF2F9',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
    },
    totalMoney: {
      fontFamily: '等线',
      fontSizePt: 12,
      bold: true,
      numFmt: '¥#,##0.00',
      fill: '#EEF2F9',
      color: '#D92D20',
      horizontalAlign: 'right',
    },
  };
  const model: WorkbookModel = {
    ...base,
    sheets: [sheet],
    activeSheetId: sheet.id,
    styles,
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: { title: '费用报销单' },
    stats: { cells: Object.keys(sheet.cells).length, sheets: 1 },
  };
  return recalculateWorkbook(model);
}
