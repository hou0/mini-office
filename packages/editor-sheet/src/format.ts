import type { CellStyle, CellScalar } from '@mini-office/document-model';
import { isErrorString } from './formula/types';
import { serialToLocalParts } from './dates';

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

export function formatGeneralNumber(value: number): string {
  if (!Number.isFinite(value)) return '#NUM!';
  if (Object.is(value, -0)) return '0';
  if (Number.isInteger(value)) {
    if (Math.abs(value) >= 1e15) return value.toExponential(10).replace(/\.?0+e/, 'e');
    return String(value);
  }
  const abs = Math.abs(value);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    return value
      .toExponential(9)
      .replace(/(\.\d*?)0+e/, '$1e')
      .replace(/\.e/, 'e')
      .replace(/e\+?(-?)/, 'E$1');
  }
  return String(Number(value.toPrecision(15)));
}

export function isDateFormat(pattern: string | undefined): boolean {
  if (!pattern) return false;
  const cleaned = pattern.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  if (/[ydhs]/i.test(cleaned) || /[年月日时分秒]/.test(cleaned)) return true;
  return /^(?:m{1,4}|d{1,4})(?:[\/\-.年]m{1,4})?$/i.test(cleaned);
}

function pad(value: number, length = 2): string {
  return String(value).padStart(length, '0');
}

export function formatDateSerial(serial: number, pattern: string, date1904 = false): string {
  const parts = serialToLocalParts(serial, date1904);
  const hour12 = parts.hour % 12 || 12;
  const hasAmPm = /AM\/PM|A\/P/i.test(pattern);
  let result = '';
  let index = 0;
  let previousLetter = '';

  const nextTokenLetter = (position: number): string => {
    while (position < pattern.length) {
      const char = pattern[position];
      if (/[a-zA-Z]/.test(char)) return char.toLowerCase();
      if (char === '"') {
        const close = pattern.indexOf('"', position + 1);
        position = close < 0 ? pattern.length : close + 1;
        continue;
      }
      if (char === '[') {
        const close = pattern.indexOf(']', position + 1);
        position = close < 0 ? pattern.length : close + 1;
        continue;
      }
      position += 1;
    }
    return '';
  };

  while (index < pattern.length) {
    const char = pattern[index];
    if (char === '"') {
      const close = pattern.indexOf('"', index + 1);
      result += pattern.slice(index + 1, close < 0 ? pattern.length : close);
      index = close < 0 ? pattern.length : close + 1;
      continue;
    }
    if (char === '\\') {
      result += pattern[index + 1] ?? '';
      index += 2;
      continue;
    }
    if (char === '[') {
      const close = pattern.indexOf(']', index + 1);
      const content = pattern.slice(index + 1, close < 0 ? pattern.length : close);
      if (/^(h+|m+|s+)$/i.test(content)) {
        const unit = content[0].toLowerCase();
        if (unit === 'h') result += pad(Math.floor(serial * 24));
        else if (unit === 'm') result += pad(Math.floor(serial * 24 * 60));
        else result += pad(Math.floor(serial * 24 * 3600));
      } else {
        result += content;
      }
      index = close < 0 ? pattern.length : close + 1;
      continue;
    }
    if (/^AM\/PM|^A\/P/i.test(pattern.slice(index))) {
      const token = /^AM\/PM|^A\/P/i.exec(pattern.slice(index))![0];
      result += token.length === 5 ? (parts.hour < 12 ? 'AM' : 'PM') : parts.hour < 12 ? 'A' : 'P';
      index += token.length;
      previousLetter = 'a';
      continue;
    }
    if (!/[a-zA-Z]/.test(char)) {
      result += char;
      index += 1;
      continue;
    }
    let end = index + 1;
    while (end < pattern.length && pattern[end].toLowerCase() === char.toLowerCase()) end += 1;
    const length = end - index;
    const letter = char.toLowerCase();
    switch (letter) {
      case 'y':
        result += length >= 4 ? String(parts.year) : pad(parts.year % 100);
        break;
      case 'm': {
        const nextLetter = nextTokenLetter(end);
        const isMinute = previousLetter === 'h' || nextLetter === 's';
        if (isMinute) {
          result += length >= 2 ? pad(parts.minute) : String(parts.minute);
        } else if (length >= 4) result += MONTHS[parts.month - 1];
        else if (length === 3) result += MONTHS[parts.month - 1].slice(0, 3);
        else result += length === 2 ? pad(parts.month) : String(parts.month);
        break;
      }
      case 'd':
        if (length >= 4) result += WEEKDAYS[parts.weekday];
        else if (length === 3) result += WEEKDAYS[parts.weekday].slice(0, 3);
        else result += length === 2 ? pad(parts.day) : String(parts.day);
        break;
      case 'h':
        result += length >= 2 ? pad(hasAmPm ? hour12 : parts.hour) : String(hasAmPm ? hour12 : parts.hour);
        break;
      case 's':
        result += length >= 2 ? pad(parts.second) : String(parts.second);
        break;
      default:
        result += pattern.slice(index, end);
        break;
    }
    previousLetter = letter;
    index = end;
  }
  return result;
}
function splitSections(pattern: string): string[] {
  const sections: string[] = [];
  let current = '';
  let quoted = false;
  for (const char of pattern) {
    if (char === '"') quoted = !quoted;
    if (char === ';' && !quoted) {
      sections.push(current);
      current = '';
      continue;
    }
    current += char;
  }
  sections.push(current);
  return sections;
}

function countDecimals(section: string): number {
  const match = /\.([0#]+)/.exec(section);
  return match ? match[1].length : 0;
}

function applyThousands(value: string): string {
  const [integer, decimals] = value.split('.');
  const sign = integer.startsWith('-') ? '-' : '';
  const digits = sign ? integer.slice(1) : integer;
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return `${sign}${grouped}${decimals !== undefined ? `.${decimals}` : ''}`;
}

function formatScientific(value: number, section: string): string {
  const decimals = countDecimals(section);
  return value
    .toExponential(decimals)
    .replace(/e([+-])(\d+)/i, (_match, sign: string, digits: string) => `E${sign}${digits.padStart(2, '0')}`);
}

export function formatNumeric(value: number, pattern: string | undefined, date1904 = false): string {
  if (!Number.isFinite(value)) return '#NUM!';
  if (!pattern || pattern === 'General' || pattern === 'general') return formatGeneralNumber(value);
  // 功能：文本格式保留数值原貌，不按整数四舍五入；时间：2026-10-24；作者：lq。
  if (pattern === '@') return formatGeneralNumber(value);
  if (isDateFormat(pattern)) return formatDateSerial(value, pattern, date1904);
  const sections = splitSections(pattern);
  let section = sections[0] ?? 'General';
  if (value < 0 && sections[1]) section = sections[1];
  else if (value === 0 && sections[2]) section = sections[2];
  // 功能：负数格式使用绝对值套用负数节，避免会计格式出现重复负号；时间：2026-10-24；作者：lq。
  const negative = value < 0;
  const absolute = Math.abs(value);
  const fraction = /([?0#]+)\/([?0#]+)/.exec(section);
  if (fraction) {
    const whole = Math.floor(absolute);
    const maxDenominator = Math.min(999, 10 ** fraction[2].length - 1);
    let numerator = 0; let denominator = 1; let best = Infinity;
    for (let candidate = 1; candidate <= maxDenominator; candidate++) {
      const current = Math.round((absolute - whole) * candidate);
      const error = Math.abs(absolute - whole - current / candidate);
      if (error < best) { best = error; numerator = current; denominator = candidate; }
    }
    const integer = whole + (numerator === denominator ? 1 : 0);
    const remainder = numerator === denominator ? 0 : numerator;
    return `${negative ? '-' : ''}${integer || !remainder ? integer : ''}${integer && remainder ? ' ' : ''}${remainder ? `${remainder}/${denominator}` : ''}`;
  }
  if (/E[+-]0+/i.test(section)) return formatScientific(absolute, section);
  const percent = section.includes('%');
  const digits = Math.min(20, countDecimals(section));
  const useThousands = section.includes(',');
  const currencyMatch = /[¥￥$€£]/.exec(section);
  const displayValue = percent ? absolute * 100 : absolute;
  let text = displayValue.toFixed(digits);
  if (useThousands) text = applyThousands(text);
  if (section.includes('(') && negative) {
    text = `(${text})`;
  } else if (negative) {
    text = `-${text}`;
  }
  if (currencyMatch) text = text.startsWith('(') ? `(${currencyMatch[0]}${text.slice(1)}` : `${currencyMatch[0]}${text}`;
  if (percent) text = `${text}%`;
  return text.replace(/"([^"]*)"/g, '$1');
}

export function formatCellValue(
  value: CellScalar | string,
  style: CellStyle | undefined,
  date1904 = false,
): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') {
    if (isErrorString(value)) return value;
    return value;
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return formatNumeric(value, style?.numFmt, date1904);
}

export const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'yyyy-MM-dd',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'yyyy/M/d h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

export function normalizeNumberFormat(pattern: string | undefined): string {
  if (!pattern) return 'General';
  const lower = pattern.toLowerCase();
  if (lower === 'general') return 'General';
  return pattern;
}
