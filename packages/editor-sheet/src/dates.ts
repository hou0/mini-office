const DAY_MS = 86_400_000;

export function dateToSerial(date: Date, date1904 = false): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  let serial = Math.round((utc - base) / DAY_MS);
  // Excel's fictional 1900-02-29 (serial 60) shifts early dates by one day.
  if (!date1904 && utc < Date.UTC(1900, 2, 1)) serial -= 1;
  return serial;
}

export function serialToDate(serial: number, date1904 = false): Date {
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  let adjusted = Math.floor(serial);
  if (!date1904) {
    if (adjusted === 60) adjusted = 59;
    if (adjusted < 60) adjusted += 1;
  }
  return new Date(base + adjusted * DAY_MS);
}

export function serialToLocalParts(serial: number, date1904 = false): {
  year: number;
  month: number;
  day: number;
  weekday: number;
  hour: number;
  minute: number;
  second: number;
} {
  const total = serial * DAY_MS;
  const days = Math.floor(serial);
  const fraction = serial - days;
  const date = serialToDate(days, date1904);
  const totalSeconds = Math.round(fraction * 86_400);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday: date.getUTCDay(),
    hour: Math.floor(totalSeconds / 3600) % 24,
    minute: Math.floor(totalSeconds / 60) % 60,
    second: totalSeconds % 60,
  };
}

export function todaySerial(now: Date, date1904 = false): number {
  return dateToSerial(new Date(now.getFullYear(), now.getMonth(), now.getDate()), date1904);
}

export function nowSerial(now: Date, date1904 = false): number {
  const midnight = dateToSerial(
    new Date(now.getFullYear(), now.getMonth(), now.getDate()),
    date1904,
  );
  const seconds = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
  return midnight + seconds / 86_400;
}

export function makeDateSerial(year: number, month: number, day: number, date1904 = false): number {
  const normalizedYear = year < 1900 ? year + 1900 : year;
  const date = new Date(Date.UTC(normalizedYear, month - 1, day));
  return dateToSerial(
    new Date(
      date.getUTCFullYear(),
      date.getUTCMonth(),
      date.getUTCDate(),
    ),
    date1904,
  );
}

export function addMonths(serial: number, months: number, date1904 = false): number {
  const parts = serialToLocalParts(serial, date1904);
  const date = new Date(Date.UTC(parts.year, parts.month - 1 + months, 1));
  const lastDay = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate();
  date.setUTCDate(Math.min(parts.day, lastDay));
  return dateToSerial(new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()), date1904);
}

export function startOfMonthSerial(serial: number, date1904 = false): number {
  const parts = serialToLocalParts(serial, date1904);
  return makeDateSerial(parts.year, parts.month, 1, date1904);
}

export function endOfMonthSerial(serial: number, date1904 = false): number {
  const parts = serialToLocalParts(serial, date1904);
  const lastDay = new Date(Date.UTC(parts.year, parts.month, 0)).getUTCDate();
  return makeDateSerial(parts.year, parts.month, lastDay, date1904);
}

export function parseDateInput(value: string): Date | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(trimmed);
  if (iso) {
    const date = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const cn = /^(\d{4})年(\d{1,2})月(\d{1,2})日$/.exec(trimmed);
  if (cn) {
    const date = new Date(Number(cn[1]), Number(cn[2]) - 1, Number(cn[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const parsed = new Date(trimmed);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}
