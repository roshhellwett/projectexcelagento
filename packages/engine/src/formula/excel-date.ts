import { isFormulaError, type FormulaErrorCode } from './errors.js';

/**
 * Excel date serials.
 *
 * Every date in a spreadsheet is stored as a count of days since an epoch, and which
 * epoch depends on the workbook: the 1900 system (the overwhelming default) or the
 * legacy 1904 system. This module is the single source of truth for that conversion so
 * the formula engine, the grid renderer, and `format_dates` cannot drift apart.
 *
 * The 1900 system carries Excel's own bug: it believes 1900 was a leap year, so serial
 * 60 is a day that never existed (1900-02-29) and everything after it is offset by one.
 * Serial 60 is therefore rejected rather than silently converted to 1900-03-01.
 */

export type DateSystem = '1900' | '1904';

const MS_PER_DAY = 86_400_000;

/** Days from the Unix epoch to Excel serial 1 (1900-01-01) in the 1900 system. */
const EPOCH_1900_PRE_LEAP_BUG = 25_568;
/** Days from the Unix epoch to Excel serial 61 (1900-03-01), i.e. after the 1900 leap-year bug. */
const EPOCH_1900_POST_LEAP_BUG = 25_569;
/** Days from the Unix epoch to 1904-01-01, which is serial 0 in the 1904 system. */
const EPOCH_1904 = 24_107;

/** The fictitious 1900-02-29 that Excel's 1900 date system believes in. */
export const FICTITIOUS_SERIAL_60 = 60;

const MAX_1900_SERIAL = 2_958_465; // 9999-12-31
const MIN_1900_SERIAL = 1;
const MAX_1904_SERIAL = 2_957_003; // 9999-12-31
const MIN_1904_SERIAL = 0;

export function isExcelSerial(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

/**
 * Converts an Excel serial to a UTC `Date`, or returns `null` when the serial does not
 * name a real day (serial 60, out-of-range, or non-finite).
 */
export function excelSerialToDate(serial: number, dateSystem: DateSystem = '1900'): Date | null {
  if (!Number.isFinite(serial)) return null;

  if (dateSystem === '1904') {
    if (serial < MIN_1904_SERIAL || serial > MAX_1904_SERIAL) return null;
    const wholeDays = Math.floor(serial);
    const fraction = serial - wholeDays;
    return new Date(Math.round((wholeDays - EPOCH_1904) * MS_PER_DAY + fraction * MS_PER_DAY));
  }

  if (serial < MIN_1900_SERIAL || serial > MAX_1900_SERIAL) return null;
  if (serial === FICTITIOUS_SERIAL_60) return null;

  const epochOffset =
    serial < FICTITIOUS_SERIAL_60 ? EPOCH_1900_PRE_LEAP_BUG : EPOCH_1900_POST_LEAP_BUG;
  const wholeDays = Math.floor(serial);
  const fraction = serial - wholeDays;
  return new Date(Math.round((wholeDays - epochOffset) * MS_PER_DAY + fraction * MS_PER_DAY));
}

/**
 * Converts a `Date` to an Excel serial, or returns `null` when the date cannot be
 * represented in the given system.
 */
export function dateToExcelSerial(date: Date, dateSystem: DateSystem = '1900'): number | null {
  const time = date.getTime();
  if (!Number.isFinite(time)) return null;

  const rawDays = Math.floor(time / MS_PER_DAY);
  const fraction = time - rawDays * MS_PER_DAY;

  if (dateSystem === '1904') {
    const serial = rawDays + EPOCH_1904 + fraction / MS_PER_DAY;
    return serial < MIN_1904_SERIAL || serial > MAX_1904_SERIAL ? null : serial;
  }

  // Before 1900-03-01 the 1900 system is one day ahead of a true day count.
  const serial =
    rawDays +
    (time < Date.UTC(1900, 2, 1) ? EPOCH_1900_PRE_LEAP_BUG : EPOCH_1900_POST_LEAP_BUG) +
    fraction / MS_PER_DAY;
  if (serial < MIN_1900_SERIAL) return serial >= 0 ? serial : null;
  return serial === FICTITIOUS_SERIAL_60 ? null : serial;
}

// Unambiguous textual date shapes only. Anything that could be read as either
// MM/DD or DD/MM is refused rather than guessed - a wrong date is worse than no date.
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;
const SLASHED_ISO_DATE = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/;
const ISO_DATE_TIME = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?$/;

function buildUtcDate(
  year: number,
  month: number,
  day: number,
  hours = 0,
  minutes = 0,
  seconds = 0,
  millis = 0,
): Date | null {
  const time = Date.UTC(year, month - 1, day, hours, minutes, seconds, millis);
  const date = new Date(time);
  // Rejects impossible calendar dates such as 2021-02-30, which Date.UTC would roll over.
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return date;
}

/**
 * Parses an unambiguous textual date. Deliberately rejects `01/02/2020`-style input,
 * which Excel itself also treats as ambiguous, and never feeds a bare numeric string to
 * `Date` - JavaScript reads `"43890"` as the *year* 43890, which is how a real Excel
 * serial turns into a forty-thousand-year-old date.
 */
export function parseUnambiguousDate(text: string): Date | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;

  const isoDateTime = ISO_DATE_TIME.exec(trimmed);
  if (isoDateTime) {
    const [, y, mo, d, h, mi, s, ms] = isoDateTime;
    return buildUtcDate(
      Number(y),
      Number(mo),
      Number(d),
      Number(h),
      Number(mi),
      s ? Number(s) : 0,
      ms ? Number(ms.padEnd(3, '0')) : 0,
    );
  }

  const iso = ISO_DATE.exec(trimmed);
  if (iso) {
    const [, y, mo, d] = iso;
    return buildUtcDate(Number(y), Number(mo), Number(d));
  }

  const slashed = SLASHED_ISO_DATE.exec(trimmed);
  if (slashed) {
    const [, y, mo, d] = slashed;
    return buildUtcDate(Number(y), Number(mo), Number(d));
  }

  return null;
}

export interface CoerceDateOptions {
  dateSystem?: DateSystem;
}

/**
 * Best-effort conversion of a formula value to a UTC `Date`.
 *
 * Numbers are Excel serials, `Date` values pass through, and strings must be
 * unambiguous. Anything else yields `null` so the caller can raise `#VALUE!` rather
 * than inventing a date.
 */
export function coerceToDate(value: unknown, options: CoerceDateOptions = {}): Date | null {
  const dateSystem = options.dateSystem ?? '1900';
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value : null;
  if (typeof value === 'number') return excelSerialToDate(value, dateSystem);
  if (typeof value === 'string') return parseUnambiguousDate(value);
  return null;
}

/**
 * Coerces to a date or returns the Excel error the coercion deserves. A numeric-looking
 * string is `#VALUE!` because Excel treats it as text, not as a serial.
 */
export function coerceToDateOrError(
  value: unknown,
  options: CoerceDateOptions = {},
): Date | FormulaErrorCode {
  if (isFormulaError(value)) return value;
  const date = coerceToDate(value, options);
  return date ?? '#VALUE!';
}

/**
 * Excel's `DATE` semantics: out-of-range months and days roll over rather than erroring,
 * so `DATE(2020, 13, 1)` is 2021-01-01 and `DATE(2020, 1, 32)` is 2020-02-01.
 */
export function excelDate(year: number, month: number, day: number): Date | null {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  const normalizedYear = year + Math.floor((month - 1) / 12);
  const normalizedMonth = ((((month - 1) % 12) + 12) % 12) + 1;
  const time = Date.UTC(normalizedYear, normalizedMonth - 1, day);
  const result = new Date(time);
  if (!Number.isFinite(result.getTime())) return null;
  return result;
}

/** Whole days between two dates, matching Excel's subtraction semantics. */
export function daysBetween(later: Date, earlier: Date): number {
  const laterDay = Date.UTC(later.getUTCFullYear(), later.getUTCMonth(), later.getUTCDate());
  const earlierDay = Date.UTC(
    earlier.getUTCFullYear(),
    earlier.getUTCMonth(),
    earlier.getUTCDate(),
  );
  return Math.round((laterDay - earlierDay) / MS_PER_DAY);
}

/** Renders a date as `YYYY-MM-DD` without dragging the whole Date through the locale layer. */
export function formatIsoDate(date: Date): string {
  const year = String(date.getUTCFullYear()).padStart(4, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Strips the parts of an Excel number format that carry no type information: quoted
 * literals, backslash escapes, `_x`/`*x` placeholders, and bracketed sections. Elapsed-time
 * codes (`[h]`, `[mm]`, `[ss]`) survive because they genuinely are date codes.
 */
function stripFormatLiterals(format: string): string {
  let out = '';
  let i = 0;
  while (i < format.length) {
    const ch = format[i]!;
    if (ch === '"') {
      i += 1;
      while (i < format.length && format[i] !== '"') {
        if (format[i] === '\\') i += 1;
        i += 1;
      }
      i += 1;
      continue;
    }
    if (ch === '\\') {
      i += 2;
      continue;
    }
    if (ch === '[') {
      const close = format.indexOf(']', i);
      if (close === -1) break;
      const inner = format.slice(i + 1, close);
      if (/^h+$/i.test(inner) || /^m+$/i.test(inner) || /^s+$/i.test(inner)) out += inner;
      i = close + 1;
      continue;
    }
    if (ch === '_' || ch === '*') {
      i += 2;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

/**
 * Decides whether a number format renders its cell as a date or a time.
 *
 * This is the single test that decides whether an imported numeric cell is a date or a plain
 * number, and therefore whether `=A2-A1` yields a day count or a meaningless subtraction. A
 * format containing numeric placeholders, currency, or percent is never a date format - `m`
 * only means "month" when it is not competing with a numeric placeholder.
 */
export function isDateNumberFormat(format: string | null | undefined): boolean {
  if (!format) return false;
  const cleaned = stripFormatLiterals(format);
  if (cleaned.length === 0) return false;
  // Fractional seconds are numeric placeholders that belong to a time format (`mm:ss.0`), so
  // they are removed before asking whether the rest of the format is purely numeric.
  const withoutFraction = cleaned.replace(/\.[#0?]{1,3}(?![#0?])/g, '');
  if (/[#0?]/.test(withoutFraction)) return false;
  if (/[$%]/.test(cleaned)) return false;
  if (/\\/.test(cleaned)) return false;
  return /[ymdhs]/i.test(cleaned);
}
