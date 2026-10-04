import {
  excelSerialToDate,
  formatIsoDate,
  isDateNumberFormat,
  type DateSystem,
} from '@excel-agent/engine';

/**
 * How a grid cell's date is drawn.
 *
 * All epoch arithmetic lives in the engine's `excel-date.ts`: this module only decides what the
 * text should say. A serial is an ordinary number until its number format says otherwise, and a
 * workbook from legacy Mac Excel counts days from 1904 instead of 1900 - so both the value and the
 * format have to be consulted before a cell renders as a date.
 */

/** Elapsed-time codes: `[h]`, `[mm]`, `[ss]` are running totals that may pass midnight. */
const ELAPSED_HOURS = /\[h+\]/i;
const ELAPSED_MINUTES = /\[m+\]/i;
const ELAPSED_SECONDS = /\[s+\]/i;

/** Three or more `m`s is a month name in Excel; minutes never repeat that often. */
const MONTH_NAME = /m{3,}/i;
const CLOCK_HOURS = /h{1,2}/i;
const CLOCK_SECONDS = /(^|[^a-z0-9])s{1,2}([^a-z0-9]|$)/i;
const AM_PM = /a\/p/i;
/** Year and day codes, or a month name, mean the format wants the calendar day as well as a clock. */
const YEAR_OR_DAY = /[yd]/i;

const SECONDS_PER_DAY = 86_400;

const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

/**
 * Drops the parts of a format that only decorate it - quoted literals, backslash escapes, and
 * bracketed locale or colour sections - so `[$-409]mmm d, yyyy;@` is read as date codes rather
 * than as a locale tag, and `[Red]dd/mm/yyyy` as a date rather than as a colour.
 */
function codeSection(format: string): string {
  return format
    .replace(/"(?:[^"]|\\.)*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '');
}

function pad(value: number): string {
  return String(value).padStart(2, '0');
}

function hasElapsedCodes(format: string): boolean {
  return ELAPSED_HOURS.test(format) || ELAPSED_MINUTES.test(format) || ELAPSED_SECONDS.test(format);
}

function hasCalendarCodes(format: string): boolean {
  return YEAR_OR_DAY.test(format) || MONTH_NAME.test(format);
}

/**
 * Elapsed time such as `[h]:mm:ss`, where the hours are a running total rather than a clock
 * reading. The total comes from the serial's fractional day, which is what keeps a duration
 * independent of the calendar day it happens to sit on.
 */
function formatElapsedTime(elapsedSeconds: number, format: string): string {
  // Bracketed codes are dropped here so `[h]` does not read as its own seconds field, leaving
  // only the unbracketed `:ss` that decides whether seconds are shown.
  const showSeconds = /s/i.test(codeSection(format));
  const totalMinutes = Math.floor(elapsedSeconds / 60);
  const seconds = elapsedSeconds % 60;

  if (ELAPSED_HOURS.test(format)) {
    const hours = Math.floor(totalMinutes / 60);
    return showSeconds
      ? `${hours}:${pad(totalMinutes % 60)}:${pad(seconds)}`
      : `${hours}:${pad(totalMinutes % 60)}`;
  }
  if (ELAPSED_MINUTES.test(format)) {
    return showSeconds ? `${totalMinutes}:${pad(seconds)}` : String(totalMinutes);
  }
  return String(elapsedSeconds);
}

function formatClockTime(date: Date, format: string): string {
  const clock = `${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`;
  return CLOCK_SECONDS.test(format) ? `${clock}:${pad(date.getUTCSeconds())}` : clock;
}

/**
 * A month-name format (`mmm d, yyyy`, the shape Excel writes for header cells) reads better as
 * prose than as digits; every other date format keeps the unambiguous ISO form.
 */
function formatCalendarDate(date: Date, format: string): string {
  if (!MONTH_NAME.test(format)) return formatIsoDate(date);
  const month = MONTH_NAMES[date.getUTCMonth()] ?? String(date.getUTCMonth() + 1);
  return `${month} ${date.getUTCDate()}, ${String(date.getUTCFullYear()).padStart(4, '0')}`;
}

interface ResolvedDate {
  date: Date;
  /** Seconds past midnight, taken from the serial's fraction so elapsed codes can total them. */
  elapsedSeconds: number;
}

/**
 * Turns a cell value into the date it stands for, or `null` when the value is not a date this
 * cell should display as one.
 *
 * A `Date` is already a date - the importer writes real `Date` values, so it needs no format to
 * vouch for it. A bare number stays a number unless its format is a date format, which is
 * Excel's own rule and the reason a price in a date-looking column does not become a year.
 */
function resolveDateValue(
  value: unknown,
  numberFormat: string | undefined,
  dateSystem: DateSystem,
): ResolvedDate | null {
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime())) return null;
    return {
      date: value,
      elapsedSeconds:
        value.getUTCHours() * 3600 + value.getUTCMinutes() * 60 + value.getUTCSeconds(),
    };
  }
  if (typeof value !== 'number' || !isDateNumberFormat(numberFormat)) return null;
  const date = excelSerialToDate(value, dateSystem);
  if (!date) return null;
  return { date, elapsedSeconds: Math.round((value - Math.floor(value)) * SECONDS_PER_DAY) };
}

/**
 * Formats a cell value as a date or time, or returns `null` when the caller should render it some
 * other way. Time codes bring the time of day along; a format without them shows the day alone.
 */
export function formatCellDateValue(
  value: unknown,
  numberFormat?: string,
  dateSystem: DateSystem = '1900',
): string | null {
  const resolved = resolveDateValue(value, numberFormat, dateSystem);
  if (!resolved) return null;
  return renderDate(resolved, numberFormat);
}

function renderDate({ date, elapsedSeconds }: ResolvedDate, numberFormat?: string): string {
  if (!numberFormat) return formatIsoDate(date);
  if (hasElapsedCodes(numberFormat)) return formatElapsedTime(elapsedSeconds, numberFormat);

  const codes = codeSection(numberFormat);
  const calendar = formatCalendarDate(date, codes);
  const hasClock = CLOCK_HOURS.test(codes) || CLOCK_SECONDS.test(codes) || AM_PM.test(codes);
  // A time-only format (`mm:ss`, `h:mm AM/PM`) shows the clock alone, exactly as Excel does.
  if (hasClock && !hasCalendarCodes(codes)) return formatClockTime(date, codes);
  return hasClock ? `${calendar} ${formatClockTime(date, codes)}` : calendar;
}
