import type { WorkSheet, WorkBook } from 'xlsx-js-style';
import {
  createCell,
  createWorkbookValueReader,
  dateToExcelSerial,
  isFormulaError,
  excelSerialToDate,
  isDateNumberFormat,
  type Cell,
  type CellStyle,
  type CellValue,
  type Workbook,
  type Sheet,
} from '@excel-agent/engine';

type XlsxModule = typeof import('xlsx-js-style');

/** Guard against pathological sheets that would freeze the browser tab. */
export const MAX_IMPORTED_CELLS = 1_500_000;
export const MAX_EXCEL_ROWS = 1_048_576;
export const MAX_EXCEL_COLUMNS = 16_384;

export interface ImportReport {
  /** How the bytes were interpreted, so the user is never guessing at what was loaded. */
  source: 'xlsx' | 'xls' | 'csv';
  encoding: string;
  delimiter: string;
  dateSystem: '1900' | '1904';
  cellsConvertedToDates: number;
  /**
   * Features present in the file that this workspace cannot represent. Reported rather than
   * swallowed: silently flattening a merged title row leaves the user exporting a mangled
   * file with no idea anything was lost.
   */
  dropped: string[];
}

/** What the file turned out to be, decided from its bytes rather than its extension. */
export function classifyWorkbookBytes(bytes: Uint8Array): 'xlsx' | 'xls' | 'csv' {
  if (
    bytes.length > 3 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  ) {
    return 'xlsx';
  }
  if (
    bytes.length > 7 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  ) {
    return 'xls';
  }
  return 'csv';
}

/**
 * Decodes text bytes without ever losing a character.
 *
 * A UTF-8 `TextDecoder` is non-fatal by default, so Latin-1 bytes become U+FFFD and the
 * damage is unrecoverable by the time anyone notices. A UTF-16 CSV - still the default
 * export of Excel itself - arrives as interleaved NUL bytes and loads as garbage unless the
 * byte-order mark is honoured.
 */
export function decodeTextBytes(bytes: Uint8Array): { text: string; encoding: string } {
  if (bytes.length >= 2) {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) {
      return {
        text: stripBom(new TextDecoder('utf-16le').decode(bytes), '﻿'),
        encoding: 'utf-16le',
      };
    }
    if (bytes[0] === 0xfe && bytes[1] === 0xff) {
      return {
        text: stripBom(new TextDecoder('utf-16be').decode(bytes), '﻿'),
        encoding: 'utf-16be',
      };
    }
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
      return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8' };
    }
  }

  // No BOM. UTF-16 without one still shows itself as a dense run of NUL bytes. In little-endian
  // UTF-16 an ASCII character is followed by NUL, so the NULs land on the odd byte offsets.
  const sample = bytes.subarray(0, Math.min(bytes.length, 4096));
  let nulls = 0;
  for (const byte of sample) if (byte === 0) nulls += 1;
  if (sample.length > 0 && nulls / sample.length > 0.3) {
    const encoding =
      countNullsAtOddOffsets(sample) > countNullsAtEvenOffsets(sample) ? 'utf-16le' : 'utf-16be';
    return { text: new TextDecoder(encoding).decode(bytes), encoding };
  }

  // Strict UTF-8 first; only fall back to a single-byte encoding if that genuinely fails.
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'utf-8' };
  } catch {
    return { text: new TextDecoder('windows-1252').decode(bytes), encoding: 'windows-1252' };
  }
}

function stripBom(text: string, bom: string): string {
  return text.startsWith(bom) ? text.slice(bom.length) : text;
}

function countNullsAtEvenOffsets(bytes: Uint8Array): number {
  let count = 0;
  for (let i = 0; i < bytes.length; i += 2) if (bytes[i] === 0) count += 1;
  return count;
}

function countNullsAtOddOffsets(bytes: Uint8Array): number {
  let count = 0;
  for (let i = 1; i < bytes.length; i += 2) if (bytes[i] === 0) count += 1;
  return count;
}

const CANDIDATE_DELIMITERS = [',', ';', '\t', '|'] as const;

/**
 * Splits CSV text into an array of records (rows), where each record is an array of string fields.
 * Follows the RFC 4180 standard:
 * - Honors double quotes and `""` escapes
 * - Preserves line breaks inside quoted fields without prematurely splitting records
 * - Only starts a new record on newline `\n` or `\r\n` outside quotes
 * - If delimiter is null, splits only on unquoted newlines into 1-field records
 */
export function parseCsvRecords(
  text: string,
  delimiter: string | null,
  maxRecords = Infinity,
  requireClosedQuotes = true,
): string[][] {
  const records: string[][] = [];
  let currentRecord: string[] = [];
  let currentField = '';
  let inQuotes = false;
  const len = text.length;

  for (let i = 0; i < len; i += 1) {
    const ch = text[i]!;

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < len && text[i + 1] === '"') {
          currentField += '"';
          i += 1; // skip escaped quote
        } else {
          inQuotes = false;
        }
      } else {
        currentField += ch;
      }
      continue;
    }

    if (ch === '"' && currentField === '') {
      inQuotes = true;
    } else if (delimiter !== null && ch === delimiter) {
      currentRecord.push(currentField);
      currentField = '';
    } else if (ch === '\r') {
      if (i + 1 < len && text[i + 1] === '\n') {
        i += 1;
      }
      currentRecord.push(currentField);
      currentField = '';
      records.push(currentRecord);
      currentRecord = [];
      if (records.length >= maxRecords) return records;
    } else if (ch === '\n') {
      currentRecord.push(currentField);
      currentField = '';
      records.push(currentRecord);
      currentRecord = [];
      if (records.length >= maxRecords) return records;
    } else {
      currentField += ch;
    }
  }

  if (inQuotes && requireClosedQuotes) {
    throw new Error('CSV contains an unterminated quoted field.');
  }

  // Push remaining field and record if present
  if (currentField !== '' || currentRecord.length > 0) {
    currentRecord.push(currentField);
    records.push(currentRecord);
  }

  // Remove trailing blank record if caused by trailing newline
  while (
    records.length > 0 &&
    records[records.length - 1]!.length === 1 &&
    records[records.length - 1]![0]!.trim() === ''
  ) {
    records.pop();
  }

  return records;
}

/** Splits one CSV line, honouring quoted fields and `""` escapes. A null delimiter keeps the line whole. */
export function splitCsvLine(line: string, delimiter: string | null): string[] {
  if (delimiter === null) return [line];
  const records = parseCsvRecords(line, delimiter, 1);
  return records[0] ?? [''];
}

function countFieldsPerRecord(text: string, delimiter: string): number[] {
  // Delimiter detection uses a bounded prefix which may end inside a perfectly valid field.
  const records = parseCsvRecords(text, delimiter, 200, false);
  return records.map((r) => r.length);
}

/**
 * Picks the delimiter that produces a consistent column count, or returns null when
 * the file is genuinely single-column.
 *
 * Counting raw occurrences is not enough: a single-column file of currency amounts contains as
 * many commas as a three-column file has separators, and splitting on them turns `$1,200.50`
 * into two cells. So a candidate must explain the shape of the lines before it wins.
 */
export function detectDelimiter(text: string): string | null {
  const sample = text.slice(0, 64 * 1024);
  let best: string | null = null;
  let bestScore = -1;

  for (const candidate of CANDIDATE_DELIMITERS) {
    const counts = countFieldsPerRecord(sample, candidate);
    if (counts.length < 2) continue;

    const frequency = new Map<number, number>();
    for (const count of counts) frequency.set(count, (frequency.get(count) ?? 0) + 1);

    // The modal column count, preferring the smaller count on a tie so a file that is mostly
    // single-column is not split by a stray separator in one row.
    let modal = 0;
    let modalFrequency = 0;
    for (const [count, hits] of frequency) {
      if (hits > modalFrequency || (hits === modalFrequency && (modal === 0 || count < modal))) {
        modal = count;
        modalFrequency = hits;
      }
    }

    if (modal < 2) continue;
    const consistency = modalFrequency / counts.length;
    if (consistency < 0.6) {
      const firstRowCols = counts[0] ?? 0;
      if (consistency < 0.4 || firstRowCols < 2) continue;
    }

    const score = consistency * 1000 + modal;
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  // A header-only file still has a shape. Count parsed fields rather than raw separators:
  // commas inside a quoted single-column value must never become column boundaries.
  if (!best) {
    for (const candidate of CANDIDATE_DELIMITERS) {
      const records = parseCsvRecords(sample, candidate, 2, false);
      if (records.length === 1 && records[0]!.length > bestScore) {
        bestScore = records[0]!.length;
        if (bestScore > 1) best = candidate;
      }
    }
  }

  return best;
}

/**
 * Converts a CSV field to a number only when doing so is unambiguous.
 *
 * Values with significant leading zeros are kept as text on purpose: a ZIP code, an account
 * number, or a product code that becomes `123` from `00123` is corrupted data, and no user
 * downstream can recover it.
 */
export function parseCsvField(raw: string, delimiter: string): CellValue {
  const text = raw.trim();
  if (text === '') return null;
  if (text.startsWith('=')) return raw;

  // Accounting notation: (450) means negative 450.
  const negativeParenthesised = /^\((.*)\)$/.exec(text);
  const body = negativeParenthesised ? negativeParenthesised[1]! : text;
  const negative = negativeParenthesised !== null;

  const isPercent = body.endsWith('%');
  const numericBody = isPercent ? body.slice(0, -1).trim() : body;

  // A leading zero marks an identifier. Only "0" itself and "0.5"-style values are numbers.
  if (/^[-+]?0\d/.test(numericBody)) return raw;

  const magnitude = stripNumericDecoration(numericBody, delimiter);
  if (magnitude === null) return raw;

  const parsed = Number(magnitude);
  if (!Number.isFinite(parsed) || (Number.isInteger(parsed) && !Number.isSafeInteger(parsed))) {
    return raw;
  }

  let value = negative ? -parsed : parsed;
  // Applied last, so "45%" and "45" take the same route through the parser.
  if (isPercent) value /= 100;
  return value;
}

function stripNumericDecoration(body: string, delimiter: string): string | null {
  // Currency symbols and spaces are decoration; a leading sign is not.
  const cleaned = body.replace(/[\s\u00a0]/g, '').replace(/^[$€£¥₹]/, '');
  if (cleaned === '') return null;
  if (!/^[-+]?[\d.,]+(?:[eE][-+]?\d+)?$/.test(cleaned)) return null;
  if (/^[-+]?0\d/.test(cleaned)) return null;

  // Scientific notation is unambiguous only with a plain decimal mantissa.
  if (/[eE]/.test(cleaned)) {
    return /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)[eE][-+]?\d+$/.test(cleaned) ? cleaned : null;
  }

  const signed = /^[-+]/.test(cleaned) ? cleaned[0]! : '';
  const magnitude = signed ? cleaned.slice(1) : cleaned;

  const hasComma = magnitude.includes(',');
  const hasDot = magnitude.includes('.');

  if (hasComma && hasDot) {
    if (/^\d{1,3}(?:,\d{3})+\.\d+$/.test(magnitude)) return signed + magnitude.replace(/,/g, '');
    if (/^\d{1,3}(?:\.\d{3})+,\d+$/.test(magnitude)) {
      return signed + magnitude.replace(/\./g, '').replace(',', '.');
    }
    return null;
  }
  if (hasComma) {
    if (delimiter === ',' || /^\d{1,3}(?:,\d{3}){2,}$/.test(magnitude)) {
      return /^\d{1,3}(?:,\d{3})+$/.test(magnitude) ? signed + magnitude.replace(/,/g, '') : null;
    }
    return /^\d+,\d+$/.test(magnitude) ? signed + magnitude.replace(',', '.') : null;
  }
  if (hasDot) {
    // A single dot is a decimal point, regardless of how many fractional digits it carries.
    return /^(?:\d+(?:\.\d*)?|\.\d+)$/.test(magnitude) ? signed + magnitude : null;
  }
  return signed + magnitude;
}

/** Builds a workbook from decoded CSV text, keeping values as faithfully as the text allows. */
export function workbookFromCsvText(
  text: string,
  delimiter: string | null,
  sheetName: string,
): Sheet {
  const records = parseCsvRecords(text, delimiter);
  if (records.length === 0) {
    return { name: sheetName, rows: [] };
  }

  // Calculate maximum column count across all records to normalize ragged rows
  let maxCols = 0;
  for (const r of records) {
    if (r.length > maxCols) maxCols = r.length;
  }
  if (records.length > MAX_EXCEL_ROWS) {
    throw new Error(`CSV contains more than Excel's ${MAX_EXCEL_ROWS} row limit.`);
  }
  if (maxCols > MAX_EXCEL_COLUMNS) {
    throw new Error(`CSV contains more than Excel's ${MAX_EXCEL_COLUMNS} column limit.`);
  }
  if (records.length * maxCols > MAX_IMPORTED_CELLS) {
    throw new Error(`CSV contains more than the ${MAX_IMPORTED_CELLS} cell import limit.`);
  }

  const effectiveDelimiter = delimiter ?? ',';
  const rows: Cell[][] = records.map((record) => {
    const rowCells: Cell[] = record.map((field) =>
      createCell(parseCsvField(field, effectiveDelimiter)),
    );
    while (rowCells.length < maxCols) {
      rowCells.push(createCell(null));
    }
    return rowCells;
  });

  return { name: sheetName, rows };
}

function rgbColor(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const hex = value.replace(/^#/, '').slice(-6);
  return /^[0-9a-f]{6}$/i.test(hex) ? `#${hex.toUpperCase()}` : undefined;
}

function styleFromSheetJs(raw: unknown): CellStyle | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const style = raw as {
    patternType?: string;
    fgColor?: { rgb?: string };
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean | string;
      color?: { rgb?: string };
    };
    fill?: { fgColor?: { rgb?: string } };
    alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
  };
  const next: CellStyle = {
    ...(style.font?.bold ? { bold: true } : {}),
    ...(style.font?.italic ? { italic: true } : {}),
    ...(style.font?.underline ? { underline: true } : {}),
    ...(rgbColor(style.font?.color?.rgb) ? { fontColor: rgbColor(style.font?.color?.rgb) } : {}),
    ...(rgbColor(style.fill?.fgColor?.rgb ?? style.fgColor?.rgb)
      ? { fillColor: rgbColor(style.fill?.fgColor?.rgb ?? style.fgColor?.rgb) }
      : {}),
    ...(style.alignment?.horizontal &&
    ['left', 'center', 'right'].includes(style.alignment.horizontal)
      ? { horizontalAlignment: style.alignment.horizontal as CellStyle['horizontalAlignment'] }
      : {}),
    ...(style.alignment?.vertical && ['top', 'center', 'bottom'].includes(style.alignment.vertical)
      ? {
          verticalAlignment: (style.alignment.vertical === 'center'
            ? 'middle'
            : style.alignment.vertical) as CellStyle['verticalAlignment'],
        }
      : {}),
    ...(style.alignment?.wrapText ? { wrapText: true } : {}),
  };
  return Object.keys(next).length > 0 ? next : undefined;
}

/**
 * Converts a parsed SheetJS sheet into engine cells, promoting numeric cells to real dates when
 * their number format says so. This is what makes `=B2-A2` produce a day count on a real file:
 * a bare serial is indistinguishable from an ordinary number until its format is consulted.
 */
function sheetFromWorksheet(
  XLSX: XlsxModule,
  sheetName: string,
  ws: WorkSheet,
  dateSystem: '1900' | '1904',
  counters: { dates: number },
  dropped: string[],
): Sheet {
  if (!ws['!ref']) return { name: sheetName, rows: [] };

  const range = XLSX.utils.decode_range(ws['!ref']);
  const rowCount = range.e.r + 1;
  const colCount = range.e.c + 1;

  if (rowCount > MAX_EXCEL_ROWS || colCount > MAX_EXCEL_COLUMNS) {
    throw new Error(
      `Sheet "${sheetName}" exceeds Excel's ${MAX_EXCEL_ROWS} row by ${MAX_EXCEL_COLUMNS} column limits.`,
    );
  }
  if (rowCount * colCount > MAX_IMPORTED_CELLS) {
    throw new Error(
      `Sheet "${sheetName}" contains ${rowCount * colCount} cells, above the ${MAX_IMPORTED_CELLS} safety limit.`,
    );
  }

  recordDroppedFeatures(ws, dropped);

  const rows: Cell[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const rowCells: Cell[] = [];
    for (let c = 0; c < colCount; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      const cellObj = ws[address] as
        { v?: unknown; f?: string; z?: string; t?: string; w?: string; s?: unknown } | undefined;

      if (!cellObj) {
        rowCells.push(createCell(null));
        continue;
      }

      const formula = typeof cellObj.f === 'string' ? cellObj.f : undefined;
      const numberFormat = typeof cellObj.z === 'string' ? cellObj.z : undefined;

      let value: CellValue = null;
      if (cellObj.t === 'e' && cellObj.v !== undefined && cellObj.v !== null) {
        value =
          typeof cellObj.w === 'string'
            ? cellObj.w
            : (EXCEL_ERROR_TEXT[Number(cellObj.v)] ?? '#VALUE!');
      } else if (cellObj.v === undefined || cellObj.v === null) {
        value = null;
      } else if (typeof cellObj.v === 'string') {
        value = cellObj.v;
      } else if (typeof cellObj.v === 'boolean') {
        value = cellObj.v;
      } else if (typeof cellObj.v === 'number') {
        // A formula's cached value must stay numeric: Excel recalculates it anyway, and
        // rewriting it as a date would corrupt a formula that merely looks date-formatted.
        const asDate =
          !formula && isDateNumberFormat(numberFormat)
            ? excelSerialToDate(cellObj.v, dateSystem)
            : null;
        if (asDate) {
          value = asDate;
          counters.dates += 1;
        } else {
          value = cellObj.v;
        }
      } else {
        value = String(cellObj.v);
      }

      rowCells.push(
        createCell(value, { formula, numberFormat, style: styleFromSheetJs(cellObj.s) }),
      );
    }
    rows.push(rowCells);
  }

  return { name: sheetName, rows };
}

/** Notes workbook features this workspace cannot represent, so the loss is never silent. */
function recordDroppedFeatures(ws: WorkSheet, dropped: string[]): void {
  const note = (feature: string): void => {
    if (!dropped.includes(feature)) dropped.push(feature);
  };
  if (ws['!merges'] && ws['!merges'].length > 0) note('merged cells');
  if (ws['!cols'] && ws['!cols'].length > 0) note('column widths');
  if (ws['!rows'] && ws['!rows'].length > 0) note('row heights');
  if (ws['!autofilter']) note('autofilter');
  if (ws['!protect']) note('sheet protection');

  for (const key of Object.keys(ws)) {
    if (key.startsWith('!')) continue;
    const cell = ws[key] as { c?: unknown; l?: unknown };
    if (cell?.c) note('comments');
    if (cell?.l) note('hyperlinks');
  }
}

function detectDateSystem(workbook: WorkBook): '1900' | '1904' {
  // Legacy Mac Excel wrote serials from a 1904 epoch; reading those as 1900 shifts dates by
  // four years and a day, which is silent and total.
  const props = (workbook as unknown as { Workbook?: { WBProps?: { date1904?: boolean } } })
    .Workbook?.WBProps;
  return props?.date1904 ? '1904' : '1900';
}

/**
 * Parses raw file bytes into a workbook plus an honest report of what was interpreted and what
 * was lost. Shared by the worker and the main-thread fallback so both behave identically.
 */
export function parseWorkbookBytes(
  XLSX: XlsxModule,
  arrayBuffer: ArrayBuffer,
  fallbackSheetName = 'Sheet1',
): { workbook: Workbook; report: ImportReport } {
  const bytes = new Uint8Array(arrayBuffer);
  const kind = classifyWorkbookBytes(bytes);
  const dropped: string[] = [];
  const counters = { dates: 0 };

  const report: ImportReport = {
    source: kind,
    encoding: 'n/a',
    delimiter: 'n/a',
    dateSystem: '1900',
    cellsConvertedToDates: 0,
    dropped,
  };

  if (kind === 'csv') {
    const { text, encoding } = decodeTextBytes(bytes);
    const delimiter = detectDelimiter(text);
    report.encoding = encoding;
    report.delimiter = delimiter === null ? 'none' : delimiter === '\t' ? 'tab' : delimiter;
    return {
      workbook: {
        sheets: [workbookFromCsvText(text, delimiter, fallbackSheetName)],
        dateSystem: '1900',
      },
      report,
    };
  }

  const parsed = XLSX.read(bytes, {
    type: 'array',
    cellDates: false,
    cellFormula: true,
    cellStyles: true,
  });
  const dateSystem = detectDateSystem(parsed);
  report.dateSystem = dateSystem;

  let totalCells = 0;
  for (const sheetName of parsed.SheetNames) {
    const ref = parsed.Sheets[sheetName]?.['!ref'];
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    totalCells += (range.e.r + 1) * (range.e.c + 1);
    if (!Number.isSafeInteger(totalCells) || totalCells > MAX_IMPORTED_CELLS) {
      throw new Error(`Workbook exceeds the ${MAX_IMPORTED_CELLS} cell import limit.`);
    }
  }

  const sheets = parsed.SheetNames.map((sheetName) =>
    sheetFromWorksheet(XLSX, sheetName, parsed.Sheets[sheetName]!, dateSystem, counters, dropped),
  );

  report.cellsConvertedToDates = counters.dates;

  return {
    // The epoch is carried on the workbook so every consumer - the grid renderer AND the
    // formula evaluator - reads it from one place instead of re-detecting it.
    workbook: {
      sheets: sheets.length > 0 ? sheets : [{ name: fallbackSheetName, rows: [] }],
      dateSystem,
    },
    report,
  };
}

/** Fallback format for a date that carries no format of its own, so it still reads as a date. */
const DEFAULT_DATE_FORMAT = 'yyyy-mm-dd';
const EXCEL_ERROR_TEXT: Record<number, string> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
};

/**
 * Excel stores every date as a serial number plus a display format. Writing a raw `Date`
 * through `aoa_to_sheet` depends on reader options, so dates are converted back to serials
 * explicitly and given a format when the cell does not already carry one.
 */
function valueForExport(cell: { value: unknown }, dateSystem: '1900' | '1904'): unknown {
  const { value } = cell;
  if (value === null || value === '') return null;
  if (value instanceof Date) return dateToExcelSerial(value, dateSystem) ?? null;
  return value;
}

function formatForExport(cell: { value: unknown; numberFormat?: string }): string | undefined {
  if (cell.value instanceof Date && !isDateNumberFormat(cell.numberFormat)) {
    return cell.numberFormat ?? DEFAULT_DATE_FORMAT;
  }
  return cell.numberFormat;
}

function styleForExport(style: CellStyle | undefined): Record<string, unknown> | undefined {
  if (!style) return undefined;
  const font: Record<string, unknown> = {};
  if (style.bold !== undefined) font.bold = style.bold;
  if (style.italic !== undefined) font.italic = style.italic;
  if (style.underline !== undefined) font.underline = style.underline;
  if (style.fontColor) font.color = { rgb: `FF${style.fontColor.replace(/^#/, '')}` };
  const fill = style.fillColor
    ? { patternType: 'solid', fgColor: { rgb: `FF${style.fillColor.replace(/^#/, '')}` } }
    : undefined;
  const alignment = {
    ...(style.horizontalAlignment ? { horizontal: style.horizontalAlignment } : {}),
    ...(style.verticalAlignment
      ? { vertical: style.verticalAlignment === 'middle' ? 'center' : style.verticalAlignment }
      : {}),
    ...(style.wrapText !== undefined ? { wrapText: style.wrapText } : {}),
  };
  return {
    ...(Object.keys(font).length > 0 ? { font } : {}),
    ...(fill ? { fill } : {}),
    ...(Object.keys(alignment).length > 0 ? { alignment } : {}),
  };
}

/** Excel rejects these characters in a sheet name, and duplicates are not allowed either. */
export function sanitizeSheetName(name: string, used: Set<string>): string {
  const base =
    (name || 'Sheet')
      .replace(/[:\\/?*[\]]/g, ' ')
      .trim()
      .slice(0, 31) || 'Sheet';
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) {
    const tag = `_${suffix}`;
    candidate = `${base.slice(0, 31 - tag.length)}${tag}`;
    suffix += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

interface SheetJsZipApi {
  read(bytes: Uint8Array, options: { type: 'buffer' }): unknown;
  find(archive: unknown, path: string): { content: Uint8Array } | null;
  utils: { cfb_add(archive: unknown, path: string, content: Uint8Array): unknown };
  write(
    archive: unknown,
    options: { fileType: 'zip'; type: 'buffer'; compression: boolean },
  ): Uint8Array | number[] | ArrayBuffer;
}

/**
 * SheetJS CE (including 0.18.5 and 0.20.3) reads CalcPr but does not write it. Use its bundled
 * ZIP API to put the instructions in the actual workbook XML. No Node APIs or extra ZIP
 * dependency are needed, so worker and browser exports use the same path. Never pretend
 * recalculation was requested if the XML part cannot be updated.
 */
function requestExcelRecalculation(XLSX: XlsxModule, bytes: Uint8Array): Uint8Array {
  const zip = XLSX.CFB as SheetJsZipApi;
  const archive = zip.read(bytes, { type: 'buffer' });
  const part = zip.find(archive, '/xl/workbook.xml');
  if (!part)
    throw new Error(
      'Cannot export formulas: workbook.xml is missing; recalculation could not be enabled.',
    );
  const xml = new TextDecoder().decode(part.content);
  if (!xml.includes('</workbook>'))
    throw new Error('Cannot export formulas: invalid workbook XML.');
  const calcPr = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  const updated = /<calcPr\b/.test(xml)
    ? xml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/, calcPr)
    : xml.replace('</workbook>', `${calcPr}</workbook>`);
  zip.utils.cfb_add(archive, '/xl/workbook.xml', new TextEncoder().encode(updated));
  return new Uint8Array(zip.write(archive, { fileType: 'zip', type: 'buffer', compression: true }));
}

/**
 * Serializes a workbook to real .xlsx bytes.
 *
 * Blank cells stay blank, formula caches are evaluated against this immutable snapshot,
 * per-cell formats and the date epoch are retained, and actual XML requests a full Excel
 * recalculation for unsupported formulas. Sheet names are made Excel-legal only when doing
 * so cannot silently invalidate formulas; conflicting formula-bearing exports are refused.
 */
export function workbookToXlsxBytes(XLSX: XlsxModule, workbook: Workbook): Uint8Array {
  const book = XLSX.utils.book_new();
  const usedNames = new Set<string>();
  const sheetNames = workbook.sheets.map((sheet) => sanitizeSheetName(sheet.name, usedNames));
  const hasFormulas = workbook.sheets.some((sheet) =>
    sheet.rows.some((row) => row.some((cell) => cell.formula !== undefined)),
  );
  const conflict = workbook.sheets.findIndex((sheet, index) => sheet.name !== sheetNames[index]);
  if (hasFormulas && conflict >= 0) {
    const original = workbook.sheets[conflict]!.name;
    throw new Error(
      `Cannot export: sheet name "${original}" would be renamed to "${sheetNames[conflict]}". ` +
        'That could break sheet-qualified formula references. Rename the sheets and update their references in Excel or LibreOffice, ' +
        'or replace the formulas with values before exporting.',
    );
  }
  const dateSystem = workbook.dateSystem ?? '1900';
  const read = createWorkbookValueReader(workbook);

  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    let widest = 0;
    for (const row of sheet.rows) if (row.length > widest) widest = row.length;
    if (sheet.rows.length > MAX_EXCEL_ROWS || widest > MAX_EXCEL_COLUMNS) {
      throw new Error(
        `Cannot export "${sheet.name}": Excel supports at most ${MAX_EXCEL_ROWS} rows and ${MAX_EXCEL_COLUMNS} columns per sheet.`,
      );
    }
    const aoa: unknown[][] = sheet.rows.map((row) =>
      row.map((cell) => valueForExport(cell, dateSystem)),
    );
    const worksheet = XLSX.utils.aoa_to_sheet(aoa) as Record<string, unknown>;

    // Write formula cells explicitly: aoa_to_sheet does not understand {f, v} objects.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!cell.formula) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const live = read(sheet.name, columnIndex, rowIndex + 1);
        const cached = live instanceof Date ? dateToExcelSerial(live, dateSystem) : live;
        // The evaluator represents both Excel errors and literal error-looking text as strings,
        // and an unknown function may be valid in Excel. Do not invent a typed error cache or
        // reuse a stale numeric cache in either case: preserve the formula for recalculation.
        const uncached = isFormulaError(cached);
        const type = uncached
          ? 'e'
          : typeof cached === 'number' || cached === null
            ? 'n'
            : typeof cached === 'boolean'
              ? 'b'
              : 'str';
        worksheet[address] = {
          t: type,
          f: cell.formula.replace(/^=/, ''),
          ...(!uncached ? { v: cached ?? null } : {}),
        };
      });
    });

    // Formats are applied after the value pass so a cell created above still receives one.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        const format = formatForExport(cell);
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const target = worksheet[address] as
          { z?: string; s?: Record<string, unknown> } | undefined;
        if (target) {
          if (format) target.z = format;
          const style = styleForExport(cell.style);
          if (style) target.s = style;
        } else if ((format || cell.style) && cell.value !== null && cell.value !== '') {
          worksheet[address] = {
            t: 's',
            v: String(cell.value),
            ...(format ? { z: format } : {}),
            ...(styleForExport(cell.style) ? { s: styleForExport(cell.style) } : {}),
          };
        }
      });
    });

    // A spread-based `Math.max(...rows.map(...))` overflows the stack past ~125k rows, which is
    // reachable at a fraction of the import cell limit. A linear scan has no such ceiling.
    const rowCount = Math.max(1, sheet.rows.length);
    const columnCount = Math.max(1, widest);
    // Pin the used range so trailing blank columns and rows are not silently dropped.
    worksheet['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: rowCount - 1, c: columnCount - 1 },
    });

    XLSX.utils.book_append_sheet(book, worksheet, sheetNames[sheetIndex]!);
  }

  book.Workbook = {
    ...book.Workbook,
    WBProps: { ...book.Workbook?.WBProps, date1904: dateSystem === '1904' },
  };

  const bytes = new Uint8Array(
    XLSX.write(book, { bookType: 'xlsx', type: 'array', cellStyles: true }),
  );
  return hasFormulas ? requestExcelRecalculation(XLSX, bytes) : bytes;
}
