import type { WorkSheet, WorkBook } from 'xlsx';
import {
  createCell,
  dateToExcelSerial,
  excelSerialToDate,
  isDateNumberFormat,
  type Cell,
  type CellValue,
  type Workbook,
  type Sheet,
} from '@excel-agent/engine';

type XlsxModule = typeof import('xlsx');

/** Guard against pathological sheets that would freeze the browser tab. */
export const MAX_IMPORTED_CELLS = 1_500_000;

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

    if (ch === '"') {
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
  const records = parseCsvRecords(text, delimiter, 200);
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

  // Fallback: If no candidate reached the threshold yet the first non-empty line has commas or delimiters,
  // select the most frequent candidate rather than dumping everything into a single un-split column.
  if (!best) {
    const firstNonEmpty = text.split(/\r\n|\n|\r/).find((l) => l.trim().length > 0);
    if (firstNonEmpty) {
      const commaCount = (firstNonEmpty.match(/,/g) || []).length;
      const tabCount = (firstNonEmpty.match(/\t/g) || []).length;
      const semiCount = (firstNonEmpty.match(/;/g) || []).length;
      const pipeCount = (firstNonEmpty.match(/\|/g) || []).length;
      if (commaCount >= 1 && commaCount >= tabCount && commaCount >= semiCount && commaCount >= pipeCount) {
        best = ',';
      } else if (tabCount >= 1 && tabCount >= semiCount && tabCount >= pipeCount) {
        best = '\t';
      } else if (semiCount >= 1 && semiCount >= pipeCount) {
        best = ';';
      } else if (pipeCount >= 1) {
        best = '|';
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
  let text = raw.trim();
  if (text === '' || text.toLowerCase() === 'null') return null;
  // Strip redundant surrounding quotes if the field is wrapped in literal quotes from exports like """..."""
  if (text.length >= 2 && text.startsWith('"') && text.endsWith('"') && !text.slice(1, -1).includes('"')) {
    text = text.slice(1, -1).trim();
    if (text === '' || text.toLowerCase() === 'null') return null;
  }
  if (text.startsWith('=')) return text;

  // Accounting notation: (450) means negative 450.
  const negativeParenthesised = /^\((.*)\)$/.exec(text);
  const body = negativeParenthesised ? negativeParenthesised[1]! : text;
  const negative = negativeParenthesised !== null;

  const isPercent = body.endsWith('%');
  const numericBody = isPercent ? body.slice(0, -1).trim() : body;

  // A leading zero marks an identifier. Only "0" itself and "0.5"-style values are numbers.
  if (/^[-+]?0\d/.test(numericBody)) return text;

  const magnitude = stripNumericDecoration(numericBody, delimiter);
  if (magnitude === null) return text;

  const parsed = Number(magnitude);
  if (!Number.isFinite(parsed)) return text;

  let value = negative ? -parsed : parsed;
  // Applied last, so "45%" and "45" take the same route through the parser.
  if (isPercent) value /= 100;
  return value;
}

function stripNumericDecoration(body: string, delimiter: string): string | null {
  // Currency symbols and spaces are decoration; a leading sign is not.
  const cleaned = body.replace(/[\s\u00a0]/g, '').replace(/^[$€£¥₹]/, '');
  if (cleaned === '') return null;
  if (!/^[-+]?[\d.,]+$/.test(cleaned)) return null;

  const signed = /^[-+]/.test(cleaned) ? cleaned[0]! : '';
  const magnitude = signed ? cleaned.slice(1) : cleaned;

  const hasComma = magnitude.includes(',');
  const hasDot = magnitude.includes('.');

  if (hasComma && hasDot) {
    // Whichever separator comes last is the decimal point.
    return (
      signed +
      (magnitude.lastIndexOf(',') > magnitude.lastIndexOf('.')
        ? magnitude.replace(/\./g, '').replace(',', '.')
        : magnitude.replace(/,/g, ''))
    );
  }
  if (hasComma) {
    // With a comma delimiter a comma is a thousands separator; otherwise it is a decimal comma.
    return signed + (delimiter === ',' ? magnitude.replace(/,/g, '') : magnitude.replace(',', '.'));
  }
  if (hasDot) {
    const parts = magnitude.split('.');
    if (parts.length > 2) return null;
    // "1.234.567" is a grouped integer; "1.23" is a decimal. Only an exact three-digit tail
    // with a short leading group counts as grouping, so "0.123" stays a decimal.
    if (
      parts.length === 2 &&
      parts[1]!.length === 3 &&
      /^\d+$/.test(parts[0]!) &&
      parts[0]!.length <= 3
    ) {
      return signed + magnitude.replace(/\./g, '');
    }
    return signed + magnitude;
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
        { v?: unknown; f?: string; z?: string; t?: string; w?: string } | undefined;

      if (!cellObj || cellObj.v === undefined || cellObj.v === null) {
        rowCells.push(createCell(null));
        continue;
      }

      const formula = typeof cellObj.f === 'string' ? cellObj.f : undefined;
      const numberFormat = typeof cellObj.z === 'string' ? cellObj.z : undefined;

      let value: CellValue = null;
      if (typeof cellObj.v === 'string') {
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
      } else if (cellObj.t === 'e') {
        // An Excel error cell travels as an object; its text lives in `w`.
        value = typeof cellObj.w === 'string' ? cellObj.w : '#VALUE!';
      } else {
        value = String(cellObj.v);
      }

      rowCells.push(createCell(value, { formula, numberFormat }));
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

/**
 * Excel stores every date as a serial number plus a display format. Writing a raw `Date`
 * through `aoa_to_sheet` depends on reader options, so dates are converted back to serials
 * explicitly and given a format when the cell does not already carry one.
 */
function valueForExport(cell: { value: unknown }): unknown {
  const { value } = cell;
  if (value === null || value === '') return null;
  if (value instanceof Date) return dateToExcelSerial(value) ?? null;
  return value;
}

function formatForExport(cell: { value: unknown; numberFormat?: string }): string | undefined {
  if (cell.value instanceof Date && !isDateNumberFormat(cell.numberFormat)) {
    return cell.numberFormat ?? DEFAULT_DATE_FORMAT;
  }
  return cell.numberFormat;
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

/**
 * Serializes a workbook to real .xlsx bytes.
 *
 * Blank cells stay blank rather than becoming empty strings, formulas keep their cached value,
 * per-cell number formats are reapplied, dates are written as serials, sheet names are made
 * Excel-legal and unique, and Excel is asked to recalculate on load rather than trusting values
 * this workspace computed itself.
 */
export function workbookToXlsxBytes(XLSX: XlsxModule, workbook: Workbook): Uint8Array {
  const book = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  for (const sheet of workbook.sheets) {
    const aoa: unknown[][] = sheet.rows.map((row) => row.map(valueForExport));
    const worksheet = XLSX.utils.aoa_to_sheet(aoa) as Record<string, unknown>;

    // Write formula cells explicitly: aoa_to_sheet does not understand {f, v} objects.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!cell.formula) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const cached = cell.value instanceof Date ? dateToExcelSerial(cell.value) : cell.value;
        const type =
          typeof cached === 'number'
            ? 'n'
            : typeof cached === 'boolean'
              ? 'b'
              : cached === null || cached === undefined
                ? 'n'
                : 'str';
        worksheet[address] = {
          t: type,
          f: cell.formula.replace(/^=/, ''),
          v: cached ?? null,
        };
      });
    });

    // Formats are applied after the value pass so a cell created above still receives one.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        const format = formatForExport(cell);
        if (!format) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const target = worksheet[address] as { z?: string } | undefined;
        if (target) {
          target.z = format;
        } else if (cell.value !== null && cell.value !== '') {
          worksheet[address] = { t: 's', v: String(cell.value), z: format };
        }
      });
    });

    // `Math.max(...rows.map(...))` overflows the stack past ~125k rows, which is reachable at a
    // fraction of the import cell limit. A reduce has no such ceiling.
    let widest = 0;
    for (const row of sheet.rows) if (row.length > widest) widest = row.length;
    const rowCount = Math.max(1, sheet.rows.length);
    const columnCount = Math.max(1, widest);
    // Pin the used range so trailing blank columns and rows are not silently dropped.
    worksheet['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: rowCount - 1, c: columnCount - 1 },
    });

    XLSX.utils.book_append_sheet(book, worksheet, sanitizeSheetName(sheet.name, usedNames));
  }

  book.Workbook = {
    ...book.Workbook,
    CalcPr: { fullCalcOnLoad: true },
  } as typeof book.Workbook;

  return new Uint8Array(XLSX.write(book, { bookType: 'xlsx', type: 'array' }));
}
