import { z } from 'zod';

import type {
  Cell,
  CellLocation,
  CellRange,
  InvariantResult,
  Operation,
  Preview,
  Report,
  ValidationIssue,
  ValidationResult,
  Workbook,
} from './types.js';
import { FICTITIOUS_SERIAL_60, excelSerialToDate } from './formula/excel-date.js';
import { runInvariants } from './invariants.js';
import {
  cellEquals,
  cloneCell,
  columnToIndex,
  getSheet,
  indexToColumn,
  invertPatch,
  locationFor,
  maxColumnCount,
  patchForChange,
  applyPatch,
} from './workbook.js';
import { previewForTransition } from './operation-utils.js';

export const dateFormats = ['YYYY-MM-DD', 'MM/DD/YYYY', 'DD/MM/YYYY'] as const;
export type DateFormat = (typeof dateFormats)[number];

export const formatDatesArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: z
    .string()
    .trim()
    .regex(/^[A-Za-z]+$/),
  format: z.enum(dateFormats),
  headerRow: z.number().int().positive().default(1),
  /** Defaults to the workbook epoch; callers may override it for explicitly foreign serial data. */
  dateSystem: z.enum(['1900', '1904']).optional(),
});

export type FormatDatesArgs = z.infer<typeof formatDatesArgsSchema>;

type ParsedDate =
  | { kind: 'date'; date: Date }
  | { kind: 'empty' }
  | { kind: 'not-date' }
  | { kind: 'ambiguous'; candidates: Date[] }
  | { kind: 'invalid'; reason: string };

interface ScannedCell {
  location: CellLocation;
  cell: Cell;
  parsed: ParsedDate;
}

function issue(code: string, message: string, location?: CellLocation): ValidationIssue {
  return { code, message, ...(location ? { location } : {}) };
}

function isBlank(value: Cell['value']): boolean {
  return value === null || (typeof value === 'string' && value.trim() === '');
}

function utcDate(year: number, month: number, day: number): Date | undefined {
  if (!Number.isInteger(year) || year < 1000 || year > 9999) {
    return undefined;
  }

  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return undefined;
  }
  return date;
}

/**
 * Resolves an Excel serial to a day, deferring the epoch arithmetic to `excel-date.ts` so this
 * operation cannot drift from the formula engine or the grid renderer.
 *
 * The 1900 system carries Excel's own leap-year bug: serial 60 names 1900-02-29, a day that never
 * existed. The shared converter refuses it, and this wrapper keeps the distinction the report
 * needs - serial 60 is a warning about a real value, while an out-of-range serial is simply not a
 * date and must be left alone.
 */
function parseExcelSerial(serial: number, dateSystem: FormatDatesArgs['dateSystem']): ParsedDate {
  if (!Number.isFinite(serial)) {
    return { kind: 'invalid', reason: 'The numeric value is not finite.' };
  }

  if (dateSystem === '1900' && serial === FICTITIOUS_SERIAL_60) {
    return {
      kind: 'invalid',
      reason: 'Excel serial 60 is the fictitious 1900-02-29 and cannot be represented.',
    };
  }

  const date = excelSerialToDate(serial, dateSystem);
  return date ? { kind: 'date', date } : { kind: 'not-date' };
}

function parseTextDate(text: string): ParsedDate {
  const trimmed = text.trim();
  if (trimmed === '') {
    return { kind: 'empty' };
  }

  const isoMatch = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ].*)?$/.exec(trimmed);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const date = utcDate(year, month, day);
    return date
      ? { kind: 'date', date }
      : { kind: 'invalid', reason: `The date ${trimmed} is not a real calendar date.` };
  }

  const slashMatch = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(trimmed);
  if (!slashMatch) {
    return { kind: 'not-date' };
  }

  const first = Number(slashMatch[1]);
  const second = Number(slashMatch[2]);
  const year = Number(slashMatch[3]);
  const monthFirst = utcDate(year, first, second);
  const dayFirst = utcDate(year, second, first);

  if (first > 12 && dayFirst) {
    return { kind: 'date', date: dayFirst };
  }
  if (second > 12 && monthFirst) {
    return { kind: 'date', date: monthFirst };
  }
  if (monthFirst && dayFirst) {
    return { kind: 'ambiguous', candidates: [monthFirst, dayFirst] };
  }
  if (monthFirst) {
    return { kind: 'date', date: monthFirst };
  }
  if (dayFirst) {
    return { kind: 'date', date: dayFirst };
  }

  return {
    kind: 'invalid',
    reason: `The date ${trimmed} is not a real calendar date in either month/day or day/month order.`,
  };
}

function parseDate(value: Cell['value'], dateSystem: FormatDatesArgs['dateSystem']): ParsedDate {
  if (isBlank(value)) {
    return { kind: 'empty' };
  }
  if (typeof value === 'number') {
    return parseExcelSerial(value, dateSystem);
  }
  if (typeof value === 'string') {
    return parseTextDate(value);
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return { kind: 'invalid', reason: 'The cell contains an invalid Date value.' };
    }
    const date = utcDate(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
    return date
      ? { kind: 'date', date }
      : {
          kind: 'invalid',
          reason: 'The Date value is outside the supported four-digit year range.',
        };
  }
  if (typeof value === 'string') {
    return parseTextDate(value);
  }
  return { kind: 'not-date' };
}

export function formatDate(date: Date, format: DateFormat): string {
  const year = date.getUTCFullYear().toString().padStart(4, '0');
  const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
  const day = date.getUTCDate().toString().padStart(2, '0');

  switch (format) {
    case 'YYYY-MM-DD':
      return `${year}-${month}-${day}`;
    case 'MM/DD/YYYY':
      return `${month}/${day}/${year}`;
    case 'DD/MM/YYYY':
      return `${day}/${month}/${year}`;
  }
}

export function displayCell(cell: Cell): string {
  if (cell.value instanceof Date) {
    const format = dateFormats.includes(cell.numberFormat as DateFormat)
      ? (cell.numberFormat as DateFormat)
      : 'YYYY-MM-DD';
    return formatDate(cell.value, format);
  }
  if (cell.value === null) {
    return '';
  }
  return String(cell.value);
}

function scan(workbook: Workbook, args: FormatDatesArgs): ScannedCell[] {
  const sheet = getSheet(workbook, args.sheet);
  const columnIndex = columnToIndex(args.column);
  if (!sheet || columnIndex === undefined) {
    return [];
  }

  const scanned: ScannedCell[] = [];
  const dateSystem = args.dateSystem ?? workbook.dateSystem ?? '1900';
  for (let rowIndex = args.headerRow; rowIndex < sheet.rows.length; rowIndex += 1) {
    const cell = sheet.rows[rowIndex]?.[columnIndex];
    if (!cell) {
      continue;
    }
    scanned.push({
      location: locationFor(args.sheet, rowIndex + 1, columnIndex),
      cell,
      parsed: parseDate(cell.value, dateSystem),
    });
  }
  return scanned;
}

function structuralValidation(workbook: Workbook, args: FormatDatesArgs): ValidationIssue[] {
  const errors: ValidationIssue[] = [];
  const sheet = getSheet(workbook, args.sheet);
  const columnIndex = columnToIndex(args.column);

  if (!sheet) {
    errors.push(issue('missing-sheet', `Sheet "${args.sheet}" does not exist.`));
    return errors;
  }
  if (columnIndex === undefined) {
    errors.push(
      issue('invalid-column', `Column "${args.column}" is not a valid column reference.`),
    );
    return errors;
  }
  if (args.headerRow > sheet.rows.length) {
    errors.push(
      issue('invalid-header-row', `Header row ${args.headerRow} is outside sheet "${args.sheet}".`),
    );
  }
  if (maxColumnCount(sheet.rows) <= columnIndex) {
    errors.push(
      issue(
        'missing-column',
        `Column "${indexToColumn(columnIndex)}" does not exist in the sheet.`,
      ),
    );
  }
  return errors;
}

function scanWarnings(scanned: ScannedCell[]): ValidationIssue[] {
  const warnings: ValidationIssue[] = [];
  for (const item of scanned) {
    if (item.parsed.kind === 'ambiguous') {
      const [first, second] = item.parsed.candidates;
      if (!first || !second) {
        continue;
      }
      warnings.push(
        issue(
          'ambiguous-date',
          `${item.location.column}${item.location.row} is ambiguous; it could be ${formatDate(first, 'MM/DD/YYYY')} or ${formatDate(second, 'MM/DD/YYYY')}. Choose a locale instead of guessing.`,
          item.location,
        ),
      );
    } else if (item.parsed.kind === 'invalid') {
      warnings.push(issue('invalid-date', item.parsed.reason, item.location));
    }
  }
  return warnings;
}

function makeAfterCell(cell: Cell, date: Date, format: DateFormat): Cell {
  return {
    ...cloneCell(cell),
    value: new Date(date.getTime()),
    numberFormat: format,
  };
}

function changesFor(
  workbook: Workbook,
  args: FormatDatesArgs,
): { changes: ReturnType<typeof patchForChange>[]; warnings: ValidationIssue[]; skipped: number } {
  const scanned = scan(workbook, args);
  const changes: ReturnType<typeof patchForChange>[] = [];
  let skipped = 0;

  for (const item of scanned) {
    if (item.parsed.kind === 'date' && !item.cell.formula) {
      const after = makeAfterCell(item.cell, item.parsed.date, args.format);
      if (!cellEquals(item.cell, after)) {
        changes.push(patchForChange(item.location, item.cell, after));
      }
    } else if (item.parsed.kind !== 'empty') {
      skipped += 1;
    }
  }

  return { changes, warnings: scanWarnings(scanned), skipped };
}

function targetRange(workbook: Workbook, args: FormatDatesArgs): CellRange {
  return {
    sheet: args.sheet,
    startRow: args.headerRow + 1,
    endRow: Math.max(args.headerRow + 1, getSheet(workbook, args.sheet)?.rows.length ?? 0),
    startColumn: args.column,
    endColumn: args.column,
  };
}

function buildPreview(workbook: Workbook, args: FormatDatesArgs): Preview {
  const errors = structuralValidation(workbook, args);
  if (errors.length > 0) {
    return previewForTransition(workbook, workbook, [targetRange(workbook, args)], [], errors);
  }

  const { changes, warnings } = changesFor(workbook, args);
  const after = applyPatch(workbook, changes);
  return previewForTransition(
    workbook,
    after,
    [targetRange(workbook, args)],
    warnings,
    [],
    changes.length > 100 || warnings.some((warning) => warning.code === 'ambiguous-date'),
  );
}

function validate(workbook: Workbook, args: FormatDatesArgs): ValidationResult {
  const errors = structuralValidation(workbook, args);
  const warnings = errors.length === 0 ? scanWarnings(scan(workbook, args)) : [];
  return { valid: errors.length === 0, errors, warnings };
}

function apply(workbook: Workbook, args: FormatDatesArgs) {
  const validation = validate(workbook, args);
  if (!validation.valid) {
    throw new Error(validation.errors.map((error) => error.message).join(' '));
  }

  const { changes, warnings, skipped } = changesFor(workbook, args);
  const result = applyPatch(workbook, changes);

  return {
    workbook: result,
    report: {
      affectedCells: changes.length,
      skippedCells: skipped,
      unchangedCells: scan(workbook, args).filter(
        (item) => item.parsed.kind === 'empty' || item.parsed.kind === 'not-date',
      ).length,
      warnings,
    } satisfies Report,
    patch: changes,
    inverse: invertPatch(changes),
  };
}

function invariants(before: Workbook, after: Workbook, args: FormatDatesArgs): InvariantResult {
  return runInvariants(before, after, {
    targetRanges: [targetRange(before, args)],
    rowCountUnchanged: true,
  });
}

export const formatDatesOperation: Operation<FormatDatesArgs> = {
  name: 'format_dates',
  schema: formatDatesArgsSchema,
  targetRanges: (workbook, args) => [targetRange(workbook, args)],
  validate,
  preview: buildPreview,
  apply,
  invariants,
};
