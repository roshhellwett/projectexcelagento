import { z } from 'zod';

import type {
  Cell,
  CellLocation,
  CellRange,
  CellValue,
  Operation,
  Preview,
  Report,
  ValidationIssue,
  ValidationResult,
  Workbook,
} from './types.js';
import {
  cellEquals,
  cloneCell,
  cloneWorkbook,
  columnToIndex,
  createCell,
  getSheet,
  indexToColumn,
  invertPatch,
  locationFor,
  maxColumnCount,
  patchBetween,
} from './workbook.js';

export const cellValueSchema = z.union([
  z.string(),
  z.number().finite(),
  z.boolean(),
  z.date(),
  z.null(),
]);

export const sheetSchema = z.object({ sheet: z.string().trim().min(1) });

export function issue(code: string, message: string, location?: CellLocation): ValidationIssue {
  return { code, message, ...(location ? { location } : {}) };
}

export function validResult(warnings: ValidationIssue[] = []): ValidationResult {
  return { valid: true, errors: [], warnings };
}

export function invalidResult(...errors: ValidationIssue[]): ValidationResult {
  return { valid: false, errors, warnings: [] };
}

/**
 * Structural edits must not just move formula text: relative, absolute, range, and cross-sheet
 * references need a full Excel-aware rewriter. Until one exists, refuse these operations for
 * any formula-bearing workbook. This deliberately includes formulas on other sheets and
 * unsupported reference syntax, rather than claiming safety from our subset evaluator.
 * Value edits and formula-free structural workflows remain available.
 */
export function withFormulaStructureSafety<Args>(operation: Operation<Args>): Operation<Args> {
  const errorsFor = (workbook: Workbook): ValidationIssue[] => {
    for (const sheet of workbook.sheets) {
      for (const [rowIndex, row] of sheet.rows.entries()) {
        const column = row.findIndex((cell) => cell.formula !== undefined);
        if (column < 0) continue;
        const address = `${sheet.name}!${indexToColumn(column)}${rowIndex + 1}`;
        return [
          issue(
            'formula-structural-edit',
            `${operation.name} cannot safely move, copy, insert, or delete cells while this workbook contains formulas (${address}). ` +
              'Formula reference rewriting, including cross-sheet references, is not supported. ' +
              'Perform this structural change in Excel or LibreOffice, or replace the formulas with values first; ordinary value edits are still available.',
            { sheet: sheet.name, row: rowIndex + 1, column: indexToColumn(column) },
          ),
        ];
      }
    }
    return [];
  };
  return {
    ...operation,
    validate(workbook, args) {
      const validation = operation.validate(workbook, args);
      const errors = [...validation.errors, ...errorsFor(workbook)];
      return { ...validation, valid: validation.valid && errors.length === 0, errors };
    },
    preview(workbook, args) {
      const errors = errorsFor(workbook);
      return errors.length > 0
        ? previewForTransition(workbook, workbook, [], [], errors)
        : operation.preview(workbook, args);
    },
    apply(workbook, args) {
      const errors = errorsFor(workbook);
      if (errors.length > 0) throw new Error(errors[0]!.message);
      return operation.apply(workbook, args);
    },
  };
}

export function getSheetOrUndefined(workbook: Workbook, sheetName: string) {
  return getSheet(workbook, sheetName);
}

export function validateSheet(workbook: Workbook, sheetName: string): ValidationIssue[] {
  return getSheet(workbook, sheetName)
    ? []
    : [issue('missing-sheet', `Sheet "${sheetName}" does not exist.`)];
}

export function validateColumn(
  workbook: Workbook,
  sheetName: string,
  column: string,
): ValidationIssue[] {
  const sheet = getSheet(workbook, sheetName);
  const columnIndex = columnToIndex(column);
  if (columnIndex === undefined) {
    return [issue('invalid-column', `Column "${column}" is not a valid column reference.`)];
  }
  if (!sheet) {
    return [issue('missing-sheet', `Sheet "${sheetName}" does not exist.`)];
  }
  if (maxColumnCount(sheet.rows) <= columnIndex) {
    return [
      issue(
        'missing-column',
        `Column "${indexToColumn(columnIndex)}" does not exist in the sheet.`,
      ),
    ];
  }
  return [];
}

export function headerRowError(
  workbook: Workbook,
  sheet: string,
  headerRow: number,
): ValidationIssue[] {
  const sheetData = getSheet(workbook, sheet);
  return !sheetData || headerRow <= sheetData.rows.length
    ? []
    : [issue('invalid-header-row', `Header row ${headerRow} is outside sheet "${sheet}".`)];
}

export function maxRow(workbook: Workbook, sheetName: string): number {
  return getSheet(workbook, sheetName)?.rows.length ?? 0;
}

export function maxColumn(workbook: Workbook, sheetName: string): number {
  return maxColumnCount(getSheet(workbook, sheetName)?.rows ?? []);
}

export function allDataRange(workbook: Workbook, sheetName: string, headerRow: number): CellRange {
  return {
    sheet: sheetName,
    startRow: Math.min(headerRow + 1, Math.max(1, maxRow(workbook, sheetName))),
    endRow: Math.max(headerRow + 1, maxRow(workbook, sheetName)),
    startColumn: 'A',
    endColumn: indexToColumn(Math.max(0, maxColumn(workbook, sheetName) - 1)),
  };
}

export function fullSheetRange(workbook: Workbook, sheetName: string): CellRange {
  return {
    sheet: sheetName,
    startRow: 1,
    endRow: Math.max(1, maxRow(workbook, sheetName)),
    startColumn: 'A',
    endColumn: indexToColumn(Math.max(0, maxColumn(workbook, sheetName) - 1)),
  };
}

export function rangeForColumns(
  workbook: Workbook,
  sheetName: string,
  headerRow: number,
  columns: string[],
): CellRange[] {
  const endRow = Math.max(headerRow + 1, maxRow(workbook, sheetName));
  return columns.flatMap((column) => {
    const index = columnToIndex(column);
    return index === undefined
      ? []
      : [
          {
            sheet: sheetName,
            startRow: headerRow + 1,
            endRow,
            startColumn: indexToColumn(index),
            endColumn: indexToColumn(index),
          },
        ];
  });
}

export function cloneWithValue(
  cell: Cell,
  value: CellValue,
  options: { formula?: string; numberFormat?: string; type?: Cell['type'] } = {},
): Cell {
  return createCell(value, {
    type: options.type ?? (options.formula === undefined ? undefined : 'formula'),
    ...(options.formula !== undefined ? { formula: options.formula } : {}),
    numberFormat: options.numberFormat ?? cell.numberFormat,
  });
}

export function cellAt(
  workbook: Workbook,
  sheetName: string,
  row: number,
  column: number,
): Cell | undefined {
  return getSheet(workbook, sheetName)?.rows[row - 1]?.[column];
}

interface NumericRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/**
 * Parse the column bounds of every declared range for one sheet exactly once.
 * `columnToIndex` normalizes and regex-tests its argument, so resolving the
 * ranges inside the per-cell loop made every operation O(cells x ranges) with a
 * string parse on each comparison.
 */
function numericRangesForSheet(ranges: CellRange[], sheetName: string): NumericRange[] {
  const resolved: NumericRange[] = [];
  for (const range of ranges) {
    if (range.sheet !== sheetName) {
      continue;
    }
    const startColumn = columnToIndex(range.startColumn);
    const endColumn = columnToIndex(range.endColumn);
    if (startColumn === undefined || endColumn === undefined) {
      continue;
    }
    resolved.push({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn,
      endColumn,
    });
  }
  return resolved;
}

function rangeContains(ranges: NumericRange[], row: number, column: number): boolean {
  for (const range of ranges) {
    if (
      row >= range.startRow &&
      row <= range.endRow &&
      column >= range.startColumn &&
      column <= range.endColumn
    ) {
      return true;
    }
  }
  return false;
}

export function previewForTransition(
  before: Workbook,
  after: Workbook,
  ranges: CellRange[],
  warnings: ValidationIssue[] = [],
  errors: ValidationIssue[] = [],
  requiresConfirmation = false,
): Preview {
  const changes: Preview['changes'] = [];
  let affectedCells = 0;
  const sheetNames = new Set([
    ...before.sheets.map((sheet) => sheet.name),
    ...after.sheets.map((sheet) => sheet.name),
  ]);

  for (const sheetName of sheetNames) {
    // Ranges are resolved per sheet up front and membership is only tested for
    // cells that actually changed, so the scan stays O(cells + ranges).
    const numericRanges = numericRangesForSheet(ranges, sheetName);
    if (numericRanges.length === 0) {
      continue;
    }
    const beforeSheet = getSheet(before, sheetName);
    const afterSheet = getSheet(after, sheetName);
    const rows = Math.max(beforeSheet?.rows.length ?? 0, afterSheet?.rows.length ?? 0);
    const columns = Math.max(
      maxColumnCount(beforeSheet?.rows ?? []),
      maxColumnCount(afterSheet?.rows ?? []),
    );
    for (let row = 1; row <= rows; row += 1) {
      for (let column = 0; column < columns; column += 1) {
        const beforeCell = beforeSheet?.rows[row - 1]?.[column];
        const afterCell = afterSheet?.rows[row - 1]?.[column];
        const changed =
          beforeCell === undefined
            ? afterCell !== undefined
            : afterCell === undefined || !cellEquals(beforeCell, afterCell);
        if (changed && rangeContains(numericRanges, row, column)) {
          affectedCells += 1;
          if (changes.length < 20) {
            changes.push({
              location: locationFor(sheetName, row, column),
              before: cloneCell(beforeCell ?? createCell(null)),
              after: cloneCell(afterCell ?? createCell(null)),
            });
          }
        }
      }
    }
  }

  return {
    valid: errors.length === 0,
    affectedCells,
    changes,
    warnings,
    errors,
    requiresConfirmation,
  };
}

export function transitionResult(before: Workbook, after: Workbook, report: Report) {
  const patch = patchBetween(before, after);
  return {
    workbook: cloneWorkbook(after),
    report,
    patch,
    inverse: invertPatch(patch),
  };
}

export function reportForTransition(
  before: Workbook,
  after: Workbook,
  warnings: ValidationIssue[] = [],
  extras: Partial<Report> = {},
): Report {
  const preview = previewForTransition(before, after, allWorkbookRange(before, after));
  return {
    affectedCells: preview.affectedCells,
    skippedCells: 0,
    unchangedCells: Math.max(0, countCells(before) - preview.affectedCells),
    warnings,
    ...extras,
  };
}

function allWorkbookRange(before: Workbook, after: Workbook): CellRange[] {
  const sheets = new Set([
    ...before.sheets.map((sheet) => sheet.name),
    ...after.sheets.map((sheet) => sheet.name),
  ]);
  return [...sheets].map((sheet) => ({
    sheet,
    startRow: 1,
    endRow: Math.max(maxRow(before, sheet), maxRow(after, sheet), 1),
    startColumn: 'A',
    endColumn: indexToColumn(Math.max(maxColumn(before, sheet), maxColumn(after, sheet), 1) - 1),
  }));
}

function countCells(workbook: Workbook): number {
  return workbook.sheets.reduce(
    (total, sheet) => total + sheet.rows.reduce((rows, row) => rows + row.length, 0),
    0,
  );
}

/**
 * A sheet name that is not taken yet. Two operations can legitimately ask for the same derived
 * sheet name, and overwriting the first result in place would destroy work the user still has in
 * their history - so a collision appends a counter instead of silently replacing anything.
 */
export function uniqueSheetName(workbook: Workbook, desired: string): string {
  const names = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  if (!names.has(desired.toLowerCase())) return desired;
  let suffix = 2;
  while (names.has(`${desired} (${suffix})`.toLowerCase())) suffix += 1;
  return `${desired} (${suffix})`;
}

export function textForCell(cell: Cell | undefined): string {
  if (!cell || cell.value === null) {
    return '';
  }
  return cell.value instanceof Date ? cell.value.toISOString() : String(cell.value);
}

export function normalizedText(value: string, caseMode: 'none' | 'lower' | 'upper' | 'title') {
  const collapsed = value.trim().replace(/\s+/g, ' ');
  switch (caseMode) {
    case 'lower':
      return collapsed.toLowerCase();
    case 'upper':
      return collapsed.toUpperCase();
    case 'title':
      return collapsed.replace(/\b\w/g, (character) => character.toUpperCase());
    case 'none':
      return collapsed;
  }
}
