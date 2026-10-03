import { z } from 'zod';

import { runInvariants } from './invariants.js';
import {
  allDataRange,
  cellAt,
  cellValueSchema,
  cloneWithValue,
  fullSheetRange,
  issue,
  maxColumn,
  maxRow,
  previewForTransition,
  rangeForColumns,
  textForCell,
  transitionResult,
  validResult,
  validateColumn,
  validateSheet,
} from './operation-utils.js';
import type {
  Cell,
  CellRange,
  CellValue,
  Operation,
  OperationResult,
  Preview,
  Report,
  ValidationResult,
  Workbook,
} from './types.js';
import {
  cellValueEquals,
  cloneCell,
  cloneWorkbook,
  columnToIndex,
  createCell,
  getSheet,
  indexToColumn,
  maxColumnCount,
} from './workbook.js';

function operationReport(
  before: Workbook,
  after: Workbook,
  ranges: CellRange[],
  warnings: ValidationResult['warnings'] = [],
  extras: Partial<Report> = {},
): Report {
  const preview = previewForTransition(before, after, ranges, warnings);
  return {
    affectedCells: preview.affectedCells,
    skippedCells: 0,
    unchangedCells: Math.max(
      0,
      before.sheets.reduce((total, sheet) => total + sheet.rows.flat().length, 0) -
        preview.affectedCells,
    ),
    warnings,
    ...extras,
  };
}

function invalidPreview(
  workbook: Workbook,
  ranges: CellRange[],
  errors: ValidationResult['errors'],
): Preview {
  return previewForTransition(workbook, workbook, ranges, [], errors);
}

function headerRowError(
  workbook: Workbook,
  sheet: string,
  headerRow: number,
): ValidationResult['errors'] {
  const sheetData = getSheet(workbook, sheet);
  return !sheetData || headerRow <= sheetData.rows.length
    ? []
    : [issue('invalid-header-row', `Header row ${headerRow} is outside sheet "${sheet}".`)];
}

function columnRange(
  workbook: Workbook,
  sheet: string,
  startRow: number,
  startColumn: number,
  endColumn: number,
  endRow = maxRow(workbook, sheet),
): CellRange {
  return {
    sheet,
    startRow,
    endRow: Math.max(startRow, endRow),
    startColumn: indexToColumn(startColumn),
    endColumn: indexToColumn(Math.max(startColumn, endColumn)),
  };
}

function textMatches(text: string, find: string, matchCase: boolean, wholeCell: boolean): boolean {
  const source = matchCase ? text : text.toLocaleLowerCase();
  const needle = matchCase ? find : find.toLocaleLowerCase();
  return wholeCell ? source === needle : source.includes(needle);
}

function comparableValue(value: CellValue): string | number | boolean | null {
  if (value instanceof Date) {
    return value.getTime();
  }
  return value;
}

function compareValues(left: CellValue, right: CellValue): number {
  const a = comparableValue(left);
  const b = comparableValue(right);
  if (a === null && b === null) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
  }
  // Rank by type so mixed-type sorts are deterministic instead of JS-coerced
  const rank = (v: unknown): number => (typeof v === 'number' ? 0 : typeof v === 'boolean' ? 1 : 2);
  if (typeof a !== typeof b) return rank(a) - rank(b);
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;
  return 0;
}

function rangeForSort(workbook: Workbook, args: SortRangeArgs): CellRange {
  const endColumn =
    columnToIndex(
      args.endColumn ?? indexToColumn(Math.max(0, maxColumn(workbook, args.sheet) - 1)),
    ) ?? 0;
  return columnRange(
    workbook,
    args.sheet,
    args.startRow,
    columnToIndex(args.startColumn) ?? 0,
    endColumn,
    args.endRow ?? maxRow(workbook, args.sheet),
  );
}

const rangeColumn = z
  .string()
  .trim()
  .regex(/^[A-Za-z]+$/);
const headerRow = z.number().int().positive().default(1);

export const sortRangeArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  direction: z.enum(['asc', 'desc']).default('asc'),
  startRow: z.number().int().positive().default(2),
  endRow: z.number().int().positive().optional(),
  startColumn: rangeColumn.default('A'),
  endColumn: rangeColumn.optional(),
});
export type SortRangeArgs = z.infer<typeof sortRangeArgsSchema>;

function validateSort(workbook: Workbook, args: SortRangeArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...headerRowError(workbook, args.sheet, args.startRow),
  ];
  const sheet = getSheet(workbook, args.sheet);
  const sortColumn = columnToIndex(args.column);
  const startColumn = columnToIndex(args.startColumn);
  const range = rangeForSort(workbook, args);
  if (sortColumn === undefined || startColumn === undefined) {
    errors.push(issue('invalid-column', 'Sort columns must be valid column references.'));
  }
  if (sheet && args.endRow !== undefined && args.endRow > sheet.rows.length) {
    errors.push(issue('invalid-range', 'The sort end row is outside the sheet.'));
  }
  if (
    sortColumn !== undefined &&
    (sortColumn < (startColumn ?? 0) || sortColumn > (columnToIndex(range.endColumn) ?? 0))
  ) {
    errors.push(issue('invalid-range', 'The sort column must be inside the requested sort range.'));
  }
  if (args.endRow !== undefined && args.endRow < args.startRow) {
    errors.push(issue('invalid-range', 'The sort end row must not be before the start row.'));
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function sortTarget(workbook: Workbook, args: SortRangeArgs): CellRange[] {
  return [rangeForSort(workbook, args)];
}

function applySort(workbook: Workbook, args: SortRangeArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  const range = rangeForSort(workbook, args);
  const start = range.startRow - 1;
  const end = Math.min(range.endRow, sheet?.rows.length ?? 0);
  const startColumn = columnToIndex(range.startColumn) ?? 0;
  const endColumn = columnToIndex(range.endColumn) ?? startColumn;
  const sortColumn = columnToIndex(args.column) ?? startColumn;
  const rows = sheet?.rows.slice(start, end).map((row) => row.map(cloneCell)) ?? [];
  const sortedRows = rows.slice().sort((left, right) => {
    const comparison = compareValues(
      left[sortColumn]?.value ?? null,
      right[sortColumn]?.value ?? null,
    );
    return args.direction === 'asc' ? comparison : -comparison;
  });
  if (sheet) {
    sortedRows.forEach((source, rowOffset) => {
      const target = sheet.rows[start + rowOffset];
      if (!target) return;
      for (let column = startColumn; column <= endColumn; column += 1) {
        const sourceCell = source[column];
        if (sourceCell) target[column] = cloneCell(sourceCell);
      }
    });
  }
  return transitionResult(before, after, operationReport(before, after, [range]));
}

export const sortRangeOperation: Operation<SortRangeArgs> = {
  name: 'sort_range',
  schema: sortRangeArgsSchema,
  targetRanges: sortTarget,
  validate: validateSort,
  preview(workbook, args) {
    const validation = validateSort(workbook, args);
    const ranges = sortTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applySort(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply(workbook, args) {
    return applySort(workbook, args);
  },
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: sortTarget(before, args),
      rowCountUnchanged: true,
      rowsMultisetEqual: true,
    });
  },
};

export const filterRowsArgsSchema = z
  .object({
    sheet: z.string().trim().min(1),
    column: rangeColumn,
    operator: z.enum([
      'equals',
      'not_equals',
      'contains',
      'starts_with',
      'ends_with',
      'is_blank',
      'is_not_blank',
      'gt',
      'gte',
      'lt',
      'lte',
    ]),
    value: cellValueSchema.optional(),
    headerRow,
  })
  .refine(
    (args) =>
      args.operator === 'is_blank' || args.operator === 'is_not_blank' || args.value !== undefined,
    { message: 'A filter value is required for this operator.', path: ['value'] },
  );
export type FilterRowsArgs = z.infer<typeof filterRowsArgsSchema>;

function filterTarget(workbook: Workbook, args: FilterRowsArgs): CellRange[] {
  return [allDataRange(workbook, args.sheet, args.headerRow)];
}

function matchesFilter(cell: Cell | undefined, args: FilterRowsArgs): boolean {
  const value = cell?.value ?? null;
  const text = textForCell(cell);
  const expected = args.value;
  switch (args.operator) {
    case 'is_blank':
      return value === null || (typeof value === 'string' && value.trim() === '');
    case 'is_not_blank':
      return !(value === null || (typeof value === 'string' && value.trim() === ''));
    case 'equals': {
      if (expected === undefined) return false;
      if (cellValueEquals(value, expected)) return true;
      if (String(value).trim().toLowerCase() === String(expected).trim().toLowerCase()) return true;
      const numVal =
        typeof value === 'number'
          ? value
          : Number(
              String(value ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      const numExp =
        typeof expected === 'number'
          ? expected
          : Number(
              String(expected ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      if (
        !isNaN(numVal) &&
        !isNaN(numExp) &&
        String(value).trim() !== '' &&
        String(expected).trim() !== ''
      ) {
        return numVal === numExp;
      }
      return false;
    }
    case 'not_equals': {
      if (expected === undefined) return false;
      if (cellValueEquals(value, expected)) return false;
      if (String(value).trim().toLowerCase() === String(expected).trim().toLowerCase())
        return false;
      const numVal =
        typeof value === 'number'
          ? value
          : Number(
              String(value ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      const numExp =
        typeof expected === 'number'
          ? expected
          : Number(
              String(expected ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      if (
        !isNaN(numVal) &&
        !isNaN(numExp) &&
        String(value).trim() !== '' &&
        String(expected).trim() !== ''
      ) {
        return numVal !== numExp;
      }
      return true;
    }
    case 'contains':
      return expected !== undefined && text.toLowerCase().includes(String(expected).toLowerCase());
    case 'starts_with':
      return (
        expected !== undefined && text.toLowerCase().startsWith(String(expected).toLowerCase())
      );
    case 'ends_with':
      return expected !== undefined && text.toLowerCase().endsWith(String(expected).toLowerCase());
    case 'gt':
    case 'gte':
    case 'lt':
    case 'lte': {
      if (expected === undefined) return false;
      const numVal =
        typeof value === 'number'
          ? value
          : Number(
              String(value ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      const numExp =
        typeof expected === 'number'
          ? expected
          : Number(
              String(expected ?? '')
                .replace(/,/g, '')
                .trim(),
            );
      if (
        !isNaN(numVal) &&
        !isNaN(numExp) &&
        String(value).trim() !== '' &&
        String(expected).trim() !== ''
      ) {
        if (args.operator === 'gt') return numVal > numExp;
        if (args.operator === 'gte') return numVal >= numExp;
        if (args.operator === 'lt') return numVal < numExp;
        if (args.operator === 'lte') return numVal <= numExp;
      }
      const cmp = compareValues(value, expected);
      if (args.operator === 'gt') return cmp > 0;
      if (args.operator === 'gte') return cmp >= 0;
      if (args.operator === 'lt') return cmp < 0;
      return cmp <= 0;
    }
  }
}

function validateFilter(workbook: Workbook, args: FilterRowsArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateColumn(workbook, args.sheet, args.column),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyFilter(workbook: Workbook, args: FilterRowsArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  const start = args.headerRow;
  const rows = sheet?.rows.slice(start) ?? [];
  const kept = rows.filter((row) => matchesFilter(row[columnToIndex(args.column) ?? 0], args));
  sheet?.rows.splice(start, rows.length, ...kept);
  const removedRows = rows.length - kept.length;
  const ranges = filterTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { removedRows }),
  );
}

export const filterRowsOperation: Operation<FilterRowsArgs> = {
  name: 'filter_rows',
  schema: filterRowsArgsSchema,
  targetRanges: filterTarget,
  validate: validateFilter,
  preview(workbook, args) {
    const validation = validateFilter(workbook, args);
    const ranges = filterTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyFilter(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges, [], [], true);
  },
  apply: applyFilter,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: filterTarget(before, args),
      allowFormulaChanges: true,
    });
  },
};

const findReplaceRangeSchema = z.object({
  startRow: z.number().int().positive(),
  endRow: z.number().int().positive(),
  startColumn: rangeColumn,
  endColumn: rangeColumn,
});
export const findReplaceArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  find: z.string().min(1),
  replace: z.string(),
  matchCase: z.boolean().default(false),
  wholeCell: z.boolean().default(false),
  includeFormulas: z.boolean().default(false),
  range: findReplaceRangeSchema.optional(),
});
export type FindReplaceArgs = z.infer<typeof findReplaceArgsSchema>;

function findTarget(workbook: Workbook, args: FindReplaceArgs): CellRange[] {
  if (args.range) {
    return [
      {
        sheet: args.sheet,
        startRow: args.range.startRow,
        endRow: args.range.endRow,
        startColumn: args.range.startColumn,
        endColumn: args.range.endColumn,
      },
    ];
  }
  return [fullSheetRange(workbook, args.sheet)];
}

function validateFindReplace(workbook: Workbook, args: FindReplaceArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  if (args.range) {
    const startColumn = columnToIndex(args.range.startColumn);
    const endColumn = columnToIndex(args.range.endColumn);
    if (
      args.range.endRow < args.range.startRow ||
      startColumn === undefined ||
      endColumn === undefined ||
      endColumn < (startColumn ?? 0)
    ) {
      errors.push(issue('invalid-range', 'The find/replace range is invalid.'));
    }
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyFindReplace(workbook: Workbook, args: FindReplaceArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const ranges = findTarget(workbook, args);
  let skipped = 0;
  for (const range of ranges) {
    const startColumn = columnToIndex(range.startColumn) ?? 0;
    const endColumn = columnToIndex(range.endColumn) ?? startColumn;
    const sheet = getSheet(after, range.sheet);
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        const current = sheet?.rows[row - 1]?.[column];
        if (!current) continue;
        if (current.formula !== undefined) {
          const formulaMatches = textMatches(
            current.formula,
            args.find,
            args.matchCase,
            args.wholeCell,
          );
          if (!args.includeFormulas) {
            if (formulaMatches) skipped += 1;
            continue;
          }
          if (!formulaMatches) continue;
          const formulaPattern = new RegExp(
            escapeRegExp(args.find),
            args.matchCase ? (args.wholeCell ? '' : 'g') : args.wholeCell ? 'i' : 'gi',
          );
          const nextFormula = current.formula.replace(formulaPattern, args.replace);
          sheet.rows[row - 1]![column] = cloneWithValue(current, current.value, {
            formula: nextFormula,
          });
        } else if (typeof current.value === 'string') {
          if (!textMatches(current.value, args.find, args.matchCase, args.wholeCell)) continue;
          const textPattern = new RegExp(escapeRegExp(args.find), args.matchCase ? 'g' : 'gi');
          const next = args.wholeCell
            ? args.replace
            : current.value.replace(textPattern, args.replace);
          sheet.rows[row - 1]![column] = cloneWithValue(current, next);
        }
      }
    }
  }
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { skippedCells: skipped }),
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export const findReplaceOperation: Operation<FindReplaceArgs> = {
  name: 'find_replace',
  schema: findReplaceArgsSchema,
  targetRanges: findTarget,
  validate: validateFindReplace,
  preview(workbook, args) {
    const validation = validateFindReplace(workbook, args);
    const ranges = findTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyFindReplace(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyFindReplace,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: findTarget(before, args),
      rowCountUnchanged: true,
      allowFormulaChanges: args.includeFormulas,
    });
  },
};

export const deleteDuplicatesArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  columns: z
    .array(rangeColumn)
    .min(1)
    .refine(
      (columns) => new Set(columns.map((column) => column.toUpperCase())).size === columns.length,
      {
        message: 'Duplicate columns are not allowed.',
      },
    ),
  headerRow,
  keep: z.enum(['first', 'last']).default('first'),
});
export type DeleteDuplicatesArgs = z.infer<typeof deleteDuplicatesArgsSchema>;

function duplicateTarget(workbook: Workbook, args: DeleteDuplicatesArgs): CellRange[] {
  return [allDataRange(workbook, args.sheet, args.headerRow)];
}

function validateDuplicates(workbook: Workbook, args: DeleteDuplicatesArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  for (const column of args.columns) errors.push(...validateColumn(workbook, args.sheet, column));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function duplicateKey(row: Cell[], columns: number[]): string {
  return JSON.stringify(
    columns.map((column) => {
      const value = row[column]?.value ?? null;
      return value instanceof Date ? value.toISOString() : value;
    }),
  );
}

function applyDeleteDuplicates(workbook: Workbook, args: DeleteDuplicatesArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  const columns = args.columns.map((column) => columnToIndex(column) ?? 0);
  const rows = sheet?.rows.slice(args.headerRow) ?? [];
  const kept: Cell[][] = [];
  const seen = new Set<string>();
  const source = args.keep === 'last' ? rows.slice().reverse() : rows;
  for (const row of source) {
    const key = duplicateKey(row, columns);
    if (!seen.has(key)) {
      seen.add(key);
      kept.push(row);
    }
  }
  if (args.keep === 'last') kept.reverse();
  sheet?.rows.splice(args.headerRow, rows.length, ...kept);
  const ranges = duplicateTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { removedRows: rows.length - kept.length }),
  );
}

export const deleteDuplicatesOperation: Operation<DeleteDuplicatesArgs> = {
  name: 'delete_duplicates',
  schema: deleteDuplicatesArgsSchema,
  targetRanges: duplicateTarget,
  validate: validateDuplicates,
  preview(workbook, args) {
    const validation = validateDuplicates(workbook, args);
    const ranges = duplicateTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyDeleteDuplicates(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges, [], [], true);
  },
  apply: applyDeleteDuplicates,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: duplicateTarget(before, args),
      allowFormulaChanges: true,
    });
  },
};

export const renameColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  newName: z.string().trim().min(1),
  headerRow,
});
export type RenameColumnArgs = z.infer<typeof renameColumnArgsSchema>;

function renameTarget(_workbook: Workbook, args: RenameColumnArgs): CellRange[] {
  return [
    columnRange(
      _workbook,
      args.sheet,
      args.headerRow,
      columnToIndex(args.column) ?? 0,
      columnToIndex(args.column) ?? 0,
      args.headerRow,
    ),
  ];
}

function validateRename(workbook: Workbook, args: RenameColumnArgs): ValidationResult {
  const errors = [
    ...validateColumn(workbook, args.sheet, args.column),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  const column = columnToIndex(args.column);
  const header =
    column === undefined ? undefined : cellAt(workbook, args.sheet, args.headerRow, column);
  if (header?.formula !== undefined)
    errors.push(issue('formula-header', 'A formula header cannot be renamed implicitly.'));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyRename(workbook: Workbook, args: RenameColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const column = columnToIndex(args.column) ?? 0;
  const sheet = getSheet(after, args.sheet);
  if (sheet?.rows[args.headerRow - 1]) {
    const current = sheet.rows[args.headerRow - 1]![column] ?? createCell(null);
    sheet.rows[args.headerRow - 1]![column] = cloneWithValue(current, args.newName, {
      type: 'string',
    });
  }
  const ranges = renameTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const renameColumnOperation: Operation<RenameColumnArgs> = {
  name: 'rename_column',
  schema: renameColumnArgsSchema,
  targetRanges: renameTarget,
  validate: validateRename,
  preview(workbook, args) {
    const validation = validateRename(workbook, args);
    const ranges = renameTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyRename(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyRename,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: renameTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

export const deleteColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
});
export type DeleteColumnArgs = z.infer<typeof deleteColumnArgsSchema>;

function deleteColumnTarget(workbook: Workbook, args: DeleteColumnArgs): CellRange[] {
  const column = columnToIndex(args.column) ?? 0;
  return [
    columnRange(
      workbook,
      args.sheet,
      1,
      column,
      Math.max(column, maxColumn(workbook, args.sheet) - 1),
    ),
  ];
}

function validateDeleteColumn(workbook: Workbook, args: DeleteColumnArgs): ValidationResult {
  const errors = [...validateColumn(workbook, args.sheet, args.column)];
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyDeleteColumn(workbook: Workbook, args: DeleteColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const column = columnToIndex(args.column) ?? 0;
  const sheet = getSheet(after, args.sheet);
  sheet?.rows.forEach((row) => row.splice(column, 1));
  const ranges = deleteColumnTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { deletedColumns: 1 }),
  );
}

export const deleteColumnOperation: Operation<DeleteColumnArgs> = {
  name: 'delete_column',
  schema: deleteColumnArgsSchema,
  targetRanges: deleteColumnTarget,
  validate: validateDeleteColumn,
  preview(workbook, args) {
    const validation = validateDeleteColumn(workbook, args);
    const ranges = deleteColumnTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyDeleteColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges, [], [], true);
  },
  apply: applyDeleteColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: deleteColumnTarget(before, args),
      rowCountUnchanged: true,
      allowFormulaChanges: true,
    });
  },
};

export const addColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  headerName: z.string().trim().min(1),
  defaultValue: cellValueSchema.optional(),
  headerRow,
});
export type AddColumnArgs = z.infer<typeof addColumnArgsSchema>;

function addColumnTarget(workbook: Workbook, args: AddColumnArgs): CellRange[] {
  const column = columnToIndex(args.column) ?? 0;
  return [
    columnRange(workbook, args.sheet, 1, column, Math.max(column, maxColumn(workbook, args.sheet))),
  ];
}

function validateAddColumn(workbook: Workbook, args: AddColumnArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  const sheet = getSheet(workbook, args.sheet);
  const column = columnToIndex(args.column);
  if (column === undefined)
    errors.push(issue('invalid-column', 'The insertion column is invalid.'));
  if (sheet && column !== undefined && column > maxColumnCount(sheet.rows)) {
    errors.push(
      issue(
        'invalid-column',
        'A column can only be inserted at the end or before an existing column.',
      ),
    );
  }
  errors.push(...headerRowError(workbook, args.sheet, args.headerRow));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyAddColumn(workbook: Workbook, args: AddColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const column = columnToIndex(args.column) ?? 0;
  const sheet = getSheet(after, args.sheet);
  sheet?.rows.forEach((row, rowIndex) => {
    const value = rowIndex + 1 === args.headerRow ? args.headerName : (args.defaultValue ?? null);
    row.splice(column, 0, createCell(value));
  });
  const ranges = addColumnTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { addedColumns: 1 }),
  );
}

export const addColumnOperation: Operation<AddColumnArgs> = {
  name: 'add_column',
  schema: addColumnArgsSchema,
  targetRanges: addColumnTarget,
  validate: validateAddColumn,
  preview(workbook, args) {
    const validation = validateAddColumn(workbook, args);
    const ranges = addColumnTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyAddColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyAddColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: addColumnTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

export const normalizeTextArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  columns: z.array(rangeColumn).min(1).optional(),
  headerRow,
  trim: z.boolean().default(true),
  collapseWhitespace: z.boolean().default(true),
  case: z.enum(['none', 'lower', 'upper', 'title']).default('none'),
});
export type NormalizeTextArgs = z.infer<typeof normalizeTextArgsSchema>;

function normalizeColumns(workbook: Workbook, args: NormalizeTextArgs): string[] {
  return (
    args.columns ??
    Array.from({ length: maxColumn(workbook, args.sheet) }, (_, index) => indexToColumn(index))
  );
}

function normalizeTarget(workbook: Workbook, args: NormalizeTextArgs): CellRange[] {
  return rangeForColumns(workbook, args.sheet, args.headerRow, normalizeColumns(workbook, args));
}

function validateNormalize(workbook: Workbook, args: NormalizeTextArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  for (const column of normalizeColumns(workbook, args))
    errors.push(...validateColumn(workbook, args.sheet, column));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function normalizeValue(value: string, args: NormalizeTextArgs): string {
  let result = value;
  if (args.trim) result = result.trim();
  if (args.collapseWhitespace) result = result.replace(/\s+/g, ' ');
  if (args.case === 'lower') result = result.toLocaleLowerCase();
  if (args.case === 'upper') result = result.toLocaleUpperCase();
  if (args.case === 'title') {
    result = result
      .toLocaleLowerCase()
      .replace(/\b\w/g, (character) => character.toLocaleUpperCase());
  }
  return result;
}

function applyNormalize(workbook: Workbook, args: NormalizeTextArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const ranges = normalizeTarget(workbook, args);
  let skipped = 0;
  const sheet = getSheet(after, args.sheet);
  for (const range of ranges) {
    const start = columnToIndex(range.startColumn) ?? 0;
    const end = columnToIndex(range.endColumn) ?? start;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = start; column <= end; column += 1) {
        const current = sheet?.rows[row - 1]?.[column];
        if (!current || current.formula !== undefined || typeof current.value !== 'string') {
          skipped += 1;
          continue;
        }
        const next = normalizeValue(current.value, args);
        sheet.rows[row - 1]![column] = cloneWithValue(current, next);
      }
    }
  }
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { skippedCells: skipped }),
  );
}

export const normalizeTextOperation: Operation<NormalizeTextArgs> = {
  name: 'normalize_text',
  schema: normalizeTextArgsSchema,
  targetRanges: normalizeTarget,
  validate: validateNormalize,
  preview(workbook, args) {
    const validation = validateNormalize(workbook, args);
    const ranges = normalizeTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyNormalize(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyNormalize,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: normalizeTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

const setCellArgsSchema = z.object({
  row: z.number().int().positive(),
  column: rangeColumn,
  value: cellValueSchema,
  formula: z.string().optional(),
  numberFormat: z.string().optional(),
});
export const setCellsArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  cells: z
    .array(setCellArgsSchema)
    .min(1)
    .refine(
      (cells) =>
        new Set(cells.map((cell) => `${cell.row}:${cell.column.toUpperCase()}`)).size ===
        cells.length,
      { message: 'Each cell address may only appear once.', path: ['cells'] },
    ),
});
export type SetCellsArgs = z.infer<typeof setCellsArgsSchema>;

function setCellsTarget(_workbook: Workbook, args: SetCellsArgs): CellRange[] {
  return args.cells.map((cell) => ({
    sheet: args.sheet,
    startRow: cell.row,
    endRow: cell.row,
    startColumn: cell.column,
    endColumn: cell.column,
  }));
}

function validateSetCells(workbook: Workbook, args: SetCellsArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  const sheet = getSheet(workbook, args.sheet);
  for (const cell of args.cells) {
    const column = columnToIndex(cell.column);
    if (column === undefined || column >= maxColumnCount(sheet?.rows ?? [])) {
      errors.push(
        issue(
          'missing-cell',
          `Cell ${args.sheet}!${cell.column}${cell.row} is outside the used range.`,
        ),
      );
    }
    if (!sheet?.rows[cell.row - 1]) {
      errors.push(
        issue(
          'missing-cell',
          `Cell ${args.sheet}!${cell.column}${cell.row} is outside the used range.`,
        ),
      );
    }
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applySetCells(workbook: Workbook, args: SetCellsArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  for (const requested of args.cells) {
    const column = columnToIndex(requested.column) ?? 0;
    const current = sheet?.rows[requested.row - 1]?.[column] ?? createCell(null);
    const next = cloneWithValue(current, requested.value, {
      ...(requested.formula !== undefined ? { formula: requested.formula } : {}),
      ...(requested.numberFormat !== undefined ? { numberFormat: requested.numberFormat } : {}),
    });
    if (sheet?.rows[requested.row - 1]) sheet.rows[requested.row - 1]![column] = next;
  }
  const ranges = setCellsTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const setCellsOperation: Operation<SetCellsArgs> = {
  name: 'set_cells',
  schema: setCellsArgsSchema,
  targetRanges: setCellsTarget,
  validate: validateSetCells,
  preview(workbook, args) {
    const validation = validateSetCells(workbook, args);
    const ranges = setCellsTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applySetCells(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applySetCells,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: setCellsTarget(before, args),
      rowCountUnchanged: true,
      allowFormulaChanges: true,
    });
  },
};

export const initialOperations: Operation<unknown>[] = [
  sortRangeOperation,
  filterRowsOperation,
  findReplaceOperation,
  deleteDuplicatesOperation,
  renameColumnOperation,
  deleteColumnOperation,
  addColumnOperation,
  normalizeTextOperation,
  setCellsOperation,
];
