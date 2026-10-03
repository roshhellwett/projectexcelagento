import { z } from 'zod';

import type {
  Cell,
  CellRange,
  CellValue,
  Operation,
  OperationResult,
  Preview,
  Report,
  ValidationIssue,
  ValidationResult,
  Workbook,
} from './types.js';
import {
  cellValueSchema,
  cloneWithValue,
  maxColumn,
  maxRow,
  previewForTransition,
  transitionResult,
  validateColumn,
  validateSheet,
  validResult,
} from './operation-utils.js';
import { runInvariants } from './invariants.js';
import { columnToIndex, createCell, getSheet, indexToColumn } from './workbook.js';

function operationReport(
  before: Workbook,
  after: Workbook,
  ranges: CellRange[],
  warnings: ValidationIssue[] = [],
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

const rangeColumn = z
  .string()
  .trim()
  .min(1)
  .regex(/^[A-Za-z]+$/, 'Columns must be letters');
const headerRowSchema = z.number().int().positive().default(1);

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

function isBlank(val: CellValue | undefined): boolean {
  if (val === null || val === undefined) return true;
  if (typeof val === 'string' && val.trim() === '') return true;
  return false;
}

// ============================================================================
// 1. FILL BLANKS (fill_blanks)
// ============================================================================

export const fillBlanksArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  strategy: z.enum(['value', 'forward', 'backward', 'mean']).default('forward'),
  fillValue: cellValueSchema.optional(),
  headerRow: headerRowSchema,
});
export type FillBlanksArgs = z.infer<typeof fillBlanksArgsSchema>;

function fillBlanksTarget(workbook: Workbook, args: FillBlanksArgs): CellRange[] {
  const colIdx = columnToIndex(args.column) ?? 0;
  return [columnRange(workbook, args.sheet, args.headerRow + 1, colIdx, colIdx)];
}

function validateFillBlanks(workbook: Workbook, args: FillBlanksArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  errors.push(...validateColumn(workbook, args.sheet, args.column));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyFillBlanks(workbook: Workbook, args: FillBlanksArgs): OperationResult {
  const before = workbook;
  const after = JSON.parse(JSON.stringify(workbook)) as Workbook;
  const sheet = getSheet(after, args.sheet);
  const colIdx = columnToIndex(args.column) ?? 0;

  if (sheet) {
    const startRow = args.headerRow; // 0-indexed index of first data row is args.headerRow
    const totalRows = sheet.rows.length;

    if (args.strategy === 'forward') {
      let lastVal: CellValue = null;
      for (let r = startRow; r < totalRows; r++) {
        const cell = sheet.rows[r]?.[colIdx];
        if (!cell || isBlank(cell.value)) {
          if (lastVal !== null && sheet.rows[r]) {
            sheet.rows[r]![colIdx] = cloneWithValue(cell ?? createCell(null), lastVal);
          }
        } else {
          lastVal = cell.value;
        }
      }
    } else if (args.strategy === 'backward') {
      let nextVal: CellValue = null;
      for (let r = totalRows - 1; r >= startRow; r--) {
        const cell = sheet.rows[r]?.[colIdx];
        if (!cell || isBlank(cell.value)) {
          if (nextVal !== null && sheet.rows[r]) {
            sheet.rows[r]![colIdx] = cloneWithValue(cell ?? createCell(null), nextVal);
          }
        } else {
          nextVal = cell.value;
        }
      }
    } else if (args.strategy === 'mean') {
      let sum = 0;
      let count = 0;
      for (let r = startRow; r < totalRows; r++) {
        const val = sheet.rows[r]?.[colIdx]?.value;
        if (typeof val === 'number' && !isNaN(val)) {
          sum += val;
          count++;
        }
      }
      const meanVal = count > 0 ? Math.round((sum / count) * 100) / 100 : 0;
      for (let r = startRow; r < totalRows; r++) {
        const cell = sheet.rows[r]?.[colIdx];
        if (!cell || isBlank(cell.value)) {
          if (sheet.rows[r]) {
            sheet.rows[r]![colIdx] = cloneWithValue(cell ?? createCell(null), meanVal);
          }
        }
      }
    } else if (args.strategy === 'value') {
      const fallback = args.fillValue ?? '';
      for (let r = startRow; r < totalRows; r++) {
        const cell = sheet.rows[r]?.[colIdx];
        if (!cell || isBlank(cell.value)) {
          if (sheet.rows[r]) {
            sheet.rows[r]![colIdx] = cloneWithValue(cell ?? createCell(null), fallback);
          }
        }
      }
    }
  }

  const ranges = fillBlanksTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const fillBlanksOperation: Operation<FillBlanksArgs> = {
  name: 'fill_blanks',
  schema: fillBlanksArgsSchema,
  targetRanges: fillBlanksTarget,
  validate: validateFillBlanks,
  preview(workbook, args) {
    const validation = validateFillBlanks(workbook, args);
    const ranges = fillBlanksTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyFillBlanks(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyFillBlanks,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: fillBlanksTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

// ============================================================================
// 2. ADD COMPUTED COLUMN (add_computed_column)
// ============================================================================

export const addComputedColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  headerName: z.string().trim().min(1),
  expression: z.string().trim().min(1),
  afterColumn: rangeColumn.optional(),
  headerRow: headerRowSchema,
});
export type AddComputedColumnArgs = z.infer<typeof addComputedColumnArgsSchema>;

function addComputedColumnTarget(workbook: Workbook, args: AddComputedColumnArgs): CellRange[] {
  const col = args.afterColumn
    ? (columnToIndex(args.afterColumn) ?? maxColumn(workbook, args.sheet)) + 1
    : maxColumn(workbook, args.sheet);
  return [columnRange(workbook, args.sheet, 1, col, maxColumn(workbook, args.sheet) + 1)];
}

function validateAddComputedColumn(
  workbook: Workbook,
  args: AddComputedColumnArgs,
): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  if (args.afterColumn) {
    errors.push(...validateColumn(workbook, args.sheet, args.afterColumn));
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

/**
 * Evaluates an expression against a single row's cell values.
 * Supports:
 * - Column identifiers: A, B, Col C, col('Header'), [Header]
 * - Operators: +, -, *, /, %
 * - Math/string functions: round, floor, ceil, abs, upper, lower, trim, concat
 */
function evaluateRowExpression(
  expression: string,
  row: Cell[],
  headerMap: Map<string, number>,
): CellValue {
  let expr = expression;

  // 1. Resolve named column brackets [Col Name] or col('Col Name') or col("Col Name")
  expr = expr.replace(/col\(\s*['"]([^'"]+)['"]\s*\)/gi, (_, colName: string) => {
    return resolveColValue(colName, row, headerMap);
  });
  expr = expr.replace(/\[([^\]]+)\]/g, (_, colName: string) => {
    return resolveColValue(colName, row, headerMap);
  });

  // 2. Resolve single/double letter columns like A, B, C, D (case-insensitive) preceded by boundary
  expr = expr.replace(/\b([A-Z]{1,2})\b(?!\()/gi, (match) => {
    const colIdx = columnToIndex(match);
    if (colIdx !== undefined && colIdx < row.length) {
      const val = row[colIdx]?.value;
      if (typeof val === 'number') return String(val);
      if (typeof val === 'string') return JSON.stringify(val);
      if (val === null || val === undefined) return '0';
    }
    return match;
  });

  // 3. Resolve string helpers: upper("..."), lower("..."), trim("..."), concat(...)
  expr = expr.replace(/upper\(([^)]+)\)/gi, (_, arg: string) => {
    const unquoted = arg.replace(/^['"]|['"]$/g, '');
    return JSON.stringify(unquoted.toUpperCase());
  });
  expr = expr.replace(/lower\(([^)]+)\)/gi, (_, arg: string) => {
    const unquoted = arg.replace(/^['"]|['"]$/g, '');
    return JSON.stringify(unquoted.toLowerCase());
  });
  expr = expr.replace(/trim\(([^)]+)\)/gi, (_, arg: string) => {
    const unquoted = arg.replace(/^['"]|['"]$/g, '');
    return JSON.stringify(unquoted.trim());
  });
  expr = expr.replace(/concat\(([^)]+)\)/gi, (_, args: string) => {
    const parts = args.split(',').map((p) => {
      const trimmed = p.trim();
      return trimmed.replace(/^['"]|['"]$/g, '');
    });
    return JSON.stringify(parts.join(''));
  });

  // 4. Resolve math functions: round(x, decimals), floor(x), ceil(x), abs(x)
  expr = expr.replace(
    /round\(\s*([^,]+)\s*,\s*(\d+)\s*\)/gi,
    (_, numStr: string, decStr: string) => {
      const num = safeEvalArithmetic(numStr);
      const decimals = parseInt(decStr, 10);
      const factor = Math.pow(10, decimals);
      return String(Math.round(num * factor) / factor);
    },
  );
  expr = expr.replace(/floor\(\s*([^)]+)\s*\)/gi, (_, numStr: string) => {
    return String(Math.floor(safeEvalArithmetic(numStr)));
  });
  expr = expr.replace(/ceil\(\s*([^)]+)\s*\)/gi, (_, numStr: string) => {
    return String(Math.ceil(safeEvalArithmetic(numStr)));
  });
  expr = expr.replace(/abs\(\s*([^)]+)\s*\)/gi, (_, numStr: string) => {
    return String(Math.abs(safeEvalArithmetic(numStr)));
  });

  // 5. If it's a string literal or string concat
  if (expr.includes('"') || expr.includes("'")) {
    const stringConcat = expr
      .split('+')
      .map((part) => part.trim().replace(/^['"]|['"]$/g, ''))
      .join('');
    return stringConcat;
  }

  // 6. Otherwise evaluate arithmetic
  const num = safeEvalArithmetic(expr);
  return isNaN(num) ? null : Math.round(num * 10000) / 10000;
}

function resolveColValue(
  colIdentifier: string,
  row: Cell[],
  headerMap: Map<string, number>,
): string {
  let idx = headerMap.get(colIdentifier.toLowerCase().trim());
  if (idx === undefined) {
    const directIdx = columnToIndex(colIdentifier);
    if (directIdx !== undefined) idx = directIdx;
  }
  if (idx !== undefined && idx < row.length) {
    const cellVal = row[idx]?.value;
    if (typeof cellVal === 'number') return String(cellVal);
    if (typeof cellVal === 'string') return JSON.stringify(cellVal);
    if (cellVal === null || cellVal === undefined) return '0';
  }
  return '0';
}

function safeEvalArithmetic(str: string): number {
  const sanitized = str.replace(/[^0-9+\-*/().% ]/g, '');
  if (!sanitized.trim()) return 0;
  try {
    const fn = new Function(`return (${sanitized});`);
    const val = fn();
    return typeof val === 'number' && isFinite(val) ? val : 0;
  } catch {
    return 0;
  }
}

function applyAddComputedColumn(workbook: Workbook, args: AddComputedColumnArgs): OperationResult {
  const before = workbook;
  const after = JSON.parse(JSON.stringify(workbook)) as Workbook;
  const sheet = getSheet(after, args.sheet);

  if (sheet) {
    const insertionCol = args.afterColumn
      ? (columnToIndex(args.afterColumn) ?? sheet.rows[0]?.length ?? 0) + 1
      : (sheet.rows[0]?.length ?? 0);

    // Build header map
    const headerRowIdx = args.headerRow - 1;
    const headerRow = sheet.rows[headerRowIdx] ?? [];
    const headerMap = new Map<string, number>();
    headerRow.forEach((cell, idx) => {
      if (cell.value !== null && cell.value !== undefined) {
        headerMap.set(String(cell.value).toLowerCase().trim(), idx);
      }
    });

    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r]!;
      if (r === headerRowIdx) {
        row.splice(insertionCol, 0, createCell(args.headerName));
      } else if (r > headerRowIdx) {
        const computed = evaluateRowExpression(args.expression, row, headerMap);
        row.splice(insertionCol, 0, createCell(computed));
      } else {
        row.splice(insertionCol, 0, createCell(null));
      }
    }
  }

  const ranges = addComputedColumnTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { addedColumns: 1 }),
  );
}

export const addComputedColumnOperation: Operation<AddComputedColumnArgs> = {
  name: 'add_computed_column',
  schema: addComputedColumnArgsSchema,
  targetRanges: addComputedColumnTarget,
  validate: validateAddComputedColumn,
  preview(workbook, args) {
    const validation = validateAddComputedColumn(workbook, args);
    const ranges = addComputedColumnTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyAddComputedColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyAddComputedColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: addComputedColumnTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

// ============================================================================
// 3. SPLIT COLUMN (split_column)
// ============================================================================

export const splitColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  delimiter: z.string().min(1),
  newColumnNames: z.array(z.string().trim().min(1)).min(2).optional(),
  headerRow: headerRowSchema,
});
export type SplitColumnArgs = z.infer<typeof splitColumnArgsSchema>;

function splitColumnTarget(workbook: Workbook, args: SplitColumnArgs): CellRange[] {
  const colIdx = columnToIndex(args.column) ?? 0;
  return [columnRange(workbook, args.sheet, 1, colIdx, maxColumn(workbook, args.sheet) + 2)];
}

function validateSplitColumn(workbook: Workbook, args: SplitColumnArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  errors.push(...validateColumn(workbook, args.sheet, args.column));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applySplitColumn(workbook: Workbook, args: SplitColumnArgs): OperationResult {
  const before = workbook;
  const after = JSON.parse(JSON.stringify(workbook)) as Workbook;
  const sheet = getSheet(after, args.sheet);
  const colIdx = columnToIndex(args.column) ?? 0;

  if (sheet) {
    const headerRowIdx = args.headerRow - 1;
    const origHeader = sheet.rows[headerRowIdx]?.[colIdx]?.value ?? 'Split';
    const names = args.newColumnNames ?? [`${origHeader} 1`, `${origHeader} 2`];
    const partsCount = names.length;

    // For the original column, it receives part 0. New columns are inserted for parts 1..n-1
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r]!;
      if (r === headerRowIdx) {
        row[colIdx] = createCell(names[0] ?? 'Split 1');
        for (let i = 1; i < partsCount; i++) {
          row.splice(colIdx + i, 0, createCell(names[i] ?? `Split ${i + 1}`));
        }
      } else if (r > headerRowIdx) {
        const cellVal = row[colIdx]?.value;
        const parts =
          cellVal !== null && cellVal !== undefined ? String(cellVal).split(args.delimiter) : [];
        const p0 = parts[0]?.trim();
        row[colIdx] = createCell(p0 ? p0 : null);
        for (let i = 1; i < partsCount; i++) {
          const pi = parts[i]?.trim();
          row.splice(colIdx + i, 0, createCell(pi ? pi : null));
        }
      } else {
        for (let i = 1; i < partsCount; i++) {
          row.splice(colIdx + i, 0, createCell(null));
        }
      }
    }
  }

  const ranges = splitColumnTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], {
      addedColumns: (args.newColumnNames?.length ?? 2) - 1,
    }),
  );
}

export const splitColumnOperation: Operation<SplitColumnArgs> = {
  name: 'split_column',
  schema: splitColumnArgsSchema,
  targetRanges: splitColumnTarget,
  validate: validateSplitColumn,
  preview(workbook, args) {
    const validation = validateSplitColumn(workbook, args);
    const ranges = splitColumnTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applySplitColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applySplitColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: splitColumnTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

// ============================================================================
// 4. MERGE COLUMNS (merge_columns)
// ============================================================================

export const mergeColumnsArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  columns: z.array(rangeColumn).min(2),
  separator: z.string().default(' '),
  headerName: z.string().trim().min(1),
  headerRow: headerRowSchema,
});
export type MergeColumnsArgs = z.infer<typeof mergeColumnsArgsSchema>;

function mergeColumnsTarget(workbook: Workbook, args: MergeColumnsArgs): CellRange[] {
  const lastCol = maxColumn(workbook, args.sheet);
  return [columnRange(workbook, args.sheet, 1, lastCol, lastCol + 1)];
}

function validateMergeColumns(workbook: Workbook, args: MergeColumnsArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  for (const c of args.columns) {
    errors.push(...validateColumn(workbook, args.sheet, c));
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyMergeColumns(workbook: Workbook, args: MergeColumnsArgs): OperationResult {
  const before = workbook;
  const after = JSON.parse(JSON.stringify(workbook)) as Workbook;
  const sheet = getSheet(after, args.sheet);

  if (sheet) {
    const colIndices = args.columns
      .map((c) => columnToIndex(c))
      .filter((idx): idx is number => idx !== undefined);
    const headerRowIdx = args.headerRow - 1;
    const destIdx = sheet.rows[0]?.length ?? 0;

    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r]!;
      if (r === headerRowIdx) {
        row.splice(destIdx, 0, createCell(args.headerName));
      } else if (r > headerRowIdx) {
        const mergedValue = colIndices
          .map((ci) => {
            const val = row[ci]?.value;
            return val !== null && val !== undefined ? String(val).trim() : '';
          })
          .filter(Boolean)
          .join(args.separator);
        row.splice(destIdx, 0, createCell(mergedValue));
      } else {
        row.splice(destIdx, 0, createCell(null));
      }
    }
  }

  const ranges = mergeColumnsTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { addedColumns: 1 }),
  );
}

export const mergeColumnsOperation: Operation<MergeColumnsArgs> = {
  name: 'merge_columns',
  schema: mergeColumnsArgsSchema,
  targetRanges: mergeColumnsTarget,
  validate: validateMergeColumns,
  preview(workbook, args) {
    const validation = validateMergeColumns(workbook, args);
    const ranges = mergeColumnsTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyMergeColumns(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyMergeColumns,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: mergeColumnsTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

// ============================================================================
// 5. LOOKUP MERGE (lookup_merge / VLOOKUP)
// ============================================================================

export const lookupMergeArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  keyColumn: rangeColumn,
  lookupSheet: z.string().trim().min(1),
  lookupKeyColumn: rangeColumn,
  lookupValueColumn: rangeColumn,
  headerName: z.string().trim().min(1),
  headerRow: headerRowSchema,
});
export type LookupMergeArgs = z.infer<typeof lookupMergeArgsSchema>;

function lookupMergeTarget(workbook: Workbook, args: LookupMergeArgs): CellRange[] {
  const destCol = maxColumn(workbook, args.sheet);
  return [columnRange(workbook, args.sheet, 1, destCol, destCol + 1)];
}

function validateLookupMerge(workbook: Workbook, args: LookupMergeArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  errors.push(...validateSheet(workbook, args.lookupSheet));
  errors.push(...validateColumn(workbook, args.sheet, args.keyColumn));
  errors.push(...validateColumn(workbook, args.lookupSheet, args.lookupKeyColumn));
  errors.push(...validateColumn(workbook, args.lookupSheet, args.lookupValueColumn));
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyLookupMerge(workbook: Workbook, args: LookupMergeArgs): OperationResult {
  const before = workbook;
  const after = JSON.parse(JSON.stringify(workbook)) as Workbook;
  const targetSheet = getSheet(after, args.sheet);
  const lookupSheet = getSheet(after, args.lookupSheet);

  if (targetSheet && lookupSheet) {
    const keyColIdx = columnToIndex(args.keyColumn) ?? 0;
    const lookupKeyIdx = columnToIndex(args.lookupKeyColumn) ?? 0;
    const lookupValIdx = columnToIndex(args.lookupValueColumn) ?? 0;
    const headerRowIdx = args.headerRow - 1;

    // Build lookup map from lookupSheet
    const lookupMap = new Map<string, CellValue>();
    for (let r = headerRowIdx + 1; r < lookupSheet.rows.length; r++) {
      const k = lookupSheet.rows[r]?.[lookupKeyIdx]?.value;
      const v = lookupSheet.rows[r]?.[lookupValIdx]?.value;
      if (k !== null && k !== undefined) {
        lookupMap.set(String(k).trim().toLowerCase(), v ?? null);
      }
    }

    const destColIdx = targetSheet.rows[0]?.length ?? 0;
    for (let r = 0; r < targetSheet.rows.length; r++) {
      const row = targetSheet.rows[r]!;
      if (r === headerRowIdx) {
        row.splice(destColIdx, 0, createCell(args.headerName));
      } else if (r > headerRowIdx) {
        const k = row[keyColIdx]?.value;
        const matched =
          k !== null && k !== undefined
            ? (lookupMap.get(String(k).trim().toLowerCase()) ?? null)
            : null;
        row.splice(destColIdx, 0, createCell(matched));
      } else {
        row.splice(destColIdx, 0, createCell(null));
      }
    }
  }

  const ranges = lookupMergeTarget(workbook, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, ranges, [], { addedColumns: 1 }),
  );
}

export const lookupMergeOperation: Operation<LookupMergeArgs> = {
  name: 'lookup_merge',
  schema: lookupMergeArgsSchema,
  targetRanges: lookupMergeTarget,
  validate: validateLookupMerge,
  preview(workbook, args) {
    const validation = validateLookupMerge(workbook, args);
    const ranges = lookupMergeTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyLookupMerge(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyLookupMerge,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: lookupMergeTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

export const advancedOperations: Operation<unknown>[] = [
  fillBlanksOperation as Operation<unknown>,
  addComputedColumnOperation as Operation<unknown>,
  splitColumnOperation as Operation<unknown>,
  mergeColumnsOperation as Operation<unknown>,
  lookupMergeOperation as Operation<unknown>,
];
