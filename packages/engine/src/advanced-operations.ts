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
  allDataRange,
  cellValueSchema,
  cloneWithValue,
  fullSheetRange,
  headerRowError,
  issue,
  maxColumn,
  maxRow,
  previewForTransition,
  transitionResult,
  uniqueSheetName,
  validateColumn,
  validateSheet,
  validResult,
} from './operation-utils.js';
import { runInvariants } from './invariants.js';
import { matchesFilter } from './operations.js';
import {
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
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
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
  const slots: string[] = [];
  const protect = (value: string): string => {
    slots.push(value);
    return '\u0000' + String(slots.length - 1) + '\u0000';
  };
  const unprotect = (token: string): string => {
    const t = token.trim();
    if (t.charCodeAt(0) !== 0 || t.charCodeAt(t.length - 1) !== 0) return '';
    const inner = t.slice(1, t.length - 1);
    return /^\d+$/.test(inner) ? (slots[Number(inner)] ?? '') : '';
  };
  const isPlaceholder = (token: string): boolean => {
    const t = token.trim();
    return (
      t.charCodeAt(0) === 0 && t.charCodeAt(t.length - 1) === 0 && /^\d+$/.test(t.slice(1, -1))
    );
  };

  let expr = expression.trim();

  // 1. Resolve named column brackets [Col Name] or col('Col Name') (before protecting literals)
  const resolveColToken = (colName: string): string => {
    let idx = headerMap.get(colName.toLowerCase().trim());
    if (idx === undefined) {
      const directIdx = columnToIndex(colName);
      if (directIdx !== undefined) idx = directIdx;
    }
    if (idx !== undefined && idx < row.length) {
      const cellVal = row[idx]?.value;
      if (typeof cellVal === 'number') return String(cellVal);
      if (typeof cellVal === 'string') return protect(cellVal);
      return '0';
    }
    return '0';
  };
  expr = expr.replace(/col\(\s*['"]([^'"]+)['"]\s*\)/gi, (_m, n: string) => resolveColToken(n));
  expr = expr.replace(/\[([^\]]+)\]/g, (_m, n: string) => resolveColToken(n));

  // 2. Protect string literals so column resolution never rewrites letters inside them
  expr = expr.replace(/"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'/g, (_m, a, b) =>
    protect(a !== undefined ? a : b),
  );

  // 3. Resolve bare single/double letter columns (A, B, ...). String literals are protected.
  expr = expr.replace(/\b([A-Za-z]{1,2})\b(?!\()/g, (match) => {
    const colIdx = columnToIndex(match);
    if (colIdx !== undefined && colIdx < row.length) {
      const val = row[colIdx]?.value;
      if (typeof val === 'number') return String(val);
      if (typeof val === 'string') return protect(val);
      return '0';
    }
    return match;
  });

  // 4. String helpers operate on resolved operands (literal or column value)
  const resolveOperand = (arg: string): string => {
    const t = arg.trim();
    if (isPlaceholder(t)) return unprotect(t);
    return t.replace(/^['"]|['"]$/g, '');
  };
  expr = expr.replace(/upper\(([^)]+)\)/gi, (_m, arg: string) =>
    protect(resolveOperand(arg).toUpperCase()),
  );
  expr = expr.replace(/lower\(([^)]+)\)/gi, (_m, arg: string) =>
    protect(resolveOperand(arg).toLowerCase()),
  );
  expr = expr.replace(/trim\(([^)]+)\)/gi, (_m, arg: string) =>
    protect(resolveOperand(arg).trim()),
  );
  expr = expr.replace(/concat\(([^)]+)\)/gi, (_m, args: string) =>
    protect(
      args
        .split(',')
        .map((p) => resolveOperand(p))
        .join(''),
    ),
  );

  // 5. Math functions
  expr = expr.replace(
    /round\(\s*([^,]+)\s*,\s*(\d+)\s*\)/gi,
    (_m, numStr: string, decStr: string) => {
      const num = safeEvalArithmetic(numStr);
      const decimals = parseInt(decStr, 10);
      const factor = Math.pow(10, decimals);
      return String(Math.round(num * factor) / factor);
    },
  );
  expr = expr.replace(/floor\(\s*([^)]+)\s*\)/gi, (_m, numStr: string) =>
    String(Math.floor(safeEvalArithmetic(numStr))),
  );
  expr = expr.replace(/ceil\(\s*([^)]+)\s*\)/gi, (_m, numStr: string) =>
    String(Math.ceil(safeEvalArithmetic(numStr))),
  );
  expr = expr.replace(/abs\(\s*([^)]+)\s*\)/gi, (_m, numStr: string) =>
    String(Math.abs(safeEvalArithmetic(numStr))),
  );

  // 6. Any remaining string operand -> string concatenation semantics
  if (expr.includes('\u0000')) {
    return expr
      .split('+')
      .map((part) => (isPlaceholder(part) ? unprotect(part) : part.trim()))
      .join('');
  }

  // 7. Otherwise arithmetic
  const num = safeEvalArithmetic(expr);
  return isNaN(num) ? null : Math.round(num * 10000) / 10000;
}

function safeEvalArithmetic(str: string): number {
  const sanitized = str.replace(/[^0-9+\-*/().%\s]/g, '');
  if (!sanitized.trim()) return 0;
  let i = 0;
  const skipWs = () => {
    while (i < sanitized.length && sanitized[i] === ' ') i++;
  };
  const parseExpr = (): number => {
    let v = parseTerm();
    for (;;) {
      skipWs();
      if (sanitized[i] === '+') {
        i++;
        v = v + parseTerm();
      } else if (sanitized[i] === '-') {
        i++;
        v = v - parseTerm();
      } else {
        return v;
      }
    }
  };
  const parseTerm = (): number => {
    let v = parsePowerless();
    for (;;) {
      skipWs();
      if (sanitized[i] === '*') {
        i++;
        v = v * parsePowerless();
      } else if (sanitized[i] === '/') {
        i++;
        v = v / parsePowerless();
      } else {
        return v;
      }
    }
  };
  const parsePowerless = (): number => {
    let v = parseFactor();
    for (;;) {
      skipWs();
      if (sanitized[i] === '%') {
        v = v / 100;
        i++;
      } else {
        return v;
      }
    }
  };
  const parseFactor = (): number => {
    skipWs();
    if (sanitized[i] === '-') {
      i++;
      return -parseFactor();
    }
    if (sanitized[i] === '+') {
      i++;
      return parseFactor();
    }
    if (sanitized[i] === '(') {
      i++;
      const v = parseExpr();
      skipWs();
      if (sanitized[i] === ')') i++;
      return v;
    }
    let num = '';
    while (i < sanitized.length && /[0-9.]/.test(sanitized[i]!)) num += sanitized[i++]!;
    const v = num ? parseFloat(num) : NaN;
    return v;
  };
  try {
    const v = parseExpr();
    return typeof v === 'number' && isFinite(v) ? v : 0;
  } catch {
    return 0;
  }
}

function applyAddComputedColumn(workbook: Workbook, args: AddComputedColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
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
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
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
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
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
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
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

// ============================================================================
// 6. CLEAN TO NEW SHEET (clean_to_new_sheet)
// ============================================================================

export const cleanToNewSheetArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  targetSheet: z.string().trim().min(1),
  /** 1-based row index of the header in the source. Detected automatically when omitted. */
  headerRow: z.number().int().min(1).optional(),
  trim: z.boolean().default(true),
  collapseWhitespace: z.boolean().default(true),
  dropEmptyRows: z.boolean().default(true),
  dropEmptyColumns: z.boolean().default(true),
  coerceNumbers: z.boolean().default(true),
});
export type CleanToNewSheetArgs = z.infer<typeof cleanToNewSheetArgsSchema>;

function cellIsBlank(cell: Cell | undefined): boolean {
  const v = cell?.value;
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

function coerceCell(cell: Cell, args: CleanToNewSheetArgs): Cell {
  if (cell.formula !== undefined) return cell;
  let value = cell.value;
  if (typeof value === 'string') {
    let v = value;
    if (args.trim) v = v.trim();
    if (args.collapseWhitespace) v = v.replace(/\s+/g, ' ');
    if (args.coerceNumbers) {
      const normalized = v.replace(/,/g, '').replace(/\$/g, '').replace(/%$/, '');
      const parens = /^\((.*)\)$/.exec(normalized);
      const numeric = parens ? `-${parens[1]}` : normalized;
      if (numeric.trim() !== '' && !isNaN(Number(numeric))) {
        value = Number(numeric);
      } else {
        value = v;
      }
    } else {
      value = v;
    }
  }
  return createCell(value, { numberFormat: cell.numberFormat });
}

function detectHeaderRowIndex(rows: Cell[][]): number {
  let bestIdx = 0;
  let bestCount = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
    const nonBlank = (rows[i] ?? []).filter((c) => !cellIsBlank(c)).length;
    if (nonBlank > bestCount) {
      bestCount = nonBlank;
      bestIdx = i;
    }
  }
  return bestIdx;
}

/** Shared cleaning pipeline used by both apply() and targetRanges(). */
function buildCleanedRows(sourceRows: Cell[][], args: CleanToNewSheetArgs): Cell[][] {
  const cleaned = sourceRows.map((row) => row.map((cell) => coerceCell(cell, args)));
  const headerIdx =
    args.headerRow !== undefined
      ? Math.min(args.headerRow - 1, Math.max(0, cleaned.length - 1))
      : detectHeaderRowIndex(cleaned);
  let rows = cleaned.slice(headerIdx);

  if (args.dropEmptyRows) {
    rows = rows.filter((row, idx) => {
      if (idx === 0) return true;
      const nonBlank = row.filter((cell) => !cellIsBlank(cell));
      if (nonBlank.length === 0) return false;
      // Section banners like "Cash Flow statement" float alone in one cell — drop them.
      if (nonBlank.length === 1 && typeof nonBlank[0]!.value === 'string') return false;
      return true;
    });
  }
  if (args.dropEmptyColumns) {
    const width = Math.max(0, ...rows.map((row) => row.length));
    const keepColumns: number[] = [];
    for (let c = 0; c < width; c += 1) {
      if (rows.some((row) => !cellIsBlank(row[c]))) keepColumns.push(c);
    }
    rows = rows.map((row) => keepColumns.map((c) => row[c] ?? createCell(null)));
  }
  return rows;
}

function cleanToNewSheetTarget(workbook: Workbook, args: CleanToNewSheetArgs): CellRange[] {
  const source = getSheet(workbook, args.sheet);
  if (!source) return [];
  const cleaned = buildCleanedRows(source.rows, args);
  const name = uniqueSheetName(workbook, args.targetSheet);
  const cols = Math.max(1, ...cleaned.map((row) => row.length));
  return [
    {
      sheet: name,
      startColumn: 'A',
      endColumn: indexToColumn(Math.max(0, cols - 1)),
      startRow: 1,
      endRow: Math.max(1, cleaned.length),
    },
  ];
}

function validateCleanToNewSheet(workbook: Workbook, args: CleanToNewSheetArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  if (args.targetSheet.trim().toLowerCase() === args.sheet.trim().toLowerCase()) {
    errors.push({
      code: 'same_sheet',
      message: 'Target sheet must be a different sheet than the source.',
    });
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyCleanToNewSheet(workbook: Workbook, args: CleanToNewSheetArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const source = getSheet(after, args.sheet);

  if (source) {
    const cleaned = buildCleanedRows(source.rows, args);
    const name = uniqueSheetName(after, args.targetSheet);
    after.sheets.push({ name, rows: cleaned });
  }

  const ranges = cleanToNewSheetTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const cleanToNewSheetOperation: Operation<CleanToNewSheetArgs> = {
  name: 'clean_to_new_sheet',
  schema: cleanToNewSheetArgsSchema,
  targetRanges: cleanToNewSheetTarget,
  validate: validateCleanToNewSheet,
  preview(workbook, args) {
    const validation = validateCleanToNewSheet(workbook, args);
    const ranges = cleanToNewSheetTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyCleanToNewSheet(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyCleanToNewSheet,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: cleanToNewSheetTarget(before, args),
    });
  },
};

// ============================================================================
// 7. EDIT CELLS (edit_cells)
// ============================================================================

/** Excel's own sheet limits. A grid edit must never be able to build a sheet no reader can open. */
const MAX_ROW_NUMBER = 1_048_576;
const MAX_COLUMN_INDEX = 16_383; // XFD
/** A bound on one call, so a runaway paste fails loudly instead of allocating a million cells. */
const MAX_EDITS_PER_OPERATION = 20_000;

export const editCellsArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  edits: z
    .array(
      z.object({
        /** 1-based row. */
        row: z.number().int().positive(),
        column: rangeColumn,
        /** Omitted together with `formula` clears the cell. */
        value: cellValueSchema.optional(),
        /** Stored with its leading `=`, matching how every other operation writes formulas. */
        formula: z.string().optional(),
        numberFormat: z.string().optional(),
      }),
    )
    .min(1),
});
export type EditCellsArgs = z.infer<typeof editCellsArgsSchema>;
export type CellEdit = EditCellsArgs['edits'][number];

/**
 * The declared target is the whole grown sheet, because this operation may pad a short row or add
 * rows to reach an address. Those padded cells are real writes from the invariant checker's point
 * of view, and a range covering only the requested addresses would - correctly - be rejected for
 * touching cells outside itself.
 */
function editCellsTarget(workbook: Workbook, args: EditCellsArgs): CellRange[] {
  const rows = getSheet(workbook, args.sheet)?.rows ?? [];
  let endRow = Math.max(1, rows.length);
  let endColumn = Math.max(0, maxColumnCount(rows) - 1);
  for (const edit of args.edits) {
    endRow = Math.max(endRow, Math.min(edit.row, MAX_ROW_NUMBER));
    endColumn = Math.max(endColumn, columnToIndex(edit.column) ?? 0);
  }
  return [
    {
      sheet: args.sheet,
      startRow: 1,
      endRow,
      startColumn: 'A',
      endColumn: indexToColumn(Math.min(MAX_COLUMN_INDEX, endColumn)),
    },
  ];
}

function validateEditCells(workbook: Workbook, args: EditCellsArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  if (getSheet(workbook, args.sheet)) {
    if (args.edits.length > MAX_EDITS_PER_OPERATION) {
      errors.push(
        issue(
          'too-many-edits',
          `A single edit_cells call accepts at most ${MAX_EDITS_PER_OPERATION} cells; got ${args.edits.length}.`,
        ),
      );
    }
    for (const edit of args.edits) {
      const columnIndex = columnToIndex(edit.column);
      if (columnIndex === undefined) {
        errors.push(
          issue('invalid-column', `Column "${edit.column}" is not a valid column reference.`),
        );
        continue;
      }
      if (columnIndex > MAX_COLUMN_INDEX) {
        errors.push(
          issue(
            'cell-out-of-range',
            `Column "${indexToColumn(columnIndex)}" is past the last addressable column (XFD).`,
          ),
        );
      }
      if (edit.row > MAX_ROW_NUMBER) {
        errors.push(
          issue(
            'cell-out-of-range',
            `Row ${edit.row} is past the last addressable row (${MAX_ROW_NUMBER}).`,
          ),
        );
      }
    }
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyEditCells(workbook: Workbook, args: EditCellsArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);

  if (sheet) {
    for (const edit of args.edits) {
      const columnIndex = columnToIndex(edit.column);
      if (
        columnIndex === undefined ||
        columnIndex > MAX_COLUMN_INDEX ||
        edit.row > MAX_ROW_NUMBER
      ) {
        continue;
      }
      // A paste into the empty space right of a short row, or below the last row, is ordinary
      // spreadsheet behaviour, so the sheet grows to fit rather than refusing the edit.
      while (sheet.rows.length < edit.row) sheet.rows.push([]);
      const row = sheet.rows[edit.row - 1];
      if (!row) continue;
      while (row.length < columnIndex) row.push(createCell(null));
      row[columnIndex] = cloneWithValue(row[columnIndex] ?? createCell(null), edit.value ?? null, {
        ...(edit.formula !== undefined ? { formula: edit.formula } : {}),
        ...(edit.numberFormat !== undefined ? { numberFormat: edit.numberFormat } : {}),
      });
    }
  }

  const ranges = editCellsTarget(workbook, args);
  // `requiresConfirmation` stays false on purpose: this is the operation behind a keystroke in a
  // cell. A confirmation gate on every character typed would make the grid unusable, and the user
  // asking for the edit is the confirmation. Bulk and destructive work still goes through the
  // range operations, which do demand it.
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const editCellsOperation: Operation<EditCellsArgs> = {
  name: 'edit_cells',
  schema: editCellsArgsSchema,
  targetRanges: editCellsTarget,
  validate: validateEditCells,
  preview(workbook, args) {
    const validation = validateEditCells(workbook, args);
    const ranges = editCellsTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyEditCells(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyEditCells,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: editCellsTarget(before, args),
      // No `rowCountUnchanged`: typing into the first empty row below the data adds one, which is
      // the whole point of a grid. Formulas may also be replaced, exactly as typing over one does.
      allowFormulaChanges: true,
    });
  },
};

// ============================================================================
// 8. FILTER TO NEW SHEET (filter_to_new_sheet)
// ============================================================================

export const filterToNewSheetArgsSchema = z
  .object({
    sheet: z.string().trim().min(1),
    targetSheet: z.string().trim().min(1),
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
    headerRow: headerRowSchema,
    dropFromSource: z.boolean().default(false),
  })
  .refine(
    (args) =>
      args.operator === 'is_blank' || args.operator === 'is_not_blank' || args.value !== undefined,
    { message: 'A filter value is required for this operator.', path: ['value'] },
  );
export type FilterToNewSheetArgs = z.infer<typeof filterToNewSheetArgsSchema>;

function filterToNewSheetTarget(workbook: Workbook, args: FilterToNewSheetArgs): CellRange[] {
  const source = getSheet(workbook, args.sheet);
  if (!source) return [];
  const colIndex = columnToIndex(args.column) ?? 0;
  const start = args.headerRow;
  const dataRows = source.rows.slice(start);
  const matchingCount = dataRows.filter((r) => matchesFilter(r[colIndex], args)).length;
  const totalCols = Math.max(1, maxColumnCount(source.rows));
  const newSheetName = uniqueSheetName(workbook, args.targetSheet);
  const ranges: CellRange[] = [
    {
      sheet: newSheetName,
      startColumn: 'A',
      endColumn: indexToColumn(totalCols - 1),
      startRow: 1,
      endRow: Math.max(1, start + matchingCount),
    },
  ];
  if (args.dropFromSource) {
    ranges.push(allDataRange(workbook, args.sheet, args.headerRow));
  }
  return ranges;
}

function validateFilterToNewSheet(
  workbook: Workbook,
  args: FilterToNewSheetArgs,
): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateColumn(workbook, args.sheet, args.column),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  if (args.targetSheet.trim().toLowerCase() === args.sheet.trim().toLowerCase()) {
    errors.push({
      code: 'same_sheet',
      message: 'Target sheet must be a different sheet than the source.',
    });
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyFilterToNewSheet(
  workbook: Workbook,
  args: FilterToNewSheetArgs,
): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const source = getSheet(after, args.sheet);
  if (source) {
    const colIndex = columnToIndex(args.column) ?? 0;
    const start = args.headerRow;
    const headerRows = source.rows.slice(0, start).map((r) => r.map((c) => cloneCell(c)));
    const dataRows = source.rows.slice(start);
    const matchingRows = dataRows
      .filter((r) => matchesFilter(r[colIndex], args))
      .map((r) => r.map((c) => cloneCell(c)));

    if (args.dropFromSource) {
      const remainingRows = dataRows.filter((r) => !matchesFilter(r[colIndex], args));
      source.rows.splice(start, dataRows.length, ...remainingRows);
    }

    const newSheetName = uniqueSheetName(after, args.targetSheet);
    after.sheets.push({
      name: newSheetName,
      rows: [...headerRows, ...matchingRows],
    });
  }
  const ranges = filterToNewSheetTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const filterToNewSheetOperation: Operation<FilterToNewSheetArgs> = {
  name: 'filter_to_new_sheet',
  schema: filterToNewSheetArgsSchema,
  targetRanges: filterToNewSheetTarget,
  validate: validateFilterToNewSheet,
  preview(workbook, args) {
    const validation = validateFilterToNewSheet(workbook, args);
    const ranges = filterToNewSheetTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyFilterToNewSheet(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyFilterToNewSheet,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: filterToNewSheetTarget(before, args),
    });
  },
};

// ============================================================================
// 9. CREATE SHEET (create_sheet)
// ============================================================================

export const createSheetArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  headers: z.array(z.string()).optional(),
});
export type CreateSheetArgs = z.infer<typeof createSheetArgsSchema>;

function createSheetTarget(workbook: Workbook, args: CreateSheetArgs): CellRange[] {
  const name = uniqueSheetName(workbook, args.sheetName);
  const colCount = Math.max(1, args.headers?.length ?? 1);
  return [
    {
      sheet: name,
      startColumn: 'A',
      endColumn: indexToColumn(colCount - 1),
      startRow: 1,
      endRow: 1,
    },
  ];
}

function validateCreateSheet(_workbook: Workbook, _args: CreateSheetArgs): ValidationResult {
  return validResult();
}

function applyCreateSheet(workbook: Workbook, args: CreateSheetArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const name = uniqueSheetName(after, args.sheetName);
  const headerRow: Cell[] = (args.headers ?? []).map((h) => createCell(h));
  after.sheets.push({
    name,
    rows: headerRow.length > 0 ? [headerRow] : [],
  });
  const ranges = createSheetTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const createSheetOperation: Operation<CreateSheetArgs> = {
  name: 'create_sheet',
  schema: createSheetArgsSchema,
  targetRanges: createSheetTarget,
  validate: validateCreateSheet,
  preview(workbook, args) {
    const ranges = createSheetTarget(workbook, args);
    const result = applyCreateSheet(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyCreateSheet,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: createSheetTarget(before, args),
    });
  },
};

// ============================================================================
// 10. DUPLICATE SHEET (duplicate_sheet)
// ============================================================================

export const duplicateSheetArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  targetSheet: z.string().trim().min(1),
});
export type DuplicateSheetArgs = z.infer<typeof duplicateSheetArgsSchema>;

function duplicateSheetTarget(workbook: Workbook, args: DuplicateSheetArgs): CellRange[] {
  const source = getSheet(workbook, args.sheet);
  if (!source) return [];
  const name = uniqueSheetName(workbook, args.targetSheet);
  const cols = Math.max(1, maxColumnCount(source.rows));
  return [
    {
      sheet: name,
      startColumn: 'A',
      endColumn: indexToColumn(cols - 1),
      startRow: 1,
      endRow: Math.max(1, source.rows.length),
    },
  ];
}

function validateDuplicateSheet(workbook: Workbook, args: DuplicateSheetArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyDuplicateSheet(workbook: Workbook, args: DuplicateSheetArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const source = getSheet(after, args.sheet);
  if (source) {
    const name = uniqueSheetName(after, args.targetSheet);
    const clonedRows = source.rows.map((row) => row.map((c) => cloneCell(c)));
    after.sheets.push({ name, rows: clonedRows });
  }
  const ranges = duplicateSheetTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const duplicateSheetOperation: Operation<DuplicateSheetArgs> = {
  name: 'duplicate_sheet',
  schema: duplicateSheetArgsSchema,
  targetRanges: duplicateSheetTarget,
  validate: validateDuplicateSheet,
  preview(workbook, args) {
    const validation = validateDuplicateSheet(workbook, args);
    const ranges = duplicateSheetTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyDuplicateSheet(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyDuplicateSheet,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: duplicateSheetTarget(before, args),
    });
  },
};

// ============================================================================
// 11. DELETE SHEET (delete_sheet)
// ============================================================================

export const deleteSheetArgsSchema = z.object({
  sheet: z.string().trim().min(1),
});
export type DeleteSheetArgs = z.infer<typeof deleteSheetArgsSchema>;

function deleteSheetTarget(workbook: Workbook, args: DeleteSheetArgs): CellRange[] {
  return [fullSheetRange(workbook, args.sheet)];
}

function validateDeleteSheet(workbook: Workbook, args: DeleteSheetArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  if (workbook.sheets.length <= 1) {
    errors.push(issue('cannot-delete-last-sheet', 'Cannot delete the only sheet in the workbook.'));
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyDeleteSheet(workbook: Workbook, args: DeleteSheetArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheetIdx = after.sheets.findIndex(
    (s) => s.name.toLowerCase() === args.sheet.toLowerCase().trim(),
  );
  if (sheetIdx >= 0) {
    after.sheets.splice(sheetIdx, 1);
  }
  const ranges = deleteSheetTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const deleteSheetOperation: Operation<DeleteSheetArgs> = {
  name: 'delete_sheet',
  schema: deleteSheetArgsSchema,
  targetRanges: deleteSheetTarget,
  validate: validateDeleteSheet,
  preview(workbook, args) {
    const validation = validateDeleteSheet(workbook, args);
    const ranges = deleteSheetTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyDeleteSheet(workbook, args);
    return {
      ...previewForTransition(workbook, result.workbook, ranges),
      requiresConfirmation: true,
    };
  },
  apply: applyDeleteSheet,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: deleteSheetTarget(before, args),
    });
  },
};

// ============================================================================
// 12. ADD SUMMARY ROW (add_summary_row)
// ============================================================================

export const addSummaryRowArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  columns: z.array(rangeColumn).optional(),
  aggregation: z.enum(['sum', 'average', 'count', 'min', 'max']).default('sum'),
  label: z.string().default('Total'),
  labelColumn: rangeColumn.default('A'),
  headerRow: headerRowSchema,
});
export type AddSummaryRowArgs = z.infer<typeof addSummaryRowArgsSchema>;

function addSummaryRowTarget(workbook: Workbook, args: AddSummaryRowArgs): CellRange[] {
  const sheet = getSheet(workbook, args.sheet);
  if (!sheet) return [];
  const rowIdx = sheet.rows.length + 1;
  const cols = Math.max(1, maxColumnCount(sheet.rows));
  return [
    {
      sheet: args.sheet,
      startColumn: 'A',
      endColumn: indexToColumn(cols - 1),
      startRow: rowIdx,
      endRow: rowIdx,
    },
  ];
}

function validateAddSummaryRow(workbook: Workbook, args: AddSummaryRowArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...headerRowError(workbook, args.sheet, args.headerRow),
  ];
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyAddSummaryRow(workbook: Workbook, args: AddSummaryRowArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  if (sheet) {
    const totalCols = Math.max(1, maxColumnCount(sheet.rows));
    const newRow: Cell[] = Array.from({ length: totalCols }, () => createCell(null));
    const labelColIdx = columnToIndex(args.labelColumn) ?? 0;
    if (labelColIdx < totalCols) {
      newRow[labelColIdx] = createCell(args.label);
    }
    const targetCols =
      args.columns?.map((c) => columnToIndex(c) ?? -1).filter((i) => i >= 0) ??
      Array.from({ length: totalCols }, (_, i) => i);

    const start = args.headerRow;
    const dataRows = sheet.rows.slice(start);

    for (const cIdx of targetCols) {
      if (cIdx === labelColIdx) continue;
      const values: number[] = [];
      for (const row of dataRows) {
        const val = row[cIdx]?.value;
        const num = typeof val === 'number' ? val : Number(String(val ?? '').replace(/,/g, ''));
        if (!isNaN(num) && String(val ?? '').trim() !== '') {
          values.push(num);
        }
      }
      if (values.length > 0) {
        let resultVal = 0;
        switch (args.aggregation) {
          case 'sum':
            resultVal = values.reduce((a, b) => a + b, 0);
            break;
          case 'average':
            resultVal = values.reduce((a, b) => a + b, 0) / values.length;
            break;
          case 'count':
            resultVal = values.length;
            break;
          case 'min':
            resultVal = Math.min(...values);
            break;
          case 'max':
            resultVal = Math.max(...values);
            break;
        }
        newRow[cIdx] = createCell(Math.round(resultVal * 100) / 100);
      }
    }
    sheet.rows.push(newRow);
  }
  const ranges = addSummaryRowTarget(workbook, args);
  return transitionResult(before, after, operationReport(before, after, ranges));
}

export const addSummaryRowOperation: Operation<AddSummaryRowArgs> = {
  name: 'add_summary_row',
  schema: addSummaryRowArgsSchema,
  targetRanges: addSummaryRowTarget,
  validate: validateAddSummaryRow,
  preview(workbook, args) {
    const validation = validateAddSummaryRow(workbook, args);
    const ranges = addSummaryRowTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyAddSummaryRow(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyAddSummaryRow,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: addSummaryRowTarget(before, args),
    });
  },
};

export const advancedOperations: Operation<unknown>[] = [
  fillBlanksOperation as Operation<unknown>,
  addComputedColumnOperation as Operation<unknown>,
  splitColumnOperation as Operation<unknown>,
  mergeColumnsOperation as Operation<unknown>,
  lookupMergeOperation as Operation<unknown>,
  cleanToNewSheetOperation as Operation<unknown>,
  editCellsOperation as Operation<unknown>,
  filterToNewSheetOperation as Operation<unknown>,
  createSheetOperation as Operation<unknown>,
  duplicateSheetOperation as Operation<unknown>,
  deleteSheetOperation as Operation<unknown>,
  addSummaryRowOperation as Operation<unknown>,
];
