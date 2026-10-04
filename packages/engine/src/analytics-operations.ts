import { z } from 'zod';

import { matchesCriteria, toNumericOrNull } from './formula/functions.js';
import { runInvariants } from './invariants.js';
import {
  cellValueSchema,
  issue,
  maxRow,
  previewForTransition,
  transitionResult,
  uniqueSheetName,
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
  Sheet,
  ValidationIssue,
  ValidationResult,
  Workbook,
} from './types.js';
import {
  cellValueEquals,
  cloneWorkbook,
  columnToIndex,
  createCell,
  getSheet,
  indexToColumn,
  maxColumnCount,
} from './workbook.js';

// ============================================================================
// Shared plumbing
// ============================================================================

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

function headerRowIssues(workbook: Workbook, sheet: string, headerRow: number): ValidationIssue[] {
  const sheetData = getSheet(workbook, sheet);
  return !sheetData || headerRow <= sheetData.rows.length
    ? []
    : [issue('invalid-header-row', `Header row ${headerRow} is outside sheet "${sheet}".`)];
}

function validationFailed(errors: ValidationIssue[], warnings: ValidationIssue[] = []) {
  return errors.length === 0 ? validResult(warnings) : { valid: false, errors, warnings };
}

function valueText(value: CellValue): string {
  if (value === null) return '';
  return value instanceof Date ? value.toISOString() : String(value);
}

function isBlankValue(value: CellValue | undefined): boolean {
  return (
    value === null || value === undefined || (typeof value === 'string' && value.trim() === '')
  );
}

function rowIsBlank(row: Cell[] | undefined): boolean {
  return !row || row.every((cell) => isBlankValue(cell.value));
}

function cloneValue(value: CellValue): CellValue {
  return value instanceof Date ? new Date(value.getTime()) : value;
}

/**
 * Inserts a cell at an absolute column index in a row that may be narrower than that index.
 *
 * `Array.prototype.splice` clamps an out-of-range index to the end of the array, so on a ragged
 * sheet the naive call would drop column G of a two-cell row into column C. Padding first keeps
 * every value in the column the caller asked for; `paddingFloor` declares those padding cells in
 * the target range so the write stays inside what was promised.
 */
function insertCell(row: Cell[], index: number, cell: Cell): void {
  while (row.length < index) row.push(createCell(null));
  row.splice(index, 0, cell);
}

/**
 * The first column index that has to be inside a declared target range for an insertion at
 * `insertion`: every row narrower than the insertion point is padded with blank cells before it.
 */
function paddingFloor(sheet: Sheet | undefined, insertion: number): number {
  if (!sheet) return insertion;
  let narrowest = insertion;
  for (const row of sheet.rows) {
    if (row.length < narrowest) narrowest = row.length;
  }
  return narrowest;
}

function comparable(value: CellValue): string | number | boolean | null {
  return value instanceof Date ? value.getTime() : value;
}

/** Type-ranked, locale-aware ordering. Blanks sort last, as they do in the engine's own sorts. */
function compareValues(left: CellValue, right: CellValue): number {
  const a = comparable(left);
  const b = comparable(right);
  if (a === null && b === null) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
  }
  const rank = (value: unknown): number =>
    typeof value === 'number' ? 0 : typeof value === 'boolean' ? 1 : 2;
  if (typeof a !== typeof b) return rank(a) - rank(b);
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;
  return 0;
}

function truthiness(value: CellValue): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const text = value.trim().toUpperCase();
    if (text === 'TRUE') return true;
    if (text === 'FALSE') return false;
    return text.length > 0;
  }
  return false;
}

/**
 * Equality that forgives the differences a spreadsheet actually produces: case, padding, and
 * "1,234" against 1234.
 */
function looselyEqual(left: CellValue, right: CellValue): boolean {
  if (cellValueEquals(left, right)) return true;
  if (isBlankValue(left) || isBlankValue(right)) {
    return isBlankValue(left) && isBlankValue(right);
  }
  const a = toNumericOrNull(left);
  const b = toNumericOrNull(right);
  if (a !== null && b !== null) return a === b;
  if (typeof left === 'boolean' || typeof right === 'boolean') {
    return truthiness(left) === truthiness(right);
  }
  return valueText(left).trim().toLowerCase() === valueText(right).trim().toLowerCase();
}

/** Round away binary-float dust so an extrapolated series reads back as the numbers a user typed. */
function tidy(value: number): number {
  return Number.isFinite(value) ? Math.round(value * 1e10) / 1e10 : 0;
}

/** Join key normalization: trimmed and case-insensitive, which is how a lookup key behaves. */
function keyFor(value: CellValue): string {
  return valueText(value).trim().toLowerCase();
}

// ============================================================================
// Aggregation primitives
// ============================================================================

export type Aggregation =
  'sum' | 'average' | 'count' | 'count_distinct' | 'min' | 'max' | 'median' | 'stdev';

const aggregationSchema = z.enum([
  'sum',
  'average',
  'count',
  'count_distinct',
  'min',
  'max',
  'median',
  'stdev',
]);

interface AggregateOutcome {
  value: number;
  /** Cells that carried text or a boolean where a magnitude was required. */
  nonNumeric: number;
  /** Cells whose value was empty. */
  blank: number;
  warnings: ValidationIssue[];
}

function medianOf(values: number[]): number {
  const sorted = values.slice().sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) return sorted[middle] as number;
  return ((sorted[middle - 1] as number) + (sorted[middle] as number)) / 2;
}

/** Sample standard deviation, matching Excel's STDEV. A single observation has no spread. */
function sampleStdev(values: number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((total, value) => total + value, 0) / values.length;
  const variance =
    values.reduce((total, value) => total + (value - mean) * (value - mean), 0) /
    (values.length - 1);
  return Math.sqrt(variance);
}

/**
 * Reduces one group of cells to a single number.
 *
 * `count` and `count_distinct` measure occupancy (COUNTA semantics) because a group whose value
 * column is text still has a row count. Every other aggregation measures magnitude, and counts
 * what it had to skip, so "10 total" is visibly a total of two of five rows rather than a total
 * that quietly treated the other three as zero.
 */
export function aggregateCells(cells: Cell[], aggregation: Aggregation): AggregateOutcome {
  const warnings: ValidationIssue[] = [];
  const present = cells.filter((cell) => !isBlankValue(cell.value));
  const blank = cells.length - present.length;

  if (aggregation === 'count') {
    return { value: present.length, nonNumeric: 0, blank, warnings };
  }

  if (aggregation === 'count_distinct') {
    return {
      value: new Set(present.map((cell) => keyFor(cell.value))).size,
      nonNumeric: 0,
      blank,
      warnings,
    };
  }

  const numbers: number[] = [];
  let nonNumeric = 0;
  for (const cell of present) {
    const numeric = toNumericOrNull(cell.value);
    if (numeric === null) nonNumeric += 1;
    else numbers.push(numeric);
  }

  if (numbers.length === 0 && present.length > 0) {
    warnings.push(
      issue(
        'no-numeric-values',
        `${aggregation} found no numeric values in ${present.length} non-empty cell(s); reported 0.`,
      ),
    );
  }

  switch (aggregation) {
    case 'sum':
      return {
        value: tidy(numbers.reduce((total, value) => total + value, 0)),
        nonNumeric,
        blank,
        warnings,
      };
    case 'average':
      return {
        value:
          numbers.length === 0
            ? 0
            : tidy(numbers.reduce((total, value) => total + value, 0) / numbers.length),
        nonNumeric,
        blank,
        warnings,
      };
    case 'min':
      return {
        value: numbers.length === 0 ? 0 : Math.min(...numbers),
        nonNumeric,
        blank,
        warnings,
      };
    case 'max':
      return {
        value: numbers.length === 0 ? 0 : Math.max(...numbers),
        nonNumeric,
        blank,
        warnings,
      };
    case 'median':
      return {
        value: numbers.length === 0 ? 0 : tidy(medianOf(numbers)),
        nonNumeric,
        blank,
        warnings,
      };
    default:
      if (numbers.length === 1) {
        warnings.push(
          issue(
            'insufficient-values',
            'stdev needs at least two numeric values; reported 0 for a single value.',
          ),
        );
      }
      return {
        value: numbers.length < 2 ? 0 : tidy(sampleStdev(numbers)),
        nonNumeric,
        blank,
        warnings,
      };
  }
}

// ============================================================================
// 1. AGGREGATE COLUMN (aggregate_column)
// ============================================================================

/** Operator tokens that may be split across `criteria` and `criteriaValue`. */
const CRITERIA_OPERATORS = new Set(['=', '==', '!=', '<>', '<', '<=', '>', '>=']);

interface CriteriaArgs {
  criteria?: string;
  criteriaValue?: string;
}

/**
 * One criteria string out of the two-part form.
 *
 * `criteriaValue` alone is the value to compare against. `criteria` alone is a complete criteria
 * expression, so the whole COUNTIF vocabulary ("*North*", "<>Closed", ">=2024-01-01") is available.
 * Given both, `criteria` must be a bare operator and the pair is joined, which is how a caller
 * naturally writes { criteria: '>=', criteriaValue: '100' }.
 */
function criteriaExpression(
  args: CriteriaArgs,
): { ok: true; expression?: string } | { ok: false; message: string } {
  if (args.criteria === undefined) {
    return args.criteriaValue === undefined
      ? { ok: true }
      : { ok: true, expression: args.criteriaValue };
  }
  if (args.criteriaValue === undefined) return { ok: true, expression: args.criteria };
  if (!CRITERIA_OPERATORS.has(args.criteria)) {
    return {
      ok: false,
      message: `criteria "${args.criteria}" is not a comparison operator. Pass a complete expression such as ">=100", or pass a bare operator together with criteriaValue.`,
    };
  }
  return { ok: true, expression: `${args.criteria}${args.criteriaValue}` };
}

export const aggregateColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  aggregation: aggregationSchema,
  criteria: z.string().trim().min(1).optional(),
  criteriaValue: z.string().trim().min(1).optional(),
  criteriaColumn: rangeColumn.optional(),
  headerRow: headerRowSchema,
});
export type AggregateColumnArgs = z.infer<typeof aggregateColumnArgsSchema>;

function aggregateColumnTarget(_workbook: Workbook, _args: AggregateColumnArgs): CellRange[] {
  // The operation reads a column and reports a number. It writes nothing, so it claims nothing.
  return [];
}

function validateAggregateColumn(workbook: Workbook, args: AggregateColumnArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateColumn(workbook, args.sheet, args.column),
    ...headerRowIssues(workbook, args.sheet, args.headerRow),
  ];
  if (args.criteriaColumn !== undefined) {
    errors.push(...validateColumn(workbook, args.sheet, args.criteriaColumn));
  }
  const composed = criteriaExpression(args);
  if (!composed.ok) errors.push(issue('invalid-criteria', composed.message));
  return validationFailed(errors);
}

/**
 * The cells of `column` that survive the criteria. Fully blank rows are skipped so a
 * formatted-but-empty tail under the data cannot inflate a count.
 */
function selectAggregateCells(workbook: Workbook, args: AggregateColumnArgs): Cell[] {
  const sheet = getSheet(workbook, args.sheet);
  if (!sheet) return [];
  const valueColumn = columnToIndex(args.column) ?? 0;
  const filterColumn = args.criteriaColumn
    ? (columnToIndex(args.criteriaColumn) ?? valueColumn)
    : valueColumn;
  const composed = criteriaExpression(args);
  const expression = composed.ok ? composed.expression : undefined;

  const selected: Cell[] = [];
  for (let rowIndex = args.headerRow; rowIndex < sheet.rows.length; rowIndex += 1) {
    const row = sheet.rows[rowIndex];
    if (rowIsBlank(row)) continue;
    if (
      expression !== undefined &&
      !matchesCriteria(row?.[filterColumn]?.value ?? null, expression)
    ) {
      continue;
    }
    selected.push(row?.[valueColumn] ?? createCell(null));
  }
  return selected;
}

function applyAggregateColumn(workbook: Workbook, args: AggregateColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const outcome = aggregateCells(selectAggregateCells(workbook, args), args.aggregation);
  return transitionResult(
    before,
    after,
    operationReport(before, after, aggregateColumnTarget(workbook, args), outcome.warnings, {
      aggregate: outcome.value,
      skippedCells: outcome.nonNumeric,
    }),
  );
}

export const aggregateColumnOperation: Operation<AggregateColumnArgs> = {
  name: 'aggregate_column',
  schema: aggregateColumnArgsSchema,
  targetRanges: aggregateColumnTarget,
  validate: validateAggregateColumn,
  preview(workbook, args) {
    const validation = validateAggregateColumn(workbook, args);
    const ranges = aggregateColumnTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyAggregateColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges, result.report.warnings);
  },
  apply: applyAggregateColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, { targetRanges: aggregateColumnTarget(before, args) });
  },
};

// ============================================================================
// 2. GROUP AND SUMMARIZE (group_and_summarize)
// ============================================================================

const summarizableSchema = aggregationSchema.exclude(['count_distinct', 'median', 'stdev']);

export const groupAndSummarizeArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  groupBy: z
    .array(rangeColumn)
    .min(1)
    .refine(
      (columns) => new Set(columns.map((column) => column.toUpperCase())).size === columns.length,
      { message: 'Duplicate group-by columns are not allowed.' },
    ),
  valueColumn: rangeColumn,
  aggregation: summarizableSchema,
  targetSheet: z.string().trim().min(1),
  headerRow: headerRowSchema,
});
export type GroupAndSummarizeArgs = z.infer<typeof groupAndSummarizeArgsSchema>;

interface Summary {
  rows: Cell[][];
  unmatchedGroups: number;
  nonNumeric: number;
  warnings: ValidationIssue[];
}

function headerText(workbook: Workbook, args: GroupAndSummarizeArgs, column: string): string {
  const sheet = getSheet(workbook, args.sheet);
  const index = columnToIndex(column) ?? 0;
  const text = valueText(sheet?.rows[args.headerRow - 1]?.[index]?.value ?? null).trim();
  return text === '' ? column.toUpperCase() : text;
}

/**
 * The pivot itself, shared by `apply` and `targetRanges` so the declared write area and the cells
 * actually written can never disagree.
 */
function buildSummary(workbook: Workbook, args: GroupAndSummarizeArgs): Summary {
  const sheet = getSheet(workbook, args.sheet);
  const groupColumns = args.groupBy.map((column) => columnToIndex(column) ?? 0);
  const valueColumn = columnToIndex(args.valueColumn) ?? 0;
  const warnings: ValidationIssue[] = [];

  const groups = new Map<string, { key: CellValue[]; cells: Cell[] }>();
  for (let rowIndex = args.headerRow; rowIndex < (sheet?.rows.length ?? 0); rowIndex += 1) {
    const row = sheet?.rows[rowIndex];
    if (rowIsBlank(row)) continue;
    const key = groupColumns.map((column) => row?.[column]?.value ?? null);
    const id = JSON.stringify(key.map((value) => keyFor(value)));
    const cell = row?.[valueColumn] ?? createCell(null);
    const group = groups.get(id);
    if (group) group.cells.push(cell);
    else groups.set(id, { key, cells: [cell] });
  }

  const ordered = [...groups.values()].sort((left, right) => {
    for (let index = 0; index < left.key.length; index += 1) {
      const comparison = compareValues(left.key[index] ?? null, right.key[index] ?? null);
      if (comparison !== 0) return comparison;
    }
    return 0;
  });

  const rows: Cell[][] = [
    [
      ...args.groupBy.map((column) => createCell(headerText(workbook, args, column))),
      createCell(`${args.aggregation} of ${headerText(workbook, args, args.valueColumn)}`),
    ],
  ];

  let nonNumeric = 0;
  let unmatchedGroups = 0;
  for (const group of ordered) {
    const outcome = aggregateCells(group.cells, args.aggregation);
    nonNumeric += outcome.nonNumeric;
    const measurable = outcome.blank + outcome.nonNumeric < group.cells.length;
    // `count` always has an answer; a mean or a minimum over nothing does not. Blank beats 0 here
    // because 0 is a real result that would be indistinguishable from "the numbers really were 0".
    if (!measurable && args.aggregation !== 'count') unmatchedGroups += 1;
    const value = measurable || args.aggregation === 'count' ? outcome.value : null;
    rows.push([...group.key.map((key) => createCell(key)), createCell(value)]);
  }

  if (unmatchedGroups > 0) {
    warnings.push(
      issue(
        'groups-without-values',
        `${unmatchedGroups} group(s) had no numeric value in "${args.valueColumn}"; they are reported as blank.`,
      ),
    );
  }

  return { rows, unmatchedGroups, nonNumeric, warnings };
}

function groupAndSummarizeTarget(workbook: Workbook, args: GroupAndSummarizeArgs): CellRange[] {
  if (!getSheet(workbook, args.sheet)) return [];
  const { rows } = buildSummary(workbook, args);
  const width = Math.max(1, ...rows.map((row) => row.length));
  return [
    {
      sheet: uniqueSheetName(workbook, args.targetSheet),
      startRow: 1,
      endRow: Math.max(1, rows.length),
      startColumn: 'A',
      endColumn: indexToColumn(width - 1),
    },
  ];
}

function validateGroupAndSummarize(
  workbook: Workbook,
  args: GroupAndSummarizeArgs,
): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...headerRowIssues(workbook, args.sheet, args.headerRow),
  ];
  for (const column of [...args.groupBy, args.valueColumn]) {
    errors.push(...validateColumn(workbook, args.sheet, column));
  }
  if (args.targetSheet.trim().toLowerCase() === args.sheet.trim().toLowerCase()) {
    errors.push(
      issue('same-sheet', 'targetSheet must be a different sheet than the sheet being summarized.'),
    );
  }
  return validationFailed(errors);
}

function applyGroupAndSummarize(workbook: Workbook, args: GroupAndSummarizeArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const summary = buildSummary(workbook, args);
  after.sheets.push({ name: uniqueSheetName(after, args.targetSheet), rows: summary.rows });
  return transitionResult(
    before,
    after,
    operationReport(before, after, groupAndSummarizeTarget(workbook, args), summary.warnings, {
      addedRows: summary.rows.length,
      addedColumns: args.groupBy.length + 1,
      unmatchedRows: summary.unmatchedGroups,
      skippedCells: summary.nonNumeric,
    }),
  );
}

export const groupAndSummarizeOperation: Operation<GroupAndSummarizeArgs> = {
  name: 'group_and_summarize',
  schema: groupAndSummarizeArgsSchema,
  targetRanges: groupAndSummarizeTarget,
  validate: validateGroupAndSummarize,
  preview(workbook, args) {
    const validation = validateGroupAndSummarize(workbook, args);
    const ranges = groupAndSummarizeTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyGroupAndSummarize(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges, result.report.warnings);
  },
  apply: applyGroupAndSummarize,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: groupAndSummarizeTarget(before, args),
    });
  },
};

// ============================================================================
// 3. JOIN SHEETS (join_sheets)
// ============================================================================

const joinHeaderNameSchema = z.union([
  z.string().trim().min(1),
  z.array(z.string().trim().min(1)).min(1),
]);

export const joinSheetsArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  keyColumn: rangeColumn,
  lookupSheet: z.string().trim().min(1),
  lookupKeyColumn: rangeColumn,
  lookupValueColumn: z.union([
    rangeColumn,
    z
      .array(rangeColumn)
      .min(1)
      .refine(
        (columns) => new Set(columns.map((column) => column.toUpperCase())).size === columns.length,
        { message: 'Duplicate lookup value columns are not allowed.' },
      ),
  ]),
  headerName: joinHeaderNameSchema,
  joinType: z.enum(['inner', 'left']).default('left'),
  headerRow: headerRowSchema,
});
export type JoinSheetsArgs = z.infer<typeof joinSheetsArgsSchema>;

function valueColumnsOf(args: JoinSheetsArgs): string[] {
  return Array.isArray(args.lookupValueColumn) ? args.lookupValueColumn : [args.lookupValueColumn];
}

function headerNamesOf(args: JoinSheetsArgs): string[] {
  return Array.isArray(args.headerName) ? args.headerName : [args.headerName];
}

interface LookupIndex {
  rows: Map<string, CellValue[]>;
  duplicates: string[];
}

/**
 * The lookup side of the join.
 *
 * A duplicated key keeps the FIRST match, which is what VLOOKUP and XLOOKUP already do inside this
 * engine. `lookup_merge` builds its map with `Map.set`, so it silently last-wins and can report a
 * different value than a formula written against the same lookup sheet. Joining first-wins keeps
 * the operation and the formula engine telling the same story, and the duplicate is surfaced as a
 * warning instead of being resolved quietly, because a repeated lookup key is a data defect and
 * only the caller can say which row was meant.
 */
function buildLookupIndex(workbook: Workbook, args: JoinSheetsArgs): LookupIndex {
  const sheet = getSheet(workbook, args.lookupSheet);
  const keyColumn = columnToIndex(args.lookupKeyColumn) ?? 0;
  const valueColumns = valueColumnsOf(args).map((column) => columnToIndex(column) ?? 0);
  const rows = new Map<string, CellValue[]>();
  const duplicates: string[] = [];

  for (let rowIndex = args.headerRow; rowIndex < (sheet?.rows.length ?? 0); rowIndex += 1) {
    const row = sheet?.rows[rowIndex];
    if (rowIsBlank(row)) continue;
    const key = keyFor(row?.[keyColumn]?.value ?? null);
    if (key === '') continue;
    if (rows.has(key)) {
      if (duplicates.length < 10 && !duplicates.includes(key)) duplicates.push(key);
      continue;
    }
    rows.set(
      key,
      valueColumns.map((column) => row?.[column]?.value ?? null),
    );
  }
  return { rows, duplicates };
}

function duplicateKeyWarnings(index: LookupIndex, args: JoinSheetsArgs): ValidationIssue[] {
  if (index.duplicates.length === 0) return [];
  const shown = index.duplicates
    .slice(0, 3)
    .map((key) => `"${key}"`)
    .join(', ');
  return [
    issue(
      'duplicate-lookup-key',
      `Lookup sheet "${args.lookupSheet}" repeats ${index.duplicates.length} key(s) (${shown}${
        index.duplicates.length > 3 ? ', ...' : ''
      }). The first occurrence wins.`,
    ),
  ];
}

/** The joined columns always land past every existing column, in every row, ragged or not. */
function joinInsertColumn(workbook: Workbook, sheetName: string): number {
  const sheet = getSheet(workbook, sheetName);
  return sheet ? maxColumnCount(sheet.rows) : 0;
}

function joinTarget(workbook: Workbook, args: JoinSheetsArgs): CellRange[] {
  const insertAt = joinInsertColumn(workbook, args.sheet);
  const lastColumn = insertAt + valueColumnsOf(args).length - 1;
  // An inner join removes rows, which moves every surviving cell, so the whole block genuinely
  // changes and the declared area has to cover it. A left join only appends.
  const startColumn =
    args.joinType === 'inner' ? 0 : paddingFloor(getSheet(workbook, args.sheet), insertAt);
  return [columnRange(workbook, args.sheet, 1, startColumn, lastColumn)];
}

function validateJoinSheets(workbook: Workbook, args: JoinSheetsArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateSheet(workbook, args.lookupSheet),
    ...headerRowIssues(workbook, args.sheet, args.headerRow),
    ...headerRowIssues(workbook, args.lookupSheet, args.headerRow),
    ...validateColumn(workbook, args.sheet, args.keyColumn),
    ...validateColumn(workbook, args.lookupSheet, args.lookupKeyColumn),
  ];
  for (const column of valueColumnsOf(args)) {
    errors.push(...validateColumn(workbook, args.lookupSheet, column));
  }
  if (valueColumnsOf(args).length !== headerNamesOf(args).length) {
    errors.push(
      issue(
        'header-name-mismatch',
        `headerName supplies ${headerNamesOf(args).length} name(s) for ${valueColumnsOf(args).length} lookup value column(s); they must line up.`,
      ),
    );
  }
  if (args.sheet.trim().toLowerCase() === args.lookupSheet.trim().toLowerCase()) {
    errors.push(issue('same-sheet', 'A join needs two different sheets.'));
  }
  return validationFailed(errors, duplicateKeyWarnings(buildLookupIndex(workbook, args), args));
}

function applyJoinSheets(workbook: Workbook, args: JoinSheetsArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const index = buildLookupIndex(workbook, args);
  const names = headerNamesOf(args);
  const sheet = getSheet(after, args.sheet);
  const headerRowIndex = args.headerRow - 1;
  const keyColumn = columnToIndex(args.keyColumn) ?? 0;
  let matched = 0;
  let unmatched = 0;

  if (sheet) {
    const insertAt = joinInsertColumn(workbook, args.sheet);
    const kept: Cell[][] = [];
    sheet.rows.forEach((row, rowIndex) => {
      const match = index.rows.get(keyFor(row[keyColumn]?.value ?? null));
      if (rowIndex > headerRowIndex) {
        if (match) matched += 1;
        else {
          unmatched += 1;
          if (args.joinType === 'inner') return;
        }
      }
      if (rowIndex === headerRowIndex) {
        names.forEach((name, offset) => insertCell(row, insertAt + offset, createCell(name)));
      } else {
        // Rows above the header are titles and banners, not joinable records: they stay blank
        // rather than picking up whichever lookup key happens to sit in their key column.
        const joined =
          rowIndex < headerRowIndex
            ? names.map((): CellValue => null)
            : (match ?? names.map(() => null));
        joined.forEach((value, offset) => insertCell(row, insertAt + offset, createCell(value)));
      }
      kept.push(row);
    });
    sheet.rows = kept;
  }

  const warnings = duplicateKeyWarnings(index, args);
  return transitionResult(
    before,
    after,
    operationReport(before, after, joinTarget(workbook, args), warnings, {
      addedColumns: names.length,
      removedRows: args.joinType === 'inner' ? unmatched : undefined,
      matchedRows: matched,
      unmatchedRows: unmatched,
    }),
  );
}

export const joinSheetsOperation: Operation<JoinSheetsArgs> = {
  name: 'join_sheets',
  schema: joinSheetsArgsSchema,
  targetRanges: joinTarget,
  validate: validateJoinSheets,
  preview(workbook, args) {
    const validation = validateJoinSheets(workbook, args);
    const ranges = joinTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyJoinSheets(workbook, args);
    // Dropping rows is irreversible from the user's point of view - the values only come back
    // through undo - so an inner join always asks first. A left join only appends.
    return previewForTransition(
      workbook,
      result.workbook,
      ranges,
      result.report.warnings,
      [],
      args.joinType === 'inner',
    );
  },
  apply: applyJoinSheets,
  invariants(before, after, args) {
    return runInvariants(before, after, { targetRanges: joinTarget(before, args) });
  },
};

// ============================================================================
// 4. FILL SERIES (fill_series)
// ============================================================================

interface ParsedRange {
  startRow: number;
  endRow: number;
  column: string;
  columnIndex: number;
}

/** Parses "A2", "A2:A3", and their spaced variants. Refuses anything multi-column. */
function parseRange(text: string): ParsedRange | undefined {
  const match = /^\s*([A-Za-z]+)\s*(\d+)\s*(?::\s*([A-Za-z]+)?\s*(\d+)?\s*)?$/.exec(text);
  if (!match) return undefined;
  const column = match[1] as string;
  const startRow = Number(match[2]);
  const endColumn = match[3];
  const endRow = match[4] !== undefined ? Number(match[4]) : startRow;
  const columnIndex = columnToIndex(column);
  if (columnIndex === undefined || !Number.isInteger(startRow) || startRow < 1) return undefined;
  if (endColumn !== undefined && columnToIndex(endColumn) !== columnIndex) return undefined;
  if (!Number.isInteger(endRow) || endRow < startRow) return undefined;
  return { startRow, endRow, column, columnIndex };
}

export const fillSeriesArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  column: rangeColumn,
  strategy: z.enum(['copy', 'linear', 'date', 'text']),
  sourceRange: z.string().trim().min(1),
  targetStartRow: z.number().int().positive(),
  targetEndRow: z.number().int().positive(),
  step: z.number().finite().optional(),
  dateUnit: z.enum(['day', 'week', 'month', 'year']).default('day'),
  headerRow: headerRowSchema,
});
export type FillSeriesArgs = z.infer<typeof fillSeriesArgsSchema>;

function fillSeriesTarget(workbook: Workbook, args: FillSeriesArgs): CellRange[] {
  const column = columnToIndex(args.column) ?? 0;
  return [
    columnRange(workbook, args.sheet, args.targetStartRow, column, column, args.targetEndRow),
  ];
}

const DAY_MS = 86_400_000;

function addMonths(date: Date, months: number): Date {
  const result = new Date(date.getTime());
  const day = result.getUTCDate();
  result.setUTCDate(1);
  result.setUTCMonth(result.getUTCMonth() + months);
  const daysInTargetMonth = new Date(
    Date.UTC(result.getUTCFullYear(), result.getUTCMonth() + 1, 0),
  ).getUTCDate();
  result.setUTCDate(Math.min(day, daysInTargetMonth));
  return result;
}

function addUnits(date: Date, amount: number, unit: 'day' | 'week' | 'month' | 'year'): Date {
  if (unit === 'month') return addMonths(date, amount);
  if (unit === 'year') return addMonths(date, amount * 12);
  return new Date(date.getTime() + amount * (unit === 'week' ? 7 : 1) * DAY_MS);
}

/**
 * How many `unit`s separate the last two source dates, or why the source cannot drive a series in
 * that unit. A zero interval is refused: it would fill the target with one repeated value.
 */
function dateInterval(
  previous: Date,
  last: Date,
  unit: 'day' | 'week' | 'month' | 'year',
): { ok: true; interval: number } | { ok: false; message: string } {
  const days = Math.round((last.getTime() - previous.getTime()) / DAY_MS);
  if (unit === 'day' || unit === 'week') {
    if (unit === 'week' && days % 7 !== 0) {
      return {
        ok: false,
        message: `The source dates are ${days} day(s) apart, which is not a whole number of weeks.`,
      };
    }
    const interval = unit === 'week' ? days / 7 : days;
    return interval === 0
      ? { ok: false, message: 'The source dates are the same day; pass an explicit step.' }
      : { ok: true, interval };
  }
  const interval =
    unit === 'month'
      ? (last.getUTCFullYear() - previous.getUTCFullYear()) * 12 +
        (last.getUTCMonth() - previous.getUTCMonth())
      : last.getUTCFullYear() - previous.getUTCFullYear();
  return interval === 0
    ? {
        ok: false,
        message: `The source dates fall in the same ${unit}; pass an explicit step.`,
      }
    : { ok: true, interval };
}

interface SeriesPlan {
  /** The value written `offset` rows below `targetStartRow`. */
  valueAt(offset: number): CellValue;
}

/**
 * Turns a source range into the rule that extends it. Every refusal is a message the caller can
 * act on, never a crash and never a silent default: a linear fill with one seed and no step, a
 * date fill over two identical dates, and a text fill over a value with no trailing number are all
 * cases where autofilling would invent data.
 */
function planSeries(
  sourceCells: Cell[],
  args: FillSeriesArgs,
): { ok: true; plan: SeriesPlan } | { ok: false; message: string } {
  const step = args.step ?? 1;
  const present = sourceCells.filter((cell) => !isBlankValue(cell.value));

  if (args.strategy === 'copy') {
    if (present.length === 0) {
      return { ok: false, message: 'The source range holds no values to copy.' };
    }
    const values = present.map((cell) => cell.value);
    return {
      ok: true,
      plan: { valueAt: (offset) => cloneValue(values[offset % values.length] as CellValue) },
    };
  }

  if (args.strategy === 'linear') {
    const numbers = present
      .map((cell) => toNumericOrNull(cell.value))
      .filter((value): value is number => value !== null);
    if (numbers.length === 0) {
      return { ok: false, message: 'A linear series needs at least one numeric source value.' };
    }
    const last = numbers[numbers.length - 1] as number;
    if (numbers.length === 1 && args.step === undefined) {
      return {
        ok: false,
        message:
          'A linear series needs two source values to infer a step, or an explicit step argument.',
      };
    }
    const inferred =
      numbers.length >= 2 ? last - (numbers[numbers.length - 2] as number) : (args.step ?? 1);
    if (numbers.length >= 2 && inferred === 0) {
      return {
        ok: false,
        message:
          'The last two source values are identical, so no step can be inferred; pass an explicit step.',
      };
    }
    const delta = args.step ?? inferred;
    return { ok: true, plan: { valueAt: (offset) => tidy(last + (offset + 1) * delta) } };
  }

  if (args.strategy === 'date') {
    const dates = present
      .map((cell) => (cell.value instanceof Date ? cell.value : null))
      .filter((date): date is Date => date !== null);
    if (dates.length === 0) {
      return { ok: false, message: 'A date series needs date values in the source range.' };
    }
    if (dates.length !== present.length) {
      return {
        ok: false,
        message: 'A date series needs date values in every non-empty source cell.',
      };
    }
    const last = dates[dates.length - 1] as Date;
    if (dates.length === 1) {
      return {
        ok: true,
        plan: { valueAt: (offset) => addUnits(last, (offset + 1) * step, args.dateUnit) },
      };
    }
    if (args.step !== undefined) {
      if (args.step === 0)
        return { ok: false, message: 'step must not be zero for a date series.' };
      return {
        ok: true,
        plan: { valueAt: (offset) => addUnits(last, (offset + 1) * args.step!, args.dateUnit) },
      };
    }
    const interval = dateInterval(dates[dates.length - 2] as Date, last, args.dateUnit);
    return interval.ok
      ? {
          ok: true,
          plan: {
            valueAt: (offset) => addUnits(last, (offset + 1) * interval.interval, args.dateUnit),
          },
        }
      : { ok: false, message: interval.message };
  }

  const lastCell = present[present.length - 1];
  const lastText = lastCell === undefined ? '' : valueText(lastCell.value).trim();
  const match = /^([\s\S]*?)(\d+)$/.exec(lastText);
  if (lastCell === undefined || match === null) {
    return {
      ok: false,
      message: `A text series needs a trailing number in the last source value (for example "Week1"); "${lastText}" has none.`,
    };
  }
  if (step === 0) return { ok: false, message: 'step must not be zero for a text series.' };
  const prefix = match[1] as string;
  const digits = match[2] as string;
  const width = digits.length;
  const base = Number(digits);
  return {
    ok: true,
    plan: {
      valueAt: (offset) => `${prefix}${String(base + (offset + 1) * step).padStart(width, '0')}`,
    },
  };
}

function sourceCellsOf(workbook: Workbook, args: FillSeriesArgs, source: ParsedRange): Cell[] {
  const sheet = getSheet(workbook, args.sheet);
  const cells: Cell[] = [];
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    cells.push(sheet?.rows[row - 1]?.[source.columnIndex] ?? createCell(null));
  }
  return cells;
}

function validateFillSeries(workbook: Workbook, args: FillSeriesArgs): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateColumn(workbook, args.sheet, args.column),
    ...headerRowIssues(workbook, args.sheet, args.headerRow),
  ];
  const source = parseRange(args.sourceRange);
  if (source === undefined) {
    errors.push(
      issue('invalid-source-range', `"${args.sourceRange}" is not a range such as "A2:A3".`),
    );
    return validationFailed(errors);
  }
  if (source.columnIndex !== (columnToIndex(args.column) ?? -1)) {
    errors.push(
      issue(
        'source-column-mismatch',
        `The source range is in column ${source.column.toUpperCase()} but the fill targets column ${args.column.toUpperCase()}.`,
      ),
    );
  }
  if (args.targetEndRow < args.targetStartRow) {
    errors.push(issue('invalid-range', 'targetEndRow must be at least targetStartRow.'));
  }
  if (args.targetStartRow <= source.endRow) {
    errors.push(
      issue(
        'overlapping-target',
        `The fill must start below the source range (row ${source.endRow + 1} or later); otherwise it overwrites its own seed.`,
      ),
    );
  }

  const sheet = getSheet(workbook, args.sheet);
  if (sheet) {
    if (source.endRow > sheet.rows.length) {
      errors.push(
        issue(
          'invalid-source-range',
          `The source range ends at row ${source.endRow}, past the end of "${args.sheet}".`,
        ),
      );
    }
    if (args.targetEndRow > sheet.rows.length) {
      errors.push(
        issue(
          'target-outside-sheet',
          `Row ${args.targetEndRow} is outside "${args.sheet}" (${sheet.rows.length} rows); add the rows first.`,
        ),
      );
    }
    const lastTarget = Math.min(args.targetEndRow, sheet.rows.length);
    for (let row = args.targetStartRow; row <= lastTarget; row += 1) {
      if (sheet.rows[row - 1]?.[source.columnIndex]?.formula !== undefined) {
        errors.push(
          issue(
            'formula-in-target',
            `Cell ${args.sheet}!${source.column.toUpperCase()}${row} holds a formula; fill_series will not overwrite a formula.`,
          ),
        );
        break;
      }
    }
  }

  if (errors.length === 0) {
    const plan = planSeries(sourceCellsOf(workbook, args, source), args);
    if (!plan.ok) errors.push(issue('unsupported-source', plan.message));
  }
  return validationFailed(errors);
}

function fillSeriesOverwritesContent(workbook: Workbook, args: FillSeriesArgs): boolean {
  const sheet = getSheet(workbook, args.sheet);
  const column = columnToIndex(args.column) ?? 0;
  const lastTarget = Math.min(args.targetEndRow, sheet?.rows.length ?? 0);
  for (let row = args.targetStartRow; row <= lastTarget; row += 1) {
    if (!isBlankValue(sheet?.rows[row - 1]?.[column]?.value)) return true;
  }
  return false;
}

function applyFillSeries(workbook: Workbook, args: FillSeriesArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const source = parseRange(args.sourceRange) as ParsedRange;
  const planned = planSeries(sourceCellsOf(workbook, args, source), args);
  // `validateFillSeries` already refused any source the strategy cannot use, so reaching the
  // failure branch here would mean the two paths disagree - which is a bug worth surfacing, not
  // a user input to paper over with a default series.
  if (!planned.ok) throw new Error(planned.message);
  const plan = planned.plan;
  const sheet = getSheet(after, args.sheet);
  if (sheet) {
    for (let row = args.targetStartRow; row <= args.targetEndRow; row += 1) {
      const target = sheet.rows[row - 1];
      if (!target) continue;
      target[source.columnIndex] = createCell(plan.valueAt(row - args.targetStartRow));
    }
  }
  return transitionResult(
    before,
    after,
    operationReport(before, after, fillSeriesTarget(workbook, args)),
  );
}

export const fillSeriesOperation: Operation<FillSeriesArgs> = {
  name: 'fill_series',
  schema: fillSeriesArgsSchema,
  targetRanges: fillSeriesTarget,
  validate: validateFillSeries,
  preview(workbook, args) {
    const validation = validateFillSeries(workbook, args);
    const ranges = fillSeriesTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyFillSeries(workbook, args);
    // Filling empty cells is the ordinary case and needs no second click. Replacing values that
    // are already there destroys data, so it asks.
    return previewForTransition(
      workbook,
      result.workbook,
      ranges,
      [],
      [],
      fillSeriesOverwritesContent(workbook, args),
    );
  },
  apply: applyFillSeries,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: fillSeriesTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

// ============================================================================
// 5. CATEGORIZE COLUMN (categorize_column)
// ============================================================================

const categorizeOperatorSchema = z.enum([
  'equals',
  'not_equals',
  'contains',
  'starts_with',
  'ends_with',
  'gt',
  'gte',
  'lt',
  'lte',
  'between',
]);

const ruleSchema = z
  .object({
    operator: categorizeOperatorSchema,
    // `between` carries a [low, high] pair; the bounds are numbers because every other operator
    // compares against a single value and a text range has no business in a numeric band.
    value: z.union([cellValueSchema, z.tuple([z.number(), z.number()])]),
    label: z.string().trim().min(1),
  })
  .refine(
    (rule) =>
      rule.operator === 'between' ? Array.isArray(rule.value) : !Array.isArray(rule.value),
    {
      message:
        'A between rule needs a [low, high] pair; every other operator needs a single value.',
      path: ['value'],
    },
  );

export type CategorizeRule = z.infer<typeof ruleSchema>;

export const categorizeColumnArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  sourceColumn: rangeColumn,
  newColumnName: z.string().trim().min(1),
  rules: z.array(ruleSchema).min(1),
  otherwise: z.string().trim().min(1),
  afterColumn: rangeColumn.optional(),
  headerRow: headerRowSchema,
});
export type CategorizeColumnArgs = z.infer<typeof categorizeColumnArgsSchema>;

function matchesRule(value: CellValue, rule: CategorizeRule): boolean {
  if (rule.operator === 'between') {
    const [low, high] = rule.value as [CellValue, CellValue];
    const numeric = toNumericOrNull(value);
    const from = toNumericOrNull(low);
    const to = toNumericOrNull(high);
    return numeric !== null && from !== null && to !== null && numeric >= from && numeric <= to;
  }
  const expected = rule.value as CellValue;
  const needle = valueText(expected).toLowerCase();
  const text = valueText(value).toLowerCase();
  switch (rule.operator) {
    case 'equals':
      return looselyEqual(value, expected);
    case 'not_equals':
      return !looselyEqual(value, expected);
    case 'contains':
      return text.includes(needle);
    case 'starts_with':
      return text.startsWith(needle);
    case 'ends_with':
      return text.endsWith(needle);
    case 'gt':
      return orderedComparison(value, expected, (comparison) => comparison > 0);
    case 'gte':
      return orderedComparison(value, expected, (comparison) => comparison >= 0);
    case 'lt':
      return orderedComparison(value, expected, (comparison) => comparison < 0);
    default:
      return orderedComparison(value, expected, (comparison) => comparison <= 0);
  }
}

/**
 * Ordering for gt/gte/lt/lte. A numeric bound only matches a numeric cell: "unknown" sorting after
 * 1000 because it is a string would put the text in the Large band, which is the kind of plausible
 * wrong answer that is worse than no label at all. Text bounds fall back to a text comparison.
 */
function orderedComparison(
  value: CellValue,
  expected: CellValue,
  accept: (comparison: number) => boolean,
): boolean {
  const numeric = toNumericOrNull(value);
  const bound = toNumericOrNull(expected);
  if (numeric !== null || bound !== null) {
    return numeric !== null && bound !== null && accept(numeric - bound);
  }
  return accept(compareValues(value, expected));
}

/** Rules are evaluated top to bottom and the first match wins; `otherwise` catches the rest. */
function labelFor(value: CellValue, args: CategorizeColumnArgs): string {
  for (const rule of args.rules) {
    if (matchesRule(value, rule)) return rule.label;
  }
  return args.otherwise;
}

function categorizeInsertIndex(workbook: Workbook, args: CategorizeColumnArgs): number {
  const sheet = getSheet(workbook, args.sheet);
  const width = sheet ? maxColumnCount(sheet.rows) : 0;
  return args.afterColumn === undefined ? width : (columnToIndex(args.afterColumn) ?? width) + 1;
}

function categorizeTarget(workbook: Workbook, args: CategorizeColumnArgs): CellRange[] {
  const sheet = getSheet(workbook, args.sheet);
  const insertion = categorizeInsertIndex(workbook, args);
  const width = sheet ? maxColumnCount(sheet.rows) : 0;
  // Everything from the insertion point onwards moves one column right, so the declared area ends
  // one past the widest column the sheet had. Appending writes exactly that column.
  const lastColumn = Math.max(insertion, width);
  return [columnRange(workbook, args.sheet, 1, paddingFloor(sheet, insertion), lastColumn)];
}

function existingHeaderNames(workbook: Workbook, args: CategorizeColumnArgs): string[] {
  const sheet = getSheet(workbook, args.sheet);
  const width = sheet ? maxColumnCount(sheet.rows) : 0;
  const names: string[] = [];
  for (let column = 0; column < width; column += 1) {
    names.push(
      valueText(sheet?.rows[args.headerRow - 1]?.[column]?.value ?? null)
        .trim()
        .toLowerCase(),
    );
  }
  return names;
}

function validateCategorizeColumn(
  workbook: Workbook,
  args: CategorizeColumnArgs,
): ValidationResult {
  const errors = [
    ...validateSheet(workbook, args.sheet),
    ...validateColumn(workbook, args.sheet, args.sourceColumn),
    ...headerRowIssues(workbook, args.sheet, args.headerRow),
  ];
  if (args.afterColumn !== undefined) {
    errors.push(...validateColumn(workbook, args.sheet, args.afterColumn));
  }
  // Two columns with the same header make every later reference to the column ambiguous, so one is
  // refused. Naming afterColumn states the placement explicitly and takes the decision off the
  // engine: the caller is saying they know which column they mean.
  if (
    args.afterColumn === undefined &&
    existingHeaderNames(workbook, args).includes(args.newColumnName.trim().toLowerCase())
  ) {
    errors.push(
      issue(
        'duplicate-column-name',
        `A column named "${args.newColumnName}" already exists on "${args.sheet}". Pass afterColumn to place the new column explicitly.`,
      ),
    );
  }
  return validationFailed(errors);
}

function applyCategorizeColumn(workbook: Workbook, args: CategorizeColumnArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  const sourceColumn = columnToIndex(args.sourceColumn) ?? 0;
  const insertion = categorizeInsertIndex(workbook, args);
  const headerRowIndex = args.headerRow - 1;

  sheet?.rows.forEach((row, rowIndex) => {
    // The source is read before the insert shifts every column to its right.
    const value = row[sourceColumn]?.value ?? null;
    const cell =
      rowIndex === headerRowIndex
        ? createCell(args.newColumnName)
        : rowIndex > headerRowIndex
          ? createCell(labelFor(value, args))
          : createCell(null);
    insertCell(row, insertion, cell);
  });

  return transitionResult(
    before,
    after,
    operationReport(before, after, categorizeTarget(workbook, args), [], { addedColumns: 1 }),
  );
}

export const categorizeColumnOperation: Operation<CategorizeColumnArgs> = {
  name: 'categorize_column',
  schema: categorizeColumnArgsSchema,
  targetRanges: categorizeTarget,
  validate: validateCategorizeColumn,
  preview(workbook, args) {
    const validation = validateCategorizeColumn(workbook, args);
    const ranges = categorizeTarget(workbook, args);
    if (!validation.valid) return invalidPreview(workbook, ranges, validation.errors);
    const result = applyCategorizeColumn(workbook, args);
    return previewForTransition(workbook, result.workbook, ranges);
  },
  apply: applyCategorizeColumn,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: categorizeTarget(before, args),
      rowCountUnchanged: true,
    });
  },
};

export const analyticsOperations: Operation<unknown>[] = [
  aggregateColumnOperation as Operation<unknown>,
  groupAndSummarizeOperation as Operation<unknown>,
  joinSheetsOperation as Operation<unknown>,
  fillSeriesOperation as Operation<unknown>,
  categorizeColumnOperation as Operation<unknown>,
];
