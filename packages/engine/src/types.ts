import type { ZodType } from 'zod';
import type { DateSystem } from './formula/excel-date.js';

export type CellValue = string | number | boolean | Date | null;
export type CellType = 'blank' | 'string' | 'number' | 'boolean' | 'date' | 'formula';

export interface Cell {
  value: CellValue;
  type: CellType;
  formula?: string;
  numberFormat?: string;
}

export interface Sheet {
  name: string;
  rows: Cell[][];
}

export interface Workbook {
  sheets: Sheet[];
  /**
   * Which epoch this workbook's date serials are counted from.
   *
   * This belongs to the workbook, not to any one reader: a legacy Mac Excel file written in
   * the 1904 system renders correctly in the grid but evaluates its formulas four years and a
   * day off unless every consumer reads the epoch from here. Optional and defaulting to '1900'
   * so existing workbooks and callers keep working unchanged.
   */
  dateSystem?: DateSystem;
}

export interface CellLocation {
  sheet: string;
  row: number;
  column: string;
}

export interface CellRange {
  sheet: string;
  startRow: number;
  endRow: number;
  startColumn: string;
  endColumn: string;
}

export interface ValidationIssue {
  code: string;
  message: string;
  location?: CellLocation;
}

export interface ValidationResult {
  valid: boolean;
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
}

export interface PreviewChange {
  location: CellLocation;
  before: Cell;
  after: Cell;
}

export interface Preview {
  valid: boolean;
  affectedCells: number;
  changes: PreviewChange[];
  warnings: ValidationIssue[];
  errors: ValidationIssue[];
  requiresConfirmation: boolean;
}

export interface WorkbookArtifact {
  id: string;
  kind: 'reconciliation';
  title: string;
  sheets: { name: string; role: 'summary' | 'matched' | 'exceptions' | 'methodology' }[];
  sources: CellRange[];
  facts: { label: string; value: string }[];
  checks: { id: string; label: string; status: 'passed' | 'warning' | 'failed'; detail: string }[];
  notes: string[];
}

export interface Report {
  /** Actual generated deliverables; never inferred from model prose or a proposed preview. */
  artifacts?: WorkbookArtifact[];
  affectedCells: number;
  skippedCells: number;
  unchangedCells: number;
  warnings: ValidationIssue[];
  removedRows?: number;
  addedRows?: number;
  deletedColumns?: number;
  addedColumns?: number;
  /**
   * The computed aggregate of an `aggregate_column` call. Always a number: an aggregation that
   * matched nothing reports 0 plus a warning rather than a missing value, so a caller reading the
   * result never has to guess whether the number is real.
   */
  aggregate?: number;
  /** Rows or groups a criteria, a join, or an aggregation could not resolve to a value. */
  unmatchedRows?: number;
  /** Rows of a join that found a matching lookup key. */
  matchedRows?: number;
}

/** A reversible change to one cell. Addresses and rows are one-based. */
export interface CellPatch {
  kind: 'cell';
  address: CellLocation;
  oldValue: CellValue;
  oldFormula?: string;
  oldType: CellType;
  oldNumberFormat?: string;
  newValue: CellValue;
  newFormula?: string;
  newType: CellType;
  newNumberFormat?: string;
}

/** Structural operations use a snapshot entry because row/column shape can change. */
export interface WorkbookSnapshotPatch {
  kind: 'workbook';
  oldWorkbook: Workbook;
  newWorkbook: Workbook;
}

export type PatchEntry = CellPatch | WorkbookSnapshotPatch;
export type Patch = PatchEntry[];

export interface InvariantResult {
  valid: boolean;
  errors: string[];
}

export interface InvariantOptions {
  targetRanges: CellRange[];
  rowCountUnchanged?: boolean;
  rowsMultisetEqual?: boolean;
  allowFormulaChanges?: boolean;
}

export interface OperationResult {
  workbook: Workbook;
  report: Report;
  patch: Patch;
  inverse: Patch;
}

export interface Operation<Args> {
  name: string;
  schema: ZodType<Args>;
  targetRanges(workbook: Workbook, args: Args): CellRange[];
  validate(workbook: Workbook, args: Args): ValidationResult;
  preview(workbook: Workbook, args: Args): Preview;
  apply(workbook: Workbook, args: Args): OperationResult;
  invariants(before: Workbook, after: Workbook, args: Args): InvariantResult;
}

export interface HistoryEntry {
  operationName: string;
  patch: Patch;
}
