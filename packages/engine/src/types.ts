import type { ZodType } from 'zod';

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

export interface Report {
  affectedCells: number;
  skippedCells: number;
  unchangedCells: number;
  warnings: ValidationIssue[];
  removedRows?: number;
  addedRows?: number;
  deletedColumns?: number;
  addedColumns?: number;
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
  inverse: Patch;
}
