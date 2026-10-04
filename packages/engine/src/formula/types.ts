import type { DateSystem } from './excel-date.js';

export type FormulaValue = string | number | boolean | Date | null;

export interface FormulaCellAddress {
  sheet?: string;
  column: string;
  row: number;
}

export interface FormulaRangeAddress {
  sheet?: string;
  startColumn: string;
  startRow: number;
  endColumn: string;
  endRow: number;
}

export interface FormulaContext {
  activeSheet: string;
  getCellValue: (sheet: string, column: string, row: number) => FormulaValue;
  getRangeValues: (
    sheet: string,
    startColumn: string,
    startRow: number,
    endColumn: string,
    endRow: number,
  ) => FormulaValue[][];
  /**
   * Which epoch the workbook's date serials are counted from. 1900 is the default;
   * a workbook created in legacy Mac Excel uses 1904 and is otherwise off by 4 years.
   */
  dateSystem?: DateSystem;
  /**
   * Lets the evaluator tell "this cell is empty" apart from "that sheet does not exist".
   * Without it a typo like `=NoSuchSheet!A1` silently reads as a blank cell instead of
   * raising `#REF!`.
   */
  hasSheet?: (sheet: string) => boolean;
}
