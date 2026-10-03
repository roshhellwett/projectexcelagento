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
}
