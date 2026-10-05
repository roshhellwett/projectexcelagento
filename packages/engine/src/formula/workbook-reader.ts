import type { Workbook } from '../types.js';
import { columnToIndex } from '../workbook.js';
import { evaluateFormula } from './evaluator.js';
import type { FormulaValue } from './types.js';

/** Memoized live values for one immutable snapshot; never trusts a stale formula cache. */
export function createWorkbookValueReader(workbook: Workbook) {
  const sheets = new Map(workbook.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]));
  const cache = new Map<string, FormulaValue>();
  const visiting = new Set<string>();
  const read = (sheetName: string, columnIndex: number, rowNumber: number): FormulaValue => {
    const sheet = sheets.get(sheetName.toLowerCase());
    if (!sheet) return '#REF!';
    const key = JSON.stringify([sheet.name, columnIndex, rowNumber]);
    if (cache.has(key)) return cache.get(key)!;
    if (visiting.has(key) || visiting.size >= 64) return '#ERROR!';
    const cell = sheet.rows[rowNumber - 1]?.[columnIndex];
    if (!cell?.formula) return cell?.value ?? null;
    visiting.add(key);
    try {
      const value = evaluateFormula(cell.formula, {
        activeSheet: sheet.name, dateSystem: workbook.dateSystem,
        hasSheet: (name) => sheets.has(name.toLowerCase()),
        getCellValue: (name, column, row) => read(name, columnToIndex(column) ?? -1, row),
        getRangeValues: (name, startColumn, startRow, endColumn, endRow) => {
          const from = columnToIndex(startColumn) ?? -1;
          const to = columnToIndex(endColumn) ?? -1;
          const count = (endRow - startRow + 1) * (to - from + 1);
          if (from < 0 || to < from || startRow < 1 || endRow < startRow) return [['#REF!']];
          if (count > 1_500_000) return [['#NUM!']];
          const values: FormulaValue[][] = [];
          for (let row = startRow; row <= endRow; row += 1) {
            const line: FormulaValue[] = [];
            for (let column = from; column <= to; column += 1) line.push(read(name, column, row));
            values.push(line);
          }
          return values;
        },
      });
      cache.set(key, value);
      return value;
    } finally {
      visiting.delete(key);
    }
  };
  return read;
}
