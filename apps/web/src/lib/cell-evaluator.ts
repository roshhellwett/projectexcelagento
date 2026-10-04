import { useMemo } from 'react';
import {
  type Cell,
  type DateSystem,
  type FormulaValue,
  type Sheet,
  type Workbook,
  columnToIndex,
  evaluateFormula,
} from '@excel-agent/engine';

import { formatCellDateValue } from './cell-format.js';

/**
 * A formula evaluator bound to one workbook snapshot.
 *
 * `cell.formula` stores the leading `=`, so the text is handed to the engine untouched. A cell
 * without a formula returns `undefined` rather than its value: a number or a date needs no
 * evaluation, and pretending otherwise would make every plain cell pay for the evaluator.
 */
export type CellEvaluator = (
  cell: Cell | undefined,
  sheetName: string,
  colIdx: number,
  rowNumber: number,
) => FormulaValue | undefined;

/**
 * Builds the evaluator for the live workbook.
 *
 * Results are memoised per workbook snapshot, so a formula that a hundred cells reference is
 * evaluated once per edit instead of a hundred times, and nothing is re-evaluated at all until the
 * workbook actually changes. Circular references resolve to `null` instead of recursing forever,
 * which is what a spreadsheet shows for one.
 */
export function useCellEvaluator(workbook: Workbook, dateSystem: DateSystem): CellEvaluator {
  return useMemo(() => {
    const cache = new Map<string, FormulaValue>();
    const visiting = new Set<string>();
    const findSheet = (name: string): Sheet | undefined =>
      workbook.sheets.find((sheet) => sheet.name === name);

    const evaluateAddress = (
      sheetName: string,
      colIdx: number,
      rowNumber: number,
    ): FormulaValue => {
      const key = `${sheetName}!${colIdx}:${rowNumber}`;
      const cached = cache.get(key);
      if (cached !== undefined) return cached;
      if (visiting.has(key)) return null;
      visiting.add(key);
      try {
        const cell = findSheet(sheetName)?.rows[rowNumber - 1]?.[colIdx];
        let value: FormulaValue;
        if (cell?.formula) {
          value = evaluateFormula(cell.formula, {
            activeSheet: sheetName,
            getCellValue: (sheet: string, column: string, row: number) =>
              evaluateAddress(sheet, columnToIndex(column) ?? 0, row),
            getRangeValues: (
              sheet: string,
              startCol: string,
              startRow: number,
              endCol: string,
              endRow: number,
            ) => {
              const target = findSheet(sheet);
              const from = columnToIndex(startCol) ?? 0;
              const to = columnToIndex(endCol) ?? from;
              const out: FormulaValue[][] = [];
              for (let row = startRow; row <= endRow; row += 1) {
                const line: FormulaValue[] = [];
                for (let colIdx = from; colIdx <= to; colIdx += 1) {
                  const cellInRange = target?.rows[row - 1]?.[colIdx];
                  line.push(
                    cellInRange?.formula
                      ? evaluateAddress(sheet, colIdx, row)
                      : (cellInRange?.value ?? null),
                  );
                }
                out.push(line);
              }
              return out;
            },
            // Without this a 1904 workbook's dates render correctly but every date formula in it
            // evaluates four years and a day off.
            dateSystem,
            hasSheet: (name: string) => workbook.sheets.some((sheet) => sheet.name === name),
          });
        } else {
          value = (cell?.value ?? null) as FormulaValue;
        }
        cache.set(key, value);
        return value;
      } catch {
        // A formula the engine cannot parse shows whatever was cached for that address rather than
        // taking the whole grid down with it.
        return cache.get(key) ?? null;
      } finally {
        visiting.delete(key);
      }
    };

    return (cell, sheetName, colIdx, rowNumber) => {
      if (!cell?.formula) return undefined;
      // A formula parsed from an address has no cell of its own to fall back to, so a failure
      // resolves to null (Excel's own #VALUE! behaviour for unparseable text) rather than throwing.
      return evaluateAddress(sheetName, colIdx, rowNumber);
    };
  }, [workbook, dateSystem]);
}

export interface CellDisplay {
  text: string;
  /** The value-class suffix the stylesheet keys formatting off. */
  className: string;
}

/**
 * How one cell is drawn - and the same answer the clipboard copies, so a copied cell carries the
 * text the user can see rather than the raw value underneath it.
 *
 * A numeric cell is a date only when its number format says so; the shared module owns that
 * decision and the serial-to-date conversion, so the grid never does date arithmetic itself. The
 * importer promotes date-formatted numbers to real `Date` values, so a cell is a date either
 * because its value already is one or because its number format says the bare serial is one.
 */
export function cellDisplay(
  cell: Cell | undefined,
  evaluated: FormulaValue | undefined,
  dateSystem: DateSystem,
): CellDisplay {
  if (cell?.formula) {
    return {
      text: evaluated === null || evaluated === undefined ? '' : String(evaluated),
      className: 'cell-val-formula-result',
    };
  }

  const value = cell?.value;
  const dateText =
    value === null || value === undefined
      ? null
      : formatCellDateValue(value, cell?.numberFormat, dateSystem);

  if (dateText !== null) return { text: dateText, className: 'cell-val-date' };
  if (typeof value === 'number')
    return { text: value.toLocaleString(), className: 'cell-val-number' };
  if (value === null || value === undefined || value === '') {
    return { text: '', className: 'cell-val-null' };
  }
  return { text: String(value), className: '' };
}

/** The formula text as the formula bar shows it: with its leading `=`, exactly once. */
export function formulaTextFor(cell: Cell | undefined): string {
  if (cell?.formula === undefined) return '';
  return cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
}
