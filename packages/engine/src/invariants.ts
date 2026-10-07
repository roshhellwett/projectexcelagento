import type { Cell, CellRange, InvariantOptions, InvariantResult, Workbook } from './types.js';
import { cellEquals, columnToIndex, indexToColumn } from './workbook.js';

function effectiveDateSystem(workbook: Workbook): '1900' | '1904' {
  return workbook.dateSystem ?? '1900';
}

function rangeBounds(range: CellRange): { startColumn: number; endColumn: number } | undefined {
  const startColumn = columnToIndex(range.startColumn);
  const endColumn = columnToIndex(range.endColumn);
  if (
    startColumn === undefined ||
    endColumn === undefined ||
    range.startRow < 1 ||
    range.endRow < range.startRow ||
    endColumn < startColumn
  ) {
    return undefined;
  }
  return { startColumn, endColumn };
}

function rangeContains(range: CellRange, sheet: string, row: number, column: number): boolean {
  const bounds = rangeBounds(range);
  return (
    bounds !== undefined &&
    range.sheet === sheet &&
    row >= range.startRow &&
    row <= range.endRow &&
    column >= bounds.startColumn &&
    column <= bounds.endColumn
  );
}

function isInRanges(ranges: CellRange[], sheet: string, row: number, column: number): boolean {
  return ranges.some((range) => rangeContains(range, sheet, row, column));
}

function valueFingerprint(cell: Cell | undefined): string {
  if (!cell) {
    return '<missing>';
  }
  const value = cell.value instanceof Date ? cell.value.toISOString() : cell.value;
  return JSON.stringify([
    cell.type,
    value,
    cell.formula ?? null,
    cell.numberFormat ?? null,
    cell.style ?? null,
  ]);
}

function rowFingerprint(row: Cell[] | undefined, startColumn: number, endColumn: number): string {
  const cells: string[] = [];
  for (let column = startColumn; column <= endColumn; column += 1) {
    cells.push(valueFingerprint(row?.[column]));
  }
  return JSON.stringify(cells);
}

function sortedCounts(values: string[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const value of values) {
    counts.set(value, (counts.get(value) ?? 0) + 1);
  }
  return counts;
}

function mapsEqual(left: Map<string, number>, right: Map<string, number>): boolean {
  if (left.size !== right.size) {
    return false;
  }
  for (const [key, count] of left) {
    if (right.get(key) !== count) {
      return false;
    }
  }
  return true;
}

export function invariantRowCountUnchanged(before: Workbook, after: Workbook): string[] {
  const errors: string[] = [];
  if (before.sheets.length !== after.sheets.length) {
    return ['The operation changed the number of sheets.'];
  }
  before.sheets.forEach((sheet, index) => {
    const afterSheet = after.sheets[index];
    if (!afterSheet || afterSheet.name !== sheet.name) {
      errors.push(`Sheet structure changed for "${sheet.name}".`);
    } else if (sheet.rows.length !== afterSheet.rows.length) {
      errors.push(`Row count changed for "${sheet.name}".`);
    }
  });
  return errors;
}

export function invariantRowsMultisetEqual(
  before: Workbook,
  after: Workbook,
  range: CellRange,
): string[] {
  const bounds = rangeBounds(range);
  const beforeSheet = before.sheets.find((sheet) => sheet.name === range.sheet);
  const afterSheet = after.sheets.find((sheet) => sheet.name === range.sheet);
  if (!bounds || !beforeSheet || !afterSheet) {
    return [`Cannot compare sorted rows for "${range.sheet}".`];
  }

  const beforeRows = beforeSheet.rows
    .slice(range.startRow - 1, range.endRow)
    .map((row) => rowFingerprint(row, bounds.startColumn, bounds.endColumn));
  const afterRows = afterSheet.rows
    .slice(range.startRow - 1, range.endRow)
    .map((row) => rowFingerprint(row, bounds.startColumn, bounds.endColumn));
  return mapsEqual(sortedCounts(beforeRows), sortedCounts(afterRows))
    ? []
    : [
        `Rows in ${range.sheet}!${range.startColumn}${range.startRow}:${range.endColumn}${range.endRow} were not preserved.`,
      ];
}

export function invariantFormulaCellsRemainFormulas(
  before: Workbook,
  after: Workbook,
  allowFormulaChanges = false,
): string[] {
  if (allowFormulaChanges) {
    return [];
  }

  const formulas = (workbook: Workbook): string[] =>
    workbook.sheets.flatMap((sheet) =>
      sheet.rows.flatMap((row) =>
        row.flatMap((cell) => (cell.formula === undefined ? [] : [cell.formula])),
      ),
    );
  return mapsEqual(sortedCounts(formulas(before)), sortedCounts(formulas(after)))
    ? []
    : ['Formula cells changed or were removed without explicit permission.'];
}

export function invariantNoCellsOutsideTargetRange(
  before: Workbook,
  after: Workbook,
  ranges: CellRange[],
): string[] {
  const errors: string[] = [];
  if (effectiveDateSystem(before) !== effectiveDateSystem(after)) {
    errors.push('Workbook date system metadata changed outside the declared cell target ranges.');
  }
  const allSheetNames = new Set([
    ...before.sheets.map((sheet) => sheet.name),
    ...after.sheets.map((sheet) => sheet.name),
  ]);

  for (const sheetName of allSheetNames) {
    const beforeSheet = before.sheets.find((sheet) => sheet.name === sheetName);
    const afterSheet = after.sheets.find((sheet) => sheet.name === sheetName);
    const rowCount = Math.max(beforeSheet?.rows.length ?? 0, afterSheet?.rows.length ?? 0);
    for (let row = 1; row <= rowCount; row += 1) {
      const beforeRow = beforeSheet?.rows[row - 1];
      const afterRow = afterSheet?.rows[row - 1];
      const columnCount = Math.max(beforeRow?.length ?? 0, afterRow?.length ?? 0);
      for (let column = 0; column < columnCount; column += 1) {
        const beforeCell = beforeRow?.[column];
        const afterCell = afterRow?.[column];
        const changed =
          beforeCell === undefined
            ? afterCell !== undefined
            : afterCell === undefined || !cellEquals(beforeCell, afterCell);
        if (changed && !isInRanges(ranges, sheetName, row, column)) {
          errors.push(
            `Cell ${sheetName}!${indexToColumn(column)}${row} changed outside the declared target range.`,
          );
        }
      }
    }
  }
  return errors;
}

export function runInvariants(
  before: Workbook,
  after: Workbook,
  options: InvariantOptions,
): InvariantResult {
  const errors: string[] = [];
  if (options.rowCountUnchanged) {
    errors.push(...invariantRowCountUnchanged(before, after));
  }
  if (options.rowsMultisetEqual) {
    for (const range of options.targetRanges) {
      errors.push(...invariantRowsMultisetEqual(before, after, range));
    }
  }
  errors.push(
    ...invariantFormulaCellsRemainFormulas(before, after, options.allowFormulaChanges ?? false),
  );
  errors.push(...invariantNoCellsOutsideTargetRange(before, after, options.targetRanges));
  return { valid: errors.length === 0, errors };
}
