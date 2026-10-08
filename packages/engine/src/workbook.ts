import type {
  Cell,
  CellLocation,
  CellPatch,
  CellStyle,
  CellType,
  CellValue,
  Patch,
  Workbook,
  WorkbookSnapshotPatch,
} from './types.js';

export function cellTypeForValue(value: CellValue, formula?: string): CellType {
  if (formula !== undefined) {
    return 'formula';
  }
  if (value === null || (typeof value === 'string' && value.length === 0)) {
    return 'blank';
  }
  if (value instanceof Date) {
    return 'date';
  }
  if (typeof value === 'string') {
    return 'string';
  }
  if (typeof value === 'number') {
    return 'number';
  }
  return 'boolean';
}

export function createCell(
  value: CellValue,
  options: { formula?: string; numberFormat?: string; type?: CellType; style?: CellStyle } = {},
): Cell {
  return {
    value,
    type: options.formula !== undefined ? 'formula' : (options.type ?? cellTypeForValue(value)),
    ...(options.formula !== undefined ? { formula: options.formula } : {}),
    ...(options.numberFormat !== undefined ? { numberFormat: options.numberFormat } : {}),
    ...(options.style !== undefined ? { style: { ...options.style } } : {}),
  };
}

export function cloneCell(cell: Cell): Cell {
  return createCell(cell.value instanceof Date ? new Date(cell.value.getTime()) : cell.value, {
    type: cell.formula === undefined ? cell.type : 'formula',
    ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
    ...(cell.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
    ...(cell.style !== undefined ? { style: { ...cell.style } } : {}),
  });
}

export function cloneWorkbook(workbook: Workbook): Workbook {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      rows: sheet.rows.map((row) => row.map(cloneCell)),
    })),
  };
}

export function getSheet(workbook: Workbook, name: string) {
  return workbook.sheets.find((sheet) => sheet.name === name);
}

export function columnToIndex(column: string): number | undefined {
  const normalized = column.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    return undefined;
  }

  let index = 0;
  for (const character of normalized) {
    index = index * 26 + character.charCodeAt(0) - 64;
    if (!Number.isSafeInteger(index)) return undefined;
  }
  return index - 1;
}

export function indexToColumn(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`Column index must be a non-negative integer: ${index}`);
  }

  let remaining = index + 1;
  let column = '';
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return column;
}

export function locationFor(sheet: string, row: number, columnIndex: number): CellLocation {
  return { sheet, row, column: indexToColumn(columnIndex) };
}

export function getCell(workbook: Workbook, location: CellLocation): Cell | undefined {
  const sheet = getSheet(workbook, location.sheet);
  const columnIndex = columnToIndex(location.column);
  if (!sheet || columnIndex === undefined) {
    return undefined;
  }
  return sheet.rows[location.row - 1]?.[columnIndex];
}

export function setCell(workbook: Workbook, location: CellLocation, cell: Cell): void {
  const sheet = getSheet(workbook, location.sheet);
  const columnIndex = columnToIndex(location.column);
  if (!sheet || columnIndex === undefined || location.row < 1) {
    throw new Error(
      `Cell location does not exist: ${location.sheet}!${location.column}${location.row}`,
    );
  }

  const row = sheet.rows[location.row - 1];
  if (!row) {
    throw new Error(`Row does not exist: ${location.sheet}!${location.column}${location.row}`);
  }
  row[columnIndex] = cloneCell(cell);
}

export function cellValueEquals(left: CellValue, right: CellValue): boolean {
  if (left instanceof Date && right instanceof Date) {
    return left.getTime() === right.getTime();
  }
  if (typeof left === 'number' && typeof right === 'number') {
    // `NaN === NaN` is false, so a single NaN cell made `workbookEquals` fail
    // forever. The operation registry verifies that a patch restores the
    // workbook before and after every operation, so one NaN anywhere would fail
    // EVERY subsequent operation with "The inverse patch does not restore the
    // original workbook." and permanently lock the user out of their workbook.
    // NaN is treated as equal to NaN so a round-tripped NaN stays a no-op.
    return left === right || (Number.isNaN(left) && Number.isNaN(right));
  }
  return left === right;
}

export function effectiveCellType(cell: Cell): CellType {
  return cell.type ?? cellTypeForValue(cell.value, cell.formula);
}

/**
 * Structural equality for a single cell. One rule for blanks: `createCell(null)`
 * and `createCell('')` are the same empty cell. `cellTypeForValue` already types
 * both as `blank`, so treating them as different made a patch that normalized
 * '' to null (or null to '') look like a real change forever, which broke
 * round-trips and made empty-string normalization impossible. Callers that need
 * the literal difference (for example filter matching on `equals`) should use
 * `cellValueEquals`, which stays strict.
 */
export function cellEquals(left: Cell, right: Cell): boolean {
  return (
    cellValueEquals(blankEquivalent(left.value), blankEquivalent(right.value)) &&
    effectiveCellType(left) === effectiveCellType(right) &&
    left.formula === right.formula &&
    left.numberFormat === right.numberFormat &&
    stylesEqual(left.style, right.style)
  );
}

function stylesEqual(left: CellStyle | undefined, right: CellStyle | undefined): boolean {
  return (
    left?.bold === right?.bold &&
    left?.italic === right?.italic &&
    left?.underline === right?.underline &&
    left?.fillColor === right?.fillColor &&
    left?.fontColor === right?.fontColor &&
    left?.horizontalAlignment === right?.horizontalAlignment &&
    left?.verticalAlignment === right?.verticalAlignment &&
    left?.wrapText === right?.wrapText
  );
}

function blankEquivalent(value: CellValue): CellValue {
  return value === '' ? null : value;
}

export function patchForChange(address: CellLocation, before: Cell, after: Cell): CellPatch {
  return {
    kind: 'cell',
    address,
    oldValue: before.value instanceof Date ? new Date(before.value.getTime()) : before.value,
    oldFormula: before.formula,
    oldType: effectiveCellType(before),
    oldNumberFormat: before.numberFormat,
    oldStyle: before.style ? { ...before.style } : undefined,
    newValue: after.value instanceof Date ? new Date(after.value.getTime()) : after.value,
    newFormula: after.formula,
    newType: effectiveCellType(after),
    newNumberFormat: after.numberFormat,
    newStyle: after.style ? { ...after.style } : undefined,
  };
}

function cellFromPatch(patch: CellPatch, side: 'old' | 'new'): Cell {
  if (side === 'old') {
    return createCell(
      patch.oldValue instanceof Date ? new Date(patch.oldValue.getTime()) : patch.oldValue,
      {
        type: patch.oldType,
        ...(patch.oldFormula !== undefined ? { formula: patch.oldFormula } : {}),
        ...(patch.oldNumberFormat !== undefined ? { numberFormat: patch.oldNumberFormat } : {}),
        ...(patch.oldStyle !== undefined ? { style: { ...patch.oldStyle } } : {}),
      },
    );
  }
  return createCell(
    patch.newValue instanceof Date ? new Date(patch.newValue.getTime()) : patch.newValue,
    {
      type: patch.newType,
      ...(patch.newFormula !== undefined ? { formula: patch.newFormula } : {}),
      ...(patch.newNumberFormat !== undefined ? { numberFormat: patch.newNumberFormat } : {}),
      ...(patch.newStyle !== undefined ? { style: { ...patch.newStyle } } : {}),
    },
  );
}

export function snapshotPatch(before: Workbook, after: Workbook): WorkbookSnapshotPatch {
  return {
    kind: 'workbook',
    oldWorkbook: cloneWorkbook(before),
    newWorkbook: cloneWorkbook(after),
  };
}

export function applyPatch(workbook: Workbook, patch: Patch): Workbook {
  let result = cloneWorkbook(workbook);
  for (const change of patch) {
    if (change.kind === 'workbook') {
      result = cloneWorkbook(change.newWorkbook);
      continue;
    }
    setCell(result, change.address, cellFromPatch(change, 'new'));
  }
  return result;
}

function invertEntry(entry: Patch[number]): Patch[number] {
  if (entry.kind === 'workbook') {
    return snapshotPatch(entry.newWorkbook, entry.oldWorkbook);
  }
  return {
    kind: 'cell',
    address: { ...entry.address },
    oldValue: entry.newValue instanceof Date ? new Date(entry.newValue.getTime()) : entry.newValue,
    oldFormula: entry.newFormula,
    oldType: entry.newType,
    oldNumberFormat: entry.newNumberFormat,
    oldStyle: entry.newStyle ? { ...entry.newStyle } : undefined,
    newValue: entry.oldValue instanceof Date ? new Date(entry.oldValue.getTime()) : entry.oldValue,
    newFormula: entry.oldFormula,
    newType: entry.oldType,
    newNumberFormat: entry.oldNumberFormat,
    newStyle: entry.oldStyle ? { ...entry.oldStyle } : undefined,
  };
}

export function invertPatch(patch: Patch): Patch {
  return patch.slice().reverse().map(invertEntry);
}

export function maxColumnCount(rows: Cell[][]): number {
  return rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
}

function effectiveDateSystem(workbook: Workbook): '1900' | '1904' {
  return workbook.dateSystem ?? '1900';
}

function sameShape(left: Workbook, right: Workbook): boolean {
  return (
    left.sheets.length === right.sheets.length &&
    left.sheets.every((sheet, index) => {
      const other = right.sheets[index];
      return (
        other !== undefined &&
        sheet.name === other.name &&
        sheet.rows.length === other.rows.length &&
        sheet.rows.every((row, rowIndex) => row.length === (other.rows[rowIndex]?.length ?? -1))
      );
    })
  );
}

export function patchBetween(before: Workbook, after: Workbook): Patch {
  if (effectiveDateSystem(before) !== effectiveDateSystem(after) || !sameShape(before, after)) {
    return [snapshotPatch(before, after)];
  }

  const changes: Patch = [];
  before.sheets.forEach((sheet, sheetIndex) => {
    const afterSheet = after.sheets[sheetIndex];
    if (!afterSheet) {
      return;
    }
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((beforeCell, columnIndex) => {
        const afterCell = afterSheet.rows[rowIndex]?.[columnIndex];
        if (afterCell && !cellEquals(beforeCell, afterCell)) {
          changes.push(
            patchForChange(
              locationFor(sheet.name, rowIndex + 1, columnIndex),
              beforeCell,
              afterCell,
            ),
          );
        }
      });
    });
  });
  return changes;
}

export function workbookEquals(left: Workbook, right: Workbook): boolean {
  if (effectiveDateSystem(left) !== effectiveDateSystem(right) || !sameShape(left, right)) {
    return false;
  }
  return left.sheets.every((sheet, sheetIndex) => {
    const other = right.sheets[sheetIndex];
    return (
      other !== undefined &&
      sheet.rows.every((row, rowIndex) =>
        row.every((cell, columnIndex) => {
          const otherCell = other.rows[rowIndex]?.[columnIndex];
          return otherCell !== undefined && cellEquals(cell, otherCell);
        }),
      )
    );
  });
}
