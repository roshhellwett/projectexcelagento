import { describe, expect, it } from 'vitest';

import {
  columnToIndex,
  createCell,
  indexToColumn,
  maxColumnCount,
  setCellsOperation,
  type Cell,
  type CellRange,
  type Workbook,
} from '../src/index.js';
import { previewForTransition } from '../src/operation-utils.js';

function sheetOfRows(rowCount: number, columnCount: number): Workbook {
  const rows: Cell[][] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const cells: Cell[] = [];
    for (let column = 0; column < columnCount; column += 1) {
      cells.push(createCell(row * columnCount + column));
    }
    rows.push(cells);
  }
  return { sheets: [{ name: 'Data', rows }] };
}

/** 500 x 500 = 250,000 cells. */
function largeWorkbook(): Workbook {
  return sheetOfRows(500, 500);
}

function scatteredRanges(count: number, rows: number, columns: number): CellRange[] {
  const ranges: CellRange[] = [];
  for (let index = 0; index < count; index += 1) {
    const startColumn = (index * 2) % columns;
    const endColumn = Math.min(columns - 1, startColumn + 1);
    const startRow = ((index * 5) % rows) + 1;
    ranges.push({
      sheet: 'Data',
      startRow,
      endRow: Math.min(rows, startRow + 4),
      startColumn: indexToColumn(startColumn),
      endColumn: indexToColumn(endColumn),
    });
  }
  return ranges;
}

function milliseconds(run: () => void): number {
  const started = performance.now();
  run();
  return performance.now() - started;
}

function addressesInRanges(ranges: CellRange[]): Set<string> {
  const covered = new Set<string>();
  for (const range of ranges) {
    const startColumn = columnToIndex(range.startColumn);
    const endColumn = columnToIndex(range.endColumn);
    if (startColumn === undefined || endColumn === undefined) {
      throw new Error(`Range ${range.startColumn}${range.startRow} is not parseable.`);
    }
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        covered.add(`${range.sheet}!${indexToColumn(column)}${row}`);
      }
    }
  }
  return covered;
}

describe('previewForTransition at sheet scale', () => {
  it('checks 250k cells against 200 declared ranges well under a second', () => {
    const before = largeWorkbook();
    const after = largeWorkbook();
    const ranges = scatteredRanges(200, 500, 500);
    // Change two cells that are genuinely covered by two of the declared ranges.
    const targets = [ranges[3]!, ranges[7]!];
    for (const range of targets) {
      const column = columnToIndex(range.startColumn) as number;
      after.sheets[0]!.rows[range.startRow - 1]![column] = createCell('changed');
    }

    const preview = previewForTransition(before, after, ranges);
    const elapsed = milliseconds(() => {
      previewForTransition(before, after, ranges);
    });

    expect(preview.affectedCells).toBe(2);
    expect(elapsed).toBeLessThan(1000);
  });

  it('only counts changed cells that fall inside a declared range', () => {
    const before = largeWorkbook();
    const after = largeWorkbook();
    const range = scatteredRanges(200, 500, 500)[3]!;
    const column = columnToIndex(range.startColumn) as number;
    after.sheets[0]!.rows[range.startRow - 1]![column] = createCell('inside');
    // A change far outside every declared range must not be counted.
    after.sheets[0]!.rows[0]![0] = createCell('outside');

    expect(previewForTransition(before, after, [range]).affectedCells).toBe(1);
    expect(previewForTransition(before, after, []).affectedCells).toBe(0);
  });
});

describe('set_cells target ranges', () => {
  const args = (cells: { row: number; column: string; value: number }[]) => ({
    sheet: 'Data',
    cells,
  });

  it('coalesces a 300-cell batch in 3 columns into a handful of ranges', () => {
    const cells = [];
    for (let index = 0; index < 300; index += 1) {
      cells.push({
        row: (index % 100) + 1,
        column: indexToColumn(Math.floor(index / 100)),
        value: index,
      });
    }
    const workbook = sheetOfRows(200, 5);
    const parsed = setCellsOperation.schema.parse(args(cells));
    const ranges = setCellsOperation.targetRanges(workbook, parsed);

    expect(ranges.length).toBeLessThanOrEqual(10);
    const written = new Set(cells.map((cell) => `Data!${cell.column}${cell.row}`));
    expect(addressesInRanges(ranges)).toEqual(written);
  });

  it('merges runs that are only horizontally contiguous', () => {
    const cells = [
      { row: 1, column: 'A', value: 1 },
      { row: 2, column: 'A', value: 2 },
      { row: 3, column: 'A', value: 3 },
      { row: 1, column: 'B', value: 4 },
      { row: 2, column: 'B', value: 5 },
      { row: 3, column: 'B', value: 6 },
    ];
    const workbook = sheetOfRows(5, 3);
    const ranges = setCellsOperation.targetRanges(
      workbook,
      setCellsOperation.schema.parse(args(cells)),
    );

    expect(ranges).toEqual([
      {
        sheet: 'Data',
        startRow: 1,
        endRow: 3,
        startColumn: 'A',
        endColumn: 'B',
      },
    ]);
  });

  it('covers exactly the written addresses for a scattered batch', () => {
    const cells = [];
    for (let index = 0; index < 300; index += 1) {
      cells.push({
        row: ((index * 37) % 500) + 1,
        column: indexToColumn((index * 11) % 20),
        value: index,
      });
    }
    const workbook = sheetOfRows(600, 25);
    const ranges = setCellsOperation.targetRanges(
      workbook,
      setCellsOperation.schema.parse(args(cells)),
    );
    const written = new Set(cells.map((cell) => `Data!${cell.column}${cell.row}`));

    // Ranges are a SET of covered cells: every written address is covered, and a
    // rectangle never covers a cell the operation did not write (an under-report
    // or an over-report would weaken the safety invariant either way).
    expect(addressesInRanges(ranges)).toEqual(written);
    expect(ranges.length).toBeLessThanOrEqual(cells.length);
  });
});

describe('maxColumnCount at sheet scale', () => {
  it('counts columns for a 200k-row sheet without overflowing the stack', () => {
    const workbook = sheetOfRows(200_000, 3);

    expect(maxColumnCount(workbook.sheets[0]!.rows)).toBe(3);
    expect(maxColumnCount([])).toBe(0);
  });

  it('counts columns for a ragged 200k-row sheet', () => {
    const workbook = sheetOfRows(200_000, 1);
    workbook.sheets[0]!.rows[199_999] = [createCell('a'), createCell('b'), createCell('c')];

    expect(maxColumnCount(workbook.sheets[0]!.rows)).toBe(3);
  });
});
