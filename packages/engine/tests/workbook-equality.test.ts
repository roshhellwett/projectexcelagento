import { describe, expect, it } from 'vitest';

import {
  cellEquals,
  cloneWorkbook,
  createCell,
  invariantNoCellsOutsideTargetRange,
  patchBetween,
  workbookEquals,
  type Workbook,
} from '../src/index.js';
import { previewForTransition } from '../src/operation-utils.js';
import { cellValueEquals } from '../src/workbook.js';

const RANGE = {
  sheet: 'Data',
  startRow: 1,
  endRow: 1,
  startColumn: 'A',
  endColumn: 'A',
};

function workbook(...rows: (ReturnType<typeof createCell> | undefined)[][]): Workbook {
  return { sheets: [{ name: 'Data', rows: rows as ReturnType<typeof createCell>[][] }] };
}

describe('cell value equality', () => {
  it('treats NaN as equal to NaN', () => {
    // Regression: `NaN === NaN` is false, so one NaN cell made `workbookEquals`
    // fail and every later operation was rejected by the registry.
    expect(cellValueEquals(Number.NaN, Number.NaN)).toBe(true);
    expect(cellEquals(createCell(Number.NaN), createCell(Number.NaN))).toBe(true);
  });

  it('still distinguishes NaN from a number and a date from a string', () => {
    expect(cellValueEquals(Number.NaN, 0)).toBe(false);
    expect(cellValueEquals(new Date(0), '1970-01-01')).toBe(false);
    expect(cellValueEquals(new Date(0), new Date(0))).toBe(true);
  });

  it('uses one rule for blanks: null and the empty string are the same cell', () => {
    // `cellTypeForValue` types both as 'blank', so keeping them different made a
    // patch that normalized '' to null look like a permanent change.
    expect(cellValueEquals(null, '')).toBe(false);
    expect(cellEquals(createCell(null), createCell(''))).toBe(true);
    expect(cellEquals(createCell(null), createCell(0))).toBe(false);
  });
});

describe('workbook equality with NaN and blanks', () => {
  it('compares a workbook containing NaN and mixed blanks to itself', () => {
    const before = workbook(
      [createCell('Value')],
      [createCell(null)],
      [createCell(Number.NaN), createCell('')],
    );

    expect(workbookEquals(before, cloneWorkbook(before))).toBe(true);
    expect(workbookEquals(before, workbook([createCell('Value')]))).toBe(false);
  });
});

describe('diffs under the blank rule', () => {
  it('produces no patch entry for a null to empty-string normalization', () => {
    const before = workbook([createCell('Value')], [createCell(null), createCell('')]);
    const after = workbook([createCell('Value')], [createCell(''), createCell('')]);

    expect(patchBetween(before, after)).toEqual([]);
    expect(previewForTransition(before, after, [RANGE]).affectedCells).toBe(0);
  });

  it('still produces a patch entry for a real change next to a NaN cell', () => {
    const before = workbook([createCell('Value')], [createCell(Number.NaN)]);
    const after = workbook([createCell('Value')], [createCell(5)]);

    const patch = patchBetween(before, after);
    expect(patch).toHaveLength(1);
    expect(patchBetween(after, cloneWorkbook(after))).toEqual([]);
  });

  it('does not report a blank normalization as a write outside the target range', () => {
    const before = workbook([createCell('Value')], [createCell(null)]);
    const after = workbook([createCell('Value')], [createCell('')]);

    expect(invariantNoCellsOutsideTargetRange(before, after, [RANGE])).toEqual([]);
  });

  it('still reports a real write outside the target range', () => {
    const before = workbook([createCell('Value')], [createCell(null)]);
    const after = workbook([createCell('Value')], [createCell(7)]);

    expect(invariantNoCellsOutsideTargetRange(before, after, [RANGE])).toEqual([
      'Cell Data!A2 changed outside the declared target range.',
    ]);
  });
});
