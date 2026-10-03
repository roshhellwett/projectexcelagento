import { describe, expect, it } from 'vitest';

import {
  createCell,
  invariantFormulaCellsRemainFormulas,
  invariantNoCellsOutsideTargetRange,
  invariantRowCountUnchanged,
  invariantRowsMultisetEqual,
  runInvariants,
  type Workbook,
} from '../src/index.js';

function workbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Data',
        rows: [
          [createCell('ID'), createCell('Value'), createCell('Total')],
          [createCell('A'), createCell(2), createCell(4, { formula: '=B2*2' })],
          [createCell('B'), createCell(1), createCell(2, { formula: '=B3*2' })],
        ],
      },
    ],
  };
}

const dataRange = {
  sheet: 'Data',
  startRow: 2,
  endRow: 3,
  startColumn: 'A',
  endColumn: 'C',
};

describe('shared invariants', () => {
  it('detects row-count changes', () => {
    const before = workbook();
    const after = workbook();
    after.sheets[0]!.rows.pop();

    expect(invariantRowCountUnchanged(before, after)).toHaveLength(1);
  });

  it('checks that a sort preserves the row multiset', () => {
    const before = workbook();
    const after = workbook();
    after.sheets[0]!.rows.splice(1, 2, after.sheets[0]!.rows[2]!, after.sheets[0]!.rows[1]!);

    expect(invariantRowsMultisetEqual(before, after, dataRange)).toEqual([]);
    after.sheets[0]!.rows[1]![0] = createCell('corrupted');
    expect(invariantRowsMultisetEqual(before, after, dataRange)).not.toEqual([]);
  });

  it('requires formulas to remain formulas unless explicitly allowed', () => {
    const before = workbook();
    const after = workbook();
    after.sheets[0]!.rows[1]![2] = createCell(4);

    expect(invariantFormulaCellsRemainFormulas(before, after)).not.toEqual([]);
    expect(invariantFormulaCellsRemainFormulas(before, after, true)).toEqual([]);
  });

  it('rejects changes outside declared ranges', () => {
    const before = workbook();
    const after = workbook();
    after.sheets[0]!.rows[1]![0] = createCell('changed');

    expect(
      invariantNoCellsOutsideTargetRange(before, after, [
        { sheet: 'Data', startRow: 2, endRow: 3, startColumn: 'B', endColumn: 'C' },
      ]),
    ).toHaveLength(1);
  });

  it('combines the reusable checks', () => {
    const before = workbook();
    const after = workbook();
    after.sheets[0]!.rows[1]![1] = createCell(5);

    expect(
      runInvariants(before, after, {
        targetRanges: [dataRange],
        rowCountUnchanged: true,
      }),
    ).toMatchObject({ valid: true, errors: [] });
  });
});
