import { describe, expect, it } from 'vitest';

import {
  applyOperation,
  applyPatch,
  createCell,
  createOperationRegistry,
  type Workbook,
} from '../src/index.js';

describe('format_cells operation', () => {
  it('applies workbook-native formatting transactionally and restores it through the inverse patch', () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Sheet1',
          rows: [
            [createCell('Name'), createCell('Amount')],
            [createCell('Ada'), createCell(12.5)],
          ],
        },
      ],
    };
    const registry = createOperationRegistry();
    const result = applyOperation(
      workbook,
      'format_cells',
      {
        sheet: 'Sheet1',
        startRow: 1,
        endRow: 2,
        startColumn: 'A',
        endColumn: 'B',
        style: { italic: true, fillColor: '#FFF2CC' },
      },
      { registry },
    );

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.workbook.sheets[0]?.rows[0]?.[0]?.style).toEqual({
      italic: true,
      fillColor: '#FFF2CC',
    });
    expect(result.report.affectedCells).toBe(4);
    expect(applyPatch(result.workbook, result.inverse)).toEqual(workbook);
  });
});
