import { z } from 'zod';
import { describe, expect, it } from 'vitest';

import {
  HistoryStack,
  OperationRegistry,
  applyOperation,
  createCell,
  createOperationRegistry,
  patchForChange,
  type Operation,
  type Workbook,
} from '../src/index.js';

function workbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Data',
        rows: [[createCell('Name')], [createCell('  Ada  ')]],
      },
    ],
  };
}

describe('operation registry', () => {
  it('registers all M0 and M2 operations with their schemas', () => {
    const registry = createOperationRegistry();

    expect(registry.names).toEqual([
      'format_dates',
      'sort_range',
      'filter_rows',
      'find_replace',
      'delete_duplicates',
      'rename_column',
      'delete_column',
      'add_column',
      'normalize_text',
      'set_cells',
    ]);
    expect(registry.schemas.has('format_dates')).toBe(true);
    expect(registry.schemas.has('set_cells')).toBe(true);
  });

  it('runs validation, preview, apply, invariants, and history as one transaction', () => {
    const before = workbook();
    const history = new HistoryStack(before);
    const result = applyOperation(
      before,
      'normalize_text',
      { sheet: 'Data', columns: ['A'], case: 'upper' },
      { history },
    );

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.preview.affectedCells).toBe(1);
      expect(result.workbook.sheets[0]?.rows[1]?.[0]?.value).toBe('ADA');
      expect(result.patch.length).toBe(1);
    }
    expect(history.position).toBe(1);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe('  Ada  ');
  });

  it('returns structured errors for unknown operations and invalid schemas', () => {
    const before = workbook();
    const unknown = applyOperation(before, 'missing', {});
    const invalid = applyOperation(before, 'normalize_text', { sheet: '' });

    expect(unknown).toMatchObject({
      ok: false,
      error: { code: 'unknown-operation', rolledBack: false },
    });
    expect(invalid).toMatchObject({
      ok: false,
      error: { code: 'schema-error', rolledBack: false },
    });
  });

  it('auto-rolls back when a registered operation violates an invariant', () => {
    const before = workbook();
    const badOperation: Operation<{ sheet: string }> = {
      name: 'bad_operation',
      schema: z.object({ sheet: z.string() }),
      targetRanges: () => [
        { sheet: 'Data', startRow: 1, endRow: 1, startColumn: 'A', endColumn: 'A' },
      ],
      validate: () => ({ valid: true, errors: [], warnings: [] }),
      preview: () => ({
        valid: true,
        affectedCells: 1,
        changes: [],
        warnings: [],
        errors: [],
        requiresConfirmation: false,
      }),
      apply: (input) => {
        const after = {
          sheets: input.sheets.map((sheet) => ({
            ...sheet,
            rows: sheet.rows.map((row) => row.map(() => createCell('corrupted'))),
          })),
        };
        const patch = [
          patchForChange(
            { sheet: 'Data', row: 2, column: 'A' },
            input.sheets[0]!.rows[1]![0]!,
            createCell('corrupted'),
          ),
        ];
        return {
          workbook: after,
          report: { affectedCells: 1, skippedCells: 0, unchangedCells: 0, warnings: [] },
          patch,
          inverse: [],
        };
      },
      invariants: () => ({ valid: false, errors: ['intentional invariant failure'] }),
    };
    const registry = new OperationRegistry().register(badOperation);
    const result = applyOperation(before, 'bad_operation', { sheet: 'Data' }, { registry });

    expect(result).toMatchObject({
      ok: false,
      error: { code: 'invariant-error', rolledBack: true },
    });
    expect(result.workbook).toEqual(before);
  });
});
