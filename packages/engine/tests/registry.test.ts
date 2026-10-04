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
      'fill_blanks',
      'add_computed_column',
      'split_column',
      'merge_columns',
      'lookup_merge',
      'clean_to_new_sheet',
      'edit_cells',
      'filter_to_new_sheet',
      'create_sheet',
      'duplicate_sheet',
      'delete_sheet',
      'add_summary_row',
      'aggregate_column',
      'group_and_summarize',
      'join_sheets',
      'fill_series',
      'categorize_column',
    ]);
    expect(registry.schemas.has('format_dates')).toBe(true);
    expect(registry.schemas.has('set_cells')).toBe(true);
    expect(registry.schemas.has('fill_blanks')).toBe(true);
    expect(registry.schemas.has('add_computed_column')).toBe(true);
    expect(registry.schemas.has('filter_to_new_sheet')).toBe(true);
    expect(registry.schemas.has('create_sheet')).toBe(true);
    expect(registry.schemas.has('duplicate_sheet')).toBe(true);
    expect(registry.schemas.has('delete_sheet')).toBe(true);
    expect(registry.schemas.has('add_summary_row')).toBe(true);
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

  it('keeps applying operations to a workbook that contains a NaN cell', () => {
    // Regression: `NaN === NaN` is false, so one NaN cell made the forward/inverse
    // patch verification fail and locked the user out of EVERY later operation.
    const before = {
      sheets: [
        {
          name: 'Data',
          rows: [[createCell('Value')], [createCell(Number.NaN)], [createCell(3)]],
        },
      ],
    };
    const history = new HistoryStack(before);

    for (const value of [4, 5, 6]) {
      const result = applyOperation(
        before,
        'set_cells',
        { sheet: 'Data', cells: [{ row: 3, column: 'A', value }] },
        { history },
      );
      expect(result.ok).toBe(true);
    }

    expect(history.position).toBe(3);
    expect(history.stepBack(history.length).sheets[0]?.rows[2]?.[0]?.value).toBe(6);
    expect(Number.isNaN(history.stepBack(0).sheets[0]?.rows[1]?.[0]?.value as number)).toBe(true);
  });

  it('round-trips a workbook that mixes null and empty-string blanks', () => {
    const before = {
      sheets: [
        {
          name: 'Data',
          rows: [[createCell('Value')], [createCell(null)], [createCell('')]],
        },
      ],
    };
    const history = new HistoryStack(before);

    const result = applyOperation(
      before,
      'set_cells',
      { sheet: 'Data', cells: [{ row: 2, column: 'A', value: '' }] },
      { history },
    );

    expect(result.ok).toBe(true);
    expect(history.position).toBe(1);
    expect(history.undo()).toEqual(before);
    const afterFirst = history.redo() as Workbook;

    const value = applyOperation(
      afterFirst,
      'set_cells',
      { sheet: 'Data', cells: [{ row: 2, column: 'A', value: 9 }] },
      { history },
    );
    expect(value.ok).toBe(true);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBeNull();
  });
});

/**
 * A destructive change must never be applied just because a caller invoked it. The engine is
 * the only layer that cannot be bypassed, so the confirmation flag has to be a guarantee here
 * rather than a hint that the UI is free to ignore.
 */
describe('destructive-change confirmation', () => {
  function duplicates(): Workbook {
    return {
      sheets: [
        {
          name: 'Data',
          rows: [
            [createCell('Name'), createCell('Amount')],
            [createCell('Ada'), createCell(1)],
            [createCell('Grace'), createCell(2)],
            [createCell('Ada'), createCell(1)],
          ],
        },
      ],
    };
  }

  const args = { sheet: 'Data', columns: ['A'], headerRow: 1 };

  it('refuses a row-deleting operation until it is confirmed', () => {
    const before = duplicates();
    const registry = createOperationRegistry();

    // The operation genuinely asks for confirmation.
    const preview = registry.get('delete_duplicates')!.preview(before, args);
    expect(preview.valid).toBe(true);
    expect(preview.requiresConfirmation).toBe(true);

    const result = applyOperation(before, 'delete_duplicates', args, { registry });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.error.code).toBe('confirmation-required');
    expect(result.error.messages.join(' ')).toMatch(/not confirmed/i);
  });

  it('leaves the workbook completely untouched when it refuses', () => {
    const before = duplicates();
    const result = applyOperation(before, 'delete_duplicates', args);
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    // Row count is unchanged: the refusal happened before any mutation.
    expect(result.workbook.sheets[0]?.rows).toHaveLength(4);
    expect(result.workbook).toEqual(before);
  });

  it('records nothing in history when it refuses', () => {
    const before = duplicates();
    const history = new HistoryStack(before);
    const result = applyOperation(before, 'delete_duplicates', args, { history });
    expect(result.ok).toBe(false);
    expect(history.position).toBe(0);
    expect(history.canUndo).toBe(false);
  });

  it('applies once confirmed', () => {
    const result = applyOperation(duplicates(), 'delete_duplicates', args, { confirmed: true });
    expect(result.ok).toBe(true);
    expect(result.workbook.sheets[0]?.rows).toHaveLength(3);
  });

  it('applies immediately when the operation does not need confirmation', () => {
    // A safe operation must not be made to wait for a pointless second click.
    const before = duplicates();
    const result = applyOperation(before, 'rename_column', {
      sheet: 'Data',
      column: 'A',
      newName: 'Customer',
    });
    expect(result.ok).toBe(true);
  });

  /**
   * An inner join and a fill that overwrites values both destroy data, so both have to hit the
   * same wall `delete_duplicates` does. The refusal is checked on the preview and again through
   * the registry, because the flag is only a guarantee if the registry enforces it.
   */
  const destructiveCases = [
    {
      label: 'an inner join that drops rows',
      operation: 'join_sheets',
      input: {
        sheet: 'Orders',
        keyColumn: 'A',
        lookupSheet: 'Customers',
        lookupKeyColumn: 'A',
        lookupValueColumn: 'B',
        headerName: 'Name',
        joinType: 'inner',
      },
      book: (): Workbook => ({
        sheets: [
          {
            name: 'Orders',
            rows: [[createCell('Customer')], [createCell('acme')], [createCell('nobody')]],
          },
          {
            name: 'Customers',
            rows: [
              [createCell('Key'), createCell('Name')],
              [createCell('acme'), createCell('Acme')],
            ],
          },
        ],
      }),
    },
    {
      label: 'a fill that overwrites existing values',
      operation: 'fill_series',
      input: {
        sheet: 'Data',
        column: 'A',
        strategy: 'linear',
        sourceRange: 'A2:A3',
        targetStartRow: 4,
        targetEndRow: 5,
      },
      book: (): Workbook => ({
        sheets: [
          {
            name: 'Data',
            rows: [
              [createCell('Amount'), createCell('Label')],
              [createCell(10), createCell('a')],
              [createCell(20), createCell('b')],
              [createCell(999), createCell('keep me')],
              [createCell(5), createCell('also keep me')],
            ],
          },
        ],
      }),
    },
  ];

  it.each(destructiveCases)('refuses $label until it is confirmed', (testCase) => {
    const before = testCase.book();
    const registry = createOperationRegistry();
    const operation = registry.get(testCase.operation)!;
    const args = operation.schema.parse(testCase.input) as never;

    expect(operation.preview(before, args).requiresConfirmation).toBe(true);

    const history = new HistoryStack(before);
    const refused = applyOperation(before, testCase.operation, testCase.input, {
      registry,
      history,
    });
    expect(refused.ok).toBe(false);
    if (refused.ok) throw new Error('unreachable');
    expect(refused.error.code).toBe('confirmation-required');
    expect(refused.workbook).toEqual(before);
    expect(history.position).toBe(0);

    const confirmed = applyOperation(before, testCase.operation, testCase.input, {
      registry,
      confirmed: true,
    });
    expect(confirmed.ok).toBe(true);
  });
});
