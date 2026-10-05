import { describe, expect, it } from 'vitest';
import { z } from 'zod';

import {
  applyOperation,
  applyPatch,
  cloneWorkbook,
  createCell,
  createOperationRegistry,
  createWorkbookValueReader,
  HistoryStack,
  invertPatch,
  OperationRegistry,
  patchBetween,
  workbookEquals,
  type Operation,
  type Workbook,
} from '../src/index.js';

function dataWorkbook(): Workbook {
  return {
    dateSystem: '1904',
    sheets: [
      {
        name: 'Data',
        rows: [
          [createCell('ID'), createCell('Amount')],
          [createCell('b'), createCell(2)],
          [createCell('a'), createCell(1)],
          [createCell('b'), createCell(2)],
        ],
      },
    ],
  };
}

describe('workbook metadata integrity', () => {
  it.each([1, 10])(
    'preserves the 1904 epoch through an edit, undo, redo, and structural snapshot (snapshot every %i)',
    (snapshotEvery) => {
      const before: Workbook = {
        dateSystem: '1904',
        sheets: [
          { name: 'Dates', rows: [[createCell(1), createCell(0, { formula: 'YEAR(A1)' })]] },
        ],
      };
      const history = new HistoryStack(before, { snapshotEvery });
      const result = applyOperation(
        before,
        'set_cells',
        {
          sheet: 'Dates',
          cells: [{ row: 1, column: 'A', value: 366 }],
        },
        { history },
      );
      expect(result.ok).toBe(true);
      expect(result.workbook.dateSystem).toBe('1904');
      expect(createWorkbookValueReader(result.workbook)('Dates', 1, 1)).toBe(1905);
      const undone = history.undo()!;
      expect(undone).toEqual(before);
      expect(createWorkbookValueReader(undone)('Dates', 1, 1)).toBe(1904);
      const redone = history.redo()!;
      expect(redone).toEqual(result.workbook);
      expect(createWorkbookValueReader(redone)('Dates', 1, 1)).toBe(1905);

      const structuralBefore = dataWorkbook();
      const structural = applyOperation(
        structuralBefore,
        'delete_column',
        {
          sheet: 'Data',
          column: 'B',
        },
        { confirmed: true },
      );
      expect(structural.ok).toBe(true);
      if (!structural.ok) throw new Error('Structural operation failed');
      expect(structural.patch[0]?.kind).toBe('workbook');
      expect(structural.workbook.dateSystem).toBe('1904');
      expect(applyPatch(structural.workbook, structural.inverse)).toEqual(structuralBefore);
    },
  );

  it('compares and patches metadata-only changes exactly while treating omitted epoch as 1900', () => {
    const before: Workbook = { sheets: [{ name: 'Data', rows: [[createCell(1)]] }] };
    const after = cloneWorkbook(before);
    after.dateSystem = '1904';
    expect(workbookEquals(before, after)).toBe(false);
    const patch = patchBetween(before, after);
    expect(patch).toHaveLength(1);
    expect(patch[0]?.kind).toBe('workbook');
    expect(applyPatch(before, patch)).toEqual(after);
    expect(applyPatch(after, invertPatch(patch))).toEqual(before);
    const history = new HistoryStack(before);
    history.commit('change_epoch', after, patch);
    expect(history.undo()).toEqual(before);
    expect(history.redo()).toEqual(after);
    const explicitDefault = cloneWorkbook(before);
    explicitDefault.dateSystem = '1900';
    expect(workbookEquals(before, explicitDefault)).toBe(true);
    expect(patchBetween(before, explicitDefault)).toEqual([]);
  });

  it('rejects an operation that silently changes metadata without a corresponding patch', () => {
    const before = dataWorkbook();
    const bad: Operation<Record<string, never>> = {
      name: 'bad_metadata',
      schema: z.object({}),
      targetRanges: () => [],
      validate: () => ({ valid: true, errors: [], warnings: [] }),
      preview: () => ({
        valid: true,
        affectedCells: 0,
        changes: [],
        errors: [],
        warnings: [],
        requiresConfirmation: false,
      }),
      apply: (input) => ({
        workbook: { ...input, dateSystem: '1900' },
        patch: [],
        inverse: [],
        report: { affectedCells: 0, skippedCells: 0, unchangedCells: 0, warnings: [] },
      }),
      invariants: () => ({ valid: true, errors: [] }),
    };
    const result = applyOperation(
      before,
      bad.name,
      {},
      { registry: new OperationRegistry().register(bad) },
    );
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('Metadata corruption was accepted');
    expect(result.error.code).toBe('invariant-error');
    expect(result.error.messages.join(' ')).toMatch(/metadata|date system|patch/i);
    expect(result.workbook).toEqual(before);
  });
});

const structuralCases: [string, Record<string, unknown>][] = [
  ['add_column', { sheet: 'Data', column: 'B', headerName: 'New' }],
  ['delete_column', { sheet: 'Data', column: 'B' }],
  ['sort_range', { sheet: 'Data', column: 'A' }],
  ['filter_rows', { sheet: 'Data', column: 'A', operator: 'equals', value: 'b' }],
  ['delete_duplicates', { sheet: 'Data', columns: ['A'] }],
  [
    'filter_rows',
    { sheet: 'Data', column: 'A', operator: 'equals', value: 'b', targetSheet: 'Filtered' },
  ],
  [
    'filter_to_new_sheet',
    {
      sheet: 'Data',
      targetSheet: 'Filtered',
      column: 'A',
      operator: 'equals',
      value: 'b',
      dropFromSource: true,
    },
  ],
  ['clean_to_new_sheet', { sheet: 'Data', targetSheet: 'Clean' }],
  [
    'add_computed_column',
    { sheet: 'Data', afterColumn: 'A', headerName: 'Computed', expression: 'B*2' },
  ],
  ['split_column', { sheet: 'Data', column: 'A', delimiter: '-' }],
  ['merge_columns', { sheet: 'Data', columns: ['A', 'B'], headerName: 'Merged' }],
  [
    'lookup_merge',
    {
      sheet: 'Data',
      keyColumn: 'A',
      lookupSheet: 'Lookup',
      lookupKeyColumn: 'A',
      lookupValueColumn: 'B',
      headerName: 'Lookup value',
    },
  ],
  [
    'join_sheets',
    {
      sheet: 'Data',
      keyColumn: 'A',
      lookupSheet: 'Lookup',
      lookupKeyColumn: 'A',
      lookupValueColumn: 'B',
      headerName: 'Joined',
      joinType: 'inner',
    },
  ],
  [
    'categorize_column',
    {
      sheet: 'Data',
      sourceColumn: 'B',
      afterColumn: 'A',
      newColumnName: 'Category',
      otherwise: 'Low',
      rules: [{ operator: 'gt', value: 1, label: 'High' }],
    },
  ],
  ['delete_sheet', { sheet: 'Data' }],
];

describe('formula-sensitive structural safety', () => {
  it.each(structuralCases)(
    'refuses %s atomically even when only another sheet references the source',
    (name, args) => {
      const before = dataWorkbook();
      before.sheets.push({
        name: 'Lookup',
        rows: [
          [createCell('ID'), createCell('Value')],
          [createCell('b'), createCell(7)],
        ],
      });
      before.sheets.push({
        name: 'Calculations',
        rows: [[createCell(2, { formula: "'Data'!$B$2" })]],
      });
      const history = new HistoryStack(before);
      const registry = createOperationRegistry();
      const operation = registry.get(name)!;
      const parsed = operation.schema.parse(args);
      const validation = operation.validate(before, parsed);
      expect(validation.valid).toBe(false);
      expect(validation.errors.some((error) => error.code === 'formula-structural-edit')).toBe(
        true,
      );
      expect(operation.preview(before, parsed).valid).toBe(false);
      // Direct operation consumers must not be able to bypass the safety check either.
      expect(() => operation.apply(before, parsed)).toThrow(/formula|reference/i);
      const result = applyOperation(before, name, args, { registry, history, confirmed: true });
      expect(result.ok).toBe(false);
      if (result.ok) throw new Error('Unsafe operation was accepted');
      expect(result.error.code).toBe('validation-error');
      expect(result.error.messages.join(' ')).toMatch(/Excel|LibreOffice|replace.*formulas/i);
      expect(result.workbook).toEqual(before);
      expect(history.length).toBe(0);
    },
  );

  it('also protects formulas moved inside the edited sheet, but allows ordinary value edits', () => {
    const before = dataWorkbook();
    before.sheets[0]!.rows[1]!.push(createCell(4, { formula: '=B2*2' }));
    const sorted = applyOperation(before, 'sort_range', { sheet: 'Data', column: 'A' });
    expect(sorted.ok).toBe(false);
    expect(sorted.workbook).toEqual(before);
    const edited = applyOperation(before, 'set_cells', {
      sheet: 'Data',
      cells: [{ row: 2, column: 'B', value: 3 }],
    });
    expect(edited.ok).toBe(true);
    expect(createWorkbookValueReader(edited.workbook)('Data', 2, 2)).toBe(6);
  });

  it.each(structuralCases)('keeps formula-free %s workflows available', (name, args) => {
    const before = dataWorkbook();
    before.sheets.push({
      name: 'Lookup',
      rows: [
        [createCell('ID'), createCell('Value')],
        [createCell('b'), createCell(7)],
      ],
    });
    const result = applyOperation(before, name, args, { confirmed: true });
    expect(result.ok, result.ok ? undefined : result.error.messages.join(' ')).toBe(true);
    expect(result.workbook.dateSystem).toBe('1904');
  });
});

describe('tall row operations', () => {
  const count = 150_000;
  function tallWorkbook(): Workbook {
    return {
      dateSystem: '1904',
      sheets: [
        {
          name: 'Tall',
          rows: [
            [createCell('ID'), createCell('Payload')],
            ...Array.from({ length: count }, (_, index) => [
              createCell(index),
              createCell(`row-${index}`),
            ]),
          ],
        },
      ],
    };
  }

  it('filters a full 150k-row sheet without truncation or argument-stack overflow', () => {
    const before = tallWorkbook();
    before.sheets[0]!.rows.splice(20, 0, [createCell(-1), createCell('remove')]);
    const result = applyOperation(
      before,
      'filter_rows',
      {
        sheet: 'Tall',
        column: 'A',
        operator: 'gte',
        value: 0,
      },
      { confirmed: true },
    );
    expect(result.ok, result.ok ? undefined : result.error.messages.join(' ')).toBe(true);
    if (!result.ok) throw new Error('Tall filter failed');
    expect(result.report.removedRows).toBe(1);
    expect(result.workbook.sheets[0]!.rows).toHaveLength(count + 1);
    expect(
      result.workbook.sheets[0]!.rows.slice(1).map((row) => row.map((cell) => cell.value)),
    ).toEqual(Array.from({ length: count }, (_, index) => [index, `row-${index}`]));
    expect(applyPatch(result.workbook, result.inverse)).toEqual(before);
  }, 60_000);

  it('extracts a row from a tall sheet without overflowing while replacing the source rows', () => {
    const before = tallWorkbook();
    const result = applyOperation(
      before,
      'filter_to_new_sheet',
      {
        sheet: 'Tall',
        targetSheet: 'Extracted',
        column: 'A',
        operator: 'equals',
        value: count - 1,
        dropFromSource: true,
      },
      { confirmed: true },
    );
    expect(result.ok, result.ok ? undefined : result.error.messages.join(' ')).toBe(true);
    expect(result.workbook.sheets[0]!.rows).toHaveLength(count);
    expect(
      result.workbook.sheets[0]!.rows.slice(1).map((row) => row.map((cell) => cell.value)),
    ).toEqual(Array.from({ length: count - 1 }, (_, index) => [index, `row-${index}`]));
    expect(result.workbook.sheets[1]!.rows.map((row) => row.map((cell) => cell.value))).toEqual([
      ['ID', 'Payload'],
      [count - 1, `row-${count - 1}`],
    ]);
    expect(before.sheets[0]!.rows).toHaveLength(count + 1);
  }, 60_000);

  it.each(['first', 'last'] as const)(
    'deduplicates a full 150k-row sheet, keeping %s occurrences and all payloads',
    (keep) => {
      const before = tallWorkbook();
      before.sheets[0]!.rows.push([createCell(0), createCell('last-zero')]);
      const result = applyOperation(
        before,
        'delete_duplicates',
        {
          sheet: 'Tall',
          columns: ['A'],
          keep,
        },
        { confirmed: true },
      );
      expect(result.ok, result.ok ? undefined : result.error.messages.join(' ')).toBe(true);
      if (!result.ok) throw new Error('Tall dedupe failed');
      expect(result.report.removedRows).toBe(1);
      expect(result.workbook.sheets[0]!.rows).toHaveLength(count + 1);
      const expected = Array.from({ length: count }, (_, index) => [index, `row-${index}`]);
      if (keep === 'last') {
        expected.shift();
        expected.push([0, 'last-zero']);
      }
      expect(
        result.workbook.sheets[0]!.rows.slice(1).map((row) => row.map((cell) => cell.value)),
      ).toEqual(expected);
      expect(before.sheets[0]!.rows).toHaveLength(count + 2);
    },
    60_000,
  );
});
