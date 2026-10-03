import { describe, expect, it } from 'vitest';

import {
  addColumnOperation,
  createCell,
  deleteColumnOperation,
  deleteDuplicatesOperation,
  filterRowsOperation,
  findReplaceOperation,
  normalizeTextOperation,
  renameColumnOperation,
  setCellsOperation,
  sortRangeOperation,
  type Cell,
  type Workbook,
} from '../src/index.js';

function row(...values: Cell['value'][]): Cell[] {
  return values.map((value) => createCell(value));
}

function workbook(rows: Cell[][]): Workbook {
  return { sheets: [{ name: 'Data', rows }] };
}

function values(result: Workbook, rowNumber: number): Cell['value'][] {
  return result.sheets[0]?.rows[rowNumber - 1]?.map((cell) => cell.value) ?? [];
}

describe('engine operations', () => {
  it('sort_range keeps complete rows together and supports descending order', () => {
    const before = workbook([
      row('ID', 'Name', 'Score', 'Outside'),
      row('A', 'Zed', 2, 'row-A'),
      row('B', 'Ada', 1, 'row-B'),
      row('C', 'Mia', 3, 'row-C'),
    ]);
    const args = sortRangeOperation.schema.parse({
      sheet: 'Data',
      column: 'C',
      direction: 'desc',
      startColumn: 'A',
      endColumn: 'C',
    });
    const result = sortRangeOperation.apply(before, args);

    expect(values(result.workbook, 2)).toEqual(['C', 'Mia', 3, 'row-A']);
    expect(values(result.workbook, 4)).toEqual(['B', 'Ada', 1, 'row-C']);
    expect(sortRangeOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('filter_rows keeps matching rows and leaves the header intact', () => {
    const before = workbook([
      row('ID', 'Status'),
      row('A', 'open'),
      row('B', 'closed'),
      row('C', 'open'),
    ]);
    const args = filterRowsOperation.schema.parse({
      sheet: 'Data',
      column: 'B',
      operator: 'equals',
      value: 'open',
    });
    const result = filterRowsOperation.apply(before, args);

    expect(result.workbook.sheets[0]?.rows).toHaveLength(3);
    expect(values(result.workbook, 1)).toEqual(['ID', 'Status']);
    expect(values(result.workbook, 3)).toEqual(['C', 'open']);
    expect(result.report.removedRows).toBe(1);
    expect(filterRowsOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('find_replace changes text but protects formulas by default', () => {
    const before = workbook([
      row('Description', 'Calculated'),
      row('old value', 10),
      row('keep old value', 20),
    ]);
    const formula = createCell(30, { formula: '=SUM(old_value)' });
    before.sheets[0]!.rows[2]!.push(formula);
    const args = findReplaceOperation.schema.parse({
      sheet: 'Data',
      find: 'old',
      replace: 'new',
    });
    const result = findReplaceOperation.apply(before, args);

    expect(values(result.workbook, 2)[0]).toBe('new value');
    expect(values(result.workbook, 3)[0]).toBe('keep new value');
    expect(result.workbook.sheets[0]?.rows[2]?.[2]?.formula).toBe('=SUM(old_value)');
    expect(findReplaceOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('find_replace can explicitly update formulas without changing their type', () => {
    const before = workbook([row('Value'), row('old')]);
    before.sheets[0]!.rows[1]![0] = createCell(1, { formula: '=old+1' });
    const args = findReplaceOperation.schema.parse({
      sheet: 'Data',
      find: 'old',
      replace: 'new',
      includeFormulas: true,
    });
    const result = findReplaceOperation.apply(before, args);

    expect(result.workbook.sheets[0]?.rows[1]?.[0]?.formula).toBe('=new+1');
    expect(result.workbook.sheets[0]?.rows[1]?.[0]?.type).toBe('formula');
  });

  it('delete_duplicates removes duplicate data rows while keeping the first row', () => {
    const before = workbook([
      row('ID', 'Region', 'Amount'),
      row('A', 'West', 10),
      row('A', 'West', 10),
      row('A', 'East', 12),
    ]);
    const args = deleteDuplicatesOperation.schema.parse({
      sheet: 'Data',
      columns: ['A', 'B'],
    });
    const result = deleteDuplicatesOperation.apply(before, args);

    expect(result.workbook.sheets[0]?.rows).toHaveLength(3);
    expect(values(result.workbook, 2)).toEqual(['A', 'West', 10]);
    expect(result.report.removedRows).toBe(1);
  });

  it('rename_column changes only the requested header cell', () => {
    const before = workbook([row('old_name', 'Other'), row('A', 1)]);
    const args = renameColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'A',
      newName: 'new_name',
    });
    const result = renameColumnOperation.apply(before, args);

    expect(values(result.workbook, 1)).toEqual(['new_name', 'Other']);
    expect(values(result.workbook, 2)).toEqual(['A', 1]);
    expect(renameColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('delete_column removes the column and preserves remaining column order', () => {
    const before = workbook([row('A', 'B', 'C'), row(1, 2, 3)]);
    const args = deleteColumnOperation.schema.parse({ sheet: 'Data', column: 'B' });
    const result = deleteColumnOperation.apply(before, args);

    expect(values(result.workbook, 1)).toEqual(['A', 'C']);
    expect(values(result.workbook, 2)).toEqual([1, 3]);
    expect(result.report.deletedColumns).toBe(1);
    expect(deleteColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('add_column inserts a typed default value before the requested column', () => {
    const before = workbook([row('A', 'B'), row(1, 2), row(3, 4)]);
    const args = addColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'B',
      headerName: 'Inserted',
      defaultValue: 'pending',
    });
    const result = addColumnOperation.apply(before, args);

    expect(values(result.workbook, 1)).toEqual(['A', 'Inserted', 'B']);
    expect(values(result.workbook, 3)).toEqual([3, 'pending', 4]);
    expect(result.report.addedColumns).toBe(1);
  });

  it('normalize_text trims, collapses whitespace, and applies case conversion', () => {
    const before = workbook([row('Name'), row('  ada   lovelace  '), row('LINDA')]);
    const args = normalizeTextOperation.schema.parse({
      sheet: 'Data',
      columns: ['A'],
      case: 'title',
    });
    const result = normalizeTextOperation.apply(before, args);

    expect(values(result.workbook, 2)).toEqual(['Ada Lovelace']);
    expect(values(result.workbook, 3)).toEqual(['Linda']);
    expect(normalizeTextOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('set_cells overwrites values and explicitly permits formulas', () => {
    const before = workbook([row('Amount', 'Total'), row(10, 20)]);
    const args = setCellsOperation.schema.parse({
      sheet: 'Data',
      cells: [
        { row: 2, column: 'A', value: 30, numberFormat: '$0.00' },
        { row: 2, column: 'B', value: 60, formula: '=A2*2' },
      ],
    });
    const result = setCellsOperation.apply(before, args);

    expect(result.workbook.sheets[0]?.rows[1]?.[0]).toMatchObject({
      value: 30,
      type: 'number',
      numberFormat: '$0.00',
    });
    expect(result.workbook.sheets[0]?.rows[1]?.[1]).toMatchObject({
      value: 60,
      formula: '=A2*2',
      type: 'formula',
    });
  });
});
