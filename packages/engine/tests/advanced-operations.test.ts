import { describe, expect, it } from 'vitest';

import {
  addComputedColumnOperation,
  applyOperation,
  createCell,
  createOperationRegistry,
  fillBlanksOperation,
  lookupMergeOperation,
  mergeColumnsOperation,
  splitColumnOperation,
  type Cell,
  type Workbook,
} from '../src/index.js';

function row(...values: Cell['value'][]): Cell[] {
  return values.map((value) => createCell(value));
}

function workbook(rows: Cell[][], sheetName = 'Data'): Workbook {
  return { sheets: [{ name: sheetName, rows }] };
}

function values(result: Workbook, rowNumber: number, sheetName = 'Data'): Cell['value'][] {
  const sheet = result.sheets.find((s) => s.name === sheetName);
  return sheet?.rows[rowNumber - 1]?.map((cell) => cell.value) ?? [];
}

describe('advanced operations', () => {
  describe('fill_blanks', () => {
    it('forward fills empty values from preceding rows', () => {
      const before = workbook([
        row('Category', 'Item'),
        row('Electronics', 'Laptop'),
        row(null, 'Mouse'),
        row('', 'Keyboard'),
        row('Furniture', 'Desk'),
        row(null, 'Chair'),
      ]);

      const args = fillBlanksOperation.schema.parse({
        sheet: 'Data',
        column: 'A',
        strategy: 'forward',
      });

      const result = fillBlanksOperation.apply(before, args);
      expect(values(result.workbook, 3)).toEqual(['Electronics', 'Mouse']);
      expect(values(result.workbook, 4)).toEqual(['Electronics', 'Keyboard']);
      expect(values(result.workbook, 6)).toEqual(['Furniture', 'Chair']);
      expect(fillBlanksOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });

    it('backward fills empty values from following rows', () => {
      const before = workbook([
        row('Item', 'Status'),
        row('A', null),
        row('B', 'Shipped'),
        row('C', null),
        row('D', 'Delivered'),
      ]);

      const args = fillBlanksOperation.schema.parse({
        sheet: 'Data',
        column: 'B',
        strategy: 'backward',
      });

      const result = fillBlanksOperation.apply(before, args);
      expect(values(result.workbook, 2)).toEqual(['A', 'Shipped']);
      expect(values(result.workbook, 4)).toEqual(['C', 'Delivered']);
      expect(fillBlanksOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });

    it('fills numeric column blanks with mean', () => {
      const before = workbook([
        row('ID', 'Score'),
        row(1, 10),
        row(2, null),
        row(3, 30),
      ]);

      const args = fillBlanksOperation.schema.parse({
        sheet: 'Data',
        column: 'B',
        strategy: 'mean',
      });

      const result = fillBlanksOperation.apply(before, args);
      expect(values(result.workbook, 3)).toEqual([2, 20]);
      expect(fillBlanksOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });

    it('fills blanks with static value', () => {
      const before = workbook([
        row('Name', 'Note'),
        row('Alice', 'VIP'),
        row('Bob', null),
      ]);

      const args = fillBlanksOperation.schema.parse({
        sheet: 'Data',
        column: 'B',
        strategy: 'value',
        fillValue: 'Standard',
      });

      const result = fillBlanksOperation.apply(before, args);
      expect(values(result.workbook, 3)).toEqual(['Bob', 'Standard']);
      expect(fillBlanksOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });
  });

  describe('add_computed_column', () => {
    it('computes arithmetic product of two columns', () => {
      const before = workbook([
        row('Item', 'Price', 'Qty'),
        row('Pen', 10, 5),
        row('Notebook', 25, 2),
      ]);

      const args = addComputedColumnOperation.schema.parse({
        sheet: 'Data',
        headerName: 'Total',
        expression: "col('Price') * col('Qty')",
      });

      const result = addComputedColumnOperation.apply(before, args);
      expect(values(result.workbook, 1)).toEqual(['Item', 'Price', 'Qty', 'Total']);
      expect(values(result.workbook, 2)).toEqual(['Pen', 10, 5, 50]);
      expect(values(result.workbook, 3)).toEqual(['Notebook', 25, 2, 50]);
      expect(addComputedColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });

    it('computes string concatenation with uppercase', () => {
      const before = workbook([
        row('First', 'Last'),
        row('john', 'doe'),
        row('jane', 'smith'),
      ]);

      const args = addComputedColumnOperation.schema.parse({
        sheet: 'Data',
        headerName: 'Full Name',
        expression: "upper(col('First')) + ' ' + upper(col('Last'))",
      });

      const result = addComputedColumnOperation.apply(before, args);
      expect(values(result.workbook, 1)).toEqual(['First', 'Last', 'Full Name']);
      expect(values(result.workbook, 2)).toEqual(['john', 'doe', 'JOHN DOE']);
      expect(values(result.workbook, 3)).toEqual(['jane', 'smith', 'JANE SMITH']);
      expect(addComputedColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });
  });

  describe('split_column', () => {
    it('splits full name by space into two new columns', () => {
      const before = workbook([
        row('ID', 'Full Name', 'Department'),
        row(1, 'Alice Smith', 'Sales'),
        row(2, 'Bob Jones', 'Engineering'),
      ]);

      const args = splitColumnOperation.schema.parse({
        sheet: 'Data',
        column: 'B',
        delimiter: ' ',
        newColumnNames: ['First Name', 'Last Name'],
      });

      const result = splitColumnOperation.apply(before, args);
      expect(values(result.workbook, 1)).toEqual(['ID', 'First Name', 'Last Name', 'Department']);
      expect(values(result.workbook, 2)).toEqual([1, 'Alice', 'Smith', 'Sales']);
      expect(values(result.workbook, 3)).toEqual([2, 'Bob', 'Jones', 'Engineering']);
      expect(splitColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });
  });

  describe('merge_columns', () => {
    it('merges City and Country columns with comma separator', () => {
      const before = workbook([
        row('Company', 'City', 'Country'),
        row('Acme', 'Paris', 'France'),
        row('Beta', 'Tokyo', 'Japan'),
      ]);

      const args = mergeColumnsOperation.schema.parse({
        sheet: 'Data',
        columns: ['B', 'C'],
        separator: ', ',
        headerName: 'Location',
      });

      const result = mergeColumnsOperation.apply(before, args);
      expect(values(result.workbook, 1)).toEqual(['Company', 'City', 'Country', 'Location']);
      expect(values(result.workbook, 2)).toEqual(['Acme', 'Paris', 'France', 'Paris, France']);
      expect(values(result.workbook, 3)).toEqual(['Beta', 'Tokyo', 'Japan', 'Tokyo, Japan']);
      expect(mergeColumnsOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });
  });

  describe('lookup_merge', () => {
    it('merges values from lookup sheet based on key match', () => {
      const before: Workbook = {
        sheets: [
          {
            name: 'Orders',
            rows: [
              row('OrderID', 'CustID'),
              row(101, 'C1'),
              row(102, 'C2'),
              row(103, 'C99'), // Unmatched
            ],
          },
          {
            name: 'Customers',
            rows: [
              row('CustID', 'CustName'),
              row('C1', 'Acme Corp'),
              row('C2', 'Globex'),
            ],
          },
        ],
      };

      const args = lookupMergeOperation.schema.parse({
        sheet: 'Orders',
        keyColumn: 'B',
        lookupSheet: 'Customers',
        lookupKeyColumn: 'A',
        lookupValueColumn: 'B',
        headerName: 'Customer Name',
      });

      const result = lookupMergeOperation.apply(before, args);
      expect(values(result.workbook, 1, 'Orders')).toEqual(['OrderID', 'CustID', 'Customer Name']);
      expect(values(result.workbook, 2, 'Orders')).toEqual([101, 'C1', 'Acme Corp']);
      expect(values(result.workbook, 3, 'Orders')).toEqual([102, 'C2', 'Globex']);
      expect(values(result.workbook, 4, 'Orders')).toEqual([103, 'C99', null]);
      expect(lookupMergeOperation.invariants(before, result.workbook, args).valid).toBe(true);
    });
  });

  describe('registry integration', () => {
    it('applies advanced operations through OperationRegistry with undo/redo capability', () => {
      const reg = createOperationRegistry();
      expect(reg.names).toContain('fill_blanks');
      expect(reg.names).toContain('add_computed_column');
      expect(reg.names).toContain('split_column');
      expect(reg.names).toContain('merge_columns');
      expect(reg.names).toContain('lookup_merge');

      const before = workbook([
        row('A', 'B'),
        row(1, null),
        row(2, 10),
      ]);

      const res = applyOperation(
        before,
        'fill_blanks',
        { sheet: 'Data', column: 'B', strategy: 'value', fillValue: 99 },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(values(res.workbook, 2)).toEqual([1, 99]);
        expect(res.report.affectedCells).toBe(1);
      }
    });
  });
});
