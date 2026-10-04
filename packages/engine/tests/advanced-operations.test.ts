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
  HistoryStack,
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
      const before = workbook([row('ID', 'Score'), row(1, 10), row(2, null), row(3, 30)]);

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
      const before = workbook([row('Name', 'Note'), row('Alice', 'VIP'), row('Bob', null)]);

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
      const before = workbook([row('First', 'Last'), row('john', 'doe'), row('jane', 'smith')]);

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

    it('does not rewrite column letters inside string values', () => {
      const before = workbook([row('Code', 'Note'), row('AB', 'keep'), row('x', 'A')]);

      const args = addComputedColumnOperation.schema.parse({
        sheet: 'Data',
        headerName: 'Mirror',
        expression: "col('Code') + '-' + col('Note')",
      });

      const result = addComputedColumnOperation.apply(before, args);
      expect(values(result.workbook, 2)).toEqual(['AB', 'keep', 'AB-keep']);
      expect(values(result.workbook, 3)).toEqual(['x', 'A', 'x-A']);
    });

    it('treats % as the Excel postfix percent operator', () => {
      const before = workbook([row('Rate'), row(50), row(200)]);

      const args = addComputedColumnOperation.schema.parse({
        sheet: 'Data',
        headerName: 'Half',
        expression: "col('Rate') * 50%",
      });

      const result = addComputedColumnOperation.apply(before, args);
      expect(values(result.workbook, 2)).toEqual([50, 25]);
      expect(values(result.workbook, 3)).toEqual([200, 100]);
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
            rows: [row('CustID', 'CustName'), row('C1', 'Acme Corp'), row('C2', 'Globex')],
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

  describe('edit_cells', () => {
    it('writes a value and undoes it through the registry', () => {
      const before = workbook([row('Name', 'Qty'), row('Widget', 2)]);
      const registry = createOperationRegistry();

      const result = applyOperation(
        before,
        'edit_cells',
        { sheet: 'Data', edits: [{ row: 2, column: 'B', value: 7 }] },
        { registry, history: new HistoryStack(before, { snapshotEvery: 5 }) },
      );

      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(values(result.workbook, 2)).toEqual(['Widget', 7]);
        // A keystroke in a cell must not stop and ask a human to confirm itself.
        expect(result.preview.requiresConfirmation).toBe(false);
        expect(result.report.affectedCells).toBe(1);
      }
    });

    it('grows a short row and the sheet so a paste past the data lands', () => {
      const before = workbook([row('A', 'B'), row(1)]);

      const result = applyOperation(before, 'edit_cells', {
        sheet: 'Data',
        edits: [
          { row: 1, column: 'D', value: 'far right' },
          { row: 3, column: 'A', value: 'new row' },
        ],
      });

      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.workbook.sheets[0]?.rows).toHaveLength(3);
        expect(values(result.workbook, 1)).toEqual(['A', 'B', null, 'far right']);
        expect(values(result.workbook, 3)).toEqual(['new row']);
      }
    });

    it('stores a formula and clears a cell back to blank', () => {
      const before = workbook([row('A'), row(5)]);

      const written = applyOperation(before, 'edit_cells', {
        sheet: 'Data',
        edits: [{ row: 2, column: 'A', formula: '=A1*2' }],
      });
      expect(written.ok).toBe(true);
      if (written.ok) {
        expect(written.workbook.sheets[0]?.rows[1]?.[0]?.formula).toBe('=A1*2');
        expect(written.workbook.sheets[0]?.rows[1]?.[0]?.type).toBe('formula');
      }

      const cleared = applyOperation(written.ok ? written.workbook : before, 'edit_cells', {
        sheet: 'Data',
        edits: [{ row: 2, column: 'A' }],
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) {
        expect(cleared.workbook.sheets[0]?.rows[1]?.[0]?.value).toBe(null);
        expect(cleared.workbook.sheets[0]?.rows[1]?.[0]?.formula).toBeUndefined();
      }
    });

    it('refuses an address past the last column of a spreadsheet', () => {
      const before = workbook([row('A')]);

      const result = applyOperation(before, 'edit_cells', {
        sheet: 'Data',
        edits: [{ row: 1, column: 'XFE', value: 'nope' }],
      });

      expect(result.ok).toBe(false);
      expect(result.error.code).toBe('validation-error');
      expect(result.error.messages.join(' ')).toMatch(/XFD/);
    });
  });

  describe('filter_to_new_sheet', () => {
    it('filters rows matching condition into a new worksheet preserving headers and source data', () => {
      const reg = createOperationRegistry();
      const before: Workbook = {
        sheets: [
          {
            name: 'Transactions',
            rows: [
              row('Date', 'Qty', 'Type', 'Handled By'),
              row('2026-09-07', 10, 'OUT (Removed/Dispatched)', 'Ronak'),
              row('2026-09-07', 5, 'IN (Added to Stock)', 'Ronak'),
              row('2026-09-08', 8, 'OUT (Removed/Dispatched)', 'Amit'),
              row('2026-09-08', 12, 'IN (Added to Stock)', 'Amit'),
            ],
          },
        ],
      };

      const res = applyOperation(
        before,
        'filter_to_new_sheet',
        {
          sheet: 'Transactions',
          targetSheet: 'IN_Data',
          column: 'C',
          operator: 'contains',
          value: 'IN',
          headerRow: 1,
        },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.workbook.sheets.length).toBe(2);
        const inSheet = res.workbook.sheets.find((s) => s.name === 'IN_Data');
        expect(inSheet).toBeDefined();
        expect(inSheet?.rows.length).toBe(3); // header + 2 matching rows
        expect(inSheet?.rows[0]?.map((c) => c.value)).toEqual([
          'Date',
          'Qty',
          'Type',
          'Handled By',
        ]);
        expect(inSheet?.rows[1]?.map((c) => c.value)).toEqual([
          '2026-09-07',
          5,
          'IN (Added to Stock)',
          'Ronak',
        ]);
        expect(inSheet?.rows[2]?.map((c) => c.value)).toEqual([
          '2026-09-08',
          12,
          'IN (Added to Stock)',
          'Amit',
        ]);
        // Source sheet remains intact
        const srcSheet = res.workbook.sheets.find((s) => s.name === 'Transactions');
        expect(srcSheet?.rows.length).toBe(5);
      }
    });

    it('supports filter_rows with targetSheet option', () => {
      const reg = createOperationRegistry();
      const before: Workbook = {
        sheets: [
          {
            name: 'Data',
            rows: [
              row('ID', 'Status'),
              row(1, 'Active'),
              row(2, 'Inactive'),
              row(3, 'Active'),
            ],
          },
        ],
      };

      const res = applyOperation(
        before,
        'filter_rows',
        {
          sheet: 'Data',
          column: 'B',
          operator: 'equals',
          value: 'Active',
          targetSheet: 'Active_Users',
        },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.workbook.sheets.length).toBe(2);
        const activeSheet = res.workbook.sheets.find((s) => s.name === 'Active_Users');
        expect(activeSheet?.rows.length).toBe(3); // Header + 2 active rows
      }
    });
  });

  describe('create_sheet and delete_sheet', () => {
    it('creates a new sheet with optional headers', () => {
      const reg = createOperationRegistry();
      const before = workbook([row('A')]);

      const res = applyOperation(
        before,
        'create_sheet',
        { sheetName: 'Summary', headers: ['ID', 'Total', 'Notes'] },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.workbook.sheets.length).toBe(2);
        const sheet = res.workbook.sheets.find((s) => s.name === 'Summary');
        expect(sheet?.rows[0]?.map((c) => c.value)).toEqual(['ID', 'Total', 'Notes']);
      }
    });

    it('duplicates an existing sheet', () => {
      const reg = createOperationRegistry();
      const before = workbook([row('Col1', 'Col2'), row('V1', 'V2')]);

      const res = applyOperation(
        before,
        'duplicate_sheet',
        { sheet: 'Data', targetSheet: 'Data_Copy' },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.workbook.sheets.length).toBe(2);
        const copy = res.workbook.sheets.find((s) => s.name === 'Data_Copy');
        expect(copy?.rows.length).toBe(2);
      }
    });

    it('deletes a sheet with confirmation', () => {
      const reg = createOperationRegistry();
      const before: Workbook = {
        sheets: [
          { name: 'Sheet1', rows: [row('A')] },
          { name: 'Sheet2', rows: [row('B')] },
        ],
      };

      // Refuses without confirmation
      const unconfirmed = applyOperation(
        before,
        'delete_sheet',
        { sheet: 'Sheet2' },
        { registry: reg },
      );
      expect(unconfirmed.ok).toBe(false);

      // Applies with confirmation
      const confirmed = applyOperation(
        before,
        'delete_sheet',
        { sheet: 'Sheet2' },
        { registry: reg, confirmed: true },
      );
      expect(confirmed.ok).toBe(true);
      if (confirmed.ok) {
        expect(confirmed.workbook.sheets.length).toBe(1);
        expect(confirmed.workbook.sheets[0]?.name).toBe('Sheet1');
      }
    });
  });

  describe('add_summary_row', () => {
    it('appends a Total sum row at the bottom of data', () => {
      const reg = createOperationRegistry();
      const before = workbook([
        row('Item', 'Qty', 'Price'),
        row('Laptop', 2, 1000),
        row('Mouse', 5, 25),
      ]);

      const res = applyOperation(
        before,
        'add_summary_row',
        { sheet: 'Data', aggregation: 'sum', label: 'Total', columns: ['B', 'C'] },
        { registry: reg },
      );

      expect(res.ok).toBe(true);
      if (res.ok) {
        const sheet = res.workbook.sheets[0];
        expect(sheet?.rows.length).toBe(4);
        expect(sheet?.rows[3]?.map((c) => c.value)).toEqual(['Total', 7, 1025]);
      }
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
      expect(reg.names).toContain('filter_to_new_sheet');

      const before = workbook([row('A', 'B'), row(1, null), row(2, 10)]);

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
