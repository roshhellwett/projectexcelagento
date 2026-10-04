import { describe, expect, it } from 'vitest';

import { createCell, type Sheet, type Workbook } from '@excel-agent/engine';

import {
  analyzeSpreadsheetIntentAndData,
  auditSheet,
  buildSheetContext,
  buildSystemPrompt,
  resolveColumn,
  getColumnProfiles,
} from '../src/index.js';

const SHEET = 'Orders';

function workbook(): Workbook {
  return {
    sheets: [
      {
        name: SHEET,
        rows: [
          [
            createCell('Order ID'),
            createCell('Order Date'),
            createCell('Sort Priority'),
            createCell('Amount'),
          ],
          [createCell('ORD-1001'), createCell('2026-03-01'), createCell(1), createCell(1420.5)],
          [createCell('ORD-1002'), createCell('03/15/2026'), createCell(2), createCell(840)],
        ],
      },
    ],
  };
}

function columns() {
  return getColumnProfiles(workbook().sheets[0]!);
}

function plan(query: string) {
  return analyzeSpreadsheetIntentAndData(query, workbook(), SHEET);
}

describe('column resolution', () => {
  it('resolves a bare letter to that column, never to a substring of a header', () => {
    // Regression: "A" must not match the "a" inside "Amount" / "Date".
    expect(resolveColumn('A', columns())?.letter).toBe('A');
    expect(resolveColumn('C', columns())?.letter).toBe('C');
  });

  it('resolves multi-word header names', () => {
    expect(resolveColumn('sort priority', columns())?.letter).toBe('C');
  });
});

describe('explicit structural commands outrank keyword heuristics', () => {
  it('renames the column instead of treating "Order Date" as a date request', () => {
    const result = plan('rename column B to Invoice Date');
    expect(result.proposedAction?.name).toBe('rename_column');
    expect(result.proposedAction?.args).toEqual({
      sheet: SHEET,
      column: 'B',
      newName: 'Invoice Date',
      headerRow: 1,
    });
  });

  it('keeps the casing the user typed for the new header', () => {
    const result = plan('rename column A to Order Reference');
    expect(result.proposedAction?.args.newName).toBe('Order Reference');
  });

  it('deletes the named column instead of sorting by it', () => {
    const result = plan('delete column Sort Priority');
    expect(result.proposedAction?.name).toBe('delete_column');
    expect(result.proposedAction?.args.column).toBe('C');
  });
});

describe('find & replace', () => {
  it('preserves the casing of the search and replacement text', () => {
    const result = plan('replace Pending with Completed');
    expect(result.proposedAction?.name).toBe('find_replace');
    expect(result.proposedAction?.args).toMatchObject({ find: 'Pending', replace: 'Completed' });
  });

  it('captures the whole replacement phrase, not just its first character', () => {
    // Regression: a lazy quantifier with no end anchor captured only "A".
    const result = plan('find Acme and replace with Acme Corporation');
    expect(result.proposedAction?.args).toMatchObject({
      find: 'Acme',
      replace: 'Acme Corporation',
    });
  });
});

describe('date format selection', () => {
  it('only treats "us"/"eu" as locales when they are standalone words', () => {
    // Regression: "must" contains the substring "us".
    expect(plan('the order date must be YYYY-MM-DD').proposedAction?.args.format).toBe(
      'YYYY-MM-DD',
    );
    expect(plan('convert the order date to US format').proposedAction?.args.format).toBe(
      'MM/DD/YYYY',
    );
    expect(plan('show the order date in EU format').proposedAction?.args.format).toBe('DD/MM/YYYY');
  });
});

describe('informational turns never mutate the workbook', () => {
  it('answers aggregations and overviews without proposing an operation', () => {
    for (const query of [
      'what is the total amount',
      'average of the amount column',
      'who handled the most transactions',
      'find missing values',
      'thanks, talk later',
    ]) {
      expect(plan(query).proposedAction).toBeUndefined();
    }
  });
});

describe('filter to new sheet intent', () => {
  it('identifies IN data in Type column and proposes filter_to_new_sheet', () => {
    const wb: Workbook = {
      sheets: [
        {
          name: 'inventory_transactions_export',
          rows: [
            [
              createCell('Date (IST)'),
              createCell('Qty'),
              createCell('Type'),
              createCell('Handled By'),
            ],
            [
              createCell('2026-09-07'),
              createCell(10),
              createCell('OUT (Removed/Dispatched)'),
              createCell('Ronak'),
            ],
            [
              createCell('2026-09-07'),
              createCell(5),
              createCell('IN (Added to Stock)'),
              createCell('Ronak'),
            ],
            [
              createCell('2026-09-08'),
              createCell(8),
              createCell('OUT (Removed/Dispatched)'),
              createCell('Amit'),
            ],
            [
              createCell('2026-09-08'),
              createCell(12),
              createCell('IN (Added to Stock)'),
              createCell('Amit'),
            ],
          ],
        },
      ],
    };

    const res = analyzeSpreadsheetIntentAndData(
      'filter out the IN data into a separate sheet',
      wb,
      'inventory_transactions_export',
    );

    expect(res.proposedAction).toBeDefined();
    expect(res.proposedAction?.name).toBe('filter_to_new_sheet');
    expect(res.proposedAction?.args).toMatchObject({
      sheet: 'inventory_transactions_export',
      column: 'C',
      operator: 'contains',
      value: 'IN',
      targetSheet: 'IN_Data',
    });
    expect(res.message).toContain('Column C (Type)');
  });

  it('proposes create_sheet for explicit sheet creation request', () => {
    const res = plan('create a new sheet named Summary');
    expect(res.proposedAction?.name).toBe('create_sheet');
    expect(res.proposedAction?.args).toMatchObject({ sheetName: 'Summary' });
  });

  it('proposes add_summary_row for bottom totals request', () => {
    const res = plan('add a total row at the bottom');
    expect(res.proposedAction?.name).toBe('add_summary_row');
    expect(res.proposedAction?.args).toMatchObject({
      sheet: SHEET,
      aggregation: 'sum',
      label: 'Total',
    });
  });
});

describe('sheet-scale column counting', () => {
  /** 200k data rows: past the ~125k limit where spreading an array into Math.max throws. */
  function tallSheet(rowCount: number): Sheet {
    const rows = [[createCell('Amount')]];
    for (let index = 1; index <= rowCount; index += 1) rows.push([createCell(index)]);
    return { name: SHEET, rows };
  }

  it('profiles a 200k-row column without overflowing the stack', () => {
    const sheet = tallSheet(200_000);
    const [profile] = getColumnProfiles(sheet);

    expect(profile?.letter).toBe('A');
    expect(profile?.min).toBe(1);
    expect(profile?.max).toBe(200_000);
    expect(profile?.nonBlankCount).toBe(200_000);
  });

  it('audits and builds the agent context for a 200k-row sheet', () => {
    const sheet = tallSheet(200_000);

    expect(auditSheet(sheet).totalCols).toBe(1);
    expect(JSON.parse(buildSheetContext(sheet)).cols).toBe(1);
    expect(() => buildSystemPrompt(sheet, [])).not.toThrow();
  });
});
