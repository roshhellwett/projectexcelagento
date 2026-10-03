import { describe, expect, it } from 'vitest';

import { createCell, type Workbook } from '@excel-agent/engine';

import { analyzeSpreadsheetIntentAndData, resolveColumn, getColumnProfiles } from '../src/index.js';

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
