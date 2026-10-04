import { describe, expect, it } from 'vitest';

import {
  HistoryStack,
  aggregateColumnOperation,
  applyOperation,
  categorizeColumnOperation,
  createCell,
  createOperationRegistry,
  fillSeriesOperation,
  groupAndSummarizeOperation,
  invariantNoCellsOutsideTargetRange,
  joinSheetsOperation,
  type Cell,
  type Workbook,
} from '../src/index.js';

function row(...values: Cell['value'][]): Cell[] {
  return values.map((value) => createCell(value));
}

function sheet(name: string, rows: Cell[][]) {
  return { name, rows };
}

function workbook(...sheets: { name: string; rows: Cell[][] }[]): Workbook {
  return { sheets };
}

function values(result: Workbook, rowNumber: number, sheetName = 'Data'): Cell['value'][] {
  return (
    result.sheets
      .find((s) => s.name === sheetName)
      ?.rows[rowNumber - 1]?.map((cell) => cell.value) ?? []
  );
}

function utc(year: number, month: number, day: number): Date {
  return new Date(Date.UTC(year, month - 1, day));
}

function iso(value: Cell['value']): string {
  return value instanceof Date ? value.toISOString().slice(0, 10) : String(value);
}

/** fill_series always targets column A below the seeds; only the knobs change per test. */
function fillSeriesArgs(options: {
  strategy: 'copy' | 'linear' | 'date' | 'text';
  sourceRange: string;
  start: number;
  end: number;
  column?: string;
  step?: number;
  dateUnit?: 'day' | 'week' | 'month' | 'year';
}) {
  return fillSeriesOperation.schema.parse({
    sheet: 'Data',
    column: options.column ?? 'A',
    strategy: options.strategy,
    sourceRange: options.sourceRange,
    targetStartRow: options.start,
    targetEndRow: options.end,
    ...(options.step === undefined ? {} : { step: options.step }),
    ...(options.dateUnit === undefined ? {} : { dateUnit: options.dateUnit }),
  });
}

/**
 * A sales sheet with the messiness real ones have: repeated categories, a text value where a
 * number belongs, and a blank in the middle of a column.
 */
function salesWorkbook(): Workbook {
  return workbook(
    sheet('Data', [
      row('Region', 'Rep', 'Amount', 'Status'),
      row('North', 'Ada', 120, 'Closed'),
      row('North', 'Bob', 80, 'Open'),
      row('South', 'Cyd', 200, 'Closed'),
      row('South', 'Dee', 50, 'Open'),
      row('East', 'Eli', 0, 'Closed'),
      row('North', 'Fay', null, 'Open'),
    ]),
  );
}

describe('aggregate_column', () => {
  it('sums a whole column and reports the number without touching the workbook', () => {
    const before = salesWorkbook();
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'C',
      aggregation: 'sum',
    });

    const result = aggregateColumnOperation.apply(before, args);

    expect(result.report.aggregate).toBe(450);
    expect(result.workbook).toEqual(before);
    expect(aggregateColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('applies a split operator criteria, a complete expression, and a wildcard', () => {
    const before = salesWorkbook();

    const overHundred = aggregateColumnOperation.apply(
      before,
      aggregateColumnOperation.schema.parse({
        sheet: 'Data',
        column: 'C',
        aggregation: 'sum',
        criteria: '>',
        criteriaValue: '100',
      }),
    );
    const fromTwoHundred = aggregateColumnOperation.apply(
      before,
      aggregateColumnOperation.schema.parse({
        sheet: 'Data',
        column: 'C',
        aggregation: 'sum',
        criteria: '>=200',
      }),
    );
    const northern = aggregateColumnOperation.apply(
      before,
      aggregateColumnOperation.schema.parse({
        sheet: 'Data',
        column: 'C',
        aggregation: 'count',
        criteria: '*orth',
        criteriaColumn: 'A',
      }),
    );

    expect(overHundred.report.aggregate).toBe(320);
    expect(fromTwoHundred.report.aggregate).toBe(200);
    // Three rows are North, but Fay's amount is empty, and count measures values rather than rows.
    expect(northern.report.aggregate).toBe(2);
  });

  it('filters by a second column while aggregating the first', () => {
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'C',
      aggregation: 'average',
      criteria: '<>Closed',
      criteriaColumn: 'D',
    });

    const result = aggregateColumnOperation.apply(salesWorkbook(), args);

    expect(result.report.aggregate).toBe(65);
  });

  it('counts non-numeric cells instead of reading them as zero', () => {
    const before = workbook(
      sheet('Data', [row('Label', 'Amount'), row('a', 10), row('b', 'n/a'), row('c', 20)]),
    );
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'B',
      aggregation: 'sum',
    });

    const result = aggregateColumnOperation.apply(before, args);

    expect(result.report.aggregate).toBe(30);
    expect(result.report.skippedCells).toBe(1);
  });

  it('averages the two middle values for an even median and 0 for an empty one', () => {
    const before = workbook(sheet('Data', [row('V'), row(4), row(1), row(3), row(2)]));
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'A',
      aggregation: 'median',
    });

    const result = aggregateColumnOperation.apply(before, args);
    expect(result.report.aggregate).toBe(2.5);

    const noMatch = aggregateColumnOperation.apply(
      before,
      aggregateColumnOperation.schema.parse({
        sheet: 'Data',
        column: 'A',
        aggregation: 'median',
        criteria: '>1000',
      }),
    );
    expect(noMatch.report.aggregate).toBe(0);
    expect(noMatch.report.warnings).toEqual([]);
  });

  it('warns when the matching rows hold nothing a magnitude can be read from', () => {
    const before = workbook(
      sheet('Data', [row('Label', 'Amount'), row('a', 'n/a'), row('b', 'unknown')]),
    );
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'B',
      aggregation: 'min',
    });

    const result = aggregateColumnOperation.apply(before, args);

    expect(result.report.aggregate).toBe(0);
    expect(result.report.skippedCells).toBe(2);
    expect(result.report.warnings.map((issue) => issue.code)).toEqual(['no-numeric-values']);
  });

  it('counts distinct values ignoring case and padding', () => {
    const before = workbook(
      sheet('Data', [row('Status'), row('Open'), row(' open '), row('Closed'), row(null)]),
    );
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'A',
      aggregation: 'count_distinct',
    });

    const result = aggregateColumnOperation.apply(before, args);

    expect(result.report.aggregate).toBe(2);
    expect(aggregateColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('ignores a formatted-but-empty tail below the data', () => {
    const before = workbook(
      sheet('Data', [row('A'), row('x'), row('y'), row(null), row(''), row(null)]),
    );
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'A',
      aggregation: 'count',
    });

    expect(aggregateColumnOperation.apply(before, args).report.aggregate).toBe(2);
  });

  it('refuses a column that is not on the sheet', () => {
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'Z',
      aggregation: 'sum',
    });

    const validation = aggregateColumnOperation.validate(salesWorkbook(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('missing-column');
  });

  it('refuses criteria that is neither an operator nor a complete expression', () => {
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'C',
      aggregation: 'sum',
      criteria: 'North',
      criteriaValue: '100',
    });

    const validation = aggregateColumnOperation.validate(salesWorkbook(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors[0]?.message).toMatch(/not a comparison operator/i);
  });

  it('runs through the registry, needs no confirmation, and round-trips through undo', () => {
    const before = salesWorkbook();
    const history = new HistoryStack(before);
    const result = applyOperation(
      before,
      'aggregate_column',
      { sheet: 'Data', column: 'C', aggregation: 'sum', criteria: '>=', criteriaValue: '80' },
      { history },
    );

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.report.aggregate).toBe(400);
      expect(result.preview.requiresConfirmation).toBe(false);
      expect(result.preview.affectedCells).toBe(0);
    }
    expect(history.position).toBe(1);
    expect(history.undo()).toEqual(before);
  });

  it('passes the target-range invariant on a realistic sheet', () => {
    const before = salesWorkbook();
    const args = aggregateColumnOperation.schema.parse({
      sheet: 'Data',
      column: 'C',
      aggregation: 'max',
      criteria: '<>',
      criteriaValue: 'Closed',
      criteriaColumn: 'D',
    });

    expect(aggregateColumnOperation.targetRanges(before, args)).toEqual([]);
    expect(
      invariantNoCellsOutsideTargetRange(
        before,
        aggregateColumnOperation.apply(before, args).workbook,
        aggregateColumnOperation.targetRanges(before, args),
      ),
    ).toEqual([]);
  });
});

describe('group_and_summarize', () => {
  it('pivots one group per key, sorted, into a new sheet', () => {
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'C',
      aggregation: 'sum',
      targetSheet: 'By Region',
    });

    const result = groupAndSummarizeOperation.apply(salesWorkbook(), args);

    expect(result.workbook.sheets.map((s) => s.name)).toEqual(['Data', 'By Region']);
    expect(values(result.workbook, 1, 'By Region')).toEqual(['Region', 'sum of Amount']);
    expect(values(result.workbook, 2, 'By Region')).toEqual(['East', 0]);
    expect(values(result.workbook, 3, 'By Region')).toEqual(['North', 200]);
    expect(values(result.workbook, 4, 'By Region')).toEqual(['South', 250]);
    expect(
      groupAndSummarizeOperation.invariants(salesWorkbook(), result.workbook, args).valid,
    ).toBe(true);
  });

  it('groups by more than one column', () => {
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A', 'D'],
      valueColumn: 'C',
      aggregation: 'count',
      targetSheet: 'Matrix',
    });

    const result = groupAndSummarizeOperation.apply(salesWorkbook(), args);

    expect(values(result.workbook, 1, 'Matrix')).toEqual(['Region', 'Status', 'count of Amount']);
    expect(values(result.workbook, 2, 'Matrix')).toEqual(['East', 'Closed', 1]);
    expect(values(result.workbook, 5, 'Matrix')).toEqual(['South', 'Closed', 1]);
    expect(values(result.workbook, 6, 'Matrix')).toEqual(['South', 'Open', 1]);
    expect(result.workbook.sheets[1]?.rows).toHaveLength(6);
  });

  it('reports a group with no numeric value instead of writing a zero', () => {
    const before = workbook(
      sheet('Data', [
        row('Region', 'Amount'),
        row('North', 10),
        row('South', 'n/a'),
        row('West', null),
      ]),
    );
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'B',
      aggregation: 'average',
      targetSheet: 'Averages',
    });

    const result = groupAndSummarizeOperation.apply(before, args);

    expect(values(result.workbook, 2, 'Averages')).toEqual(['North', 10]);
    expect(values(result.workbook, 3, 'Averages')).toEqual(['South', null]);
    expect(values(result.workbook, 4, 'Averages')).toEqual(['West', null]);
    expect(result.report.unmatchedRows).toBe(2);
    expect(result.report.skippedCells).toBe(1);
    expect(result.report.warnings.map((issue) => issue.code)).toContain('groups-without-values');
  });

  it('takes a name collision by suffixing rather than overwriting a sheet', () => {
    const before = workbook(
      sheet('Data', [row('Region', 'Amount'), row('North', 1)]),
      sheet('Summary', [row('do not lose me')]),
    );
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'B',
      aggregation: 'sum',
      targetSheet: 'Summary',
    });

    const result = groupAndSummarizeOperation.apply(before, args);

    expect(result.workbook.sheets.map((s) => s.name)).toEqual(['Data', 'Summary', 'Summary (2)']);
    expect(values(result.workbook, 1, 'Summary')).toEqual(['do not lose me']);
    expect(values(result.workbook, 2, 'Summary (2)')).toEqual(['North', 1]);
  });

  it('refuses to summarize a sheet into itself', () => {
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'C',
      aggregation: 'sum',
      targetSheet: 'data',
    });

    const validation = groupAndSummarizeOperation.validate(salesWorkbook(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('same-sheet');
  });

  it('refuses a value column that is not on the sheet', () => {
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'T',
      aggregation: 'sum',
      targetSheet: 'Pivot',
    });

    const validation = groupAndSummarizeOperation.validate(salesWorkbook(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('missing-column');
  });

  it('round-trips the added sheet through undo', () => {
    const before = salesWorkbook();
    const history = new HistoryStack(before);
    const result = applyOperation(
      before,
      'group_and_summarize',
      {
        sheet: 'Data',
        groupBy: ['A'],
        valueColumn: 'C',
        aggregation: 'average',
        targetSheet: 'Avg',
      },
      { history },
    );

    expect(result.ok).toBe(true);
    expect(history.undo()).toEqual(before);
    expect(history.redo()?.sheets.map((s) => s.name)).toEqual(['Data', 'Avg']);
  });

  it('passes the target-range invariant on a realistic sheet', () => {
    const before = salesWorkbook();
    const args = groupAndSummarizeOperation.schema.parse({
      sheet: 'Data',
      groupBy: ['A', 'D'],
      valueColumn: 'C',
      aggregation: 'max',
      targetSheet: 'Pivot',
    });
    const result = groupAndSummarizeOperation.apply(before, args);

    expect(
      invariantNoCellsOutsideTargetRange(
        before,
        result.workbook,
        groupAndSummarizeOperation.targetRanges(before, args),
      ),
    ).toEqual([]);
    expect(groupAndSummarizeOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });
});

describe('join_sheets', () => {
  function joined(): Workbook {
    return workbook(
      sheet('Orders', [
        row('Order', 'Customer'),
        row('O-1', '  ACME  '),
        row('O-2', 'globex'),
        row('O-3', 'initech'),
        row('O-4', 'nowhere'),
      ]),
      sheet('Customers', [
        row('Key', 'Name', 'Tier'),
        row('acme', 'Acme Corp', 'Gold'),
        row('globex', 'Globex', 'Silver'),
        row('acme', 'Acme Copy', 'Bronze'),
      ]),
    );
  }

  const leftArgs = {
    sheet: 'Orders',
    keyColumn: 'B',
    lookupSheet: 'Customers',
    lookupKeyColumn: 'A',
    lookupValueColumn: ['B', 'C'],
    headerName: ['Customer Name', 'Tier'],
    joinType: 'left',
  } as const;

  it('appends several lookup columns, matching keys case-insensitively after trimming', () => {
    const before = joined();
    const args = joinSheetsOperation.schema.parse(leftArgs);

    const result = joinSheetsOperation.apply(before, args);

    expect(values(result.workbook, 1, 'Orders')).toEqual([
      'Order',
      'Customer',
      'Customer Name',
      'Tier',
    ]);
    expect(values(result.workbook, 2, 'Orders')).toEqual(['O-1', '  ACME  ', 'Acme Corp', 'Gold']);
    expect(values(result.workbook, 3, 'Orders')).toEqual(['O-2', 'globex', 'Globex', 'Silver']);
    expect(result.report.matchedRows).toBe(2);
    expect(result.report.unmatchedRows).toBe(2);
    expect(joinSheetsOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('keeps unmatched rows blank on a left join', () => {
    const args = joinSheetsOperation.schema.parse(leftArgs);

    const result = joinSheetsOperation.apply(joined(), args);

    expect(values(result.workbook, 4, 'Orders')).toEqual(['O-3', 'initech', null, null]);
    expect(values(result.workbook, 5, 'Orders')).toEqual(['O-4', 'nowhere', null, null]);
    expect(result.workbook.sheets[0]?.rows).toHaveLength(5);
    expect(result.report.removedRows).toBeUndefined();
  });

  it('keeps the first of two duplicate lookup keys and warns about it', () => {
    const args = joinSheetsOperation.schema.parse(leftArgs);

    const result = joinSheetsOperation.apply(joined(), args);

    // lookup_merge last-wins here; VLOOKUP and XLOOKUP first-win. The join follows the formulas.
    expect(values(result.workbook, 2, 'Orders')[2]).toBe('Acme Corp');
    expect(result.report.warnings.map((issue) => issue.code)).toEqual(['duplicate-lookup-key']);
    expect(result.report.warnings[0]?.message).toMatch(/first occurrence wins/i);
  });

  it('drops unmatched rows on an inner join and asks for confirmation first', () => {
    const args = joinSheetsOperation.schema.parse({ ...leftArgs, joinType: 'inner' });
    const before = joined();

    const preview = joinSheetsOperation.preview(before, args);
    expect(preview.requiresConfirmation).toBe(true);

    const result = joinSheetsOperation.apply(before, args);
    expect(values(result.workbook, 1, 'Orders')).toEqual([
      'Order',
      'Customer',
      'Customer Name',
      'Tier',
    ]);
    expect(values(result.workbook, 2, 'Orders')).toEqual(['O-1', '  ACME  ', 'Acme Corp', 'Gold']);
    expect(values(result.workbook, 3, 'Orders')).toEqual(['O-2', 'globex', 'Globex', 'Silver']);
    expect(result.workbook.sheets[0]?.rows).toHaveLength(3);
    expect(result.report.removedRows).toBe(2);
    expect(joinSheetsOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('does not ask for confirmation for a left join', () => {
    const args = joinSheetsOperation.schema.parse(leftArgs);

    const result = applyOperation(joined(), 'join_sheets', leftArgs);

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.preview.requiresConfirmation).toBe(false);
    expect(joinSheetsOperation.preview(joined(), args).requiresConfirmation).toBe(false);
  });

  it('refuses to join a sheet to itself', () => {
    const args = joinSheetsOperation.schema.parse({ ...leftArgs, lookupSheet: 'Orders' });

    const validation = joinSheetsOperation.validate(joined(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('same-sheet');
  });

  it('refuses one header name per missing lookup column', () => {
    const args = joinSheetsOperation.schema.parse({ ...leftArgs, headerName: ['Customer Name'] });

    const validation = joinSheetsOperation.validate(joined(), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('header-name-mismatch');
  });

  it('round-trips both join types through undo', () => {
    const before = joined();
    const history = new HistoryStack(before);

    const left = applyOperation(before, 'join_sheets', leftArgs, { history });
    expect(left.ok).toBe(true);
    expect(history.undo()).toEqual(before);

    const inner = applyOperation(
      before,
      'join_sheets',
      { ...leftArgs, joinType: 'inner' },
      { history, confirmed: true },
    );
    expect(inner.ok).toBe(true);
    expect(history.undo()).toEqual(before);
    expect(history.redo()?.sheets[0]?.rows).toHaveLength(3);
  });

  it('passes the target-range invariant on a realistic sheet', () => {
    const before = joined();
    for (const joinType of ['left', 'inner'] as const) {
      const args = joinSheetsOperation.schema.parse({ ...leftArgs, joinType });
      const result = joinSheetsOperation.apply(before, args);
      expect(
        invariantNoCellsOutsideTargetRange(
          before,
          result.workbook,
          joinSheetsOperation.targetRanges(before, args),
        ),
      ).toEqual([]);
    }
  });
});

describe('fill_series', () => {
  function ladder(): Workbook {
    return workbook(
      sheet('Data', [
        row('Week'),
        row('Week1'),
        row('Week2'),
        row(null),
        row(null),
        row(null),
        row(null),
      ]),
    );
  }

  it('increments the trailing number of a text series', () => {
    const args = fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 7 });

    const result = fillSeriesOperation.apply(ladder(), args);

    expect(values(result.workbook, 4)).toEqual(['Week3']);
    expect(values(result.workbook, 7)).toEqual(['Week6']);
    expect(fillSeriesOperation.invariants(ladder(), result.workbook, args).valid).toBe(true);
  });

  it('keeps zero padding and can step backwards', () => {
    const before = workbook(
      sheet('Data', [row('L'), row('W01'), row('W02'), row(null), row(null)]),
    );
    const padded = fillSeriesOperation.apply(
      before,
      fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 5 }),
    );
    expect(values(padded.workbook, 4)).toEqual(['W03']);
    expect(values(padded.workbook, 5)).toEqual(['W04']);

    const back = fillSeriesOperation.apply(
      before,
      fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 5, step: -1 }),
    );
    expect(values(back.workbook, 4)).toEqual(['W01']);
  });

  it('extrapolates a linear step from the source', () => {
    const before = workbook(
      sheet('Data', [row('N'), row(10), row(20), row(null), row(null), row(null)]),
    );
    const args = fillSeriesArgs({
      strategy: 'linear',
      sourceRange: 'A2:A3',
      start: 4,
      end: 6,
    });

    const result = fillSeriesOperation.apply(before, args);

    expect(values(result.workbook, 4)).toEqual([30]);
    expect(values(result.workbook, 5)).toEqual([40]);
    expect(values(result.workbook, 6)).toEqual([50]);
    expect(fillSeriesOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('honours an explicit step over a single seed', () => {
    const before = workbook(sheet('Data', [row('N'), row(5), row(null), row(null)]));
    const args = fillSeriesArgs({
      strategy: 'linear',
      sourceRange: 'A2',
      start: 3,
      end: 4,
      step: 2.5,
    });

    const result = fillSeriesOperation.apply(before, args);

    expect(values(result.workbook, 3)).toEqual([7.5]);
    expect(values(result.workbook, 4)).toEqual([10]);
  });

  it('repeats the interval between two source dates', () => {
    const before = workbook(
      sheet('Data', [
        row('When'),
        row(utc(2024, 1, 1)),
        row(utc(2024, 1, 8)),
        row(null),
        row(null),
        row(null),
      ]),
    );
    const weekly = fillSeriesOperation.apply(
      before,
      fillSeriesArgs({ strategy: 'date', sourceRange: 'A2:A3', start: 4, end: 6 }),
    );
    expect(values(weekly.workbook, 4).map(iso)).toEqual(['2024-01-15']);
    expect(values(weekly.workbook, 6).map(iso)).toEqual(['2024-01-29']);

    const monthly = workbook(
      sheet('Data', [row('When'), row(utc(2024, 1, 31)), row(utc(2024, 2, 29)), row(null)]),
    );
    const byMonth = fillSeriesOperation.apply(
      monthly,
      fillSeriesArgs({
        strategy: 'date',
        sourceRange: 'A2:A3',
        start: 4,
        end: 4,
        dateUnit: 'month',
      }),
    );
    expect(values(byMonth.workbook, 4).map(iso)).toEqual(['2024-03-29']);
  });

  it('cycles the source values when copying', () => {
    const args = fillSeriesArgs({ strategy: 'copy', sourceRange: 'A2:A3', start: 4, end: 7 });

    const result = fillSeriesOperation.apply(ladder(), args);

    expect(values(result.workbook, 4)).toEqual(['Week1']);
    expect(values(result.workbook, 5)).toEqual(['Week2']);
    expect(values(result.workbook, 6)).toEqual(['Week1']);
    expect(values(result.workbook, 7)).toEqual(['Week2']);
  });

  it('refuses a linear series it cannot infer a step for', () => {
    const before = workbook(sheet('Data', [row('N'), row(5), row(null), row(null)]));
    const args = fillSeriesArgs({
      strategy: 'linear',
      sourceRange: 'A2',
      start: 3,
      end: 4,
    });

    const validation = fillSeriesOperation.validate(before, args);

    expect(validation.valid).toBe(false);
    expect(validation.errors[0]?.message).toMatch(/two source values/i);
  });

  it('refuses a text source with no trailing number', () => {
    const before = workbook(
      sheet('Data', [row('L'), row('alpha'), row('beta'), row(null), row(null)]),
    );
    const args = fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 5 });

    const validation = fillSeriesOperation.validate(before, args);

    expect(validation.valid).toBe(false);
    expect(validation.errors[0]?.message).toMatch(/trailing number/i);
  });

  it('refuses a date source whose rows are not a whole number of weeks apart', () => {
    const before = workbook(
      sheet('Data', [row('D'), row(utc(2024, 1, 1)), row(utc(2024, 1, 4)), row(null)]),
    );
    const args = fillSeriesArgs({
      strategy: 'date',
      sourceRange: 'A2:A3',
      start: 4,
      end: 4,
      dateUnit: 'week',
    });

    const validation = fillSeriesOperation.validate(before, args);

    expect(validation.valid).toBe(false);
    expect(validation.errors[0]?.message).toMatch(/whole number of weeks/i);
  });

  it('refuses to fill past the last row, over its own seed, or in another column', () => {
    const before = ladder();

    expect(
      fillSeriesOperation
        .validate(
          before,
          fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 40 }),
        )
        .errors.map((error) => error.code),
    ).toContain('target-outside-sheet');
    expect(
      fillSeriesOperation
        .validate(
          before,
          fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 3, end: 4 }),
        )
        .errors.map((error) => error.code),
    ).toContain('overlapping-target');
    expect(
      fillSeriesOperation
        .validate(
          before,
          fillSeriesArgs({ strategy: 'text', sourceRange: 'B2:B3', start: 4, end: 5 }),
        )
        .errors.map((error) => error.code),
    ).toContain('source-column-mismatch');
  });

  it('refuses to overwrite a formula in the target range', () => {
    const rows = [row('L'), row('W01'), row('W02'), [createCell(null, { formula: '=1+1' })]];
    const args = fillSeriesArgs({ strategy: 'text', sourceRange: 'A2:A3', start: 4, end: 4 });

    const validation = fillSeriesOperation.validate(workbook(sheet('Data', rows)), args);

    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('formula-in-target');
  });

  it('round-trips through undo', () => {
    const before = ladder();
    const history = new HistoryStack(before);
    const input = {
      sheet: 'Data',
      column: 'A',
      strategy: 'text',
      sourceRange: 'A2:A3',
      targetStartRow: 4,
      targetEndRow: 7,
    };

    const result = applyOperation(before, 'fill_series', input, { history });

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.preview.requiresConfirmation).toBe(false);
    expect(history.undo()).toEqual(before);
  });

  it('passes the target-range invariant on a realistic sheet', () => {
    const before = salesWorkbook();
    const args = fillSeriesArgs({
      strategy: 'linear',
      sourceRange: 'C2:C3',
      column: 'C',
      start: 7,
      end: 8,
    });

    const result = fillSeriesOperation.apply(before, args);

    expect(
      invariantNoCellsOutsideTargetRange(
        before,
        result.workbook,
        fillSeriesOperation.targetRanges(before, args),
      ),
    ).toEqual([]);
    expect(fillSeriesOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });
});

describe('categorize_column', () => {
  function banded(): Workbook {
    return workbook(
      sheet('Data', [
        row('Item', 'Amount', 'Owner'),
        row('Widget', 1500, 'Ada'),
        row('Gadget', 500, 'Bob'),
        row('Doohickey', 100, 'Cyd'),
        row('Thing', 50, 'Dee'),
        row('Gizmo', 'unknown', 'Eli'),
      ]),
    );
  }

  const rules = {
    sheet: 'Data',
    sourceColumn: 'B',
    newColumnName: 'Band',
    rules: [
      { operator: 'gte', value: 1000, label: 'Large' },
      { operator: 'between', value: [100, 999], label: 'Medium' },
      { operator: 'lt', value: 100, label: 'Small' },
    ],
    otherwise: 'Unclassified',
  } as const;

  it('labels each row with the first matching rule', () => {
    const args = categorizeColumnOperation.schema.parse(rules);

    const result = categorizeColumnOperation.apply(banded(), args);

    expect(values(result.workbook, 1)).toEqual(['Item', 'Amount', 'Owner', 'Band']);
    expect(values(result.workbook, 2)[3]).toBe('Large');
    expect(values(result.workbook, 3)[3]).toBe('Medium');
    expect(values(result.workbook, 4)[3]).toBe('Medium');
    expect(values(result.workbook, 5)[3]).toBe('Small');
    expect(categorizeColumnOperation.invariants(banded(), result.workbook, args).valid).toBe(true);
  });

  it('falls back to the otherwise label', () => {
    const args = categorizeColumnOperation.schema.parse(rules);

    const result = categorizeColumnOperation.apply(banded(), args);

    expect(values(result.workbook, 6)[3]).toBe('Unclassified');
  });

  it('inserts after a column and leaves the rows above the header blank', () => {
    const before = workbook(
      sheet('Data', [
        row('2024 sales'),
        row('Item', 'Amount', 'Owner'),
        row('Widget', 1500, 'Ada'),
      ]),
    );
    const args = categorizeColumnOperation.schema.parse({
      ...rules,
      afterColumn: 'A',
      headerRow: 2,
    });

    const result = categorizeColumnOperation.apply(before, args);

    expect(values(result.workbook, 1)).toEqual(['2024 sales', null]);
    expect(values(result.workbook, 2)).toEqual(['Item', 'Band', 'Amount', 'Owner']);
    expect(values(result.workbook, 3)[1]).toBe('Large');
    expect(categorizeColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });

  it('refuses a duplicate column name unless afterColumn places it explicitly', () => {
    const before = workbook(sheet('Data', [row('Band', 'Amount'), row('existing', 10)]));
    const args = categorizeColumnOperation.schema.parse({
      sheet: 'Data',
      sourceColumn: 'B',
      newColumnName: 'band',
      rules: [{ operator: 'gte', value: 1, label: 'Big' }],
      otherwise: 'Small',
    });

    const validation = categorizeColumnOperation.validate(before, args);
    expect(validation.valid).toBe(false);
    expect(validation.errors.map((error) => error.code)).toContain('duplicate-column-name');

    const placed = categorizeColumnOperation.schema.parse({
      ...categorizeColumnOperation.schema.parse({
        sheet: 'Data',
        sourceColumn: 'B',
        newColumnName: 'band',
        rules: [{ operator: 'gte', value: 1, label: 'Big' }],
        otherwise: 'Small',
        afterColumn: 'A',
      }),
      afterColumn: 'A',
    });
    expect(categorizeColumnOperation.validate(before, placed).valid).toBe(true);
    expect(values(categorizeColumnOperation.apply(before, placed).workbook, 1)).toEqual([
      'Band',
      'band',
      'Amount',
    ]);
  });

  it('refuses a between rule that was not given a pair, at the schema level', () => {
    const parsed = categorizeColumnOperation.schema.safeParse({
      ...rules,
      rules: [{ operator: 'between', value: 5, label: 'Only one bound' }],
    });

    expect(parsed.success).toBe(false);
  });

  it('refuses between bounds that are not numbers, also at the schema level', () => {
    const parsed = categorizeColumnOperation.schema.safeParse({
      ...rules,
      rules: [{ operator: 'between', value: ['a', 'z'], label: 'Letters' }],
    });

    expect(parsed.success).toBe(false);
  });

  it('needs no confirmation and round-trips through undo', () => {
    const before = banded();
    const history = new HistoryStack(before);

    const result = applyOperation(before, 'categorize_column', rules, { history });

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.preview.requiresConfirmation).toBe(false);
    expect(history.undo()).toEqual(before);
    expect(history.redo()?.sheets[0]?.rows[1]?.[3]?.value).toBe('Large');
  });

  it('passes the target-range invariant on a realistic sheet', () => {
    const before = banded();
    const args = categorizeColumnOperation.schema.parse({ ...rules, afterColumn: 'B' });
    const result = categorizeColumnOperation.apply(before, args);

    expect(
      invariantNoCellsOutsideTargetRange(
        before,
        result.workbook,
        categorizeColumnOperation.targetRanges(before, args),
      ),
    ).toEqual([]);
    expect(categorizeColumnOperation.invariants(before, result.workbook, args).valid).toBe(true);
  });
});

/**
 * A destructive change must never be applied just because a caller invoked it, so the engine
 * refuses rather than trusting the caller to have asked the user first.
 */
describe('business operations that destroy data', () => {
  function withOverwrittenTarget(): Workbook {
    return workbook(
      sheet('Data', [
        row('Amount', 'Label'),
        row(10, 'a'),
        row(20, 'b'),
        row(999, 'keep me'),
        row(5, 'also keep me'),
      ]),
    );
  }

  it('refuses an inner join that drops rows until it is confirmed', () => {
    const before = workbook(
      sheet('Orders', [row('Customer'), row('acme'), row('unknown')]),
      sheet('Customers', [row('Key', 'Name'), row('acme', 'Acme Corp')]),
    );
    const input = {
      sheet: 'Orders',
      keyColumn: 'A',
      lookupSheet: 'Customers',
      lookupKeyColumn: 'A',
      lookupValueColumn: 'B',
      headerName: 'Name',
      joinType: 'inner',
    };
    const args = joinSheetsOperation.schema.parse(input);
    const registry = createOperationRegistry();

    const preview = registry.get('join_sheets')!.preview(before, args);
    expect(preview.valid).toBe(true);
    expect(preview.requiresConfirmation).toBe(true);

    const result = applyOperation(before, 'join_sheets', input, { registry });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.error.code).toBe('confirmation-required');
    expect(result.workbook).toEqual(before);

    const confirmed = applyOperation(before, 'join_sheets', input, { registry, confirmed: true });
    expect(confirmed.ok).toBe(true);
    if (confirmed.ok) expect(confirmed.workbook.sheets[0]?.rows).toHaveLength(2);
  });

  it('refuses a fill that would overwrite existing values until it is confirmed', () => {
    const before = withOverwrittenTarget();
    const input = {
      sheet: 'Data',
      column: 'A',
      strategy: 'linear',
      sourceRange: 'A2:A3',
      targetStartRow: 4,
      targetEndRow: 5,
    };
    const registry = createOperationRegistry();

    expect(
      registry.get('fill_series')!.preview(before, fillSeriesOperation.schema.parse(input))
        .requiresConfirmation,
    ).toBe(true);

    const result = applyOperation(before, 'fill_series', input, { registry });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.error.code).toBe('confirmation-required');
    expect(result.workbook).toEqual(before);

    const confirmed = applyOperation(before, 'fill_series', input, { registry, confirmed: true });
    expect(confirmed.ok).toBe(true);
    if (confirmed.ok) expect(values(confirmed.workbook, 5)[0]).toBe(40);
  });

  it('records nothing in history when it refuses either one', () => {
    const before = withOverwrittenTarget();
    const history = new HistoryStack(before);

    applyOperation(
      before,
      'fill_series',
      {
        sheet: 'Data',
        column: 'A',
        strategy: 'linear',
        sourceRange: 'A2:A3',
        targetStartRow: 4,
        targetEndRow: 5,
      },
      { history },
    );

    expect(history.position).toBe(0);
    expect(history.canUndo).toBe(false);
  });
});

describe('tool catalog coverage', () => {
  it('registers all five business operations in the engine', () => {
    const names = createOperationRegistry().names;

    for (const name of [
      'aggregate_column',
      'group_and_summarize',
      'join_sheets',
      'fill_series',
      'categorize_column',
    ]) {
      expect(names).toContain(name);
    }
  });
});
