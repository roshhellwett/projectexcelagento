import { describe, expect, it, vi } from 'vitest';
import { createCell, type Workbook } from '@excel-agent/engine';
import {
  BRIEFING_LIMITS,
  compareWorkbookBaseline,
  createAnalystBriefing,
} from '../src/briefing.js';

function workbook(values: (string | number | boolean | Date | null)[][]): Workbook {
  return {
    sheets: [{ name: 'Data', rows: values.map((row) => row.map((value) => createCell(value))) }],
  };
}

describe('deterministic analyst briefing', () => {
  it('reports full negative/blank distributions, strict numeric text, and rectangular missing cells', () => {
    const wb = workbook([
      ['Amount', 'Segment'],
      [-10, 'A'],
      [-5, 'B'],
      [0, 'A'],
      [null, ''],
      ['  ', 'B'],
      ['-2', true],
      ['12 USD', 'A'],
      [],
    ]);
    const before = structuredClone(wb);
    const report = createAnalystBriefing(wb, ' data ');
    expect(report.status).toBe('ready');
    expect(report.columns[0]?.numeric).toMatchObject({
      totalCount: 8,
      numericCount: 4,
      missingCount: 3,
      nonNumericCount: 1,
      sum: -17,
      mean: -4.25,
      median: -3.5,
      min: -10,
      max: 0,
    });
    expect(report.columns[0]?.numericTextCount).toBe(1);
    expect(report.columns[0]?.sourceRange).toMatchObject({
      a1: "'Data'!A2:A9",
      rowCount: 8,
      cellCount: 8,
    });
    expect(report.dataRange?.a1).toBe("'Data'!A2:B9");
    expect(report.sourceRange?.a1).toBe("'Data'!A1:B9");
    expect(report.quality).toMatchObject({
      totalCells: 16,
      missingCount: 5,
      numericCount: 4,
      nonNumericCount: 7,
      raggedRowCount: 1,
      fullyBlankRowCount: 2,
    });
    expect(
      report.quality.missingCount + report.quality.numericCount + report.quality.nonNumericCount,
    ).toBe(report.quality.totalCells);
    const histogram = report.charts.find((chart) => chart.kind === 'histogram');
    expect(histogram?.series.reduce((total, bin) => total + bin.count, 0)).toBe(4);
    expect(wb).toEqual(before);
    expect(createAnalystBriefing(wb, 'Data')).toEqual(report);
  });

  it('reads supported formulas live, excluding errors, unknown functions and missing formula source instead of caches', () => {
    const wb = workbook([
      ['Input', 'Result'],
      [2, 999],
      [4, 999],
      [6, 999],
      [8, 999],
      [10, 999],
    ]);
    const rows = wb.sheets[0]!.rows;
    rows[1]![1] = createCell(999, { formula: '=A2*3' });
    rows[2]![1] = createCell(999, { formula: '=1/0' });
    rows[3]![1] = createCell(999, { formula: '=NOT_SUPPORTED(A4)' });
    rows[4]![1] = createCell(999, { type: 'formula' });
    rows[5]![1] = createCell(999, { formula: '=IFERROR(B4,123)' });
    const report = createAnalystBriefing(wb, 'Data');
    expect(report.columns[1]).toMatchObject({
      formulaCount: 5,
      formulaErrorCount: 1,
      unsupportedFormulaCount: 3,
      numeric: { totalCount: 5, numericCount: 1, nonNumericCount: 4, missingCount: 0, sum: 6 },
    });
    expect(
      report.quality.examples.some(
        (example) => example.address === "'Data'!B3:B3" && example.reason.includes('#DIV/0!'),
      ),
    ).toBe(true);
    expect(report.columns[1]?.categories.totalCount).toBe(0);
  });

  it('uses the workbook date epoch for live serial formulas and excludes circular results', () => {
    const wb = workbook([
      ['Year', 'Circular'],
      [999, 999],
    ]);
    wb.sheets[0]!.rows[1]![0] = createCell(999, { formula: '=YEAR(1)' });
    wb.sheets[0]!.rows[1]![1] = createCell(999, { formula: '=B2+1' });
    expect(
      createAnalystBriefing({ ...wb, dateSystem: '1904' }, 'Data').columns[0]?.numeric.sum,
    ).toBe(1904);
    expect(
      createAnalystBriefing({ ...wb, dateSystem: '1900' }, 'Data').columns[0]?.numeric.sum,
    ).toBe(1900);
    expect(createAnalystBriefing(wb, 'Data').columns[1]).toMatchObject({
      formulaErrorCount: 1,
      numeric: { numericCount: 0, sum: null, nonNumericCount: 1 },
    });
  });

  it('omits an unrepresentable histogram instead of emitting non-finite chart coordinates', () => {
    const report = createAnalystBriefing(workbook([['Extreme'], [-1e308], [1e308]]), 'Data');
    expect(report.columns[0]?.numeric.numericCount).toBe(2);
    expect(report.charts).toEqual([]);
    expect(report.findings.join(' ')).toContain('histogram omitted');
    expect(report.columns[0]?.numeric.warnings.length).toBeGreaterThan(0);
  });

  it('excludes clock/random formulas and transitive cross-sheet/range dependents deterministically', () => {
    const wb = workbook([
      ['Volatile', 'Derived', 'Static'],
      [111, 222, 333],
      [444, 555, 666],
    ]);
    wb.sheets[0]!.rows[1]![0] = createCell(111, { formula: '=NOW()' });
    wb.sheets[0]!.rows[1]![1] = createCell(222, { formula: '=IFERROR(A2,99)' });
    wb.sheets[0]!.rows[1]![2] = createCell(333, { formula: '=2+2' });
    wb.sheets[0]!.rows[2]![0] = createCell(444, { formula: "=SUM('Other Sheet'!A2:A3)" });
    wb.sheets[0]!.rows[2]![1] = createCell(555, { formula: '=SUM(A2:A3)' });
    wb.sheets[0]!.rows[2]![2] = createCell(666, { formula: '=LEN("NOW()")' });
    wb.sheets.push({
      name: 'Other Sheet',
      rows: [[createCell('Clock')], [createCell(777, { formula: '=TODAY()' })], [createCell(0)]],
    });
    vi.useFakeTimers();
    try {
      vi.setSystemTime(new Date('2024-01-01T00:00:00Z'));
      const report = createAnalystBriefing(wb, 'Data');
      vi.setSystemTime(new Date('2030-06-01T00:00:00Z'));
      expect(createAnalystBriefing(wb, 'Data')).toEqual(report);
      expect(report.quality.volatileFormulaCount).toBe(4);
      expect(report.columns[2]?.numeric.sum).toBe(9);
    } finally {
      vi.useRealTimers();
    }
  });

  it('reports constant distributions and no-numeric columns without inventing zero totals', () => {
    const report = createAnalystBriefing(
      workbook([
        ['Constant', 'Words'],
        [5, 'alpha'],
        [5, 'beta'],
        [5, null],
      ]),
      'Data',
    );
    expect(report.columns[0]?.numeric).toMatchObject({
      sum: 15,
      sampleStandardDeviation: 0,
      iqr: 0,
      outlierCount: 0,
    });
    expect(report.charts[0]?.series).toEqual([
      { label: '5', count: 3, lowerBound: 5, upperBound: 5 },
    ]);
    expect(report.columns[1]?.numeric).toMatchObject({
      sum: null,
      mean: null,
      median: null,
      numericCount: 0,
    });
    expect(report.columns[1]?.numeric.warnings[0]).toContain('No numeric observations');
    const nonnumeric = createAnalystBriefing(
      workbook([['Text'], ['alpha'], [false], [new Date('2024-01-01T00:00:00Z')]]),
      'Data',
    );
    expect(nonnumeric.quality.numericCount).toBe(0);
    expect(nonnumeric.findings.join(' ')).toContain('undefined, not zero');
  });

  it('counts all categories but bounds output and charts, with stable tie order and honest omitted counts', () => {
    const rows = Array.from({ length: 20 }, (_, index) => [
      `Group ${String(index).padStart(2, '0')}`,
      index,
      index,
      index,
      index,
      index,
    ]);
    const report = createAnalystBriefing(
      workbook([['Category', 'B', 'C', 'D', 'E', 'F'], ...rows, ['Group 19', 0, 0, 0, 0, 0]]),
      'Data',
    );
    const categories = report.columns[0]!.categories;
    expect(categories).toMatchObject({
      totalCount: 21,
      distinctCount: 20,
      omittedCount: 12,
      omittedCategoryCount: 12,
    });
    expect(categories.top).toHaveLength(BRIEFING_LIMITS.categories);
    expect(categories.top[0]).toMatchObject({ label: 'Group 19', count: 2 });
    expect(categories.top[1]?.label).toBe('Group 00');
    expect(report.charts).toHaveLength(BRIEFING_LIMITS.charts);
    expect(report.omittedChartCount).toBe(2);
    expect(report.charts[0]).toMatchObject({
      kind: 'categorical',
      includedCount: 9,
      omittedCount: 12,
      sourceRange: { a1: "'Data'!A2:A22" },
    });
    expect(report.charts[0]?.series).toHaveLength(8);
  });

  it('does not merge categories of different types or swallow legitimate hash text', () => {
    const report = createAnalystBriefing(
      workbook([['Label'], [true], ['true'], ['#SKU'], ['#DIV/0!']]),
      'Data',
    );
    expect(report.columns[0]?.categories.top).toEqual([
      { label: '#SKU', valueType: 'string', count: 1 },
      { label: 'true', valueType: 'boolean', count: 1 },
      { label: 'true', valueType: 'string', count: 1 },
    ]);
    expect(report.quality.formulaErrorCount).toBe(1);
  });

  it('handles missing sheets, empty sheets and header-only data with exact ranges', () => {
    expect(createAnalystBriefing({ sheets: [] }, 'Missing')).toMatchObject({
      status: 'sheet-not-found',
      sourceRange: null,
      dataRange: null,
    });
    expect(createAnalystBriefing(workbook([]), 'Data')).toMatchObject({
      status: 'empty',
      columns: [],
      sourceRange: null,
    });
    const headerOnly = createAnalystBriefing(workbook([['Header']]), 'Data');
    expect(headerOnly).toMatchObject({ status: 'empty', dataRange: null, dataRowCount: 0 });
    expect(headerOnly.columns[0]?.numeric.sum).toBeNull();
    const quoted = workbook([['Header'], [1]]);
    quoted.sheets[0]!.name = "O'Brien";
    expect(createAnalystBriefing(quoted, "O'Brien").sourceRange?.a1).toBe("'O''Brien'!A1:A2");
  });

  it('bounds quality examples but audits every row and reports repeated/blank headers', () => {
    const report = createAnalystBriefing(
      workbook([['Same', 'Same', null], ...Array.from({ length: 50 }, () => [null, null, null])]),
      'Data',
    );
    expect(report.quality).toMatchObject({
      missingCount: 150,
      fullyBlankRowCount: 50,
      blankHeaderCount: 1,
      duplicateHeaders: ['Same'],
      omittedExampleCount: 138,
    });
    expect(report.quality.examples).toHaveLength(BRIEFING_LIMITS.qualityExamples);
    expect(report.columns.every((column) => column.numeric.totalCount === 50)).toBe(true);
  });
});

describe('workbook baseline comparison', () => {
  it('compares sheet presence, stored rows/columns and all added/removed cell counts', () => {
    const baseline = workbook([['A'], [1]]);
    baseline.sheets.push({ name: 'Removed', rows: [[createCell('X')], [createCell(2)]] });
    const current = workbook([
      ['A', 'B'],
      [1, 3],
      [2, 4],
    ]);
    current.sheets.push({ name: 'Added', rows: [[createCell('Y')]] });
    const result = compareWorkbookBaseline(baseline, current);
    expect(result.sheets).toMatchObject([
      {
        sheetName: 'Data',
        presence: 'both',
        rowDelta: 1,
        columnDelta: 1,
        changedCellCount: 4,
        layoutChanged: true,
      },
      {
        sheetName: 'Removed',
        presence: 'removed',
        rowDelta: -2,
        columnDelta: -1,
        changedCellCount: 2,
      },
      { sheetName: 'Added', presence: 'added', rowDelta: 1, columnDelta: 1, changedCellCount: 1 },
    ]);
    expect(result.changedCellCount).toBe(7);
    expect(result.changedCells.find((cell) => cell.sheet === 'Removed')?.after).toEqual({
      present: false,
    });
  });

  it('flags shifted headers as a layout change instead of inventing comparable aggregates or trends', () => {
    const baseline = workbook([
      ['Revenue', 'Cost'],
      [100, 10],
      [200, 20],
    ]);
    const current = workbook([
      ['Cost', 'Revenue'],
      [10, 100],
      [20, 200],
    ]);
    const result = compareWorkbookBaseline(baseline, current);
    expect(result.sheets[0]).toMatchObject({
      headerChanged: true,
      layoutChanged: true,
      rowDelta: 0,
      columnDelta: 0,
      changedCellCount: 6,
    });
    expect(result.notes.join(' ')).toContain('Layouts differ');
    expect(result.notes.join(' ')).toContain(
      'No record alignment, trends, or comparable aggregate deltas',
    );
    expect(result).not.toHaveProperty('aggregates');
    expect(result.changedCells[0]).toMatchObject({
      address: "'Data'!A1:A1",
      before: { value: 'Revenue' },
      after: { value: 'Cost' },
    });
  });

  it('retains exact Date instants, formula edits and formula cache changes independently of epoch metadata', () => {
    const baseline = workbook([
      ['Date', 'Formula', 'Cache'],
      [new Date('2024-01-01T00:00:00.000Z'), 4, 4],
    ]);
    baseline.sheets[0]!.rows[1]![1] = createCell(4, { formula: '=2+2' });
    baseline.sheets[0]!.rows[1]![2] = createCell(4, { formula: '=2+2' });
    const current = structuredClone(baseline);
    current.dateSystem = '1904';
    current.sheets[0]!.rows[1]![0] = createCell(new Date('2024-01-01T00:00:00.001Z'));
    current.sheets[0]!.rows[1]![1] = createCell(4, { formula: '=1+3' });
    current.sheets[0]!.rows[1]![2] = createCell(999, { formula: '=2+2' });
    const result = compareWorkbookBaseline(baseline, current);
    expect(result).toMatchObject({
      baselineDateSystem: '1900',
      currentDateSystem: '1904',
      dateSystemChanged: true,
      changedCellCount: 3,
    });
    expect(result.changedCells[0]?.before.value).toEqual({
      kind: 'date',
      epochMilliseconds: 1704067200000,
      iso: '2024-01-01T00:00:00.000Z',
    });
    expect(result.changedCells[0]?.after.value).toMatchObject({ epochMilliseconds: 1704067200001 });
    expect(result.changedCells[1]).toMatchObject({
      before: { formula: '=2+2' },
      after: { formula: '=1+3' },
    });
    expect(result.changedCells[2]).toMatchObject({
      before: { value: 4 },
      after: { value: 999, formula: '=2+2' },
    });
    const epochOnly = compareWorkbookBaseline(baseline, { ...baseline, dateSystem: '1904' });
    expect(epochOnly.changedCellCount).toBe(0);
    expect(epochOnly.notes.join(' ')).toContain('may now mean different dates');
    expect(compareWorkbookBaseline(baseline, structuredClone(baseline)).changedCellCount).toBe(0);
  });

  it('bounds exact changed-cell examples globally but counts the complete workbook', () => {
    const baseline = workbook([['Number'], ...Array.from({ length: 100 }, (_, index) => [index])]);
    const current = workbook([
      ['Number'],
      ...Array.from({ length: 100 }, (_, index) => [index + 1]),
    ]);
    const result = compareWorkbookBaseline(baseline, current);
    expect(result.changedCellCount).toBe(100);
    expect(result.changedCells).toHaveLength(BRIEFING_LIMITS.changedCellExamples);
    expect(result.omittedChangedCellCount).toBe(80);
    expect(result.changedCells[19]?.address).toBe("'Data'!A21:A21");
    expect(compareWorkbookBaseline(baseline, current)).toEqual(result);
  });

  it('compares absent versus blank, type/format, and exact special number values', () => {
    const baseline = workbook([['A', 'B'], [], [-0, NaN]]);
    const current = workbook([['A', 'B'], [null], [0, NaN]]);
    current.sheets[0]!.rows[0]![1]!.numberFormat = '0.00';
    const result = compareWorkbookBaseline(baseline, current);
    expect(result.changedCellCount).toBe(3);
    expect(result.changedCells[1]).toMatchObject({
      before: { present: false },
      after: { present: true, value: null },
    });
    expect(compareWorkbookBaseline(workbook([[NaN]]), workbook([[NaN]])).changedCellCount).toBe(0);
  });
});
