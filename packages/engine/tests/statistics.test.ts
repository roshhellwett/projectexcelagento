import { describe, expect, it } from 'vitest';
import {
  aggregateCells,
  applyOperation,
  createCell,
  createWorkbookValueReader,
  fitLinearRegression,
  matchesCriteria,
  summarizeNumericValues,
  type Workbook,
} from '../src/index.js';

describe('descriptive statistics', () => {
  it('matches a known distribution and inclusive quartiles', () => {
    const result = summarizeNumericValues([2, 4, 4, 4, 5, 5, 7, 9]);
    expect(result).toMatchObject({
      numericCount: 8,
      sum: 40,
      mean: 5,
      min: 2,
      max: 9,
      median: 4.5,
      q1: 4,
      q3: 5.5,
      iqr: 1.5,
      outlierCount: 1,
      populationStandardDeviation: 2,
    });
    expect(result.sampleVariance).toBeCloseTo(32 / 7, 12);
    expect(result.sampleStandardDeviation).toBeCloseTo(Math.sqrt(32 / 7), 12);
  });

  it('reports missing and invalid observations without coercing them to zero', () => {
    const result = summarizeNumericValues([
      null,
      '',
      '  ',
      'bad',
      true,
      Infinity,
      new Date(),
      '1,200',
      '1e-5',
    ]);
    expect(result).toMatchObject({
      totalCount: 9,
      missingCount: 3,
      nonNumericCount: 4,
      numericCount: 2,
      sum: 1200.00001,
    });
    expect(summarizeNumericValues([])).toMatchObject({
      mean: null,
      min: null,
      sampleVariance: null,
    });
    expect(summarizeNumericValues([5])).toMatchObject({
      mean: 5,
      sampleVariance: null,
      populationStandardDeviation: 0,
    });
  });

  it('preserves variance when observations have a large common offset', () => {
    const result = summarizeNumericValues([1e12 + 1, 1e12 + 2, 1e12 + 3, 1e12 + 4]);
    expect(result.mean).toBe(1e12 + 2.5);
    expect(result.sampleVariance).toBeCloseTo(5 / 3, 12);
    expect(summarizeNumericValues([1e-12, 3e-12]).mean).toBeCloseTo(2e-12, 20);
  });

  it('surfaces numerical overflow instead of inventing a zero result', () => {
    const result = summarizeNumericValues([1e308, 1e308]);
    expect(result.sum).toBeNull();
    expect(result.warnings.join(' ')).toMatch(/finite numeric range/);
  });
});

describe('Pearson correlation and ordinary least squares', () => {
  it('recovers an exact line while excluding incomplete pairs', () => {
    const result = fitLinearRegression([
      [1, 5],
      [2, 8],
      [3, 11],
      [4, 14],
      [null, 20],
      [5, 'bad'],
    ]);
    expect(result).toMatchObject({
      pairCount: 4,
      excludedPairCount: 2,
      slope: 3,
      intercept: 2,
      residualStandardError: 0,
    });
    expect(result.correlation).toBeCloseTo(1, 14);
    expect(result.rSquared).toBeCloseTo(1, 14);
  });

  it('matches a non-perfect reference fit', () => {
    const result = fitLinearRegression([
      [1, 2],
      [2, 4],
      [3, 5],
      [4, 4],
    ]);
    expect(result.slope).toBeCloseTo(0.7, 12);
    expect(result.intercept).toBeCloseTo(2, 12);
    expect(result.rSquared).toBeCloseTo(12.25 / 23.75, 12);
    expect(result.residualStandardError).toBeCloseTo(Math.sqrt(1.15), 12);
  });

  it('handles negative association and identifies undefined fits', () => {
    const negative = fitLinearRegression([
      [1, 6],
      [2, 4],
      [3, 2],
    ]);
    expect(negative).toMatchObject({ slope: -2, intercept: 8 });
    expect(negative.correlation).toBeCloseTo(-1, 14);
    expect(fitLinearRegression([[1, 2]])).toMatchObject({ slope: null, correlation: null });
    expect(
      fitLinearRegression([
        [1, 2],
        [1, 3],
      ]),
    ).toMatchObject({ slope: null, correlation: null });
    expect(
      fitLinearRegression([
        [1, 3],
        [2, 3],
      ]),
    ).toMatchObject({ slope: 0, intercept: 3, correlation: null, rSquared: null });
  });
});

describe('live workbook values used by analytics', () => {
  it.each([
    ['median', 3],
    ['stdev', Math.sqrt(2)],
    ['count_distinct', 2],
  ] as const)('supports %s in grouped summaries', (aggregation, expected) => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Data',
          rows: [
            [createCell('Group'), createCell('Value')],
            [createCell('North'), createCell(2)],
            [createCell('North'), createCell(4)],
          ],
        },
      ],
    };
    const result = applyOperation(workbook, 'group_and_summarize', {
      sheet: 'Data',
      groupBy: ['A'],
      valueColumn: 'B',
      aggregation,
      targetSheet: 'Summary',
    });
    expect(result.ok).toBe(true);
    expect(result.workbook.sheets[1]?.rows[1]?.[1]?.value).toBeCloseTo(expected, 12);
  });
  it('honors the 1904 epoch without leaking it to other workbook evaluations', () => {
    const cells = [[createCell(1), createCell(0, { formula: 'YEAR(A1)' })]];
    const legacy = createWorkbookValueReader({
      dateSystem: '1904',
      sheets: [{ name: 'Data', rows: cells }],
    });
    const modern = createWorkbookValueReader({
      dateSystem: '1900',
      sheets: [{ name: 'Data', rows: cells }],
    });
    expect(legacy('Data', 1, 1)).toBe(1904);
    expect(modern('Data', 1, 1)).toBe(1900);
  });
  it('recalculates dependencies and contains cycles without using stale cached values', () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Data',
          rows: [
            [
              createCell(7),
              createCell(999, { formula: 'A1*2' }),
              createCell(888, { formula: 'B1+1' }),
              createCell(5, { formula: 'D1' }),
            ],
          ],
        },
      ],
    };
    const read = createWorkbookValueReader(workbook);
    expect(read('data', 1, 1)).toBe(14);
    expect(read('Data', 2, 1)).toBe(15);
    expect(read('Data', 3, 1)).toBe('#ERROR!');
    expect(read('Missing', 0, 1)).toBe('#REF!');
    expect(workbook.sheets[0]?.rows[0]?.[1]?.value).toBe(999);
  });

  it('aggregates a column beyond the function-argument stack limit', () => {
    const cells = Array.from({ length: 150_000 }, (_, index) => createCell(index));
    expect(aggregateCells(cells, 'max').value).toBe(149_999);
    expect(aggregateCells(cells, 'min').value).toBe(0);
  });

  it('compares grouped criteria as whole numeric values, with operator aliases', () => {
    expect(matchesCriteria(1200, '=1,200')).toBe(true);
    expect(matchesCriteria(1200, '==1,200')).toBe(true);
    expect(matchesCriteria(1200, '!=1,200')).toBe(false);
    expect(matchesCriteria(1200, '>100oops')).toBe(false);
  });
});
