import { describe, expect, it } from 'vitest';
import { createCell, type Workbook } from '@excel-agent/engine';
import {
  calculateAggregate,
  getWorkbookOverview,
  profileColumn,
  readCellRange,
  searchSheet,
} from '../src/read-tools.js';

function sampleWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Sales',
        rows: [
          [
            createCell('Order ID'),
            createCell('Customer'),
            createCell('Revenue'),
            createCell('Date'),
          ],
          [createCell('ORD-1'), createCell('Alice'), createCell(100), createCell('2026-01-01')],
          [createCell('ORD-2'), createCell('Bob'), createCell(250), createCell('2026-01-05')],
          [createCell('ORD-3'), createCell('Charlie'), createCell(150), createCell('2026-01-10')],
          [createCell('ORD-4'), createCell('Alice'), createCell(500), createCell('2026-01-15')],
        ],
      },
    ],
  };
}

describe('read tools for workbook inspection', () => {
  it('returns high level workbook overview', () => {
    const wb = sampleWorkbook();
    const overview = getWorkbookOverview(wb);
    expect(overview.sheets).toHaveLength(1);
    expect(overview.sheets[0]?.name).toBe('Sales');
    expect(overview.sheets[0]?.rowCount).toBe(5);
    expect(overview.sheets[0]?.headers).toEqual(['Order ID', 'Customer', 'Revenue', 'Date']);
  });

  it('profiles a column accurately with statistics and cardinality', () => {
    const wb = sampleWorkbook();
    const profile = profileColumn(wb, 'Sales', 'C');
    expect('error' in profile).toBe(false);
    if (!('error' in profile)) {
      expect(profile.column).toBe('C');
      expect(profile.headerName).toBe('Revenue');
      expect(profile.inferredType).toBe('numeric');
      expect(profile.sum).toBe(1000);
      expect(profile.average).toBe(250);
      expect(profile.min).toBe(100);
      expect(profile.max).toBe(500);
      expect(profile.distinctCount).toBe(4);
    }
  });

  it('reads a bounded cell range', () => {
    const wb = sampleWorkbook();
    const result = readCellRange(wb, 'Sales', 1, 3, 'A', 'B');
    expect('error' in result).toBe(false);
    if (!('error' in result)) {
      expect(result.rows).toHaveLength(3);
      expect(result.startRow).toBe(1);
      expect(result.endRow).toBe(3);
      expect(result.rows[1]?.cells['A (Order ID)']).toBe('ORD-1');
      expect(result.rows[1]?.cells['B (Customer)']).toBe('Alice');
    }
  });

  it('searches for values across rows', () => {
    const wb = sampleWorkbook();
    const searchRes = searchSheet(wb, 'Sales', 'Alice');
    expect('error' in searchRes).toBe(false);
    if (!('error' in searchRes)) {
      expect(searchRes.matches.length).toBe(2);
      expect(searchRes.matches[0]?.cell).toBe('B2');
      expect(searchRes.matches[1]?.cell).toBe('B5');
    }
  });

  it('calculates deterministic aggregates on numerical columns', () => {
    const wb = sampleWorkbook();
    const sumRes = calculateAggregate(wb, 'Sales', 'C', 'sum');
    expect('error' in sumRes).toBe(false);
    if (!('error' in sumRes)) {
      expect(sumRes.value).toBe(1000);
    }

    const avgRes = calculateAggregate(wb, 'Sales', 'C', 'avg');
    expect('error' in avgRes).toBe(false);
    if (!('error' in avgRes)) {
      expect(avgRes.value).toBe(250);
    }
  });
});
