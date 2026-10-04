import { describe, expect, it } from 'vitest';

import { formatCellDateValue } from './cell-format.js';

/**
 * The grid is where a wrong date becomes a wrong number on screen: the old renderer hard-coded the
 * 1900 epoch, so a 1904 workbook was off by four years and a day, and it truncated every date to
 * its day. These tests pin the rendering rules the shared engine module now owns.
 */

// 44197 = 2021-01-01 in the 1900 system. Its fractional day carries the time of day.
const JAN_1_2021 = 44197;

describe('grid cell date formatting', () => {
  it('renders a date-only serial as an ISO day', () => {
    expect(formatCellDateValue(JAN_1_2021, 'yyyy-mm-dd')).toBe('2021-01-01');
    expect(formatCellDateValue(JAN_1_2021, 'MM/DD/YYYY')).toBe('2021-01-01');
  });

  it('shows the time of day when the format asks for it', () => {
    expect(formatCellDateValue(JAN_1_2021 + 0.5, 'yyyy-mm-dd hh:mm')).toBe('2021-01-01 12:00');
    expect(formatCellDateValue(JAN_1_2021 + 0.75, 'yyyy-mm-dd hh:mm:ss')).toBe(
      '2021-01-01 18:00:00',
    );
  });

  it('drops the time when the format has no time codes', () => {
    expect(formatCellDateValue(JAN_1_2021 + 0.5, 'yyyy-mm-dd')).toBe('2021-01-01');
  });

  it('reads a locale-tagged month-name format that a naive regex rejects', () => {
    expect(formatCellDateValue(JAN_1_2021, '[$-409]mmm d, yyyy;@')).toBe('Jan 1, 2021');
    expect(formatCellDateValue(JAN_1_2021, 'mmmm d, yyyy')).toBe('Jan 1, 2021');
  });

  it('renders an elapsed duration rather than a clock time', () => {
    // `[mm]` keeps counting past the hour, where `mm` in a clock format restarts every hour.
    expect(formatCellDateValue(JAN_1_2021 + 0.75, '[h]:mm:ss')).toBe('18:00:00');
    expect(formatCellDateValue(JAN_1_2021 + 0.75, '[mm]:ss')).toBe('1080:00');
    expect(formatCellDateValue(JAN_1_2021 + 0.75, 'hh:mm:ss')).toBe('18:00:00');
  });

  it('leaves numeric formats alone', () => {
    expect(formatCellDateValue(JAN_1_2021, 'General')).toBeNull();
    expect(formatCellDateValue(JAN_1_2021, '#,##0.00')).toBeNull();
    expect(formatCellDateValue(JAN_1_2021, undefined)).toBeNull();
    expect(formatCellDateValue(1234.5, '0.00')).toBeNull();
  });

  it('renders a cell that already holds a Date, with or without a format', () => {
    expect(formatCellDateValue(new Date(Date.UTC(2021, 0, 1, 12, 30)))).toBe('2021-01-01');
    expect(formatCellDateValue(new Date(Date.UTC(2021, 0, 1, 12, 30)), 'yyyy-mm-dd')).toBe(
      '2021-01-01',
    );
    expect(formatCellDateValue(new Date(Date.UTC(2021, 0, 1, 12, 30)), 'yyyy-mm-dd hh:mm')).toBe(
      '2021-01-01 12:30',
    );
  });

  it('never renders text, booleans, blanks, or an invalid Date as a date', () => {
    expect(formatCellDateValue('2021-01-01', 'yyyy-mm-dd')).toBeNull();
    expect(formatCellDateValue('not a date')).toBeNull();
    expect(formatCellDateValue(true, 'yyyy-mm-dd')).toBeNull();
    expect(formatCellDateValue(null, 'yyyy-mm-dd')).toBeNull();
    expect(formatCellDateValue(new Date('nonsense'), 'yyyy-mm-dd')).toBeNull();
  });

  it('honours the 1904 date system, where the same serial is four years and a day later', () => {
    expect(formatCellDateValue(JAN_1_2021, 'yyyy-mm-dd', '1900')).toBe('2021-01-01');
    expect(formatCellDateValue(JAN_1_2021, 'yyyy-mm-dd', '1904')).toBe('2025-01-02');
  });

  it('rejects the fictitious 1900-02-29 instead of showing a day that never existed', () => {
    expect(formatCellDateValue(60, 'yyyy-mm-dd')).toBeNull();
  });
});
