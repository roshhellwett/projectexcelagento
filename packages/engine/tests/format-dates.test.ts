import { describe, expect, it } from 'vitest';

import {
  applyPatch,
  applyPatch as applyWorkbookPatch,
  cloneWorkbook,
  displayCell,
  formatDatesArgsSchema,
  formatDatesOperation,
  createCell,
  type FormatDatesArgs,
  type Cell,
  type Workbook,
} from '../src/index.js';

function cell(value: Cell['value'], extras: Partial<Omit<Cell, 'value'>> = {}): Cell {
  return createCell(value, extras);
}

function workbook(rows: Cell[][]): Workbook {
  return { sheets: [{ name: 'Orders', rows }] };
}

function args(overrides: Partial<FormatDatesArgs> = {}) {
  return formatDatesArgsSchema.parse({
    sheet: 'Orders',
    column: 'B',
    format: 'YYYY-MM-DD',
    ...overrides,
  });
}

describe('format_dates operation', () => {
  it('converts an Excel 1900-system serial, including a leap day', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell(43890)], // 2020-02-29 in the Excel 1900 date system.
    ]);

    const preview = formatDatesOperation.preview(before, args());
    const result = formatDatesOperation.apply(before, args());
    const converted = result.workbook.sheets[0]?.rows[1]?.[1];

    expect(preview).toMatchObject({ valid: true, affectedCells: 1, requiresConfirmation: false });
    expect(preview.changes[0]?.after.value).toEqual(new Date(Date.UTC(2020, 1, 29)));
    expect(converted?.value).toEqual(new Date(Date.UTC(2020, 1, 29)));
    expect(converted?.numberFormat).toBe('YYYY-MM-DD');
    expect(displayCell(converted as Cell)).toBe('2020-02-29');
    expect(result.report.affectedCells).toBe(1);
    expect(result.report.warnings).toHaveLength(0);
  });

  it('handles the Excel 1900 leap-year boundary without accepting fake February 29', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell(59)],
      [cell('A-2'), cell(60)],
      [cell('A-3'), cell(61)],
    ]);

    const result = formatDatesOperation.apply(before, args());
    const rows = result.workbook.sheets[0]?.rows ?? [];

    expect(rows[1]?.[1]?.value).toEqual(new Date(Date.UTC(1900, 1, 28)));
    expect(rows[2]?.[1]?.value).toBe(60);
    expect(rows[3]?.[1]?.value).toEqual(new Date(Date.UTC(1900, 2, 1)));
    expect(result.report.affectedCells).toBe(2);
    expect(result.report.warnings[0]?.code).toBe('invalid-date');
  });

  it('converts unambiguous ISO and locale-shaped text dates', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell('2024-02-29')],
      [cell('A-2'), cell('31/12/2024')],
      [cell('A-3'), cell('12/31/2024')],
    ]);

    const result = formatDatesOperation.apply(before, args({ format: 'MM/DD/YYYY' }));
    const rows = result.workbook.sheets[0]?.rows ?? [];

    expect(displayCell(rows[1]?.[1] as Cell)).toBe('02/29/2024');
    expect(displayCell(rows[2]?.[1] as Cell)).toBe('12/31/2024');
    expect(displayCell(rows[3]?.[1] as Cell)).toBe('12/31/2024');
    expect(result.report.affectedCells).toBe(3);
  });

  it('flags DD/MM versus MM/DD ambiguity and does not guess', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell('01/02/2020')],
    ]);

    const validation = formatDatesOperation.validate(before, args());
    const preview = formatDatesOperation.preview(before, args());
    const result = formatDatesOperation.apply(before, args());

    expect(validation.valid).toBe(true);
    expect(validation.warnings).toMatchObject([
      { code: 'ambiguous-date', location: { sheet: 'Orders', row: 2, column: 'B' } },
    ]);
    expect(preview.requiresConfirmation).toBe(true);
    expect(preview.affectedCells).toBe(0);
    expect(result.workbook).toEqual(before);
    expect(result.report.affectedCells).toBe(0);
  });

  it('keeps empty cells empty and skips non-date values', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell(null)],
      [cell('A-2'), cell('')],
      [cell('A-3'), cell('not a date')],
      [cell('A-4'), cell(0)],
      [cell('A-5'), cell('44986')], // Numeric text is not guessed as an Excel serial.
    ]);

    const result = formatDatesOperation.apply(before, args());
    const rows = result.workbook.sheets[0]?.rows ?? [];

    expect(rows.slice(1).map((row) => row[1]?.value)).toEqual([null, '', 'not a date', 0, '44986']);
    expect(result.report.affectedCells).toBe(0);
    expect(result.report.unchangedCells).toBe(5);
  });

  it('flags an impossible leap day instead of normalizing it', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell('29/02/2019')],
    ]);

    const result = formatDatesOperation.apply(before, args());

    expect(result.workbook).toEqual(before);
    expect(result.report.warnings).toMatchObject([{ code: 'invalid-date' }]);
  });

  it('preserves formulas and cells outside the target column', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date'), cell('Total')],
      [cell('A-1'), cell(43890), cell(10, { formula: '=1+9' })],
      [cell('A-2'), cell('2024-01-01'), cell('keep me')],
    ]);

    const result = formatDatesOperation.apply(before, args());
    const invariant = formatDatesOperation.invariants(before, result.workbook, args());

    expect(invariant).toEqual({ valid: true, errors: [] });
    expect(result.workbook.sheets[0]?.rows[1]?.[2]).toEqual(cell(10, { formula: '=1+9' }));
    expect(result.workbook.sheets[0]?.rows[2]?.[2]?.value).toBe('keep me');
  });

  it('provides a reversible inverse patch', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell(43890, { numberFormat: 'General' })],
      [cell('A-2'), cell('2024-01-01')],
    ]);

    const result = formatDatesOperation.apply(before, args());
    const restored = applyPatch(result.workbook, result.inverse);

    expect(restored).toEqual(before);
    expect(formatDatesOperation.invariants(before, result.workbook, args()).valid).toBe(true);
  });

  it('supports the 1904 Excel date system', () => {
    const before = workbook([
      [cell('Order date'), cell('Ship date')],
      [cell('A-1'), cell(0)],
      [cell('A-2'), cell(1)],
    ]);

    const result = formatDatesOperation.apply(before, args({ dateSystem: '1904' }));
    const rows = result.workbook.sheets[0]?.rows ?? [];

    expect(rows[1]?.[1]?.value).toEqual(new Date(Date.UTC(1904, 0, 1)));
    expect(rows[2]?.[1]?.value).toEqual(new Date(Date.UTC(1904, 0, 2)));
  });

  it('reports structural validation errors before applying', () => {
    const before = workbook([[cell('Order date')], [cell('A-1')]]);
    const invalidArgs = args({ column: 'C' });

    expect(formatDatesOperation.validate(before, invalidArgs)).toMatchObject({
      valid: false,
      errors: [{ code: 'missing-column' }],
    });
    expect(() => formatDatesOperation.apply(before, invalidArgs)).toThrow('does not exist');
    expect(cloneWorkbook(before)).toEqual(before);
  });
});

describe('workbook helpers used by operations', () => {
  it('applies an inverse patch without mutating the original workbook', () => {
    const before = workbook([[cell('Date')], [cell(43890)]]);
    const result = formatDatesOperation.apply(before, args({ column: 'A' }));
    const restored = applyWorkbookPatch(result.workbook, result.inverse);

    expect(restored).toEqual(before);
    expect(before.sheets[0]?.rows[1]?.[0]?.value).toBe(43890);
  });
});
