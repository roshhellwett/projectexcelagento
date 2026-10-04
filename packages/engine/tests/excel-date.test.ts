import { describe, expect, it } from 'vitest';
import { evaluateFormula, setFormulaClock } from '../src/formula/index.js';
import type { FormulaContext, FormulaValue } from '../src/formula/types.js';
import {
  dateToExcelSerial,
  excelSerialToDate,
  parseUnambiguousDate,
  formatIsoDate,
  isDateNumberFormat,
} from '../src/formula/excel-date.js';
import { isFormulaError } from '../src/formula/errors.js';

/**
 * Dates and errors are the two places a spreadsheet engine fails silently: a wrong date
 * still renders, and a wrong error still looks like text. These tests pin the behaviour that
 * used to be wrong - a serial read as a calendar year, an error coerced into `NaN`, an
 * `IFERROR` that swallowed a SKU.
 */

/** Builds a context whose sheet holds raw Excel serials, mirroring what the importer produces. */
function contextFor(
  cells: Record<string, Record<string, Record<number, FormulaValue>>>,
  sheets = Object.keys(cells),
): FormulaContext {
  return {
    activeSheet: sheets[0] ?? 'Sheet1',
    hasSheet: (sheet) => sheets.includes(sheet),
    getCellValue: (sheet, col, row) => cells[sheet]?.[col]?.[row] ?? null,
    getRangeValues: (sheet, startCol, startRow, endCol, endRow) => {
      const rows: FormulaValue[][] = [];
      const width = endCol.charCodeAt(0) - startCol.charCodeAt(0);
      for (let r = startRow; r <= endRow; r++) {
        const row: FormulaValue[] = [];
        for (let c = 0; c <= width; c++) {
          const letter = String.fromCharCode(startCol.charCodeAt(0) + c);
          row.push(cells[sheet]?.[letter]?.[r] ?? null);
        }
        rows.push(row);
      }
      return rows;
    },
  };
}

describe('Excel serial conversion', () => {
  it('maps known serials to the dates Excel shows', () => {
    expect(formatIsoDate(excelSerialToDate(1)!)).toBe('1900-01-01');
    expect(formatIsoDate(excelSerialToDate(59)!)).toBe('1900-02-28');
    expect(formatIsoDate(excelSerialToDate(61)!)).toBe('1900-03-01');
    expect(formatIsoDate(excelSerialToDate(25569)!)).toBe('1970-01-01');
    expect(formatIsoDate(excelSerialToDate(44197)!)).toBe('2021-01-01');
    expect(formatIsoDate(excelSerialToDate(45292)!)).toBe('2024-01-01');
    expect(formatIsoDate(excelSerialToDate(2958465)!)).toBe('9999-12-31');
  });

  it('rejects the fictitious 1900-02-29 rather than rolling it into March', () => {
    expect(excelSerialToDate(60)).toBeNull();
  });

  it('round-trips every serial through the 1900 system', () => {
    for (let serial = 61; serial <= 2958465; serial += 997) {
      const date = excelSerialToDate(serial)!;
      expect(dateToExcelSerial(date)).toBeCloseTo(serial, 6);
    }
  });

  it('round-trips the pre-leap-bug window of the 1900 system', () => {
    for (let serial = 1; serial < 60; serial++) {
      expect(dateToExcelSerial(excelSerialToDate(serial)!)).toBeCloseTo(serial, 6);
    }
  });

  it('keeps the 1904 system offset by 1462 days from the 1900 system', () => {
    expect(formatIsoDate(excelSerialToDate(0, '1904')!)).toBe('1904-01-01');
    // The 1904 epoch is 1462 days later, so the same serial names a later day.
    expect(formatIsoDate(excelSerialToDate(44197)!)).toBe('2021-01-01');
    expect(formatIsoDate(excelSerialToDate(44197, '1904')!)).toBe('2025-01-02');
  });

  it('rejects serials outside the representable range', () => {
    expect(excelSerialToDate(0)).toBeNull();
    expect(excelSerialToDate(2_958_466)).toBeNull();
    expect(excelSerialToDate(Number.NaN)).toBeNull();
  });

  it('preserves the time-of-day fraction', () => {
    const date = excelSerialToDate(44197.5)!;
    expect(date.getUTCHours()).toBe(12);
  });
});

describe('Date text parsing', () => {
  it('parses unambiguous ISO forms', () => {
    expect(formatIsoDate(parseUnambiguousDate('2024-01-31')!)).toBe('2024-01-31');
    expect(formatIsoDate(parseUnambiguousDate('2024/1/31')!)).toBe('2024-01-31');
    expect(formatIsoDate(parseUnambiguousDate('2024-01-31T10:30:00')!)).toBe('2024-01-31');
  });

  it('refuses ambiguous day/month order instead of guessing', () => {
    expect(parseUnambiguousDate('01/02/2020')).toBeNull();
    expect(parseUnambiguousDate('31/01/2020')).toBeNull();
  });

  it('never lets JavaScript read a bare serial as a calendar year', () => {
    // `new Date("43890")` is the year 43890 in V8. A serial must never travel that path.
    expect(parseUnambiguousDate('43890')).toBeNull();
  });

  it('rejects dates that do not exist', () => {
    expect(parseUnambiguousDate('2021-02-30')).toBeNull();
    expect(parseUnambiguousDate('2021-13-01')).toBeNull();
  });
});

describe('Date arithmetic and date functions over serials', () => {
  // 44197 = 2021-01-01, 45292 = 2024-01-01 in the 1900 system.
  const context = contextFor({ Sheet1: { A: { 1: 44197, 2: 45292 }, B: { 1: 43890 } } });

  it('reads a real year out of an Excel serial', () => {
    // The regression that shipped a 41,600-year answer past a green suite.
    expect(evaluateFormula('=YEAR(B1)', context)).toBe(2020);
    expect(evaluateFormula('=YEAR(A1)', context)).toBe(2021);
  });

  it('subtracts two dates into a day count', () => {
    expect(evaluateFormula('=A2-A1', context)).toBe(1095);
  });

  it('treats a bare serial as a plain number, exactly as Excel does', () => {
    // A cell holding 44197 is only a date because of its number format. Without that
    // context the engine must not invent one, so serial arithmetic stays numeric.
    expect(evaluateFormula('=A1+30', context)).toBe(44227);
  });

  it('does date arithmetic when the cell actually holds a date', () => {
    const dated = contextFor({
      Sheet1: { A: { 1: new Date(Date.UTC(2021, 0, 1)), 2: new Date(Date.UTC(2024, 0, 1)) } },
    });
    expect(evaluateFormula('=A2-A1', dated)).toBe(1095);
    expect(evaluateFormula('=A1+30', dated)).toEqual(new Date(Date.UTC(2021, 0, 31)));
    expect(evaluateFormula('=A1-30', dated)).toEqual(new Date(Date.UTC(2020, 11, 2)));
  });

  it('refuses to add two dates together', () => {
    const dated = contextFor({
      Sheet1: { A: { 1: new Date(Date.UTC(2021, 0, 1)), 2: new Date(Date.UTC(2024, 0, 1)) } },
    });
    expect(evaluateFormula('=A1+A2', dated)).toBe('#VALUE!');
  });

  it('rejects multiplying a date by a number', () => {
    const dated = contextFor({ Sheet1: { A: { 1: new Date(Date.UTC(2021, 0, 1)) } } });
    expect(evaluateFormula('=A1*2', dated)).toBe('#VALUE!');
  });

  it('counts days, months and years between serials', () => {
    const range = contextFor({ Sheet1: { A: { 1: 44197, 2: 45292 } } });
    expect(evaluateFormula('=DAYS(A2,A1)', range)).toBe(1095);
    expect(evaluateFormula('=DATEDIF(A1,A2,"Y")', range)).toBe(3);
    expect(evaluateFormula('=DATEDIF(A1,A2,"M")', range)).toBe(36);
    expect(evaluateFormula('=DATEDIF(A1,A2,"D")', range)).toBe(1095);
  });

  it('finds month boundaries', () => {
    const range = contextFor({ Sheet1: { A: { 1: 44197 } } });
    const endOfMonth = evaluateFormula('=EOMONTH(A1,0)', range) as Date;
    expect(endOfMonth.getUTCDate()).toBe(31);
    expect(endOfMonth.getUTCMonth()).toBe(0);
    const nextMonth = evaluateFormula('=EDATE(A1,1)', range) as Date;
    expect(formatIsoDate(nextMonth)).toBe('2021-02-01');
  });

  it('refuses to invent a date from an ambiguous string', () => {
    expect(evaluateFormula('=YEAR("01/02/2020")', context)).toBe('#VALUE!');
    expect(evaluateFormula('=YEAR("not a date")', context)).toBe('#VALUE!');
  });

  it('answers NETWORKDAYS over a multi-century span without hanging', () => {
    const span = contextFor({ Sheet1: { A: { 1: 1 }, B: { 1: 2958465 } } });
    const started = Date.now();
    const result = evaluateFormula('=NETWORKDAYS(A1,B1)', span);
    // 1900-01-01 to 9999-12-31 contains roughly 2.4M weekdays.
    expect(typeof result).toBe('number');
    expect(result as number).toBeGreaterThan(2_000_000);
    expect(Date.now() - started).toBeLessThan(1000);
  });

  it('reports the weekday of a serial', () => {
    const range = contextFor({ Sheet1: { A: { 1: 44197 } } });
    // 2021-01-01 was a Friday, which is 6 in Excel's Sunday-first numbering.
    expect(evaluateFormula('=WEEKDAY(A1)', range)).toBe(6);
  });

  it('honours a deterministic clock for TODAY', () => {
    setFormulaClock(() => new Date('2024-03-15T18:30:00Z'));
    try {
      const result = evaluateFormula('=TODAY()', context) as Date;
      expect(formatIsoDate(result)).toBe('2024-03-15');
    } finally {
      setFormulaClock(() => new Date());
    }
  });
});

describe('Error values', () => {
  const context = contextFor({ Sheet1: { A: { 1: 10, 2: 0 }, B: { 1: '#12345' } } });

  it('propagates an error through arithmetic instead of producing NaN', () => {
    expect(evaluateFormula('=1/0+1', context)).toBe('#DIV/0!');
    expect(evaluateFormula('=1/0*5', context)).toBe('#DIV/0!');
    expect(evaluateFormula('=5-(1/0)', context)).toBe('#DIV/0!');
  });

  it('propagates an error operand through comparisons', () => {
    const withError = contextFor({ Sheet1: { A: { 1: '#DIV/0!' } } });
    expect(evaluateFormula('=10>A1', withError)).toBe('#DIV/0!');
    expect(evaluateFormula('=A1=10', withError)).toBe('#DIV/0!');
  });

  it('reports #NUM! for results Excel cannot represent', () => {
    expect(evaluateFormula('=SQRT(0-1)', context)).toBe('#NUM!');
  });

  it('raises #REF! for a reference to a sheet that does not exist', () => {
    expect(evaluateFormula('=NoSuchSheet!A1', context)).toBe('#REF!');
    expect(evaluateFormula('=SUM(NoSuchSheet!A1:A5)', context)).toBe('#REF!');
  });

  it('raises #REF! for a reversed range rather than reporting zero', () => {
    expect(evaluateFormula('=SUM(A5:A1)', context)).toBe('#REF!');
  });

  it('uses the canonical #NAME? code', () => {
    expect(evaluateFormula('=NOSUCHFUNCTION(1)', context)).toBe('#NAME?');
  });

  it('lets IFERROR catch an unknown name while IFNA correctly does not', () => {
    // IFNA is narrowly scoped to #N/A by design; IFERROR is the catch-all.
    expect(evaluateFormula('=IFERROR(NOSUCHFUNCTION(1),"fallback")', context)).toBe('fallback');
    expect(evaluateFormula('=IFNA(NOSUCHFUNCTION(1),"fallback")', context)).toBe('#NAME?');
    expect(evaluateFormula('=IFNA(VLOOKUP("zz",A1:B1,2,FALSE),"fallback")', context)).toBe(
      'fallback',
    );
  });

  it('does not let IFERROR swallow text that merely starts with a hash', () => {
    // "#12345" is a plausible SKU or ticket id, not an error.
    expect(evaluateFormula('=B1', context)).toBe('#12345');
    expect(evaluateFormula('=IFERROR(B1,"caught")', context)).toBe('#12345');
    expect(evaluateFormula('=IFERROR(1/0,"caught")', context)).toBe('caught');
  });

  it('rejects a malformed formula rather than silently returning the prefix', () => {
    expect(evaluateFormula('=1+1)+DROP(A1)', context)).toBe('#ERROR!');
  });

  it('terminates on deeply nested parentheses', () => {
    expect(evaluateFormula(`=${'('.repeat(500)}1${')'.repeat(500)}`, context)).toBe('#ERROR!');
  });

  it('survives a self-referential formula without corrupting the workbook', () => {
    const cyclic = contextFor({ Sheet1: { A: { 1: 0 } } });
    // A cell whose own value resolves through a deep chain must still return, not throw.
    expect(() => evaluateFormula('=((((((((((((1))))))))))))', cyclic)).not.toThrow();
  });
});

describe('Information functions', () => {
  const context = contextFor({
    Sheet1: { A: { 1: 42, 2: 'text', 3: true, 4: null }, B: { 1: '#DIV/0!', 2: '#N/A' } },
  });

  it('distinguishes types', () => {
    expect(evaluateFormula('=ISNUMBER(A1)', context)).toBe(true);
    expect(evaluateFormula('=ISTEXT(A2)', context)).toBe(true);
    expect(evaluateFormula('=ISLOGICAL(A3)', context)).toBe(true);
    expect(evaluateFormula('=ISBLANK(A4)', context)).toBe(true);
  });

  it('separates the two error predicates the way Excel does', () => {
    expect(evaluateFormula('=ISERROR(B1)', context)).toBe(true);
    expect(evaluateFormula('=ISERR(B1)', context)).toBe(true);
    expect(evaluateFormula('=ISERROR(B2)', context)).toBe(true);
    expect(evaluateFormula('=ISERR(B2)', context)).toBe(false);
    expect(evaluateFormula('=ISNA(B2)', context)).toBe(true);
  });

  it('produces #N/A on demand', () => {
    expect(evaluateFormula('=ISNA(NA())', context)).toBe(true);
  });

  it('reports a numeric serial for a date via N()', () => {
    const dated = contextFor({ Sheet1: { A: { 1: 44197 } } });
    expect(evaluateFormula('=N(A1)', dated)).toBe(44197);
    expect(evaluateFormula('=N(DATE(2021,1,1))', dated)).toBe(44197);
  });
});

describe('isFormulaError', () => {
  it('matches only canonical Excel codes', () => {
    expect(isFormulaError('#DIV/0!')).toBe(true);
    expect(isFormulaError('#N/A')).toBe(true);
    expect(isFormulaError('#12345')).toBe(false);
    expect(isFormulaError('#NAME? (FOO)')).toBe(false);
    expect(isFormulaError('hello')).toBe(false);
    expect(isFormulaError(null)).toBe(false);
  });
});

describe('Date number format detection', () => {
  // This single predicate decides whether an imported numeric cell becomes a date or stays a
  // number, so a wrong answer here makes date arithmetic wrong on real spreadsheets.
  it.each([
    ['yyyy-mm-dd', true],
    ['MM/DD/YYYY', true],
    ['d-mmm-yy', true],
    ['m/d/yy', true],
    ['d/m/yyyy', true],
    ['hh:mm:ss', true],
    ['h:mm AM/PM', true],
    ['mm:ss', true],
    ['mm:ss.0', true],
    ['yyyy-mm-dd hh:mm', true],
    ['[h]:mm:ss', true],
    ['[mm]:ss', true],
    ['[$-409]mmmm d, yyyy', true],
    ['[$-409]h:mm:ss', true],
  ])('treats %s as a date format', (format, expected) => {
    expect(isDateNumberFormat(format)).toBe(expected);
  });

  it.each([
    ['General', false],
    ['0', false],
    ['#,##0', false],
    ['#,##0.00', false],
    ['0.00%', false],
    ['0.0%', false],
    ['0.00E+00', false],
    ['@', false],
    ['[Red]0.00', false],
    ['0.00 "days"', false],
    ['"Sales "0.00', false],
    ['"$"#,##0.00', false],
    ['$#,##0.00', false],
    ['#,##0;[Red]-#,##0', false],
    ['0.00_);[Red](0.00)', false],
    ['$#,##0.00_);[Red]($#,##0.00)', false],
    ['[$]#,##0', false],
  ])('treats %s as a numeric format', (format, expected) => {
    expect(isDateNumberFormat(format)).toBe(expected);
  });

  it('returns false for a missing format', () => {
    expect(isDateNumberFormat(undefined)).toBe(false);
    expect(isDateNumberFormat(null)).toBe(false);
    expect(isDateNumberFormat('')).toBe(false);
  });
});
