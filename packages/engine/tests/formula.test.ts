import { describe, expect, it } from 'vitest';
import { evaluateFormula, tokenize } from '../src/formula/index.js';
import type { FormulaContext, FormulaValue } from '../src/formula/types.js';

describe('Formula Evaluator Engine', () => {
  const mockSheetData: Record<string, Record<string, Record<number, FormulaValue>>> = {
    Sheet1: {
      A: { 1: 10, 2: 20, 3: 30, 4: 'Apple', 5: 'Banana' },
      B: { 1: 100, 2: 200, 3: 300, 4: 5, 5: 15 },
      C: { 1: 'North', 2: 'South', 3: 'North', 4: 'West', 5: 'North' },
    },
    Sales: {
      A: { 1: 'P100', 2: 'P200', 3: 'P300' },
      B: { 1: 'Widget', 2: 'Gadget', 3: 'Doohickey' },
      C: { 1: 49.99, 2: 99.99, 3: 19.99 },
    },
  };

  const context: FormulaContext = {
    activeSheet: 'Sheet1',
    getCellValue(sheet, col, row) {
      return mockSheetData[sheet]?.[col]?.[row] ?? null;
    },
    getRangeValues(sheet, startCol, startRow, endCol, endRow) {
      const rows: FormulaValue[][] = [];
      const startColCode = startCol.charCodeAt(0);
      const endColCode = endCol.charCodeAt(0);

      for (let r = startRow; r <= endRow; r++) {
        const rowVals: FormulaValue[] = [];
        for (let c = startColCode; c <= endColCode; c++) {
          const colLetter = String.fromCharCode(c);
          rowVals.push(mockSheetData[sheet]?.[colLetter]?.[r] ?? null);
        }
        rows.push(rowVals);
      }
      return rows;
    },
  };

  describe('Basic Expressions & Arithmetic', () => {
    it('evaluates constant arithmetic expressions with correct precedence', () => {
      expect(evaluateFormula('=1 + 2 * 3', context)).toBe(7);
      expect(evaluateFormula('=(1 + 2) * 3', context)).toBe(9);
      expect(evaluateFormula('=10 - 4 / 2', context)).toBe(8);
      expect(evaluateFormula('=2 ^ 3', context)).toBe(8);
      expect(evaluateFormula('=10 % 3', context)).toBe(1);
      expect(evaluateFormula('=-5 + 10', context)).toBe(5);
    });

    it('evaluates comparison operators', () => {
      expect(evaluateFormula('=10 > 5', context)).toBe(true);
      expect(evaluateFormula('=5 >= 5', context)).toBe(true);
      expect(evaluateFormula('=10 < 5', context)).toBe(false);
      expect(evaluateFormula('=10 = 10', context)).toBe(true);
      expect(evaluateFormula('=10 <> 5', context)).toBe(true);
    });

    it('handles string concatenation with &', () => {
      expect(evaluateFormula('="Hello " & "World"', context)).toBe('Hello World');
    });

    it('handles division by zero gracefully', () => {
      expect(evaluateFormula('=10 / 0', context)).toBe('#DIV/0!');
    });
  });

  describe('Cell & Range References', () => {
    it('reads single cell references from active sheet', () => {
      expect(evaluateFormula('=A1 + B1', context)).toBe(110);
      expect(evaluateFormula('=A2 * 2', context)).toBe(40);
    });

    it('reads cross-sheet references', () => {
      expect(evaluateFormula('=Sales!C1 * 2', context)).toBeCloseTo(99.98);
      expect(evaluateFormula('=Sales!B2', context)).toBe('Gadget');
    });

    it('handles ranges in functions', () => {
      expect(evaluateFormula('=SUM(A1:A3)', context)).toBe(60);
      expect(evaluateFormula('=AVERAGE(B1:B3)', context)).toBe(200);
      expect(evaluateFormula('=MAX(A1:A3)', context)).toBe(30);
      expect(evaluateFormula('=MIN(A1:A3)', context)).toBe(10);
    });
  });

  describe('Mathematical & Statistical Functions', () => {
    it('evaluates SUM, AVERAGE, COUNT, COUNTA, COUNTBLANK', () => {
      expect(evaluateFormula('=SUM(A1:A3, 40, 50)', context)).toBe(150);
      expect(evaluateFormula('=AVERAGE(10, 20, 30)', context)).toBe(20);
      expect(evaluateFormula('=COUNT(A1:A5)', context)).toBe(3); // A4, A5 are strings
      expect(evaluateFormula('=COUNTA(A1:A5)', context)).toBe(5);
      expect(evaluateFormula('=COUNTBLANK(A1:A5)', context)).toBe(0);
    });

    it('evaluates ROUND, ROUNDUP, ROUNDDOWN, ABS, CEILING, FLOOR', () => {
      expect(evaluateFormula('=ROUND(123.456, 2)', context)).toBe(123.46);
      expect(evaluateFormula('=ROUNDUP(123.411, 2)', context)).toBe(123.42);
      expect(evaluateFormula('=ROUNDDOWN(123.499, 2)', context)).toBe(123.49);
      expect(evaluateFormula('=ABS(-42)', context)).toBe(42);
      expect(evaluateFormula('=CEILING(4.2, 1)', context)).toBe(5);
      expect(evaluateFormula('=FLOOR(4.9, 1)', context)).toBe(4);
      expect(evaluateFormula('=SQRT(16)', context)).toBe(4);
    });
  });

  describe('Logical Functions', () => {
    it('evaluates IF', () => {
      expect(evaluateFormula('=IF(A1 > 5, "High", "Low")', context)).toBe('High');
      expect(evaluateFormula('=IF(A1 < 5, "High", "Low")', context)).toBe('Low');
    });

    it('evaluates IFS', () => {
      expect(evaluateFormula('=IFS(A1 = 5, "Five", A1 = 10, "Ten", TRUE, "Other")', context)).toBe(
        'Ten',
      );
    });

    it('evaluates AND, OR, NOT', () => {
      expect(evaluateFormula('=AND(10 > 5, 20 > 10)', context)).toBe(true);
      expect(evaluateFormula('=AND(10 > 5, 5 > 10)', context)).toBe(false);
      expect(evaluateFormula('=OR(10 > 5, 5 > 10)', context)).toBe(true);
      expect(evaluateFormula('=NOT(10 > 5)', context)).toBe(false);
    });

    it('evaluates IFERROR', () => {
      expect(evaluateFormula('=IFERROR(10 / 0, "Error Occurred")', context)).toBe('Error Occurred');
      expect(evaluateFormula('=IFERROR(10 / 2, "Error Occurred")', context)).toBe(5);
    });
  });

  describe('Conditional Aggregate Functions (SUMIF, COUNTIF, SUMIFS, COUNTIFS)', () => {
    it('evaluates COUNTIF with numeric and string criteria', () => {
      expect(evaluateFormula('=COUNTIF(C1:C5, "North")', context)).toBe(3);
      expect(evaluateFormula('=COUNTIF(A1:A3, ">15")', context)).toBe(2);
    });

    it('evaluates SUMIF with criteria and sum_range', () => {
      // Where C is "North", sum B: B1 (100) + B3 (300) + B5 (15) = 415
      expect(evaluateFormula('=SUMIF(C1:C5, "North", B1:B5)', context)).toBe(415);
    });

    it('evaluates COUNTIFS and SUMIFS', () => {
      // COUNTIFS(C1:C5, "North", B1:B5, ">50") -> B1 (100) and B3 (300) = 2
      expect(evaluateFormula('=COUNTIFS(C1:C5, "North", B1:B5, ">50")', context)).toBe(2);

      // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2)
      // Sum B1:B5 where C is North and A is >= 20 -> row 3 (A=30, C=North, B=300)
      expect(evaluateFormula('=SUMIFS(B1:B5, C1:C5, "North", A1:A5, ">=20")', context)).toBe(300);
    });
  });

  describe('Lookup & Reference Functions (VLOOKUP, XLOOKUP, INDEX, MATCH)', () => {
    it('evaluates VLOOKUP exact match', () => {
      // Lookup 'P200' in Sales!A1:C3, return col 2 (Gadget)
      expect(evaluateFormula('=VLOOKUP("P200", Sales!A1:C3, 2, FALSE)', context)).toBe('Gadget');
      // Return price (col 3)
      expect(evaluateFormula('=VLOOKUP("P300", Sales!A1:C3, 3, FALSE)', context)).toBe(19.99);
      // Not found
      expect(evaluateFormula('=VLOOKUP("P999", Sales!A1:C3, 2, FALSE)', context)).toBe('#N/A');
    });

    it('evaluates INDEX and MATCH', () => {
      // Match 'Gadget' in Sales!B1:B3 -> index 2
      expect(evaluateFormula('=MATCH("Gadget", Sales!B1:B3, 0)', context)).toBe(2);
      // INDEX(Sales!C1:C3, 2) -> 99.99
      expect(evaluateFormula('=INDEX(Sales!C1:C3, 2)', context)).toBe(99.99);
    });

    it('evaluates XLOOKUP', () => {
      // XLOOKUP(lookup_value, lookup_array, return_array, if_not_found)
      expect(evaluateFormula('=XLOOKUP("P100", Sales!A1:A3, Sales!B1:B3)', context)).toBe('Widget');
      expect(
        evaluateFormula('=XLOOKUP("P999", Sales!A1:A3, Sales!B1:B3, "Not Found")', context),
      ).toBe('Not Found');
    });
  });

  describe('Text Functions', () => {
    it('evaluates CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN', () => {
      expect(evaluateFormula('=CONCAT("Super", " ", "Excel")', context)).toBe('Super Excel');
      expect(evaluateFormula('=TEXTJOIN("-", TRUE, "A", "B", "", "C")', context)).toBe('A-B-C');
      expect(evaluateFormula('=LEFT("Antigravity", 4)', context)).toBe('Anti');
      expect(evaluateFormula('=RIGHT("Antigravity", 7)', context)).toBe('gravity');
      expect(evaluateFormula('=MID("Antigravity", 5, 4)', context)).toBe('grav');
      expect(evaluateFormula('=LEN("ExcelAgento")', context)).toBe(11);
    });

    it('evaluates UPPER, LOWER, PROPER, TRIM, SUBSTITUTE, REPLACE', () => {
      expect(evaluateFormula('=UPPER("hello")', context)).toBe('HELLO');
      expect(evaluateFormula('=LOWER("HELLO")', context)).toBe('hello');
      expect(evaluateFormula('=PROPER("john doe")', context)).toBe('John Doe');
      expect(evaluateFormula('=TRIM("   spaced   out   ")', context)).toBe('spaced out');
      expect(evaluateFormula('=SUBSTITUTE("Banana", "a", "o")', context)).toBe('Bonono');
      expect(evaluateFormula('=REPLACE("Hello World", 7, 5, "Agent")', context)).toBe(
        'Hello Agent',
      );
    });
  });

  describe('Date Functions', () => {
    it('evaluates DATE, YEAR, MONTH, DAY', () => {
      expect(evaluateFormula('=DATE(2026, 10, 4)', context)).toBe('2026-10-04');
      expect(evaluateFormula('=YEAR("2026-10-04")', context)).toBe(2026);
      expect(evaluateFormula('=MONTH("2026-10-04")', context)).toBe(10);
      expect(evaluateFormula('=DAY("2026-10-04")', context)).toBe(4);
    });
  });

  describe('Tokenization Edge Cases', () => {
    it('tokenizes correctly without error', () => {
      const tokens = tokenize('=IF(A1 >= 10, "Yes!", "No")');
      expect(tokens.map((t) => t.type)).toEqual([
        'IDENT',
        'LPAREN',
        'CELL',
        'OP',
        'NUMBER',
        'COMMA',
        'STRING',
        'COMMA',
        'STRING',
        'RPAREN',
      ]);
    });
  });
});
