import { describe, expect, it } from 'vitest';

import { createCell, type Workbook } from '@excel-agent/engine';

import { workbookToXlsxBuffer, xlsxToWorkbook } from './engine-adapter.js';

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('xlsx adapter', () => {
  it('round-trips values, formulas, blanks, and number formats', () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Data',
          rows: [
            [createCell('Name'), createCell('Amount'), createCell('When')],
            [
              createCell('Ada'),
              createCell(12.5),
              createCell(new Date(Date.UTC(2026, 2, 1)), { numberFormat: 'YYYY-MM-DD' }),
            ],
            [createCell('Grace'), createCell(null), createCell(0, { formula: 'B3*2' })],
          ],
        },
      ],
    };

    const restored = xlsxToWorkbook(toArrayBuffer(workbookToXlsxBuffer(workbook)));
    const sheet = restored.sheets[0];
    expect(sheet?.name).toBe('Data');
    expect(sheet?.rows[1]?.[0]?.value).toBe('Ada');
    expect(sheet?.rows[1]?.[1]?.value).toBe(12.5);
    // Dates survive as Excel serials plus the original number format.
    expect(sheet?.rows[1]?.[2]?.numberFormat).toBe('YYYY-MM-DD');
    expect(typeof sheet?.rows[1]?.[2]?.value).toBe('number');
    // Blank cells stay blank rather than becoming empty strings.
    expect(sheet?.rows[2]?.[1]?.value).toBe(null);
    // Formulas are preserved.
    expect(sheet?.rows[2]?.[2]?.formula).toBe('B3*2');
  });

  it('writes Excel-legal, unique sheet names', () => {
    const workbook: Workbook = {
      sheets: [
        { name: 'A:B[C]', rows: [[createCell('x')]] },
        { name: 'A:B[C]', rows: [[createCell('y')]] },
      ],
    };

    const restored = xlsxToWorkbook(toArrayBuffer(workbookToXlsxBuffer(workbook)));
    expect(restored.sheets.map((sheet) => sheet.name)).toEqual(['A B C', 'A B C_2']);
  });

  it('imports plain CSV uploads by sniffing the ZIP magic bytes', () => {
    const csv = 'Name,Amount\nAda,12.5\nGrace,7\n';
    const bytes = new TextEncoder().encode(csv);
    const restored = xlsxToWorkbook(toArrayBuffer(bytes));

    expect(restored.sheets[0]?.rows[0]?.[0]?.value).toBe('Name');
    expect(restored.sheets[0]?.rows[1]?.[0]?.value).toBe('Ada');
    expect(String(restored.sheets[0]?.rows[1]?.[1]?.value)).toBe('12.5');
  });

  it('keeps multi-sheet structure and trailing blanks', () => {
    const workbook: Workbook = {
      sheets: [
        { name: 'One', rows: [[createCell('a'), createCell('b')]] },
        { name: 'Two', rows: [[createCell('c'), createCell(null)]] },
      ],
    };

    const restored = xlsxToWorkbook(toArrayBuffer(workbookToXlsxBuffer(workbook)));
    expect(restored.sheets.map((sheet) => sheet.name)).toEqual(['One', 'Two']);
    expect(restored.sheets[0]?.rows[0]).toHaveLength(2);
    expect(restored.sheets[1]?.rows[0]?.[1]?.value).toBe(null);
  });
});
