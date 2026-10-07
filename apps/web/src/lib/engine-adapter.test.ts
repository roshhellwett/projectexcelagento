import { describe, expect, it } from 'vitest';

import { applyOperation, createCell, type Workbook } from '@excel-agent/engine';

import { workbookToXlsxBuffer, xlsxToWorkbook } from './engine-adapter.js';
import {
  classifyWorkbookBytes,
  decodeTextBytes,
  detectDelimiter,
  parseCsvField,
  workbookFromCsvText,
  MAX_IMPORTED_CELLS,
} from './workbook-io.js';

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/** Runs a workbook through the real writer and the real reader. */
async function roundTrip(workbook: Workbook): Promise<Workbook> {
  const { workbook: restored } = await xlsxToWorkbook(
    toArrayBuffer(await workbookToXlsxBuffer(workbook)),
  );
  return restored;
}

describe('xlsx adapter', () => {
  it('round-trips values, formulas, blanks, and number formats', async () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Data',
          rows: [
            [createCell('Name'), createCell('Amount'), createCell('When')],
            [
              createCell('Ada'),
              createCell(12.5),
              createCell(new Date(Date.UTC(2026, 2, 1)), { numberFormat: 'yyyy-mm-dd' }),
            ],
            [createCell('Grace'), createCell(null), createCell(0, { formula: 'B3*2' })],
          ],
        },
      ],
    };

    const sheet = (await roundTrip(workbook)).sheets[0];
    expect(sheet?.name).toBe('Data');
    expect(sheet?.rows[1]?.[0]?.value).toBe('Ada');
    expect(sheet?.rows[1]?.[1]?.value).toBe(12.5);
    // A date-formatted cell comes back as a real Date, not a bare serial. That is what makes
    // `=C3-C2` a day count on a real spreadsheet instead of a meaningless subtraction.
    expect(sheet?.rows[1]?.[2]?.value).toEqual(new Date(Date.UTC(2026, 2, 1)));
    expect(sheet?.rows[1]?.[2]?.numberFormat).toBe('yyyy-mm-dd');
    // Blank cells stay blank rather than becoming empty strings.
    expect(sheet?.rows[2]?.[1]?.value).toBe(null);
    // Formulas are preserved.
    expect(sheet?.rows[2]?.[2]?.formula).toBe('B3*2');
  });

  it('round-trips workbook-native cell styles', async () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: 'Styled',
          rows: [
            [
              createCell('Header', {
                style: {
                  bold: true,
                  fillColor: '#FFF2CC',
                  fontColor: '#20342B',
                  horizontalAlignment: 'center',
                  verticalAlignment: 'middle',
                  wrapText: true,
                },
              }),
            ],
          ],
        },
      ],
    };
    const sheet = (await roundTrip(workbook)).sheets[0];

    // xlsx-js-style exposes the fill on read; font and alignment are emitted in the XLSX
    // styles table and are consumed by Excel/LibreOffice when opening the exported file.
    expect(sheet?.rows[0]?.[0]?.style).toEqual({ fillColor: '#FFF2CC' });
    const XLSX = await import('xlsx-js-style');
    const parsed = XLSX.read(await workbookToXlsxBuffer(workbook), {
      type: 'array',
      cellStyles: true,
    }) as unknown as {
      Styles?: {
        Fonts?: Array<{ bold?: number; color?: { rgb?: string } }>;
        CellXf?: Array<{
          alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
        }>;
      };
    };
    expect(parsed.Styles?.Fonts?.some((font) => font.bold === 1)).toBe(true);
    expect(
      parsed.Styles?.CellXf?.some(
        (xf) =>
          xf.alignment?.horizontal === 'center' &&
          xf.alignment.vertical === 'center' &&
          xf.alignment.wrapText === true,
      ),
    ).toBe(true);
  });

  it('gives a date with no format of its own one that still reads as a date', async () => {
    const sheet = (
      await roundTrip({
        sheets: [{ name: 'D', rows: [[createCell(new Date(Date.UTC(2021, 0, 1)))]] }],
      })
    ).sheets[0];
    expect(sheet?.rows[0]?.[0]?.value).toEqual(new Date(Date.UTC(2021, 0, 1)));
    expect(sheet?.rows[0]?.[0]?.numberFormat).toBe('yyyy-mm-dd');
  });

  it('leaves an unformatted number as a number', async () => {
    const sheet = (await roundTrip({ sheets: [{ name: 'N', rows: [[createCell(44197)]] }] }))
      .sheets[0];
    // No date format means it is just a number, so it must not be reinterpreted.
    expect(sheet?.rows[0]?.[0]?.value).toBe(44197);
  });

  it('preserves the 1904 epoch, including raw serials referenced by formulas', async () => {
    const restored = await roundTrip({
      dateSystem: '1904',
      sheets: [
        {
          name: 'Dates',
          rows: [
            [
              createCell(new Date(Date.UTC(2026, 0, 1))),
              createCell(44561),
              createCell(44561, { formula: 'B1', numberFormat: 'yyyy-mm-dd' }),
            ],
          ],
        },
      ],
    });
    expect(restored.dateSystem).toBe('1904');
    expect(restored.sheets[0]?.rows[0]?.[0]?.value).toEqual(new Date(Date.UTC(2026, 0, 1)));
    expect(restored.sheets[0]?.rows[0]?.[1]?.value).toBe(44561);
    expect(restored.sheets[0]?.rows[0]?.[2]?.formula).toBe('B1');
  });

  it('keeps formulas even when Excel has not supplied a cached result', async () => {
    const restored = await roundTrip({
      sheets: [{ name: 'Formulas', rows: [[createCell(null, { formula: 'SUM(B1:B2)' })]] }],
    });
    expect(restored.sheets[0]?.rows[0]?.[0]?.formula).toBe('SUM(B1:B2)');
  });

  it('imports an Excel error code as an error instead of a magnitude', async () => {
    const XLSX = await import('xlsx-js-style');
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, { '!ref': 'A1', A1: { t: 'e', v: 7, f: '1/0' } }, 'Errors');
    const { workbook } = await xlsxToWorkbook(
      XLSX.write(book, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer,
    );
    expect(workbook.sheets[0]?.rows[0]?.[0]?.value).toBe('#DIV/0!');
    expect(workbook.sheets[0]?.rows[0]?.[0]?.formula).toBe('1/0');
  });

  it('exports live formula caches after editing a dependency, including cross-sheet chains', async () => {
    const before: Workbook = {
      sheets: [
        { name: 'Data', rows: [[createCell(3), createCell(6, { formula: '=A1*2' })]] },
        { name: 'Summary', rows: [[createCell(7, { formula: "'Data'!B1+1" })]] },
      ],
    };
    const edited = applyOperation(before, 'set_cells', {
      sheet: 'Data',
      cells: [{ row: 1, column: 'A', value: 10 }],
    });
    expect(edited.ok).toBe(true);
    // The engine deliberately stores the imported cache; the writer must evaluate the current snapshot.
    expect(edited.workbook.sheets[0]?.rows[0]?.[1]?.value).toBe(6);
    const restored = await roundTrip(edited.workbook);
    expect(restored.sheets[0]?.rows[0]?.[1]).toMatchObject({ value: 20, formula: 'A1*2' });
    expect(restored.sheets[1]?.rows[0]?.[0]).toMatchObject({ value: 21, formula: "'Data'!B1+1" });
    expect(edited.workbook.sheets[0]?.rows[0]?.[1]?.value).toBe(6);
  });

  it('exports live boolean, text, and 1904 date formula results with their formats', async () => {
    const restored = await roundTrip({
      dateSystem: '1904',
      sheets: [
        {
          name: 'Data',
          rows: [
            [
              createCell(false, { formula: '1=1' }),
              createCell('stale', { formula: '"current"' }),
              createCell(1, { formula: 'DATE(2026,1,1)', numberFormat: 'yyyy-mm-dd' }),
            ],
          ],
        },
      ],
    });
    expect(restored.sheets[0]?.rows[0]?.[0]?.value).toBe(true);
    expect(restored.sheets[0]?.rows[0]?.[1]?.value).toBe('current');
    expect(restored.sheets[0]?.rows[0]?.[2]).toMatchObject({
      value: 44561,
      formula: 'DATE(2026,1,1)',
      numberFormat: 'yyyy-mm-dd',
    });
    expect(restored.dateSystem).toBe('1904');
  });

  it.each(['UNSUPPORTED(A2)', '1/0', '"#N/A"'])(
    'never exports stale or invented error caches for %s',
    async (formula) => {
      const workbook: Workbook = {
        sheets: [{ name: 'Data', rows: [[createCell(999, { formula })], [createCell(5)]] }],
      };
      const XLSX = await import('xlsx-js-style');
      const bytes = await workbookToXlsxBuffer(workbook);
      const raw = XLSX.read(bytes, { type: 'array' }).Sheets['Data']!['A1'] as {
        f?: string;
        v?: unknown;
      };
      expect(raw.f).toBe(formula);
      expect(raw.v).toBeUndefined();
      const { workbook: restored } = await xlsxToWorkbook(toArrayBuffer(bytes));
      expect(restored.sheets[0]?.rows[0]?.[0]).toMatchObject({ formula, value: null });
    },
  );

  it('writes real recalculation instructions into workbook.xml, not only an ignored API property', async () => {
    const XLSX = await import('xlsx-js-style');
    const bytes = await workbookToXlsxBuffer({
      sheets: [{ name: 'Data', rows: [[createCell(999, { formula: 'A2*2' })], [createCell(5)]] }],
    });
    const parsed = XLSX.read(bytes, { type: 'array' });
    expect((parsed.Workbook as unknown as { CalcPr?: unknown })?.CalcPr).toMatchObject({
      calcMode: 'auto',
      fullCalcOnLoad: '1',
      forceFullCalc: '1',
    });
  });

  it.each(['A:B', 'Data'.repeat(10), 'data'])(
    'rejects export renaming conflicts that could break sheet-qualified formulas (%s)',
    async (name) => {
      const workbook: Workbook = {
        sheets: [
          { name: 'Data', rows: [[createCell(1)]] },
          { name, rows: [[createCell(2)]] },
          { name: 'Summary', rows: [[createCell(2, { formula: `'${name}'!A1` })]] },
        ],
      };
      await expect(workbookToXlsxBuffer(workbook)).rejects.toThrow(/sheet.*name|renam/i);
      expect(workbook.sheets[1]?.name).toBe(name);
    },
  );

  it('preserves valid qualified references with spaces without renaming sheets', async () => {
    const restored = await roundTrip({
      sheets: [
        { name: 'Source Data', rows: [[createCell(9)]] },
        { name: 'Summary', rows: [[createCell(0, { formula: "'Source Data'!$A$1*2" })]] },
      ],
    });
    expect(restored.sheets[1]?.rows[0]?.[0]).toMatchObject({
      value: 18,
      formula: "'Source Data'!$A$1*2",
    });
  });

  it('writes Excel-legal, unique sheet names', async () => {
    const workbook: Workbook = {
      sheets: [
        { name: 'A:B[C]', rows: [[createCell('x')]] },
        { name: 'A:B[C]', rows: [[createCell('y')]] },
      ],
    };

    const restored = await roundTrip(workbook);
    expect(restored.sheets.map((sheet) => sheet.name)).toEqual(['A B C', 'A B C_2']);
  });

  it('keeps multi-sheet structure and trailing blanks', async () => {
    const workbook: Workbook = {
      sheets: [
        { name: 'One', rows: [[createCell('a'), createCell('b')]] },
        { name: 'Two', rows: [[createCell('c'), createCell(null)]] },
      ],
    };

    const restored = await roundTrip(workbook);
    expect(restored.sheets.map((sheet) => sheet.name)).toEqual(['One', 'Two']);
    expect(restored.sheets[0]?.rows[0]).toHaveLength(2);
    expect(restored.sheets[1]?.rows[0]?.[1]?.value).toBe(null);
  });

  it('reports what a file could not carry over instead of losing it silently', async () => {
    // Merged cells and a chart cannot be represented by this engine's sheet model. The point
    // of the report is that the user finds out from us rather than from their own boss.
    const XLSX = await import('xlsx-js-style');
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Quarterly Report'],
      ['Region', 'Revenue'],
      ['North', 100],
    ]);
    sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, 'Report');
    const bytes = XLSX.write(book, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;

    const { workbook: restored, report } = await xlsxToWorkbook(bytes);
    expect(restored.sheets[0]?.rows[0]?.[0]?.value).toBe('Quarterly Report');
    expect(report.dropped).toContain('merged cells');
  });

  it('exports a sheet wide enough to overflow a spread-based column count', async () => {
    // `Math.max(...rows.map(...))` used to throw past ~125k rows. A two-row sheet with many
    // columns exercises the same code path without the runtime cost.
    const wide = Array.from({ length: 200 }, (_, i) => createCell(`c${i}`));
    const sheet = (await roundTrip({ sheets: [{ name: 'W', rows: [wide] }] })).sheets[0];
    expect(sheet?.rows[0]?.[0]?.value).toBe('c0');
    expect(sheet?.rows[0]?.[199]?.value).toBe('c199');
  });
});

describe('byte classification', () => {
  it('decides the format from magic bytes rather than the file extension', () => {
    expect(classifyWorkbookBytes(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00]))).toBe('xlsx');
    expect(
      classifyWorkbookBytes(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])),
    ).toBe('xls');
    expect(classifyWorkbookBytes(new TextEncoder().encode('a,b\n1,2'))).toBe('csv');
  });
});

describe('text decoding', () => {
  it('honours a UTF-8 byte-order mark instead of leaking it into the first header', () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode('Name\nAda')]);
    const { text, encoding } = decodeTextBytes(bytes);
    expect(encoding).toBe('utf-8');
    // A BOM on the header breaks every header-keyed feature downstream.
    expect(text).toBe('Name\nAda');
  });

  it('decodes UTF-16 with a byte-order mark', () => {
    const utf16 = new Uint8Array([0xff, 0xfe, 0x41, 0x00, 0x64, 0x00]);
    expect(decodeTextBytes(utf16).text).toBe('Ad');
  });

  it('detects UTF-16 that carries no byte-order mark', () => {
    const bytes = new Uint8Array([0x41, 0x00, 0x64, 0x00, 0x61, 0x00]);
    const { text, encoding } = decodeTextBytes(bytes);
    expect(encoding).toBe('utf-16le');
    expect(text).toBe('Ada');
  });

  it('falls back to a single-byte encoding rather than producing replacement characters', () => {
    // 0xE9 is "e with acute" in latin-1. A non-fatal UTF-8 decode turns it into U+FFFD and
    // the original character is gone for good.
    const bytes = new Uint8Array([...new TextEncoder().encode('Ren'), 0xe9]);
    const { text, encoding } = decodeTextBytes(bytes);
    expect(encoding).toBe('windows-1252');
    expect(text).toBe('Ren\u00e9');
  });

  it('decodes plain ASCII as UTF-8', () => {
    expect(decodeTextBytes(new TextEncoder().encode('Name,Amount')).encoding).toBe('utf-8');
  });
});

describe('delimiter detection', () => {
  it('detects the delimiter that yields consistent columns', () => {
    expect(detectDelimiter('Name,Amount\nAda,12.5\nGrace,7\n')).toBe(',');
    expect(detectDelimiter('Name;Amount\nAda;12,5\nGrace;7\n')).toBe(';');
    expect(detectDelimiter('Name\tAmount\nAda\t12.5\n')).toBe('\t');
    expect(detectDelimiter('Name|Amount\nAda|12.5\n')).toBe('|');
  });

  it('does not split a European decimal comma into a second column', () => {
    expect(detectDelimiter('Name;Amount\nAda;12,5\nGrace;7,25\n')).toBe(';');
  });

  it('handles quoted fields containing the delimiter', () => {
    expect(detectDelimiter('Name,Address\n"Doe, John","1 Main St"\n"Kane, Sue","2 Oak Rd"\n')).toBe(
      ',',
    );
  });

  it('does not use quoted punctuation as evidence of a delimiter', () => {
    expect(detectDelimiter('"Doe, John"')).toBeNull();
    expect(detectDelimiter('"Doe, John"\n"Kane, Sue"')).toBeNull();
    expect(detectDelimiter('Name,Amount')).toBe(',');
  });
});

describe('CSV field typing', () => {
  it('converts unambiguous numbers', () => {
    expect(parseCsvField('12.5', ',')).toBe(12.5);
    expect(parseCsvField('42', ',')).toBe(42);
    expect(parseCsvField('-7', ',')).toBe(-7);
    expect(parseCsvField('0.123', ',')).toBe(0.123);
    expect(parseCsvField('12.345', ',')).toBe(12.345);
    expect(parseCsvField('1.25e3', ',')).toBe(1250);
  });

  it('preserves identifiers that begin with a zero', () => {
    // 00123 is a ZIP code or an account number. Turning it into 123 is unrecoverable corruption.
    expect(parseCsvField('00123', ',')).toBe('00123');
    expect(parseCsvField('007', ',')).toBe('007');
    expect(parseCsvField('0', ',')).toBe(0);
    expect(parseCsvField('9007199254740993', ',')).toBe('9007199254740993');
  });

  it('reads a European decimal comma when the delimiter is a semicolon', () => {
    expect(parseCsvField('12,5', ';')).toBe(12.5);
    expect(parseCsvField('7,25', ';')).toBe(7.25);
  });

  it('strips thousands separators that match the locale', () => {
    expect(parseCsvField('1.200,50', ';')).toBe(1200.5);
    expect(parseCsvField('1,200.50', ',')).toBe(1200.5);
  });

  it('handles currency, percent, and accounting negatives', () => {
    expect(parseCsvField('$1,200.50', ',')).toBe(1200.5);
    expect(parseCsvField('45%', ',')).toBeCloseTo(0.45);
    expect(parseCsvField('(450)', ',')).toBe(-450);
  });

  it('keeps text as text', () => {
    expect(parseCsvField('Ada', ',')).toBe('Ada');
    expect(parseCsvField('12abc', ',')).toBe('12abc');
    expect(parseCsvField('  ', ',')).toBe(null);
    expect(parseCsvField('=A1+B1', ',')).toBe('=A1+B1');
    expect(parseCsvField('  Ada  ', ',')).toBe('  Ada  ');
    expect(parseCsvField('null', ',')).toBe('null');
    expect(parseCsvField('"quoted text"', ',')).toBe('"quoted text"');
    expect(parseCsvField('1,2,3', ',')).toBe('1,2,3');
  });
});

describe('CSV import end to end', () => {
  it('enforces the cell limit before allocating a padded CSV workbook', () => {
    const wideHeader = Array.from({ length: 1500 }, () => 'h').join(',');
    const csv = `${wideHeader}\n${'x\n'.repeat(Math.floor(MAX_IMPORTED_CELLS / 1500) + 1)}`;
    expect(() => workbookFromCsvText(csv, ',', 'Large')).toThrow(/cell import limit/);
  });
  it('imports plain CSV uploads', async () => {
    const csv = 'Name,Amount\nAda,12.5\nGrace,7\n';
    const { workbook, report } = await xlsxToWorkbook(toArrayBuffer(new TextEncoder().encode(csv)));

    expect(workbook.sheets[0]?.rows[0]?.[0]?.value).toBe('Name');
    expect(workbook.sheets[0]?.rows[1]?.[0]?.value).toBe('Ada');
    expect(workbook.sheets[0]?.rows[1]?.[1]?.value).toBe(12.5);
    expect(report.source).toBe('csv');
    expect(report.encoding).toBe('utf-8');
  });

  it('keeps quoted commas inside a single field', async () => {
    const csv = 'Name,Address\n"Doe, John","1 Main St, Apt 2"\n';
    const { workbook } = await xlsxToWorkbook(toArrayBuffer(new TextEncoder().encode(csv)));
    expect(workbook.sheets[0]?.rows[1]?.[0]?.value).toBe('Doe, John');
    expect(workbook.sheets[0]?.rows[1]?.[1]?.value).toBe('1 Main St, Apt 2');
  });

  it('does not split a single-column file on currency grouping commas', async () => {
    const csv = 'Amount\n$1,200.50\n$99.00\n';
    const { workbook } = await xlsxToWorkbook(toArrayBuffer(new TextEncoder().encode(csv)));
    expect(workbook.sheets[0]?.rows[1]).toHaveLength(1);
    expect(workbook.sheets[0]?.rows[1]?.[0]?.value).toBe(1200.5);
  });

  it('parses Supabase-style log CSV with multiline quotes and JSON without breaking columns', async () => {
    const csv = [
      'id,date,method,pathname,status,timestamp,level,event_message,log_type',
      'row-1,2026-10-04,POST,/rest/v1/users,200,1728000000,info,"{\n  ""user_id"": 123,\n  ""action"": ""login""\n}",api',
      'row-2,2026-10-04,GET,/auth/v1/user,200,1728000060,info,"user session refreshed\nwith token",auth',
      'row-3,2026-10-04,DELETE,/rest/v1/items,204,1728000120,info,"deleted",api',
    ].join('\n');

    const { workbook, report } = await xlsxToWorkbook(toArrayBuffer(new TextEncoder().encode(csv)));
    const sheet = workbook.sheets[0]!;

    expect(report.delimiter).toBe(',');
    expect(sheet.rows).toHaveLength(4);
    // Header row has 9 columns
    expect(sheet.rows[0]).toHaveLength(9);
    expect(sheet.rows[0]?.[0]?.value).toBe('id');
    expect(sheet.rows[0]?.[3]?.value).toBe('pathname');
    expect(sheet.rows[0]?.[7]?.value).toBe('event_message');
    expect(sheet.rows[0]?.[8]?.value).toBe('log_type');

    // Data row 1 contains the multiline JSON inside column 7
    expect(sheet.rows[1]?.[0]?.value).toBe('row-1');
    expect(sheet.rows[1]?.[2]?.value).toBe('POST');
    expect(sheet.rows[1]?.[4]?.value).toBe(200);
    expect(sheet.rows[1]?.[7]?.value).toContain('"user_id": 123');
    expect(sheet.rows[1]?.[8]?.value).toBe('api');

    // Data row 2 contains multiline string inside column 7
    expect(sheet.rows[2]?.[0]?.value).toBe('row-2');
    expect(sheet.rows[2]?.[7]?.value).toBe('user session refreshed\nwith token');
    expect(sheet.rows[2]?.[8]?.value).toBe('auth');
  });

  it('normalizes ragged rows by padding with empty cells to the maximum column count', async () => {
    const csv = 'colA,colB,colC\n1,2\n3,4,5\n';
    const { workbook } = await xlsxToWorkbook(toArrayBuffer(new TextEncoder().encode(csv)));
    const sheet = workbook.sheets[0]!;

    expect(sheet.rows[0]).toHaveLength(3);
    expect(sheet.rows[1]).toHaveLength(3);
    expect(sheet.rows[1]?.[2]?.value).toBe(null);
    expect(sheet.rows[2]).toHaveLength(3);
    expect(sheet.rows[2]?.[2]?.value).toBe(5);
  });
});
