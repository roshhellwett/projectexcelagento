import {
  createCell,
  type Cell,
  type CellValue,
  type Workbook,
  type Sheet,
} from '@excel-agent/engine';

type XlsxModule = typeof import('xlsx');

let xlsxLoad: Promise<XlsxModule> | null = null;

/**
 * Lazily load the heavy xlsx codec the first time a file is parsed or
 * exported. The spreadsheet code is one of the largest chunks in the bundle;
 * deferring it keeps the initial page load small for visitors who only
 * browse, while repeat calls reuse the memoized module.
 */
function loadXlsx(): Promise<XlsxModule> {
  xlsxLoad ??= import('xlsx');
  return xlsxLoad;
}

/** Guard against pathological sheets that would freeze the browser tab. */
export const MAX_IMPORTED_CELLS = 1_500_000;

function looksLikeZip(bytes: Uint8Array): boolean {
  return (
    bytes.length > 3 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

/**
 * Parse an uploaded workbook. Handles both real .xlsx/.xlsm archives (detected by
 * their ZIP magic bytes) and plain-text CSV/TSV files, and refuses sheets large
 * enough to hang the tab.
 */
export async function xlsxToWorkbook(arrayBuffer: ArrayBuffer): Promise<Workbook> {
  const XLSX = await loadXlsx();
  const bytes = new Uint8Array(arrayBuffer);
  const wb = looksLikeZip(bytes)
    ? XLSX.read(bytes, {
        type: 'array',
        cellDates: false,
        cellFormula: true,
        cellStyles: true,
      })
    : XLSX.read(new TextDecoder().decode(bytes), {
        type: 'string',
        cellDates: false,
        cellFormula: true,
      });

  const sheets: Sheet[] = wb.SheetNames.map((sheetName) => {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) {
      return { name: sheetName, rows: [] };
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const rowCount = range.e.r + 1;
    const colCount = range.e.c + 1;

    if (rowCount * colCount > MAX_IMPORTED_CELLS) {
      throw new Error(
        `Sheet "${sheetName}" contains ${rowCount * colCount} cells, above the ${MAX_IMPORTED_CELLS} safety limit.`,
      );
    }

    const rows: Cell[][] = [];
    for (let r = 0; r < rowCount; r += 1) {
      const rowCells: Cell[] = [];
      for (let c = 0; c < colCount; c += 1) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cellObj = ws[cellAddress];

        if (!cellObj) {
          rowCells.push(createCell(null));
          continue;
        }

        let val: CellValue = null;
        if (typeof cellObj.v === 'string') {
          val = cellObj.v;
        } else if (typeof cellObj.v === 'number') {
          val = cellObj.v;
        } else if (typeof cellObj.v === 'boolean') {
          val = cellObj.v;
        } else if (cellObj.v !== undefined && cellObj.v !== null) {
          val = String(cellObj.v);
        }

        const formula = typeof cellObj.f === 'string' ? cellObj.f : undefined;
        const numberFormat = typeof cellObj.z === 'string' ? cellObj.z : undefined;

        rowCells.push(createCell(val, { formula, numberFormat }));
      }
      rows.push(rowCells);
    }

    return {
      name: sheetName,
      rows,
    };
  });

  return {
    sheets: sheets.length > 0 ? sheets : [{ name: 'Sheet1', rows: [] }],
  };
}

/** Excel forbids these characters in sheet names and caps names at 31 characters. */
function sanitizeSheetName(name: string, used: Set<string>): string {
  const base =
    (name || 'Sheet')
      .replace(/[:\\/?*[\]]/g, ' ')
      .trim()
      .slice(0, 31) || 'Sheet';
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) {
    const tag = `_${suffix}`;
    candidate = `${base.slice(0, 31 - tag.length)}${tag}`;
    suffix += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

/**
 * Serialize a workbook to a real .xlsx byte array. Blank cells stay blank (not
 * empty strings), formulas keep their cached value, per-cell number formats are
 * reapplied, and sheet names are made Excel-legal and unique.
 */
export async function workbookToXlsxBuffer(workbook: Workbook): Promise<Uint8Array> {
  const XLSX = await loadXlsx();
  const wb = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  for (const sheet of workbook.sheets) {
    const aoa: unknown[][] = sheet.rows.map((row) =>
      row.map((cell) => {
        if (cell.formula) {
          return { f: cell.formula, v: cell.value ?? '' };
        }
        if (cell.value === null || cell.value === '') {
          return null;
        }
        return cell.value;
      }),
    );

    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    // Preserve per-cell number formats so dates and currencies survive the trip.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!cell.numberFormat) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const target = ws[address] as { z?: string } | undefined;
        if (target) target.z = cell.numberFormat;
      });
    });

    // Pin the used range so trailing blank columns/rows are not silently dropped.
    const rows = Math.max(1, sheet.rows.length);
    const columns = Math.max(1, ...sheet.rows.map((row) => row.length), 0);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows - 1, c: columns - 1 } });

    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheet.name, usedNames));
  }

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(out);
}

export async function downloadWorkbookAsXlsx(
  workbook: Workbook,
  filename = 'exported-data.xlsx',
): Promise<boolean> {
  try {
    const buffer = await workbookToXlsxBuffer(workbook);
    const blob = new Blob([buffer as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
    return true;
  } catch {
    return false;
  }
}

export function createSampleWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Orders & Deliveries',
        rows: [
          [
            createCell('Order ID'),
            createCell('Customer Name'),
            createCell('Order Date'),
            createCell('Region'),
            createCell('Amount (USD)'),
            createCell('Status'),
          ],
          [
            createCell('ORD-1001'),
            createCell('  john DOE  '),
            createCell('2026-03-01'),
            createCell('North America'),
            createCell(1420.5),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1002'),
            createCell('acme corp'),
            createCell('03/15/2026'),
            createCell('Europe'),
            createCell(840.0),
            createCell('Pending'),
          ],
          [
            createCell('ORD-1003'),
            createCell('Alice M. Smith'),
            createCell('24/03/2026'),
            createCell('North America'),
            createCell(2150.0),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1004'),
            createCell('  TECH CORP  '),
            createCell('2026/04/05'),
            createCell('Asia-Pacific'),
            createCell(430.75),
            createCell('Pending'),
          ],
          [
            createCell('ORD-1002'),
            createCell('acme corp'),
            createCell('03/15/2026'),
            createCell('Europe'),
            createCell(840.0),
            createCell('Pending'),
          ],
          [
            createCell('ORD-1005'),
            createCell('robert johnson'),
            createCell('2026-04-18'),
            createCell('South America'),
            createCell(3100.25),
            createCell('Processing'),
          ],
          [
            createCell('ORD-1006'),
            createCell(' GLOBAL LOGISTICS '),
            createCell('05-02-2026'),
            createCell('Europe'),
            createCell(950.0),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1007'),
            createCell('Elena Rostova'),
            createCell('2026-05-12'),
            createCell('Europe'),
            createCell(1840.0),
            createCell('Cancelled'),
          ],
          [
            createCell('ORD-1008'),
            createCell('David Zhang'),
            createCell('16/05/2026'),
            createCell('Asia-Pacific'),
            createCell(5400.0),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1009'),
            createCell('  apex holdings  '),
            createCell('06/01/2026'),
            createCell('North America'),
            createCell(720.5),
            createCell('Pending'),
          ],
        ],
      },
      {
        name: 'Summary & Metrics',
        rows: [
          [createCell('Metric'), createCell('Value'), createCell('Notes')],
          [createCell('Total Target'), createCell(25000), createCell('Q1 & Q2 combined')],
          [createCell('Active Regions'), createCell(4), createCell('Global coverage')],
          [createCell('Default Currency'), createCell('USD'), createCell('Base pricing')],
        ],
      },
    ],
  };
}
