import * as XLSX from 'xlsx';
import {
  createCell,
  type Cell,
  type CellValue,
  type Workbook,
  type Sheet,
} from '@excel-agent/engine';

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

function parseXlsxInternal(arrayBuffer: ArrayBuffer): Workbook {
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

function exportXlsxInternal(workbook: Workbook): Uint8Array {
  const wb = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  for (const sheet of workbook.sheets) {
    const aoa: unknown[][] = sheet.rows.map((row) =>
      row.map((cell) => {
        if (cell.value === null || cell.value === '') {
          return null;
        }
        return cell.value;
      }),
    );

    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    // Write formula cells explicitly: aoa_to_sheet does not understand {f, v} objects.
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!cell.formula) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const cached = cell.value;
        const type =
          typeof cached === 'number'
            ? 'n'
            : typeof cached === 'boolean'
              ? 'b'
              : cached instanceof Date
                ? 'd'
                : 'str';
        (ws as Record<string, unknown>)[address] = {
          t: type,
          f: cell.formula.replace(/^=/, ''),
          v: cached ?? null,
        };
      });
    });

    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!cell.numberFormat) return;
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const target = ws[address] as { z?: string } | undefined;
        if (target) target.z = cell.numberFormat;
      });
    });

    const rows = Math.max(1, sheet.rows.length);
    const columns = Math.max(1, ...sheet.rows.map((row) => row.length), 0);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows - 1, c: columns - 1 } });

    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheet.name, usedNames));
  }

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(out);
}

// Listen to messages when running inside Web Worker
if (typeof self !== 'undefined' && 'addEventListener' in self) {
  self.addEventListener('message', (e: MessageEvent) => {
    const { id, type, payload } = e.data || {};
    try {
      if (type === 'PING') {
        self.postMessage({ id, type: 'PING_SUCCESS' });
      } else if (type === 'PARSE_XLSX') {
        const workbook = parseXlsxInternal(payload.arrayBuffer);
        self.postMessage({ id, type: 'PARSE_XLSX_SUCCESS', payload: workbook });
      } else if (type === 'EXPORT_XLSX') {
        const buffer = exportXlsxInternal(payload.workbook);
        (
          self as unknown as { postMessage: (message: unknown, transfer: Transferable[]) => void }
        ).postMessage({ id, type: 'EXPORT_XLSX_SUCCESS', payload: buffer }, [buffer.buffer]);
      } else {
        self.postMessage({ id, type: 'ERROR', error: `Unknown worker message type: ${type}` });
      }
    } catch (err) {
      self.postMessage({
        id,
        type: 'ERROR',
        error: err instanceof Error ? err.message : String(err),
      });
    }
  });
}
