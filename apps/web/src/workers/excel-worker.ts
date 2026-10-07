import * as XLSX from 'xlsx-js-style';
import { parseWorkbookBytes, workbookToXlsxBytes, MAX_IMPORTED_CELLS } from '../lib/workbook-io.js';

export { MAX_IMPORTED_CELLS };

function parseXlsxInternal(arrayBuffer: ArrayBuffer): ReturnType<typeof parseWorkbookBytes> {
  return parseWorkbookBytes(XLSX, arrayBuffer);
}

function exportXlsxInternal(workbook: Parameters<typeof workbookToXlsxBytes>[1]): Uint8Array {
  return workbookToXlsxBytes(XLSX, workbook);
}

// Listen to messages when running inside Web Worker
if (typeof self !== 'undefined' && 'addEventListener' in self) {
  self.addEventListener('message', (e: MessageEvent) => {
    const { id, type, payload } = e.data || {};
    try {
      if (type === 'PING') {
        self.postMessage({ id, type: 'PING_SUCCESS' });
      } else if (type === 'PARSE_XLSX') {
        // The report travels inside the payload so the client's generic request helper can
        // carry it without special-casing this one message type.
        self.postMessage({
          id,
          type: 'PARSE_XLSX_SUCCESS',
          payload: parseXlsxInternal(payload.arrayBuffer),
        });
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
