import { parseWorkbookBytes, workbookToXlsxBytes, MAX_IMPORTED_CELLS } from '../lib/workbook-io.js';

type XlsxModule = typeof import('xlsx-js-style');

export { MAX_IMPORTED_CELLS };

let xlsxLoad: Promise<XlsxModule> | null = null;

function loadXlsx(): Promise<XlsxModule> {
  xlsxLoad ??= import(
    /* @vite-ignore */ new URL(/* @vite-ignore */ __XLSX_CODEC_ASSET__, import.meta.url).href
  ).then((module) => {
    const loaded = module as { default?: XlsxModule } & XlsxModule;
    return loaded.default ?? loaded;
  });
  return xlsxLoad;
}

async function parseXlsxInternal(
  arrayBuffer: ArrayBuffer,
): Promise<ReturnType<typeof parseWorkbookBytes>> {
  const XLSX = await loadXlsx();
  return parseWorkbookBytes(XLSX, arrayBuffer);
}

async function exportXlsxInternal(
  workbook: Parameters<typeof workbookToXlsxBytes>[1],
): Promise<Uint8Array> {
  const XLSX = await loadXlsx();
  return workbookToXlsxBytes(XLSX, workbook);
}

// Listen to messages when running inside Web Worker
if (typeof self !== 'undefined' && 'addEventListener' in self) {
  self.addEventListener('message', async (e: MessageEvent) => {
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
          payload: await parseXlsxInternal(payload.arrayBuffer),
        });
      } else if (type === 'EXPORT_XLSX') {
        const buffer = await exportXlsxInternal(payload.workbook);
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
