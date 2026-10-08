import { createCell, type Workbook } from '@excel-agent/engine';
import { parseXlsxWorker, exportXlsxWorker } from './worker-client.js';
import {
  parseWorkbookBytes,
  workbookToXlsxBytes,
  MAX_IMPORTED_CELLS,
  type ImportReport,
} from './workbook-io.js';

type XlsxModule = typeof import('xlsx-js-style');

export { MAX_IMPORTED_CELLS };
export type { ImportReport };

let xlsxLoad: Promise<XlsxModule> | null = null;

/**
 * Lazily load the heavy xlsx codec the first time a file is parsed or
 * exported. The spreadsheet code is one of the largest chunks in the bundle;
 * deferring it keeps the initial page load small for visitors who only
 * browse, while repeat calls reuse the memoized module.
 */
function loadXlsx(): Promise<XlsxModule> {
  xlsxLoad ??= loadSpreadsheetCodec();
  return xlsxLoad;
}

async function loadSpreadsheetCodec(): Promise<XlsxModule> {
  // Vitest runs without Vite's emitted public codec asset. Keep the test fallback explicit so the
  // production graph contains no static dependency on the monolithic package.
  if (import.meta.env.MODE === 'test') return import('xlsx-js-style');
  const module = (await import(
    /* @vite-ignore */ new URL(/* @vite-ignore */ './xlsx-codec.js', import.meta.url).href
  )) as { default?: XlsxModule } & XlsxModule;
  return module.default ?? module;
}

/**
 * Parse an uploaded workbook. Delegates to the same parser the worker uses, so the
 * main-thread fallback and the worker path can never disagree about what a file contains.
 */
export async function xlsxToWorkbook(
  arrayBuffer: ArrayBuffer,
): Promise<{ workbook: Workbook; report: ImportReport }> {
  // Offload to background Web Worker thread to keep the UI at 60 FPS
  const workerResult = await parseXlsxWorker(arrayBuffer);
  if (workerResult) return workerResult;

  const XLSX = await loadXlsx();
  return parseWorkbookBytes(XLSX, arrayBuffer);
}

/**
 * Serialize a workbook to a real .xlsx byte array. Delegates to the same writer the worker
 * uses, so the two paths cannot drift apart in how they treat dates, blanks, or formats.
 */
export async function workbookToXlsxBuffer(workbook: Workbook): Promise<Uint8Array> {
  // Offload to background Web Worker thread if available
  const workerResult = await exportXlsxWorker(workbook);
  if (workerResult) return workerResult;

  const XLSX = await loadXlsx();
  return workbookToXlsxBytes(XLSX, workbook);
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
  } catch (error) {
    console.error('Excel export failed:', error);
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
