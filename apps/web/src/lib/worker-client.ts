import type { Workbook, FormulaValue } from '@excel-agent/engine';

interface PendingRequest {
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
  timer: ReturnType<typeof setTimeout>;
}

let workerInstance: Worker | null = null;
let workerSupported: boolean | null = null;
const pendingRequests = new Map<string, PendingRequest>();

function isWorkerSupported(): boolean {
  if (workerSupported !== null) return workerSupported;
  try {
    workerSupported =
      typeof window !== 'undefined' &&
      typeof window.Worker !== 'undefined' &&
      typeof Blob !== 'undefined';
  } catch {
    workerSupported = false;
  }
  return workerSupported;
}

function getWorker(): Worker | null {
  if (!isWorkerSupported()) return null;
  if (workerInstance) return workerInstance;

  try {
    workerInstance = new Worker(new URL('../workers/excel-worker.ts', import.meta.url), {
      type: 'module',
    });

    workerInstance.onmessage = (e: MessageEvent) => {
      const { id, type, payload, error } = e.data || {};
      const pending = pendingRequests.get(id);
      if (!pending) return;

      clearTimeout(pending.timer);
      pendingRequests.delete(id);

      if (type === 'ERROR') {
        pending.reject(new Error(error || 'Worker execution failed'));
      } else {
        pending.resolve(payload);
      }
    };

    workerInstance.onerror = (e) => {
      // In case of fatal worker script error, reject all pending requests
      const err = new Error(e.message || 'Worker thread crashed');
      for (const req of pendingRequests.values()) {
        clearTimeout(req.timer);
        req.reject(err);
      }
      pendingRequests.clear();
      workerInstance = null;
    };

    return workerInstance;
  } catch {
    workerSupported = false;
    return null;
  }
}

function postToWorker<T>(type: string, payload?: unknown, timeoutMs = 60000): Promise<T> {
  const worker = getWorker();
  if (!worker) {
    return Promise.reject(new Error('Web Workers not available in this environment'));
  }

  const id = `req_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;

  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      reject(new Error(`Worker request ${type} timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    pendingRequests.set(id, { resolve: resolve as (value: unknown) => void, reject, timer });

    try {
      worker.postMessage({ id, type, payload });
    } catch (err) {
      clearTimeout(timer);
      pendingRequests.delete(id);
      reject(err);
    }
  });
}

/**
 * Offloads Excel parsing to a Web Worker thread.
 * Returns null if Web Workers are unavailable, allowing fallback.
 */
export async function parseXlsxWorker(arrayBuffer: ArrayBuffer): Promise<Workbook | null> {
  if (!isWorkerSupported()) return null;
  try {
    return await postToWorker<Workbook>('PARSE_XLSX', { arrayBuffer });
  } catch (err) {
    console.warn('Worker parsing failed, falling back to main thread:', err);
    return null;
  }
}

/**
 * Offloads Excel workbook serialization to a Web Worker thread.
 * Returns null if Web Workers are unavailable, allowing fallback.
 */
export async function exportXlsxWorker(workbook: Workbook): Promise<Uint8Array | null> {
  if (!isWorkerSupported()) return null;
  try {
    return await postToWorker<Uint8Array>('EXPORT_XLSX', { workbook });
  } catch (err) {
    console.warn('Worker export failed, falling back to main thread:', err);
    return null;
  }
}

/**
 * Evaluates Excel formula on the worker thread.
 */
export async function evaluateFormulaWorker(
  formula: string,
  activeSheet: string,
  sheetData: Record<string, Record<string, Record<number, FormulaValue>>>,
): Promise<FormulaValue | null> {
  if (!isWorkerSupported()) return null;
  try {
    return await postToWorker<FormulaValue>('EVALUATE_FORMULA', {
      formula,
      activeSheet,
      sheetData,
    });
  } catch {
    return null;
  }
}
