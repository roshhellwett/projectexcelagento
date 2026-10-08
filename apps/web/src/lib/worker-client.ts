import type { Workbook } from '@excel-agent/engine';
import type { ImportReport } from './workbook-io.js';

class WorkerRequestError extends Error {
  constructor(
    message: string,
    readonly kind: 'unavailable' | 'crash' | 'timeout' | 'execution',
  ) {
    super(message);
    this.name = 'WorkerRequestError';
  }
}

interface PendingRequest {
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
  timer: ReturnType<typeof setTimeout>;
}

let workerInstance: Worker | null = null;
let workerSupported: boolean | null = null;
const pendingRequests = new Map<string, PendingRequest>();

function disposeWorker(error?: WorkerRequestError): void {
  const worker = workerInstance;
  workerInstance = null;
  worker?.terminate();
  if (!error) return;
  for (const request of pendingRequests.values()) {
    clearTimeout(request.timer);
    request.reject(error);
  }
  pendingRequests.clear();
}

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
        pending.reject(new WorkerRequestError(error || 'Worker execution failed', 'execution'));
      } else {
        pending.resolve(payload);
      }
    };

    workerInstance.onerror = (e) => {
      // In case of fatal worker script error, reject all pending requests
      const err = new WorkerRequestError(e.message || 'Worker thread crashed', 'crash');
      disposeWorker(err);
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
    return Promise.reject(new WorkerRequestError('Web Workers not available', 'unavailable'));
  }

  const id = `req_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;

  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      const error = new WorkerRequestError(
        `Worker request ${type} timed out after ${timeoutMs}ms`,
        'timeout',
      );
      // A timed-out worker may still be parsing or serializing a large workbook. Terminate it so a
      // failed request cannot continue consuming CPU in the background or race a later retry.
      disposeWorker(error);
      reject(error);
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
export async function parseXlsxWorker(
  arrayBuffer: ArrayBuffer,
): Promise<{ workbook: Workbook; report: ImportReport } | null> {
  if (!isWorkerSupported()) return null;
  try {
    return await postToWorker<{ workbook: Workbook; report: ImportReport }>('PARSE_XLSX', {
      arrayBuffer,
    });
  } catch (err) {
    if (err instanceof WorkerRequestError && (err.kind === 'unavailable' || err.kind === 'crash')) {
      console.warn('Worker unavailable, falling back to main thread:', err);
      return null;
    }
    // A genuine parse failure or timeout must surface, not silently double-parse.
    throw err;
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
    if (err instanceof WorkerRequestError && (err.kind === 'unavailable' || err.kind === 'crash')) {
      console.warn('Worker unavailable, falling back to main thread:', err);
      return null;
    }
    throw err;
  }
}
