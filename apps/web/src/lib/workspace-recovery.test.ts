import { describe, expect, it, vi } from 'vitest';
import { createCell } from '@excel-agent/engine';
import { createIndexedDBRecoveryStore, type WorkspaceCheckpoint } from './workspace-recovery.js';

const checkpoint = (fileName = 'dates.xlsx'): WorkspaceCheckpoint => ({
  version: 1,
  fileName,
  activeSheetName: 'Dates',
  dateSystem: '1904',
  hasUserUploadedFile: true,
  savedAt: 1770000000000,
  workbook: {
    dateSystem: '1904',
    sheets: [
      {
        name: 'Dates',
        rows: [
          [
            { ...createCell(new Date('2024-02-29T00:00:00.000Z')), numberFormat: 'yyyy-mm-dd' },
            { ...createCell(null), type: 'formula', formula: '=A1+1' },
          ],
        ],
      },
    ],
  },
});

/** Small transaction-aware IndexedDB fake; writes only become visible when committed. */
function indexedDBFake() {
  let stored: unknown;
  let exists = false;
  let paused = false;
  let abortWrite = false;
  let completePending: (() => void) | undefined;
  const close = vi.fn();
  const database = {
    objectStoreNames: { contains: () => exists },
    createObjectStore: vi.fn(() => {
      exists = true;
    }),
    close,
    transaction: (_name: string, mode: string) => {
      const transaction: {
        oncomplete?: () => void;
        onabort?: () => void;
        error?: Error;
        objectStore: () => Record<string, (...args: unknown[]) => unknown>;
      } = {
        objectStore: () => ({
          get: () => request('get'),
          put: (value: unknown) => request('put', value),
          delete: () => request('delete'),
        }),
      };
      const request = (action: string, value?: unknown) => {
        const req: { result?: unknown; onsuccess?: () => void } = {};
        // Clone at dispatch, just as IndexedDB does (not at commit).
        const copy = action === 'put' ? structuredClone(value) : undefined;
        queueMicrotask(() => {
          req.result = action === 'get' ? structuredClone(stored) : undefined;
          req.onsuccess?.();
          const complete = () => {
            if (mode === 'readwrite' && abortWrite) {
              abortWrite = false;
              transaction.error = new Error('Quota exceeded');
              transaction.onabort?.();
            } else {
              if (action === 'put') stored = copy;
              if (action === 'delete') stored = undefined;
              transaction.oncomplete?.();
            }
          };
          if (paused) {
            paused = false;
            completePending = complete;
          } else queueMicrotask(complete);
        });
        return req;
      };
      return transaction;
    },
  };
  const factory = {
    open: vi.fn(() => {
      const request: {
        result: typeof database;
        onupgradeneeded?: () => void;
        onsuccess?: () => void;
      } = { result: database };
      queueMicrotask(() => {
        if (!exists) request.onupgradeneeded?.();
        request.onsuccess?.();
      });
      return request;
    }),
  } as unknown as IDBFactory;
  return {
    factory,
    close,
    pauseNext: () => {
      paused = true;
    },
    complete: () => {
      completePending?.();
      completePending = undefined;
    },
    hasPending: () => completePending !== undefined,
    abortNextWrite: () => {
      abortWrite = true;
    },
    seed: (value: unknown) => {
      stored = value;
    },
  };
}

describe('durable IndexedDB workspace checkpoint', () => {
  it('round-trips Dates, formulas, formats, filename and 1904 epoch without localStorage', async () => {
    const fake = indexedDBFake();
    const store = createIndexedDBRecoveryStore(() => fake.factory);
    expect(await store.load()).toBeNull();
    const original = checkpoint();
    await store.save(original);
    const restored = await store.load();
    expect(restored).toEqual(original);
    expect(restored?.workbook.sheets[0]?.rows[0]?.[0]?.value).toBeInstanceOf(Date);
    expect(restored?.workbook).not.toBe(original.workbook);
    expect(fake.close).toHaveBeenCalledTimes(3);
  });

  it('reports success only after commit and preserves the previous checkpoint on abort', async () => {
    const fake = indexedDBFake();
    const store = createIndexedDBRecoveryStore(() => fake.factory);
    await store.save(checkpoint());
    fake.pauseNext();
    fake.abortNextWrite();
    let saved = false;
    const writing = store.save(checkpoint('failed.xlsx')).then(() => {
      saved = true;
    });
    const rejected = expect(writing).rejects.toThrow('Quota exceeded');
    await vi.waitFor(() => expect(fake.hasPending()).toBe(true));
    expect(saved).toBe(false);
    fake.complete();
    await rejected;
    expect((await store.load())?.fileName).toBe('dates.xlsx');
  });

  it('orders privacy deletion after an in-flight save and can save again afterward', async () => {
    const fake = indexedDBFake();
    const store = createIndexedDBRecoveryStore(() => fake.factory);
    fake.pauseNext();
    const writing = store.save(checkpoint());
    const clearing = store.clear();
    await vi.waitFor(() => expect(fake.hasPending()).toBe(true));
    fake.complete();
    await Promise.all([writing, clearing]);
    expect(await store.load()).toBeNull();
    await store.save(checkpoint('new.xlsx'));
    expect((await store.load())?.fileName).toBe('new.xlsx');
  });

  it('makes unsupported storage and malformed checkpoints explicit failures', async () => {
    await expect(createIndexedDBRecoveryStore(() => undefined).load()).rejects.toThrow(
      'unavailable',
    );
    const fake = indexedDBFake();
    fake.seed({ version: 10, workbook: { sheets: [] } });
    await expect(createIndexedDBRecoveryStore(() => fake.factory).load()).rejects.toThrow(
      'could not be read',
    );
  });
});
