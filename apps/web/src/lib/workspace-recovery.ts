import type { DateSystem, Workbook } from '@excel-agent/engine';
import { MAX_EXCEL_COLUMNS, MAX_EXCEL_ROWS, MAX_IMPORTED_CELLS } from './workbook-io.js';

/** Only the document is checkpointed: no keys, conversations, selections or undo history. */
export interface WorkspaceCheckpoint {
  version: 1;
  workbook: Workbook;
  fileName: string;
  activeSheetName: string;
  dateSystem: DateSystem;
  hasUserUploadedFile: boolean;
  savedAt: number;
}

/** Injectable seam for integration tests and environments without IndexedDB. */
export interface WorkspaceRecoveryStore {
  load(): Promise<WorkspaceCheckpoint | null>;
  save(checkpoint: WorkspaceCheckpoint): Promise<void>;
  clear(): Promise<void>;
}

export type CheckpointStatus =
  'checking' | 'idle' | 'recovery' | 'saving' | 'saved' | 'unavailable' | 'off';

const DATABASE = 'excelagento-workspace';
const STORE = 'checkpoints';
const KEY = 'latest';
const MAX_CHECKPOINT_CELLS = MAX_IMPORTED_CELLS;
const MAX_FORMULA_LENGTH = 8_192;

function validateCheckpoint(value: unknown): WorkspaceCheckpoint {
  const checkpoint = value as Partial<WorkspaceCheckpoint> | null;
  if (
    !checkpoint ||
    checkpoint.version !== 1 ||
    typeof checkpoint.fileName !== 'string' ||
    typeof checkpoint.activeSheetName !== 'string' ||
    (checkpoint.dateSystem !== '1900' && checkpoint.dateSystem !== '1904') ||
    typeof checkpoint.hasUserUploadedFile !== 'boolean' ||
    typeof checkpoint.savedAt !== 'number' ||
    !Number.isFinite(checkpoint.savedAt) ||
    checkpoint.savedAt < 0 ||
    !checkpoint.workbook ||
    !Array.isArray(checkpoint.workbook.sheets) ||
    checkpoint.workbook.sheets.length === 0
  ) {
    throw new Error(
      'The local checkpoint could not be read. Export your current workbook or clear the stored checkpoint.',
    );
  }
  if (
    checkpoint.workbook.dateSystem !== undefined &&
    checkpoint.workbook.dateSystem !== checkpoint.dateSystem
  ) {
    throw new Error('The local checkpoint contains an invalid date system.');
  }
  let cellCount = 0;
  const sheetNames = new Set<string>();
  for (const sheet of checkpoint.workbook.sheets) {
    if (
      !sheet ||
      typeof sheet.name !== 'string' ||
      !sheet.name.trim() ||
      sheetNames.has(sheet.name.toLowerCase()) ||
      !Array.isArray(sheet.rows) ||
      sheet.rows.length > MAX_EXCEL_ROWS
    ) {
      throw new Error('The local checkpoint contains an invalid worksheet.');
    }
    sheetNames.add(sheet.name.toLowerCase());
    for (const row of sheet.rows) {
      cellCount += Array.isArray(row) ? row.length : 0;
      if (cellCount > MAX_CHECKPOINT_CELLS) {
        throw new Error('The local checkpoint exceeds the workbook cell safety limit.');
      }
      if (
        !Array.isArray(row) ||
        row.length > MAX_EXCEL_COLUMNS ||
        Array.from(row).some((cell) => {
          if (
            !cell ||
            !['blank', 'string', 'number', 'boolean', 'date', 'formula'].includes(cell.type)
          )
            return true;
          const value = cell.value;
          const style = cell.style as
            | {
                bold?: unknown;
                italic?: unknown;
                underline?: unknown;
                fillColor?: unknown;
                fontColor?: unknown;
                horizontalAlignment?: unknown;
                verticalAlignment?: unknown;
                wrapText?: unknown;
              }
            | undefined;
          const validStyle =
            style === undefined ||
            (style !== null &&
              typeof style === 'object' &&
              !Array.isArray(style) &&
              (style.bold === undefined || typeof style.bold === 'boolean') &&
              (style.italic === undefined || typeof style.italic === 'boolean') &&
              (style.underline === undefined || typeof style.underline === 'boolean') &&
              (style.fillColor === undefined || /^#[0-9a-f]{6}$/i.test(String(style.fillColor))) &&
              (style.fontColor === undefined || /^#[0-9a-f]{6}$/i.test(String(style.fontColor))) &&
              (style.horizontalAlignment === undefined ||
                ['left', 'center', 'right'].includes(String(style.horizontalAlignment))) &&
              (style.verticalAlignment === undefined ||
                ['top', 'middle', 'bottom'].includes(String(style.verticalAlignment))) &&
              (style.wrapText === undefined || typeof style.wrapText === 'boolean'));
          const formula = (cell as { formula?: unknown }).formula;
          const numberFormat = (cell as { numberFormat?: unknown }).numberFormat;
          const validFormula =
            cell.type === 'formula'
              ? typeof formula === 'string' && formula.length <= MAX_FORMULA_LENGTH
              : formula === undefined;
          const validNumberFormat =
            numberFormat === undefined ||
            (typeof numberFormat === 'string' && numberFormat.length <= 255);
          return (
            !validStyle ||
            !validFormula ||
            !validNumberFormat ||
            !(
              value === null ||
              typeof value === 'string' ||
              typeof value === 'boolean' ||
              (typeof value === 'number' && Number.isFinite(value)) ||
              (value instanceof Date && Number.isFinite(value.getTime()))
            )
          );
        })
      )
        throw new Error('The local checkpoint contains invalid cells.');
    }
  }
  return {
    ...(checkpoint as WorkspaceCheckpoint),
    workbook: { ...checkpoint.workbook, dateSystem: checkpoint.dateSystem },
  };
}

/**
 * IndexedDB uses structured cloning, preserving Date cells and workbook metadata without JSON
 * or workbook-sized localStorage strings. Success means the transaction committed, not just
 * that its put request ran. Operations are ordered so a privacy clear cannot be overtaken by
 * an earlier save. A failed write leaves the previous successful checkpoint intact.
 */
export function createIndexedDBRecoveryStore(
  getFactory: () => IDBFactory | undefined = () => globalThis.indexedDB,
): WorkspaceRecoveryStore {
  let queue: Promise<unknown> = Promise.resolve();
  const ordered = <T>(task: () => Promise<T>): Promise<T> => {
    const result = queue.then(task, task);
    queue = result.catch(() => undefined);
    return result;
  };

  const open = (): Promise<IDBDatabase> =>
    new Promise((resolve, reject) => {
      let settled = false;
      let request: IDBOpenDBRequest;
      const fail = (error: unknown) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(error);
      };
      const timer = setTimeout(
        () => fail(new Error('Local checkpoint storage did not respond.')),
        8000,
      );
      try {
        const factory = getFactory();
        if (!factory) throw new Error('Local checkpoint storage is unavailable in this browser.');
        request = factory.open(DATABASE, 1);
        request.onupgradeneeded = () => {
          if (!request.result.objectStoreNames.contains(STORE))
            request.result.createObjectStore(STORE);
        };
        request.onerror = () =>
          fail(request.error ?? new Error('Could not open local checkpoint storage.'));
        request.onblocked = () =>
          fail(
            new Error('Local checkpoint storage is blocked by another tab. Close it and retry.'),
          );
        request.onsuccess = () => {
          if (settled) {
            request.result.close();
            return;
          }
          settled = true;
          clearTimeout(timer);
          request.result.onversionchange = () => request.result.close();
          resolve(request.result);
        };
      } catch (error) {
        fail(error);
      }
    });

  const transact = async (
    mode: IDBTransactionMode,
    action: 'load' | 'save' | 'clear',
    checkpoint?: WorkspaceCheckpoint,
  ): Promise<unknown> => {
    const database = await open();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = database.transaction(STORE, mode);
        const store = transaction.objectStore(STORE);
        const request =
          action === 'load'
            ? store.get(KEY)
            : action === 'clear'
              ? store.delete(KEY)
              : store.put(checkpoint, KEY);
        let value: unknown;
        request.onsuccess = () => {
          value = request.result;
        };
        transaction.oncomplete = () => resolve(value);
        transaction.onerror = () =>
          reject(
            transaction.error ?? request.error ?? new Error('Local checkpoint storage failed.'),
          );
        transaction.onabort = () =>
          reject(transaction.error ?? request.error ?? new Error('The checkpoint was not saved.'));
      });
    } finally {
      database.close();
    }
  };

  return {
    load: () =>
      ordered(async () => {
        const value = await transact('readonly', 'load');
        return value === undefined ? null : validateCheckpoint(value);
      }),
    save: async (checkpoint) => {
      const snapshot = structuredClone(validateCheckpoint(checkpoint));
      await ordered(async () => {
        await transact('readwrite', 'save', snapshot);
      });
    },
    clear: () =>
      ordered(async () => {
        await transact('readwrite', 'clear');
      }),
  };
}

export const workspaceRecoveryStore = createIndexedDBRecoveryStore();
