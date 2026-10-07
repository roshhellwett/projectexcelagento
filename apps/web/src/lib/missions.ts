import type { EvidenceItem, ExecutionPlan, ProposedAction } from '@excel-agent/agent';
import type { Preview, Workbook } from '@excel-agent/engine';

export type MissionStatus =
  | 'planning'
  | 'executing'
  | 'prepared'
  | 'analyzed'
  | 'applied'
  | 'undone'
  | 'failed'
  | 'stale'
  | 'cancelled'
  | 'interrupted';
export type MissionKind = 'action' | 'plan' | 'analysis';
export interface MissionReceipt {
  id: string;
  status: 'applied' | 'undone' | 'failed';
  createdAt: number;
  completedAt: number;
  operations: { name: string; affectedCells: number; warnings: string[] }[];
}
export interface MissionRecord {
  version: 1;
  id: string;
  kind: MissionKind;
  status: MissionStatus;
  title: string;
  request: string;
  answer?: string;
  stage?: string;
  createdAt: number;
  updatedAt: number;
  workbook: {
    generation: number;
    revision: number;
    fileName: string;
    sheetName: string;
    signature: string | null;
  };
  action?: ProposedAction;
  plan?: ExecutionPlan;
  preview?: Preview;
  evidence?: EvidenceItem[];
  receipt?: MissionReceipt;
  error?: string;
}
export interface MissionStore {
  list(): Promise<MissionRecord[]>;
  save(record: MissionRecord): Promise<void>;
  delete(id: string): Promise<void>;
  clear(): Promise<void>;
}
export const MAX_MISSIONS = 50;
const STATUSES: MissionStatus[] = [
  'planning',
  'executing',
  'prepared',
  'analyzed',
  'applied',
  'undone',
  'failed',
  'stale',
  'cancelled',
  'interrupted',
];
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown, max = 10000): value is string =>
  typeof value === 'string' && value.length <= max;
const count = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
const timestamp = (value: unknown): value is number => count(value) && value <= 8640000000000000;

/** Storage is not a trusted execution contract. Engine schemas are checked again on resume. */
export function validateMission(value: unknown): MissionRecord {
  if (
    !object(value) ||
    value.version !== 1 ||
    !text(value.id, 200) ||
    !value.id ||
    !STATUSES.includes(value.status as MissionStatus) ||
    !['action', 'plan', 'analysis'].includes(String(value.kind)) ||
    !text(value.title, 1000) ||
    !text(value.request) ||
    !timestamp(value.createdAt) ||
    !timestamp(value.updatedAt) ||
    !object(value.workbook) ||
    !count(value.workbook.generation) ||
    !count(value.workbook.revision) ||
    !text(value.workbook.fileName, 1000) ||
    !text(value.workbook.sheetName, 1000) ||
    !(
      value.workbook.signature === null ||
      (typeof value.workbook.signature === 'string' &&
        /^sha256:[a-f0-9]{64}$/.test(value.workbook.signature))
    )
  ) {
    throw new Error('Saved mission metadata is invalid. It cannot be resumed safely.');
  }
  if (value.stage !== undefined && !text(value.stage, 1000))
    throw new Error('Saved mission stage is invalid.');
  if (value.answer !== undefined && !text(value.answer, 20000))
    throw new Error('Saved mission answer is invalid.');
  if (value.error !== undefined && !text(value.error, 20000))
    throw new Error('Saved mission error is invalid.');
  if (
    value.action !== undefined &&
    (!object(value.action) ||
      !text(value.action.name, 100) ||
      !object(value.action.args) ||
      !text(value.action.explanation) ||
      !['format', 'transform', 'filter', 'columns', 'structure'].includes(
        String(value.action.category),
      ))
  )
    throw new Error('Saved mission action is invalid.');
  if (
    value.plan !== undefined &&
    (!object(value.plan) ||
      !text(value.plan.id, 200) ||
      !text(value.plan.title, 1000) ||
      !text(value.plan.description) ||
      !['pending', 'applied', 'error'].includes(String(value.plan.status)) ||
      !Array.isArray(value.plan.steps) ||
      value.plan.steps.length < 1 ||
      value.plan.steps.length > 25 ||
      !value.plan.steps.every(
        (step) =>
          object(step) &&
          text(step.id, 200) &&
          text(step.operation, 100) &&
          object(step.args) &&
          text(step.description),
      ))
  )
    throw new Error('Saved mission plan is invalid.');
  if (value.status === 'prepared' && !value.action && !value.plan)
    throw new Error('A prepared mission needs a complete action or plan.');
  if (
    value.evidence !== undefined &&
    (!Array.isArray(value.evidence) ||
      value.evidence.length > 12 ||
      !value.evidence.every(
        (item) =>
          object(item) &&
          text(item.id, 500) &&
          text(item.title) &&
          text(item.source) &&
          ['statistic', 'aggregate', 'profile', 'inspection'].includes(String(item.kind)) &&
          Array.isArray(item.facts) &&
          item.facts.length <= 30 &&
          item.facts.every((fact) => object(fact) && text(fact.label) && text(fact.value)) &&
          (item.note === undefined || text(item.note)),
      ))
  )
    throw new Error('Saved mission evidence is invalid.');
  if (
    value.receipt !== undefined &&
    (!object(value.receipt) ||
      !text(value.receipt.id, 200) ||
      !['applied', 'undone', 'failed'].includes(String(value.receipt.status)) ||
      !timestamp(value.receipt.createdAt) ||
      !timestamp(value.receipt.completedAt) ||
      !Array.isArray(value.receipt.operations) ||
      value.receipt.operations.length > 25 ||
      !value.receipt.operations.every(
        (op) =>
          object(op) &&
          text(op.name, 100) &&
          count(op.affectedCells) &&
          Array.isArray(op.warnings) &&
          op.warnings.every((warning) => text(warning)),
      ))
  )
    throw new Error('Saved mission receipt is invalid.');
  // Pick known fields: never retain accidental provider configs, keys, or workbook snapshots.
  const record: MissionRecord = {
    version: 1,
    id: value.id,
    kind: value.kind as MissionKind,
    status: value.status as MissionStatus,
    title: value.title,
    request: value.request,
    createdAt: value.createdAt,
    updatedAt: value.updatedAt,
    workbook: {
      generation: value.workbook.generation,
      revision: value.workbook.revision,
      fileName: value.workbook.fileName,
      sheetName: value.workbook.sheetName,
      signature: value.workbook.signature as string | null,
    },
  };
  for (const key of [
    'action',
    'plan',
    'preview',
    'evidence',
    'receipt',
    'error',
    'answer',
    'stage',
  ] as const) {
    if (value[key] !== undefined) Object.assign(record, { [key]: value[key] });
  }
  return record;
}

/** Length-framed JSON + SHA-256 avoids delimiter and 32-bit collisions in resume authorization. */
export async function workbookSignature(workbook: Workbook): Promise<string> {
  if (!globalThis.crypto?.subtle)
    throw new Error(
      'Secure workbook verification is unavailable. This mission cannot be resumed after refresh.',
    );
  const canonical = JSON.stringify([
    workbook.dateSystem ?? '1900',
    workbook.sheets.map((sheet) => [
      sheet.name,
      sheet.rows.map((row) =>
        row.map((cell) => [
          cell.type,
          cell.value instanceof Date
            ? ['date', cell.value.getTime()]
            : typeof cell.value === 'number'
              ? ['number', Object.is(cell.value, -0) ? '-0' : String(cell.value)]
              : [typeof cell.value, cell.value],
          cell.formula ?? null,
          cell.numberFormat ?? null,
          cell.style ?? null,
        ]),
      ),
    ]),
  ]);
  const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(canonical));
  return `sha256:${Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
}
export async function missionMatchesWorkbook(
  mission: MissionRecord,
  workbook: Workbook,
): Promise<boolean> {
  return (
    mission.workbook.signature !== null &&
    mission.workbook.signature === (await workbookSignature(workbook))
  );
}

/** One transaction contains the write and retention trim; success means the full transaction committed. */
export function createMissionStore(
  getFactory: () => IDBFactory | undefined = () => globalThis.indexedDB,
): MissionStore {
  let queue: Promise<unknown> = Promise.resolve();
  const ordered = <T>(task: () => Promise<T>): Promise<T> => {
    const next = queue.then(task, task);
    queue = next.catch(() => undefined);
    return next;
  };
  const open = (): Promise<IDBDatabase> =>
    new Promise((resolve, reject) => {
      let settled = false;
      const fail = (error: unknown) => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          reject(error);
        }
      };
      const timer = setTimeout(() => fail(new Error('Mission storage did not respond.')), 8000);
      try {
        const factory = getFactory();
        if (!factory) throw new Error('Mission storage is unavailable in this browser.');
        const request = factory.open('excelagento-missions', 1);
        request.onupgradeneeded = () => {
          if (!request.result.objectStoreNames.contains('missions'))
            request.result.createObjectStore('missions', { keyPath: 'id' });
        };
        request.onerror = () => fail(request.error ?? new Error('Could not open mission storage.'));
        request.onblocked = () =>
          fail(new Error('Mission storage is blocked by another tab. Close that tab and retry.'));
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
    action: 'list' | 'save' | 'delete' | 'clear',
    record?: MissionRecord,
    id?: string,
  ): Promise<unknown> => {
    const database = await open();
    try {
      return await new Promise((resolve, reject) => {
        const tx = database.transaction('missions', action === 'list' ? 'readonly' : 'readwrite');
        const store = tx.objectStore('missions');
        let output: unknown;
        let failure: unknown;
        const timer = setTimeout(() => {
          failure = new Error('Mission storage transaction timed out.');
          tx.abort();
        }, 10000);
        tx.oncomplete = () => {
          clearTimeout(timer);
          resolve(output);
        };
        tx.onabort = () => {
          clearTimeout(timer);
          reject(failure ?? tx.error ?? new Error('Mission storage transaction aborted.'));
        };
        tx.onerror = () => {
          clearTimeout(timer);
          reject(failure ?? tx.error ?? new Error('Mission storage failed.'));
        };
        const request =
          action === 'list'
            ? store.getAll()
            : action === 'save'
              ? store.put(record)
              : action === 'delete'
                ? store.delete(id!)
                : store.clear();
        request.onsuccess = () => {
          output = request.result;
          if (action === 'save') {
            const reading = store.getAll();
            reading.onsuccess = () => {
              const records = reading.result as MissionRecord[];
              records.sort((a, b) => b.updatedAt - a.updatedAt || a.id.localeCompare(b.id));
              for (const old of records.slice(MAX_MISSIONS)) store.delete(old.id);
            };
          }
        };
      });
    } finally {
      database.close();
    }
  };
  return {
    list: () =>
      ordered(async () => {
        const values = (await transact('list')) as unknown[];
        return values
          .map(validateMission)
          .sort((a, b) => b.updatedAt - a.updatedAt || a.id.localeCompare(b.id))
          .slice(0, MAX_MISSIONS);
      }),
    save: (record) => {
      const snapshot = structuredClone(validateMission(record));
      return ordered(async () => {
        await transact('save', snapshot);
      });
    },
    delete: (id) =>
      ordered(async () => {
        await transact('delete', undefined, id);
      }),
    clear: () =>
      ordered(async () => {
        await transact('clear');
      }),
  };
}
export const missionStore = createMissionStore();
