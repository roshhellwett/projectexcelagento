import { describe, expect, it, vi, afterEach } from 'vitest';
import { IDBFactory, IDBObjectStore } from 'fake-indexeddb';
import { cloneWorkbook, createCell, type Workbook } from '@excel-agent/engine';
import {
  createMissionStore,
  MAX_MISSIONS,
  missionMatchesWorkbook,
  validateMission,
  workbookSignature,
  type MissionRecord,
} from './missions.js';

const workbook: Workbook = {
  sheets: [{ name: 'Sales', rows: [[createCell('Income')], [createCell(-300)]] }],
};
async function record(
  id = 'task-1',
  status: MissionRecord['status'] = 'prepared',
): Promise<MissionRecord> {
  return {
    version: 1,
    id,
    kind: 'action',
    status,
    title: 'Trim orders',
    request: 'trim column A',
    createdAt: 1,
    updatedAt: 1,
    workbook: {
      generation: 0,
      revision: 0,
      fileName: 'sales.xlsx',
      sheetName: 'Sales',
      signature: await workbookSignature(workbook),
    },
    action: {
      name: 'normalize_text',
      category: 'transform',
      explanation: 'Trim values',
      args: { sheet: 'Sales', columns: ['A'], trim: true },
    },
  };
}
afterEach(() => vi.restoreAllMocks());

describe('mission workbook authorization', () => {
  it('uses full content SHA-256, framing, formula, formatting, Dates, and epoch', async () => {
    const signature = await workbookSignature(workbook);
    expect(signature).toMatch(/^sha256:[a-f0-9]{64}$/);
    expect(await workbookSignature(cloneWorkbook(workbook))).toBe(signature);
    expect(await workbookSignature({ ...workbook, dateSystem: '1900' })).toBe(signature);
    for (const replacement of [
      createCell(10),
      createCell('-300'),
      createCell(-300, { formula: '1-301' }),
      createCell(-300, { numberFormat: '0.00' }),
      createCell(new Date('2026-01-01')),
    ]) {
      const changed = cloneWorkbook(workbook);
      changed.sheets[0]!.rows[1]![0] = replacement;
      expect(await missionMatchesWorkbook(await record(), changed)).toBe(false);
    }
    expect(await workbookSignature({ ...workbook, dateSystem: '1904' })).not.toBe(signature);
    const piped = { sheets: [{ name: 'Sales', rows: [[createCell('x|y')]] }] };
    const formula = { sheets: [{ name: 'Sales', rows: [[createCell('x', { formula: 'y' })]] }] };
    expect(await workbookSignature(piped)).not.toBe(await workbookSignature(formula));
  });
  it('rejects absent or legacy weak signatures for resume', async () => {
    const item = await record();
    item.workbook.signature = null;
    expect(await missionMatchesWorkbook(item, workbook)).toBe(false);
    expect(() =>
      validateMission({ ...item, workbook: { ...item.workbook, signature: '1234abcd' } }),
    ).toThrow(/invalid/i);
  });
});

describe('durable mission storage', () => {
  it('round-trips a completed mission, evidence, receipt and Dates without a workbook snapshot or keys', async () => {
    const factory = new IDBFactory();
    const first = createMissionStore(() => factory);
    const item = await record('completed', 'applied');
    item.receipt = {
      id: 'receipt-1',
      status: 'applied',
      createdAt: 1,
      completedAt: 2,
      operations: [{ name: 'normalize_text', affectedCells: 1, warnings: [] }],
    };
    item.action!.args.date = new Date('2026-01-01');
    item.evidence = [
      {
        id: 'source',
        kind: 'statistic',
        title: 'Income',
        source: 'Sales!A2',
        facts: [{ label: 'Sum', value: '-300' }],
      },
    ];
    await first.save({
      ...item,
      apiKey: 'not-a-real-key',
      wholeWorkbook: workbook,
    } as MissionRecord);
    const restored = await createMissionStore(() => factory).list();
    expect(restored).toEqual([item]);
    expect(restored[0]?.action?.args.date).toBeInstanceOf(Date);
    expect(restored[0]).not.toHaveProperty('apiKey');
    expect(restored[0]).not.toHaveProperty('wholeWorkbook');
  });
  it('keeps the last committed record when a transaction aborts after put', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    const item = await record();
    await store.save(item);
    const original = IDBObjectStore.prototype.put;
    const spy = vi.spyOn(IDBObjectStore.prototype, 'put').mockImplementation(function (
      this: IDBObjectStore,
      ...args
    ) {
      const request = original.apply(this, args);
      request.addEventListener('success', () => this.transaction.abort());
      return request;
    });
    await expect(store.save({ ...item, status: 'applied' })).rejects.toThrow(/abort/i);
    spy.mockRestore();
    expect((await store.list())[0]?.status).toBe('prepared');
  });
  it('bounds retention atomically and does not silently drop a newer task', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    for (let i = 0; i <= MAX_MISSIONS; i += 1)
      await store.save({ ...(await record(`task-${i}`)), updatedAt: i + 1 });
    const restored = await store.list();
    expect(restored).toHaveLength(MAX_MISSIONS);
    expect(restored[0]?.id).toBe(`task-${MAX_MISSIONS}`);
    expect(restored.some((item) => item.id === 'task-0')).toBe(false);
  });
  it('orders delete/clear after pending saves so old writes cannot resurrect deleted history', async () => {
    const factory = new IDBFactory();
    const ordered = createMissionStore(() => factory);
    const item = await record();
    await Promise.all([ordered.save(item), ordered.delete(item.id)]);
    expect(await ordered.list()).toEqual([]);
    await Promise.all([ordered.save(item), ordered.clear()]);
    expect(await ordered.list()).toEqual([]);
  });
  it('rejects malformed steps, receipts and unavailable storage rather than making work executable', async () => {
    const item = await record();
    expect(() => validateMission({ ...item, updatedAt: Number.MAX_SAFE_INTEGER })).toThrow(
      /invalid/i,
    );
    expect(() => validateMission({ ...item, action: { ...item.action, args: null } })).toThrow(
      /invalid/i,
    );
    expect(() =>
      validateMission({ ...item, receipt: { status: 'applied', operations: [] } }),
    ).toThrow(/invalid/i);
    await expect(createMissionStore(() => undefined).list()).rejects.toThrow(/unavailable/i);
  });
});
