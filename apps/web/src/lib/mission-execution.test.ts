import { afterEach, describe, expect, it, vi } from 'vitest';
import { createCell, HistoryStack, type Workbook } from '@excel-agent/engine';
import {
  executeMissionMutation,
  stageMissionMutation,
  type MissionMutationRequest,
} from './mission-execution.js';
const workbook: Workbook = {
  sheets: [
    {
      name: 'Orders',
      rows: [
        [createCell('ID'), createCell('Name')],
        [createCell('1'), createCell(' Alice ')],
        [createCell('1'), createCell(' Alice ')],
      ],
    },
  ],
};
const request: MissionMutationRequest = {
  workbook,
  confirmed: false,
  steps: [
    { operation: 'normalize_text', args: { sheet: 'Orders', columns: ['B'], trim: true } },
    { operation: 'delete_duplicates', args: { sheet: 'Orders', columns: ['A'], headerRow: 1 } },
  ],
};
afterEach(() => vi.unstubAllGlobals());
describe('private mission execution', () => {
  it('does not leak partial steps before confirmation, then produces one reversible commit', () => {
    const before = structuredClone(workbook);
    const progress: unknown[] = [];
    const pending = stageMissionMutation(request, (event) => progress.push(event));
    expect(pending.ok).toBe(false);
    if (!pending.ok) expect(pending.error.code).toBe('confirmation-required');
    expect(workbook).toEqual(before);
    expect(progress).toEqual([
      { index: 0, phase: 'verifying' },
      { index: 0, phase: 'verified' },
      { index: 1, phase: 'verifying' },
      { index: 1, phase: 'verified' },
    ]);
    const result = stageMissionMutation({ ...request, confirmed: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.workbook.sheets[0]!.rows).toHaveLength(2);
    expect(result.workbook.sheets[0]!.rows[1]![1]!.value).toBe('Alice');
    expect(result.steps[0]?.report.affectedCells).toBe(2);
    const history = new HistoryStack(workbook);
    history.commit('mission', result.workbook, result.patch, result.inverse);
    expect(history.length).toBe(1);
    expect(history.undo()).toEqual(workbook);
  });
  it('rejects an invalid late step without a patch or changed caller data', () => {
    const result = stageMissionMutation({
      ...request,
      confirmed: true,
      steps: [...request.steps, { operation: 'not_an_operation', args: {} }],
    });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.failedStep).toBe(2);
    expect(result).not.toHaveProperty('patch');
    expect(workbook.sheets[0]!.rows[1]![1]!.value).toBe(' Alice ');
  });
  it('terminates its dedicated worker on abort and ignores late worker output', async () => {
    let worker!: Worker;
    const terminate = vi.fn();
    class FakeWorker {
      onmessage?: Worker['onmessage'];
      onerror?: Worker['onerror'];
      constructor() {
        worker = this as unknown as Worker;
      }
      terminate = terminate;
      postMessage = vi.fn();
    }
    vi.stubGlobal('Worker', FakeWorker);
    const controller = new AbortController();
    const pending = executeMissionMutation(request, { signal: controller.signal });
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    expect(terminate).toHaveBeenCalledOnce();
    worker.onmessage?.call(worker, {
      data: { type: 'result', result: { ok: true } },
    } as MessageEvent);
    expect(terminate).toHaveBeenCalledOnce();
  });
  it('surfaces worker crashes and timeouts without silently executing again on the main thread', async () => {
    let worker!: Worker;
    const terminate = vi.fn();
    class FakeWorker {
      onmessage?: Worker['onmessage'];
      onerror?: Worker['onerror'];
      constructor() {
        worker = this as unknown as Worker;
      }
      terminate = terminate;
      postMessage = vi.fn();
    }
    vi.stubGlobal('Worker', FakeWorker);
    const pending = executeMissionMutation(request);
    worker.onerror?.call(worker, {} as ErrorEvent);
    await expect(pending).rejects.toThrow(/worker failed/i);
    await expect(executeMissionMutation(request, { timeoutMs: 5 })).rejects.toThrow(/timed out/i);
    expect(terminate).toHaveBeenCalledTimes(2);
  });
});
