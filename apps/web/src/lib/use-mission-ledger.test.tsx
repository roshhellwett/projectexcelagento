// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { useMissionLedger } from './use-mission-ledger.js';
import type { MissionRecord, MissionStore } from './missions.js';
const record = (id = 'task'): MissionRecord => ({
  version: 1,
  id,
  kind: 'analysis',
  status: 'analyzed',
  title: 'Analysis',
  request: 'inspect',
  createdAt: 1,
  updatedAt: 1,
  workbook: {
    generation: 0,
    revision: 0,
    fileName: 'sample.xlsx',
    sheetName: 'Sales',
    signature: null,
  },
});
const deferred = () => {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
};
function memoryStore(): MissionStore {
  const items = new Map<string, MissionRecord>();
  return {
    list: vi.fn(async () => [...items.values()]),
    save: vi.fn(async (item) => {
      items.set(item.id, item);
    }),
    delete: vi.fn(async (id) => {
      items.delete(id);
    }),
    clear: vi.fn(async () => {
      items.clear();
    }),
  };
}
describe('mission ledger commit semantics', () => {
  it('never lets an older save clear a newer unsaved revision', async () => {
    const store = memoryStore();
    const first = deferred();
    const second = deferred();
    vi.mocked(store.save)
      .mockImplementationOnce(() => first.promise)
      .mockImplementationOnce(() => second.promise);
    const { result } = renderHook(() => useMissionLedger(store));
    await waitFor(() => expect(result.current.storageStatus).toBe('saved'));
    act(() => {
      result.current.upsert(record());
      result.current.update('task', { answer: 'new answer' });
    });
    await waitFor(() => expect(store.save).toHaveBeenCalledTimes(1));
    await act(async () => first.resolve());
    expect(result.current.unsavedIds.has('task')).toBe(true);
    expect(result.current.storageStatus).toBe('saving');
    expect(result.current.missions[0]?.answer).toBe('new answer');
    await act(async () => second.resolve());
    await waitFor(() => expect(result.current.storageStatus).toBe('saved'));
    expect(result.current.unsavedIds.size).toBe(0);
  });
  it('keeps history after clear failure and removes it only after the clear transaction resolves', async () => {
    const store = memoryStore();
    await store.save(record());
    vi.mocked(store.clear).mockRejectedValueOnce(new Error('blocked clear'));
    const { result } = renderHook(() => useMissionLedger(store));
    await waitFor(() => expect(result.current.missions).toHaveLength(1));
    await act(async () => {
      await expect(result.current.clear()).rejects.toThrow('blocked clear');
    });
    expect(result.current.missions).toHaveLength(1);
    expect(result.current.storageStatus).toBe('unavailable');
    const committed = deferred();
    vi.mocked(store.clear).mockImplementationOnce(() => committed.promise);
    let deletion!: Promise<void>;
    act(() => {
      deletion = result.current.clear();
    });
    expect(result.current.missions).toHaveLength(1);
    await act(async () => {
      committed.resolve();
      await deletion;
    });
    expect(result.current.missions).toEqual([]);
  });
  it('bounds unsaved history with visible retention rather than leaving hidden records on retry', async () => {
    const store = memoryStore();
    vi.mocked(store.save).mockRejectedValue(new Error('quota'));
    const { result } = renderHook(() => useMissionLedger(store));
    await waitFor(() => expect(result.current.storageStatus).toBe('saved'));
    act(() => {
      for (let i = 0; i < 51; i += 1)
        result.current.upsert({ ...record(`task-${i}`), updatedAt: i + 1 });
    });
    await waitFor(() => expect(result.current.storageStatus).toBe('unavailable'));
    expect(result.current.missions).toHaveLength(50);
    expect(result.current.unsavedIds.size).toBe(50);
    expect(result.current.unsavedIds.has('task-0')).toBe(false);
  });

  it('serializes delete after pending saves and suppresses late updates to removed tasks', async () => {
    const store = memoryStore();
    const writing = deferred();
    vi.mocked(store.save).mockImplementationOnce(() => writing.promise);
    const { result } = renderHook(() => useMissionLedger(store));
    await waitFor(() => expect(result.current.storageStatus).toBe('saved'));
    act(() => result.current.upsert(record()));
    let removal!: Promise<void>;
    act(() => {
      removal = result.current.remove('task');
      result.current.update('task', { answer: 'late callback' });
    });
    expect(store.delete).not.toHaveBeenCalled();
    await act(async () => {
      writing.resolve();
      await removal;
    });
    expect(result.current.missions).toEqual([]);
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(store.delete).toHaveBeenCalledExactlyOnceWith('task');
  });
});
