import { useCallback, useEffect, useRef, useState } from 'react';
import { MAX_MISSIONS, type MissionRecord, type MissionStore } from './missions.js';

export type MissionStorageStatus = 'checking' | 'saving' | 'saved' | 'unavailable';
const sorted = (records: MissionRecord[]) =>
  [...records]
    .sort((a, b) => b.updatedAt - a.updatedAt || a.id.localeCompare(b.id))
    .slice(0, MAX_MISSIONS);

/** In-memory work remains usable on storage failure; never claim a durable save/deletion before commit. */
export function useMissionLedger(store: MissionStore) {
  const [missions, setMissions] = useState<MissionRecord[]>([]);
  const records = useRef<MissionRecord[]>([]);
  const [storageStatus, setStorageStatus] = useState<MissionStorageStatus>('checking');
  const [storageError, setStorageError] = useState('');
  const [unsavedIds, setUnsavedIds] = useState<Set<string>>(new Set());
  const unsaved = useRef(new Map<string, MissionRecord>());
  const queue = useRef<Promise<unknown>>(Promise.resolve());
  const mounted = useRef(true);
  const readSucceeded = useRef(false);
  const deletingIds = useRef(new Set<string>());
  const publish = useCallback((values: MissionRecord[]) => {
    records.current = sorted(values);
    const retained = new Set(records.current.map((item) => item.id));
    let dropped = false;
    for (const id of unsaved.current.keys())
      if (!retained.has(id)) {
        unsaved.current.delete(id);
        dropped = true;
      }
    if (mounted.current) {
      setMissions(records.current);
      if (dropped) setUnsavedIds(new Set(unsaved.current.keys()));
    }
  }, []);
  const enqueue = useCallback(<T>(task: () => Promise<T>): Promise<T> => {
    const result = queue.current.then(task, task);
    queue.current = result.catch(() => undefined);
    return result;
  }, []);
  const fail = useCallback((error: unknown) => {
    if (!mounted.current) return;
    setStorageStatus('unavailable');
    setStorageError(error instanceof Error ? error.message : 'Mission storage is unavailable.');
  }, []);
  const refreshUnsaved = useCallback(() => {
    if (mounted.current) setUnsavedIds(new Set(unsaved.current.keys()));
  }, []);
  const save = useCallback(
    (record: MissionRecord) => {
      unsaved.current.set(record.id, record);
      refreshUnsaved();
      if (mounted.current) setStorageStatus('saving');
      void enqueue(() => store.save(record))
        .then(() => {
          if (!mounted.current) return;
          // An older committed revision cannot mark a newer pending record saved.
          if (unsaved.current.get(record.id) === record) unsaved.current.delete(record.id);
          refreshUnsaved();
          if (!unsaved.current.size && readSucceeded.current) {
            setStorageStatus('saved');
            setStorageError('');
          }
        })
        .catch(fail);
    },
    [enqueue, fail, refreshUnsaved, store],
  );
  const upsert = useCallback(
    (record: MissionRecord) => {
      if (deletingIds.current.has(record.id)) return;
      publish([record, ...records.current.filter((item) => item.id !== record.id)]);
      if (records.current.some((item) => item.id === record.id)) save(record);
    },
    [publish, save],
  );
  const update = useCallback(
    (id: string, patch: Partial<MissionRecord>) => {
      const current = records.current.find((item) => item.id === id);
      if (!current) return;
      upsert({ ...current, ...patch, id: current.id, updatedAt: Date.now() });
    },
    [upsert],
  );
  const load = useCallback(async () => {
    try {
      const values = await enqueue(() => store.list());
      if (!mounted.current) return;
      readSucceeded.current = true;
      const currentIds = new Set(records.current.map((item) => item.id));
      const recovered = values
        .filter((item) => !currentIds.has(item.id))
        .map((item) =>
          item.status === 'planning' || item.status === 'executing'
            ? {
                ...item,
                status: 'interrupted' as const,
                stage: undefined,
                error:
                  'The last saved state was in progress when this page closed. The task will not resume automatically. Inspect the restored workbook before replanning; a final commit may not have reached mission storage.',
              }
            : item,
        );
      publish([...records.current, ...recovered]);
      for (const item of recovered) if (item.status === 'interrupted') save(item);
      if (!unsaved.current.size) {
        setStorageStatus('saved');
        setStorageError('');
      }
    } catch (error) {
      fail(error);
    }
  }, [enqueue, fail, publish, save, store]);
  useEffect(() => {
    mounted.current = true;
    void load();
    return () => {
      mounted.current = false;
    };
  }, [load]);
  const remove = useCallback(
    async (id: string) => {
      deletingIds.current.add(id);
      try {
        await enqueue(() => store.delete(id));
        unsaved.current.delete(id);
        // A successful delete only removes this in-flight guard. Keeping the id here would make
        // every future mission with the same id silently disappear from the in-memory ledger.
        deletingIds.current.delete(id);
        refreshUnsaved();
        publish(records.current.filter((item) => item.id !== id));
        if (mounted.current && !unsaved.current.size) {
          setStorageStatus('saved');
          setStorageError('');
        }
      } catch (error) {
        deletingIds.current.delete(id);
        fail(error);
        throw error;
      }
    },
    [enqueue, fail, publish, refreshUnsaved, store],
  );
  const clear = useCallback(async () => {
    const ids = new Set(records.current.map((item) => item.id));
    for (const id of ids) deletingIds.current.add(id);
    try {
      await enqueue(() => store.clear());
      // Work created while deletion was in flight stays visible and is queued for a later save.
      for (const id of ids) unsaved.current.delete(id);
      for (const id of ids) deletingIds.current.delete(id);
      refreshUnsaved();
      publish(records.current.filter((item) => !ids.has(item.id)));
      if (mounted.current && !unsaved.current.size) {
        setStorageStatus('saved');
        setStorageError('');
      }
    } catch (error) {
      for (const id of ids) deletingIds.current.delete(id);
      fail(error);
      throw error;
    }
  }, [enqueue, fail, publish, refreshUnsaved, store]);
  const retry = useCallback(async () => {
    await load();
    for (const record of unsaved.current.values()) save(record);
  }, [load, save]);
  return {
    missions,
    records,
    storageStatus,
    storageError,
    unsavedIds,
    upsert,
    update,
    remove,
    clear,
    retry,
  };
}
