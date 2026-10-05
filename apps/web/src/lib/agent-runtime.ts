import { SupabaseMemoryStore, createOrchestrator } from '@excel-agent/agent';
import { createOperationRegistry } from '@excel-agent/engine';

const MEMORY_KEY = 'excel_agent_memory_v1';

const SUPABASE_URL =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_URL) ||
  'https://fsepapdadtrlddkyqqxu.supabase.co';
const SUPABASE_KEY =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_ANON_KEY) ||
  'sb_publishable_8JPAZFfCaS_U8nAAqc1rrQ_V6OSKAic';

function storage(): Storage | undefined {
  try {
    return typeof localStorage === 'undefined' ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

function readMemory(): string | undefined {
  try {
    return storage()?.getItem(MEMORY_KEY) ?? undefined;
  } catch {
    return undefined;
  }
}

/**
 * Single shared runtime for the workspace: one engine registry, one dual-store Supabase
 * self-learning memory (with 100MB ephemeral scratchpad and 400MB collective cortex),
 * and one multi-layer orchestrator.
 */
export const registry = createOperationRegistry();

const isTestMode =
  typeof import.meta !== 'undefined' &&
  (import.meta.env?.MODE === 'test' || Boolean(import.meta.env?.VITEST));

export const memory = new SupabaseMemoryStore({
  url: SUPABASE_URL,
  apiKey: SUPABASE_KEY,
  enabled: !isTestMode,
  timeoutMs: 2500,
  minConfidence: 0.5,
});

// Preload any existing browser local storage into the local hot-cache
const localCache = readMemory();
if (localCache) {
  try {
    // If local memory serialized JSON exists, seed the in-memory cache
    const parsed: unknown = JSON.parse(localCache);
    const records =
      typeof parsed === 'object' && parsed !== null && 'records' in parsed
        ? parsed.records
        : undefined;
    if (Array.isArray(records)) {
      for (const rec of records) {
        if (!isMemorySeed(rec)) continue;
        memory.remember(rec);
      }
    }
  } catch {
    // Ignore corrupt local cache
  }
}

function isMemorySeed(value: unknown): value is Parameters<typeof memory.remember>[0] {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return (
    typeof record.key === 'string' &&
    typeof record.rawQuery === 'string' &&
    typeof record.operation === 'string' &&
    typeof record.sheetName === 'string' &&
    typeof record.args === 'object' &&
    record.args !== null &&
    !Array.isArray(record.args)
  );
}

export const orchestrator = createOrchestrator({ registry, memory });

export function persistMemory(): void {
  try {
    storage()?.setItem(MEMORY_KEY, memory.toJSON());
  } catch {
    // Storage may be full or unavailable; learning simply stays in-memory.
  }
}

export function forgetLearnedActions(): void {
  memory.clear();
  try {
    storage()?.removeItem(MEMORY_KEY);
  } catch {
    // Ignore storage failures.
  }
}

export function learnedActionCount(): number {
  return memory.entries().length;
}

export async function initCloudMemory(onUpdate?: (count: number) => void): Promise<number> {
  try {
    const loaded = await memory.hydrateFromCloud();
    if (loaded > 0) {
      persistMemory();
    }
    const total = learnedActionCount();
    onUpdate?.(total);
    return total;
  } catch {
    const total = learnedActionCount();
    onUpdate?.(total);
    return total;
  }
}

export function getMemoryEntries() {
  return memory.entries();
}

export function getMemoryCloudStatus() {
  return memory.getCloudStatus();
}

export { getTabSessionId, clearCurrentTabWorkingMemory } from './session';
