import { createMemoryStore, createOrchestrator } from '@excel-agent/agent';
import { createOperationRegistry } from '@excel-agent/engine';

const MEMORY_KEY = 'excel_agent_memory_v1';

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
 * Single shared runtime for the workspace: one engine registry, one 100% in-browser local
 * self-learning memory (0ms latency, zero external network calls, total data privacy),
 * and one multi-layer orchestrator.
 */
export const registry = createOperationRegistry();

export const cloudMemoryConfig = {
  url: '',
  apiKey: '',
  enabled: false,
  timeoutMs: 0,
  minConfidence: 0.5,
};

// Initialize 100% in-browser local memory from localStorage
export const memory = createMemoryStore(readMemory());

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
  const total = learnedActionCount();
  onUpdate?.(total);
  return total;
}

export function getMemoryEntries() {
  return memory.entries();
}

export function getMemoryCloudStatus() {
  return {
    enabled: false,
    url: '',
    syncedCount: memory.entries().length,
    minConfidence: 0.5,
  };
}

export { getTabSessionId, clearCurrentTabWorkingMemory } from './session';
