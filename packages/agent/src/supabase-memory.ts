import type { Workbook } from '@excel-agent/engine';
import { InMemoryMemoryStore, normalizeQuery, querySimilarity, sheetFingerprint } from './memory.js';
import type { MemoryRecord, MemoryStore } from './types.js';

export interface SupabaseMemoryConfig {
  url: string;
  apiKey: string;
  /** Whether cloud synchronization is enabled. Defaults to true. When false, acts purely as local in-memory store. */
  enabled?: boolean;
  /** Timeout for Supabase network calls in milliseconds. Defaults to 2500ms so the user is never kept waiting. */
  timeoutMs?: number;
  /** Minimum confidence score required to replay a learned cloud action (0 to 1). Defaults to 0.5. */
  minConfidence?: number;
  /** Local in-memory capacity. Defaults to 500 records. */
  maxLocalRecords?: number;
}

export interface WorkingMemoryEntry {
  id?: string;
  sessionId: string;
  stepType: string;
  payload: Record<string, unknown>;
  createdAt?: string;
  expiresAt?: string;
}

interface CortexRow {
  id: string;
  raw_query: string;
  normalized_query: string;
  schema_fingerprint: string;
  operation: string;
  args: Record<string, unknown>;
  success_count: number;
  failure_count: number;
  confidence: number;
  first_learned_at: string;
  last_used_at: string;
}

/**
 * Enterprise Production Dual-Store Supabase Memory Provider.
 *
 * Partitions:
 * 1. 100 MB Ephemeral Ring-Buffer Pool (agent_working_memory) - Scratchpad calculations, auto-recycled.
 * 2. 400 MB Collective Intelligence Cortex (agent_learned_cortex) - Long-term persistent learning.
 *
 * Reliability guarantees:
 * - Instant local hot-cache for 0ms lookups.
 * - Non-blocking asynchronous cloud synchronization.
 * - Strict timeout (default 2500ms) with seamless local fallback if offline.
 */
export class SupabaseMemoryStore implements MemoryStore {
  private readonly url: string;
  private readonly apiKey: string;
  private readonly enabled: boolean;
  private readonly timeoutMs: number;
  private readonly minConfidence: number;
  private readonly localStore: InMemoryMemoryStore;

  constructor(config: SupabaseMemoryConfig) {
    this.url = config.url.replace(/\/+$/, '');
    this.apiKey = config.apiKey;
    this.enabled = config.enabled ?? true;
    this.timeoutMs = config.timeoutMs ?? 2500;
    this.minConfidence = config.minConfidence ?? 0.5;
    this.localStore = new InMemoryMemoryStore({ maxRecords: config.maxLocalRecords ?? 500 });
  }

  private headers(): Record<string, string> {
    return {
      apikey: this.apiKey,
      Authorization: `Bearer ${this.apiKey}`,
      'Content-Type': 'application/json',
      Prefer: 'return=representation',
    };
  }

  private async fetchWithTimeout(endpoint: string, options: RequestInit = {}): Promise<Response> {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await fetch(`${this.url}${endpoint}`, {
        ...options,
        signal: controller.signal,
      });
      return response;
    } finally {
      clearTimeout(timer);
    }
  }

  remember(
    record: Omit<MemoryRecord, 'id' | 'createdAt' | 'lastUsedAt' | 'successes' | 'failures'>,
  ): MemoryRecord {
    // 1. Immediately update local cache
    const local = this.localStore.remember(record);

    if (!this.enabled) return local;

    // 2. Asynchronously sync to Supabase agent_learned_cortex (non-blocking)
    const norm = normalizeQuery(record.rawQuery || record.key || '');
    const fingerprint = record.schemaFingerprint || 'unknown';

    if (norm && fingerprint !== 'unknown') {
      void this.syncRememberToCloud(record.rawQuery || record.key, norm, fingerprint, record.operation, record.args);
    }

    return local;
  }

  private async syncRememberToCloud(
    rawQuery: string,
    normalizedQuery: string,
    schemaFingerprint: string,
    operation: string,
    args: Record<string, unknown>,
  ): Promise<void> {
    try {
      // Check if this pattern already exists for the given schema
      const searchRes = await this.fetchWithTimeout(
        `/rest/v1/agent_learned_cortex?schema_fingerprint=eq.${encodeURIComponent(
          schemaFingerprint,
        )}&normalized_query=eq.${encodeURIComponent(normalizedQuery)}&operation=eq.${encodeURIComponent(
          operation,
        )}&select=id,success_count`,
        { headers: this.headers() },
      );

      if (searchRes.ok) {
        const existing = (await searchRes.json()) as Array<{ id: string; success_count: number }>;
        if (existing.length > 0 && existing[0]) {
          // Increment success count and update last used
          await this.fetchWithTimeout(
            `/rest/v1/agent_learned_cortex?id=eq.${existing[0].id}`,
            {
              method: 'PATCH',
              headers: this.headers(),
              body: JSON.stringify({
                success_count: (existing[0].success_count || 1) + 1,
                last_used_at: new Date().toISOString(),
                args,
              }),
            },
          );
          return;
        }
      }

      // Insert new record
      await this.fetchWithTimeout(`/rest/v1/agent_learned_cortex`, {
        method: 'POST',
        headers: this.headers(),
        body: JSON.stringify({
          raw_query: rawQuery,
          normalized_query: normalizedQuery,
          schema_fingerprint: schemaFingerprint,
          operation,
          args,
          success_count: 1,
          failure_count: 0,
        }),
      });
    } catch {
      // Non-blocking: background sync failure does not disrupt the user
    }
  }

  recordOutcome(operation: string, sheetName: string, success: boolean, key?: string): void {
    this.localStore.recordOutcome(operation, sheetName, success, key);

    if (!this.enabled || !key) return;

    const norm = normalizeQuery(key);
    void this.syncOutcomeToCloud(norm, operation, success);
  }

  private async syncOutcomeToCloud(
    normalizedQuery: string,
    operation: string,
    success: boolean,
  ): Promise<void> {
    try {
      const searchRes = await this.fetchWithTimeout(
        `/rest/v1/agent_learned_cortex?normalized_query=eq.${encodeURIComponent(
          normalizedQuery,
        )}&operation=eq.${encodeURIComponent(operation)}&select=id,success_count,failure_count&limit=1`,
        { headers: this.headers() },
      );

      if (searchRes.ok) {
        const rows = (await searchRes.json()) as Array<{
          id: string;
          success_count: number;
          failure_count: number;
        }>;
        if (rows.length > 0 && rows[0]) {
          const row = rows[0];
          await this.fetchWithTimeout(`/rest/v1/agent_learned_cortex?id=eq.${row.id}`, {
            method: 'PATCH',
            headers: this.headers(),
            body: JSON.stringify({
              success_count: success ? (row.success_count || 0) + 1 : row.success_count,
              failure_count: !success ? (row.failure_count || 0) + 1 : row.failure_count,
              last_used_at: new Date().toISOString(),
            }),
          });
        }
      }
    } catch {
      // Non-blocking
    }
  }

  retrieve(query: string, sheetName: string, threshold = 0.6): MemoryRecord | undefined {
    return this.localStore.retrieve(query, sheetName, threshold);
  }

  async retrieveForWorkbook(
    query: string,
    workbook: Workbook,
    sheetName: string,
    threshold = 0.6,
  ): Promise<MemoryRecord | undefined> {
    // 1. Check local cache first (instant 0ms)
    const local = this.localStore.retrieveForWorkbook(query, workbook, sheetName, threshold);
    if (local && this.localStore.confidenceOf(local) >= this.minConfidence) {
      return local;
    }

    if (!this.enabled) return local;

    // 2. Query Supabase 400 MB Cortex
    const currentFingerprint = sheetFingerprint(workbook, sheetName);
    if (!currentFingerprint) return local;

    try {
      const res = await this.fetchWithTimeout(
        `/rest/v1/agent_learned_cortex?schema_fingerprint=eq.${encodeURIComponent(
          currentFingerprint,
        )}&confidence=gte.${this.minConfidence}&order=confidence.desc,last_used_at.desc&limit=10`,
        { headers: this.headers() },
      );

      if (!res.ok) return local;

      const candidates = (await res.json()) as CortexRow[];
      if (!Array.isArray(candidates) || candidates.length === 0) return local;

      let bestMatch: CortexRow | undefined;
      let highestSimilarity = threshold;

      for (const candidate of candidates) {
        const sim = querySimilarity(query, candidate.normalized_query || candidate.raw_query);
        if (sim >= highestSimilarity) {
          highestSimilarity = sim;
          bestMatch = candidate;
        }
      }

      if (bestMatch && bestMatch.schema_fingerprint === currentFingerprint) {
        // Cache in local memory store for subsequent turns
        const learned = this.localStore.remember({
          key: bestMatch.normalized_query,
          rawQuery: bestMatch.raw_query,
          sheetName,
          operation: bestMatch.operation,
          args: bestMatch.args,
          schemaFingerprint: bestMatch.schema_fingerprint,
        });

        // Seed with cloud success counts
        learned.successes = bestMatch.success_count;
        learned.failures = bestMatch.failure_count;
        return learned;
      }
    } catch {
      // Fallback seamlessly to local cache on any network failure
    }

    return local;
  }

  entries(): MemoryRecord[] {
    return this.localStore.entries();
  }

  confidenceOf(record: MemoryRecord): number {
    return this.localStore.confidenceOf(record);
  }

  clear(): void {
    this.localStore.clear();
  }

  toJSON(): string {
    return this.localStore.toJSON();
  }

  // ==========================================================================
  // Ephemeral Working Memory Pool (100 MB Recycled Scratchpad)
  // ==========================================================================

  async saveWorkingMemory(
    sessionId: string,
    stepType: string,
    payload: Record<string, unknown>,
  ): Promise<boolean> {
    if (!this.enabled) return true;
    try {
      const res = await this.fetchWithTimeout(`/rest/v1/agent_working_memory`, {
        method: 'POST',
        headers: this.headers(),
        body: JSON.stringify({
          session_id: sessionId,
          step_type: stepType,
          payload,
        }),
      });
      return res.ok;
    } catch {
      return false;
    }
  }

  async clearWorkingMemory(sessionId: string): Promise<boolean> {
    if (!this.enabled) return true;
    try {
      const res = await this.fetchWithTimeout(
        `/rest/v1/agent_working_memory?session_id=eq.${encodeURIComponent(sessionId)}`,
        {
          method: 'DELETE',
          headers: this.headers(),
        },
      );
      return res.ok;
    } catch {
      return false;
    }
  }

  async getWorkingMemory(sessionId: string): Promise<WorkingMemoryEntry[]> {
    if (!this.enabled) return [];
    try {
      const res = await this.fetchWithTimeout(
        `/rest/v1/agent_working_memory?session_id=eq.${encodeURIComponent(
          sessionId,
        )}&order=created_at.asc`,
        { headers: this.headers() },
      );
      if (!res.ok) return [];
      const rows = (await res.json()) as Array<{
        id: string;
        session_id: string;
        step_type: string;
        payload: Record<string, unknown>;
        created_at: string;
        expires_at: string;
      }>;
      return rows.map((r) => ({
        id: r.id,
        sessionId: r.session_id,
        stepType: r.step_type,
        payload: r.payload,
        createdAt: r.created_at,
        expiresAt: r.expires_at,
      }));
    } catch {
      return [];
    }
  }
}
