import { describe, expect, it, vi, afterEach } from 'vitest';
import { createCell, type Workbook } from '@excel-agent/engine';
import { SupabaseMemoryStore } from '../src/supabase-memory.js';

function mockWorkbook(sheetName = 'Sheet1', headers = ['ID', 'Type', 'Amount']): Workbook {
  return {
    sheets: [
      {
        name: sheetName,
        rows: [
          headers.map((h) => createCell(h)),
          [createCell('1'), createCell('IN'), createCell(100)],
        ],
      },
    ],
  };
}

describe('SupabaseMemoryStore', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('immediately updates and retrieves from local cache without waiting for cloud', () => {
    const store = new SupabaseMemoryStore({
      url: 'https://test.supabase.co',
      apiKey: 'test-key',
    });

    const record = store.remember({
      key: 'filter in data',
      rawQuery: 'filter out the IN data into a separate sheet',
      sheetName: 'Sheet1',
      operation: 'filter_to_new_sheet',
      args: { column: 'B', operator: 'contains', value: 'IN' },
      schemaFingerprint: 'id\u0001type\u0001amount',
    });

    expect(record.operation).toBe('filter_to_new_sheet');
    expect(store.entries()).toHaveLength(1);
    store.recordOutcome('filter_to_new_sheet', 'Sheet1', true);
    const retrieved = store.retrieve('filter in data', 'Sheet1');
    expect(retrieved?.operation).toBe('filter_to_new_sheet');
  });

  it('queries cloud cortex when local memory does not have the record', async () => {
    const store = new SupabaseMemoryStore({
      url: 'https://test.supabase.co',
      apiKey: 'test-key',
      minConfidence: 0.5,
    });

    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL) => {
        const urlStr = String(input);
        if (urlStr.includes('agent_learned_cortex')) {
          return new Response(
            JSON.stringify([
              {
                id: 'cloud-1',
                raw_query: 'extract IN data to new sheet',
                normalized_query: 'extract in data new sheet',
                schema_fingerprint: 'id\u0001type\u0001amount',
                operation: 'filter_to_new_sheet',
                args: { column: 'B', operator: 'contains', value: 'IN' },
                success_count: 5,
                failure_count: 0,
                confidence: 1.0,
                first_learned_at: new Date().toISOString(),
                last_used_at: new Date().toISOString(),
              },
            ]),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }
        return new Response('[]', { status: 200 });
      }),
    );

    const wb = mockWorkbook('Sheet1', ['ID', 'Type', 'Amount']);
    const hit = await store.retrieveForWorkbook('extract in data to new sheet', wb, 'Sheet1', 0.5);

    expect(hit).toBeDefined();
    expect(hit?.operation).toBe('filter_to_new_sheet');
    expect(hit?.args).toEqual({ column: 'B', operator: 'contains', value: 'IN' });
  });

  it('refuses cloud memory hit when schema fingerprint differs from the sheet layout', async () => {
    const store = new SupabaseMemoryStore({
      url: 'https://test.supabase.co',
      apiKey: 'test-key',
      minConfidence: 0.5,
    });

    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        new Response(
          JSON.stringify([
            {
              id: 'cloud-1',
              raw_query: 'delete column B',
              normalized_query: 'delete column b',
              schema_fingerprint: 'customer\u0001notes\u0001price', // different fingerprint
              operation: 'delete_column',
              args: { column: 'B' },
              success_count: 5,
              failure_count: 0,
              confidence: 1.0,
            },
          ]),
          { status: 200, headers: { 'Content-Type': 'application/json' } },
        ),
      ),
    );

    const wb = mockWorkbook('Sheet1', ['ID', 'Type', 'Amount']);
    const hit = await store.retrieveForWorkbook('delete column B', wb, 'Sheet1', 0.5);
    // Should not match because fingerprint is different
    expect(hit).toBeUndefined();
  });

  it('gracefully falls back to local cache when cloud fetch fails or times out', async () => {
    const store = new SupabaseMemoryStore({
      url: 'https://test.supabase.co',
      apiKey: 'test-key',
      timeoutMs: 100,
    });

    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        throw new Error('Network offline');
      }),
    );

    const wb = mockWorkbook('Sheet1', ['ID', 'Type', 'Amount']);
    // Should not throw, returns undefined gracefully
    const hit = await store.retrieveForWorkbook('unknown query', wb, 'Sheet1', 0.5);
    expect(hit).toBeUndefined();
  });

  it('handles ephemeral working memory operations (save, get, clear)', async () => {
    const store = new SupabaseMemoryStore({
      url: 'https://test.supabase.co',
      apiKey: 'test-key',
    });

    const calls: string[] = [];
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const method = init?.method ?? 'GET';
        calls.push(`${method} ${String(input)}`);
        if (method === 'GET') {
          return new Response(
            JSON.stringify([
              {
                id: 'uuid-1',
                session_id: 'session-123',
                step_type: 'analysis',
                payload: { candidateCol: 'D' },
                created_at: new Date().toISOString(),
                expires_at: new Date().toISOString(),
              },
            ]),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }
        return new Response('[]', { status: 200 });
      }),
    );

    const saved = await store.saveWorkingMemory('session-123', 'analysis', { candidateCol: 'D' });
    expect(saved).toBe(true);

    const entries = await store.getWorkingMemory('session-123');
    expect(entries.length).toBe(1);
    expect(entries[0]?.stepType).toBe('analysis');

    const cleared = await store.clearWorkingMemory('session-123');
    expect(cleared).toBe(true);
    expect(calls.some((c) => c.startsWith('DELETE'))).toBe(true);
  });
});
