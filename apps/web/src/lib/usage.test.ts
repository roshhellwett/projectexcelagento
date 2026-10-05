// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';

import type { LlmTelemetry } from '@excel-agent/agent';

import {
  MAX_USAGE_ENTRIES,
  USAGE_KEY,
  appendUsageEntry,
  clearUsageLog,
  createUsageEntry,
  formatTokenCount,
  loadUsageLog,
  maskApiKey,
  summarizeUsage,
  type UsageEntry,
} from './usage.js';

const TELEMETRY: LlmTelemetry = {
  provider: 'groq',
  model: 'llama-3.3-70b-versatile',
  promptTokens: 900,
  completionTokens: 80,
  totalTokens: 980,
  latencyMs: 420,
  ok: true,
};

function entry(overrides: Partial<UsageEntry> = {}): UsageEntry {
  return {
    id: `usage-${Math.random().toString(36).slice(2)}`,
    timestamp: 1_700_000_000_000,
    provider: 'groq',
    model: 'llama-3.3-70b-versatile',
    source: 'llm',
    llmUsed: true,
    promptTokens: 100,
    completionTokens: 25,
    totalTokens: 125,
    latencyMs: 300,
    ok: true,
    query: 'clean the dates',
    ...overrides,
  };
}

beforeEach(() => {
  localStorage.clear();
});

describe('createUsageEntry', () => {
  it('records provider-reported tokens for a model call', () => {
    const created = createUsageEntry({
      query: '  clean dates  ',
      source: 'llm',
      telemetry: TELEMETRY,
    });

    expect(created).toMatchObject({
      provider: 'groq',
      model: 'llama-3.3-70b-versatile',
      llmUsed: true,
      promptTokens: 900,
      completionTokens: 80,
      totalTokens: 980,
      latencyMs: 420,
      ok: true,
      query: 'clean dates',
      source: 'llm',
    });
  });

  it('derives a total when the provider omits one', () => {
    const created = createUsageEntry({
      query: 'q',
      source: 'llm',
      telemetry: {
        provider: 'gemini',
        model: 'gemini-2.0-flash',
        promptTokens: 30,
        completionTokens: 12,
        latencyMs: 100,
        ok: true,
      },
    });

    expect(created.totalTokens).toBe(42);
  });

  it('labels missing provider usage as unknown rather than a free model turn', () => {
    const created = createUsageEntry({
      query: 'q',
      source: 'llm',
      telemetry: {
        provider: 'groq',
        model: 'test',
        latencyMs: 10,
        ok: true,
      },
    });
    expect(created.llmUsed).toBe(true);
    expect(created.tokenUsageReported).toBe(false);
    expect(summarizeUsage([created]).unreportedUsageRequests).toBe(1);
  });

  it('records failures with their reason', () => {
    const created = createUsageEntry({
      query: 'q',
      source: 'fallback',
      telemetry: {
        provider: 'openrouter',
        model: 'unknown',
        latencyMs: 12,
        ok: false,
        error: 'HTTP 401: Invalid API Key',
      },
    });

    expect(created.ok).toBe(false);
    expect(created.error).toContain('401');
    expect(created.totalTokens).toBe(0);
  });

  it('marks a locally served turn as zero tokens rather than guessing', () => {
    const created = createUsageEntry({ query: 'remove duplicate rows', source: 'heuristic' });

    expect(created).toMatchObject({
      provider: 'none',
      llmUsed: false,
      promptTokens: 0,
      completionTokens: 0,
      totalTokens: 0,
      ok: true,
    });
    expect(created.model).toContain('deterministic');
  });

  it('truncates the stored prompt so the ledger cannot grow unbounded', () => {
    const created = createUsageEntry({
      query: 'x'.repeat(500),
      source: 'llm',
      telemetry: TELEMETRY,
    });
    expect(created.query).toHaveLength(120);
  });
});

describe('persistence', () => {
  it('round-trips entries through localStorage', () => {
    const created = createUsageEntry({ query: 'q', source: 'llm', telemetry: TELEMETRY });
    const saved = appendUsageEntry(created);

    expect(saved).toHaveLength(1);
    expect(loadUsageLog()).toEqual([created]);
  });

  it('keeps only the most recent entries beyond the cap', () => {
    for (let index = 0; index < MAX_USAGE_ENTRIES + 5; index += 1) {
      appendUsageEntry(entry({ id: `e-${index}` }));
    }

    const loaded = loadUsageLog();
    expect(loaded).toHaveLength(MAX_USAGE_ENTRIES);
    expect(loaded.at(-1)?.id).toBe(`e-${MAX_USAGE_ENTRIES + 4}`);
    expect(loaded.some((item) => item.id === 'e-0')).toBe(false);
  });

  it('survives corrupt, non-array, and partially invalid storage', () => {
    localStorage.setItem(USAGE_KEY, '{not json');
    expect(loadUsageLog()).toEqual([]);

    localStorage.setItem(USAGE_KEY, '"a string"');
    expect(loadUsageLog()).toEqual([]);

    localStorage.setItem(USAGE_KEY, JSON.stringify([entry(), { garbage: true }, 42, null]));
    const loaded = loadUsageLog();
    expect(loaded).toHaveLength(1);
    expect(loaded[0]?.totalTokens).toBe(125);
  });

  it('clears the ledger', () => {
    appendUsageEntry(entry());
    clearUsageLog();

    expect(loadUsageLog()).toEqual([]);
    expect(localStorage.getItem(USAGE_KEY)).toBeNull();
  });
});

describe('summarizeUsage', () => {
  it('handles an empty ledger', () => {
    const summary = summarizeUsage([]);

    expect(summary).toMatchObject({
      requests: 0,
      llmRequests: 0,
      localRequests: 0,
      failures: 0,
      totalTokens: 0,
      avgLatencyMs: 0,
    });
    expect(summary.lastRequestAt).toBeUndefined();
    expect(summary.byProvider).toEqual([]);
  });

  it('aggregates tokens, failures, latency, and buckets', () => {
    const summary = summarizeUsage([
      entry({ totalTokens: 980, latencyMs: 400, provider: 'groq' }),
      entry({ totalTokens: 20, latencyMs: 200, provider: 'groq' }),
      entry({
        totalTokens: 0,
        latencyMs: 0,
        llmUsed: false,
        provider: 'none',
        model: 'deterministic engine (no model)',
      }),
      entry({
        totalTokens: 5,
        latencyMs: 100,
        ok: false,
        provider: 'gemini',
        model: 'gemini-2.0-flash',
      }),
    ]);

    expect(summary.requests).toBe(4);
    expect(summary.llmRequests).toBe(3);
    expect(summary.localRequests).toBe(1);
    expect(summary.failures).toBe(1);
    expect(summary.totalTokens).toBe(1005);
    // Average latency covers model calls only, never the local ones.
    expect(summary.avgLatencyMs).toBe(233);
    expect(summary.byProvider[0]?.label).toBe('groq');
    expect(summary.byProvider[0]?.totalTokens).toBe(1000);
    expect(summary.byProvider[0]?.requests).toBe(2);
    expect(summary.byModel.map((bucket) => bucket.label)).toContain('gemini-2.0-flash');
  });

  it('tracks the most recent request time', () => {
    const summary = summarizeUsage([
      entry({ timestamp: 100 }),
      entry({ timestamp: 500 }),
      entry({ timestamp: 250 }),
    ]);

    expect(summary.lastRequestAt).toBe(500);
  });
});

describe('display helpers', () => {
  it('formats token counts compactly', () => {
    expect(formatTokenCount(0)).toBe('0');
    expect(formatTokenCount(-5)).toBe('0');
    expect(formatTokenCount(999)).toBe('999');
    expect(formatTokenCount(1234)).toBe('1,234');
    expect(formatTokenCount(1_500_000)).toBe('1.5M');
  });

  it('never exposes a full API key', () => {
    expect(maskApiKey(undefined)).toBe('—');
    expect(maskApiKey('   ')).toBe('—');
    expect(maskApiKey('short')).toBe('sh••••');
    const masked = maskApiKey('gsk_1234567890abcdef');
    expect(masked.startsWith('gsk_')).toBe(true);
    expect(masked.endsWith('cdef')).toBe(true);
    expect(masked).not.toContain('1234567890');
  });
});
