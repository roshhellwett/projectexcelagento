import type { DecisionSource, LlmTelemetry, ProviderName } from '@excel-agent/agent';

export const USAGE_KEY = 'excel_agent_usage_v1';

/** Ring-buffer cap: the ledger is a diagnostic record, not an analytics store. */
export const MAX_USAGE_ENTRIES = 200;

/** Placeholder provider used when a turn was answered locally with no inference call. */
export const LOCAL_PROVIDER = 'none' as const;
export const LOCAL_MODEL_LABEL = 'deterministic engine (no model)';

export interface UsageEntry {
  id: string;
  timestamp: number;
  /** `'none'` when the turn never left the browser. */
  provider: ProviderName | typeof LOCAL_PROVIDER;
  model: string;
  source: DecisionSource;
  llmUsed: boolean;
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  latencyMs: number;
  ok: boolean;
  error?: string;
  /** Truncated copy of the prompt, kept short so the ledger stays small. */
  query: string;
}

export interface UsageBucket {
  label: string;
  requests: number;
  totalTokens: number;
}

export interface UsageSummary {
  requests: number;
  llmRequests: number;
  localRequests: number;
  failures: number;
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  avgLatencyMs: number;
  lastRequestAt?: number;
  byProvider: UsageBucket[];
  byModel: UsageBucket[];
}

const MAX_QUERY_PREVIEW = 120;

function storage(): Storage | undefined {
  try {
    return typeof localStorage === 'undefined' ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

let usageSequence = 0;

function nextUsageId(): string {
  usageSequence += 1;
  return `usage-${Date.now().toString(36)}-${usageSequence.toString(36)}`;
}

function isUsageEntry(value: unknown): value is UsageEntry {
  if (typeof value !== 'object' || value === null) return false;
  const entry = value as Partial<UsageEntry>;
  return (
    typeof entry.id === 'string' &&
    typeof entry.timestamp === 'number' &&
    typeof entry.model === 'string' &&
    typeof entry.totalTokens === 'number'
  );
}

/** Read the ledger, tolerating absent, corrupt, or partially-invalid storage. */
export function loadUsageLog(): UsageEntry[] {
  const store = storage();
  if (!store) return [];
  try {
    const raw = store.getItem(USAGE_KEY);
    if (!raw) return [];
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isUsageEntry).slice(-MAX_USAGE_ENTRIES);
  } catch {
    return [];
  }
}

export function saveUsageLog(entries: UsageEntry[]): void {
  const store = storage();
  if (!store) return;
  try {
    store.setItem(USAGE_KEY, JSON.stringify(entries.slice(-MAX_USAGE_ENTRIES)));
  } catch {
    // Quota or privacy mode: the ledger degrades to in-memory state.
  }
}

/** Append one entry, trimming the oldest records beyond the cap. */
export function appendUsageEntry(entry: UsageEntry): UsageEntry[] {
  const entries = [...loadUsageLog(), entry].slice(-MAX_USAGE_ENTRIES);
  saveUsageLog(entries);
  return entries;
}

export function clearUsageLog(): void {
  const store = storage();
  if (!store) return;
  try {
    store.removeItem(USAGE_KEY);
  } catch {
    // Ignore storage failures.
  }
}

/**
 * Build a ledger entry from one agent turn. When no telemetry is present the turn
 * was resolved locally, which is recorded explicitly so the usage page can prove
 * that zero tokens were sent.
 */
export function createUsageEntry(input: {
  query: string;
  source: DecisionSource;
  telemetry?: LlmTelemetry;
}): UsageEntry {
  const query = input.query.trim().slice(0, MAX_QUERY_PREVIEW);
  if (!input.telemetry) {
    return {
      id: nextUsageId(),
      timestamp: Date.now(),
      provider: LOCAL_PROVIDER,
      model: LOCAL_MODEL_LABEL,
      source: input.source,
      llmUsed: false,
      promptTokens: 0,
      completionTokens: 0,
      totalTokens: 0,
      latencyMs: 0,
      ok: true,
      query,
    };
  }

  const telemetry = input.telemetry;
  const promptTokens = telemetry.promptTokens ?? 0;
  const completionTokens = telemetry.completionTokens ?? 0;
  return {
    id: nextUsageId(),
    timestamp: Date.now(),
    provider: telemetry.provider,
    model: telemetry.model,
    source: input.source,
    llmUsed: true,
    promptTokens,
    completionTokens,
    totalTokens: telemetry.totalTokens ?? promptTokens + completionTokens,
    latencyMs: telemetry.latencyMs,
    ok: telemetry.ok,
    ...(telemetry.error ? { error: telemetry.error } : {}),
    query,
  };
}

function accumulate(buckets: Map<string, UsageBucket>, label: string, tokens: number): void {
  const existing = buckets.get(label);
  if (existing) {
    existing.requests += 1;
    existing.totalTokens += tokens;
  } else {
    buckets.set(label, { label, requests: 1, totalTokens: tokens });
  }
}

/** Aggregate ledger entries into the numbers shown on the usage page. */
export function summarizeUsage(entries: UsageEntry[]): UsageSummary {
  const providerBuckets = new Map<string, UsageBucket>();
  const modelBuckets = new Map<string, UsageBucket>();
  let promptTokens = 0;
  let completionTokens = 0;
  let totalTokens = 0;
  let failures = 0;
  let llmRequests = 0;
  let latencySum = 0;
  let latencySamples = 0;
  let lastRequestAt: number | undefined;

  for (const entry of entries) {
    promptTokens += entry.promptTokens;
    completionTokens += entry.completionTokens;
    totalTokens += entry.totalTokens;
    if (!entry.ok) failures += 1;
    if (entry.llmUsed) {
      llmRequests += 1;
      latencySum += entry.latencyMs;
      latencySamples += 1;
    }
    if (lastRequestAt === undefined || entry.timestamp > lastRequestAt) {
      lastRequestAt = entry.timestamp;
    }
    accumulate(providerBuckets, entry.provider, entry.totalTokens);
    accumulate(modelBuckets, entry.model, entry.totalTokens);
  }

  const byTokens = (a: UsageBucket, b: UsageBucket): number =>
    b.totalTokens - a.totalTokens || a.label.localeCompare(b.label);

  return {
    requests: entries.length,
    llmRequests,
    localRequests: entries.length - llmRequests,
    failures,
    promptTokens,
    completionTokens,
    totalTokens,
    avgLatencyMs: latencySamples === 0 ? 0 : Math.round(latencySum / latencySamples),
    ...(lastRequestAt === undefined ? {} : { lastRequestAt }),
    byProvider: [...providerBuckets.values()].sort(byTokens),
    byModel: [...modelBuckets.values()].sort(byTokens),
  };
}

/** Compact token display: 1234 -> "1,234", 1500000 -> "1.5M". */
export function formatTokenCount(value: number): string {
  if (!Number.isFinite(value) || value <= 0) return '0';
  if (value >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
  return value.toLocaleString('en-US');
}

/** Never render a full API key; keep a recognisable prefix/suffix only. */
export function maskApiKey(key: string | undefined | null): string {
  const trimmed = (key ?? '').trim();
  if (!trimmed) return '—';
  if (trimmed.length <= 8) return `${trimmed.slice(0, 2)}${'•'.repeat(4)}`;
  return `${trimmed.slice(0, 4)}${'•'.repeat(6)}${trimmed.slice(-4)}`;
}

/** Human host label for each provider, shown for transparency in the UI. */
export const PROVIDER_HOSTS: Record<ProviderName, string> = {
  groq: 'api.groq.com',
  openrouter: 'openrouter.ai',
  gemini: 'generativelanguage.googleapis.com',
};
