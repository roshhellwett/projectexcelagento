import type { MemoryRecord, MemoryStore } from './types.js';

const STOP_WORDS = new Set([
  'the',
  'a',
  'an',
  'to',
  'of',
  'in',
  'on',
  'for',
  'and',
  'or',
  'my',
  'me',
  'please',
  'all',
  'is',
  'are',
  'it',
  'this',
  'that',
  'with',
  'from',
  'by',
  'do',
  'can',
  'you',
  'i',
  'rows',
  'row',
  'column',
  'columns',
  'col',
  'sheet',
]);

/** Lowercase, strip punctuation, drop stop words, collapse whitespace. */
export function normalizeQuery(query: string): string {
  return query
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter((token) => token.length > 0 && !STOP_WORDS.has(token))
    .join(' ')
    .trim();
}

function tokenSet(value: string): Set<string> {
  return new Set(value.split(' ').filter(Boolean));
}

/** Jaccard index over query tokens, with a small bonus for exact normalized equality. */
export function querySimilarity(left: string, right: string): number {
  const a = normalizeQuery(left);
  const b = normalizeQuery(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  const setA = tokenSet(a);
  const setB = tokenSet(b);
  let intersection = 0;
  for (const token of setA) {
    if (setB.has(token)) intersection += 1;
  }
  const union = setA.size + setB.size - intersection;
  return union === 0 ? 0 : intersection / union;
}

let sequence = 0;

function nextId(): string {
  sequence += 1;
  return `mem_${Date.now().toString(36)}_${sequence.toString(36)}`;
}

export class InMemoryMemoryStore implements MemoryStore {
  private records: MemoryRecord[] = [];

  private readonly maxRecords: number;

  constructor(options: { maxRecords?: number } = {}) {
    this.maxRecords = options.maxRecords ?? 500;
  }

  remember(
    record: Omit<MemoryRecord, 'id' | 'createdAt' | 'lastUsedAt' | 'successes' | 'failures'>,
  ): MemoryRecord {
    const key = normalizeQuery(record.key || record.rawQuery);
    const existing = this.records.find(
      (entry) =>
        entry.key === key &&
        entry.sheetName === record.sheetName &&
        entry.operation === record.operation,
    );
    if (existing) {
      existing.args = { ...record.args };
      existing.rawQuery = record.rawQuery;
      existing.lastUsedAt = Date.now();
      this.records.splice(this.records.indexOf(existing), 1);
      this.records.push(existing);
      return existing;
    }
    const created: MemoryRecord = {
      ...record,
      key,
      id: nextId(),
      successes: 0,
      failures: 0,
      createdAt: Date.now(),
      lastUsedAt: Date.now(),
    };
    this.records.push(created);
    if (this.records.length > this.maxRecords) {
      this.records.sort((a, b) => this.score(b) - this.score(a));
      this.records = this.records.slice(0, this.maxRecords);
    }
    return created;
  }

  recordOutcome(operation: string, sheetName: string, success: boolean, key?: string): void {
    const normalizedKey = key === undefined ? undefined : normalizeQuery(key);
    for (let index = this.records.length - 1; index >= 0; index -= 1) {
      const record = this.records[index];
      if (record?.operation !== operation || record.sheetName !== sheetName) continue;
      if (normalizedKey !== undefined && record.key !== normalizedKey) continue;
      if (success) record.successes += 1;
      else record.failures += 1;
      record.lastUsedAt = Date.now();
      return;
    }
  }

  retrieve(query: string, sheetName: string, threshold = 0.6): MemoryRecord | undefined {
    let best: MemoryRecord | undefined;
    let bestScore = threshold;
    for (const record of this.records) {
      if (record.sheetName !== sheetName) continue;
      // A record that has never succeeded is untrusted until proven.
      const confidence = this.score(record);
      if (confidence <= 0) continue;
      const similarity = querySimilarity(query, record.key);
      const combined = similarity * confidence;
      if (similarity >= threshold && combined >= bestScore) {
        bestScore = combined;
        best = record;
      }
    }
    if (best) best.lastUsedAt = Date.now();
    return best;
  }

  entries(): MemoryRecord[] {
    return this.records.map((record) => ({ ...record, args: { ...record.args } }));
  }

  confidenceOf(record: MemoryRecord): number {
    return this.score(record);
  }

  clear(): void {
    this.records = [];
  }

  toJSON(): string {
    return JSON.stringify({ version: 1, records: this.records });
  }

  load(serialized: string): void {
    try {
      const parsed = JSON.parse(serialized) as { records?: MemoryRecord[] };
      if (Array.isArray(parsed.records)) {
        this.records = parsed.records.filter(
          (record): record is MemoryRecord =>
            typeof record === 'object' &&
            record !== null &&
            typeof record.key === 'string' &&
            typeof record.operation === 'string',
        );
      }
    } catch {
      // Corrupt memory is ignored rather than crashing the session.
    }
  }

  /** Reliability in [0, 1]: reliability ratio damped by low observation counts. */
  private score(record: MemoryRecord): number {
    const total = record.successes + record.failures;
    if (total === 0) return 0.5;
    const ratio = record.successes / total;
    const confidence = Math.min(1, total / 3);
    return ratio * (0.4 + 0.6 * confidence);
  }
}

export function createMemoryStore(serialized?: string): InMemoryMemoryStore {
  const store = new InMemoryMemoryStore();
  if (serialized) store.load(serialized);
  return store;
}
