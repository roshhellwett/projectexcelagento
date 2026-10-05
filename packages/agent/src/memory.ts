import type { Workbook } from '@excel-agent/engine';
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

/** Levenshtein distance between two strings */
export function levenshteinDistance(s1: string, s2: string): number {
  if (s1 === s2) return 0;
  if (!s1.length) return s2.length;
  if (!s2.length) return s1.length;

  const m = s1.length;
  const n = s2.length;
  const d: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) d[i]![0] = i;
  for (let j = 0; j <= n; j++) d[0]![j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
      d[i]![j] = Math.min(d[i - 1]![j]! + 1, d[i]![j - 1]! + 1, d[i - 1]![j - 1]! + cost);
    }
  }

  return d[m]![n]!;
}

const CANONICAL_MAP: Record<string, string> = {
  // Duplicates
  dupli: 'duplicate',
  dups: 'duplicate',
  dupe: 'duplicate',
  dupes: 'duplicate',
  duplicatez: 'duplicate',
  dublicate: 'duplicate',
  dublicats: 'duplicate',
  repeted: 'duplicate',
  repeated: 'duplicate',
  // Remove / delete
  remov: 'remove',
  rm: 'remove',
  del: 'delete',
  delet: 'delete',
  dlt: 'delete',
  hatao: 'delete',
  nikalo: 'delete',
  // Sort
  sorrt: 'sort',
  srot: 'sort',
  shrt: 'sort',
  arrang: 'sort',
  ordr: 'sort',
  // Amount / value
  amunt: 'amount',
  amnt: 'amount',
  amont: 'amount',
  amt: 'amount',
  val: 'value',
  vals: 'values',
  paise: 'amount',
  // Direction
  decs: 'descending',
  desc: 'descending',
  desending: 'descending',
  asce: 'ascending',
  asc: 'ascending',
  acending: 'ascending',
  // Format / date
  formt: 'format',
  farmat: 'format',
  mak: 'make',
  convrt: 'convert',
  tarikh: 'date',
  dat: 'date',
  dats: 'date',
  dte: 'date',
  dt: 'date',
  // Columns / rows
  colum: 'column',
  colm: 'column',
  clm: 'column',
  cols: 'columns',
  clmn: 'column',
  // Calculation / aggregate
  calclate: 'calculate',
  clculate: 'calculate',
  clc: 'calculate',
  hisab: 'total',
  hisaab: 'total',
  totl: 'total',
  ttl: 'total',
  // Text
  cap: 'capitalize',
  caps: 'capitalize',
  capital: 'capitalize',
  capitaliz: 'capitalize',
  capitalise: 'capitalize',
  spac: 'space',
  spce: 'space',
  safai: 'clean',
  khatam: 'clean',
};

/** Canonicalize a single token to handle typos, slang, and phonetic spelling. */
export function canonicalizeToken(token: string): string {
  const lower = token.toLowerCase().trim();
  if (CANONICAL_MAP[lower]) return CANONICAL_MAP[lower];
  return lower;
}

/** Canonicalize a whole query string by replacing typos, slang, and phonetic tokens with canonical domain terms. */
export function canonicalizeSpreadsheetQuery(query: string): string {
  if (!query || typeof query !== 'string') return '';
  return query
    .split(/\b/)
    .map((part) => {
      if (/^[a-zA-Z0-9]+$/.test(part)) {
        return canonicalizeToken(part);
      }
      return part;
    })
    .join('');
}

/** Check if two tokens are equivalent either directly, canonicalized, prefix-matched, or fuzzy Levenshtein. */
export function isTokenMatch(a: string, b: string): boolean {
  if (a === b) return true;
  const canA = canonicalizeToken(a);
  const canB = canonicalizeToken(b);
  if (canA === canB) return true;

  // Prefix matching for prefixes of length >= 4 (e.g. dupli -> duplicate, decs -> descending)
  if (canA.length >= 4 && canB.startsWith(canA)) return true;
  if (canB.length >= 4 && canA.startsWith(canB)) return true;

  // Levenshtein distance: 1 edit for len >= 4, 2 edits for len >= 6
  const maxDist =
    Math.min(canA.length, canB.length) >= 6 ? 2 : Math.min(canA.length, canB.length) >= 4 ? 1 : 0;
  if (maxDist > 0 && Math.abs(canA.length - canB.length) <= maxDist) {
    return levenshteinDistance(canA, canB) <= maxDist;
  }

  return false;
}

function tokenSet(value: string): Set<string> {
  return new Set(value.split(' ').filter(Boolean));
}

/**
 * Typo-tolerant and fuzzy query similarity.
 * Evaluates semantic and token equivalence so queries with typos, slang, and
 * non-native grammar (e.g. "remov dupli" vs "remove duplicate") score accurately.
 */
export function querySimilarity(left: string, right: string): number {
  const a = normalizeQuery(left);
  const b = normalizeQuery(right);
  if (!a || !b) return 0;
  if (a === b) return 1;

  const tokensA = Array.from(tokenSet(a));
  const tokensB = Array.from(tokenSet(b));

  const matchedB = new Set<number>();
  let intersection = 0;

  for (const tA of tokensA) {
    for (let idx = 0; idx < tokensB.length; idx++) {
      if (matchedB.has(idx)) continue;
      const tB = tokensB[idx]!;
      if (isTokenMatch(tA, tB)) {
        matchedB.add(idx);
        intersection += 1;
        break;
      }
    }
  }

  const union = tokensA.length + tokensB.length - intersection;
  return union === 0 ? 0 : intersection / union;
}

/**
 * Identifies a sheet's shape by its header names in order.
 *
 * Column letters are positional, so any argument naming a column stops meaning what it meant
 * as soon as the sheet is reshaped. Fingerprinting the headers turns that invisible hazard into
 * a check that can be run before a replay. The row count is deliberately excluded: appending
 * rows is routine and must not invalidate what was learned.
 */
export function schemaFingerprint(headers: string[]): string {
  return headers.map((header) => header.trim().toLowerCase()).join('\u0001');
}

/** Reads the header row of a sheet, tolerating the ragged rows an import can produce. */
export function sheetFingerprint(workbook: Workbook, sheetName: string): string | undefined {
  const sheet = workbook.sheets.find((candidate) => candidate.name === sheetName);
  const headerRow = sheet?.rows[0];
  if (!headerRow) return undefined;
  return schemaFingerprint(headerRow.map((cell) => String(cell.value ?? '')));
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

  /**
   * Retrieves a replayable record, refusing any whose learned arguments no longer fit the sheet.
   *
   * This is separate from `retrieve` because staleness is a property of the data, not of the
   * query: the words may match perfectly while the column letters they referred to have shifted.
   * Callers must use this before replaying a learned action against real user data.
   */
  retrieveForWorkbook(
    query: string,
    workbook: Workbook,
    sheetName: string,
    threshold = 0.6,
  ): MemoryRecord | undefined {
    const best = this.retrieve(query, sheetName, threshold);
    if (!best) return undefined;

    // A record learned before fingerprints existed cannot be checked, so it is not replayed.
    // Trusting an unverifiable record is exactly the failure this guards against.
    if (typeof best.schemaFingerprint !== 'string') return undefined;

    const current = sheetFingerprint(workbook, sheetName);
    if (current === undefined || current !== best.schemaFingerprint) return undefined;

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
