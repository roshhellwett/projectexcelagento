export const FEEDBACK_KEY = 'excel_agent_feedback_v1';
export const MAX_FEEDBACK_ENTRIES = 100;
export const MAX_FEEDBACK_NOTE_LENGTH = 240;

export type FeedbackResult = 'completed' | 'blocked';
export type FeedbackRating = 'helpful' | 'mixed' | 'not_helpful';
export type FeedbackCategory =
  'cleanup' | 'formatting' | 'analysis' | 'sorting_filtering' | 'other';

export interface FeedbackInput {
  result: FeedbackResult;
  rating: FeedbackRating;
  category: FeedbackCategory;
  note?: string;
}

export interface FeedbackEntry extends FeedbackInput {
  version: 1;
  id: string;
  createdAt: number;
}

export interface FeedbackSummary {
  total: number;
  helpful: number;
  mixed: number;
  notHelpful: number;
  completed: number;
  blocked: number;
  byCategory: Array<{ category: FeedbackCategory; count: number }>;
  lastSubmittedAt?: number;
}

export type FeedbackPersistence = 'saved-in-browser' | 'session-only';
export type FeedbackStorageIssue = 'read' | 'write' | 'clear';

export interface FeedbackStorageStatus {
  persistence: FeedbackPersistence;
  issue?: FeedbackStorageIssue;
}

export interface FeedbackAppendResult {
  entries: FeedbackEntry[];
  status: FeedbackStorageStatus;
}

export interface FeedbackClearResult {
  entries: FeedbackEntry[];
  cleared: boolean;
  status: FeedbackStorageStatus;
}

/** The export intentionally omits notes and any context that was never part of the schema. */
export interface FeedbackExportEntry {
  id: string;
  createdAt: number;
  result: FeedbackResult;
  rating: FeedbackRating;
  category: FeedbackCategory;
}

export interface FeedbackExport {
  version: 1;
  exportedAt: number;
  summary: FeedbackSummary;
  entries: FeedbackExportEntry[];
}

export const FEEDBACK_CATEGORY_LABELS: Record<FeedbackCategory, string> = {
  cleanup: 'Clean-up & duplicates',
  formatting: 'Formatting & dates',
  analysis: 'Analysis & insights',
  sorting_filtering: 'Sorting & filtering',
  other: 'Something else',
};

const FEEDBACK_RESULTS = new Set<FeedbackResult>(['completed', 'blocked']);
const FEEDBACK_RATINGS = new Set<FeedbackRating>(['helpful', 'mixed', 'not_helpful']);
const FEEDBACK_CATEGORIES = new Set<FeedbackCategory>([
  'cleanup',
  'formatting',
  'analysis',
  'sorting_filtering',
  'other',
]);

function storage(): Storage | undefined {
  try {
    return typeof localStorage === 'undefined' ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

let feedbackSequence = 0;

let inMemoryFeedbackLog: FeedbackEntry[] = [];
let inMemoryPersistence: FeedbackPersistence = 'saved-in-browser';
let feedbackStorageStatus: FeedbackStorageStatus = {
  persistence: 'saved-in-browser',
};

function nextFeedbackId(): string {
  feedbackSequence += 1;
  return `feedback-${Date.now().toString(36)}-${feedbackSequence.toString(36)}`;
}

/**
 * Keep the optional note useful without turning the feedback ledger into a place for credentials.
 * The UI also explains that notes should stay generic; these redactions are a second local guard.
 */
export function sanitizeFeedbackNote(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const cleaned = value
    .replace(/\bBearer\s+[^\s]+/gi, '[redacted]')
    .replace(/\b(?:sk|rk|pk|ghp|github_pat|xox[baprs])[-_][A-Za-z0-9_-]{8,}/g, '[redacted]')
    .replace(/\bAIza[0-9A-Za-z_-]{20,}\b/g, '[redacted]')
    .replace(/\b((?:api[_ -]?key|token|password|secret))\s*[:=]\s*[^\s,;]+/gi, '$1: [redacted]')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_FEEDBACK_NOTE_LENGTH);
  return cleaned || undefined;
}

function isFeedbackEntry(value: unknown): value is FeedbackEntry {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const entry = value as Partial<FeedbackEntry>;
  return (
    entry.version === 1 &&
    typeof entry.id === 'string' &&
    entry.id.length > 0 &&
    entry.id.length <= 100 &&
    typeof entry.createdAt === 'number' &&
    Number.isFinite(entry.createdAt) &&
    entry.createdAt >= 0 &&
    FEEDBACK_RESULTS.has(entry.result as FeedbackResult) &&
    FEEDBACK_RATINGS.has(entry.rating as FeedbackRating) &&
    FEEDBACK_CATEGORIES.has(entry.category as FeedbackCategory) &&
    (entry.note === undefined ||
      (typeof entry.note === 'string' && entry.note.length <= MAX_FEEDBACK_NOTE_LENGTH))
  );
}

function normalizeFeedbackEntry(value: unknown): FeedbackEntry | undefined {
  if (!isFeedbackEntry(value)) return undefined;
  const entry = value as FeedbackEntry;
  const note = sanitizeFeedbackNote(entry.note);
  return {
    version: 1,
    id: entry.id,
    createdAt: entry.createdAt,
    result: entry.result,
    rating: entry.rating,
    category: entry.category,
    ...(note ? { note } : {}),
  };
}

function normalizeFeedbackEntries(entries: unknown): FeedbackEntry[] {
  if (!Array.isArray(entries)) return [];
  return entries
    .map(normalizeFeedbackEntry)
    .filter((entry): entry is FeedbackEntry => entry !== undefined)
    .slice(-MAX_FEEDBACK_ENTRIES);
}

function setFeedbackStorageStatus(
  persistence: FeedbackPersistence,
  issue?: FeedbackStorageIssue,
): FeedbackStorageStatus {
  feedbackStorageStatus = { persistence, ...(issue ? { issue } : {}) };
  return feedbackStorageStatus;
}

function copyFeedbackStorageStatus(): FeedbackStorageStatus {
  return { ...feedbackStorageStatus };
}

function copyFeedbackEntries(entries: FeedbackEntry[]): FeedbackEntry[] {
  return entries.map((entry) => ({
    version: 1,
    id: entry.id,
    createdAt: entry.createdAt,
    result: entry.result,
    rating: entry.rating,
    category: entry.category,
    ...(entry.note ? { note: entry.note } : {}),
  }));
}

function mergeFeedbackEntries(
  persistedEntries: FeedbackEntry[],
  sessionEntries: FeedbackEntry[],
): FeedbackEntry[] {
  const byId = new Map<string, FeedbackEntry>();
  for (const entry of persistedEntries) byId.set(entry.id, entry);
  for (const entry of sessionEntries) byId.set(entry.id, entry);
  return [...byId.values()].slice(-MAX_FEEDBACK_ENTRIES);
}

/** Read only the small, structured local feedback ledger; malformed entries are ignored. */
export function loadFeedbackLog(): FeedbackEntry[] {
  const store = storage();
  if (!store) {
    inMemoryPersistence = 'session-only';
    setFeedbackStorageStatus('session-only', 'read');
    return copyFeedbackEntries(inMemoryFeedbackLog);
  }
  let raw: string | null;
  try {
    raw = store.getItem(FEEDBACK_KEY);
  } catch {
    inMemoryPersistence = 'session-only';
    setFeedbackStorageStatus('session-only', 'read');
    return copyFeedbackEntries(inMemoryFeedbackLog);
  }
  try {
    const persistedEntries = raw ? normalizeFeedbackEntries(JSON.parse(raw)) : [];
    const entries =
      inMemoryPersistence === 'session-only'
        ? mergeFeedbackEntries(persistedEntries, inMemoryFeedbackLog)
        : persistedEntries;
    inMemoryFeedbackLog = entries;
    if (inMemoryPersistence !== 'session-only') {
      inMemoryPersistence = 'saved-in-browser';
      setFeedbackStorageStatus('saved-in-browser');
    } else {
      setFeedbackStorageStatus('session-only');
    }
    return copyFeedbackEntries(inMemoryFeedbackLog);
  } catch {
    // A malformed value is treated as an empty browser ledger. It must never cause stale entries
    // or unknown fields from a previous read to be copied back into a later save.
    if (inMemoryPersistence === 'session-only') {
      setFeedbackStorageStatus('session-only', 'read');
      return copyFeedbackEntries(inMemoryFeedbackLog);
    }
    inMemoryFeedbackLog = [];
    inMemoryPersistence = 'saved-in-browser';
    setFeedbackStorageStatus('saved-in-browser');
    return [];
  }
}

export function saveFeedbackLog(entries: FeedbackEntry[]): void {
  const normalizedEntries = normalizeFeedbackEntries(entries);
  inMemoryFeedbackLog = normalizedEntries;
  const store = storage();
  if (!store) {
    inMemoryPersistence = 'session-only';
    setFeedbackStorageStatus('session-only', 'write');
    return;
  }
  const serialized = JSON.stringify(normalizedEntries);
  try {
    // Serialize only the known schema fields. This also strips fields injected into malformed
    // entries before they can become part of a later saved ledger or export.
    store.setItem(FEEDBACK_KEY, serialized);
    if (store.getItem(FEEDBACK_KEY) !== serialized)
      throw new Error('Feedback storage did not retain the new value');
    inMemoryPersistence = 'saved-in-browser';
    setFeedbackStorageStatus('saved-in-browser');
  } catch {
    // Retain the complete current-session ledger when quota/private-mode failures prevent a write.
    inMemoryPersistence = 'session-only';
    setFeedbackStorageStatus('session-only', 'write');
  }
}

export function getFeedbackStorageStatus(): FeedbackStorageStatus {
  return copyFeedbackStorageStatus();
}

export function appendFeedbackEntryWithStatus(input: FeedbackInput): FeedbackAppendResult {
  const note = sanitizeFeedbackNote(input.note);
  const entry: FeedbackEntry = {
    version: 1,
    id: nextFeedbackId(),
    createdAt: Date.now(),
    result: input.result,
    rating: input.rating,
    category: input.category,
    ...(note ? { note } : {}),
  };
  const entries = [...loadFeedbackLog(), entry].slice(-MAX_FEEDBACK_ENTRIES);
  saveFeedbackLog(entries);
  return { entries: copyFeedbackEntries(inMemoryFeedbackLog), status: getFeedbackStorageStatus() };
}

export function appendFeedbackEntry(input: FeedbackInput): FeedbackEntry[] {
  return appendFeedbackEntryWithStatus(input).entries;
}

export function clearFeedbackLogWithStatus(): FeedbackClearResult {
  const store = storage();
  if (!store) {
    inMemoryPersistence = 'session-only';
    setFeedbackStorageStatus('session-only', 'clear');
    return {
      entries: copyFeedbackEntries(inMemoryFeedbackLog),
      cleared: false,
      status: getFeedbackStorageStatus(),
    };
  }
  try {
    store.removeItem(FEEDBACK_KEY);
    // A successful call is not enough to claim deletion when a restricted or test storage
    // implementation leaves the value behind (or cannot be read back for verification).
    if (store.getItem(FEEDBACK_KEY) !== null)
      throw new Error('Feedback storage still contains data');
    inMemoryFeedbackLog = [];
    inMemoryPersistence = 'saved-in-browser';
    setFeedbackStorageStatus('saved-in-browser');
    return { entries: [], cleared: true, status: getFeedbackStorageStatus() };
  } catch {
    // Keep the entries visible: a failed remove must not make the UI claim that browser data is
    // gone. The user can retry after storage permissions change.
    setFeedbackStorageStatus(inMemoryPersistence, 'clear');
    return {
      entries: copyFeedbackEntries(inMemoryFeedbackLog),
      cleared: false,
      status: getFeedbackStorageStatus(),
    };
  }
}

export function clearFeedbackLog(): void {
  void clearFeedbackLogWithStatus();
}

export function summarizeFeedback(entries: FeedbackEntry[]): FeedbackSummary {
  const counts = new Map<FeedbackCategory, number>();
  let helpful = 0;
  let mixed = 0;
  let notHelpful = 0;
  let completed = 0;
  let blocked = 0;
  let lastSubmittedAt: number | undefined;

  for (const entry of entries) {
    if (entry.rating === 'helpful') helpful += 1;
    else if (entry.rating === 'mixed') mixed += 1;
    else notHelpful += 1;

    if (entry.result === 'completed') completed += 1;
    else blocked += 1;

    counts.set(entry.category, (counts.get(entry.category) ?? 0) + 1);
    if (lastSubmittedAt === undefined || entry.createdAt > lastSubmittedAt)
      lastSubmittedAt = entry.createdAt;
  }

  return {
    total: entries.length,
    helpful,
    mixed,
    notHelpful,
    completed,
    blocked,
    byCategory: [...counts.entries()]
      .map(([category, count]) => ({ category, count }))
      .sort((a, b) => b.count - a.count || a.category.localeCompare(b.category)),
    ...(lastSubmittedAt === undefined ? {} : { lastSubmittedAt }),
  };
}

export function createFeedbackExport(
  entries: FeedbackEntry[],
  exportedAt = Date.now(),
): FeedbackExport {
  const safeFeedbackEntries = entries
    .map(normalizeFeedbackEntry)
    .filter((entry): entry is FeedbackEntry => entry !== undefined);
  const safeEntries = safeFeedbackEntries.map(
    ({ id, createdAt, result, rating, category }): FeedbackExportEntry => ({
      id,
      createdAt,
      result,
      rating,
      category,
    }),
  );
  return {
    version: 1,
    exportedAt,
    summary: summarizeFeedback(safeFeedbackEntries),
    entries: safeEntries,
  };
}

export function serializeFeedbackExport(entries: FeedbackEntry[], exportedAt?: number): string {
  return JSON.stringify(createFeedbackExport(entries, exportedAt), null, 2);
}

/** Download is intentionally local and user-triggered; no network or analytics call is made. */
export function downloadFeedbackSummary(entries: FeedbackEntry[]): boolean {
  if (
    typeof document === 'undefined' ||
    typeof URL === 'undefined' ||
    typeof URL.createObjectURL !== 'function'
  ) {
    return false;
  }
  const blob = new Blob([serializeFeedbackExport(entries)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `feedback-summary-${new Date().toISOString().slice(0, 10)}.json`;
  link.setAttribute('aria-label', 'Download aggregate feedback summary');
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  return true;
}

/** Use only a non-sensitive operation label to suggest a category; the label is never persisted. */
export function feedbackCategoryForOperation(
  operation: string | undefined,
  kind?: 'analysis' | 'action' | 'plan',
): FeedbackCategory {
  const name = (operation ?? '').toLowerCase();
  if (/duplicate|clean|normalize|trim|empty|coerce/.test(name)) return 'cleanup';
  if (/format|date|style|number/.test(name)) return 'formatting';
  if (/sort|filter|row/.test(name)) return 'sorting_filtering';
  if (kind === 'analysis' || /analy|stat|summary|reconcile/.test(name)) return 'analysis';
  return 'other';
}
