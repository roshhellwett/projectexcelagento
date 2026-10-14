// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';

import {
  FEEDBACK_KEY,
  MAX_FEEDBACK_ENTRIES,
  appendFeedbackEntry,
  appendFeedbackEntryWithStatus,
  clearFeedbackLog,
  clearFeedbackLogWithStatus,
  createFeedbackExport,
  feedbackCategoryForOperation,
  getFeedbackStorageStatus,
  loadFeedbackLog,
  sanitizeFeedbackNote,
  summarizeFeedback,
} from './feedback.js';

beforeEach(() => {
  vi.restoreAllMocks();
  localStorage.clear();
  clearFeedbackLogWithStatus();
  localStorage.clear();
});

describe('local feedback persistence', () => {
  it('stores only the small structured signal and redacts common credentials in notes', () => {
    const entries = appendFeedbackEntry({
      result: 'completed',
      rating: 'mixed',
      category: 'cleanup',
      note: 'The preview needed work. apiKey=sk-or-v1-super-secret-value',
    });

    expect(entries).toHaveLength(1);
    expect(loadFeedbackLog()).toEqual(entries);
    expect(entries[0]).toMatchObject({
      version: 1,
      result: 'completed',
      rating: 'mixed',
      category: 'cleanup',
    });
    expect(entries[0]).not.toHaveProperty('request');
    expect(entries[0]).not.toHaveProperty('workbook');
    expect(entries[0]?.note).not.toContain('sk-or-v1-super-secret-value');
    expect(localStorage.getItem(FEEDBACK_KEY)).not.toContain('sk-or-v1-super-secret-value');
  });

  it('tolerates corrupt storage and trims the ledger to the newest responses', () => {
    localStorage.setItem(FEEDBACK_KEY, '{not json');
    expect(loadFeedbackLog()).toEqual([]);

    for (let index = 0; index < MAX_FEEDBACK_ENTRIES + 4; index += 1) {
      appendFeedbackEntry({
        result: index % 2 === 0 ? 'completed' : 'blocked',
        rating: 'helpful',
        category: 'analysis',
        note: `response ${index}`,
      });
    }

    const loaded = loadFeedbackLog();
    expect(loaded).toHaveLength(MAX_FEEDBACK_ENTRIES);
    expect(loaded[0]?.note).toBe('response 4');
    expect(loaded.at(-1)?.note).toBe('response 103');
  });

  it('clears local feedback without affecting other ledgers', () => {
    appendFeedbackEntry({
      result: 'blocked',
      rating: 'not_helpful',
      category: 'other',
    });
    localStorage.setItem('unrelated-local-setting', 'keep');

    clearFeedbackLog();

    expect(loadFeedbackLog()).toEqual([]);
    expect(localStorage.getItem(FEEDBACK_KEY)).toBeNull();
    expect(localStorage.getItem('unrelated-local-setting')).toBe('keep');
  });

  it('retains repeat submissions in memory when browser reads and writes are denied', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('read denied');
    });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('write denied');
    });

    const first = appendFeedbackEntryWithStatus({
      result: 'completed',
      rating: 'helpful',
      category: 'analysis',
    });
    const second = appendFeedbackEntryWithStatus({
      result: 'blocked',
      rating: 'mixed',
      category: 'other',
    });

    expect(first.status.persistence).toBe('session-only');
    expect(second.status).toMatchObject({ persistence: 'session-only', issue: 'write' });
    expect(second.entries).toHaveLength(2);
    expect(second.entries.map(({ result }) => result)).toEqual(['completed', 'blocked']);
    expect(loadFeedbackLog()).toHaveLength(2);
  });

  it('retains repeat submissions when only browser writes are denied', () => {
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('write denied');
    });

    appendFeedbackEntry({ result: 'completed', rating: 'helpful', category: 'cleanup' });
    const entries = appendFeedbackEntry({
      result: 'completed',
      rating: 'mixed',
      category: 'formatting',
    });

    expect(entries).toHaveLength(2);
    expect(getFeedbackStorageStatus()).toMatchObject({
      persistence: 'session-only',
      issue: 'write',
    });
  });

  it('does not claim browser deletion when removeItem is denied', () => {
    appendFeedbackEntry({ result: 'blocked', rating: 'not_helpful', category: 'other' });
    vi.spyOn(Storage.prototype, 'removeItem').mockImplementation(() => {
      throw new Error('remove denied');
    });

    const result = clearFeedbackLogWithStatus();

    expect(result.cleared).toBe(false);
    expect(result.status).toMatchObject({ persistence: 'saved-in-browser', issue: 'clear' });
    expect(result.entries).toHaveLength(1);
    expect(loadFeedbackLog()).toHaveLength(1);
  });

  it('strips unknown fields while reading and re-saving malformed entries', () => {
    localStorage.setItem(
      FEEDBACK_KEY,
      JSON.stringify([
        {
          version: 1,
          id: 'known-entry',
          createdAt: 123,
          result: 'completed',
          rating: 'helpful',
          category: 'analysis',
          note: 'A useful note',
          workbook: 'secret workbook contents',
          prompt: 'secret prompt',
          injected: { accountToken: 'secret' },
        },
      ]),
    );

    const loaded = loadFeedbackLog();
    expect(loaded).toEqual([
      {
        version: 1,
        id: 'known-entry',
        createdAt: 123,
        result: 'completed',
        rating: 'helpful',
        category: 'analysis',
        note: 'A useful note',
      },
    ]);

    appendFeedbackEntry({ result: 'blocked', rating: 'mixed', category: 'other' });
    const stored = localStorage.getItem(FEEDBACK_KEY) ?? '';
    expect(stored).not.toContain('secret workbook contents');
    expect(stored).not.toContain('secret prompt');
    expect(stored).not.toContain('accountToken');
    expect(JSON.parse(stored)[0]).not.toHaveProperty('injected');
  });
});

describe('feedback summaries and privacy helpers', () => {
  it('summarizes result, rating, and workflow categories without retaining notes', () => {
    const entries = [
      appendFeedbackEntry({ result: 'completed', rating: 'helpful', category: 'cleanup' })[0]!,
      appendFeedbackEntry({ result: 'blocked', rating: 'not_helpful', category: 'analysis' })[1]!,
      appendFeedbackEntry({ result: 'completed', rating: 'mixed', category: 'cleanup' })[2]!,
    ];

    expect(summarizeFeedback(entries)).toMatchObject({
      total: 3,
      helpful: 1,
      mixed: 1,
      notHelpful: 1,
      completed: 2,
      blocked: 1,
      byCategory: [
        { category: 'cleanup', count: 2 },
        { category: 'analysis', count: 1 },
      ],
    });
  });

  it('suggests categories from operation labels without storing the labels', () => {
    expect(feedbackCategoryForOperation('delete_duplicates')).toBe('cleanup');
    expect(feedbackCategoryForOperation('format_dates')).toBe('formatting');
    expect(feedbackCategoryForOperation('sort_range')).toBe('sorting_filtering');
    expect(feedbackCategoryForOperation(undefined, 'analysis')).toBe('analysis');
    expect(feedbackCategoryForOperation('unknown_operation')).toBe('other');
  });

  it('bounds and normalizes optional notes', () => {
    expect(sanitizeFeedbackNote('  clear   preview\nplease  ')).toBe('clear preview please');
    expect(sanitizeFeedbackNote('x'.repeat(300))).toHaveLength(240);
    expect(sanitizeFeedbackNote('   ')).toBeUndefined();
  });

  it('exports only the aggregate signal and safe per-entry metadata', () => {
    const entry = appendFeedbackEntry({
      result: 'completed',
      rating: 'helpful',
      category: 'cleanup',
      note: 'User-entered workbook detail that must stay local',
    })[0]!;
    const injected = {
      ...entry,
      note: 'private note',
      workbook: 'workbook contents',
      prompt: 'prompt contents',
      fileName: 'secret.xlsx',
      accountSecret: 'secret-token',
    } as typeof entry & Record<string, unknown>;

    const exported = createFeedbackExport([injected], 456);

    expect(exported).toMatchObject({ version: 1, exportedAt: 456 });
    expect(exported.summary.total).toBe(1);
    expect(exported.entries).toEqual([
      {
        id: entry.id,
        createdAt: entry.createdAt,
        result: 'completed',
        rating: 'helpful',
        category: 'cleanup',
      },
    ]);
    expect(exported.entries[0]).not.toHaveProperty('note');
    const serialized = JSON.stringify(exported);
    expect(serialized).not.toContain('private note');
    expect(serialized).not.toContain('workbook contents');
    expect(serialized).not.toContain('prompt contents');
    expect(serialized).not.toContain('secret.xlsx');
    expect(serialized).not.toContain('secret-token');
  });
});
