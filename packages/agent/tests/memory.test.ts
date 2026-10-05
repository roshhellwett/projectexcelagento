import { describe, expect, it } from 'vitest';

import { createMemoryStore, normalizeQuery, querySimilarity } from '../src/index.js';

describe('self-learning memory', () => {
  it('normalizes queries by removing punctuation, casing, and stop words', () => {
    expect(normalizeQuery('Please, Remove the DUP-licate rows!')).toBe('remove dup licate');
    expect(normalizeQuery('   ')).toBe('');
  });

  it('scores identical queries highest and disjoint queries zero', () => {
    expect(querySimilarity('sort by amount', 'sort by amount')).toBe(1);
    expect(querySimilarity('sort by amount', 'delete empty column')).toBe(0);
  });

  it('deduplicates identical associations instead of appending', () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'trim names',
      rawQuery: 'trim names',
      operation: 'normalize_text',
      args: {
        sheet: 'Orders',
        columns: ['B'],
        trim: true,
        collapseWhitespace: true,
        case: 'none',
        headerRow: 1,
      },
      sheetName: 'Orders',
    });
    memory.remember({
      key: 'trim names',
      rawQuery: 'trim names',
      operation: 'normalize_text',
      args: {
        sheet: 'Orders',
        columns: ['B'],
        trim: true,
        collapseWhitespace: true,
        case: 'none',
        headerRow: 1,
      },
      sheetName: 'Orders',
    });

    expect(memory.entries()).toHaveLength(1);
  });

  it('does not replay a record that has never succeeded', () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'sort by amount',
      rawQuery: 'sort by amount',
      operation: 'sort_range',
      args: { sheet: 'Orders', column: 'C', direction: 'asc', startRow: 2, startColumn: 'A' },
      sheetName: 'Orders',
    });

    expect(memory.retrieve('sort by amount', 'Orders')).toBeUndefined();
    expect(memory.retrieve('sort by amount', 'Orders', 0)).toBeUndefined();
  });

  it('replays a record once it has accumulated successes', () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'sort by amount',
      rawQuery: 'sort by amount',
      operation: 'sort_range',
      args: { sheet: 'Orders', column: 'C', direction: 'asc', startRow: 2, startColumn: 'A' },
      sheetName: 'Orders',
    });
    memory.recordOutcome('sort_range', 'Orders', true);
    memory.recordOutcome('sort_range', 'Orders', true);

    const hit = memory.retrieve('sort by amount', 'Orders');
    expect(hit?.operation).toBe('sort_range');
    expect(hit?.successes).toBe(2);
  });

  it('round-trips through JSON and clears', () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'dedupe rows',
      rawQuery: 'dedupe rows',
      operation: 'delete_duplicates',
      args: { sheet: 'Orders', columns: ['A'], headerRow: 1, keep: 'first' },
      sheetName: 'Orders',
    });
    memory.recordOutcome('delete_duplicates', 'Orders', true);
    memory.recordOutcome('delete_duplicates', 'Orders', true);

    const restored = createMemoryStore(memory.toJSON());
    expect(restored.entries()).toHaveLength(1);
    expect(restored.retrieve('dedupe rows', 'Orders')?.operation).toBe('delete_duplicates');

    restored.clear();
    expect(restored.entries()).toHaveLength(0);
  });

  it('matches broken English and typos with high similarity', () => {
    // "remov dupli" vs "remove duplicate"
    expect(querySimilarity('remov dupli', 'remove duplicate')).toBeGreaterThanOrEqual(0.6);
    // "sorrt amunt decs" vs "sort amount descending"
    expect(querySimilarity('sorrt amunt decs', 'sort amount descending')).toBeGreaterThanOrEqual(
      0.6,
    );
    // "mak date formt" vs "format dates"
    expect(querySimilarity('mak date formt', 'make date format')).toBeGreaterThanOrEqual(0.6);
  });
});
