import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import {
  HistoryStack,
  applyPatch,
  cloneWorkbook,
  createCell,
  invertPatch,
  patchForChange,
  snapshotPatch,
  type Patch,
  type Workbook,
} from '../src/index.js';

function workbook(values: number[]): Workbook {
  return {
    sheets: [
      {
        name: 'Data',
        rows: [[createCell('Value')], ...values.map((value) => [createCell(value)])],
      },
    ],
  };
}

function setValuePatch(before: Workbook, row: number, value: number) {
  const oldCell = before.sheets[0]!.rows[row - 1]![0]!;
  const newCell = createCell(value);
  return patchForChange({ sheet: 'Data', row, column: 'A' }, oldCell, newCell);
}

/** A structural patch: it retains two whole workbooks, so it costs real memory. */
function setValueSnapshot(before: Workbook, row: number, value: number): Patch {
  const after = cloneWorkbook(before);
  after.sheets[0]!.rows[row - 1]![0] = createCell(value);
  return [snapshotPatch(before, after)];
}

describe('HistoryStack', () => {
  it('supports undo, redo, step-back, and periodic snapshots', () => {
    const initial = workbook([1, 2, 3]);
    const history = new HistoryStack(initial, { snapshotEvery: 2 });
    let current = initial;
    for (let index = 0; index < 5; index += 1) {
      const next = applyPatch(current, [setValuePatch(current, 2, index + 10)]);
      const patch = [setValuePatch(current, 2, index + 10)];
      history.commit('set_cells', next, patch, invertPatch(patch));
      current = next;
    }

    expect(history.position).toBe(5);
    expect(history.length).toBe(5);
    expect(history.snapshotCount).toBe(3);
    expect(history.stepBack(2).sheets[0]?.rows[1]?.[0]?.value).toBe(11);
    expect(history.redo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(12);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(11);
    expect(history.stepBack(0).sheets[0]?.rows[1]?.[0]?.value).toBe(1);
    expect(history.redo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(10);
  });

  it('drops the redo branch after a new commit', () => {
    const initial = workbook([1]);
    const history = new HistoryStack(initial);
    const first = [setValuePatch(initial, 2, 2)];
    const afterFirst = applyPatch(initial, first);
    history.commit('set_cells', afterFirst, first, invertPatch(first));
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(1);

    const second = [setValuePatch(initial, 2, 3)];
    history.commit('set_cells', applyPatch(initial, second), second, invertPatch(second));

    expect(history.length).toBe(1);
    expect(history.canRedo).toBe(false);
    expect(history.redo()).toBeUndefined();
  });
});

describe('history retention budget', () => {
  it('keeps undoing recent operations after compacting the oldest ones', () => {
    const initial = workbook([0]);
    const history = new HistoryStack(initial, { snapshotEvery: 2, maxEntries: 3 });
    let current = initial;
    for (let value = 1; value <= 10; value += 1) {
      const patch = [setValuePatch(current, 2, value)];
      current = applyPatch(current, patch);
      history.commit('set_cells', current, patch);
    }

    expect(history.length).toBeLessThanOrEqual(3);
    expect(history.compactedCount).toBeGreaterThan(0);
    expect(history.position).toBe(history.length);
    expect(history.stepBack(history.length).sheets[0]?.rows[1]?.[0]?.value).toBe(10);

    // The most recent operations are still individually undoable.
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(9);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(8);

    // Undo crosses the compaction boundary and lands on the compacted snapshot.
    const boundary = history.undo()?.sheets[0]?.rows[1]?.[0]?.value;
    expect(boundary).toBe(8);
    expect(history.canUndo).toBe(false);
    expect(history.undo()).toBeUndefined();

    // And the session still replays forward.
    expect(history.redo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(8);
    expect(history.redo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(9);
  });

  it('stays consistent and never throws on an explicit tiny budget', () => {
    const initial = workbook([0]);
    const history = new HistoryStack(initial, {
      snapshotEvery: 1,
      maxEntries: 1,
      maxRetainedCells: 1,
    });
    let current = initial;
    for (let value = 1; value <= 6; value += 1) {
      const patch = setValueSnapshot(current, 2, value);
      current = applyPatch(current, patch);
      history.commit('add_column', current, patch);
    }

    expect(history.stepBack(history.length)).toEqual(current);
    expect(history.compactedCount).toBe(5);
    // A one-entry budget folds the whole session into a single snapshot, which is
    // the newest state; the session is still consistent and still replays forward.
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(6);
    expect(history.canUndo).toBe(false);
    expect(history.redo()).toEqual(current);
    expect(history.history).toHaveLength(history.length);
    expect(history.history[0]).not.toHaveProperty('inverse');
  });

  it('compacts structural patches that blow the cell budget', () => {
    const initial = workbook([1, 2, 3, 4, 5]);
    const history = new HistoryStack(initial, { snapshotEvery: 100, maxRetainedCells: 25 });
    let current = initial;
    for (let value = 6; value <= 10; value += 1) {
      const patch = setValueSnapshot(current, 2, value);
      current = applyPatch(current, patch);
      history.commit('add_column', current, patch);
    }

    // Two 12-cell snapshots fit; the third pushes the budget over.
    expect(history.length).toBeLessThanOrEqual(2);
    expect(history.compactedCount).toBeGreaterThan(0);
    expect(history.stepBack(history.length)).toEqual(current);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(9);
  });

  it('bounds a long session with the default budgets', () => {
    const initial = workbook([0]);
    const history = new HistoryStack(initial);
    let current = initial;
    for (let value = 1; value <= 400; value += 1) {
      const patch = [setValuePatch(current, 2, value)];
      current = applyPatch(current, patch);
      history.commit('set_cells', current, patch);
    }

    expect(history.length).toBeLessThanOrEqual(200);
    expect(history.length).toBeGreaterThan(0);
    expect(history.compactedCount).toBe(400 - history.length);
    expect(history.stepBack(history.length)).toEqual(current);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(399);
  });

  it('rejects nonsensical budgets instead of silently ignoring them', () => {
    const initial = workbook([1]);
    expect(() => new HistoryStack(initial, { maxEntries: 0 })).toThrow();
    expect(() => new HistoryStack(initial, { maxRetainedCells: 0 })).toThrow();
    expect(() => new HistoryStack(initial, { snapshotEvery: 0 })).toThrow();
  });

  it('materializes every retained position to a real operation state', () => {
    // Compaction may shorten the window, but it must never reorder it, invent a
    // state, or drop the newest one.
    fc.assert(
      fc.property(
        fc.uniqueArray(fc.integer({ min: 1, max: 100_000 }), {
          minLength: 1,
          maxLength: 25,
        }),
        fc.integer({ min: 1, max: 6 }),
        fc.integer({ min: 1, max: 5 }),
        (values, maxEntries, snapshotEvery) => {
          const initial = workbook([0]);
          const history = new HistoryStack(initial, {
            maxEntries,
            snapshotEvery,
            maxRetainedCells: 1000,
          });
          let current = initial;
          const written = [0];
          for (const value of values) {
            const patch = [setValuePatch(current, 2, value)];
            current = applyPatch(current, patch);
            history.commit('set_cells', current, patch);
            written.push(value);
          }

          const operationOfPosition: number[] = [];
          for (let position = 0; position <= history.length; position += 1) {
            const value = history.stepBack(position).sheets[0]?.rows[1]?.[0]?.value as number;
            expect(written).toContain(value);
            operationOfPosition.push(written.indexOf(value));
          }

          expect(operationOfPosition.at(-1)).toBe(written.length - 1);
          for (let index = 1; index < operationOfPosition.length; index += 1) {
            expect(operationOfPosition[index]!).toBeGreaterThanOrEqual(
              operationOfPosition[index - 1]!,
            );
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  it('re-bases a new branch on the compacted snapshot after a full undo', () => {
    // Regression: once the oldest entries are compacted the origin snapshot is
    // gone, so a fresh commit has to replay from the compacted state instead of
    // from a workbook that no longer exists.
    const initial = workbook([0]);
    const history = new HistoryStack(initial, { snapshotEvery: 2, maxEntries: 2 });
    let current = initial;
    for (let value = 1; value <= 6; value += 1) {
      const patch = [setValuePatch(current, 2, value)];
      current = applyPatch(current, patch);
      history.commit('set_cells', current, patch);
    }

    while (history.canUndo) {
      history.undo();
    }
    const base = history.stepBack(0);
    const patch = [setValuePatch(base, 2, 99)];
    const after = applyPatch(base, patch);
    history.commit('set_cells', after, patch);

    expect(history.stepBack(history.length)).toEqual(after);
    expect(history.undo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(
      base.sheets[0]?.rows[1]?.[0]?.value,
    );
    expect(history.redo()?.sheets[0]?.rows[1]?.[0]?.value).toBe(99);
  });
});

describe('undo and redo property tests', () => {
  it('round-trips arbitrary sequences of cell changes', () => {
    fc.assert(
      fc.property(
        fc.array(fc.integer({ min: -1000, max: 1000 }), { minLength: 1, maxLength: 30 }),
        (values) => {
          const initial = workbook([0]);
          const history = new HistoryStack(initial, { snapshotEvery: 3 });
          let current = initial;
          for (const value of values) {
            const patch = [setValuePatch(current, 2, value)];
            current = applyPatch(current, patch);
            history.commit('set_cells', current, patch, invertPatch(patch));
          }

          const final = history.stepBack(values.length);
          expect(final).toEqual(current);
          const origin = history.stepBack(0);
          expect(origin).toEqual(initial);
          const replayed = history.stepBack(values.length);
          expect(replayed).toEqual(current);
        },
      ),
      { numRuns: 100 },
    );
  });
});
