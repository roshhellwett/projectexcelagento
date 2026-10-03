import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import {
  HistoryStack,
  applyPatch,
  createCell,
  invertPatch,
  patchForChange,
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
