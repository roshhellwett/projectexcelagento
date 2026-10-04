import { describe, expect, it } from 'vitest';
import { createCell, type Sheet } from '@excel-agent/engine';

import {
  MAX_BULK_EDITS,
  clearEditsForRect,
  clipboardTextForRect,
  coerceTypedValue,
  editsForClipboardBlock,
  fillEditsFor,
  parseClipboardGrid,
  rectFromPositions,
  unionRects,
  type CellRect,
} from './grid-edit.js';

function sheet(...rows: Array<Array<string | number | null>>): Sheet {
  return { name: 'Data', rows: rows.map((row) => row.map((value) => createCell(value))) };
}

const A1: CellRect = { startRow: 1, endRow: 1, startColIdx: 0, endColIdx: 0 };

describe('grid selection maths', () => {
  it('builds a rectangle from an anchor and a focus cell in any direction', () => {
    expect(rectFromPositions({ row: 4, colIdx: 2 }, { row: 2, colIdx: 5 })).toEqual({
      startRow: 2,
      endRow: 4,
      startColIdx: 2,
      endColIdx: 5,
    });
    expect(rectFromPositions({ row: 2, colIdx: 5 }, { row: 4, colIdx: 2 })).toEqual({
      startRow: 2,
      endRow: 4,
      startColIdx: 2,
      endColIdx: 5,
    });
  });

  it('unions a fill drag target with the range it started from', () => {
    const draggedUp = unionRects(
      { startRow: 3, endRow: 5, startColIdx: 0, endColIdx: 1 },
      { startRow: 1, endRow: 1, startColIdx: 2, endColIdx: 2 },
    );
    expect(draggedUp).toEqual({ startRow: 1, endRow: 5, startColIdx: 0, endColIdx: 2 });
  });
});

describe('coerceTypedValue', () => {
  it('keeps text that only looks numeric from becoming a number', () => {
    expect(coerceTypedValue('ORD-1001')).toEqual({ value: 'ORD-1001' });
    expect(coerceTypedValue('1-2')).toEqual({ value: '1-2' });
    expect(coerceTypedValue('1,234')).toEqual({ value: '1,234' });
    expect(coerceTypedValue('2026-03-01')).toEqual({ value: '2026-03-01' });
  });

  it('turns numbers, booleans, blanks and formulas into real cell values', () => {
    expect(coerceTypedValue('42')).toEqual({ value: 42 });
    expect(coerceTypedValue('-3.5')).toEqual({ value: -3.5 });
    expect(coerceTypedValue('1e3')).toEqual({ value: 1000 });
    expect(coerceTypedValue('true')).toEqual({ value: true });
    expect(coerceTypedValue('')).toEqual({ value: null });
    // The engine stores formulas with their leading `=`, so typing one must not lose it.
    expect(coerceTypedValue('=SUM(A1:A9)')).toEqual({ formula: '=SUM(A1:A9)' });
  });
});

describe('clipboard text', () => {
  it('round-trips a block of cells through tab-separated text', () => {
    const data = sheet(['a', 'b'], ['c', 'd']);
    const text = clipboardTextForRect(
      data,
      { startRow: 1, endRow: 2, startColIdx: 0, endColIdx: 1 },
      (cell) => (cell ? String(cell.value ?? '') : ''),
    );
    expect(text).toBe('a\tb\nc\td');
    expect(parseClipboardGrid(text)).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('normalises Windows line endings and drops one trailing newline', () => {
    expect(parseClipboardGrid('1\t2\r\n3\t4\r\n')).toEqual([
      ['1', '2'],
      ['3', '4'],
    ]);
  });

  it('anchors a pasted block at the selection, padding a ragged block to a rectangle', () => {
    // "x\ty\nz" is what text copied out of a terminal looks like: the second row is one column short.
    const edits = editsForClipboardBlock(parseClipboardGrid('x\ty\nz'), 2, 1);
    expect(edits).toEqual([
      { row: 2, column: 'B', value: 'x' },
      { row: 2, column: 'C', value: 'y' },
      { row: 3, column: 'B', value: 'z' },
      { row: 3, column: 'C', value: null },
    ]);
  });

  it('caps a paste that would try to write more cells than the engine accepts', () => {
    const huge = Array.from({ length: MAX_BULK_EDITS + 50 }, () => ['v']);
    expect(editsForClipboardBlock(huge, 1, 0)).toHaveLength(MAX_BULK_EDITS);
  });
});

describe('clearEditsForRect', () => {
  it('clears only the cells that actually hold something', () => {
    const data = sheet(['a', null, '  '], [null, null, null]);
    const edits = clearEditsForRect(data, { startRow: 1, endRow: 2, startColIdx: 0, endColIdx: 2 });
    expect(edits).toEqual([
      { row: 1, column: 'A', value: null },
      { row: 1, column: 'C', value: null },
    ]);
  });

  it('asks for nothing when the range is already empty', () => {
    const data = sheet([null, null]);
    expect(clearEditsForRect(data, A1)).toEqual([]);
  });
});

describe('fillEditsFor', () => {
  it('counts a numeric series downwards', () => {
    const data = sheet(['x'], [7]);
    const edits = fillEditsFor(
      data,
      { startRow: 2, endRow: 2, startColIdx: 0, endColIdx: 0 },
      {
        startRow: 2,
        endRow: 4,
        startColIdx: 0,
        endColIdx: 0,
      },
    );
    expect(edits).toEqual([
      { row: 3, column: 'A', value: 8 },
      { row: 4, column: 'A', value: 9 },
    ]);
  });

  it('counts backwards when the drag goes up', () => {
    const data = sheet(['x'], [7], [null]);
    const edits = fillEditsFor(
      data,
      { startRow: 2, endRow: 2, startColIdx: 0, endColIdx: 0 },
      {
        startRow: 1,
        endRow: 2,
        startColIdx: 0,
        endColIdx: 0,
      },
    );
    expect(edits).toEqual([{ row: 1, column: 'A', value: 6 }]);
  });

  it('steps a date series by a day', () => {
    const data = {
      name: 'Data',
      rows: [[createCell(new Date(Date.UTC(2026, 2, 1)))]],
    };
    const edits = fillEditsFor(data, A1, { startRow: 1, endRow: 3, startColIdx: 0, endColIdx: 0 });
    expect(edits).toHaveLength(2);
    expect((edits[0] as { value: Date }).value.toISOString().slice(0, 10)).toBe('2026-03-02');
    expect((edits[1] as { value: Date }).value.toISOString().slice(0, 10)).toBe('2026-03-03');
  });

  it('repeats text unchanged', () => {
    const data = sheet(['done']);
    const edits = fillEditsFor(data, A1, { startRow: 1, endRow: 3, startColIdx: 0, endColIdx: 0 });
    expect(edits).toEqual([
      { row: 2, column: 'A', value: 'done' },
      { row: 3, column: 'A', value: 'done' },
    ]);
  });

  it('repeats a multi-cell source as a pattern', () => {
    const data = sheet(['a', 'b'], [null, null], [null, null]);
    const edits = fillEditsFor(
      data,
      { startRow: 1, endRow: 1, startColIdx: 0, endColIdx: 1 },
      {
        startRow: 1,
        endRow: 3,
        startColIdx: 0,
        endColIdx: 1,
      },
    );
    expect(edits).toEqual([
      { row: 2, column: 'A', value: 'a' },
      { row: 2, column: 'B', value: 'b' },
      { row: 3, column: 'A', value: 'a' },
      { row: 3, column: 'B', value: 'b' },
    ]);
  });

  it('clamps a drag that would run past the last addressable row', () => {
    const data = sheet([1]);
    const edits = fillEditsFor(data, A1, {
      startRow: 1,
      endRow: 1_048_600,
      startColIdx: 0,
      endColIdx: 0,
    });
    expect(edits.length).toBeLessThanOrEqual(MAX_BULK_EDITS);
  });
});
