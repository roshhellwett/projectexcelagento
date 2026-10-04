import { type Cell, type CellValue, type Sheet, indexToColumn } from '@excel-agent/engine';

/**
 * The spreadsheet interaction model, expressed as pure functions.
 *
 * Everything here is deliberately free of React so the rules that decide what a keystroke, a
 * paste, or a fill-handle drag *mean* can be tested directly, instead of only through the DOM.
 */

/** A cell address in grid terms: `row` is 1-based (as in the formula bar), `colIdx` is 0-based. */
export interface CellPosition {
  row: number;
  colIdx: number;
}

/** A rectangular selection. A single cell is a rectangle with all four bounds equal. */
export interface CellRect {
  startRow: number;
  endRow: number;
  startColIdx: number;
  endColIdx: number;
}

/** One cell write, in the shape the engine's `edit_cells` operation accepts. */
export interface CellEdit {
  row: number;
  column: string;
  value?: CellValue;
  formula?: string;
}

export const DEFAULT_COLUMN_WIDTH = 120;
export const MIN_COLUMN_WIDTH = 56;
export const MAX_COLUMN_WIDTH = 640;

/** Excel's own addressable limits, so a fill or paste cannot build a sheet nobody can open. */
export const MAX_ROW_NUMBER = 1_048_576;
export const MAX_COLUMN_INDEX = 16_383;

/**
 * Ceiling on the cells one gesture may write. A drag across the whole sheet would otherwise build
 * a hundred million edit objects and lock the tab; the gesture is clamped to this instead.
 */
export const MAX_BULK_EDITS = 20_000;

export function rectContains(rect: CellRect, row: number, colIdx: number): boolean {
  return (
    row >= rect.startRow &&
    row <= rect.endRow &&
    colIdx >= rect.startColIdx &&
    colIdx <= rect.endColIdx
  );
}

/** The rectangle spanned by an anchor and the cell the pointer or caret reached. */
export function rectFromPositions(anchor: CellPosition, focus: CellPosition): CellRect {
  return {
    startRow: Math.min(anchor.row, focus.row),
    endRow: Math.max(anchor.row, focus.row),
    startColIdx: Math.min(anchor.colIdx, focus.colIdx),
    endColIdx: Math.max(anchor.colIdx, focus.colIdx),
  };
}

export function unionRects(left: CellRect, right: CellRect): CellRect {
  return {
    startRow: Math.min(left.startRow, right.startRow),
    endRow: Math.max(left.endRow, right.endRow),
    startColIdx: Math.min(left.startColIdx, right.startColIdx),
    endColIdx: Math.max(left.endColIdx, right.endColIdx),
  };
}

export function clampPosition(
  position: CellPosition,
  totalRows: number,
  totalCols: number,
): CellPosition {
  return {
    row: Math.min(Math.max(1, position.row), Math.max(1, totalRows)),
    colIdx: Math.min(Math.max(0, position.colIdx), Math.max(0, totalCols - 1)),
  };
}

/** A plain decimal number, the only shape typed into a cell that becomes a number. */
const NUMERIC_TEXT = /^-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * Turns what the user typed into a cell write.
 *
 * A leading `=` is a formula and keeps its sign, because that is how this engine stores formulas
 * everywhere else. `true`/`false` become booleans, anything numeric becomes a number so the sheet
 * can sort and total it, and everything else - including `0123`-style codes and `1-2` ranges - stays
 * the exact text that was typed rather than being silently reinterpreted.
 */
export function coerceTypedValue(raw: string): { value?: CellValue; formula?: string } {
  if (raw.length === 0) return { value: null };
  if (raw.startsWith('=')) return { formula: raw };

  const trimmed = raw.trim();
  if (NUMERIC_TEXT.test(trimmed)) {
    const numeric = Number(trimmed);
    if (Number.isFinite(numeric)) return { value: numeric };
  }
  if (trimmed.toUpperCase() === 'TRUE') return { value: true };
  if (trimmed.toUpperCase() === 'FALSE') return { value: false };
  return { value: raw };
}

/**
 * Splits clipboard text into a grid of rows and columns.
 *
 * Tab-separated values and newlines are the interchange every spreadsheet and terminal accepts, so
 * they are the only thing parsed here: one trailing newline is dropped so a copy that ends with a
 * break does not add a phantom empty row.
 */
export function parseClipboardGrid(text: string): string[][] {
  const normalized = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (normalized.length === 0) return [];
  const lines = normalized.split('\n').map((line) => line.split('\t'));
  // Text copied out of another program is often ragged: "a\tb\nc". A spreadsheet pastes that as a
  // rectangle, so the short row is padded rather than silently pasting one column wide.
  const width = lines.reduce((widest, line) => Math.max(widest, line.length), 0);
  return lines.map((line) => {
    if (line.length === width) return line;
    return [...line, ...new Array<string>(width - line.length).fill('')];
  });
}

/** Serializes a rectangle to the tab-separated text a spreadsheet can read back. */
export function clipboardTextForRect(
  sheet: Sheet,
  rect: CellRect,
  display: (cell: Cell | undefined, row: number, colIdx: number) => string,
): string {
  const lines: string[] = [];
  for (let row = rect.startRow; row <= rect.endRow; row += 1) {
    const cells = sheet.rows[row - 1] ?? [];
    const columns: string[] = [];
    for (let colIdx = rect.startColIdx; colIdx <= rect.endColIdx; colIdx += 1) {
      columns.push(display(cells[colIdx], row, colIdx));
    }
    lines.push(columns.join('\t'));
  }
  return lines.join('\n');
}

/**
 * Writes every cell of a rectangle from a block of clipboard text, anchored at the rectangle's
 * top-left corner. Cells are capped so an enormous paste fails small instead of freezing the tab.
 */
export function editsForClipboardBlock(
  block: string[][],
  anchorRow: number,
  anchorColIdx: number,
): CellEdit[] {
  const edits: CellEdit[] = [];
  for (let rowOffset = 0; rowOffset < block.length; rowOffset += 1) {
    const row = anchorRow + rowOffset;
    if (row > MAX_ROW_NUMBER) break;
    const line = block[rowOffset] ?? [];
    for (let colOffset = 0; colOffset < line.length; colOffset += 1) {
      const colIdx = anchorColIdx + colOffset;
      if (colIdx > MAX_COLUMN_INDEX) break;
      edits.push({
        row,
        column: indexToColumn(colIdx),
        ...coerceTypedValue(line[colOffset] ?? ''),
      });
      if (edits.length >= MAX_BULK_EDITS) return edits;
    }
  }
  return edits;
}

/**
 * Whether a cell needs clearing at all.
 *
 * A cell holding only spaces is not blank: `ISBLANK` is false for it, and pressing Delete on it
 * does remove those spaces. Only a genuinely empty cell is left alone.
 */
function isBlankCell(cell: Cell | undefined): boolean {
  if (!cell) return true;
  return cell.value === null || cell.value === undefined || cell.value === '';
}

/**
 * Clears a rectangle. Cells that are already empty are left out, so pressing Delete on empty space
 * changes nothing at all instead of growing the sheet with blank cells that then show up as diffs.
 */
export function clearEditsForRect(sheet: Sheet, rect: CellRect): CellEdit[] {
  const edits: CellEdit[] = [];
  for (let row = rect.startRow; row <= rect.endRow; row += 1) {
    const cells = sheet.rows[row - 1];
    for (let colIdx = rect.startColIdx; colIdx <= rect.endColIdx; colIdx += 1) {
      if (isBlankCell(cells?.[colIdx])) continue;
      edits.push({ row, column: indexToColumn(colIdx), value: null });
      if (edits.length >= MAX_BULK_EDITS) return edits;
    }
  }
  return edits;
}

/** One step of a fill series: numbers count, dates advance a day, everything else repeats. */
function seriesValueFor(seed: Cell, step: number): { value?: CellValue; formula?: string } {
  if (seed.formula !== undefined) return { value: seed.formula };
  const value = seed.value;
  if (typeof value === 'number') return { value: value + step };
  if (value instanceof Date) return { value: new Date(value.getTime() + step * 86_400_000) };
  return { value };
}

/**
 * The edits a fill-handle drag from `source` to `target` produces, excluding the source itself.
 *
 * A one-cell source becomes a series along whichever axis the drag grew - numbers step, dates
 * advance a day, text repeats. Dragging upwards or leftwards subtracts instead of failing, because
 * a user extending a list backwards wants a decreasing sequence, not an error. A larger source
 * repeats its own pattern, which is what dragging a two-column block down a list means.
 */
export function fillEditsFor(sheet: Sheet, source: CellRect, target: CellRect): CellEdit[] {
  const sourceHeight = source.endRow - source.startRow + 1;
  const sourceWidth = source.endColIdx - source.startColIdx + 1;
  const single = sourceHeight === 1 && sourceWidth === 1;
  const seed = single ? sheet.rows[source.startRow - 1]?.[source.startColIdx] : undefined;
  // A drag that grew the range vertically steps by row - upwards or downwards; one that only moved
  // sideways steps by column.
  const stepsByRow = target.startRow !== source.startRow || target.endRow !== source.endRow;

  const edits: CellEdit[] = [];
  for (let row = target.startRow; row <= target.endRow; row += 1) {
    for (let colIdx = target.startColIdx; colIdx <= target.endColIdx; colIdx += 1) {
      if (rectContains(source, row, colIdx)) continue;
      if (row > MAX_ROW_NUMBER || colIdx > MAX_COLUMN_INDEX) continue;

      let write: { value?: CellValue; formula?: string };
      if (single) {
        const step = stepsByRow ? row - source.startRow : colIdx - source.startColIdx;
        write = seed ? seriesValueFor(seed, step) : { value: null };
      } else {
        const sourceCell =
          sheet.rows[source.startRow - 1 + ((row - target.startRow) % sourceHeight)]?.[
            source.startColIdx + ((colIdx - target.startColIdx) % sourceWidth)
          ];
        // A formula is copied as text: rewriting its references per cell is guesswork, and a
        // repeated formula that still points at the source row is at least visible and correctable.
        write = sourceCell
          ? sourceCell.formula !== undefined
            ? { value: sourceCell.formula }
            : { value: sourceCell.value }
          : { value: null };
      }

      edits.push({ row, column: indexToColumn(colIdx), ...write });
      if (edits.length >= MAX_BULK_EDITS) return edits;
    }
  }
  return edits;
}
