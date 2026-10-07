import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  type Workbook,
  type Sheet,
  type DateSystem,
  indexToColumn,
  maxColumnCount,
} from '@excel-agent/engine';

import { describeCellSelection, type CellSelection } from '../lib/selection-context.js';
import { cellDisplay, formulaTextFor, useCellEvaluator } from '../lib/cell-evaluator.js';
import {
  DEFAULT_COLUMN_WIDTH,
  MAX_BULK_EDITS,
  MAX_COLUMN_INDEX,
  MAX_COLUMN_WIDTH,
  MAX_ROW_NUMBER,
  MIN_COLUMN_WIDTH,
  clearEditsForRect,
  clipboardTextForRect,
  clampPosition,
  coerceTypedValue,
  editsForClipboardBlock,
  fillEditsFor,
  parseClipboardGrid,
  rectFromPositions,
  unionRects,
  type CellEdit,
  type CellPosition,
  type CellRect,
} from '../lib/grid-edit.js';
import { GridCell, type CellActions, type EditorMove } from './GridCell.js';

interface SpreadsheetGridProps {
  workbook: Workbook;
  activeSheetName: string;
  onSelectSheet: (sheetName: string) => void;
  recentChangedCells: Set<string>; // formatted as `${sheetName}:${row}:${colLetter}`
  searchHighlightCells?: Set<string>;
  onQuickSort?: (columnLetter: string, direction: 'asc' | 'desc') => void;
  onSelectCell?: (coord: { row: number; column: string; value: unknown }) => void;
  onFileDrop?: (file: File) => void;
  onAddSelectionContext?: (ctx: CellSelection) => void;
  onSelectionChange?: (selection: CellRect) => void;
  /**
   * Applies a batch of cell writes to the named sheet. Without it the grid stays a viewer: cells
   * are `aria-readonly`, editing and the fill handle are refused, and nothing is lost by trying.
   */
  onEditCells?: (sheet: string, edits: CellEdit[]) => void;
  /**
   * Which epoch the workbook's bare serials count from. Legacy Mac Excel workbooks use 1904, and
   * reading those as 1900 shifts every displayed date by four years and a day.
   */
  dateSystem?: DateSystem;
}

/**
 * Row height the virtualiser plans with. It is also applied inline to every rendered row, and
 * re-measured from the DOM below, so the plan and the pixels cannot drift apart - a virtualiser
 * that assumes 28px while rows render 31px tall makes the scrollbar lie about what is on screen.
 */
const DEFAULT_ROW_HEIGHT = 30;
const ROW_OVERSCAN = 12;
const COLUMN_OVERSCAN = 3;
/** Below this width a sheet renders whole; above it, columns are windowed too. */
const COLUMN_VIRTUALIZE_THRESHOLD = 80;
/** Keyboard column nudge, so a resizer is reachable without a pointer. */
const RESIZE_KEYBOARD_STEP = 8;
/** How many cells one gesture may ask the sheet to change before the sheet says no. */
const MAX_PASTE_BLOCK = 5_000;

interface ClipboardBuffer {
  text: string;
  mode: 'copy' | 'cut';
  sheetName: string;
  rect: CellRect;
}

function isEditableTarget(target: EventTarget | null): boolean {
  const element = target as HTMLElement | null;
  if (!element || typeof element.tagName !== 'string') return false;
  return (
    element.tagName === 'INPUT' ||
    element.tagName === 'TEXTAREA' ||
    element.tagName === 'SELECT' ||
    element.isContentEditable
  );
}

/** True while a dialog or drawer is on top of the grid, which then surrenders the keyboard. */
function isDialogOpen(): boolean {
  return typeof document !== 'undefined' && document.querySelector('[data-dialog-open]') !== null;
}

/** Largest index in the cumulative-width table whose offset is still at or before `offset`. */
function indexForOffset(offsets: number[], offset: number): number {
  let low = 0;
  let high = offsets.length - 1;
  while (low < high) {
    const mid = (low + high + 1) >> 1;
    if ((offsets[mid] as number) <= offset) low = mid;
    else high = mid - 1;
  }
  return low;
}

/** The text an edit begins from: the formula when there is one, otherwise the raw value. */
function editableTextFor(cell: { value: unknown; formula?: string } | undefined): string {
  if (!cell) return '';
  if (cell.formula !== undefined) return cell.formula;
  if (cell.value === null || cell.value === undefined) return '';
  if (cell.value instanceof Date) return cell.value.toISOString().slice(0, 10);
  return String(cell.value);
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  workbook,
  activeSheetName,
  onSelectSheet,
  recentChangedCells,
  searchHighlightCells = new Set(),
  onQuickSort,
  onSelectCell,
  onFileDrop,
  onAddSelectionContext,
  onSelectionChange,
  onEditCells,
  dateSystem = '1900',
}) => {
  const currentSheet: Sheet = workbook.sheets.find((s) => s.name === activeSheetName) ||
    workbook.sheets[0] || { name: 'Sheet1', rows: [] };

  const evaluateCell = useCellEvaluator(workbook, dateSystem);
  const editable = typeof onEditCells === 'function';

  const [anchor, setAnchor] = useState<CellPosition>({ row: 1, colIdx: 0 });
  const [activeCell, setActiveCell] = useState<CellPosition>({ row: 1, colIdx: 0 });
  const [editing, setEditing] = useState<{ row: number; colIdx: number; text: string } | null>(
    null,
  );
  const [columnWidths, setColumnWidths] = useState<Record<number, number>>({});
  const [resizeDrag, setResizeDrag] = useState<{
    colIdx: number;
    startX: number;
    startWidth: number;
  } | null>(null);
  const [fillPreview, setFillPreview] = useState<CellRect | null>(null);

  const [isDragOver, setIsDragOver] = useState(false);
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    rowIdx: number;
    colIdx: number;
  } | null>(null);

  const scrollWrapperRef = useRef<HTMLDivElement>(null);
  const probeRowRef = useRef<HTMLTableRowElement>(null);
  const activeCellRef = useRef<HTMLTableCellElement>(null);
  const gridFocusedRef = useRef(false);
  const clipboardRef = useRef<ClipboardBuffer | null>(null);
  const fillStartRef = useRef<CellRect | null>(null);
  const fillPreviewRef = useRef<CellRect | null>(null);

  const selection = useMemo(() => rectFromPositions(anchor, activeCell), [anchor, activeCell]);
  fillPreviewRef.current = fillPreview;

  useEffect(() => {
    onSelectionChange?.(selection);
  }, [onSelectionChange, selection]);

  const totalRows = currentSheet.rows.length;
  const totalCols = maxColumnCount(currentSheet.rows);

  // The cell-level actions read the sheet, the selection and the writer through this one ref so
  // their identities never change. A per-cell inline arrow would be a new prop on every render and
  // `React.memo` would never skip a single cell.
  const latest = useRef({
    sheet: currentSheet,
    editing,
    selection,
    totalRows,
    totalCols,
    editable,
  });
  latest.current = { sheet: currentSheet, editing, selection, totalRows, totalCols, editable };
  // The writer is a prop that the shell recreates whenever the workbook changes, so it is read the
  // same way as the state above.
  const onEditCellsRef = useRef(onEditCells);
  onEditCellsRef.current = onEditCells;
  const workbookRef = useRef(workbook);
  workbookRef.current = workbook;

  // Column widths belong to a sheet, the way they do in a spreadsheet, so switching sheets starts
  // again from the default width.
  useEffect(() => {
    setColumnWidths({});
  }, [activeSheetName]);

  const widthForColumn = useCallback(
    (colIdx: number) => columnWidths[colIdx] ?? DEFAULT_COLUMN_WIDTH,
    [columnWidths],
  );

  // Cumulative widths make the column window a binary search rather than a scan, which is what
  // keeps a 16k-column sheet from rendering 16k cells for every row.
  const columnOffsets = useMemo(() => {
    const offsets = new Array<number>(totalCols + 1);
    offsets[0] = 0;
    for (let colIdx = 0; colIdx < totalCols; colIdx += 1) {
      offsets[colIdx + 1] = (offsets[colIdx] as number) + widthForColumn(colIdx);
    }
    return offsets;
  }, [totalCols, widthForColumn]);

  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportWidth, setViewportWidth] = useState(0);
  const [measuredRowHeight, setMeasuredRowHeight] = useState(0);
  const rowHeight = measuredRowHeight > 0 ? measuredRowHeight : DEFAULT_ROW_HEIGHT;

  const handleScroll = useCallback(() => {
    const el = scrollWrapperRef.current;
    if (!el) return;
    setScrollTop(el.scrollTop);
    setScrollLeft(el.scrollLeft);
    // jsdom reports 0, and a 0 viewport would window the grid down to nothing at all.
    if (el.clientHeight > 0) setViewportHeight(el.clientHeight);
    if (el.clientWidth > 0) setViewportWidth(el.clientWidth);
  }, []);

  useEffect(() => {
    handleScroll();
    const onResize = () => handleScroll();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [handleScroll, totalRows, totalCols]);

  const rowWindow = useMemo(() => {
    if (totalRows <= 120 || viewportHeight <= 0) return { start: 0, end: totalRows };
    return {
      start: Math.max(0, Math.floor(scrollTop / rowHeight) - ROW_OVERSCAN),
      end: Math.min(totalRows, Math.ceil((scrollTop + viewportHeight) / rowHeight) + ROW_OVERSCAN),
    };
  }, [totalRows, scrollTop, viewportHeight, rowHeight]);

  const columnWindow = useMemo(() => {
    const total = columnOffsets[totalCols] ?? 0;
    if (totalCols === 0) return { start: 0, end: 0, left: 0, right: 0 };
    if (viewportWidth <= 0 || totalCols <= COLUMN_VIRTUALIZE_THRESHOLD) {
      return { start: 0, end: totalCols, left: 0, right: 0 };
    }
    const start = Math.max(
      0,
      Math.min(indexForOffset(columnOffsets, scrollLeft) - COLUMN_OVERSCAN, totalCols - 1),
    );
    const end = Math.min(
      totalCols,
      Math.max(
        start + 1,
        indexForOffset(columnOffsets, scrollLeft + viewportWidth) + 1 + COLUMN_OVERSCAN,
      ),
    );
    return {
      start,
      end,
      left: columnOffsets[start] as number,
      right: total - (columnOffsets[end] as number),
    };
  }, [columnOffsets, totalCols, scrollLeft, viewportWidth]);

  useEffect(() => {
    const probe = probeRowRef.current;
    if (!probe) return;
    const measured = probe.getBoundingClientRect().height;
    // Only re-plan when the rendered height genuinely disagrees, so a sub-pixel difference between
    // two passes cannot start a loop.
    if (measured > 0 && Math.abs(measured - rowHeight) > 0.5) setMeasuredRowHeight(measured);
  }, [rowHeight, totalRows, columnWidths]);

  // ---------------------------------------------------------------- selection

  const selectPosition = useCallback((position: CellPosition, extend: boolean) => {
    if (extend) {
      setActiveCell(position);
      return;
    }
    setAnchor(position);
    setActiveCell(position);
  }, []);

  const moveActive = useCallback(
    (rowDelta: number, colDelta: number, extend: boolean) => {
      const next = (previous: CellPosition) =>
        clampPosition(
          { row: previous.row + rowDelta, colIdx: previous.colIdx + colDelta },
          Math.max(1, totalRows),
          Math.max(1, totalCols),
        );
      setActiveCell(next);
      if (!extend) setAnchor(next);
    },
    [totalRows, totalCols],
  );

  // The caret follows the selection while the grid owns the keyboard, so a keyboard user keeps
  // typing into the cell the arrow keys just moved to.
  useEffect(() => {
    if (!gridFocusedRef.current) return;
    const cell = activeCellRef.current;
    if (cell && document.activeElement !== cell && !isEditableTarget(document.activeElement)) {
      cell.focus();
    }
  }, [activeCell, totalRows, totalCols]);

  useEffect(() => {
    if (!onSelectCell) return;
    const cell = currentSheet.rows[activeCell.row - 1]?.[activeCell.colIdx];
    if (!cell) return;
    onSelectCell({
      row: activeCell.row,
      column: indexToColumn(activeCell.colIdx),
      value: cell.value,
    });
  }, [activeCell, currentSheet, onSelectCell]);

  // ----------------------------------------------------------------- editing

  const beginEdit = useCallback((position: CellPosition, seed?: string) => {
    const { sheet, totalRows: rows, totalCols: cols, editable: canEdit } = latest.current;
    if (!canEdit || isDialogOpen()) return;
    if (rows === 0 || cols === 0) return;
    const cell = sheet.rows[position.row - 1]?.[position.colIdx];
    setEditing({
      row: position.row,
      colIdx: position.colIdx,
      text: seed ?? editableTextFor(cell),
    });
  }, []);

  const commitEdit = useCallback((raw: string, move: EditorMove) => {
    const target = latest.current.editing;
    setEditing(null);
    if (!target) return;
    const { sheet, totalRows: rows, totalCols: cols } = latest.current;
    const current = sheet.rows[target.row - 1]?.[target.colIdx];
    // Re-committing the text already there is not an edit, so it must not enter history.
    if (raw !== editableTextFor(current)) {
      onEditCellsRef.current?.(sheet.name, [
        { row: target.row, column: indexToColumn(target.colIdx), ...coerceTypedValue(raw) },
      ]);
    }
    if (move === 'none') return;
    const deltas: Record<Exclude<EditorMove, 'none'>, [number, number]> = {
      down: [1, 0],
      up: [-1, 0],
      right: [0, 1],
      left: [0, -1],
    };
    const [rowDelta, colDelta] = deltas[move];
    const next = clampPosition(
      { row: target.row + rowDelta, colIdx: target.colIdx + colDelta },
      Math.max(1, rows),
      Math.max(1, cols),
    );
    setAnchor(next);
    setActiveCell(next);
  }, []);

  const clearSelection = useCallback(() => {
    const { sheet, selection: rect, editable: canEdit } = latest.current;
    if (!canEdit || isDialogOpen()) return;
    const edits = clearEditsForRect(sheet, rect);
    if (edits.length > 0) onEditCellsRef.current?.(sheet.name, edits);
    setEditing(null);
  }, []);

  // --------------------------------------------------------------- clipboard

  const selectionText = useCallback(() => {
    return clipboardTextForRect(
      currentSheet,
      selection,
      (cell, row, colIdx) =>
        cellDisplay(cell, evaluateCell(cell, currentSheet.name, colIdx, row), dateSystem).text,
    );
  }, [currentSheet, selection, evaluateCell, dateSystem]);

  const pasteBlock = useCallback((block: string[][], anchorRow: number, anchorColIdx: number) => {
    const { sheet, editable: canEdit } = latest.current;
    if (!canEdit || isDialogOpen()) return;
    const pasted = editsForClipboardBlock(block, anchorRow, anchorColIdx).slice(0, MAX_PASTE_BLOCK);
    if (pasted.length === 0) return;

    const edits: CellEdit[] = [];
    const cut = clipboardRef.current;
    // A cut only moves cells when the source is the sheet being pasted into; across sheets the
    // safe reading is a copy, since the move would need the other sheet on screen.
    if (cut?.mode === 'cut' && cut.sheetName === sheet.name) {
      const source = workbookRef.current.sheets.find((s) => s.name === cut.sheetName) ?? sheet;
      edits.push(...clearEditsForRect(source, cut.rect));
    }
    edits.push(...pasted);
    onEditCellsRef.current?.(sheet.name, edits.slice(0, MAX_BULK_EDITS));

    // The pasted block becomes the selection, so what landed is visibly what was asked for.
    const rowSpan = Math.min(block.length, MAX_ROW_NUMBER - anchorRow + 1);
    const colSpan = Math.min(
      block.reduce((widest, line) => Math.max(widest, line.length), 0),
      MAX_COLUMN_INDEX - anchorColIdx + 1,
    );
    setAnchor({ row: anchorRow, colIdx: anchorColIdx });
    setActiveCell({
      row: anchorRow + Math.max(0, rowSpan - 1),
      colIdx: anchorColIdx + Math.max(0, colSpan - 1),
    });
  }, []);

  const pasteText = useCallback(
    (text: string, at?: CellPosition) => {
      const block = parseClipboardGrid(text);
      if (block.length === 0) return;
      pasteBlock(block, at?.row ?? selection.startRow, at?.colIdx ?? selection.startColIdx);
    },
    [selection, pasteBlock],
  );

  // ------------------------------------------------------------ fill handle

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const source = fillStartRef.current;
      const wrapper = scrollWrapperRef.current;
      if (!source || !wrapper) return;
      const bounds = wrapper.getBoundingClientRect();
      const y = event.clientY - bounds.top + wrapper.scrollTop;
      const x = event.clientX - bounds.left + wrapper.scrollLeft;
      const row = Math.min(Math.max(1, Math.floor(y / rowHeight) + 1), Math.max(1, totalRows));
      const colIdx = Math.min(
        Math.max(0, indexForOffset(columnOffsets, x)),
        Math.max(0, totalCols - 1),
      );
      setFillPreview(
        unionRects(source, { startRow: row, endRow: row, startColIdx: colIdx, endColIdx: colIdx }),
      );
    };
    const onUp = () => {
      const target = fillPreviewRef.current;
      const source = fillStartRef.current;
      fillStartRef.current = null;
      if (source && target) {
        const edits = fillEditsFor(currentSheet, source, target);
        if (editable && edits.length > 0) onEditCells?.(currentSheet.name, edits);
        setAnchor({ row: target.startRow, colIdx: target.startColIdx });
        setActiveCell({ row: target.endRow, colIdx: target.endColIdx });
      }
      setFillPreview(null);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [rowHeight, columnOffsets, totalRows, totalCols, currentSheet, editable, onEditCells]);

  // ------------------------------------------------------- column resizing

  useEffect(() => {
    if (!resizeDrag) return;
    const onMove = (event: MouseEvent) => {
      const next = Math.round(
        Math.min(
          MAX_COLUMN_WIDTH,
          Math.max(MIN_COLUMN_WIDTH, resizeDrag.startWidth + (event.clientX - resizeDrag.startX)),
        ),
      );
      setColumnWidths((previous) =>
        previous[resizeDrag.colIdx] === next
          ? previous
          : { ...previous, [resizeDrag.colIdx]: next },
      );
    };
    const onUp = () => setResizeDrag(null);
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [resizeDrag]);

  const applyColumnWidth = useCallback((colIdx: number, width: number) => {
    setColumnWidths((previous) =>
      previous[colIdx] === width ? previous : { ...previous, [colIdx]: width },
    );
  }, []);

  // ----------------------------------------------------------------- events

  const handleCellSelect = useCallback(
    (row: number, colIdx: number, event: React.MouseEvent) => {
      gridFocusedRef.current = true;
      setContextMenu(null);
      const editing = latest.current.editing;
      if (editing && (editing.row !== row || editing.colIdx !== colIdx)) setEditing(null);
      selectPosition({ row, colIdx }, event.shiftKey);
    },
    [selectPosition],
  );

  const handleCellActivate = useCallback(
    (row: number, colIdx: number) => {
      selectPosition({ row, colIdx }, false);
      beginEdit({ row, colIdx });
    },
    [selectPosition, beginEdit],
  );

  const cancelEdit = useCallback(() => setEditing(null), []);

  const pasteAt = useCallback(
    (text: string, row: number, colIdx: number) => {
      // The editor unmounts as this runs, so the block is anchored at this cell explicitly rather
      // than relying on a selection that is about to change.
      setEditing(null);
      pasteText(text, { row, colIdx });
    },
    [pasteText],
  );

  const beginFillDrag = useCallback((event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    fillStartRef.current = latest.current.selection;
    setFillPreview(latest.current.selection);
  }, []);

  const actionsRef = useRef<CellActions>({
    select: handleCellSelect,
    activate: handleCellActivate,
    showContextMenu: () => {},
    focus: () => {},
    commitEdit,
    cancelEdit,
    pasteBlock: pasteAt,
    beginFill: beginFillDrag,
  });
  actionsRef.current = {
    select: handleCellSelect,
    activate: handleCellActivate,
    showContextMenu: (row, colIdx, event) => {
      event.preventDefault();
      event.stopPropagation();
      handleCellSelect(row, colIdx, event);
      setContextMenu({ x: event.clientX, y: event.clientY, rowIdx: row - 1, colIdx });
    },
    focus: (row, colIdx) => {
      gridFocusedRef.current = true;
      // Moving focus to the in-cell editor bubbles a focus event to its cell; following it there
      // would drag the caret straight back.
      if (latest.current.editing) return;
      setActiveCell((previous) =>
        previous.row === row && previous.colIdx === colIdx ? previous : { row, colIdx },
      );
    },
    commitEdit,
    cancelEdit,
    pasteBlock: pasteAt,
    beginFill: beginFillDrag,
  };

  const handleKeyDown = useCallback(
    (event: KeyboardEvent) => {
      if (isDialogOpen() || isEditableTarget(event.target)) return;

      const mod = event.ctrlKey || event.metaKey;
      const key = event.key;

      if (mod && !event.altKey && key.toLowerCase() === 'a' && totalRows > 0) {
        event.preventDefault();
        setAnchor({ row: 1, colIdx: 0 });
        setActiveCell({ row: totalRows, colIdx: Math.max(0, totalCols - 1) });
        return;
      }

      const movement: Record<string, [number, number]> = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
      };
      const step = movement[key];
      if (step) {
        event.preventDefault();
        moveActive(step[0], step[1], event.shiftKey);
        return;
      }

      switch (key) {
        case 'Tab':
          event.preventDefault();
          moveActive(0, event.shiftKey ? -1 : 1, false);
          return;
        case 'Enter':
          event.preventDefault();
          moveActive(event.shiftKey ? -1 : 1, 0, false);
          return;
        case 'F2':
          event.preventDefault();
          beginEdit(activeCell);
          return;
        case 'Delete':
        case 'Backspace':
          event.preventDefault();
          clearSelection();
          return;
        case 'Home':
          event.preventDefault();
          selectPosition(
            mod ? { row: 1, colIdx: 0 } : { row: activeCell.row, colIdx: 0 },
            event.shiftKey,
          );
          return;
        case 'End': {
          event.preventDefault();
          const rowWidth = currentSheet.rows[activeCell.row - 1]?.length ?? 0;
          selectPosition(
            mod
              ? { row: totalRows, colIdx: Math.max(0, totalCols - 1) }
              : { row: activeCell.row, colIdx: Math.max(0, rowWidth - 1) },
            event.shiftKey,
          );
          return;
        }
        case 'Escape':
          // Escape abandons a fill and collapses a range back to the active cell, which is how a
          // spreadsheet backs out of a gesture.
          fillStartRef.current = null;
          setFillPreview(null);
          setAnchor(activeCell);
          return;
        default:
          break;
      }

      // Type-to-overwrite: any printable character starts editing the selected cell with it.
      if (!editable || event.altKey || mod || event.shiftKey) return;
      if (key.length !== 1) return;
      if (!gridFocusedRef.current) return;
      event.preventDefault();
      beginEdit(activeCell, key);
    },
    [
      activeCell,
      beginEdit,
      clearSelection,
      currentSheet,
      editable,
      moveActive,
      selectPosition,
      totalCols,
      totalRows,
    ],
  );

  const handleCopyEvent = useCallback(
    (event: ClipboardEvent, mode: 'copy' | 'cut') => {
      if (isDialogOpen() || isEditableTarget(event.target)) return;
      if (totalRows === 0 || totalCols === 0) return;
      const text = selectionText();
      if (text.length === 0) return;
      event.preventDefault();
      event.clipboardData?.setData('text/plain', text);
      clipboardRef.current = { text, mode, sheetName: currentSheet.name, rect: selection };
      // The system clipboard is best-effort: it needs a permission and a secure context, and the
      // internal buffer is what makes copy and paste work in either case.
      void navigator.clipboard?.writeText?.(text)?.catch?.(() => undefined);
    },
    [selectionText, selection, currentSheet.name, totalRows, totalCols],
  );

  const handlePasteEvent = useCallback(
    (event: ClipboardEvent) => {
      if (isDialogOpen() || isEditableTarget(event.target)) return;
      const fromEvent = event.clipboardData?.getData('text/plain') ?? '';
      const text = fromEvent || clipboardRef.current?.text || '';
      if (text.length === 0) {
        // No synchronous payload: ask the async clipboard once, and do nothing if it refuses.
        const pending = navigator.clipboard?.readText?.();
        if (!pending) return;
        event.preventDefault();
        void pending.then((value) => {
          if (value) pasteText(value);
        });
        return;
      }
      event.preventDefault();
      pasteText(text);
    },
    [pasteText],
  );

  // One set of window listeners reading the freshest closures through a ref, so a keypress is
  // never handled against the state of an earlier render.
  const handlersRef = useRef({ handleKeyDown, handleCopyEvent, handlePasteEvent });
  handlersRef.current = { handleKeyDown, handleCopyEvent, handlePasteEvent };

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => handlersRef.current.handleKeyDown(event);
    const onCopy = (event: ClipboardEvent) => handlersRef.current.handleCopyEvent(event, 'copy');
    const onCut = (event: ClipboardEvent) => handlersRef.current.handleCopyEvent(event, 'cut');
    const onPaste = (event: ClipboardEvent) => handlersRef.current.handlePasteEvent(event);
    window.addEventListener('keydown', onKeyDown);
    window.addEventListener('copy', onCopy);
    window.addEventListener('cut', onCut);
    window.addEventListener('paste', onPaste);
    return () => {
      window.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('copy', onCopy);
      window.removeEventListener('cut', onCut);
      window.removeEventListener('paste', onPaste);
    };
  }, []);

  useEffect(() => {
    if (!contextMenu) return;
    const close = () => setContextMenu(null);
    window.addEventListener('click', close);
    window.addEventListener('contextmenu', close);
    window.addEventListener('scroll', close, true);
    return () => {
      window.removeEventListener('click', close);
      window.removeEventListener('contextmenu', close);
      window.removeEventListener('scroll', close, true);
    };
  }, [contextMenu]);

  // -------------------------------------------------------------- rendering

  const selectedCellCoord = `${indexToColumn(activeCell.colIdx)}${activeCell.row}`;
  const currentCell = currentSheet.rows[activeCell.row - 1]?.[activeCell.colIdx];
  const cellValue = currentCell?.value;
  const cellFormula = currentCell?.formula;
  const formulaBarText = cellDisplay(
    currentCell,
    evaluateCell(currentCell, currentSheet.name, activeCell.colIdx, activeCell.row),
    dateSystem,
  ).text;

  let cellTypeStr = 'empty';
  if (cellFormula) cellTypeStr = 'formula';
  else if (cellValue instanceof Date) cellTypeStr = 'date';
  else if (typeof cellValue === 'number') cellTypeStr = 'number';
  else if (typeof cellValue === 'boolean') cellTypeStr = 'boolean';
  else if (typeof cellValue === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(cellValue) || /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cellValue)) {
      cellTypeStr = 'date';
    } else {
      cellTypeStr = 'string';
    }
  }

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };
  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file && onFileDrop) onFileDrop(file);
  };

  const topSpacerHeight = rowWindow.start * rowHeight;
  const bottomSpacerHeight = Math.max(0, (totalRows - rowWindow.end) * rowHeight);
  const spacerColSpan = Math.max(
    1,
    columnIndicesLength(columnWindow) +
      (columnWindow.left > 0 ? 1 : 0) +
      (columnWindow.right > 0 ? 1 : 0) +
      1,
  );

  const columnIndices = columnRange(columnWindow.start, columnWindow.end);
  const rowIndices = columnRange(rowWindow.start, rowWindow.end);

  return (
    <div
      className="grid-panel"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {isDragOver && (
        <div className="grid-drag-overlay">
          <div className="drag-overlay-card">
            <div className="drag-icon-bounce">
              <svg
                width="24"
                height="24"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                <line x1="3" y1="9" x2="21" y2="9" />
                <line x1="3" y1="15" x2="21" y2="15" />
                <line x1="9" y1="3" x2="9" y2="21" />
                <line x1="15" y1="3" x2="15" y2="21" />
              </svg>
            </div>
            <h3 style={{ fontSize: '15px', fontWeight: 600 }}>Drop your Excel or CSV file here</h3>
            <p style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
              Excel Agent will inspect and parse your spreadsheet locally
            </p>
          </div>
        </div>
      )}

      <div className="formula-bar">
        <div className="cell-coord-badge">{selectedCellCoord}</div>
        <div className="formula-type-badge">{cellTypeStr}</div>
        <div className="formula-input-display">
          {cellFormula ? (
            <span className="cell-val-formula">{formulaTextFor(currentCell)}</span>
          ) : cellValue !== null && cellValue !== undefined ? (
            // A Date read as a raw string becomes "Fri Jan 01 2021 00:00:00 GMT+0000"; the grid
            // shows it the way its number format renders it.
            formulaBarText || String(cellValue)
          ) : (
            <span style={{ color: 'var(--text-subtle)' }}>(empty)</span>
          )}
        </div>
      </div>

      <div className="grid-scroll-wrapper" ref={scrollWrapperRef} onScroll={handleScroll}>
        <table
          className="spreadsheet-table"
          role="grid"
          aria-label={`${currentSheet.name} spreadsheet`}
          aria-readonly={!editable}
          aria-rowcount={totalRows + 1}
          aria-colcount={totalCols + 1}
        >
          <thead>
            <tr role="row" aria-rowindex={1}>
              <th className="corner-cell" aria-hidden="true" />
              {columnWindow.left > 0 && (
                <th
                  role="presentation"
                  aria-hidden="true"
                  className="spreadsheet-spacer"
                  style={{ width: `${columnWindow.left}px` }}
                />
              )}
              {columnIndices.map((colIdx) => {
                const colLetter = indexToColumn(colIdx);
                const width = widthForColumn(colIdx);

                return (
                  <th
                    key={colLetter}
                    scope="col"
                    role="columnheader"
                    aria-colindex={colIdx + 2}
                    // Every column offers a sort, and none is sorted until one is pressed.
                    aria-sort="none"
                    className="column-header"
                    style={{
                      height: `${rowHeight}px`,
                      width: `${width}px`,
                      minWidth: `${width}px`,
                      maxWidth: `${width}px`,
                    }}
                    onContextMenu={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      setContextMenu({ x: e.clientX, y: e.clientY, rowIdx: 0, colIdx });
                    }}
                  >
                    <div className="col-header-inner">
                      <span className="col-letter">{colLetter}</span>

                      {onQuickSort && (
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            style={{ padding: '0 3px', height: '18px', fontSize: '10px' }}
                            title={`Sort Column ${colLetter} Ascending`}
                            aria-label={`Sort column ${colLetter} ascending`}
                            onClick={(e) => {
                              e.stopPropagation();
                              onQuickSort(colLetter, 'asc');
                            }}
                          >
                            ▲
                          </button>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            style={{ padding: '0 3px', height: '18px', fontSize: '10px' }}
                            title={`Sort Column ${colLetter} Descending`}
                            aria-label={`Sort column ${colLetter} descending`}
                            onClick={(e) => {
                              e.stopPropagation();
                              onQuickSort(colLetter, 'desc');
                            }}
                          >
                            ▼
                          </button>
                        </div>
                      )}
                    </div>

                    <div
                      role="separator"
                      aria-orientation="vertical"
                      aria-label={`Resize column ${colLetter}`}
                      aria-valuenow={width}
                      aria-valuemin={MIN_COLUMN_WIDTH}
                      aria-valuemax={MAX_COLUMN_WIDTH}
                      tabIndex={0}
                      className="col-resize-handle"
                      onMouseDown={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        gridFocusedRef.current = false;
                        setResizeDrag({ colIdx, startX: e.clientX, startWidth: width });
                      }}
                      onKeyDown={(e) => {
                        if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
                        e.preventDefault();
                        e.stopPropagation();
                        applyColumnWidth(
                          colIdx,
                          Math.min(
                            MAX_COLUMN_WIDTH,
                            Math.max(
                              MIN_COLUMN_WIDTH,
                              width + (e.key === 'ArrowRight' ? 1 : -1) * RESIZE_KEYBOARD_STEP,
                            ),
                          ),
                        );
                      }}
                    />
                  </th>
                );
              })}
              {columnWindow.right > 0 && (
                <th
                  role="presentation"
                  aria-hidden="true"
                  className="spreadsheet-spacer"
                  style={{ width: `${columnWindow.right}px` }}
                />
              )}
            </tr>
          </thead>
          <tbody>
            {topSpacerHeight > 0 && (
              <tr aria-hidden="true" style={{ height: `${topSpacerHeight}px` }}>
                <td
                  role="presentation"
                  colSpan={spacerColSpan}
                  style={{ padding: 0, border: 'none' }}
                />
              </tr>
            )}

            {rowIndices.map((rowIdx, rowOffset) => {
              const rowNumber = rowIdx + 1;
              const rowCells = currentSheet.rows[rowIdx] ?? [];
              const isHeaderRow = rowNumber === 1;

              return (
                <tr
                  key={rowNumber}
                  role="row"
                  aria-rowindex={rowNumber + 1}
                  ref={rowOffset === 0 ? probeRowRef : undefined}
                  style={{ height: `${rowHeight}px` }}
                >
                  <th
                    scope="row"
                    role="rowheader"
                    aria-rowindex={rowNumber + 1}
                    className="row-index-cell"
                    style={{ height: `${rowHeight}px` }}
                  >
                    {rowNumber}
                  </th>
                  {columnWindow.left > 0 && (
                    <td
                      role="presentation"
                      aria-hidden="true"
                      className="spreadsheet-spacer"
                      style={{ width: `${columnWindow.left}px` }}
                    />
                  )}
                  {columnIndices.map((colIdx) => {
                    const colLetter = indexToColumn(colIdx);
                    const cell = rowCells[colIdx];
                    const isSelected =
                      rowNumber >= selection.startRow &&
                      rowNumber <= selection.endRow &&
                      colIdx >= selection.startColIdx &&
                      colIdx <= selection.endColIdx;
                    const cellKey = `${currentSheet.name}:${rowNumber}:${colLetter}`;
                    const isDiffChanged = recentChangedCells.has(cellKey);
                    const isSearchMatch = searchHighlightCells.has(cellKey);
                    const isEditingThis = editing?.row === rowNumber && editing?.colIdx === colIdx;
                    const isActive = activeCell.row === rowNumber && activeCell.colIdx === colIdx;
                    const inFillPreview =
                      fillPreview !== null &&
                      fillPreview.startRow <= rowNumber &&
                      fillPreview.endRow >= rowNumber &&
                      fillPreview.startColIdx <= colIdx &&
                      fillPreview.endColIdx >= colIdx &&
                      !isSelected;

                    return (
                      <GridCell
                        key={colLetter}
                        cell={cell}
                        sheetName={currentSheet.name}
                        evaluate={evaluateCell}
                        dateSystem={dateSystem}
                        rowNumber={rowNumber}
                        colIdx={colIdx}
                        address={`${colLetter}${rowNumber}`}
                        isHeaderRow={isHeaderRow}
                        isSelected={isSelected}
                        isActive={isActive}
                        isDiffChanged={isDiffChanged}
                        isSearchMatch={isSearchMatch}
                        readOnly={!editable}
                        width={columnWidths[colIdx]}
                        rowHeight={rowHeight}
                        editorText={isEditingThis ? (editing?.text ?? '') : null}
                        showFillHandle={
                          editable &&
                          !editing &&
                          totalRows > 0 &&
                          selection.endRow === rowNumber &&
                          selection.endColIdx === colIdx
                        }
                        isFillPreview={inFillPreview}
                        actions={actionsRef}
                        innerRef={isActive ? activeCellRef : undefined}
                      />
                    );
                  })}
                  {columnWindow.right > 0 && (
                    <td
                      role="presentation"
                      aria-hidden="true"
                      className="spreadsheet-spacer"
                      style={{ width: `${columnWindow.right}px` }}
                    />
                  )}
                </tr>
              );
            })}

            {bottomSpacerHeight > 0 && (
              <tr aria-hidden="true" style={{ height: `${bottomSpacerHeight}px` }}>
                <td
                  role="presentation"
                  colSpan={spacerColSpan}
                  style={{ padding: 0, border: 'none' }}
                />
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {contextMenu && (
        <div
          className="cell-context-menu"
          role="menu"
          aria-label="Cell actions"
          style={{ position: 'fixed', left: contextMenu.x, top: contextMenu.y, zIndex: 1000 }}
          onClick={(e) => e.stopPropagation()}
          onContextMenu={(e) => e.preventDefault()}
        >
          {(['cell', 'row', 'column'] as const).map((kind) => (
            <button
              key={kind}
              type="button"
              role="menuitem"
              className="cell-context-menu-item"
              onClick={() => {
                const ctx = describeCellSelection(
                  kind,
                  currentSheet,
                  contextMenu.rowIdx,
                  contextMenu.colIdx,
                );
                onAddSelectionContext?.(ctx);
                setContextMenu(null);
              }}
            >
              Ask agent about this {kind}
            </button>
          ))}
        </div>
      )}

      <div className="sheet-tabs-bar" aria-label="Sheets">
        {workbook.sheets.map((sheet) => {
          const isActive = sheet.name === currentSheet.name;
          const count = sheet.rows.length;

          return (
            <button
              key={sheet.name}
              type="button"
              aria-pressed={isActive}
              className={`sheet-tab ${isActive ? 'active' : ''}`}
              onClick={() => onSelectSheet(sheet.name)}
            >
              <span className="sheet-tab-title">{sheet.name}</span>
              <span className="sheet-tab-count">({count})</span>
            </button>
          );
        })}
      </div>
    </div>
  );
};

/** `[start, end)` as a plain array. A window is small by construction, so this stays cheap. */
function columnRange(start: number, end: number): number[] {
  const out: number[] = [];
  for (let i = start; i < end; i += 1) out.push(i);
  return out;
}

function columnIndicesLength(window: { start: number; end: number }): number {
  return Math.max(0, window.end - window.start);
}
