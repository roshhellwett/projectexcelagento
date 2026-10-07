import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { Cell, DateSystem } from '@excel-agent/engine';

import { cellDisplay, type CellEvaluator } from '../lib/cell-evaluator.js';

/** Where the caret goes once an edit is committed. `none` keeps it where it is. */
export type EditorMove = 'none' | 'down' | 'up' | 'right' | 'left';

interface CellEditorProps {
  /** The text the edit began with, or the keystroke that overwrote the cell. */
  initialText: string;
  width: number | undefined;
  onCommit: (raw: string, move: EditorMove) => void;
  onCancel: () => void;
  /** A paste containing tabs or newlines is a block, not one cell's text. */
  onPasteBlock: (text: string) => void;
}

/**
 * The in-cell editor.
 *
 * It sits on top of the cell it is editing rather than in the formula bar, which is what makes it
 * feel like a spreadsheet: Enter commits and steps down, Tab commits and steps right, Escape throws
 * the edit away, and a multi-line paste is handed back to the grid to scatter over a range instead
 * of landing as one cell full of tabs.
 */
const CellEditor: React.FC<CellEditorProps> = ({
  initialText,
  width,
  onCommit,
  onCancel,
  onPasteBlock,
}) => {
  const [text, setText] = useState(initialText);
  const inputRef = useRef<HTMLInputElement>(null);
  const committedRef = useRef(false);
  const mountedRef = useRef(false);
  // Latest values, read by the unmount cleanup. Depending on them instead would re-arm that
  // cleanup on every keystroke and commit the same edit over and over.
  const textRef = useRef(text);
  textRef.current = text;
  const commitRef = useRef(onCommit);
  commitRef.current = onCommit;

  const commit = (move: EditorMove) => {
    if (committedRef.current) return;
    committedRef.current = true;
    commitRef.current(textRef.current, move);
  };

  useLayoutEffect(() => {
    const input = inputRef.current;
    if (!input) return;
    input.focus();
    // The caret goes to the end, which is where a spreadsheet puts it for F2 and a double-click.
    const caret = input.value.length;
    input.setSelectionRange(caret, caret);
  }, []);

  // Losing the grid mid-edit (a modal opens and takes focus) commits rather than discards.
  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      // Strict Mode probes cleanup and immediately re-mounts effects. Only a real unmount
      // may commit: otherwise opening the editor closes it before the first keystroke.
      queueMicrotask(() => {
        if (!mountedRef.current && !committedRef.current) {
          committedRef.current = true;
          commitRef.current(textRef.current, 'none');
        }
      });
    };
  }, []);

  return (
    <input
      ref={inputRef}
      id="grid-cell-editor"
      className="cell-editor"
      // The editor is a real text field, so it is reachable by name and by keyboard.
      aria-label="Cell editor"
      type="text"
      value={text}
      style={width === undefined ? undefined : { width: `${width}px` }}
      onChange={(e) => setText(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          e.stopPropagation();
          commit(e.shiftKey ? 'up' : 'down');
        } else if (e.key === 'Tab') {
          e.preventDefault();
          e.stopPropagation();
          commit(e.shiftKey ? 'left' : 'right');
        } else if (e.key === 'Escape') {
          e.preventDefault();
          e.stopPropagation();
          committedRef.current = true;
          onCancel();
        }
      }}
      onBlur={() => commit('none')}
      onPaste={(e) => {
        const pasted = e.clipboardData?.getData('text/plain') ?? '';
        // A single value with no separators belongs in this cell; anything structured is a block.
        if (pasted.length === 0 || !/[\t\n]/.test(pasted)) return;
        e.preventDefault();
        e.stopPropagation();
        committedRef.current = true;
        onPasteBlock(pasted);
      }}
    />
  );
};

/**
 * Everything a cell can ask the grid to do. These are stable identities, and the cell's own address
 * travels with the call: an inline arrow per cell would be a new prop on every render, and the
 * memoisation below would never once get to skip anything.
 */
export interface CellActions {
  select: (row: number, colIdx: number, event: React.MouseEvent) => void;
  activate: (row: number, colIdx: number) => void;
  showContextMenu: (row: number, colIdx: number, event: React.MouseEvent) => void;
  focus: (row: number, colIdx: number) => void;
  commitEdit: (raw: string, move: EditorMove) => void;
  cancelEdit: () => void;
  pasteBlock: (text: string, row: number, colIdx: number) => void;
  beginFill: (event: React.MouseEvent) => void;
}

interface GridCellProps {
  cell: Cell | undefined;
  sheetName: string;
  evaluate: CellEvaluator;
  dateSystem: DateSystem;
  rowNumber: number;
  colIdx: number;
  address: string;
  isHeaderRow: boolean;
  isSelected: boolean;
  isActive: boolean;
  isDiffChanged: boolean;
  isSearchMatch: boolean;
  /**
   * True when the grid cannot be edited - either no writer was supplied, or a dialog is on top of
   * it. Exposed to assistive tech as `aria-readonly`.
   */
  readOnly: boolean;
  width: number | undefined;
  rowHeight: number;
  editorText: string | null;
  showFillHandle: boolean;
  isFillPreview: boolean;
  actions: React.RefObject<CellActions>;
  innerRef?: React.Ref<HTMLTableCellElement>;
}

/**
 * One grid cell.
 *
 * Memoised on its own props, and its text is derived inside a `useMemo` keyed on the cell's *value*
 * rather than its object identity. Both matter: the engine hands back a freshly cloned workbook after
 * every edit, so identity-keying would re-derive every visible cell (and re-evaluate every visible
 * formula) for a one-character change. Keyed on the value, an untouched cell re-renders at worst and
 * never re-evaluates.
 */
function renderCellContent(text: string, isHeaderRow: boolean): React.ReactNode {
  if (isHeaderRow || !text) return text;
  const lower = text.trim().toLowerCase();
  if (lower === 'completed') {
    return (
      <span className="status-pill status-completed">
        <span className="status-pill-dot" />
        {text}
      </span>
    );
  }
  if (lower === 'pending') {
    return (
      <span className="status-pill status-pending">
        <span className="status-pill-dot" />
        {text}
      </span>
    );
  }
  if (lower === 'processing') {
    return (
      <span className="status-pill status-processing">
        <span className="status-pill-dot" />
        {text}
      </span>
    );
  }
  if (lower === 'cancelled') {
    return (
      <span className="status-pill status-cancelled">
        <span className="status-pill-dot" />
        {text}
      </span>
    );
  }
  if (lower === 'draft') {
    return (
      <span className="status-pill status-draft">
        <span className="status-pill-dot" />
        {text}
      </span>
    );
  }
  return text;
}

const GridCellBase: React.FC<GridCellProps> = ({
  cell,
  sheetName,
  evaluate,
  dateSystem,
  rowNumber,
  colIdx,
  address,
  isHeaderRow,
  isSelected,
  isActive,
  isDiffChanged,
  isSearchMatch,
  readOnly,
  width,
  rowHeight,
  editorText,
  showFillHandle,
  isFillPreview,
  actions,
  innerRef,
}) => {
  const value = cell?.value;
  // A Date is compared by its instant so cloning the workbook does not look like a change.
  const valueKey = value instanceof Date ? value.getTime() : value;
  const formulaKey = cell?.formula;
  const formatKey = cell?.numberFormat;
  const styleKey = cell?.style ? JSON.stringify(cell.style) : '';

  const display = useMemo(
    () => cellDisplay(cell, evaluate(cell, sheetName, colIdx, rowNumber), dateSystem),
    [valueKey, formulaKey, formatKey, styleKey, evaluate, dateSystem, colIdx, rowNumber, sheetName],
  );

  // A click that immediately follows this cell's own mousedown is the same gesture, already handled.
  const mouseDownKeyRef = useRef<string | null>(null);

  let cellClass = 'data-cell';
  if (isHeaderRow) cellClass += ' cell-header-row';
  if (isSelected) cellClass += ' selected';
  if (isDiffChanged) cellClass += ' diff-changed';
  if (isSearchMatch) cellClass += ' cell-search-match';
  if (isFillPreview) cellClass += ' fill-preview';
  if (display.className) cellClass += ` ${display.className}`;

  const sizeStyle: React.CSSProperties = {
    height: `${rowHeight}px`,
    ...(cell?.style?.fontColor ? { color: cell.style.fontColor } : {}),
    ...(cell?.style?.fillColor ? { backgroundColor: cell.style.fillColor } : {}),
    ...(cell?.style?.bold ? { fontWeight: 700 } : {}),
    ...(cell?.style?.italic ? { fontStyle: 'italic' } : {}),
    ...(cell?.style?.underline ? { textDecoration: 'underline' } : {}),
    ...(cell?.style?.horizontalAlignment ? { textAlign: cell.style.horizontalAlignment } : {}),
    ...(cell?.style?.verticalAlignment
      ? {
          verticalAlign:
            cell.style.verticalAlignment === 'middle' ? 'middle' : cell.style.verticalAlignment,
        }
      : {}),
    ...(cell?.style?.wrapText ? { whiteSpace: 'normal', overflowWrap: 'anywhere' } : {}),
    ...(width === undefined
      ? {}
      : { width: `${width}px`, minWidth: `${width}px`, maxWidth: `${width}px` }),
  };

  return (
    <td
      ref={innerRef}
      role="gridcell"
      // ARIA indices count the header row and the row-number column, so data cell A1 is 2,2.
      aria-colindex={colIdx + 2}
      aria-selected={isSelected}
      aria-readonly={readOnly}
      aria-label={`${address}: ${display.text || 'empty'}`}
      // Roving tabindex: exactly one cell is in the tab order, and it is the active one.
      tabIndex={isActive && !readOnly ? 0 : -1}
      className={cellClass}
      style={sizeStyle}
      title={`${address}: ${display.text || '(empty)'}`}
      onMouseDown={(e) => {
        mouseDownKeyRef.current = address;
        actions.current.select(rowNumber, colIdx, e);
      }}
      onClick={(e) => {
        if (mouseDownKeyRef.current === address) {
          mouseDownKeyRef.current = null;
          return;
        }
        actions.current.select(rowNumber, colIdx, e);
      }}
      onDoubleClick={() => actions.current.activate(rowNumber, colIdx)}
      onContextMenu={(e) => actions.current.showContextMenu(rowNumber, colIdx, e)}
      onFocus={() => actions.current.focus(rowNumber, colIdx)}
    >
      {editorText === null ? (
        renderCellContent(display.text, isHeaderRow)
      ) : (
        <CellEditor
          initialText={editorText}
          width={width}
          onCommit={actions.current.commitEdit}
          onCancel={actions.current.cancelEdit}
          onPasteBlock={(text) => actions.current.pasteBlock(text, rowNumber, colIdx)}
        />
      )}
      {showFillHandle && (
        <button
          type="button"
          className="fill-handle"
          aria-label={`Fill from ${address}`}
          title="Drag to fill"
          onMouseDown={actions.current.beginFill}
          onClick={(e) => e.stopPropagation()}
          onDoubleClick={(e) => e.stopPropagation()}
        />
      )}
    </td>
  );
};

export const GridCell = React.memo(GridCellBase);
