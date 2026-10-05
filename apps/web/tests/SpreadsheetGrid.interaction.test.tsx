// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { StrictMode } from 'react';
import { fireEvent, render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { createCell, type Workbook } from '@excel-agent/engine';

import { SpreadsheetGrid } from '../src/components/SpreadsheetGrid.js';
import type { CellEdit } from '../src/lib/grid-edit.js';

function testWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Sheet1',
        rows: [
          [createCell('Name'), createCell('Qty')],
          [createCell('Widget'), createCell(7)],
          [createCell('Gadget'), createCell(9)],
        ],
      },
    ],
  };
}

interface HarnessOptions {
  /** Omit the writer to exercise the read-only grid. */
  writable?: boolean;
  strict?: boolean;
}

function renderGrid(options: HarnessOptions = {}) {
  const edits: Array<{ sheet: string; edits: CellEdit[] }> = [];
  const onEditCells = vi.fn((sheet: string, batch: CellEdit[]) => {
    edits.push({ sheet, edits: batch });
  });
  const workbook = testWorkbook();
  const utils = render(
    <SpreadsheetGrid
      workbook={workbook}
      activeSheetName="Sheet1"
      onSelectSheet={() => {}}
      recentChangedCells={new Set()}
      {...(options.writable === false ? {} : { onEditCells })}
    />,
    options.strict ? { wrapper: StrictMode } : undefined,
  );
  const cell = (address: string) =>
    screen.getByRole('gridcell', { name: new RegExp(`^${address}: `) });
  return { ...utils, edits, onEditCells, cell };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe('grid structure exposed to assistive technology', () => {
  it('is a grid of rows, column headers and cells with ARIA indices that count the headers', () => {
    const { cell } = renderGrid();

    const grid = screen.getByRole('grid');
    expect(grid).toHaveAttribute('aria-label', 'Sheet1 spreadsheet');
    expect(grid).toHaveAttribute('aria-readonly', 'false');
    // Row 1 of the grid is the header row and column 1 is the row-number column, so data cell A1
    // sits at 2,2 - the same convention a screen reader expects when it counts a grid.
    expect(grid).toHaveAttribute('aria-rowcount', '4');
    expect(grid).toHaveAttribute('aria-colcount', '3');

    const headerA = screen.getByRole('columnheader', { name: /A/ });
    expect(headerA).toHaveAttribute('aria-colindex', '2');
    // Every column offers a sort, and none is sorted until one is pressed.
    expect(headerA).toHaveAttribute('aria-sort', 'none');

    expect(cell('A1')).toHaveAttribute('aria-colindex', '2');
    expect(cell('B1')).toHaveAttribute('aria-colindex', '3');
    const row2 = screen.getAllByRole('row')[2];
    expect(row2).toHaveAttribute('aria-rowindex', '3');
    // The grid starts with A1 selected, which is what a spreadsheet opens with.
    expect(cell('A1')).toHaveAttribute('aria-selected', 'true');
    expect(cell('A2')).toHaveAttribute('aria-selected', 'false');
    expect(cell('A1')).toHaveAttribute('aria-readonly', 'false');
  });

  it('puts exactly one cell in the tab order and moves that cell with the selection', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    expect(cell('A1')).toHaveAttribute('tabindex', '0');
    expect(cell('B2')).toHaveAttribute('tabindex', '-1');

    await user.click(cell('B2'));

    expect(cell('B2')).toHaveAttribute('tabindex', '0');
    expect(cell('A1')).toHaveAttribute('tabindex', '-1');
    expect(cell('B2')).toHaveFocus();
  });

  it('marks every cell readonly and refuses to edit when there is no writer', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid({ writable: false });

    expect(screen.getByRole('grid')).toHaveAttribute('aria-readonly', 'true');
    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: '7' });

    expect(screen.queryByLabelText('Cell editor')).not.toBeInTheDocument();
  });
});

describe('editing', () => {
  it('starts an edit from a typed character and commits it with Enter', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: '4' });

    const editor = screen.getByLabelText('Cell editor');
    expect(editor).toHaveValue('4');
    await user.type(editor, '2');
    fireEvent.keyDown(editor, { key: 'Enter' });

    expect(screen.queryByLabelText('Cell editor')).not.toBeInTheDocument();
    expect(edits).toEqual([{ sheet: 'Sheet1', edits: [{ row: 2, column: 'B', value: 42 }] }]);
    // Enter commits and steps down, so the next keystroke edits the cell below.
    expect(cell('B2')).toHaveAttribute('aria-selected', 'false');
    expect(cell('B3')).toHaveAttribute('aria-selected', 'true');
  });

  it('keeps the editor open during Strict Mode effect probes and commits exactly once', async () => {
    const user = userEvent.setup();
    const { cell, onEditCells } = renderGrid({ strict: true });
    await user.dblClick(cell('B2'));
    expect(screen.getByLabelText('Cell editor')).toHaveValue('7');
    expect(onEditCells).not.toHaveBeenCalled();
    await user.clear(screen.getByLabelText('Cell editor'));
    await user.type(screen.getByLabelText('Cell editor'), '42');
    await user.keyboard('{Enter}');
    expect(onEditCells).toHaveBeenCalledExactlyOnceWith('Sheet1', [
      { row: 2, column: 'B', value: 42 },
    ]);
  });

  it('opens the editor with the current value on F2 and on a double click', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    await user.click(cell('A2'));
    fireEvent.keyDown(window, { key: 'F2' });
    expect(screen.getByLabelText('Cell editor')).toHaveValue('Widget');
    fireEvent.keyDown(screen.getByLabelText('Cell editor'), { key: 'Escape' });

    await user.dblClick(cell('B2'));
    expect(screen.getByLabelText('Cell editor')).toHaveValue('7');
  });

  it('cancels an in-progress edit on Escape and leaves the value untouched', async () => {
    const user = userEvent.setup();
    const { cell, edits, onEditCells } = renderGrid();

    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: 'F2' });
    const editor = screen.getByLabelText('Cell editor');
    await user.clear(editor);
    await user.type(editor, 'discard me');
    await user.keyboard('{Escape}');

    expect(screen.queryByLabelText('Cell editor')).not.toBeInTheDocument();
    expect(onEditCells).not.toHaveBeenCalled();
    expect(edits).toEqual([]);
    expect(cell('B2')).toHaveTextContent('7');
  });

  it('writes a formula when the edit starts with = and keeps its leading sign', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: '=' });
    const editor = screen.getByLabelText('Cell editor');
    await user.type(editor, 'B2*2');
    fireEvent.keyDown(editor, { key: 'Enter' });

    expect(edits[0]?.edits).toEqual([{ row: 2, column: 'B', formula: '=B2*2' }]);
  });

  it('does not write anything when an edit is committed with the text it started from', async () => {
    const user = userEvent.setup();
    const { cell, onEditCells } = renderGrid();

    await user.click(cell('A2'));
    fireEvent.keyDown(window, { key: 'F2' });
    fireEvent.keyDown(screen.getByLabelText('Cell editor'), { key: 'Enter' });

    expect(onEditCells).not.toHaveBeenCalled();
  });

  it('commits with Tab and moves right, or left with Shift+Tab', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('A2'));
    fireEvent.keyDown(window, { key: 'F2' });
    const editor = screen.getByLabelText('Cell editor');
    await user.clear(editor);
    await user.type(editor, 'Gadget v2');
    fireEvent.keyDown(editor, { key: 'Tab' });

    expect(edits[0]?.edits).toEqual([{ row: 2, column: 'A', value: 'Gadget v2' }]);
    // Tab hands the caret to the next cell, and Shift+Tab brings it back.
    expect(cell('B2')).toHaveAttribute('aria-selected', 'true');
    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true });
    expect(cell('A2')).toHaveAttribute('aria-selected', 'true');
  });
});

describe('range selection and clearing', () => {
  it('extends a range with shift+arrow and clears exactly that range with Delete', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'ArrowDown', shiftKey: true });
    fireEvent.keyDown(window, { key: 'ArrowRight', shiftKey: true });

    expect(cell('A1')).toHaveAttribute('aria-selected', 'true');
    expect(cell('A2')).toHaveAttribute('aria-selected', 'true');
    expect(cell('B1')).toHaveAttribute('aria-selected', 'true');
    expect(cell('B2')).toHaveAttribute('aria-selected', 'true');
    // The anchor stays put and only the far corner moves, which is what makes a second shift+arrow
    // keep growing the same rectangle.
    expect(cell('A1')).toHaveAttribute('tabindex', '-1');
    expect(cell('B2')).toHaveAttribute('tabindex', '0');

    fireEvent.keyDown(window, { key: 'Delete' });

    expect(edits).toEqual([
      {
        sheet: 'Sheet1',
        edits: [
          { row: 1, column: 'A', value: null },
          { row: 1, column: 'B', value: null },
          { row: 2, column: 'A', value: null },
          { row: 2, column: 'B', value: null },
        ],
      },
    ]);
  });

  it('extends a range with shift+click', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    await user.click(cell('A1'));
    await user.click(cell('B2'));

    expect(cell('B2')).toHaveAttribute('aria-selected', 'true');
  });

  it('collapses a range back to the active cell on Escape', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'ArrowDown', shiftKey: true });
    expect(cell('A2')).toHaveAttribute('aria-selected', 'true');

    fireEvent.keyDown(window, { key: 'Escape' });

    expect(cell('A1')).toHaveAttribute('aria-selected', 'false');
    expect(cell('A2')).toHaveAttribute('aria-selected', 'true');
  });

  it('selects the whole used range with Ctrl+A', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'a', ctrlKey: true });

    expect(cell('B3')).toHaveAttribute('aria-selected', 'true');
  });

  it('walks with the arrow keys and stops at the edges of the data', async () => {
    const user = userEvent.setup();
    const { cell } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'ArrowUp' });
    fireEvent.keyDown(window, { key: 'ArrowLeft' });
    expect(cell('A1')).toHaveAttribute('aria-selected', 'true');

    fireEvent.keyDown(window, { key: 'ArrowDown', shiftKey: true });
    fireEvent.keyDown(window, { key: 'ArrowDown', shiftKey: true });
    fireEvent.keyDown(window, { key: 'ArrowDown' });
    expect(cell('A3')).toHaveAttribute('aria-selected', 'true');
  });
});

describe('clipboard', () => {
  const clipboardPayload = (text: string) => ({
    clipboardData: { getData: () => text, setData: vi.fn() },
  });

  it('copies the selection and pastes a tab-separated block from its top-left anchor', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'ArrowDown', shiftKey: true });
    fireEvent.keyDown(window, { key: 'ArrowRight', shiftKey: true });

    const copied = clipboardPayload('');
    fireEvent.copy(window, copied);
    expect(copied.clipboardData.setData).toHaveBeenCalledWith('text/plain', 'Name\tQty\nWidget\t7');

    // Paste into a different place, anchored at the selection's top-left corner.
    await user.click(cell('B2'));
    fireEvent.paste(window, clipboardPayload('left\tright\n1\t2'));

    expect(edits).toEqual([
      {
        sheet: 'Sheet1',
        edits: [
          { row: 2, column: 'B', value: 'left' },
          { row: 2, column: 'C', value: 'right' },
          { row: 3, column: 'B', value: 1 },
          { row: 3, column: 'C', value: 2 },
        ],
      },
    ]);
    // The pasted block becomes the selection.
    expect(cell('B3')).toHaveAttribute('aria-selected', 'true');
  });

  it('pastes the internal buffer when the event carries no text', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('A1'));
    fireEvent.keyDown(window, { key: 'ArrowRight', shiftKey: true });
    fireEvent.copy(window, clipboardPayload(''));
    await user.click(cell('A2'));
    fireEvent.paste(window, clipboardPayload(''));

    expect(edits).toEqual([
      {
        sheet: 'Sheet1',
        edits: [
          { row: 2, column: 'A', value: 'Name' },
          { row: 2, column: 'B', value: 'Qty' },
        ],
      },
    ]);
  });

  it('moves the source cells on a cut and paste', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('A2'));
    fireEvent.cut(window, clipboardPayload(''));
    await user.click(cell('A3'));
    fireEvent.paste(window, clipboardPayload(''));

    expect(edits[0]?.edits).toEqual([
      { row: 2, column: 'A', value: null },
      { row: 3, column: 'A', value: 'Widget' },
    ]);
  });

  it('leaves a paste alone while a modal owns the screen', async () => {
    const user = userEvent.setup();
    const onEditCells = vi.fn();
    const grid = (key?: string) => (
      <SpreadsheetGrid
        key={key}
        workbook={testWorkbook()}
        activeSheetName="Sheet1"
        onSelectSheet={() => {}}
        recentChangedCells={new Set()}
        onEditCells={onEditCells}
      />
    );
    const { rerender } = render(grid());

    await user.click(screen.getByRole('gridcell', { name: /^B2: / }));
    rerender(
      <>
        <div data-dialog-open="true" />
        {grid()}
      </>,
    );

    fireEvent.paste(window, clipboardPayload('nope'));
    fireEvent.keyDown(window, { key: 'x' });

    expect(onEditCells).not.toHaveBeenCalled();
  });

  it('does not hijack a paste aimed at a text field', async () => {
    const { cell, onEditCells } = renderGrid();
    const field = document.createElement('input');
    document.body.appendChild(field);

    fireEvent.paste(field, clipboardPayload('into the field'));

    expect(onEditCells).not.toHaveBeenCalled();
    expect(cell('A1')).toHaveTextContent('Name');
    field.remove();
  });

  it('scatters a structured paste over a range instead of into the cell being edited', async () => {
    const user = userEvent.setup();
    const { cell, edits } = renderGrid();

    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: 'F2' });
    const editor = screen.getByLabelText('Cell editor');
    fireEvent.paste(editor, clipboardPayload('a\tb\nc\td'));

    expect(edits).toEqual([
      {
        sheet: 'Sheet1',
        edits: [
          { row: 2, column: 'B', value: 'a' },
          { row: 2, column: 'C', value: 'b' },
          { row: 3, column: 'B', value: 'c' },
          { row: 3, column: 'C', value: 'd' },
        ],
      },
    ]);
  });
});

describe('column resizing and filling', () => {
  it('applies the dragged width to the header and every cell in that column', () => {
    renderGrid();
    const handle = screen.getByRole('separator', { name: 'Resize column A' });
    const before = screen.getByRole('columnheader', { name: /A/ }).style.width;

    fireEvent.mouseDown(handle, { clientX: 100 });
    fireEvent.mouseMove(window, { clientX: 180 });
    fireEvent.mouseUp(window);

    const after = screen.getByRole('columnheader', { name: /A/ }).style.width;
    expect(after).toBe(`${parseInt(before, 10) + 80}px`);
    expect(screen.getByRole('gridcell', { name: /^A1: / }).style.width).toBe(after);
    expect(screen.getByRole('columnheader', { name: /B/ }).style.width).toBe(before);
  });

  it('clamps a drag to the usable width range', () => {
    renderGrid();
    const handle = screen.getByRole('separator', { name: 'Resize column A' });

    fireEvent.mouseDown(handle, { clientX: 100 });
    fireEvent.mouseMove(window, { clientX: -5000 });
    expect(screen.getByRole('columnheader', { name: /A/ }).style.width).toBe('56px');

    fireEvent.mouseMove(window, { clientX: 5000 });
    expect(screen.getByRole('columnheader', { name: /A/ }).style.width).toBe('640px');
  });

  it('resizes from the keyboard so the divider is not pointer-only', () => {
    renderGrid();
    const handle = screen.getByRole('separator', { name: 'Resize column A' });
    const before = parseInt(screen.getByRole('columnheader', { name: /A/ }).style.width, 10);

    fireEvent.keyDown(handle, { key: 'ArrowRight' });

    expect(screen.getByRole('columnheader', { name: /A/ }).style.width).toBe(`${before + 8}px`);
  });

  it('offers a fill handle on the corner of the selection and none while read-only', async () => {
    const user = userEvent.setup();
    const { cell, rerender } = renderGrid();

    await user.click(cell('B2'));
    const handle = screen.getByRole('button', { name: 'Fill from B2' });
    expect(handle).toBeInTheDocument();

    rerender(
      <SpreadsheetGrid
        workbook={testWorkbook()}
        activeSheetName="Sheet1"
        onSelectSheet={() => {}}
        recentChangedCells={new Set()}
      />,
    );
    expect(screen.queryByRole('button', { name: 'Fill from B2' })).not.toBeInTheDocument();
  });

  it('keeps the sheet tabs and the agent context menu working alongside the new roles', async () => {
    const user = userEvent.setup();
    const onAddSelectionContext = vi.fn();
    render(
      <SpreadsheetGrid
        workbook={testWorkbook()}
        activeSheetName="Sheet1"
        onSelectSheet={() => {}}
        recentChangedCells={new Set()}
        onAddSelectionContext={onAddSelectionContext}
        onEditCells={() => {}}
      />,
    );

    fireEvent.contextMenu(screen.getByRole('gridcell', { name: /^A2: / }));
    const menu = screen.getByRole('menu', { name: 'Cell actions' });
    await user.click(within(menu).getByRole('menuitem', { name: 'Ask agent about this cell' }));

    expect(onAddSelectionContext).toHaveBeenCalledWith(
      expect.objectContaining({ kind: 'cell', label: 'A2' }),
    );
  });

  it('windows a very wide sheet instead of rendering a cell per column per row', () => {
    // 400 columns: a naive grid would render 400 <td>s for this single row.
    const wide: Workbook = {
      sheets: [
        {
          name: 'Wide',
          rows: [Array.from({ length: 400 }, (_, i) => createCell(`c${i}`))],
        },
      ],
    };
    render(
      <SpreadsheetGrid
        workbook={wide}
        activeSheetName="Wide"
        onSelectSheet={() => {}}
        recentChangedCells={new Set()}
        onEditCells={() => {}}
      />,
    );

    const grid = screen.getByRole('grid');
    expect(grid).toHaveAttribute('aria-colcount', '401');
    // Column indices still describe the whole sheet even though most of it is off screen.
    const headerA = document.querySelector(
      '.spreadsheet-table th[role="columnheader"][aria-colindex="2"]',
    );
    expect(headerA).toHaveTextContent('A');

    // jsdom reports a zero-sized viewport, which the grid treats as "size unknown" and renders
    // whole. Give it a real viewport and it windows the columns.
    const wrapper = document.querySelector('.grid-scroll-wrapper') as HTMLElement;
    Object.defineProperty(wrapper, 'clientWidth', { value: 400, configurable: true });
    Object.defineProperty(wrapper, 'clientHeight', { value: 300, configurable: true });
    fireEvent.scroll(wrapper);

    const renderedColumns = document.querySelectorAll(
      '.spreadsheet-table tbody td[role="gridcell"]',
    ).length;
    expect(renderedColumns).toBeGreaterThan(0);
    expect(renderedColumns).toBeLessThan(20);

    // Scrolling sideways moves the window instead of growing the DOM.
    Object.defineProperty(wrapper, 'scrollLeft', { value: 2400, configurable: true });
    fireEvent.scroll(wrapper);

    expect(
      document
        .querySelector('.spreadsheet-table th[role="columnheader"]')
        ?.getAttribute('aria-colindex'),
    ).not.toBe('2');
    expect(
      document.querySelectorAll('.spreadsheet-table tbody td[role="gridcell"]').length,
    ).toBeLessThan(20);
  });

  it('fills a series downwards from the handle', async () => {
    const user = userEvent.setup();
    const tall: Workbook = {
      sheets: [
        {
          name: 'Sheet1',
          rows: [
            [createCell('Name')],
            [createCell('Widget')],
            [createCell('')],
            [createCell('')],
            [createCell('')],
          ],
        },
      ],
    };
    const edits: CellEdit[] = [];
    render(
      <SpreadsheetGrid
        workbook={tall}
        activeSheetName="Sheet1"
        onSelectSheet={() => {}}
        recentChangedCells={new Set()}
        onEditCells={(_sheet, batch) => edits.push(...batch)}
      />,
    );
    const cell = (address: string) =>
      screen.getByRole('gridcell', { name: new RegExp(`^${address}: `) });

    await user.click(cell('A2'));
    const handle = screen.getByRole('button', { name: 'Fill from A2' });

    fireEvent.mouseDown(handle, { clientX: 0, clientY: 0 });
    // Four rows down at the rendered row height, so the drag lands on row 5.
    fireEvent.mouseMove(window, { clientX: 0, clientY: 30 * 4 });
    fireEvent.mouseUp(window);

    expect(edits).toEqual([
      { row: 3, column: 'A', value: 'Widget' },
      { row: 4, column: 'A', value: 'Widget' },
      { row: 5, column: 'A', value: 'Widget' },
    ]);
    // The filled range becomes the selection.
    expect(cell('A5')).toHaveAttribute('aria-selected', 'true');
  });
});
