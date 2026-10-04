// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import {
  enterDemoMode,
  installBrowserStubs,
  metaPillText,
  renderApp,
  resetAppState,
} from './helpers.js';

installBrowserStubs();

beforeEach(() => {
  resetAppState();
});

const cell = (address: string) =>
  screen.getByRole('gridcell', { name: new RegExp(`^${address}: `) });

describe('editing cells in the workspace', () => {
  it('types into a cell, changes the model, and puts it back on undo', async () => {
    const user = userEvent.setup();
    renderApp();

    // D3 is "Europe" in the sample workbook.
    expect(cell('D3')).toHaveTextContent('Europe');

    await user.click(cell('D3'));
    fireEvent.keyDown(window, { key: 'A' });
    const editor = screen.getByLabelText('Cell editor');
    expect(editor).toHaveValue('A');
    await user.type(editor, 'P');
    fireEvent.keyDown(editor, { key: 'Enter' });

    expect(cell('D3')).toHaveTextContent('AP');

    await user.click(screen.getByRole('button', { name: /^Undo$/ }));
    expect(cell('D3')).toHaveTextContent('Europe');

    // A cell edit must not raise a toast per keystroke, but it is a real history entry.
    expect(document.querySelectorAll('.toast')).toHaveLength(0);
  });

  it('shows the committed value in the grid, the formula bar and the history log', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    await user.click(cell('D3'));
    fireEvent.keyDown(window, { key: 'F' });
    await user.type(screen.getByLabelText('Cell editor'), 'R');
    fireEvent.keyDown(screen.getByLabelText('Cell editor'), { key: 'Enter' });

    expect(cell('D3')).toHaveTextContent('FR');
    // Enter committed and stepped down, so the badge and the bar now describe the cell below.
    expect(container.querySelector('.cell-coord-badge')).toHaveTextContent('D4');
    await user.click(cell('D3'));
    expect(container.querySelector('.formula-input-display')).toHaveTextContent('FR');

    // The change is in the audit log like any other, and undo is offered for it.
    await user.click(screen.getByRole('button', { name: /History \(/ }));
    expect(screen.getByRole('dialog', { name: 'Operation audit log' })).toHaveTextContent(
      'edit_cells',
    );
    expect(screen.getByRole('button', { name: /Undo$/ })).toBeEnabled();
    expect(metaPillText(container)).toContain('Orders & Deliveries');
  });

  it('clears a range with Delete and undoes the whole range at once', async () => {
    const user = userEvent.setup();
    renderApp();

    await user.click(cell('A2'));
    fireEvent.keyDown(window, { key: 'ArrowRight', shiftKey: true });
    fireEvent.keyDown(window, { key: 'Delete' });

    expect(cell('A2')).toHaveTextContent('');
    expect(cell('B2')).toHaveTextContent('');
    expect(cell('C2')).toHaveTextContent('2026-03-01');

    await user.click(screen.getByRole('button', { name: /^Undo$/ }));

    expect(cell('A2')).toHaveTextContent('ORD-1001');
    expect(cell('B2')).toHaveTextContent('john DOE');
  });

  it('pastes a tab-separated block into the sheet, growing it past the current data', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    // 11 rows of 6 columns today; this block needs two more columns.
    expect(metaPillText(container)).toContain('11 rows');

    await user.click(cell('F11'));
    fireEvent.paste(window, { clipboardData: { getData: () => 'Region\tNotes' } });
    expect(cell('F11')).toHaveTextContent('Region');
    expect(screen.getByRole('columnheader', { name: /G/ })).toBeInTheDocument();
    expect(cell('G11')).toHaveTextContent('Notes');
    expect(metaPillText(container)).toContain('7 cols');

    await user.click(screen.getByRole('button', { name: /^Undo$/ }));

    // Undo takes back both the overwritten cell and the column the paste added.
    expect(cell('F11')).toHaveTextContent('Pending');
    expect(screen.queryByRole('columnheader', { name: /G/ })).not.toBeInTheDocument();
    expect(metaPillText(container)).toContain('6 cols');
  });

  it('leaves the chat usable while demo mode is off', async () => {
    const user = userEvent.setup();
    renderApp();

    await user.click(cell('B2'));
    fireEvent.keyDown(window, { key: '9' });
    await user.keyboard('{Enter}');

    // 1,420.5 became 9, not a string: the sheet can still total the column.
    expect(cell('B2')).toHaveTextContent('9');
    expect(screen.getByRole('heading', { name: 'Activate Excel Agent' })).toBeInTheDocument();
  });
});

describe('accessible dialogs over a live grid', () => {
  it('moves focus into the settings dialog and restores it on Escape', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();

    const opener = screen.getByRole('button', { name: /API Keys & Settings/i });
    await user.click(opener);

    const dialog = await screen.findByRole('dialog');
    expect(dialog).toHaveAttribute('aria-modal', 'true');
    await waitForFocus(screen.getByLabelText('AI Provider'));

    fireEvent.keyDown(document, { key: 'Escape' });

    await waitForGone(() => screen.queryByRole('dialog'));
    expect(opener).toHaveFocus();
  });

  it('closes the operation modal on Escape and keeps the grid out of the way', async () => {
    const user = userEvent.setup();
    renderApp();

    const opener = screen.getByRole('button', { name: /Run Operation/i });
    await user.click(opener);

    const dialog = await screen.findByRole('dialog', { name: 'Run Engine Operation' });
    // Every label in the dialog now names the control it labels.
    expect(screen.getByLabelText('Select Operation')).toBeInTheDocument();
    expect(screen.getByLabelText('Target Column')).toBeInTheDocument();

    // A keystroke aimed at the grid must not land in the sheet behind the modal.
    fireEvent.keyDown(window, { key: 'x' });
    expect(cell('A1')).toHaveTextContent('Order ID');

    fireEvent.keyDown(document, { key: 'Escape' });
    await waitForGone(() => screen.queryByRole('dialog'));
    expect(dialog).not.toBeInTheDocument();
  });
});

async function waitForFocus(element: HTMLElement): Promise<void> {
  for (let attempt = 0; attempt < 20 && document.activeElement !== element; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  expect(document.activeElement).toBe(element);
}

async function waitForGone(query: () => HTMLElement | null): Promise<void> {
  for (let attempt = 0; attempt < 20 && query() !== null; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  expect(query()).toBeNull();
}
