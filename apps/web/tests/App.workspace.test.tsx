// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { createCell, type Workbook } from '@excel-agent/engine';

import { workbookToXlsxBuffer } from '../src/lib/engine-adapter.js';
import {
  askAgent,
  createdBlobsSnapshot,
  enterDemoMode,
  fileInput,
  fileWithSize,
  installBrowserStubs,
  metaPillText,
  renderApp,
  resetAppState,
  revokedUrlsSnapshot,
  sheetTabTexts,
  toastTexts,
} from './helpers.js';

installBrowserStubs();

beforeEach(() => {
  resetAppState();
});

describe('workspace shell', () => {
  it('renders the sample workbook with the BYOK gate closed', () => {
    const { container } = renderApp();

    expect(metaPillText(container)).toContain('sample-orders.xlsx');
    expect(metaPillText(container)).toContain('11 rows');
    expect(screen.getAllByText('Order ID').length).toBeGreaterThan(0);
    expect(
      screen.getByRole('heading', { name: 'Get started with ExcelAgento' }),
    ).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Try the local agent/i })).toBeInTheDocument();
  });

  it('unlocks demo mode, greets the user, and stores the preference', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    await user.click(screen.getByRole('button', { name: /Try the local agent/i }));

    expect(await screen.findByPlaceholderText(/Ask ExcelAgento/i)).toBeInTheDocument();
    expect(screen.getByText('LOCAL ENGINE · NO KEY')).toBeInTheDocument();
    expect(
      screen.queryByRole('heading', { name: 'Get started with ExcelAgento' }),
    ).not.toBeInTheDocument();
    expect(localStorage.getItem('excel_agent_settings_v2')).toBeNull();
    await waitFor(() => expect(container.textContent).toContain('local agent is ready'), {
      timeout: 5000,
    });
  });

  it('opens the command palette with Ctrl+K', async () => {
    renderApp();

    fireEvent.keyDown(window, { key: 'k', ctrlKey: true });

    expect(await screen.findByPlaceholderText(/Type a command/i)).toBeInTheDocument();
  });
});

describe('deterministic actions from chat', () => {
  /**
   * `delete_duplicates` removes rows irreversibly from the user's point of view, so the engine
   * refuses it until a human confirms. Clicking "Apply Changes" therefore opens a confirm gate
   * instead of mutating anything, and the second, deliberate click is what applies it.
   */
  const applyWithConfirmation = async (user: ReturnType<typeof userEvent.setup>) => {
    const apply = await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });
    await user.click(apply);
    const confirm = await screen.findByRole('button', { name: /Yes, apply this change/i });
    await user.click(confirm);
  };

  it('demands confirmation before deleting rows, then applies and undoes', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    const { container } = renderApp();

    // The sample workbook deliberately contains one duplicate row.
    await askAgent(user, 'remove duplicate rows');
    expect(metaPillText(container)).toContain('11 rows');

    const apply = await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });
    await user.click(apply);

    // The first click must not touch the data.
    const gate = await screen.findByRole('group', { name: /Confirm destructive change/i });
    expect(gate).toBeInTheDocument();
    expect(metaPillText(container)).toContain('11 rows');
    expect(toastTexts(container).join(' ')).toMatch(/was not confirmed/i);

    await user.click(screen.getByRole('button', { name: /Yes, apply this change/i }));

    await waitFor(() =>
      expect(toastTexts(container).join(' ')).toMatch(/delete_duplicates applied/i),
    );
    await waitFor(() => expect(metaPillText(container)).toContain('10 rows'));

    const undo = await screen.findByRole('button', { name: /Undo This Step/i });
    await user.click(undo);

    await waitFor(() => expect(metaPillText(container)).toContain('11 rows'));
  });

  it('applies nothing when the confirmation is cancelled', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    const { container } = renderApp();

    await askAgent(user, 'remove duplicate rows');
    const apply = await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });
    await user.click(apply);
    await screen.findByRole('group', { name: /Confirm destructive change/i });

    await user.click(screen.getByRole('button', { name: /Cancel/i }));

    await waitFor(() => expect(metaPillText(container)).toContain('11 rows'));
    expect(toastTexts(container).join(' ')).not.toMatch(/delete_duplicates applied/i);
  });

  it('persists a verified action into the self-learning memory only once confirmed', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'remove duplicate rows');
    await applyWithConfirmation(user);

    await waitFor(() =>
      expect(localStorage.getItem('excel_agent_memory_v1') ?? '').toContain('delete_duplicates'),
    );
  });

  it('does not learn from a change that is only awaiting confirmation', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'remove duplicate rows');
    const apply = await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });
    await user.click(apply);
    await screen.findByRole('group', { name: /Confirm destructive change/i });

    // A pause for a human decision is not a rejection, so nothing may be recorded either way.
    const stored = localStorage.getItem('excel_agent_memory_v1') ?? '';
    expect(stored).not.toContain('delete_duplicates');
  });

  it('runs a no-key local turn without contacting an AI provider', async () => {
    const fetchSpy = vi.fn(async () => {
      throw new Error('the network must not be used in demo mode');
    });
    vi.stubGlobal('fetch', fetchSpy);
    const user = userEvent.setup();
    renderApp();

    await user.click(screen.getByRole('button', { name: /Try the local agent/i }));
    await askAgent(user, 'remove duplicate rows');
    await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });

    expect(fetchSpy).not.toHaveBeenCalled();
    expect(localStorage.getItem('excel_agent_settings_v2')).toBeNull();
    vi.unstubAllGlobals();
  });
});

describe('file handling', () => {
  it('rejects a file larger than the 50 MB upload guard', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    await user.upload(fileInput(container), fileWithSize('huge.xlsx', 51 * 1024 * 1024));

    await waitFor(() => expect(toastTexts(container).join(' ')).toMatch(/larger than 50 MB/i));
    // The workspace keeps the current workbook rather than half-loading it.
    expect(metaPillText(container)).toContain('sample-orders.xlsx');
  });

  it('imports a real .xlsx file selected from disk', async () => {
    const uploaded: Workbook = {
      sheets: [
        {
          name: 'Revenue',
          rows: [
            [createCell('Region'), createCell('Total')],
            [createCell('EMEA'), createCell(500)],
          ],
        },
      ],
    };
    const bytes = await workbookToXlsxBuffer(uploaded);
    // Copy into a fresh, ArrayBuffer-backed view so it satisfies BlobPart under TS 6.
    const file = new File([new Uint8Array(bytes)], 'quarterly.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const user = userEvent.setup();
    const { container } = renderApp();
    await user.upload(fileInput(container), file);

    await waitFor(() => expect(metaPillText(container)).toContain('quarterly.xlsx'));
    await waitFor(() =>
      expect(toastTexts(container).join(' ')).toMatch(/Loaded "quarterly\.xlsx"/i),
    );
    expect(metaPillText(container)).toContain('2 rows');
    expect(container.textContent).toContain('Region');
    expect(sheetTabTexts(container).join(' ')).toContain('Revenue');
  });

  it('exports the workbook as a real xlsx blob', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    await user.click(screen.getByRole('button', { name: /Export \.xlsx/i }));

    await waitFor(() => expect(toastTexts(container).join(' ')).toMatch(/Exported/i));
    const blobs = createdBlobsSnapshot();
    expect(blobs).toHaveLength(1);
    expect(blobs[0]?.type).toContain('spreadsheetml.sheet');
    expect(blobs[0]?.size).toBeGreaterThan(1000);
    expect(revokedUrlsSnapshot()).toHaveLength(1);
  });
});

describe('BYOK settings modal', () => {
  it('saves a key, reflects it in the chat header, then clears it', async () => {
    const user = userEvent.setup();
    renderApp();

    await user.click(screen.getByRole('button', { name: /API Keys & Settings/i }));
    const card = document.querySelector('.modal-card') as HTMLElement;
    expect(card).not.toBeNull();
    expect(within(card).getByText(/Settings - Model Keys/i)).toBeInTheDocument();

    await user.type(within(card).getByPlaceholderText(/sk-or/i), 'sk-or-test-123456');
    await user.click(within(card).getByRole('button', { name: /Save Preferences/i }));

    await waitFor(() =>
      expect(localStorage.getItem('excel_agent_settings_v2')).toContain('sk-or-test-123456'),
    );
    expect(await screen.findByText('OPENROUTER ACTIVE')).toBeInTheDocument();
    // The modal auto-closes after a successful save.
    await waitFor(() => expect(document.querySelector('.modal-card')).toBeNull(), {
      timeout: 3000,
    });

    await user.click(screen.getByRole('button', { name: /Change Key/i }));
    await waitFor(() =>
      expect(
        screen.getByRole('heading', { name: 'Get started with ExcelAgento' }),
      ).toBeInTheDocument(),
    );
    // Clearing removes the stored configuration entirely, it does not leave a key behind.
    expect(localStorage.getItem('excel_agent_settings_v2')).toBeNull();
  });

  it('translates a raw fetch failure into an actionable message', async () => {
    const user = userEvent.setup();
    renderApp();

    await user.click(screen.getByRole('button', { name: /API Keys & Settings/i }));
    const card = document.querySelector('.modal-card') as HTMLElement;

    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        throw new TypeError('Failed to fetch');
      }),
    );

    await user.type(within(card).getByPlaceholderText(/sk-or/i), 'sk-or-test-123456');
    await user.click(within(card).getByRole('button', { name: /Test Connection/i }));

    expect(await within(card).findByText(/Could not reach the provider/i)).toBeInTheDocument();
    vi.unstubAllGlobals();
  });
});
