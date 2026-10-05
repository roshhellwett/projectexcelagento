// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { act, fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { createCell, type Workbook } from '@excel-agent/engine';
import { App } from '../src/App.js';
import { askExcelAgent, type AgentResponse } from '../src/lib/llm-service.js';
import { workbookToXlsxBuffer } from '../src/lib/engine-adapter.js';
import type { WorkspaceCheckpoint, WorkspaceRecoveryStore } from '../src/lib/workspace-recovery.js';
import {
  askAgent,
  enterDemoMode,
  fileInput,
  installBrowserStubs,
  metaPillText,
  resetAppState,
} from './helpers.js';

vi.mock('../src/lib/llm-service.js', () => ({ askExcelAgent: vi.fn() }));
installBrowserStubs();

const actionResponse: AgentResponse = {
  message: 'Review the duplicate removal.',
  proposedAction: {
    name: 'delete_duplicates',
    category: 'structure',
    explanation: 'Remove duplicate rows.',
    args: { sheet: 'Orders & Deliveries', columns: ['A'], headerRow: 1 },
  },
};
const planResponse: AgentResponse = {
  message: 'Review the cleaning plan.',
  plan: {
    id: 'clean',
    title: 'Clean orders',
    description: 'Trim and remove duplicates',
    status: 'pending',
    steps: [
      {
        id: 'trim',
        operation: 'normalize_text',
        args: { sheet: 'Orders & Deliveries', columns: ['B'], trim: true },
        description: 'Trim text',
      },
      {
        id: 'dedupe',
        operation: 'delete_duplicates',
        args: { sheet: 'Orders & Deliveries', columns: ['A'], headerRow: 1 },
        description: 'Remove duplicates',
      },
    ],
  },
};

function memoryRecovery(initial: WorkspaceCheckpoint | null = null) {
  let persisted = structuredClone(initial);
  const store: WorkspaceRecoveryStore = {
    load: vi.fn(async () => structuredClone(persisted)),
    save: vi.fn(async (checkpoint) => {
      persisted = structuredClone(checkpoint);
    }),
    clear: vi.fn(async () => {
      persisted = null;
    }),
  };
  return { store, snapshot: () => structuredClone(persisted) };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

const cell = (address: string) =>
  screen.getByRole('gridcell', { name: new RegExp(`^${address}: `) });
const status = () => screen.getByRole('status', { name: 'Local checkpoint status' });
const saved = () =>
  waitFor(() => expect(status()).toHaveTextContent('Checkpoint saved locally'), { timeout: 2500 });
const sampleFixture = () =>
  fireEvent.change(screen.getByRole('combobox', { name: 'Load test fixtures' }), {
    target: { value: 'sample' },
  });

async function editCell(user: ReturnType<typeof userEvent.setup>, address: string, value: string) {
  await user.click(cell(address));
  fireEvent.keyDown(window, { key: value[0] });
  if (value.length > 1) await user.type(screen.getByLabelText('Cell editor'), value.slice(1));
  fireEvent.keyDown(screen.getByLabelText('Cell editor'), { key: 'Enter' });
}

beforeEach(() => {
  resetAppState();
  vi.mocked(askExcelAgent).mockReset().mockResolvedValue(actionResponse);
});

describe('local workbook recovery', () => {
  it('restores an edited upload after refresh, preserving Dates, filename and date epoch with fresh history', async () => {
    const recovery = memoryRecovery();
    const user = userEvent.setup();
    const uploaded: Workbook = {
      dateSystem: '1904',
      sheets: [
        {
          name: 'Dates',
          rows: [
            [createCell('Date'), createCell('Region')],
            [
              { ...createCell(new Date('2024-02-29T00:00:00Z')), numberFormat: 'yyyy-mm-dd' },
              createCell('Europe'),
            ],
          ],
        },
      ],
    };
    const bytes = await workbookToXlsxBuffer(uploaded);
    const first = render(<App recoveryStore={recovery.store} />);
    await user.upload(
      fileInput(first.container),
      new File([new Uint8Array(bytes)], 'dates-1904.xlsx'),
    );
    await waitFor(() => expect(metaPillText(first.container)).toContain('dates-1904.xlsx'));
    await editCell(user, 'B2', 'AP');
    expect(status()).toHaveTextContent('Saving checkpoint');
    await saved();
    const stored = recovery.snapshot();
    expect(stored?.fileName).toBe('dates-1904.xlsx');
    expect(stored?.dateSystem).toBe('1904');
    expect(stored?.workbook.dateSystem).toBe('1904');
    expect(stored?.workbook.sheets[0]?.rows[1]?.[0]?.value).toBeInstanceOf(Date);
    expect(stored?.workbook.sheets[0]?.rows[1]?.[1]?.value).toBe('AP');
    first.unmount();
    vi.mocked(recovery.store.save).mockClear();

    const refreshed = render(<App recoveryStore={recovery.store} />);
    await screen.findByRole('heading', { name: 'Resume your last successful checkpoint' });
    expect(metaPillText(refreshed.container)).toContain('sample-orders.xlsx');
    // Opening the sample must not overwrite the user's checkpoint while they decide.
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 500));
    });
    expect(recovery.store.save).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: 'Restore checkpoint' }));
    expect(metaPillText(refreshed.container)).toContain('dates-1904.xlsx');
    expect(cell('B2')).toHaveTextContent('AP');
    expect(cell('A2')).toHaveTextContent('2024-02-29');
    expect(screen.getByRole('button', { name: 'Undo' })).toBeDisabled();
    expect(screen.getByRole('button', { name: 'History (0/0)' })).toBeInTheDocument();
    expect(
      screen.getByText(/Undo history starts fresh; previous conversations/),
    ).toBeInTheDocument();
    expect(localStorage.getItem('excelagento-workspace')).toBeNull();
  });

  it('does not claim a failed save succeeded and restores only the last successful checkpoint', async () => {
    const recovery = memoryRecovery();
    const user = userEvent.setup();
    const first = render(<App recoveryStore={recovery.store} />);
    await editCell(user, 'D3', 'FR');
    await saved();
    vi.mocked(recovery.store.save).mockRejectedValueOnce(new Error('Storage quota exceeded'));
    await editCell(user, 'D3', 'AP');
    expect(status()).toHaveTextContent('Saving checkpoint');
    const warning = await screen.findByRole(
      'alert',
      { name: 'Checkpoint storage warning' },
      { timeout: 2500 },
    );
    expect(status()).toHaveTextContent('Checkpoint unavailable');
    expect(warning).toHaveTextContent('Storage quota exceeded');
    expect(warning).toHaveTextContent('Only the last successful checkpoint');
    expect(cell('D3')).toHaveTextContent('AP');
    expect(recovery.snapshot()?.workbook.sheets[0]?.rows[2]?.[3]?.value).toBe('FR');
    first.unmount();
    render(<App recoveryStore={recovery.store} />);
    await user.click(await screen.findByRole('button', { name: 'Restore checkpoint' }));
    expect(cell('D3')).toHaveTextContent('FR');
  });

  it('allows edits when storage cannot be read, without overwriting an unread checkpoint', async () => {
    const recovery = memoryRecovery();
    vi.mocked(recovery.store.load).mockRejectedValueOnce(new Error('Storage blocked'));
    const user = userEvent.setup();
    render(<App recoveryStore={recovery.store} />);
    await screen.findByRole('alert', { name: 'Checkpoint storage warning' });
    await editCell(user, 'D3', 'FR');
    expect(cell('D3')).toHaveTextContent('FR');
    expect(status()).toHaveTextContent('unavailable');
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 500));
    });
    expect(recovery.store.save).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: 'Retry local storage' }));
    await saved();
    expect(recovery.snapshot()?.workbook.sheets[0]?.rows[2]?.[3]?.value).toBe('FR');
  });

  it('explicitly deletes the checkpoint and keeps autosave off so it does not silently reappear', async () => {
    const recovery = memoryRecovery();
    const user = userEvent.setup();
    const first = render(<App recoveryStore={recovery.store} />);
    await editCell(user, 'D3', 'FR');
    await saved();
    await user.click(screen.getByRole('button', { name: 'Clear checkpoint' }));
    const dialog = screen.getByRole('dialog', { name: 'Delete local checkpoint?' });
    expect(within(dialog).getByRole('button', { name: 'Keep working' })).toHaveFocus();
    await user.click(within(dialog).getByRole('button', { name: 'Delete checkpoint' }));
    await waitFor(() => expect(status()).toHaveTextContent('Checkpoints off'));
    expect(recovery.snapshot()).toBeNull();
    await editCell(user, 'D3', 'AP');
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 500));
    });
    expect(recovery.store.save).toHaveBeenCalledTimes(1);
    first.unmount();
    render(<App recoveryStore={recovery.store} />);
    await waitFor(() => expect(status()).toHaveTextContent('Not checkpointed yet'));
    expect(screen.queryByRole('button', { name: 'Restore checkpoint' })).not.toBeInTheDocument();
  });

  it('does not report deletion when clearing storage fails', async () => {
    const recovery = memoryRecovery();
    const user = userEvent.setup();
    render(<App recoveryStore={recovery.store} />);
    await editCell(user, 'D3', 'FR');
    await saved();
    vi.mocked(recovery.store.clear).mockRejectedValueOnce(new Error('Blocked'));
    await user.click(screen.getByRole('button', { name: 'Clear checkpoint' }));
    await user.click(screen.getByRole('button', { name: 'Delete checkpoint' }));
    expect(
      await screen.findByRole('alert', { name: 'Checkpoint storage warning' }),
    ).toHaveTextContent('Stored workbook data may remain');
    expect(status()).toHaveTextContent('unavailable');
    expect(recovery.snapshot()).not.toBeNull();
    expect(cell('D3')).toHaveTextContent('FR');
  });
});

describe('workbook-bound turns, previews and confirmations', () => {
  it.each([
    ['action', actionResponse, 'Apply Changes'],
    ['plan', planResponse, 'Apply All 2 Steps'],
  ] as const)(
    'clears the old %s card and selection when a clean workbook is replaced',
    async (_kind, response, applyLabel) => {
      enterDemoMode();
      vi.mocked(askExcelAgent).mockResolvedValue(response);
      const user = userEvent.setup();
      const { container } = render(<App recoveryStore={memoryRecovery().store} />);
      fireEvent.contextMenu(cell('D3'), { clientX: 20, clientY: 20 });
      await user.click(screen.getByRole('menuitem', { name: 'Ask agent about this cell' }));
      expect(screen.getByRole('button', { name: 'Clear selection' })).toBeInTheDocument();
      await askAgent(user, 'old workbook request');
      await screen.findByRole('button', { name: applyLabel });
      sampleFixture();
      expect(metaPillText(container)).toContain('sample-orders.xlsx');
      expect(screen.queryByRole('button', { name: applyLabel })).not.toBeInTheDocument();
      expect(screen.queryByRole('button', { name: 'Clear selection' })).not.toBeInTheDocument();
      expect(screen.queryByText('old workbook request')).not.toBeInTheDocument();
      expect(container.querySelector('.cell-coord-badge')).toHaveTextContent('A1');
      await askAgent(user, 'new workbook request');
      const options = vi.mocked(askExcelAgent).mock.calls.at(-1)?.[5];
      expect(
        options?.conversationHistory?.map((message) => message.content).join(' '),
      ).not.toContain('old workbook request');
    },
  );

  it('aborts an old turn and ignores all late callbacks/results without stopping the new turn', async () => {
    enterDemoMode();
    const oldTurn = deferred<AgentResponse>();
    const newTurn = deferred<AgentResponse>();
    vi.mocked(askExcelAgent)
      .mockImplementationOnce(() => oldTurn.promise)
      .mockImplementationOnce(() => newTurn.promise);
    const user = userEvent.setup();
    render(<App recoveryStore={memoryRecovery().store} />);
    await askAgent(user, 'old workbook request');
    const oldOptions = vi.mocked(askExcelAgent).mock.calls[0]?.[5];
    sampleFixture();
    expect(oldOptions?.signal?.aborted).toBe(true);
    await askAgent(user, 'new workbook request');
    await act(async () => {
      oldOptions?.callbacks?.onToken?.('LATE OLD TOKEN');
      oldOptions?.callbacks?.onThinking?.('LATE OLD THOUGHT');
      oldOptions?.callbacks?.onTokenCount?.({ totalTokens: 99999 });
      oldOptions?.onActivity?.({
        id: 'old',
        type: 'status',
        agent: 'Old',
        timestamp: Date.now(),
        summary: 'LATE OLD ACTIVITY',
      });
      oldTurn.resolve({ ...actionResponse, message: 'LATE OLD RESULT' });
      await oldTurn.promise;
    });
    expect(screen.queryByText(/LATE OLD/)).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: 'Apply Changes' })).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Stop' })).toBeInTheDocument();
    expect(screen.getByPlaceholderText(/Ask ExcelAgento/)).toBeDisabled();
    expect(localStorage.getItem('excel_agent_usage_v1') ?? '').not.toContain(
      'old workbook request',
    );
    await act(async () => {
      newTurn.resolve(actionResponse);
      await newTurn.promise;
    });
    expect(await screen.findByRole('button', { name: 'Apply Changes' })).toBeInTheDocument();
  });

  it.each([
    ['action', actionResponse, 'Apply Changes'],
    ['plan', planResponse, 'Apply All 2 Steps'],
  ] as const)(
    'invalidates an old %s confirmation after an edit and undo; repreview requires new consent',
    async (_kind, response, applyLabel) => {
      enterDemoMode();
      vi.mocked(askExcelAgent).mockResolvedValue(response);
      const user = userEvent.setup();
      const { container } = render(<App recoveryStore={memoryRecovery().store} />);
      await askAgent(user, 'remove duplicate rows');
      await user.click(await screen.findByRole('button', { name: applyLabel }));
      await screen.findByRole('group', { name: 'Confirm destructive change' });
      await editCell(user, 'D3', 'FR');
      await user.click(screen.getByRole('button', { name: 'Undo' }));
      expect(cell('D3')).toHaveTextContent('Europe');
      expect(metaPillText(container)).toContain('11 rows');
      expect(
        screen.queryByRole('button', { name: 'Yes, apply this change' }),
      ).not.toBeInTheDocument();
      expect(screen.queryByRole('button', { name: applyLabel })).not.toBeInTheDocument();
      expect(screen.getByText(/previous confirmation no longer applies/)).toBeInTheDocument();
      await user.click(screen.getByRole('button', { name: 'Preview again' }));
      await user.click(await screen.findByRole('button', { name: applyLabel }));
      expect(metaPillText(container)).toContain('11 rows');
      await user.click(await screen.findByRole('button', { name: 'Yes, apply this change' }));
      await waitFor(() => expect(metaPillText(container)).toContain('10 rows'));
    },
  );

  it('refuses a proposal prepared against a revision that was edited during the provider turn', async () => {
    enterDemoMode();
    const turn = deferred<AgentResponse>();
    vi.mocked(askExcelAgent).mockImplementationOnce(() => turn.promise);
    const user = userEvent.setup();
    const { container } = render(<App recoveryStore={memoryRecovery().store} />);
    await askAgent(user, 'remove duplicate rows');
    await editCell(user, 'D3', 'FR');
    await act(async () => {
      turn.resolve(actionResponse);
      await turn.promise;
    });
    expect(await screen.findByRole('button', { name: 'Preview again' })).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: 'Apply Changes' })).not.toBeInTheDocument();
    expect(metaPillText(container)).toContain('11 rows');
    expect(cell('D3')).toHaveTextContent('FR');
  });

  it('requires explicit in-app confirmation for dirty fixture changes and file uploads', async () => {
    const recovery = memoryRecovery();
    const user = userEvent.setup();
    const { container } = render(<App recoveryStore={recovery.store} />);
    await editCell(user, 'D3', 'FR');
    sampleFixture();
    const fixtureDialog = screen.getByRole('dialog', {
      name: 'Replace workbook with unexported changes?',
    });
    expect(fixtureDialog).toHaveAttribute('aria-modal', 'true');
    fireEvent.keyDown(document, { key: 'Escape' });
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument();
    expect(cell('D3')).toHaveTextContent('FR');

    await user.upload(
      fileInput(container),
      new File(['Region,Total\nEMEA,500'], 'quarterly.csv', { type: 'text/csv' }),
    );
    const uploadDialog = await screen.findByRole('dialog', {
      name: 'Replace workbook with unexported changes?',
    });
    expect(uploadDialog).toHaveTextContent('quarterly.csv');
    expect(metaPillText(container)).toContain('sample-orders.xlsx');
    expect(cell('D3')).toHaveTextContent('FR');
    await user.click(within(uploadDialog).getByRole('button', { name: 'Replace workbook' }));
    expect(metaPillText(container)).toContain('quarterly.csv');
    expect(screen.getByRole('button', { name: 'Undo' })).toBeDisabled();
  });
});
