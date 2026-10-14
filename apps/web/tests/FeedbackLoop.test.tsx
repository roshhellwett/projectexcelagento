// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { askExcelAgent, type AgentResponse } from '../src/lib/llm-service.js';
import { clearFeedbackLogWithStatus, FEEDBACK_KEY } from '../src/lib/feedback.js';
import {
  askAgent,
  createdBlobsSnapshot,
  enterDemoMode,
  installBrowserStubs,
  renderApp,
  resetAppState,
} from './helpers.js';

vi.mock('../src/lib/llm-service.js', () => ({ askExcelAgent: vi.fn() }));
installBrowserStubs();

const successResponse: AgentResponse = {
  message: 'Review duplicate removal.',
  proposedAction: {
    name: 'delete_duplicates',
    category: 'structure',
    explanation: 'Remove duplicate rows',
    args: { sheet: 'Orders & Deliveries', columns: ['A'], headerRow: 1, keep: 'first' },
  },
};

const blockedResponse: AgentResponse = {
  message: 'I could not safely prepare that request.',
  guardrail: {
    passed: false,
    schemaValid: true,
    engineValid: false,
    errors: ['The requested workbook operation is not supported for this sheet.'],
    warnings: [],
    requiresConfirmation: false,
  },
};

beforeEach(() => {
  resetAppState();
  clearFeedbackLogWithStatus();
  window.location.hash = '';
  vi.mocked(askExcelAgent).mockReset();
});

describe('design-partner feedback loop', () => {
  it('offers local feedback after a completed task and surfaces its local summary', async () => {
    enterDemoMode();
    vi.mocked(askExcelAgent).mockResolvedValue(successResponse);
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'remove duplicate rows');
    await user.click(await screen.findByRole('button', { name: /Apply Changes/i }));
    await user.click(await screen.findByRole('button', { name: /Yes, apply this change/i }));

    const feedback = await screen.findByTestId('task-feedback');
    await user.click(within(feedback).getByRole('button', { name: 'Worked well' }));
    await user.selectOptions(within(feedback).getByLabelText('Workflow area'), 'cleanup');
    await user.type(within(feedback).getByLabelText('Optional note'), 'The preview was clear.');
    await user.click(within(feedback).getByRole('button', { name: 'Save local feedback' }));

    expect(await screen.findByText(/feedback saved locally/i)).toBeInTheDocument();
    const stored = JSON.parse(localStorage.getItem(FEEDBACK_KEY) ?? '[]') as Record<
      string,
      unknown
    >[];
    expect(stored).toHaveLength(1);
    expect(stored[0]).toMatchObject({
      result: 'completed',
      rating: 'helpful',
      category: 'cleanup',
      note: 'The preview was clear.',
    });
    expect(stored[0]).not.toHaveProperty('request');
    expect(stored[0]).not.toHaveProperty('fileName');

    await user.click(screen.getByRole('button', { name: 'Token Ledger' }));
    expect(
      await screen.findByRole('region', { name: 'Local design-partner feedback' }),
    ).toBeInTheDocument();
    expect(screen.getByTestId('feedback-count')).toHaveTextContent('1');

    await user.click(screen.getByRole('button', { name: 'Download aggregate feedback summary' }));
    const blob = createdBlobsSnapshot().at(-1);
    expect(blob).toBeDefined();
    const exported = JSON.parse(await blob!.text()) as {
      summary: { total: number };
      entries: Record<string, unknown>[];
    };
    expect(exported.summary.total).toBe(1);
    expect(exported.entries[0]).toMatchObject({
      result: 'completed',
      rating: 'helpful',
      category: 'cleanup',
    });
    expect(exported.entries[0]).not.toHaveProperty('note');
    expect(await blob!.text()).not.toContain('The preview was clear.');
  });

  it('offers the same local signal after a blocked or unsupported task', async () => {
    enterDemoMode();
    vi.mocked(askExcelAgent).mockResolvedValue(blockedResponse);
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'run an unsupported workbook operation');

    expect(await screen.findByRole('alert', { name: 'Safety boundary' })).toBeInTheDocument();
    const feedback = await screen.findByTestId('task-feedback');
    await user.click(within(feedback).getByRole('button', { name: 'I needed this task' }));
    await user.click(within(feedback).getByRole('button', { name: 'Save local feedback' }));

    await waitFor(() => {
      const stored = JSON.parse(localStorage.getItem(FEEDBACK_KEY) ?? '[]') as Record<
        string,
        unknown
      >[];
      expect(stored[0]).toMatchObject({
        result: 'blocked',
        rating: 'not_helpful',
      });
    });
    expect(
      screen.queryByRole('button', { name: /Apply Changes|Apply All/i }),
    ).not.toBeInTheDocument();
  });

  it('shows session-only durability when browser feedback writes are denied', async () => {
    enterDemoMode();
    const originalSetItem = Storage.prototype.setItem;
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(function (
      this: Storage,
      key: string,
      value: string,
    ) {
      if (key === FEEDBACK_KEY) throw new Error('feedback writes denied');
      return originalSetItem.call(this, key, value);
    });
    vi.mocked(askExcelAgent).mockResolvedValue(blockedResponse);
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'run an unsupported workbook operation');
    const feedback = await screen.findByTestId('task-feedback');
    await user.click(within(feedback).getByRole('button', { name: 'I needed this task' }));
    await user.click(within(feedback).getByRole('button', { name: 'Save local feedback' }));

    expect(await screen.findByText(/feedback kept for this session only/i)).toBeInTheDocument();
    expect(screen.getByText(/browser storage was unavailable/i)).toBeInTheDocument();
  });

  it('keeps feedback visible and reports failed clearing', async () => {
    enterDemoMode();
    vi.mocked(askExcelAgent).mockResolvedValue(blockedResponse);
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'run an unsupported workbook operation');
    const feedback = await screen.findByTestId('task-feedback');
    await user.click(within(feedback).getByRole('button', { name: 'I needed this task' }));
    await user.click(within(feedback).getByRole('button', { name: 'Save local feedback' }));
    await user.click(screen.getByRole('button', { name: 'Token Ledger' }));
    expect(await screen.findByTestId('feedback-count')).toHaveTextContent('1');

    vi.spyOn(Storage.prototype, 'removeItem').mockImplementation(() => {
      throw new Error('feedback removal denied');
    });
    await user.click(screen.getByTestId('feedback-clear'));

    expect(screen.getByTestId('feedback-count')).toHaveTextContent('1');
    expect(screen.getByRole('alert')).toHaveTextContent(/no deletion is claimed/i);
  });
});
