// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { askExcelAgent } from '../src/lib/llm-service.js';
import {
  askAgent,
  enterDemoMode,
  installBrowserStubs,
  metaPillText,
  renderApp,
  resetAppState,
} from './helpers.js';

vi.mock('../src/lib/llm-service.js', () => ({ askExcelAgent: vi.fn() }));
installBrowserStubs();

beforeEach(() => {
  resetAppState();
  enterDemoMode();
  vi.mocked(askExcelAgent).mockResolvedValue({
    message: 'Review this plan.',
    plan: {
      id: 'plan-1',
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
  });
});

describe('plan execution in the workspace', () => {
  it('offers a confirmation gate and undoes the whole plan with one click', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();
    await askAgent(user, 'clean these orders');
    await user.click(
      await screen.findByRole('button', { name: 'Apply All 2 Steps' }, { timeout: 5000 }),
    );
    expect(
      await screen.findByRole('group', { name: 'Confirm destructive change' }),
    ).toBeInTheDocument();
    expect(metaPillText(container)).toContain('11 rows');
    await user.click(screen.getByRole('button', { name: 'Yes, apply this change' }));
    await waitFor(() => expect(metaPillText(container)).toContain('10 rows'));
    await user.click(screen.getByRole('button', { name: /Undo this step/i }));
    await waitFor(() => expect(metaPillText(container)).toContain('11 rows'));
  });

  it('leaves all rows intact when plan confirmation is cancelled', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();
    await askAgent(user, 'clean these orders');
    await user.click(
      await screen.findByRole('button', { name: 'Apply All 2 Steps' }, { timeout: 5000 }),
    );
    await user.click(await screen.findByRole('button', { name: 'Cancel' }));
    expect(metaPillText(container)).toContain('11 rows');
    expect(localStorage.getItem('excel_agent_memory_v1')).toBeNull();
  });
});
