// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import {
  enterDemoMode,
  askAgent,
  renderApp,
  resetAppState,
  installBrowserStubs,
} from './helpers.js';

installBrowserStubs();
beforeEach(() => resetAppState());

describe('analyst evidence and task receipts', () => {
  it('shows deterministic source evidence for a statistics answer', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();
    await askAgent(user, 'descriptive statistics for column E');
    expect(await screen.findByText(/Evidence & sources/i)).toBeInTheDocument();
    expect(screen.getByText(/Orders & Deliveries!E2:E11/i)).toBeInTheDocument();
  });

  it('shows an applied task receipt and changes it when the task is undone', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();
    await askAgent(user, 'trim whitespace in column B');
    const apply = await screen.findByRole('button', { name: /Apply Changes/i });
    await user.click(apply);
    expect(await screen.findByText(/Task completed/i)).toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: /Undo this step/i }));
    await waitFor(() => expect(screen.getByText(/Task reverted/i)).toBeInTheDocument());
  });
});
