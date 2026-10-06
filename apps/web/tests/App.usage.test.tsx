// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { saveSettings } from '../src/lib/settings.js';
import {
  askAgent,
  enterDemoMode,
  enterLiveMode,
  installBrowserStubs,
  renderApp,
  resetAppState,
} from './helpers.js';

/** A recorded Groq completion carrying real usage numbers. */
function recordedCompletion(): Response {
  return new Response(
    JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      choices: [{ index: 0, message: { role: 'assistant', content: 'I reviewed the sheet.' } }],
      usage: { prompt_tokens: 900, completion_tokens: 80, total_tokens: 980 },
    }),
    { status: 200, headers: { 'content-type': 'application/json' } },
  );
}

installBrowserStubs();

beforeEach(() => {
  resetAppState();
  window.location.hash = '';
});

describe('Model & Usage page', () => {
  it('opens from the workspace and returns to it', async () => {
    const user = userEvent.setup();
    const { container } = renderApp();

    await user.click(screen.getByTestId('open-usage'));

    expect(await screen.findByTestId('usage-page')).toBeInTheDocument();
    expect(screen.getByRole('heading', { name: /Model & Usage/i })).toBeInTheDocument();
    // The workspace is genuinely replaced, not overlaid.
    expect(container.querySelector('.file-meta-pill')).toBeNull();
    expect(window.location.hash).toBe('#/usage');

    await user.click(screen.getByRole('button', { name: /Back to workspace/i }));

    await waitFor(() => expect(container.querySelector('.file-meta-pill')).not.toBeNull());
    expect(window.location.hash).toBe('#/workspace');
    expect(screen.queryByTestId('usage-page')).not.toBeInTheDocument();
  });

  it('renders from a deep link at #/usage', () => {
    window.location.hash = '#/usage';

    renderApp();

    expect(screen.getByTestId('usage-page')).toBeInTheDocument();
  });

  it('reports an unconfigured workspace honestly and shows an empty ledger', async () => {
    const user = userEvent.setup();
    renderApp();
    await user.click(screen.getByTestId('open-usage'));

    expect(screen.getByTestId('usage-provider')).toHaveTextContent('OpenRouter (multi-model)');
    expect(screen.getByTestId('usage-model')).toHaveTextContent('anthropic/claude-3.5-sonnet');
    expect(screen.getByTestId('usage-key-status')).toHaveTextContent('Not set');
    expect(screen.getByTestId('usage-mode')).toHaveTextContent('Local deterministic engine');
    expect(screen.getByTestId('usage-empty')).toBeInTheDocument();
    expect(screen.getByTestId('usage-requests')).toHaveTextContent('0');
  });

  it('marks the key active, shows the model and host, and never exposes the secret', async () => {
    enterLiveMode();
    const user = userEvent.setup();
    renderApp();
    await user.click(screen.getByTestId('open-usage'));

    const page = screen.getByTestId('usage-page');
    expect(screen.getByTestId('usage-key-status')).toHaveTextContent('Active');
    expect(screen.getByTestId('usage-mode')).toHaveTextContent('BYOK model + guardrail');
    expect(page.textContent).toContain('openrouter.ai');
    expect(page.textContent).toContain('sk-o');
    // The full key must never be rendered.
    expect(page.textContent).not.toContain('sk-or-live_test_key');
  });

  it('reflects a custom model override', async () => {
    saveSettings({ provider: 'openrouter', apiKey: 'sk-or-test', model: 'deepseek/deepseek-chat' });

    const user = userEvent.setup();
    renderApp();
    await user.click(screen.getByTestId('open-usage'));

    expect(screen.getByTestId('usage-provider')).toHaveTextContent('OpenRouter (multi-model)');
    expect(screen.getByTestId('usage-model')).toHaveTextContent('deepseek/deepseek-chat');
    expect(screen.getByText('Manual override')).toBeInTheDocument();
  });
});

describe('token accounting', () => {
  it('records a demo-mode turn as a zero-token local request', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'remove duplicate rows');
    await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });

    await user.click(screen.getByTestId('open-usage'));

    expect(screen.getByTestId('usage-requests')).toHaveTextContent('1');
    expect(screen.getByTestId('usage-total-tokens')).toHaveTextContent('0');
    const rows = screen.getAllByTestId('usage-log-row');
    expect(rows).toHaveLength(1);
    expect(rows[0]?.textContent).toContain('remove duplicate rows');
    expect(rows[0]?.textContent).toContain('none');
    expect(rows[0]?.textContent).toContain('deterministic engine');
    expect(rows[0]?.textContent).toContain('OK');
  });

  it('records real provider-reported tokens for a model-backed turn', async () => {
    const fetchSpy = vi.fn(async () => recordedCompletion());
    vi.stubGlobal('fetch', fetchSpy);
    enterLiveMode();

    const user = userEvent.setup();
    renderApp();
    await askAgent(user, 'remove duplicate rows');
    await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    await user.click(screen.getByTestId('open-usage'));

    expect(screen.getByTestId('usage-requests')).toHaveTextContent('1');
    expect(screen.getByTestId('usage-total-tokens')).toHaveTextContent('980');
    expect(screen.getByTestId('usage-prompt-tokens')).toHaveTextContent('900');
    expect(screen.getByTestId('usage-completion-tokens')).toHaveTextContent('80');
    expect(screen.getByTestId('usage-failures')).toHaveTextContent('0');

    const row = screen.getAllByTestId('usage-log-row')[0]!;
    expect(row.textContent).toContain('openrouter');
    expect(row.textContent).toContain('980');
    // Totals are also broken down per provider and per model.
    expect(screen.getByText('Breakdown')).toBeInTheDocument();

    vi.unstubAllGlobals();
  });

  it('records a rejected key as a failed call with its reason', async () => {
    const fetchSpy = vi.fn(
      async () => new Response('{"error":{"message":"Invalid API Key"}}', { status: 401 }),
    );
    vi.stubGlobal('fetch', fetchSpy);
    enterLiveMode();

    const user = userEvent.setup();
    renderApp();
    await askAgent(user, 'remove duplicate rows');
    await screen.findByText(/OpenRouter \/ Model Error/i, {}, { timeout: 5000 });

    await user.click(screen.getByTestId('open-usage'));

    expect(screen.getByTestId('usage-failures')).toHaveTextContent('1');
    const row = screen.getAllByTestId('usage-log-row')[0]!;
    expect(row.textContent).toContain('Failed');
    expect(screen.getByTestId('usage-key-status')).toHaveTextContent('Active');

    vi.unstubAllGlobals();
  });

  it('clears the ledger from the page', async () => {
    enterDemoMode();
    const user = userEvent.setup();
    renderApp();
    await askAgent(user, 'remove duplicate rows');
    await screen.findByRole('button', { name: /Apply Changes/i }, { timeout: 5000 });
    await user.click(screen.getByTestId('open-usage'));
    expect(screen.getByTestId('usage-requests')).toHaveTextContent('1');

    await user.click(screen.getByTestId('usage-clear'));

    await waitFor(() => expect(screen.getByTestId('usage-requests')).toHaveTextContent('0'));
    expect(screen.getByTestId('usage-empty')).toBeInTheDocument();
  });
});
