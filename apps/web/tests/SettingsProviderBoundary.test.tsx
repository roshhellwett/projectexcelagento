// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { renderApp, resetAppState, installBrowserStubs } from './helpers.js';

installBrowserStubs();

beforeEach(() => {
  resetAppState();
  window.location.hash = '';
});

describe('provider connection diagnostics', () => {
  it('redacts an echoed key and ignores arbitrary upstream error bodies', async () => {
    const user = userEvent.setup();
    renderApp();

    await user.click(screen.getByRole('button', { name: /API Keys & Settings/i }));
    const card = document.querySelector('.modal-card') as HTMLElement;
    const secret = 'sk-or-boundary-test-secret';
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify({ error: { message: `Provider rejected ${secret}` } }), {
            status: 401,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    );

    await user.type(within(card).getByPlaceholderText(/sk-or/i), secret);
    await user.click(within(card).getByRole('button', { name: /Test Connection/i }));

    const status = await within(card).findByText('HTTP 401: Provider rejected [redacted]');
    expect(status).toHaveTextContent('[redacted]');
    expect(status).not.toHaveTextContent(secret);
    vi.unstubAllGlobals();
  });
});
