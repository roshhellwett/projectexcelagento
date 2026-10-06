// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { AuthPage } from '../src/components/AuthPage.js';
import { AuthProvider, useAuth } from '../src/lib/auth-context.js';

const auth = vi.hoisted(() => ({
  signUp: vi.fn(),
  signInWithPassword: vi.fn(),
  getSession: vi.fn().mockResolvedValue({ data: { session: null } }),
  onAuthStateChange: vi.fn().mockReturnValue({ data: { subscription: { unsubscribe: vi.fn() } } }),
}));
vi.mock('../src/lib/supabase-client.js', () => ({ supabase: { auth } }));

function AccountState() {
  const { user, session } = useAuth();
  return (
    <output data-testid="account-state">{user && session ? 'signed-in' : 'signed-out'}</output>
  );
}

beforeEach(() => vi.clearAllMocks());

describe('account access', () => {
  it('keeps confirmation-required signup signed out even when Supabase returns a user', async () => {
    auth.signUp.mockResolvedValue({
      data: { user: { id: 'confirmation-required', email: 'example@example.com' }, session: null },
      error: null,
    });
    const onSuccess = vi.fn();
    render(
      <AuthProvider>
        <AuthPage initialMode="signup" onSuccess={onSuccess} />
        <AccountState />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(screen.getByLabelText('Email address'), 'example@example.com');
    await user.type(screen.getByLabelText('Password'), 'test-only-password');
    await user.click(screen.getByRole('button', { name: 'Create account' }));
    expect(await screen.findByText(/Check your email/)).toBeInTheDocument();
    expect(screen.getByTestId('account-state')).toHaveTextContent('signed-out');
    expect(onSuccess).not.toHaveBeenCalled();
  });

  it('explains transport failures and allows signup to be retried', async () => {
    auth.signUp.mockResolvedValue({ data: {}, error: new TypeError('Failed to fetch') });
    render(
      <AuthProvider>
        <AuthPage initialMode="signup" />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(screen.getByLabelText('Email address'), 'example@example.com');
    await user.type(screen.getByLabelText('Password'), 'test-only-password');
    await user.click(screen.getByRole('button', { name: 'Create account' }));
    expect(await screen.findByRole('alert')).toHaveTextContent(
      'Could not reach the account service',
    );
    await waitFor(() =>
      expect(screen.getByRole('button', { name: 'Create account' })).toBeEnabled(),
    );
    await user.click(screen.getByRole('button', { name: 'Create account' }));
    expect(auth.signUp).toHaveBeenCalledTimes(2);
  });
});
