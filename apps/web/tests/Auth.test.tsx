// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { act, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { AuthPage } from '../src/components/AuthPage.js';
import { AuthProvider, useAuth } from '../src/lib/auth-context.js';
import { AuthCallbackPage } from '../src/components/AuthCallbackPage.js';
import type { AuthCallbackInfo } from '../src/lib/auth-redirect.js';

const auth = vi.hoisted(() => ({
  signUp: vi.fn(),
  signInWithPassword: vi.fn(),
  resend: vi.fn(),
  resetPasswordForEmail: vi.fn(),
  updateUser: vi.fn(),
  callback: null as AuthCallbackInfo | null,
  getSession: vi.fn().mockResolvedValue({ data: { session: null } }),
  onAuthStateChange: vi.fn().mockReturnValue({ data: { subscription: { unsubscribe: vi.fn() } } }),
}));
vi.mock('../src/lib/supabase-client.js', () => ({
  get initialAuthCallback() {
    return auth.callback;
  },
  supabase: {
    auth,
    from: () => ({
      select: () => ({ eq: () => ({ maybeSingle: async () => ({ data: null, error: null }) }) }),
    }),
  },
}));

function AccountState() {
  const { user, session } = useAuth();
  return (
    <output data-testid="account-state">{user && session ? 'signed-in' : 'signed-out'}</output>
  );
}

beforeEach(() => {
  vi.clearAllMocks();
  auth.callback = null;
  auth.getSession.mockResolvedValue({ data: { session: null } });
  auth.resend.mockResolvedValue({ error: null });
  auth.resetPasswordForEmail.mockResolvedValue({ error: null });
  auth.updateUser.mockResolvedValue({ error: null });
  window.history.replaceState(null, '', '/');
});

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
    expect(auth.signUp).toHaveBeenCalledWith(
      expect.objectContaining({
        options: expect.objectContaining({
          emailRedirectTo: `${window.location.origin}/?auth=confirm`,
        }),
      }),
    );
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

  it('resends verification from the sign-in fallback, preserving email and preventing immediate repeats', async () => {
    render(
      <AuthProvider>
        <AuthPage />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(screen.getByLabelText('Email address'), 'example@example.com');
    await user.click(screen.getByRole('button', { name: 'Resend confirmation email' }));
    expect(screen.queryByLabelText('Password')).not.toBeInTheDocument();
    expect(screen.getByLabelText('Email address')).toHaveValue('example@example.com');
    await user.click(screen.getByRole('button', { name: 'Resend confirmation email' }));
    expect(auth.resend).toHaveBeenCalledWith({
      type: 'signup',
      email: 'example@example.com',
      options: { emailRedirectTo: `${window.location.origin}/?auth=confirm` },
    });
    expect(await screen.findByText(/use the newest email/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Send another link in/ })).toBeDisabled();
    expect(auth.resend).toHaveBeenCalledTimes(1);
  });

  it('requests password recovery for the current app and keeps email when returning to sign in', async () => {
    render(
      <AuthProvider>
        <AuthPage />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(screen.getByLabelText('Email address'), 'example@example.com');
    await user.click(screen.getByRole('button', { name: 'Forgot password?' }));
    await user.click(screen.getByRole('button', { name: 'Send recovery link' }));
    expect(auth.resetPasswordForEmail).toHaveBeenCalledWith('example@example.com', {
      redirectTo: `${window.location.origin}/?auth=recovery`,
    });
    expect(await screen.findByText(/If an account exists/)).toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: 'Back to sign in' }));
    expect(screen.getByLabelText('Email address')).toHaveValue('example@example.com');
  });

  it('surfaces rate limits without claiming an email was sent and keeps retry available', async () => {
    auth.resend.mockResolvedValue({
      error: { status: 429, code: 'over_email_send_rate_limit', message: 'rate limited' },
    });
    render(
      <AuthProvider>
        <AuthPage initialMode="verify" />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(screen.getByLabelText('Email address'), 'example@example.com');
    await user.click(screen.getByRole('button', { name: 'Resend confirmation email' }));
    expect(await screen.findByRole('alert')).toHaveTextContent('Wait a minute');
    expect(screen.getByLabelText('Email address')).toHaveValue('example@example.com');
    expect(screen.getByRole('button', { name: 'Resend confirmation email' })).toBeEnabled();
  });
});

const session = {
  access_token: 'test-token',
  user: { id: 'confirmed-account', email: 'example@example.com' },
};

function CallbackFlow({ onContinue = vi.fn(), onVerify = vi.fn(), onRecover = vi.fn() }) {
  const { callback } = useAuth();
  return (
    callback && (
      <AuthCallbackPage
        callback={callback}
        onContinue={onContinue}
        onSignIn={vi.fn()}
        onRequestRecovery={onRecover}
        onRequestVerification={onVerify}
      />
    )
  );
}

describe('email callback fallbacks', () => {
  it('waits for session initialization before clearing URL tokens or offering workspace access', async () => {
    auth.callback = { kind: 'confirm', error: null };
    window.history.replaceState(
      null,
      '',
      '/?auth=confirm#access_token=test-token&refresh_token=test-refresh&type=signup',
    );
    let finishSession!: (result: { data: { session: typeof session } }) => void;
    auth.getSession.mockReturnValue(
      new Promise((resolve) => {
        finishSession = resolve;
      }),
    );
    const onContinue = vi.fn();
    render(
      <AuthProvider>
        <CallbackFlow onContinue={onContinue} />
      </AuthProvider>,
    );
    expect(screen.getByRole('status')).toHaveTextContent('Verifying account access');
    expect(screen.queryByRole('button', { name: 'Continue to workspace' })).not.toBeInTheDocument();
    expect(window.location.hash).toContain('access_token');
    await act(async () => finishSession({ data: { session } }));
    expect(window.location.hash).toBe('');
    expect(window.location.search).toBe('');
    await userEvent
      .setup()
      .click(await screen.findByRole('button', { name: 'Continue to workspace' }));
    expect(onContinue).toHaveBeenCalledTimes(1);
  });

  it('offers verification resend for an expired link even when a previous session exists', async () => {
    auth.callback = {
      kind: 'confirm',
      error: 'This email link has expired or has already been used.',
    };
    auth.getSession.mockResolvedValue({ data: { session } });
    const onVerify = vi.fn();
    render(
      <AuthProvider>
        <CallbackFlow onVerify={onVerify} />
      </AuthProvider>,
    );
    expect(await screen.findByRole('alert')).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: 'Continue to workspace' })).not.toBeInTheDocument();
    await userEvent
      .setup()
      .click(screen.getByRole('button', { name: 'Request a new confirmation link' }));
    expect(onVerify).toHaveBeenCalledTimes(1);
  });

  it('offers a fresh recovery link if no recovery session can be established', async () => {
    auth.callback = { kind: 'recovery', error: null };
    const onRecover = vi.fn();
    render(
      <AuthProvider>
        <CallbackFlow onRecover={onRecover} />
      </AuthProvider>,
    );
    await userEvent
      .setup()
      .click(await screen.findByRole('button', { name: 'Request a new recovery link' }));
    expect(onRecover).toHaveBeenCalledTimes(1);
    expect(auth.updateUser).not.toHaveBeenCalled();
  });

  it('checks password confirmation, preserves rejected input, and only reports saved after the provider accepts', async () => {
    auth.callback = { kind: 'recovery', error: null };
    auth.getSession.mockResolvedValue({ data: { session } });
    auth.updateUser
      .mockResolvedValueOnce({ error: { message: 'Choose a different password.' } })
      .mockResolvedValue({ error: null });
    render(
      <AuthProvider>
        <CallbackFlow />
      </AuthProvider>,
    );
    const user = userEvent.setup();
    await user.type(await screen.findByLabelText('New password'), 'new-test-password');
    await user.type(screen.getByLabelText('Confirm new password'), 'different-password');
    await user.click(screen.getByRole('button', { name: 'Save new password' }));
    expect(await screen.findByRole('alert')).toHaveTextContent('do not match');
    expect(auth.updateUser).not.toHaveBeenCalled();
    await user.clear(screen.getByLabelText('Confirm new password'));
    await user.type(screen.getByLabelText('Confirm new password'), 'new-test-password');
    await user.click(screen.getByRole('button', { name: 'Save new password' }));
    expect(await screen.findByRole('alert')).toHaveTextContent('Choose a different password');
    expect(screen.getByLabelText('New password')).toHaveValue('new-test-password');
    await user.click(screen.getByRole('button', { name: 'Save new password' }));
    expect(await screen.findByRole('heading', { name: 'Password updated' })).toBeInTheDocument();
    expect(auth.updateUser).toHaveBeenLastCalledWith({ password: 'new-test-password' });
  });
});
