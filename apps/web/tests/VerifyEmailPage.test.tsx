// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { VerifyEmailPage } from '../src/components/VerifyEmailPage.js';
import { AuthProvider } from '../src/lib/auth-context.js';
import { LicenseProvider } from '../src/lib/license-context.js';

const auth = vi.hoisted(() => ({
  getSession: vi.fn(),
  onAuthStateChange: vi.fn(),
  exchangeCodeForSession: vi.fn(),
  verifyOtp: vi.fn(),
  resend: vi.fn(),
}));

vi.mock('../src/lib/supabase-client.js', () => ({
  initialAuthCallback: null,
  supabase: {
    auth,
    from: () => ({
      select: () => ({
        eq: () => ({ maybeSingle: async () => ({ data: null, error: null }) }),
      }),
    }),
  },
}));

function renderPage(initialEmail = 'example@example.com') {
  return render(
    <AuthProvider>
      <LicenseProvider>
        <VerifyEmailPage initialEmail={initialEmail} onContinue={vi.fn()} onSignIn={vi.fn()} />
      </LicenseProvider>
    </AuthProvider>,
  );
}

beforeEach(() => {
  vi.clearAllMocks();
  window.history.replaceState(null, '', '/verify-email?auth=confirm');
  auth.getSession.mockResolvedValue({ data: { session: null }, error: null });
  auth.onAuthStateChange.mockReturnValue({
    data: { subscription: { unsubscribe: vi.fn() } },
  });
  auth.exchangeCodeForSession.mockResolvedValue({ data: { session: null }, error: null });
  auth.verifyOtp.mockResolvedValue({ data: { session: null }, error: null });
  auth.resend.mockResolvedValue({ error: null });
});

describe('email verification surface', () => {
  it('waits for Supabase session initialization before exchanging a confirmation link', async () => {
    window.history.replaceState(null, '', '/verify-email?auth=confirm&code=one-time-code');
    let finishSession!: (value: { data: { session: unknown }; error: null }) => void;
    auth.getSession.mockReturnValue(
      new Promise((resolve) => {
        finishSession = resolve;
      }),
    );
    const session = {
      user: {
        id: 'confirmed-user',
        email: 'example@example.com',
        email_confirmed_at: '2026-10-12T00:00:00.000Z',
      },
    };

    renderPage();
    expect(screen.getByRole('heading', { name: /Verifying Your Email/i })).toBeInTheDocument();
    expect(auth.exchangeCodeForSession).not.toHaveBeenCalled();

    finishSession({ data: { session }, error: null });
    await waitFor(() =>
      expect(
        screen.getByRole('heading', { name: /Email Confirmed Successfully/i }),
      ).toBeInTheDocument(),
    );
    expect(auth.exchangeCodeForSession).not.toHaveBeenCalled();
  });

  it('accepts a six-digit OTP, normalizes input, and opens the success state', async () => {
    auth.verifyOtp.mockResolvedValue({
      data: {
        session: {
          user: {
            email: 'example@example.com',
            email_confirmed_at: '2026-10-12T00:00:00.000Z',
          },
        },
      },
      error: null,
    });
    const user = userEvent.setup();
    renderPage();

    const otp = await screen.findByPlaceholderText('6-digit code');
    await user.type(otp, '12a345678');
    expect(otp).toHaveValue('123456');
    await user.click(screen.getByRole('button', { name: 'Verify' }));

    await waitFor(() =>
      expect(
        screen.getByRole('heading', { name: /Email Confirmed Successfully/i }),
      ).toBeInTheDocument(),
    );
    expect(auth.verifyOtp).toHaveBeenCalledWith({
      email: 'example@example.com',
      token: '123456',
      type: 'signup',
    });
  });
});
