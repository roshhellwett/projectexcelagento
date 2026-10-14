// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { AuthPage } from '../src/components/AuthPage.js';
import { LicensePanel } from '../src/components/LicensePanel.js';
import { PrivacyPolicyPage } from '../src/components/PrivacyPolicyPage.js';
import { SupportPage } from '../src/components/SupportPage.js';
import { TermsPage } from '../src/components/TermsPage.js';
import { TopNav } from '../src/components/TopNav.js';
import type { LicenseStatus } from '../src/lib/licensing.js';

const authMock = vi.hoisted(() => ({ useAuth: vi.fn() }));
const licenseMock = vi.hoisted(() => ({ useLicense: vi.fn() }));
const licensingMock = vi.hoisted(() => ({ submitSupportTicket: vi.fn() }));

vi.mock('../src/lib/auth-context.js', () => ({ useAuth: authMock.useAuth }));
vi.mock('../src/lib/license-context.js', () => ({ useLicense: licenseMock.useLicense }));
vi.mock('../src/lib/licensing.js', () => ({
  submitSupportTicket: licensingMock.submitSupportTicket,
}));

function createStatus(overrides: Partial<LicenseStatus> = {}): LicenseStatus {
  return {
    state: 'licensed',
    canUse: true,
    isAdmin: false,
    email: 'example@example.com',
    trialStartedAt: '2026-10-01T00:00:00.000Z',
    trialExpiresAt: '2026-10-31T00:00:00.000Z',
    daysRemaining: 30,
    license: {
      id: 'license-id',
      keyHint: 'EXCEL-••••-ABC123',
      rawKey: null,
      status: 'active',
      activatedAt: '2026-10-01T00:00:00.000Z',
      expiresAt: '2026-10-31T00:00:00.000Z',
      daysRemaining: 30,
    },
    banReason: null,
    ...overrides,
  };
}

function mockLicense(status: LicenseStatus): void {
  licenseMock.useLicense.mockReturnValue({
    status,
    loading: false,
    error: null,
    refresh: vi.fn(),
    activate: vi.fn(),
  });
}

beforeEach(() => {
  vi.clearAllMocks();
  authMock.useAuth.mockReturnValue({
    user: null,
    profile: null,
    signIn: vi.fn(),
    signUp: vi.fn(),
    resetPassword: vi.fn(),
    resendConfirmation: vi.fn(),
    signOut: vi.fn(),
  });
  licensingMock.submitSupportTicket.mockResolvedValue({ id: 'ticket-id' });
  mockLicense(createStatus());
});

describe('evaluation and activation copy', () => {
  it('describes a verified activation key instead of a card subscription', async () => {
    const onOpenSupport = vi.fn();
    render(<LicensePanel onOpenSupport={onOpenSupport} />);

    expect(screen.getByText('Activation Key Active')).toBeInTheDocument();
    expect(screen.getByText('Activation Key Access')).toBeInTheDocument();
    expect(screen.getByText(/not a recurring card subscription/i)).toBeInTheDocument();
    expect(
      screen.getByText(/not an unattended or background automation service/i),
    ).toBeInTheDocument();
    expect(screen.queryByText(/Subscription Active|Paid Subscription/i)).not.toBeInTheDocument();

    await userEvent.setup().click(screen.getByRole('button', { name: 'Open support form' }));
    expect(onOpenSupport).toHaveBeenCalledTimes(1);
  });

  it('explains the 30-day evaluation and lack of automatic renewal', () => {
    mockLicense(
      createStatus({
        state: 'trial',
        daysRemaining: 30,
        license: null,
      }),
    );
    render(<LicensePanel />);

    expect(screen.getByText('30-Day Evaluation Active')).toBeInTheDocument();
    expect(screen.getByText('30-Day Evaluation')).toBeInTheDocument();
    expect(
      screen.getByText(/No card is collected and access does not renew automatically/i),
    ).toBeInTheDocument();
  });

  it('directs an expired account to support instead of suggesting checkout', () => {
    mockLicense(
      createStatus({
        state: 'expired',
        canUse: false,
        daysRemaining: 0,
        license: null,
      }),
    );
    render(<LicensePanel />);

    expect(screen.getByText('Evaluation Ended')).toBeInTheDocument();
    expect(
      screen.getByText(/Request a server-issued activation key through support/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/this deployment has no in-app checkout/i)).toBeInTheDocument();
  });

  it('does not present an unverified account as expired when the service is unavailable', () => {
    licenseMock.useLicense.mockReturnValue({
      status: null,
      loading: false,
      error: 'The activation service could not verify this account.',
      refresh: vi.fn(),
      activate: vi.fn(),
    });
    render(<LicensePanel onOpenSupport={vi.fn()} />);

    expect(screen.getByText('Access Status Unavailable')).toBeInTheDocument();
    expect(screen.getByText(/No access or payment status has been changed/i)).toBeInTheDocument();
    expect(screen.queryByText('Evaluation Expired')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Enter activation key')).not.toBeInTheDocument();
  });

  it('sets the account entry point apart from model-provider usage', () => {
    render(<AuthPage initialMode="signup" onOpenSupport={vi.fn()} />);

    expect(screen.getByText(/30-day evaluation/i)).toBeInTheDocument();
    expect(screen.getByText(/No card or recurring subscription is required/i)).toBeInTheDocument();
    expect(
      screen.getByRole('button', { name: 'Need help with account access? Open support' }),
    ).toBeInTheDocument();
    expect(screen.getByText(/BYOK provider costs remain separate/i)).toBeInTheDocument();
  });

  it('uses access language in the signed-in navigation badge', () => {
    authMock.useAuth.mockReturnValue({
      user: { id: 'user-id', email: 'example@example.com' },
      profile: null,
      signOut: vi.fn(),
    });
    const noop = vi.fn();
    render(
      <TopNav
        fileName="sample.xlsx"
        activeSheetName="Sheet1"
        rowCount={1}
        colCount={1}
        canUndo={false}
        canRedo={false}
        historyLength={0}
        historyPosition={0}
        searchQuery=""
        searchMatchCount={0}
        onSearchChange={noop}
        onUndo={noop}
        onRedo={noop}
        onReset={noop}
        onFileUpload={noop}
        onExport={noop}
        onSelectFixture={noop}
        onOpenOperationModal={noop}
        onToggleHistory={noop}
        onOpenSettings={noop}
        licenseStatus={createStatus()}
      />,
    );

    expect(screen.getByText('30d Access')).toBeInTheDocument();
    expect(screen.queryByText('30d Pro')).not.toBeInTheDocument();
  });
});

describe('public access and support explanations', () => {
  it('explains the manual support path and separates activation from AI usage', () => {
    render(<SupportPage onBack={vi.fn()} />);

    expect(
      screen.getByText(/no in-app checkout or recurring card subscription/i),
    ).toBeInTheDocument();
    expect(screen.getByText('Does activation pay for AI model usage?')).toBeInTheDocument();
    expect(screen.getByText('Does activation include unattended automation?')).toBeInTheDocument();
    expect(screen.getByText(/does not create a background workbook queue/i)).toBeInTheDocument();
    expect(
      screen.getByText(/BYOK model-provider usage and any provider costs remain separate/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/Please do not include API keys or workbook contents/i),
    ).toBeInTheDocument();
  });

  it('states that the privacy flow has no embedded payment-card processing', () => {
    render(<PrivacyPolicyPage onBack={vi.fn()} />);

    expect(screen.getByText(/No payment-card flow is implemented here/i)).toBeInTheDocument();
    expect(
      screen.getByText(/Activation controls access to the official hosted deployment/i),
    ).toBeInTheDocument();
  });

  it('states the current activation boundary in the terms page', () => {
    render(<TermsPage onBack={vi.fn()} />);

    expect(
      screen.getByRole('heading', { name: '8. Official deployment evaluation and activation' }),
    ).toBeInTheDocument();
    expect(screen.getByText(/does not provide in-app checkout/i)).toBeInTheDocument();
    expect(screen.getByText(/does not include model-provider usage/i)).toBeInTheDocument();
  });
});
