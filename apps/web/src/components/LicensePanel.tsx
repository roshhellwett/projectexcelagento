import React, { useState } from 'react';
import {
  CheckCircle2,
  Copy,
  ExternalLink,
  KeyRound,
  Laptop,
  LockKeyhole,
  RefreshCw,
  ShieldAlert,
} from 'lucide-react';
import { getInstallId } from '../lib/device-identity.js';
import { useLicense } from '../lib/license-context.js';

interface LicensePanelProps {
  onBack?: () => void;
  onOpenAdmin?: () => void;
  locked?: boolean;
}

function formatDate(value: string | null | undefined): string {
  if (!value) return '—';
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '—' : parsed.toLocaleString();
}

function stateCopy(
  state: string | undefined,
  isAdmin?: boolean,
): { label: string; tone: string; detail: string } {
  if (isAdmin) {
    return {
      label: 'Super Admin / Owner',
      tone: 'success',
      detail:
        'You have permanent, unrestricted administrative root access across the entire platform.',
    };
  }
  if (state === 'licensed')
    return {
      label: 'Activated',
      tone: 'success',
      detail: 'This installation is licensed and verified.',
    };
  if (state === 'trial')
    return { label: 'Trial active', tone: 'success', detail: 'Your 30-day evaluation is active.' };
  if (state === 'activation_required')
    return {
      label: 'Activation required',
      tone: 'warning',
      detail:
        'This key was transferred to your account. Activate it on this installation to continue.',
    };
  if (state === 'banned')
    return {
      label: 'Access suspended',
      tone: 'danger',
      detail: 'This account has been suspended by an administrator.',
    };
  if (state === 'device_banned')
    return {
      label: 'Device suspended',
      tone: 'danger',
      detail: 'This browser installation has been suspended by an administrator.',
    };
  if (state === 'device_mismatch')
    return {
      label: 'Device transfer required',
      tone: 'warning',
      detail:
        'Your account is bound to another installation. Contact support to verify a transfer.',
    };
  if (state === 'revoked')
    return {
      label: 'License revoked',
      tone: 'danger',
      detail: 'The activation assigned to this account was revoked. Contact support for review.',
    };
  if (state === 'email_unconfirmed')
    return {
      label: 'Confirm your email',
      tone: 'warning',
      detail: 'Confirm your Supabase account email before starting the trial.',
    };
  return {
    label: 'Activation required',
    tone: 'danger',
    detail: 'Your trial has ended. Enter a valid activation key to continue.',
  };
}

export const LicensePanel: React.FC<LicensePanelProps> = ({
  onBack,
  onOpenAdmin,
  locked = false,
}) => {
  const { status, loading, error, refresh, activate } = useLicense();
  const [key, setKey] = useState('');
  const [notice, setNotice] = useState('');
  const [activating, setActivating] = useState(false);
  const [copied, setCopied] = useState(false);
  const [keyCopied, setKeyCopied] = useState(false);
  const deviceId = status?.deviceId || getInstallId();
  const state = stateCopy(status?.state, status?.isAdmin);
  const canEnterActivationKey =
    !status?.isAdmin &&
    (!status ||
      status.state === 'expired' ||
      status.state === 'activation_required' ||
      (locked &&
        !['banned', 'device_banned', 'device_mismatch', 'email_unconfirmed', 'revoked'].includes(
          status.state,
        )));

  const copyDeviceId = async () => {
    try {
      await navigator.clipboard.writeText(deviceId);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1600);
    } catch {
      setNotice('Copy was blocked by the browser. Select the ID manually.');
    }
  };

  const copyKeyText = async (textToCopy: string) => {
    try {
      await navigator.clipboard.writeText(textToCopy);
      setKeyCopied(true);
      window.setTimeout(() => setKeyCopied(false), 1600);
    } catch {
      setNotice('Copy was blocked by the browser. Select the key manually.');
    }
  };

  const handleActivate = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!key.trim() || activating) return;
    setActivating(true);
    setNotice('');
    const result = await activate(key);
    if (result.success) {
      setKey('');
      setNotice('Activation verified. Your account is unlocked on this installation.');
    } else {
      setNotice(result.error || 'The activation key could not be verified.');
    }
    setActivating(false);
  };

  return (
    <div className="app-container app-page-scroll license-page">
      <main className="license-shell" aria-labelledby="license-title">
        <div className="license-topbar">
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={onBack}
            disabled={!onBack}
          >
            ← Back to workspace
          </button>
          <span className="studio-eyebrow">ACCOUNT SECURITY · ACTIVATION</span>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => void refresh()}
            disabled={loading}
          >
            <RefreshCw size={14} className={loading ? 'license-spin' : undefined} />
            Refresh status
          </button>
        </div>

        <section className={`license-hero license-tone-${state.tone}`}>
          <div className="license-hero-icon" aria-hidden="true">
            {state.tone === 'danger' ? <LockKeyhole size={26} /> : <CheckCircle2 size={26} />}
          </div>
          <div>
            <span className="studio-eyebrow">EXCELAGENTO ACCESS CONTROL</span>
            <h1 id="license-title">
              {status?.isAdmin
                ? 'Super Admin / System Owner'
                : locked
                  ? 'Workspace access is locked'
                  : 'Account & activation'}
            </h1>
            <p>{state.detail}</p>
          </div>
          <span className={`license-state-badge ${state.tone}`}>{state.label}</span>
        </section>

        {error && (
          <div className="license-alert danger" role="alert">
            <ShieldAlert size={17} />
            <span>{error}</span>
          </div>
        )}
        {notice && (
          <div className="license-alert" role="status">
            <CheckCircle2 size={17} />
            <span>{notice}</span>
          </div>
        )}

        <section className="license-grid">
          <article className="license-card license-account-card">
            <div className="license-card-heading">
              <div>
                <span className="studio-eyebrow">IDENTITY</span>
                <h2>{status?.isAdmin ? 'Root Administrator' : 'Signed-in account'}</h2>
              </div>
              <Laptop size={21} aria-hidden="true" />
            </div>
            <dl className="license-details">
              <div>
                <dt>Verified email</dt>
                <dd>{status?.email || 'Waiting for account session…'}</dd>
              </div>
              <div>
                <dt>Unique Device ID</dt>
                <dd className="license-device-value">
                  <code>{deviceId}</code>
                  <button
                    type="button"
                    className="btn btn-ghost btn-icon btn-sm"
                    onClick={() => void copyDeviceId()}
                    aria-label="Copy unique device ID"
                    title="Copy unique device ID"
                  >
                    {copied ? <CheckCircle2 size={14} /> : <Copy size={14} />}
                  </button>
                </dd>
              </div>
              <div>
                <dt>{status?.isAdmin ? 'Access Expiration' : 'Trial expires'}</dt>
                <dd>
                  {status?.isAdmin ? (
                    <strong style={{ color: 'var(--brand-emerald-dark, #059669)' }}>
                      Never Expires (Permanent ∞)
                    </strong>
                  ) : (
                    formatDate(status?.trialExpiresAt)
                  )}
                </dd>
              </div>
              <div>
                <dt>Current access</dt>
                <dd>
                  {status?.isAdmin ? (
                    <strong style={{ color: 'var(--brand-emerald-dark, #059669)' }}>
                      Unlimited Lifetime Access
                    </strong>
                  ) : status?.daysRemaining ? (
                    `${status.daysRemaining} day(s) remaining`
                  ) : (
                    'No active days'
                  )}
                </dd>
              </div>
            </dl>
            <p className="license-muted">
              Hardware-derived Unique Device ID (HWID) based on CPU cores, GPU WebGL rendering, and
              display characteristics. Remains identical across all browsers and private browsing
              sessions on this machine. If you lose your email access, support verifies this ID to
              recover and restore your license.
            </p>
          </article>

          <article className="license-card">
            <div className="license-card-heading">
              <div>
                <span className="studio-eyebrow">ENTITLEMENT</span>
                <h2>
                  {status?.isAdmin
                    ? 'Super Admin Permanent License'
                    : status?.license?.keyHint || '30-day evaluation'}
                </h2>
              </div>
              <KeyRound size={21} aria-hidden="true" />
            </div>
            <dl className="license-details">
              <div>
                <dt>Status</dt>
                <dd>{state.label}</dd>
              </div>
              <div>
                <dt>Activation started</dt>
                <dd>{formatDate(status?.license?.activatedAt || status?.trialStartedAt)}</dd>
              </div>
              <div>
                <dt>Entitlement expires</dt>
                <dd>
                  {status?.isAdmin ? (
                    <strong style={{ color: 'var(--brand-emerald-dark, #059669)' }}>
                      Never Expires (Lifetime ∞)
                    </strong>
                  ) : (
                    formatDate(status?.license?.expiresAt || status?.trialExpiresAt)
                  )}
                </dd>
              </div>
            </dl>

            {status?.license?.rawKey && (
              <div
                style={{
                  marginTop: '1rem',
                  padding: '0.75rem',
                  background: 'var(--color-bg-secondary, rgba(255,255,255,0.04))',
                  borderRadius: '8px',
                  border: '1px solid var(--color-border, rgba(255,255,255,0.08))',
                }}
              >
                <span
                  className="studio-eyebrow"
                  style={{ display: 'block', marginBottom: '0.35rem' }}
                >
                  YOUR ACTIVATION KEY
                </span>
                <div
                  style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flexWrap: 'wrap' }}
                >
                  <code style={{ fontSize: '0.85rem', wordBreak: 'break-all', fontWeight: 600 }}>
                    {status.license.rawKey}
                  </code>
                  <button
                    type="button"
                    className="btn btn-ghost btn-sm"
                    onClick={() => void copyKeyText(status.license!.rawKey!)}
                    style={{
                      padding: '0.2rem 0.5rem',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '0.25rem',
                    }}
                  >
                    {keyCopied ? <CheckCircle2 size={13} /> : <Copy size={13} />}
                    {keyCopied ? 'Copied' : 'Copy Key'}
                  </button>
                </div>
                <small className="license-muted" style={{ display: 'block', marginTop: '0.25rem' }}>
                  Linked to your Google account and Unique Device ID. Save this for recovery.
                </small>
              </div>
            )}

            {status && status.daysRemaining <= 0 && (
              <div
                style={{
                  marginTop: '1rem',
                  padding: '0.75rem',
                  background: 'rgba(239, 68, 68, 0.1)',
                  border: '1px solid rgba(239, 68, 68, 0.3)',
                  borderRadius: '8px',
                }}
              >
                <strong style={{ display: 'block', color: '#f87171', marginBottom: '0.25rem' }}>
                  Evaluation Expired
                </strong>
                <p style={{ margin: 0, fontSize: '0.85rem' }}>
                  Your 30-day access has ended. Contact{' '}
                  <a
                    href={`mailto:zenithprojects@icloud.com?subject=ExcelAgento%20License%20Extension&body=Email:%20${encodeURIComponent(status?.email || '')}%0D%0AUnique%20Device%20ID:%20${encodeURIComponent(deviceId)}`}
                    style={{ textDecoration: 'underline' }}
                  >
                    zenithprojects@icloud.com
                  </a>{' '}
                  with your Unique Device ID (<code>{deviceId}</code>) to add days or purchase
                  extended access.
                </p>
              </div>
            )}

            {canEnterActivationKey && (
              <form
                className="license-activate-form"
                onSubmit={(event) => void handleActivate(event)}
                style={{ marginTop: '1rem' }}
              >
                <label htmlFor="license-key">Enter new activation key</label>
                <div className="license-key-row">
                  <input
                    id="license-key"
                    className="form-input"
                    value={key}
                    onChange={(event) => setKey(event.target.value.toUpperCase())}
                    placeholder="EXCEL-XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX"
                    autoComplete="off"
                    spellCheck={false}
                    disabled={activating}
                  />
                  <button
                    type="submit"
                    className="btn btn-primary"
                    disabled={!key.trim() || activating}
                  >
                    {activating ? 'Verifying…' : 'Activate'}
                  </button>
                </div>
              </form>
            )}
            <p className="license-muted" style={{ marginTop: '0.75rem' }}>
              Need a paid key, a duration change, or a recovery transfer? Contact{' '}
              <a href="mailto:zenithprojects@icloud.com">zenithprojects@icloud.com</a> and include
              your verified email and Unique Device ID.
            </p>
          </article>
        </section>

        <section className="license-support-strip">
          <div>
            <strong>Recovery and device transfer</strong>
            <span>
              Support can reassign an activation after verifying account ownership and device
              details. Every transfer is recorded in the administrative audit log.
            </span>
          </div>
          <div className="license-support-actions">
            {status?.isAdmin && onOpenAdmin && (
              <button type="button" className="btn btn-secondary" onClick={onOpenAdmin}>
                Open admin console
              </button>
            )}
            <a
              className="btn btn-ghost"
              href="mailto:zenithprojects@icloud.com?subject=ExcelAgento%20activation%20support"
            >
              Contact support <ExternalLink size={14} />
            </a>
          </div>
        </section>
      </main>
    </div>
  );
};
