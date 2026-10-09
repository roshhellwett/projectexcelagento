import React, { useEffect, useRef, useState } from 'react';
import {
  CheckCircle2,
  Copy,
  ExternalLink,
  KeyRound,
  LockKeyhole,
  Mail,
  RefreshCw,
  ShieldAlert,
  Sparkles,
  UserCheck,
} from 'lucide-react';
import { useLicense } from '../lib/license-context.js';
import { writeClipboardText } from '../lib/clipboard.js';

interface LicensePanelProps {
  onBack?: () => void;
  onOpenAdmin?: () => void;
  locked?: boolean;
}

function formatDate(value: string | null | undefined): string {
  if (!value) return '—';
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime())
    ? '—'
    : parsed.toLocaleDateString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
      });
}

function stateCopy(
  state: string | undefined,
  isAdmin?: boolean,
  daysRemaining?: number,
): { label: string; tone: string; detail: string } {
  if (isAdmin) {
    return {
      label: 'Super Admin',
      tone: 'owner',
      detail: 'You have permanent, unrestricted administrative access across the entire platform.',
    };
  }
  if (state === 'licensed') {
    return {
      label: 'Subscription Active',
      tone: 'success',
      detail: `Your paid activation is verified with ${daysRemaining ?? 0} day(s) remaining.`,
    };
  }
  if (state === 'trial') {
    return {
      label: '30-Day Trial Active',
      tone: 'success',
      detail: `Your free 30-day evaluation is active with ${daysRemaining ?? 0} day(s) remaining.`,
    };
  }
  if (state === 'banned') {
    return {
      label: 'Access Suspended',
      tone: 'danger',
      detail: 'This account has been suspended by an administrator.',
    };
  }
  return {
    label: 'Activation Required',
    tone: 'danger',
    detail:
      'Your 30-day trial has ended. Enter an activation key (30 days, 60 days, etc.) to continue.',
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
  const [keyCopied, setKeyCopied] = useState(false);
  const keyCopyTimerRef = useRef<number | null>(null);

  useEffect(
    () => () => {
      if (keyCopyTimerRef.current !== null) window.clearTimeout(keyCopyTimerRef.current);
    },
    [],
  );

  const state = stateCopy(status?.state, status?.isAdmin, status?.daysRemaining);
  const isExpired =
    !status?.isAdmin && (status?.state === 'expired' || (status?.daysRemaining ?? 0) <= 0);

  const copyKeyText = async (textToCopy: string) => {
    if (await writeClipboardText(textToCopy)) {
      setKeyCopied(true);
      if (keyCopyTimerRef.current !== null) window.clearTimeout(keyCopyTimerRef.current);
      keyCopyTimerRef.current = window.setTimeout(() => {
        keyCopyTimerRef.current = null;
        setKeyCopied(false);
      }, 1600);
    } else {
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
      setNotice('Activation successful! Your account access has been updated.');
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
          <span className="studio-eyebrow">ACCOUNT &amp; SUBSCRIPTION</span>
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
                  : 'Account & Activation'}
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
          {/* Account Profile Card */}
          <article className="license-card license-account-card">
            <div className="license-card-heading">
              <div>
                <span className="studio-eyebrow">ACCOUNT DETAILS</span>
                <h2>{status?.isAdmin ? 'Root Administrator' : 'User Account'}</h2>
              </div>
              <UserCheck size={21} aria-hidden="true" />
            </div>
            <dl className="license-details">
              <div>
                <dt>Signed-in email</dt>
                <dd>
                  <strong>{status?.email || 'Loading account session…'}</strong>
                </dd>
              </div>
              <div>
                <dt>Membership tier</dt>
                <dd>
                  {status?.isAdmin ? (
                    <span style={{ color: 'var(--brand-emerald-dark, #059669)', fontWeight: 600 }}>
                      Lifetime Administrator (∞)
                    </span>
                  ) : status?.state === 'licensed' ? (
                    <span style={{ color: 'var(--brand-emerald-dark, #059669)', fontWeight: 600 }}>
                      Paid Subscription
                    </span>
                  ) : isExpired ? (
                    <span style={{ color: '#ef4444', fontWeight: 600 }}>Evaluation Expired</span>
                  ) : (
                    <span style={{ color: 'var(--brand-emerald-dark, #059669)', fontWeight: 600 }}>
                      Free 30-Day Trial
                    </span>
                  )}
                </dd>
              </div>
              <div>
                <dt>Days remaining</dt>
                <dd>
                  {status?.isAdmin ? (
                    <strong style={{ color: 'var(--brand-emerald-dark, #059669)' }}>
                      Unlimited Lifetime
                    </strong>
                  ) : status?.daysRemaining && status.daysRemaining > 0 ? (
                    <strong>{status.daysRemaining} day(s)</strong>
                  ) : (
                    <strong style={{ color: '#ef4444' }}>0 days (Expired)</strong>
                  )}
                </dd>
              </div>
              <div>
                <dt>Access expires on</dt>
                <dd>
                  {status?.isAdmin ? (
                    <span style={{ color: 'var(--brand-emerald-dark, #059669)' }}>
                      Never Expires
                    </span>
                  ) : (
                    formatDate(
                      status?.state === 'licensed'
                        ? status?.license?.expiresAt || status?.trialExpiresAt
                        : status?.trialExpiresAt,
                    )
                  )}
                </dd>
              </div>
            </dl>
            <p className="license-muted" style={{ marginTop: '1rem' }}>
              Every email registered automatically receives a full 30-day free trial. You can log in
              from any browser, computer, or device with your email account anytime.
            </p>
          </article>

          {/* Activation & Keys Card */}
          <article className="license-card">
            <div className="license-card-heading">
              <div>
                <span className="studio-eyebrow">KEY ACTIVATION</span>
                <h2>
                  {status?.isAdmin
                    ? 'Root Permanent Entitlement'
                    : status?.license?.keyHint || 'Activation Key'}
                </h2>
              </div>
              <KeyRound size={21} aria-hidden="true" />
            </div>

            {status?.license?.rawKey && (
              <div
                style={{
                  marginBottom: '1rem',
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
                  ACTIVE KEY
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
              </div>
            )}

            {/* Trial Expiring Soon Warning */}
            {status?.state === 'trial' &&
              (status.daysRemaining ?? 0) <= 5 &&
              (status.daysRemaining ?? 0) > 0 && (
                <div
                  style={{
                    marginBottom: '1rem',
                    padding: '0.85rem',
                    background: 'rgba(245, 158, 11, 0.08)',
                    border: '1px solid rgba(245, 158, 11, 0.25)',
                    borderRadius: '8px',
                  }}
                >
                  <strong
                    style={{
                      display: 'block',
                      color: 'var(--accent-amber, #f59e0b)',
                      marginBottom: '0.25rem',
                    }}
                  >
                    Trial Expiring Soon ({status.daysRemaining} days left)
                  </strong>
                  <p
                    style={{
                      margin: 0,
                      fontSize: '0.85rem',
                      color: 'var(--color-text-secondary, #94a3b8)',
                    }}
                  >
                    Your 30-day evaluation is coming to an end. Enter an activation key (30 days, 60
                    days, etc.) below to ensure uninterrupted access.
                  </p>
                </div>
              )}

            {isExpired && (
              <div
                style={{
                  marginBottom: '1rem',
                  padding: '0.85rem',
                  background: 'rgba(239, 68, 68, 0.08)',
                  border: '1px solid rgba(239, 68, 68, 0.25)',
                  borderRadius: '8px',
                }}
              >
                <strong style={{ display: 'block', color: '#ef4444', marginBottom: '0.25rem' }}>
                  Your 30-Day Trial Has Ended
                </strong>
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--color-text-secondary)' }}>
                  Enter an activation key below to unlock your workspace. Keys can be purchased for
                  30 days, 60 days, or customized periods.
                </p>
              </div>
            )}

            {!status?.isAdmin && (
              <form
                className="license-activate-form"
                onSubmit={(event) => void handleActivate(event)}
              >
                <label htmlFor="license-key">
                  {status?.state === 'licensed'
                    ? 'Extend access with a new key'
                    : 'Enter activation key'}
                </label>
                <div className="license-key-row">
                  <input
                    id="license-key"
                    className="form-input"
                    value={key}
                    onChange={(event) =>
                      setKey(event.target.value.toUpperCase().replace(/\s+/g, ''))
                    }
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

            <div
              style={{
                marginTop: '1.25rem',
                padding: '0.85rem',
                borderRadius: '8px',
                background: 'rgba(255, 255, 255, 0.03)',
                border: '1px solid rgba(255, 255, 255, 0.06)',
              }}
            >
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '0.5rem',
                  marginBottom: '0.35rem',
                }}
              >
                <Sparkles size={15} style={{ color: '#10b981' }} />
                <strong style={{ fontSize: '0.88rem' }}>Need an activation key?</strong>
              </div>
              <p
                style={{
                  margin: 0,
                  fontSize: '0.82rem',
                  color: 'var(--color-text-secondary)',
                  lineHeight: 1.5,
                }}
              >
                Activation keys are available for <strong>30 days</strong>, <strong>60 days</strong>
                , or longer periods. To order or renew, contact{' '}
                <a
                  href={`mailto:zenithprojects@icloud.com?subject=ExcelAgento%20Activation%20Key%20Request&body=Hello,%0D%0A%0D%0AI would like to purchase an activation key for ExcelAgento.%0D%0AMy registered email: ${encodeURIComponent(status?.email || '')}`}
                  style={{ textDecoration: 'underline', color: 'var(--color-primary)' }}
                >
                  zenithprojects@icloud.com
                </a>
                .
              </p>
            </div>
          </article>
        </section>

        <section className="license-support-strip">
          <div>
            <strong>Support &amp; Assistance</strong>
            <span>
              Need help with your account or activation key? Contact our support desk anytime with
              your account email.
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
              href={`mailto:zenithprojects@icloud.com?subject=ExcelAgento%20Support&body=Account:%20${encodeURIComponent(status?.email || '')}`}
            >
              <Mail size={14} style={{ marginRight: '0.35rem' }} /> Contact support{' '}
              <ExternalLink size={14} />
            </a>
          </div>
        </section>
      </main>
    </div>
  );
};
