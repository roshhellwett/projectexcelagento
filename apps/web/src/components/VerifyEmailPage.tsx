import React, { useEffect, useState } from 'react';
import {
  CheckCircle2,
  AlertCircle,
  Mail,
  Loader2,
  ArrowRight,
  RefreshCw,
  KeyRound,
  ShieldCheck,
  LifeBuoy,
} from 'lucide-react';
import { supabase } from '../lib/supabase-client.js';
import { useAuth } from '../lib/auth-context.js';
import { useLicense } from '../lib/license-context.js';

interface VerifyEmailPageProps {
  onContinue: () => void;
  onSignIn: () => void;
  onOpenSupport?: () => void;
  onOpenAdmin?: () => void;
  initialEmail?: string;
}

type VerificationStatus = 'verifying' | 'success' | 'pending' | 'error';

export const VerifyEmailPage: React.FC<VerifyEmailPageProps> = ({
  onContinue,
  onSignIn,
  onOpenSupport,
  onOpenAdmin,
  initialEmail = '',
}) => {
  const { session, loading: authLoading, resendConfirmation, dismissCallback } = useAuth();
  const { status: licenseStatus } = useLicense();

  const [status, setStatus] = useState<VerificationStatus>('verifying');
  const [targetEmail, setTargetEmail] = useState<string>(initialEmail);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [feedbackMessage, setFeedbackMessage] = useState<string>('');
  const [resendCooldown, setResendCooldown] = useState<number>(0);
  const [otpCode, setOtpCode] = useState<string>('');
  const [verifyingOtp, setVerifyingOtp] = useState<boolean>(false);
  const [isResending, setIsResending] = useState<boolean>(false);

  // Timer for resend cooldown
  useEffect(() => {
    if (resendCooldown <= 0) return;
    const interval = setInterval(() => {
      setResendCooldown((prev) => Math.max(0, prev - 1));
    }, 1000);
    return () => clearInterval(interval);
  }, [resendCooldown]);

  // Main verification logic from URL or session
  useEffect(() => {
    let mounted = true;
    let authTimer: ReturnType<typeof setTimeout> | undefined;

    async function handleVerification(): Promise<void> {
      if (typeof window === 'undefined') return;

      const url = new URL(window.location.href);
      const hashParams = new URLSearchParams(url.hash.replace(/^#\/?/, ''));
      const queryParams = url.searchParams;

      // Extract email hint if present
      const emailParam = queryParams.get('email') || hashParams.get('email') || initialEmail;
      if (emailParam && !targetEmail) {
        setTargetEmail(emailParam);
      }

      // 1. Check for explicit error in URL parameters
      const err = queryParams.get('error') || hashParams.get('error');
      const errCode = queryParams.get('error_code') || hashParams.get('error_code');
      const errDesc = queryParams.get('error_description') || hashParams.get('error_description');

      if (err || errCode) {
        if (!mounted) return;
        setStatus('error');
        setErrorMessage(
          errCode === 'otp_expired' || /expired|already/i.test(errDesc ?? '')
            ? 'This email verification link has expired or has already been used. Please request a fresh link below.'
            : errDesc ||
                'The email verification link could not be validated. Try requesting a new one.',
        );
        return;
      }

      // 2. Check for PKCE auth code (?code=...)
      const authCode = queryParams.get('code');
      if (authCode) {
        try {
          const { data, error } = await supabase.auth.exchangeCodeForSession(authCode);
          if (error) {
            if (!mounted) return;
            setStatus('error');
            setErrorMessage(error.message || 'Failed to exchange verification code.');
            return;
          }
          if (data.session) {
            if (!mounted) return;
            setStatus('success');
            setTargetEmail(data.session.user.email || targetEmail);
            dismissCallback();
            return;
          }
        } catch (ex) {
          if (!mounted) return;
          setStatus('error');
          setErrorMessage(ex instanceof Error ? ex.message : 'Error validating verification code.');
          return;
        }
      }

      // 3. Check for token_hash OTP verification (?token_hash=...&type=signup|email)
      const tokenHash = queryParams.get('token_hash') || hashParams.get('token_hash');
      const type = (queryParams.get('type') || hashParams.get('type') || 'signup') as
        'signup' | 'email';
      if (tokenHash) {
        try {
          const { data, error } = await supabase.auth.verifyOtp({
            token_hash: tokenHash,
            type: type,
          });
          if (error) {
            if (!mounted) return;
            setStatus('error');
            setErrorMessage(error.message || 'Token verification failed.');
            return;
          }
          if (data.session) {
            if (!mounted) return;
            setStatus('success');
            setTargetEmail(data.session.user.email || targetEmail);
            dismissCallback();
            return;
          }
        } catch (ex) {
          if (!mounted) return;
          setStatus('error');
          setErrorMessage(ex instanceof Error ? ex.message : 'Verification failed.');
          return;
        }
      }

      // 4. Check if session already exists and user email is confirmed
      if (session?.user) {
        if (session.user.email_confirmed_at) {
          if (!mounted) return;
          setStatus('success');
          setTargetEmail(session.user.email || targetEmail);
          return;
        }
      }

      // 5. If arrived via explicit auth=confirm without tokens or just pending verification
      const authMarker = queryParams.get('auth');
      if (authMarker === 'confirm' && !session) {
        // Wait a brief moment in case Supabase client handles tokens in background
        authTimer = setTimeout(() => {
          if (!mounted) return;
          supabase.auth.getSession().then(({ data: { session: s } }) => {
            if (s?.user?.email_confirmed_at) {
              setStatus('success');
              setTargetEmail(s.user.email || targetEmail);
            } else {
              setStatus('pending');
            }
          });
        }, 800);
        return;
      }

      // Otherwise default to pending check inbox
      if (!authLoading) {
        setStatus('pending');
      }
    }

    void handleVerification();

    return () => {
      mounted = false;
      if (authTimer) clearTimeout(authTimer);
    };
  }, [session, authLoading, initialEmail, dismissCallback]);

  // Handle manual OTP code submission
  const handleVerifyOtp = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!otpCode.trim() || !targetEmail.trim()) {
      setErrorMessage('Please provide both your account email and 6-digit confirmation code.');
      return;
    }

    setVerifyingOtp(true);
    setErrorMessage('');
    setFeedbackMessage('');

    try {
      const { data, error } = await supabase.auth.verifyOtp({
        email: targetEmail.trim(),
        token: otpCode.trim(),
        type: 'signup',
      });

      if (error) {
        setErrorMessage(error.message || 'Invalid confirmation code. Please check and try again.');
      } else if (data.session) {
        setStatus('success');
        setTargetEmail(data.session.user.email || targetEmail);
        setFeedbackMessage('Email verified successfully!');
      }
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : 'Failed to verify code.');
    } finally {
      setVerifyingOtp(false);
    }
  };

  // Handle Resend Confirmation
  const handleResend = async () => {
    if (resendCooldown > 0 || isResending) return;
    if (!targetEmail.trim()) {
      setErrorMessage('Please enter your account email address to send a new confirmation link.');
      return;
    }

    setIsResending(true);
    setErrorMessage('');
    setFeedbackMessage('');

    try {
      const result = await resendConfirmation(targetEmail.trim());
      if (result.success) {
        setResendCooldown(60);
        setFeedbackMessage(
          `Confirmation email sent to ${targetEmail.trim()}! Please check your inbox and spam folder.`,
        );
      } else {
        setErrorMessage(
          result.error || 'Could not send verification email. Please try again shortly.',
        );
      }
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : 'Failed to send confirmation email.');
    } finally {
      setIsResending(false);
    }
  };

  const isAdmin = Boolean(licenseStatus?.isAdmin);

  return (
    <div className="auth-page verify-email-page">
      <div className="auth-page-grid" aria-hidden="true" />
      <div className="auth-layout" style={{ maxWidth: 540 }}>
        <section className="auth-card verify-email-card" aria-labelledby="verify-title">
          {/* Brand header */}
          <div className="verify-card-top">
            <button
              type="button"
              className="auth-brand"
              onClick={onSignIn}
              aria-label="ExcelAgento brand"
            >
              <span className="auth-brand-mark">
                <img src="/excel-agent-logo.svg" alt="" />
              </span>
              <span>
                <strong>
                  Excel<span>Agento</span>
                </strong>
                <small>IDENTITY &amp; AUTHENTICATION</small>
              </span>
            </button>
            <span className="verify-security-badge">
              <ShieldCheck size={13} /> 256-Bit Auth
            </span>
          </div>

          {/* Alerts */}
          {errorMessage && (
            <div className="auth-alert is-error" role="alert" style={{ marginTop: '1.25rem' }}>
              <AlertCircle size={17} />
              <span>{errorMessage}</span>
            </div>
          )}

          {feedbackMessage && (
            <div className="auth-alert is-success" role="status" style={{ marginTop: '1.25rem' }}>
              <CheckCircle2 size={17} />
              <span>{feedbackMessage}</span>
            </div>
          )}

          {/* STATE 1: VERIFYING IN PROGRESS */}
          {status === 'verifying' && (
            <div className="verify-state-box verifying">
              <div className="verify-icon-pulse">
                <Loader2 size={38} className="auth-spin text-emerald" />
              </div>
              <h1 id="verify-title" className="verify-title">
                Verifying Your Email…
              </h1>
              <p className="verify-subtitle">
                Confirming security credentials with Supabase Auth and establishing your session.
              </p>
            </div>
          )}

          {/* STATE 2: SUCCESS */}
          {status === 'success' && (
            <div className="verify-state-box success">
              <div className="verify-icon-circle success">
                <CheckCircle2 size={44} />
              </div>
              <span className="studio-eyebrow text-emerald" style={{ letterSpacing: '0.08em' }}>
                VERIFICATION COMPLETE
              </span>
              <h1 id="verify-title" className="verify-title">
                Email Confirmed Successfully!
              </h1>
              <p className="verify-subtitle">
                Your account is verified and ready. All spreadsheet operations and agent features
                are unlocked.
              </p>

              {targetEmail && (
                <div className="verify-account-pill">
                  <Mail size={15} />
                  <span>{targetEmail}</span>
                  <span className="verify-pill-tag">Active</span>
                </div>
              )}

              <div className="verify-actions" style={{ marginTop: '1.75rem' }}>
                {isAdmin ? (
                  <button
                    type="button"
                    className="auth-submit-btn verify-cta-btn"
                    onClick={onOpenAdmin || onContinue}
                  >
                    Open Admin CMS Dashboard <ArrowRight size={16} />
                  </button>
                ) : (
                  <button
                    type="button"
                    className="auth-submit-btn verify-cta-btn"
                    onClick={onContinue}
                  >
                    Launch Spreadsheet Workspace <ArrowRight size={16} />
                  </button>
                )}
                <button type="button" className="auth-link-btn" onClick={onSignIn}>
                  Sign in with another account
                </button>
              </div>
            </div>
          )}

          {/* STATE 3: PENDING / CHECK INBOX */}
          {status === 'pending' && (
            <div className="verify-state-box pending">
              <div className="verify-icon-circle pending">
                <Mail size={40} />
              </div>
              <span className="studio-eyebrow" style={{ letterSpacing: '0.08em' }}>
                ACTION REQUIRED
              </span>
              <h1 id="verify-title" className="verify-title">
                Check Your Inbox
              </h1>
              <p className="verify-subtitle">
                We sent a secure verification link to activate your account.
              </p>

              {targetEmail ? (
                <div className="verify-account-pill">
                  <Mail size={15} />
                  <strong>{targetEmail}</strong>
                </div>
              ) : (
                <div className="verify-input-group" style={{ marginBottom: '1rem' }}>
                  <label htmlFor="target-email" className="auth-label">
                    Your Account Email
                  </label>
                  <input
                    id="target-email"
                    type="email"
                    className="auth-input"
                    placeholder="you@company.com"
                    value={targetEmail}
                    onChange={(e) => setTargetEmail(e.target.value)}
                  />
                </div>
              )}

              {/* Instructions checklist */}
              <div className="verify-guide-list">
                <div className="verify-guide-item">
                  <span className="verify-step-num">1</span>
                  <span>
                    Check your email inbox for a message from <strong>ExcelAgento</strong>.
                  </span>
                </div>
                <div className="verify-guide-item">
                  <span className="verify-step-num">2</span>
                  <span>
                    Click the <strong>Confirm your mail</strong> button in the email.
                  </span>
                </div>
                <div className="verify-guide-item">
                  <span className="verify-step-num">3</span>
                  <span>
                    Don't see it? Check your <strong>Spam or Junk</strong> folder.
                  </span>
                </div>
              </div>

              {/* Alternative: Enter 6-digit OTP code directly */}
              <form onSubmit={handleVerifyOtp} className="verify-otp-card">
                <div className="verify-otp-header">
                  <KeyRound size={15} />
                  <span>Have a 6-digit confirmation code?</span>
                </div>
                <div className="verify-otp-row">
                  <input
                    type="text"
                    maxLength={8}
                    className="auth-input verify-otp-input"
                    placeholder="Enter code"
                    value={otpCode}
                    onChange={(e) => setOtpCode(e.target.value)}
                  />
                  <button
                    type="submit"
                    className="btn btn-secondary btn-sm"
                    disabled={verifyingOtp || !otpCode.trim()}
                  >
                    {verifyingOtp ? <Loader2 size={14} className="auth-spin" /> : 'Verify'}
                  </button>
                </div>
              </form>

              {/* Resend button */}
              <div className="verify-actions" style={{ marginTop: '1.25rem' }}>
                <button
                  type="button"
                  className="auth-submit-btn"
                  onClick={handleResend}
                  disabled={resendCooldown > 0 || isResending}
                >
                  {isResending ? (
                    <>
                      <Loader2 size={15} className="auth-spin" /> Sending link…
                    </>
                  ) : resendCooldown > 0 ? (
                    <>
                      <RefreshCw size={14} /> Resend in {resendCooldown}s
                    </>
                  ) : (
                    <>
                      <RefreshCw size={14} /> Resend Verification Email
                    </>
                  )}
                </button>
                <button type="button" className="auth-link-btn" onClick={onSignIn}>
                  ← Back to Sign In
                </button>
              </div>
            </div>
          )}

          {/* STATE 4: ERROR / EXPIRED */}
          {status === 'error' && (
            <div className="verify-state-box error">
              <div className="verify-icon-circle error">
                <AlertCircle size={42} />
              </div>
              <span className="studio-eyebrow text-amber" style={{ letterSpacing: '0.08em' }}>
                LINK EXPIRED OR INVALID
              </span>
              <h1 id="verify-title" className="verify-title">
                Verification Link Expired
              </h1>
              <p className="verify-subtitle">
                Email verification links expire after 24 hours or after being clicked. If you
                already verified your account, you can sign in directly.
              </p>

              <div className="verify-input-group" style={{ margin: '1.25rem 0 1rem' }}>
                <label htmlFor="resend-email" className="auth-label">
                  Enter your email address for a fresh link:
                </label>
                <input
                  id="resend-email"
                  type="email"
                  className="auth-input"
                  placeholder="you@company.com"
                  value={targetEmail}
                  onChange={(e) => setTargetEmail(e.target.value)}
                />
              </div>

              <div className="verify-actions">
                <button
                  type="button"
                  className="auth-submit-btn"
                  onClick={handleResend}
                  disabled={resendCooldown > 0 || isResending || !targetEmail.trim()}
                >
                  {isResending ? (
                    <>
                      <Loader2 size={15} className="auth-spin" /> Sending fresh link…
                    </>
                  ) : resendCooldown > 0 ? (
                    <>
                      <RefreshCw size={14} /> Resend in {resendCooldown}s
                    </>
                  ) : (
                    <>
                      <RefreshCw size={14} /> Send Fresh Verification Link
                    </>
                  )}
                </button>
                <button type="button" className="auth-link-btn" onClick={onSignIn}>
                  Return to Sign In
                </button>
              </div>
            </div>
          )}

          {/* Support helpdesk footer */}
          <div className="verify-footer-help">
            <LifeBuoy size={14} />
            <span>Still having trouble?</span>
            {onOpenSupport ? (
              <button type="button" className="verify-help-link" onClick={onOpenSupport}>
                Contact Support Desk
              </button>
            ) : (
              <a href="#/support" className="verify-help-link">
                Contact Support Desk
              </a>
            )}
          </div>
        </section>
      </div>
    </div>
  );
};
