import React, { useEffect, useState } from 'react';
import { AnimatePresence, motion, useReducedMotion } from 'framer-motion';
import {
  AlertCircle,
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  Eye,
  EyeOff,
  KeyRound,
  Loader2,
  LockKeyhole,
  Mail,
  UserRound,
} from 'lucide-react';
import { useAuth } from '../lib/auth-context.js';

interface AuthPageProps {
  initialMode?: 'signin' | 'signup' | 'reset' | 'verify';
  onSuccess?: () => void;
  onNavigateBack?: () => void;
}

type AuthMode = 'signin' | 'signup' | 'reset' | 'verify';

export const AuthPage: React.FC<AuthPageProps> = ({
  initialMode = 'signin',
  onSuccess,
  onNavigateBack,
}) => {
  const { signIn, signUp, resetPassword, resendConfirmation } = useAuth();
  const reducedMotion = useReducedMotion();
  const [mode, setMode] = useState<AuthMode>(initialMode);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [fullName, setFullName] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [resendWait, setResendWait] = useState(0);

  useEffect(() => {
    setMode(initialMode);
    setError(null);
    setSuccessMessage(null);
  }, [initialMode]);

  useEffect(() => {
    if (resendWait === 0) return;
    const timer = window.setTimeout(
      () => setResendWait((seconds) => Math.max(0, seconds - 1)),
      1000,
    );
    return () => window.clearTimeout(timer);
  }, [resendWait]);

  const switchMode = (nextMode: AuthMode) => {
    if (loading) return;
    setMode(nextMode);
    setShowPassword(false);
    setError(null);
    setSuccessMessage(null);
  };

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (loading) return;
    setError(null);
    setSuccessMessage(null);

    const normalizedEmail = email.trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
      setError('Enter a valid email address to continue.');
      return;
    }

    if (mode === 'reset') {
      if (resendWait > 0) return;
      setLoading(true);
      const result = await resetPassword(normalizedEmail);
      setLoading(false);
      if (result.success) {
        setResendWait(60);
        setSuccessMessage(
          'If an account exists for this email, a recovery link has been requested. Check your inbox and spam folder.',
        );
      } else {
        setError(result.error || 'We could not send recovery instructions.');
      }
      return;
    }

    if (mode === 'verify') {
      await handleResendConfirmation();
      return;
    }

    if (!password || password.length < 6) {
      setError('Use a password with at least 6 characters.');
      return;
    }

    setLoading(true);
    const result =
      mode === 'signup'
        ? await signUp(normalizedEmail, password, fullName)
        : await signIn(normalizedEmail, password);
    setLoading(false);

    if (!result.success) {
      setError(
        result.error ||
          (mode === 'signup' ? 'We could not create your account.' : 'We could not sign you in.'),
      );
      return;
    }

    if (mode === 'signup' && !result.authenticated) {
      setMode('verify');
      setResendWait(60);
      setPassword('');
      setShowPassword(false);
      setSuccessMessage(
        'Account created. Check your email to confirm the account, then sign in to enter the workspace.',
      );
      return;
    }

    onSuccess?.();
  };

  const handleResendConfirmation = async () => {
    if (loading || resendWait > 0) return;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      setError('Enter your account email address above, then resend the confirmation link.');
      return;
    }
    setError(null);
    setSuccessMessage(null);
    setLoading(true);
    const result = await resendConfirmation(email);
    setLoading(false);
    if (!result.success) {
      setError(result.error || 'The confirmation email could not be sent.');
    } else {
      setResendWait(60);
      setSuccessMessage(
        'If this account needs confirmation, a new email link has been requested. Check your inbox and spam folder, and use the newest email.',
      );
    }
  };

  const title =
    mode === 'signin'
      ? 'Continue to your workspace'
      : mode === 'signup'
        ? 'Create your workspace account'
        : mode === 'verify'
          ? 'Confirm your email address'
          : 'Recover your account';
  const subtitle =
    mode === 'signin'
      ? 'Sign in to load workbooks, ask for changes, and review every proposed operation.'
      : mode === 'signup'
        ? 'Create an account to enter the open-source ExcelAgento workspace.'
        : mode === 'verify'
          ? 'Already registered? Request a fresh confirmation link using your account email.'
          : 'Enter your account email to request a password recovery link.';

  return (
    <div className="auth-page">
      <div className="auth-page-grid" aria-hidden="true" />
      <div className="auth-layout">
        <aside className="auth-context">
          <button
            type="button"
            className="auth-brand"
            onClick={onNavigateBack}
            aria-label="Back to ExcelAgento overview"
          >
            <span className="auth-brand-mark">
              <img src="/excel-agent-logo.svg" alt="" />
            </span>
            <span>
              <strong>
                Excel<span>Agento</span>
              </strong>
              <small>THE AGENT WORKSPACE</small>
            </span>
          </button>

          <div className="auth-context-copy">
            <span className="auth-kicker">WORKSPACE ACCESS</span>
            <h1>Move from spreadsheet instructions to verified work.</h1>
            <p>
              ExcelAgento keeps the workbook visible, the agent’s proposal reviewable, and every
              committed change reversible.
            </p>
          </div>

          <div className="auth-context-list">
            <div>
              <span className="auth-context-index">01</span>
              <p>Load an .xlsx, .xls, or .csv workbook in the browser.</p>
            </div>
            <div>
              <span className="auth-context-index">02</span>
              <p>Describe a cleanup, transformation, or analysis in plain English.</p>
            </div>
            <div>
              <span className="auth-context-index">03</span>
              <p>Inspect the preview before anything is applied to the sheet.</p>
            </div>
          </div>

          <p className="auth-context-footnote">
            Open source · browser-first execution · optional bring-your-own-key model access
          </p>
        </aside>

        <main className="auth-panel">
          {onNavigateBack && (
            <button type="button" className="auth-back-btn" onClick={onNavigateBack}>
              <ArrowLeft size={15} />
              Back to overview
            </button>
          )}

          <motion.section
            className="auth-card"
            initial={reducedMotion ? false : { opacity: 0, y: 14 }}
            animate={reducedMotion ? undefined : { opacity: 1, y: 0 }}
            transition={{ duration: 0.35, ease: [0.2, 0.7, 0.3, 1] }}
          >
            <div className="auth-card-header">
              <div className="auth-card-mark">
                <LockKeyhole size={17} />
              </div>
              <div>
                <span className="auth-card-label">ACCOUNT</span>
                <p>Required for workspace access</p>
              </div>
            </div>

            <div className="auth-heading">
              <h2>{title}</h2>
              <p>{subtitle}</p>
            </div>

            {(mode === 'signin' || mode === 'signup') && (
              <div className="auth-tabs" role="tablist" aria-label="Account action">
                <button
                  type="button"
                  role="tab"
                  aria-selected={mode === 'signin'}
                  className={mode === 'signin' ? 'is-active' : ''}
                  disabled={loading}
                  onClick={() => switchMode('signin')}
                >
                  Sign in
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={mode === 'signup'}
                  className={mode === 'signup' ? 'is-active' : ''}
                  disabled={loading}
                  onClick={() => switchMode('signup')}
                >
                  Create account
                </button>
              </div>
            )}

            <AnimatePresence initial={false} mode="wait">
              {(error || successMessage) && (
                <motion.div
                  key={error ? 'error' : 'success'}
                  className={`auth-alert ${error ? 'is-error' : 'is-success'}`}
                  role={error ? 'alert' : 'status'}
                  initial={reducedMotion ? false : { opacity: 0, height: 0 }}
                  animate={reducedMotion ? undefined : { opacity: 1, height: 'auto' }}
                  exit={reducedMotion ? undefined : { opacity: 0, height: 0 }}
                >
                  {error ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
                  <span>{error || successMessage}</span>
                </motion.div>
              )}
            </AnimatePresence>

            <form onSubmit={handleSubmit} className="auth-form" noValidate>
              {mode === 'signup' && (
                <label className="auth-field">
                  <span className="auth-label">
                    Name <em>Optional</em>
                  </span>
                  <span className="auth-input-wrap">
                    <UserRound size={16} aria-hidden="true" />
                    <input
                      id="auth-name"
                      type="text"
                      placeholder="How should we address you?"
                      value={fullName}
                      onChange={(event) => setFullName(event.target.value)}
                      autoComplete="name"
                      disabled={loading}
                    />
                  </span>
                </label>
              )}

              <label className="auth-field">
                <span className="auth-label">Email address</span>
                <span className="auth-input-wrap">
                  <Mail size={16} aria-hidden="true" />
                  <input
                    id="auth-email"
                    type="email"
                    placeholder="you@example.com"
                    value={email}
                    onChange={(event) => setEmail(event.target.value)}
                    autoComplete="email"
                    required
                    disabled={loading}
                  />
                </span>
              </label>

              {(mode === 'signin' || mode === 'signup') && (
                <div className="auth-field">
                  <span className="auth-label-row">
                    <label className="auth-label" htmlFor="auth-password">
                      Password
                    </label>
                    {mode === 'signin' && (
                      <button
                        type="button"
                        className="auth-link-btn"
                        disabled={loading}
                        onClick={() => switchMode('reset')}
                      >
                        Forgot password?
                      </button>
                    )}
                  </span>
                  <span className="auth-input-wrap">
                    <KeyRound size={16} aria-hidden="true" />
                    <input
                      id="auth-password"
                      type={showPassword ? 'text' : 'password'}
                      placeholder="At least 6 characters"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      autoComplete={mode === 'signup' ? 'new-password' : 'current-password'}
                      required
                      disabled={loading}
                    />
                    <button
                      type="button"
                      className="auth-eye-btn"
                      onClick={() => setShowPassword((visible) => !visible)}
                      aria-label={showPassword ? 'Hide password' : 'Show password'}
                      disabled={loading}
                    >
                      {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </span>
                  {mode === 'signup' && (
                    <span className="auth-hint">Use at least 6 characters.</span>
                  )}
                </div>
              )}

              <button
                type="submit"
                className="auth-submit-btn"
                disabled={loading || ((mode === 'reset' || mode === 'verify') && resendWait > 0)}
              >
                {loading ? (
                  <>
                    <Loader2 size={16} className="auth-spin" />
                    Working…
                  </>
                ) : mode === 'signin' ? (
                  <>
                    Enter workspace <ArrowRight size={16} />
                  </>
                ) : mode === 'signup' ? (
                  <>
                    Create account <ArrowRight size={16} />
                  </>
                ) : resendWait > 0 ? (
                  <>Send another link in {resendWait}s</>
                ) : mode === 'verify' ? (
                  <>
                    Resend confirmation email <ArrowRight size={16} />
                  </>
                ) : (
                  <>
                    Send recovery link <ArrowRight size={16} />
                  </>
                )}
              </button>
            </form>

            {mode === 'signin' && (
              <div className="auth-confirmation-actions">
                <button
                  type="button"
                  className="auth-link-btn"
                  disabled={loading}
                  onClick={() => switchMode('verify')}
                >
                  Resend confirmation email
                </button>
              </div>
            )}

            {(mode === 'reset' || mode === 'verify') && (
              <button
                type="button"
                className="auth-reset-back"
                disabled={loading}
                onClick={() => switchMode('signin')}
              >
                <ArrowLeft size={14} /> Back to sign in
              </button>
            )}

            <div className="auth-data-note">
              <span className="auth-data-note-icon">
                <LockKeyhole size={14} />
              </span>
              <p>
                Workbook files and local history stay in this browser. If you connect a model
                provider later, the workspace shows what context may be sent before the request
                runs.
              </p>
            </div>
          </motion.section>

          <p className="auth-panel-footer">
            By continuing, you agree to use this open-source workspace with your own data and model
            provider configuration.
          </p>
        </main>
      </div>
    </div>
  );
};
