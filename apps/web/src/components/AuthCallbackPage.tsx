import { useState } from 'react';
import { AlertCircle, ArrowRight, CheckCircle2, KeyRound, Loader2 } from 'lucide-react';
import { useAuth } from '../lib/auth-context.js';
import type { AuthCallbackInfo } from '../lib/auth-redirect.js';

interface AuthCallbackPageProps {
  callback: AuthCallbackInfo;
  onContinue: () => void;
  onSignIn: () => void;
  onRequestRecovery: () => void;
  onRequestVerification: () => void;
}

export function AuthCallbackPage({
  callback,
  onContinue,
  onSignIn,
  onRequestRecovery,
  onRequestVerification,
}: AuthCallbackPageProps) {
  const { session, loading, error: sessionError, updatePassword } = useAuth();
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [saving, setSaving] = useState(false);
  const [updated, setUpdated] = useState(false);
  const [formError, setFormError] = useState('');
  const [showPasswords, setShowPasswords] = useState(false);
  const callbackError = callback.error || sessionError;
  const failed = !loading && (Boolean(callbackError) || !session);
  const recovery = callback.kind === 'recovery';

  const savePassword = async (event: React.FormEvent) => {
    event.preventDefault();
    if (saving) return;
    setFormError('');
    if (password.length < 6) {
      setFormError('Use at least 6 characters for the new password.');
      return;
    }
    if (password !== confirmPassword) {
      setFormError('The passwords do not match.');
      return;
    }
    setSaving(true);
    const result = await updatePassword(password);
    setSaving(false);
    if (!result.success) {
      setFormError(result.error || 'The password could not be updated.');
      return;
    }
    setPassword('');
    setConfirmPassword('');
    setUpdated(true);
  };

  return (
    <main className="auth-page auth-callback-layout">
      <section className="auth-card auth-callback-card" aria-labelledby="auth-callback-title">
        <button
          type="button"
          className="auth-brand"
          disabled={loading || saving}
          onClick={onSignIn}
          aria-label="Back to sign in"
        >
          <span className="auth-brand-mark">
            <img src="/excel-agent-logo.svg" alt="" />
          </span>
          <span>
            <strong>
              Excel<span>Agento</span>
            </strong>
            <small>ACCOUNT ACCESS</small>
          </span>
        </button>
        <div className="auth-heading">
          <span className="auth-card-label">
            {recovery ? 'PASSWORD RECOVERY' : 'EMAIL CONFIRMATION'}
          </span>
          <h1 id="auth-callback-title">
            {loading
              ? 'Checking your email link'
              : failed
                ? 'Let’s get you back in'
                : updated
                  ? 'Password updated'
                  : recovery
                    ? 'Set a new password'
                    : 'Account session ready'}
          </h1>
          <p>
            {loading
              ? 'Finishing the account session before opening the workspace…'
              : failed
                ? callbackError ||
                  'No account session was established from this link. It may have already been used. Try signing in, or request a new email link.'
                : updated
                  ? 'Your new password has been saved. You can now continue to the workspace.'
                  : recovery
                    ? 'Choose the password you will use the next time you sign in.'
                    : 'You’re signed in. Continue to your spreadsheet workspace.'}
          </p>
        </div>

        {loading ? (
          <div className="auth-callback-status" role="status">
            <Loader2 size={18} className="auth-spin" /> Verifying account access
          </div>
        ) : failed ? (
          <div className="auth-callback-actions">
            <div className="auth-alert is-error" role="alert">
              <AlertCircle size={16} /> A valid session is required to continue.
            </div>
            <button className="auth-submit-btn" type="button" onClick={onSignIn}>
              Back to sign in <ArrowRight size={16} />
            </button>
            {recovery && (
              <button type="button" className="auth-reset-back" onClick={onRequestRecovery}>
                Request a new recovery link
              </button>
            )}
            {!recovery && (
              <button type="button" className="auth-reset-back" onClick={onRequestVerification}>
                Request a new confirmation link
              </button>
            )}
          </div>
        ) : recovery && !updated ? (
          <form className="auth-form" onSubmit={savePassword}>
            {formError && (
              <div className="auth-alert is-error" role="alert">
                <AlertCircle size={16} /> {formError}
              </div>
            )}
            <label className="auth-field">
              <span className="auth-label">New password</span>
              <span className="auth-input-wrap">
                <KeyRound size={16} />
                <input
                  type={showPasswords ? 'text' : 'password'}
                  autoComplete="new-password"
                  required
                  minLength={6}
                  disabled={saving}
                  value={password}
                  onChange={(event) => setPassword(event.target.value)}
                />
              </span>
            </label>
            <label className="auth-field">
              <span className="auth-label">Confirm new password</span>
              <span className="auth-input-wrap">
                <KeyRound size={16} />
                <input
                  type={showPasswords ? 'text' : 'password'}
                  autoComplete="new-password"
                  required
                  disabled={saving}
                  value={confirmPassword}
                  onChange={(event) => setConfirmPassword(event.target.value)}
                />
              </span>
            </label>
            <button
              type="button"
              className="auth-link-btn"
              disabled={saving}
              aria-pressed={showPasswords}
              onClick={() => setShowPasswords((visible) => !visible)}
            >
              {showPasswords ? 'Hide passwords' : 'Show passwords'}
            </button>
            <button className="auth-submit-btn" type="submit" disabled={saving}>
              {saving ? (
                <>
                  <Loader2 size={16} className="auth-spin" /> Saving password…
                </>
              ) : (
                <>
                  Save new password <ArrowRight size={16} />
                </>
              )}
            </button>
          </form>
        ) : (
          <div className="auth-callback-actions">
            <div className="auth-alert is-success" role="status">
              <CheckCircle2 size={16} /> Account access is ready.
            </div>
            <button className="auth-submit-btn" type="button" onClick={onContinue}>
              Continue to workspace <ArrowRight size={16} />
            </button>
          </div>
        )}
      </section>
    </main>
  );
}
