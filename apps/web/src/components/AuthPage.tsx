import React, { useState } from 'react';
import {
  ShieldCheck,
  Mail,
  Lock,
  User,
  ArrowRight,
  Eye,
  EyeOff,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Sparkles,
  ArrowLeft,
  KeyRound,
} from 'lucide-react';
import { useAuth } from '../lib/auth-context.js';

interface AuthPageProps {
  initialMode?: 'signin' | 'signup';
  onSuccess?: () => void;
  onNavigateBack?: () => void;
}

export const AuthPage: React.FC<AuthPageProps> = ({
  initialMode = 'signin',
  onSuccess,
  onNavigateBack,
}) => {
  const { signIn, signUp, resetPassword } = useAuth();
  const [mode, setMode] = useState<'signin' | 'signup' | 'reset'>(initialMode);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [fullName, setFullName] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setSuccessMessage(null);

    if (!email.trim() || !email.includes('@')) {
      setError('Please provide a valid email address.');
      return;
    }

    if (mode === 'reset') {
      setLoading(true);
      const res = await resetPassword(email);
      setLoading(false);
      if (res.success) {
        setSuccessMessage('Password reset instructions sent. Please check your inbox.');
      } else {
        setError(res.error || 'Failed to send reset link.');
      }
      return;
    }

    if (!password || password.length < 6) {
      setError('Password must contain at least 6 characters.');
      return;
    }

    setLoading(true);
    if (mode === 'signup') {
      const res = await signUp(email, password, fullName);
      setLoading(false);
      if (res.success) {
        setSuccessMessage('Account created successfully! Welcome to ExcelAgento.');
        setTimeout(() => {
          onSuccess?.();
        }, 800);
      } else {
        setError(res.error || 'Failed to create account.');
      }
    } else {
      const res = await signIn(email, password);
      setLoading(false);
      if (res.success) {
        setSuccessMessage('Signed in successfully! Loading workspace…');
        setTimeout(() => {
          onSuccess?.();
        }, 500);
      } else {
        setError(res.error || 'Invalid email or password.');
      }
    }
  };

  return (
    <div className="auth-container">
      <div className="auth-shell">
        {onNavigateBack && (
          <button
            type="button"
            className="auth-back-btn"
            onClick={onNavigateBack}
            aria-label="Back to previous page"
          >
            <ArrowLeft size={16} />
            <span>Back to Workspace</span>
          </button>
        )}

        <div className="auth-card">
          {/* Brand Header */}
          <div className="auth-header">
            <div className="auth-logo-badge">
              <span className="auth-logo-icon">Σ</span>
              <span className="auth-logo-text">ExcelAgento</span>
              <span className="auth-version-pill">v0.2</span>
            </div>
            <h1 className="auth-title">
              {mode === 'signin' && 'Sign in to your account'}
              {mode === 'signup' && 'Create your ExcelAgento account'}
              {mode === 'reset' && 'Reset your password'}
            </h1>
            <p className="auth-subtitle">
              {mode === 'signin' && 'Access your autonomous spreadsheet workspace & agent history.'}
              {mode === 'signup' && 'Instant access to 6 autonomous agents, 100% private local computing.'}
              {mode === 'reset' && 'Enter your verified email to receive a recovery link.'}
            </p>
          </div>

          {/* Mode Switcher Tabs */}
          {mode !== 'reset' && (
            <div className="auth-tabs" role="tablist">
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'signin'}
                className={`auth-tab ${mode === 'signin' ? 'is-active' : ''}`}
                onClick={() => {
                  setMode('signin');
                  setError(null);
                  setSuccessMessage(null);
                }}
              >
                Sign In
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'signup'}
                className={`auth-tab ${mode === 'signup' ? 'is-active' : ''}`}
                onClick={() => {
                  setMode('signup');
                  setError(null);
                  setSuccessMessage(null);
                }}
              >
                Create Account
              </button>
            </div>
          )}

          {/* Feedback Messages */}
          {error && (
            <div className="auth-alert is-error" role="alert">
              <AlertCircle size={16} className="auth-alert-icon" />
              <span>{error}</span>
            </div>
          )}

          {successMessage && (
            <div className="auth-alert is-success" role="status">
              <CheckCircle2 size={16} className="auth-alert-icon" />
              <span>{successMessage}</span>
            </div>
          )}

          {/* Form */}
          <form onSubmit={handleSubmit} className="auth-form" noValidate>
            {mode === 'signup' && (
              <div className="auth-field">
                <label htmlFor="auth-name" className="auth-label">
                  Full Name
                </label>
                <div className="auth-input-wrapper">
                  <User size={16} className="auth-input-icon" />
                  <input
                    id="auth-name"
                    type="text"
                    className="auth-input"
                    placeholder="Ada Lovelace"
                    value={fullName}
                    onChange={(e) => setFullName(e.target.value)}
                    autoComplete="name"
                    disabled={loading}
                  />
                </div>
              </div>
            )}

            <div className="auth-field">
              <label htmlFor="auth-email" className="auth-label">
                Email Address
              </label>
              <div className="auth-input-wrapper">
                <Mail size={16} className="auth-input-icon" />
                <input
                  id="auth-email"
                  type="email"
                  className="auth-input"
                  placeholder="name@company.com"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  autoComplete="email"
                  required
                  disabled={loading}
                />
              </div>
            </div>

            {mode !== 'reset' && (
              <div className="auth-field">
                <div className="auth-label-row">
                  <label htmlFor="auth-password" className="auth-label">
                    Password
                  </label>
                  {mode === 'signin' && (
                    <button
                      type="button"
                      className="auth-link-btn"
                      onClick={() => {
                        setMode('reset');
                        setError(null);
                        setSuccessMessage(null);
                      }}
                    >
                      Forgot password?
                    </button>
                  )}
                </div>
                <div className="auth-input-wrapper">
                  <Lock size={16} className="auth-input-icon" />
                  <input
                    id="auth-password"
                    type={showPassword ? 'text' : 'password'}
                    className="auth-input"
                    placeholder="••••••••••••"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    autoComplete={mode === 'signup' ? 'new-password' : 'current-password'}
                    required
                    disabled={loading}
                  />
                  <button
                    type="button"
                    className="auth-eye-btn"
                    onClick={() => setShowPassword(!showPassword)}
                    title={showPassword ? 'Hide password' : 'Show password'}
                    aria-label={showPassword ? 'Hide password' : 'Show password'}
                  >
                    {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                  </button>
                </div>
                {mode === 'signup' && (
                  <span className="auth-hint">Must be at least 6 characters.</span>
                )}
              </div>
            )}

            <button
              type="submit"
              className="auth-submit-btn"
              disabled={loading}
            >
              {loading ? (
                <>
                  <Loader2 size={16} className="auth-spin" />
                  <span>Processing…</span>
                </>
              ) : mode === 'signin' ? (
                <>
                  <span>Sign In</span>
                  <ArrowRight size={16} />
                </>
              ) : mode === 'signup' ? (
                <>
                  <Sparkles size={16} />
                  <span>Create Free Account</span>
                </>
              ) : (
                <>
                  <KeyRound size={16} />
                  <span>Send Recovery Instructions</span>
                </>
              )}
            </button>
          </form>

          {mode === 'reset' && (
            <div className="auth-footer-link">
              <button
                type="button"
                className="auth-link-btn"
                onClick={() => {
                  setMode('signin');
                  setError(null);
                  setSuccessMessage(null);
                }}
              >
                Back to Sign In
              </button>
            </div>
          )}

          {/* Privacy & Tier Guarantee */}
          <div className="auth-guarantee">
            <ShieldCheck size={16} className="auth-guarantee-icon" />
            <div className="auth-guarantee-text">
              <strong>100% In-Browser Privacy:</strong> Your spreadsheets, formulas, and cells are
              computed strictly inside your browser. No file data is ever stored on external cloud
              servers.
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
