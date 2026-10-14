import React from 'react';

interface StartupConfigurationErrorProps {
  issues: readonly string[];
  onReload?: () => void;
}

/**
 * Safe, actionable fallback for a production build with missing public configuration. It avoids
 * rendering the auth/provider surfaces against an unknown backend and never echoes config values.
 */
export const StartupConfigurationError: React.FC<StartupConfigurationErrorProps> = ({
  issues,
  onReload = () => window.location.reload(),
}) => (
  <main
    role="alert"
    aria-live="assertive"
    data-testid="startup-configuration-error"
    style={{
      minHeight: '100vh',
      display: 'grid',
      placeItems: 'center',
      padding: '2rem',
      background: '#f8fafc',
      color: '#0f172a',
      fontFamily: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    }}
  >
    <section
      style={{
        width: 'min(100%, 42rem)',
        padding: '2rem',
        border: '1px solid #cbd5e1',
        borderRadius: '0.75rem',
        background: '#ffffff',
        boxShadow: '0 16px 40px rgba(15, 23, 42, 0.1)',
      }}
    >
      <p style={{ margin: 0, fontSize: '0.75rem', fontWeight: 700, letterSpacing: '0.12em' }}>
        EXCELAGENTO
      </p>
      <h1 style={{ margin: '0.75rem 0 0.5rem', fontSize: '1.5rem' }}>
        Deployment configuration needed
      </h1>
      <p style={{ margin: 0, color: '#475569', lineHeight: 1.6 }}>
        This production deployment is missing valid public configuration, so account access is
        paused until it is corrected. No workbook or provider data was loaded.
      </p>
      <ul
        style={{ margin: '1.25rem 0', paddingLeft: '1.25rem', color: '#334155', lineHeight: 1.7 }}
      >
        {issues.map((issue) => (
          <li key={issue}>{issue}</li>
        ))}
      </ul>
      <p style={{ margin: '0 0 1.25rem', color: '#64748b', fontSize: '0.9rem', lineHeight: 1.6 }}>
        Set these Vite build variables in the deployment environment, then redeploy. Keep keys in
        the deployment secret store; this screen does not display or request them.
      </p>
      <button
        type="button"
        onClick={onReload}
        style={{
          border: 0,
          borderRadius: '0.5rem',
          padding: '0.7rem 1rem',
          background: '#0f172a',
          color: '#ffffff',
          fontWeight: 700,
          cursor: 'pointer',
        }}
      >
        Reload after fixing configuration
      </button>
    </section>
  </main>
);
