import React, { createContext, useCallback, useContext, useEffect, useState } from 'react';
import { useAuth } from './auth-context.js';
import {
  activateLicense,
  fetchLicenseStatus,
  heartbeatLicense,
  type LicenseStatus,
  LicenseServiceError,
} from './licensing.js';

interface LicenseContextValue {
  status: LicenseStatus | null;
  loading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
  activate: (key: string) => Promise<{ success: boolean; error?: string }>;
}

const LicenseContext = createContext<LicenseContextValue | null>(null);

export const LicenseProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { user, loading: authLoading } = useAuth();
  const [status, setStatus] = useState<LicenseStatus | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    if (!user) {
      setStatus(null);
      setError(null);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      setStatus(await fetchLicenseStatus());
    } catch (requestError) {
      const message =
        requestError instanceof LicenseServiceError
          ? requestError.message
          : 'The activation service could not verify this account.';
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [user]);

  useEffect(() => {
    if (authLoading) return;
    void refresh();
  }, [authLoading, refresh]);

  useEffect(() => {
    if (!user || !status) return undefined;
    const timer = window.setInterval(() => {
      void heartbeatLicense()
        .then((next) => {
          setStatus(next);
          setError(null);
        })
        .catch((requestError: unknown) => {
          setError(
            requestError instanceof LicenseServiceError
              ? requestError.message
              : 'The activation service could not re-verify this account.',
          );
        });
    }, 60_000);
    return () => window.clearInterval(timer);
  }, [status, user]);

  const activate = useCallback(
    async (key: string) => {
      try {
        await activateLicense(key);
        await refresh();
        setError(null);
        return { success: true };
      } catch (requestError) {
        const message =
          requestError instanceof LicenseServiceError
            ? requestError.message
            : 'The activation key could not be verified.';
        setError(message);
        return { success: false, error: message };
      }
    },
    [refresh],
  );

  return (
    <LicenseContext.Provider value={{ status, loading, error, refresh, activate }}>
      {children}
    </LicenseContext.Provider>
  );
};

export function useLicense(): LicenseContextValue {
  const context = useContext(LicenseContext);
  if (!context) throw new Error('useLicense must be used within a LicenseProvider');
  return context;
}
