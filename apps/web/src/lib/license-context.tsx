import React, { createContext, useCallback, useContext, useEffect, useState, useRef } from 'react';
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
  const [verified, setVerified] = useState<{ userId: string; status: LicenseStatus } | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const requestSequence = useRef(0);
  const currentUserId = useRef(user?.id);
  currentUserId.current = user?.id;
  const userId = user?.id;
  const status = verified && verified.userId === userId ? verified.status : null;

  const refresh = useCallback(async () => {
    if (!userId) {
      requestSequence.current += 1;
      setVerified(null);
      setError(null);
      setLoading(false);
      return;
    }
    const requestId = ++requestSequence.current;
    const isCurrent = () =>
      requestId === requestSequence.current && userId === currentUserId.current;
    setLoading(true);
    setError(null);
    try {
      const next = await fetchLicenseStatus();
      if (!isCurrent()) return;
      setVerified({ userId, status: next });
    } catch (requestError) {
      if (!isCurrent()) return;
      const message =
        requestError instanceof LicenseServiceError
          ? requestError.message
          : 'The activation service could not verify this account.';
      setError(message);
    } finally {
      if (isCurrent()) setLoading(false);
    }
  }, [userId]);

  useEffect(() => {
    if (authLoading) return;
    void refresh();
    return () => {
      requestSequence.current += 1;
    };
  }, [authLoading, refresh]);

  useEffect(() => {
    if (!userId) return undefined;
    let active = true;
    let pending = false;
    const timer = window.setInterval(() => {
      // Never overlap verification with activation or an earlier heartbeat.
      if (pending || loading) return;
      pending = true;
      const sequence = requestSequence.current;
      const isCurrent = () =>
        active && sequence === requestSequence.current && userId === currentUserId.current;
      void heartbeatLicense()
        .then((next) => {
          if (!isCurrent()) return;
          setVerified({ userId, status: next });
          setError(null);
        })
        .catch((requestError: unknown) => {
          if (!isCurrent()) return;
          setError(
            requestError instanceof LicenseServiceError
              ? requestError.message
              : 'The activation service could not re-verify this account.',
          );
        })
        .finally(() => {
          pending = false;
        });
    }, 60_000);
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, [userId, loading]);

  const activate = useCallback(
    async (key: string) => {
      if (!userId) return { success: false, error: 'Sign in before activating a license.' };
      const requestId = ++requestSequence.current;
      const isCurrent = () =>
        requestId === requestSequence.current && userId === currentUserId.current;
      setLoading(true);
      try {
        const next = await activateLicense(key);
        if (!isCurrent())
          return { success: false, error: 'The account changed. Verify access again.' };
        setVerified({ userId, status: next });
        if (!next.canUse) {
          const message =
            next.banReason || 'This account or installation does not have active access.';
          setError(message);
          return { success: false, error: message };
        }
        setError(null);
        return { success: true };
      } catch (requestError) {
        const message =
          requestError instanceof LicenseServiceError
            ? requestError.message
            : 'The activation key could not be verified.';
        if (isCurrent()) setError(message);
        return { success: false, error: message };
      } finally {
        if (isCurrent()) setLoading(false);
      }
    },
    [userId],
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
