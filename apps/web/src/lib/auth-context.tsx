import React, { createContext, useContext, useEffect, useState, useCallback, useRef } from 'react';
import type { Session, User } from '@supabase/supabase-js';
import { initialAuthCallback, supabase } from './supabase-client.js';
import { authErrorMessage } from './auth-errors.js';
import {
  authEmailRedirectUrl,
  clearAuthCallbackUrl,
  type AuthCallbackInfo,
} from './auth-redirect.js';

export interface UserProfile {
  id: string;
  email: string;
  full_name: string | null;
  avatar_url: string | null;
  created_at: string;
  updated_at: string;
}

interface AuthContextValue {
  user: User | null;
  session: Session | null;
  profile: UserProfile | null;
  loading: boolean;
  error: string | null;
  callback: AuthCallbackInfo | null;
  dismissCallback: () => void;
  signIn: (
    email: string,
    password: string,
  ) => Promise<{ success: boolean; authenticated?: boolean; error?: string }>;
  signUp: (
    email: string,
    password: string,
    fullName?: string,
  ) => Promise<{ success: boolean; authenticated?: boolean; error?: string }>;
  signOut: () => Promise<void>;
  resetPassword: (email: string) => Promise<{ success: boolean; error?: string }>;
  resendConfirmation: (email: string) => Promise<{ success: boolean; error?: string }>;
  updatePassword: (password: string) => Promise<{ success: boolean; error?: string }>;
  refreshProfile: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | null>(null);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [session, setSession] = useState<Session | null>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [callback, setCallback] = useState<AuthCallbackInfo | null>(initialAuthCallback);
  const profileRequest = useRef(0);
  const currentUserId = useRef<string | null>(null);

  const fetchProfile = useCallback(async (userId: string, userEmail?: string) => {
    const requestId = ++profileRequest.current;
    const isCurrent = () =>
      requestId === profileRequest.current && currentUserId.current === userId;
    try {
      const { data, error: profileErr } = await supabase
        .from('profiles')
        .select('*')
        .eq('id', userId)
        .maybeSingle();

      if (profileErr) {
        console.warn('Failed to load profile from Supabase:', profileErr.message);
        // Keep the workspace usable when the optional profile row is unavailable.
        if (!isCurrent()) return;
        setProfile({
          id: userId,
          email: userEmail || '',
          full_name: null,
          avatar_url: null,
          created_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        });
        return;
      }

      if (!isCurrent()) return;
      if (data) {
        setProfile(data as UserProfile);
      } else {
        // Construct default profile if row is being provisioned by trigger
        setProfile({
          id: userId,
          email: userEmail || '',
          full_name: null,
          avatar_url: null,
          created_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        });
      }
    } catch {
      // Keep the account session usable when the optional profile request is offline. Do not let a
      // late failure from an older account overwrite the current profile.
    }
  }, []);

  const refreshProfile = useCallback(async () => {
    if (user?.id) {
      await fetchProfile(user.id, user.email);
    }
  }, [user, fetchProfile]);

  useEffect(() => {
    let mounted = true;
    let authEventReceived = false;

    // Check active session on initial load
    supabase.auth
      .getSession()
      .then(({ data: { session: currentSession }, error: sessionError }) => {
        if (!mounted || authEventReceived) return;
        if (sessionError)
          setError(authErrorMessage(sessionError, 'The email link could not be verified.'));
        setSession(currentSession);
        setUser(currentSession?.user ?? null);
        currentUserId.current = currentSession?.user.id ?? null;
        setProfile(null);
        if (currentSession?.user) {
          void fetchProfile(currentSession.user.id, currentSession.user.email);
        }
        setLoading(false);
        if (initialAuthCallback) clearAuthCallbackUrl();
      })
      .catch((sessionError: unknown) => {
        if (!mounted) return;
        setError(authErrorMessage(sessionError, 'The account session could not be loaded.'));
        setLoading(false);
        if (initialAuthCallback) clearAuthCallbackUrl();
      });

    // Listen for auth state changes (login, logout, token refresh)
    const {
      data: { subscription: authListener },
    } = supabase.auth.onAuthStateChange((_event, newSession) => {
      if (!mounted) return;
      authEventReceived = true;
      setSession(newSession);
      setUser(newSession?.user ?? null);
      const changedUser = currentUserId.current !== (newSession?.user.id ?? null);
      currentUserId.current = newSession?.user.id ?? null;
      if (changedUser) setProfile(null);
      // Invalidate a profile request belonging to the previous account before starting the next one.
      profileRequest.current += 1;
      if (newSession?.user) {
        // Supabase invokes auth callbacks while holding its auth lock. Start API work after the
        // callback returns so acquiring a session for the profile query cannot deadlock that lock.
        const nextUser = newSession.user;
        setTimeout(() => {
          if (mounted && currentUserId.current === nextUser.id)
            void fetchProfile(nextUser.id, nextUser.email);
        }, 0);
      } else {
        setProfile(null);
      }
      setLoading(false);
    });

    return () => {
      mounted = false;
      profileRequest.current += 1;
      authListener.unsubscribe();
    };
  }, [fetchProfile]);

  const signIn = async (email: string, password: string) => {
    setError(null);
    try {
      const { data, error: authErr } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      });

      if (authErr) {
        const message = authErrorMessage(authErr, 'Failed to sign in');
        setError(message);
        return { success: false, error: message };
      }

      if (data.session) {
        currentUserId.current = data.session.user.id;
        setProfile(null);
        setUser(data.session.user);
        setSession(data.session);
        void fetchProfile(data.session.user.id, data.session.user.email);
      } else {
        return {
          success: false,
          error: 'No signed-in session was returned. Please try signing in again.',
        };
      }
      return { success: true, authenticated: Boolean(data.session) };
    } catch (err) {
      const msg = authErrorMessage(err, 'Failed to sign in');
      setError(msg);
      return { success: false, error: msg };
    }
  };

  const signUp = async (email: string, password: string, fullName?: string) => {
    setError(null);
    try {
      const { data, error: authErr } = await supabase.auth.signUp({
        email: email.trim(),
        password,
        options: {
          emailRedirectTo: authEmailRedirectUrl('confirm'),
          data: {
            full_name: fullName?.trim() || undefined,
          },
        },
      });

      if (authErr) {
        const message = authErrorMessage(authErr, 'Failed to create account');
        setError(message);
        return { success: false, error: message };
      }

      if (data.session) {
        currentUserId.current = data.session.user.id;
        setProfile(null);
        setUser(data.session.user);
        setSession(data.session);
        void fetchProfile(data.session.user.id, data.session.user.email);
      }
      return { success: true, authenticated: Boolean(data.session) };
    } catch (err) {
      const msg = authErrorMessage(err, 'Failed to create account');
      setError(msg);
      return { success: false, error: msg };
    }
  };

  const signOut = async () => {
    setError(null);
    try {
      await supabase.auth.signOut();
      profileRequest.current += 1;
      currentUserId.current = null;
      setUser(null);
      setSession(null);
      setProfile(null);
    } catch {
      profileRequest.current += 1;
      currentUserId.current = null;
      setUser(null);
      setSession(null);
      setProfile(null);
    }
  };

  const resetPassword = async (email: string) => {
    setError(null);
    try {
      const { error: resetErr } = await supabase.auth.resetPasswordForEmail(email.trim(), {
        redirectTo: authEmailRedirectUrl('recovery'),
      });
      if (resetErr) {
        const message = authErrorMessage(resetErr, 'Password reset failed');
        setError(message);
        return { success: false, error: message };
      }
      return { success: true };
    } catch (err) {
      const msg = authErrorMessage(err, 'Password reset failed');
      setError(msg);
      return { success: false, error: msg };
    }
  };

  const resendConfirmation = async (email: string) => {
    setError(null);
    try {
      const { error: resendError } = await supabase.auth.resend({
        type: 'signup',
        email: email.trim(),
        options: { emailRedirectTo: authEmailRedirectUrl('confirm') },
      });
      if (resendError) {
        const message = authErrorMessage(resendError, 'The confirmation email could not be sent.');
        setError(message);
        return { success: false, error: message };
      }
      return { success: true };
    } catch (resendError) {
      const message = authErrorMessage(resendError, 'The confirmation email could not be sent.');
      setError(message);
      return { success: false, error: message };
    }
  };

  const updatePassword = async (password: string) => {
    if (!session)
      return { success: false, error: 'Open a valid recovery link before setting a new password.' };
    try {
      const { error: updateError } = await supabase.auth.updateUser({ password });
      if (updateError)
        return {
          success: false,
          error: authErrorMessage(updateError, 'The password could not be updated.'),
        };
      return { success: true };
    } catch (updateError) {
      return {
        success: false,
        error: authErrorMessage(updateError, 'The password could not be updated.'),
      };
    }
  };

  return (
    <AuthContext.Provider
      value={{
        user,
        session,
        profile,
        loading,
        error,
        callback,
        dismissCallback: () => setCallback(null),
        signIn,
        signUp,
        signOut,
        resetPassword,
        resendConfirmation,
        updatePassword,
        refreshProfile,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export function useAuth(): AuthContextValue {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
}
