import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import type { Session, User } from '@supabase/supabase-js';
import { supabase } from './supabase-client.js';
import { authErrorMessage } from './auth-errors.js';

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
  refreshProfile: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | null>(null);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [session, setSession] = useState<Session | null>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const fetchProfile = useCallback(async (userId: string, userEmail?: string) => {
    try {
      const { data, error: profileErr } = await supabase
        .from('profiles')
        .select('*')
        .eq('id', userId)
        .maybeSingle();

      if (profileErr) {
        console.warn('Failed to load profile from Supabase:', profileErr.message);
        // Keep the workspace usable when the optional profile row is unavailable.
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
      // Ignore network errors on offline startup
    }
  }, []);

  const refreshProfile = useCallback(async () => {
    if (user?.id) {
      await fetchProfile(user.id, user.email);
    }
  }, [user, fetchProfile]);

  useEffect(() => {
    let mounted = true;

    // Check active session on initial load
    supabase.auth
      .getSession()
      .then(({ data: { session: currentSession } }) => {
        if (!mounted) return;
        setSession(currentSession);
        setUser(currentSession?.user ?? null);
        if (currentSession?.user) {
          void fetchProfile(currentSession.user.id, currentSession.user.email);
        }
        setLoading(false);
      })
      .catch(() => {
        if (mounted) setLoading(false);
      });

    // Listen for auth state changes (login, logout, token refresh)
    const {
      data: { subscription: authListener },
    } = supabase.auth.onAuthStateChange(async (_event, newSession) => {
      if (!mounted) return;
      setSession(newSession);
      setUser(newSession?.user ?? null);
      if (newSession?.user) {
        void fetchProfile(newSession.user.id, newSession.user.email);
      } else {
        setProfile(null);
      }
      setLoading(false);
    });

    return () => {
      mounted = false;
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
        setUser(data.session.user);
        setSession(data.session);
        void fetchProfile(data.session.user.id, data.session.user.email);
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
      setUser(null);
      setSession(null);
      setProfile(null);
    } catch {
      setUser(null);
      setSession(null);
      setProfile(null);
    }
  };

  const resetPassword = async (email: string) => {
    setError(null);
    try {
      const { error: resetErr } = await supabase.auth.resetPasswordForEmail(email.trim());
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

  return (
    <AuthContext.Provider
      value={{
        user,
        session,
        profile,
        loading,
        error,
        signIn,
        signUp,
        signOut,
        resetPassword,
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
