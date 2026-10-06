import { createClient } from '@supabase/supabase-js';
import { readAuthCallback } from './auth-redirect.js';

const SUPABASE_PROJECT_ID = 'fsepapdadtrlddkyqqxu';
const DEFAULT_SUPABASE_URL = `https://${SUPABASE_PROJECT_ID}.supabase.co`;
const DEFAULT_SUPABASE_ANON_KEY = 'sb_publishable_8JPAZFfCaS_U8nAAqc1rrQ_V6OSKAic';

export const SUPABASE_URL =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_URL) ||
  DEFAULT_SUPABASE_URL;

export const SUPABASE_ANON_KEY =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_ANON_KEY) ||
  DEFAULT_SUPABASE_ANON_KEY;

export const initialAuthCallback =
  typeof window === 'undefined' ? null : readAuthCallback(window.location.href);

/**
 * Shared Supabase Client configured for account authentication and profiles.
 * Agent reasoning and memory remain browser-local; model calls use the provider
 * selected by the user in workspace settings.
 */
export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});
