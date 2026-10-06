import { createClient } from '@supabase/supabase-js';

const SUPABASE_PROJECT_ID = 'fsepapdadtrlddkyqqxu';
const DEFAULT_SUPABASE_URL = `https://${SUPABASE_PROJECT_ID}.supabase.co`;
const DEFAULT_SUPABASE_ANON_KEY =
  'sb_publishable_8JPAZFfCaS_U8nAAqc1rrQ_V6OSKAic';

export const SUPABASE_URL =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_URL) ||
  DEFAULT_SUPABASE_URL;

export const SUPABASE_ANON_KEY =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_ANON_KEY) ||
  DEFAULT_SUPABASE_ANON_KEY;

/**
 * Shared Supabase Client configured exclusively for User Authentication,
 * Profiles, and future Subscription state management.
 * (Agent reasoning and memory remain 100% in-browser local storage at 0ms latency).
 */
export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});
