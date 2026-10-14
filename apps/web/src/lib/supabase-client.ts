import { createClient } from '@supabase/supabase-js';
import { readAuthCallback } from './auth-redirect.js';
import { resolveSupabaseClientConfig, runtimeConfig } from './runtime-config.js';

const SUPABASE_PROJECT_ID = 'fsepapdadtrlddkyqqxu';
const DEFAULT_SUPABASE_URL = `https://${SUPABASE_PROJECT_ID}.supabase.co`;
const DEFAULT_SUPABASE_ANON_KEY = 'sb_publishable_8JPAZFfCaS_U8nAAqc1rrQ_V6OSKAic';

const clientConfig = resolveSupabaseClientConfig(runtimeConfig, {
  url: DEFAULT_SUPABASE_URL,
  anonKey: DEFAULT_SUPABASE_ANON_KEY,
});

export const SUPABASE_URL = clientConfig.url;
export const SUPABASE_ANON_KEY = clientConfig.anonKey;

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
