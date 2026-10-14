export interface RuntimeEnvironment {
  MODE?: unknown;
  VITE_SUPABASE_URL?: unknown;
  VITE_SUPABASE_ANON_KEY?: unknown;
  VITE_AUTH_REDIRECT_ORIGIN?: unknown;
}

export interface RuntimeConfig {
  mode: string;
  isProduction: boolean;
  supabaseUrl: string;
  supabaseAnonKey: string;
  authRedirectOrigin: string;
  supabaseUrlValid: boolean;
  supabaseAnonKeyValid: boolean;
  issues: string[];
}

export interface SupabaseClientDefaults {
  url: string;
  anonKey: string;
}

export interface SupabaseClientConfig {
  url: string;
  anonKey: string;
  disabled: boolean;
}

/** A valid public Supabase URL must be an origin without credentials, paths, or URL fragments. */
function isValidSupabaseUrl(value: string, requiresExplicitConfig: boolean): boolean {
  if (!value) return false;
  try {
    const url = new URL(value);
    const protocolAllowed = requiresExplicitConfig
      ? url.protocol === 'https:'
      : url.protocol === 'https:' || url.protocol === 'http:';
    return (
      protocolAllowed &&
      Boolean(url.hostname) &&
      !url.username &&
      !url.password &&
      url.pathname === '/' &&
      !url.search &&
      !url.hash
    );
  } catch {
    return false;
  }
}

function isValidAuthRedirectOrigin(value: string, requiresExplicitConfig: boolean): boolean {
  if (!value) return false;
  try {
    const url = new URL(value);
    const protocolAllowed = requiresExplicitConfig
      ? url.protocol === 'https:'
      : url.protocol === 'https:' || url.protocol === 'http:';
    return (
      protocolAllowed &&
      Boolean(url.hostname) &&
      !url.username &&
      !url.password &&
      url.pathname === '/' &&
      !url.search &&
      !url.hash
    );
  } catch {
    return false;
  }
}

function stringValue(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function hasControlCharacters(value: string): boolean {
  return Array.from(value).some((character) => {
    const code = character.charCodeAt(0);
    return code <= 0x1f || code === 0x7f;
  });
}

const PUBLISHABLE_KEY_PREFIX = 'sb_publishable_';
const PUBLISHABLE_KEY_SUFFIX = /^[A-Za-z0-9_-]{16,}$/u;
const PLACEHOLDER_PUBLISHABLE_SUFFIX =
  /^(?:placeholder|example|test|fake|dummy|changeme|redacted|replace)(?:[-_].*)?$/iu;
const JWT_SEGMENT = /^[A-Za-z0-9_-]+$/u;

function decodeBase64Url(value: string): string | null {
  try {
    const base64 = value.replace(/-/g, '+').replace(/_/g, '/');
    if (base64.length % 4 === 1) return null;
    const padded = base64.padEnd(Math.ceil(base64.length / 4) * 4, '=');
    const binary = atob(padded);
    const bytes = Uint8Array.from(Array.from(binary, (character) => character.charCodeAt(0)));
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

function hasLegacyAnonRole(value: string): boolean {
  const parts = value.split('.');
  if (parts.length !== 3 || parts.some((part) => !JWT_SEGMENT.test(part))) return false;

  const payload = decodeBase64Url(parts[1]!);
  if (!payload) return false;

  try {
    const parsed: unknown = JSON.parse(payload);
    return (
      typeof parsed === 'object' &&
      parsed !== null &&
      !Array.isArray(parsed) &&
      (parsed as { role?: unknown }).role === 'anon'
    );
  } catch {
    return false;
  }
}

function isValidSupabaseAnonKey(value: string): boolean {
  if (!value || hasControlCharacters(value)) return false;

  if (value.startsWith(PUBLISHABLE_KEY_PREFIX)) {
    const suffix = value.slice(PUBLISHABLE_KEY_PREFIX.length);
    return PUBLISHABLE_KEY_SUFFIX.test(suffix) && !PLACEHOLDER_PUBLISHABLE_SUFFIX.test(suffix);
  }

  // Supabase secret keys and legacy JWTs for service_role/admin access must never be accepted by
  // the browser configuration boundary. Only a JWT whose decoded role is exactly anon is public.
  if (value.toLowerCase().startsWith('sb_secret_')) return false;
  return hasLegacyAnonRole(value);
}

const INVALID_ANON_KEY_MESSAGE =
  'VITE_SUPABASE_ANON_KEY must be a public Supabase publishable key or a legacy JWT anon key; private or placeholder keys are not allowed.';

function invalidAnonKeyMessage(value: string): string {
  return hasControlCharacters(value)
    ? 'VITE_SUPABASE_ANON_KEY contains invalid control characters.'
    : INVALID_ANON_KEY_MESSAGE;
}

function usesLocalFallbacks(mode: string): boolean {
  return mode === 'development' || mode === 'test';
}

/**
 * Validate build-time public configuration without ever including its values in an error.
 * Development and test builds intentionally keep the existing local fallbacks; every other build
 * mode must be explicit so an incorrectly configured deployment cannot silently target a default
 * account or generate email links for the wrong origin.
 */
export function validateRuntimeConfig(
  environment: RuntimeEnvironment,
  mode = stringValue(environment.MODE) || 'development',
): RuntimeConfig {
  const isProduction = mode === 'production';
  const requiresExplicitConfig = !usesLocalFallbacks(mode);
  const supabaseUrl = stringValue(environment.VITE_SUPABASE_URL);
  const supabaseAnonKey = stringValue(environment.VITE_SUPABASE_ANON_KEY);
  const authRedirectOrigin = stringValue(environment.VITE_AUTH_REDIRECT_ORIGIN);
  const supabaseUrlValid = isValidSupabaseUrl(supabaseUrl, requiresExplicitConfig);
  const supabaseAnonKeyValid = isValidSupabaseAnonKey(supabaseAnonKey);
  const issues: string[] = [];

  if (requiresExplicitConfig) {
    if (!supabaseUrl) {
      issues.push('VITE_SUPABASE_URL is required for production authentication.');
    } else if (!supabaseUrlValid) {
      issues.push('VITE_SUPABASE_URL must be a valid HTTPS URL without embedded credentials.');
    }

    if (!supabaseAnonKey) {
      issues.push('VITE_SUPABASE_ANON_KEY is required for production authentication.');
    } else if (!supabaseAnonKeyValid) {
      issues.push(invalidAnonKeyMessage(supabaseAnonKey));
    }

    if (!authRedirectOrigin) {
      issues.push('VITE_AUTH_REDIRECT_ORIGIN is required for production email links.');
    } else if (!isValidAuthRedirectOrigin(authRedirectOrigin, requiresExplicitConfig)) {
      issues.push('VITE_AUTH_REDIRECT_ORIGIN must be a valid HTTPS origin.');
    }
  } else {
    if (supabaseAnonKey && !supabaseAnonKeyValid) {
      issues.push(invalidAnonKeyMessage(supabaseAnonKey));
    }
    if (authRedirectOrigin && !isValidAuthRedirectOrigin(authRedirectOrigin, false)) {
      issues.push('VITE_AUTH_REDIRECT_ORIGIN must be a valid HTTP(S) origin when configured.');
    }
  }

  return {
    mode,
    isProduction,
    supabaseUrl,
    supabaseAnonKey,
    authRedirectOrigin,
    supabaseUrlValid,
    supabaseAnonKeyValid,
    issues,
  };
}

/**
 * Keep the SDK import safe when a production deployment is misconfigured. The startup gate stops
 * the UI before any request is made; this disabled client is only a last line of defence against
 * a module-level SDK constructor throwing or accidentally using a checked-in fallback.
 */
export function resolveSupabaseClientConfig(
  config: RuntimeConfig,
  defaults: SupabaseClientDefaults,
): SupabaseClientConfig {
  const disabled = !usesLocalFallbacks(config.mode) && config.issues.length > 0;
  if (disabled) {
    return {
      url: 'https://invalid.excelagento.invalid',
      anonKey: 'missing-production-configuration',
      disabled: true,
    };
  }
  return {
    url: config.supabaseUrlValid && config.supabaseUrl ? config.supabaseUrl : defaults.url,
    anonKey:
      config.supabaseAnonKeyValid && config.supabaseAnonKey
        ? config.supabaseAnonKey
        : defaults.anonKey,
    disabled: false,
  };
}

const viteEnvironment: RuntimeEnvironment =
  typeof import.meta !== 'undefined' && import.meta.env ? import.meta.env : {};

export const runtimeConfig = validateRuntimeConfig(viteEnvironment);
