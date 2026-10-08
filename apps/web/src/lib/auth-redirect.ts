export type AuthEmailKind = 'confirm' | 'recovery';

export interface AuthCallbackInfo {
  kind: AuthEmailKind;
  error: string | null;
}

export const DEFAULT_PRODUCTION_ORIGIN = 'https://excelagento.vercel.app';

function configuredAuthOrigin(): string {
  const configured =
    typeof import.meta !== 'undefined' &&
    typeof import.meta.env?.VITE_AUTH_REDIRECT_ORIGIN === 'string'
      ? import.meta.env.VITE_AUTH_REDIRECT_ORIGIN.trim()
      : '';
  if (configured) {
    try {
      return new URL(configured).origin;
    } catch {
      // ignore invalid URLs
    }
  }

  // If in browser and current origin is already a deployed public domain
  if (typeof window !== 'undefined' && window.location.origin) {
    const origin = window.location.origin;
    if (!origin.includes('localhost') && !origin.includes('127.0.0.1')) {
      return origin;
    }
  }

  // Always fallback to production origin so Supabase emails never link to localhost
  return DEFAULT_PRODUCTION_ORIGIN;
}

/** Use the canonical website origin for email verification & recovery.
 * Keep routing in the query: Supabase uses the fragment for session tokens. */
export function authEmailRedirectUrl(kind: AuthEmailKind, href?: string): string {
  const isDefaultHref =
    href === undefined || (typeof window !== 'undefined' && href === window.location.href);
  const currentHref =
    href ?? (typeof window !== 'undefined' ? window.location.href : DEFAULT_PRODUCTION_ORIGIN);
  const current = new URL(currentHref);

  const isTestMode = typeof import.meta !== 'undefined' && import.meta.env?.MODE === 'test';
  const configuredOrigin = configuredAuthOrigin();
  const shouldUseConfigured =
    isDefaultHref || current.origin.includes('localhost') || current.origin.includes('127.0.0.1');

  const url = !isTestMode && shouldUseConfigured ? new URL(configuredOrigin) : current;
  url.search = '';
  url.hash = '';
  if (kind === 'confirm' && (url.pathname === '/' || url.pathname === '')) {
    url.pathname = '/verify-email';
  }
  url.searchParams.set('auth', kind);
  return url.toString();
}

/** Capture this before the SDK starts initialization: it can consume the fragment. */
export function readAuthCallback(href: string): AuthCallbackInfo | null {
  const url = new URL(href);
  const fragment = new URLSearchParams(url.hash.slice(1));
  const marker = url.searchParams.get('auth');
  const hasVerifyPath = url.pathname.includes('verify-email') || url.pathname.includes('/verify');
  const type = fragment.get('type') ?? url.searchParams.get('type');
  const errorCode = fragment.get('error_code') ?? url.searchParams.get('error_code');
  const errorDescription =
    fragment.get('error_description') ?? url.searchParams.get('error_description');
  const callbackError = fragment.get('error') ?? url.searchParams.get('error');
  const hasSessionPayload =
    fragment.has('access_token') ||
    url.searchParams.has('code') ||
    url.searchParams.has('token_hash');
  if (
    marker !== 'confirm' &&
    marker !== 'recovery' &&
    !hasSessionPayload &&
    !callbackError &&
    !hasVerifyPath
  ) {
    return null;
  }
  const hasError = Boolean(callbackError || errorCode || errorDescription);
  const error = !hasError
    ? null
    : errorCode === 'otp_expired' || /expired|already.*used/i.test(errorDescription ?? '')
      ? 'This email link has expired or has already been used. Try signing in, or request a new email link.'
      : 'This email link could not be verified. Try signing in, or request a new email link.';
  return { kind: marker === 'recovery' || type === 'recovery' ? 'recovery' : 'confirm', error };
}

/** Only called after Supabase has finished reading the callback. Never persist tokens. */
export function clearAuthCallbackUrl(): void {
  const url = new URL(window.location.href);
  for (const key of ['auth', 'code', 'error', 'error_code', 'error_description', 'type']) {
    url.searchParams.delete(key);
  }
  url.hash = '';
  window.history.replaceState(window.history.state, '', url.pathname + url.search);
}
