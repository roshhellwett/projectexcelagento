export type AuthEmailKind = 'confirm' | 'recovery';

export interface AuthCallbackInfo {
  kind: AuthEmailKind;
  error: string | null;
}

function configuredAuthOrigin(): string | undefined {
  const configured =
    typeof import.meta !== 'undefined' &&
    typeof import.meta.env?.VITE_AUTH_REDIRECT_ORIGIN === 'string'
      ? import.meta.env.VITE_AUTH_REDIRECT_ORIGIN.trim()
      : '';
  if (!configured) return undefined;
  try {
    return new URL(configured).origin;
  } catch {
    return undefined;
  }
}

/** Use the app the person signed up on, including the dev port and any base path.
 * Keep routing in the query: Supabase uses the fragment for session tokens. */
export function authEmailRedirectUrl(kind: AuthEmailKind, href = window.location.href): string {
  const current = new URL(href);
  const configuredOrigin = configuredAuthOrigin();
  const url = configuredOrigin ? new URL(configuredOrigin) : current;
  url.search = '';
  url.hash = '';
  url.searchParams.set('auth', kind);
  return url.toString();
}

/** Capture this before the SDK starts initialization: it can consume the fragment. */
export function readAuthCallback(href: string): AuthCallbackInfo | null {
  const url = new URL(href);
  const fragment = new URLSearchParams(url.hash.slice(1));
  const marker = url.searchParams.get('auth');
  const type = fragment.get('type') ?? url.searchParams.get('type');
  const errorCode = fragment.get('error_code') ?? url.searchParams.get('error_code');
  const errorDescription =
    fragment.get('error_description') ?? url.searchParams.get('error_description');
  const callbackError = fragment.get('error') ?? url.searchParams.get('error');
  const hasSessionPayload = fragment.has('access_token') || url.searchParams.has('code');
  if (marker !== 'confirm' && marker !== 'recovery' && !hasSessionPayload && !callbackError) {
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
