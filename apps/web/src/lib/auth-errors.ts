/** Translate transport failures without confusing them with rejected credentials. */
export function authErrorMessage(error: unknown, fallback: string): string {
  const message =
    typeof error === 'object' && error !== null && 'message' in error
      ? String(error.message)
      : fallback;
  if (typeof error === 'object' && error !== null) {
    const code = 'code' in error ? error.code : undefined;
    const status = 'status' in error ? error.status : undefined;
    if (code === 'over_email_send_rate_limit' || status === 429) {
      return 'Too many requests. Wait a minute before requesting another email, and use the newest link in your inbox.';
    }
    if (code === 'email_not_confirmed') {
      return 'Confirm your email before signing in. You can request a new link with “Resend confirmation email” below.';
    }
  }
  if (
    /failed to fetch|fetch failed|networkerror|network request failed|load failed/i.test(message)
  ) {
    return 'Could not reach the account service. Check your connection and retry. If this persists, the deployment may be blocking authentication requests.';
  }
  return message || fallback;
}
