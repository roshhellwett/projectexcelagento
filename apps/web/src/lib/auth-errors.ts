/** Translate transport failures without confusing them with rejected credentials. */
export function authErrorMessage(error: unknown, fallback: string): string {
  const message =
    typeof error === 'object' && error !== null && 'message' in error
      ? String(error.message)
      : fallback;
  if (
    /failed to fetch|fetch failed|networkerror|network request failed|load failed/i.test(message)
  ) {
    return 'Could not reach the account service. Check your connection and retry. If this persists, the deployment may be blocking authentication requests.';
  }
  return message || fallback;
}
