/**
 * @deprecated Device-level HWID / install ID has been removed in favor of simple email account licensing.
 * This stub is maintained only for legacy backwards-compatibility if needed.
 */

export function getRawHWID(): string {
  return 'standard_client';
}

export function getInstallId(): string {
  return '00000000-0000-4000-8000-000000000000';
}

export function getInstallIdHint(): string {
  return 'account-licensed';
}
