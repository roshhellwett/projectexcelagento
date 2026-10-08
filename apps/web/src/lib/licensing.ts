import { supabase } from './supabase-client.js';
import { getInstallId } from './device-identity.js';

export type LicenseState =
  | 'trial'
  | 'licensed'
  | 'activation_required'
  | 'expired'
  | 'banned'
  | 'device_banned'
  | 'device_mismatch'
  | 'email_unconfirmed'
  | 'revoked';

export interface LicenseEntitlement {
  id: string;
  keyHint: string;
  rawKey?: string | null;
  status: 'active' | 'expired' | 'revoked' | 'unused';
  activatedAt: string | null;
  expiresAt: string | null;
  daysRemaining: number;
}

export interface LicenseStatus {
  state: LicenseState;
  canUse: boolean;
  isAdmin: boolean;
  isOwner?: boolean;
  isLifetime?: boolean;
  role?: 'owner' | 'admin' | null;
  email: string;
  deviceId: string | null;
  deviceHint: string | null;
  trialStartedAt: string | null;
  trialExpiresAt: string | null;
  daysRemaining: number;
  license: LicenseEntitlement | null;
  banReason: string | null;
  serverNow?: string;
  verifyUntil?: string;
}

export interface AdminLicenseRecord {
  id: string;
  key_hint: string;
  raw_key?: string | null;
  status: 'unused' | 'active' | 'expired' | 'revoked';
  duration_days: number;
  days_remaining?: number;
  bound_user_id: string | null;
  bound_email: string | null;
  bound_device_id: string | null;
  activated_at: string | null;
  expires_at: string | null;
  revoked_at: string | null;
  created_at: string;
}

export interface AdminAccountRecord {
  user_id: string;
  email: string;
  status: 'owner' | 'admin' | 'trial' | 'licensed' | 'expired' | 'banned';
  role?: 'owner' | 'admin' | null;
  is_admin?: boolean;
  trial_started_at: string;
  trial_expires_at: string | null;
  days_remaining?: number | null;
  ban_reason: string | null;
  banned_at: string | null;
  created_at: string;
  updated_at?: string;
  active_device_id?: string | null;
}

export interface SupportTicketRecord {
  id: string;
  user_id: string | null;
  name: string;
  email: string;
  category: string;
  subject: string;
  message: string;
  status: 'pending' | 'in_progress' | 'resolved';
  admin_notes: string | null;
  replied_at: string | null;
  created_at: string;
  updated_at: string;
}

export interface AdminStats {
  totalUsers: number;
  activeLicenses: number;
  unusedKeys: number;
  pendingTickets: number;
  totalTickets: number;
}

export interface AdminDeviceRecord {
  id: string;
  user_id: string;
  install_id_hint: string;
  status: 'active' | 'retired' | 'banned';
  ban_reason: string | null;
  label: string | null;
  first_seen_at: string;
  last_seen_at: string;
}

export interface LicenseEventRecord {
  id: string;
  actor_user_id: string | null;
  target_user_id: string | null;
  license_id: string | null;
  device_id: string | null;
  event_type: string;
  metadata: Record<string, unknown>;
  created_at: string;
}

export interface AdminLicenseData {
  keys: AdminLicenseRecord[];
  accounts: AdminAccountRecord[];
  devices: AdminDeviceRecord[];
  events: LicenseEventRecord[];
  tickets?: SupportTicketRecord[];
  stats?: AdminStats;
}

/**
 * Resolves a real key only when it is available in the current generation response or database
 * record. A masked key hint is never returned as plaintext because copying it would produce an
 * activation value that can never work.
 */
export function resolvePlaintextLicenseKey(
  license: Pick<AdminLicenseRecord, 'raw_key' | 'key_hint'>,
  generatedKeys: readonly string[],
): string | undefined {
  if (license.raw_key) return license.raw_key;
  const suffix = license.key_hint
    .replace(/[^A-Z0-9]/gi, '')
    .slice(-6)
    .toUpperCase();
  return generatedKeys.find((key) => key.replace(/[^A-Z0-9]/gi, '').endsWith(suffix));
}

export class LicenseServiceError extends Error {
  readonly code?: string;
  readonly status?: number;

  constructor(message: string, code?: string, status?: number) {
    super(message);
    this.name = 'LicenseServiceError';
    this.code = code;
    this.status = status;
  }
}

interface LicenseResponse {
  error?: string;
  code?: string;
  [key: string]: unknown;
}

async function invoke<T>(body: Record<string, unknown>): Promise<T> {
  const { data, error } = await supabase.functions.invoke('license', { body });
  if (error) {
    let message = error.message || 'The licensing service could not be reached.';
    let code: string | undefined;
    if (
      'context' in error &&
      error.context &&
      typeof (error.context as Response).json === 'function'
    ) {
      try {
        const errorJson = (await (error.context as Response).clone().json()) as LicenseResponse;
        if (errorJson?.error) message = errorJson.error;
        if (errorJson?.code) code = errorJson.code;
      } catch {
        // Use default error message
      }
    }
    throw new LicenseServiceError(message, code);
  }
  const payload = (data ?? {}) as T & LicenseResponse;
  if (payload.error) throw new LicenseServiceError(payload.error, payload.code);
  return payload as T;
}

export async function fetchLicenseStatus(): Promise<LicenseStatus> {
  return invoke<LicenseStatus>({ action: 'status', installId: getInstallId() });
}

export async function heartbeatLicense(): Promise<LicenseStatus> {
  return invoke<LicenseStatus>({ action: 'heartbeat', installId: getInstallId() });
}

export async function activateLicense(licenseKey: string): Promise<LicenseStatus> {
  return invoke<LicenseStatus>({
    action: 'activate',
    installId: getInstallId(),
    licenseKey: licenseKey.trim(),
  });
}

export async function adminLicenseAction<T>(
  action: string,
  payload: Record<string, unknown> = {},
): Promise<T> {
  return invoke<T>({ action: `admin_${action}`, ...payload });
}

export async function fetchAdminLicenseData(): Promise<AdminLicenseData> {
  return adminLicenseAction<AdminLicenseData>('list');
}

export async function submitSupportTicket(ticket: {
  name: string;
  email: string;
  category: string;
  subject: string;
  message: string;
  userId?: string | null;
}): Promise<{ id: string }> {
  // 1. Try secure RPC first (runs with SECURITY DEFINER to avoid RLS/permission errors)
  try {
    const { data: rpcData, error: rpcErr } = await supabase.rpc('submit_support_ticket', {
      p_name: ticket.name.trim(),
      p_email: ticket.email.trim().toLowerCase(),
      p_category: ticket.category || 'general',
      p_subject: ticket.subject.trim(),
      p_message: ticket.message.trim(),
      p_user_id: ticket.userId || null,
    });
    if (!rpcErr) {
      if (
        rpcData &&
        typeof rpcData === 'object' &&
        'id' in rpcData &&
        typeof (rpcData as { id?: unknown }).id === 'string' &&
        (rpcData as { id: string }).id.trim()
      ) {
        return { id: (rpcData as { id: string }).id };
      }
      // A successful RPC must not fall through to a second write merely because its response was
      // malformed. The first call may already have committed the ticket.
      throw new Error('The support service returned an invalid ticket response. Please try again.');
    }
    if (
      rpcErr &&
      !/PGRST202|42883|function .*does not exist|could not find the function/i.test(rpcErr.message)
    ) {
      throw new Error(rpcErr.message || 'Failed to submit support ticket. Please try again.');
    }
  } catch (error) {
    if (
      !(error instanceof Error) ||
      !/PGRST202|42883|function .*does not exist|could not find the function/i.test(error.message)
    ) {
      throw error;
    }
    // Fall through only when the RPC is not deployed yet. A transient or validation error must not
    // be retried through a second write, which could create duplicate tickets after a lost response.
  }

  // 2. Fallback to direct table insertion
  const { data, error } = await supabase
    .from('support_tickets')
    .insert({
      name: ticket.name.trim(),
      email: ticket.email.trim().toLowerCase(),
      category: ticket.category || 'general',
      subject: ticket.subject.trim(),
      message: ticket.message.trim(),
      user_id: ticket.userId || null,
      status: 'pending',
    })
    .select('id')
    .single();

  if (error) {
    throw new Error(error.message || 'Failed to submit support ticket. Please try again.');
  }

  return { id: data.id };
}

export async function adminDeleteKey(licenseId: string): Promise<void> {
  await adminLicenseAction('delete_key', { licenseId });
}

export async function adminDeleteUnusedKeys(): Promise<{ deletedCount: number }> {
  return adminLicenseAction<{ deletedCount: number }>('delete_unused_keys');
}

export async function adminUpdateTicket(
  ticketId: string,
  status: 'pending' | 'in_progress' | 'resolved',
  adminNotes?: string,
  replied?: boolean,
): Promise<void> {
  await adminLicenseAction('update_ticket', {
    ticketId,
    status,
    adminNotes,
    // The current transactional function calls this flag `markReplied`. Keep the payload name in
    // sync with the database contract so resolving a ticket actually records the reply timestamp.
    markReplied: replied ? 'true' : 'false',
  });
}

export async function adminDeleteTicket(ticketId: string): Promise<void> {
  await adminLicenseAction('delete_ticket', { ticketId });
}
