import React, { useEffect, useMemo, useState } from 'react';
import {
  Clipboard,
  Copy,
  RefreshCw,
  ShieldCheck,
  UserRound,
  XCircle,
  Users,
  Key,
  LifeBuoy,
  CheckCircle2,
  Trash2,
  Mail,
  Activity,
  Plus,
  Crown,
  Infinity as InfinityIcon,
  Lock,
} from 'lucide-react';
import { useLicense } from '../lib/license-context.js';
import {
  adminLicenseAction,
  fetchAdminLicenseData,
  adminDeleteKey,
  adminDeleteUnusedKeys,
  adminUpdateTicket,
  adminDeleteTicket,
  resolvePlaintextLicenseKey,
  type AdminLicenseData,
  type AdminLicenseRecord,
  type AdminAccountRecord,
  type SupportTicketRecord,
} from '../lib/licensing.js';

interface AdminLicensePageProps {
  onBack?: () => void;
}

type CmsTab = 'overview' | 'users' | 'keys' | 'support' | 'diagnostics';

const SYSTEM_OWNER_EMAILS = new Set(['roshhellwett@gmail.com', 'zenithprojects@icloud.com']);

function isSystemOwner(account: AdminAccountRecord): boolean {
  if (
    account.is_admin ||
    account.role === 'owner' ||
    account.role === 'admin' ||
    account.status === 'owner'
  ) {
    return true;
  }
  if (account.email && SYSTEM_OWNER_EMAILS.has(account.email.toLowerCase())) {
    return true;
  }
  return false;
}

function formatDate(value: string | null | undefined): string {
  if (!value) return '—';
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '—' : parsed.toLocaleString();
}

export const AdminLicensePage: React.FC<AdminLicensePageProps> = ({ onBack }) => {
  const { status, loading: licenseLoading, error: licenseError, refresh } = useLicense();
  const [data, setData] = useState<AdminLicenseData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [feedback, setFeedback] = useState('');
  const [activeTab, setActiveTab] = useState<CmsTab>('overview');

  // Key issuance form
  const [count, setCount] = useState('5');
  const [durationDays, setDurationDays] = useState('30');
  const [generatedKeys, setGeneratedKeys] = useState<string[]>([]);

  // Search & Filters
  const [query, setQuery] = useState('');
  const [keyFilter, setKeyFilter] = useState<'all' | 'unused' | 'active' | 'revoked'>('all');
  const [ticketFilter, setTicketFilter] = useState<'all' | 'pending' | 'in_progress' | 'resolved'>(
    'all',
  );

  // Days adjustment per item
  const [customDaysMap, setCustomDaysMap] = useState<Record<string, string>>({});

  // Transfer state
  const [transfer, setTransfer] = useState<{
    id: string;
    email: string;
    note: string;
    installId: string;
    preserveDevice: boolean;
  } | null>(null);

  const load = async () => {
    setLoading(true);
    setError('');
    try {
      const res = await fetchAdminLicenseData();
      setData(res);
    } catch (requestError) {
      setError(
        requestError instanceof Error ? requestError.message : 'The admin console could not load.',
      );
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (status?.isAdmin) void load();
  }, [status?.isAdmin]);

  const accountByUser = useMemo(
    () => new Map((data?.accounts ?? []).map((account) => [account.user_id, account])),
    [data],
  );

  const filteredKeys = useMemo(() => {
    let keys = data?.keys ?? [];
    if (keyFilter !== 'all') {
      keys = keys.filter((k) => {
        if (keyFilter === 'unused') return k.status === 'unused';
        if (keyFilter === 'active') return k.status === 'active';
        if (keyFilter === 'revoked') return k.status === 'revoked' || k.status === 'expired';
        return true;
      });
    }
    const normalizedQuery = query.trim().toLowerCase();
    if (!normalizedQuery) return keys;
    return keys.filter((license) =>
      [license.key_hint, license.status, license.bound_email, license.bound_user_id].some((value) =>
        String(value ?? '')
          .toLowerCase()
          .includes(normalizedQuery),
      ),
    );
  }, [data, query, keyFilter]);

  const filteredAccounts = useMemo(() => {
    const accounts = data?.accounts ?? [];
    const normalizedQuery = query.trim().toLowerCase();
    if (!normalizedQuery) return accounts;
    return accounts.filter((account) =>
      [account.email, account.user_id, account.status].some((value) =>
        String(value ?? '')
          .toLowerCase()
          .includes(normalizedQuery),
      ),
    );
  }, [data, query]);

  const filteredTickets = useMemo(() => {
    let tickets = data?.tickets ?? [];
    if (ticketFilter !== 'all') {
      tickets = tickets.filter((t) => t.status === ticketFilter);
    }
    const normalizedQuery = query.trim().toLowerCase();
    if (!normalizedQuery) return tickets;
    return tickets.filter((t) =>
      [t.name, t.email, t.subject, t.message, t.category, t.status].some((val) =>
        String(val ?? '')
          .toLowerCase()
          .includes(normalizedQuery),
      ),
    );
  }, [data, query, ticketFilter]);

  const showFeedback = (msg: string) => {
    setFeedback(msg);
    setTimeout(() => setFeedback(''), 5000);
  };

  const run = async (
    action: string,
    payload: Record<string, unknown> = {},
    successMsg?: string,
  ) => {
    setLoading(true);
    setError('');
    try {
      await adminLicenseAction(action, payload);
      await load();
      if (successMsg) showFeedback(successMsg);
    } catch (requestError) {
      setError(
        requestError instanceof Error ? requestError.message : 'The administrator action failed.',
      );
      setLoading(false);
    }
  };

  const handleAdjustUserDays = async (userId: string, deltaDays: number, email?: string) => {
    await run(
      'adjust_trial',
      { userId, deltaDays },
      `Successfully adjusted ${deltaDays > 0 ? `+${deltaDays}` : deltaDays} days for ${email || 'user'}.`,
    );
  };

  const handleAdjustKeyDays = async (licenseId: string, deltaDays: number, keyHint?: string) => {
    await run(
      'adjust_days',
      { licenseId, deltaDays },
      `Successfully adjusted ${deltaDays > 0 ? `+${deltaDays}` : deltaDays} days for key ${keyHint || ''}.`,
    );
  };

  const handleDeleteSingleKey = async (licenseId: string, keyHint: string) => {
    if (!window.confirm(`Are you sure you want to permanently delete license key "${keyHint}"?`))
      return;
    setLoading(true);
    try {
      await adminDeleteKey(licenseId);
      await load();
      showFeedback(`License key ${keyHint} deleted.`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete key.');
      setLoading(false);
    }
  };

  const handleDeleteAllUnusedKeys = async () => {
    const unusedCount = data?.keys?.filter((k) => k.status === 'unused').length ?? 0;
    if (unusedCount === 0) {
      alert('There are no unused keys to delete.');
      return;
    }
    if (
      !window.confirm(
        `Are you sure you want to delete all ${unusedCount} unused license keys? This action cannot be undone.`,
      )
    ) {
      return;
    }
    setLoading(true);
    try {
      const res = await adminDeleteUnusedKeys();
      await load();
      showFeedback(`Successfully deleted ${res.deletedCount ?? unusedCount} unused keys.`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete unused keys.');
      setLoading(false);
    }
  };

  const handleUpdateTicketStatus = async (
    ticketId: string,
    newStatus: 'pending' | 'in_progress' | 'resolved',
  ) => {
    setLoading(true);
    try {
      await adminUpdateTicket(ticketId, newStatus);
      await load();
      showFeedback(`Support ticket status updated to "${newStatus}".`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to update ticket status.');
      setLoading(false);
    }
  };

  const handleDeleteSupportTicket = async (ticketId: string) => {
    if (!window.confirm('Are you sure you want to delete this support ticket?')) return;
    setLoading(true);
    try {
      await adminDeleteTicket(ticketId);
      await load();
      showFeedback('Support ticket deleted.');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete ticket.');
      setLoading(false);
    }
  };

  const handleReplyMailto = (ticket: SupportTicketRecord) => {
    const subjectLine = encodeURIComponent(
      `Re: [Ticket #${ticket.id.slice(0, 8).toUpperCase()}] ${ticket.subject}`,
    );
    const bodyContent = encodeURIComponent(
      `Hi ${ticket.name},\n\nThank you for reaching out to ExcelAgento Support.\nRegarding your inquiry:\n"${ticket.message}"\n\n\n\nBest regards,\nExcelAgento Operations Team\nzenithprojects@icloud.com`,
    );
    window.open(`mailto:${ticket.email}?subject=${subjectLine}&body=${bodyContent}`, '_blank');
    // Mark as in-progress or save reply timestamp
    void adminUpdateTicket(ticket.id, 'in_progress', undefined, true);
  };

  const generate = async (event: React.FormEvent) => {
    event.preventDefault();
    setLoading(true);
    setError('');
    try {
      const result = await adminLicenseAction<{ keys: string[] }>('generate', {
        count: Number(count),
        durationDays: Number(durationDays),
      });
      setGeneratedKeys(result.keys ?? []);
      await load();
      showFeedback(`Generated ${result.keys?.length ?? count} new activation keys!`);
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'Key generation failed.');
      setLoading(false);
    }
  };

  const transferLicense = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!transfer) return;
    await run(
      'transfer',
      {
        licenseId: transfer.id,
        targetEmail: transfer.email,
        verificationNote: transfer.note,
        installId: transfer.installId,
        preserveDevice: transfer.preserveDevice,
      },
      'License transferred successfully.',
    );
    setTransfer(null);
  };

  if (licenseLoading && !status) {
    return (
      <div className="app-container app-page-scroll license-page">
        <main className="license-shell admin-license-shell">
          <div className="license-alert info" role="status">
            <RefreshCw size={17} className="animate-spin" /> Verifying administrator access…
          </div>
        </main>
      </div>
    );
  }

  if (!status?.isAdmin) {
    return (
      <div className="app-container app-page-scroll license-page">
        <main className="license-shell admin-license-shell">
          <div className="license-alert danger" role="alert">
            <XCircle size={17} />
            {licenseError || 'Administrator access is required.'}
          </div>
          <div style={{ display: 'flex', gap: '0.75rem', marginTop: '1rem' }}>
            <button type="button" className="btn btn-secondary" onClick={() => void refresh()}>
              <RefreshCw size={14} /> Retry verification
            </button>
            <button type="button" className="btn btn-ghost" onClick={onBack}>
              Return to workspace
            </button>
          </div>
        </main>
      </div>
    );
  }

  const unusedKeysCount = data?.keys?.filter((k) => k.status === 'unused').length ?? 0;
  const pendingTicketsCount = data?.tickets?.filter((t) => t.status === 'pending').length ?? 0;

  return (
    <div className="app-container app-page-scroll license-page">
      <main className="license-shell admin-license-shell" aria-labelledby="admin-license-title">
        {/* Top bar */}
        <div className="license-topbar">
          <button type="button" className="btn btn-ghost btn-sm" onClick={onBack}>
            ← Preview Workspace
          </button>
          <span className="studio-eyebrow">EXCELAGENTO · OPERATIONS CONTROL PLANE</span>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => void load()}
            disabled={loading}
          >
            <RefreshCw size={14} className={loading ? 'license-spin' : undefined} /> Refresh
          </button>
        </div>

        {/* Hero Section */}
        <section className="license-hero license-tone-success">
          <div className="license-hero-icon">
            <ShieldCheck size={28} />
          </div>
          <div>
            <span className="studio-eyebrow">SERVER-AUTHORITATIVE CMS DASHBOARD</span>
            <h1 id="admin-license-title">Operations &amp; Entitlements</h1>
            <p>
              Manage users, license keys, customer support tickets, and system operations
              seamlessly.
            </p>
          </div>
          <span className="license-state-badge success">Super Admin</span>
        </section>

        {/* Alerts & Toasts */}
        {error && (
          <div className="license-alert danger" role="alert">
            <XCircle size={17} />
            <span>{error}</span>
          </div>
        )}
        {feedback && (
          <div className="license-alert success" role="status">
            <CheckCircle2 size={17} />
            <span>{feedback}</span>
          </div>
        )}

        {/* CMS Navigation Tabs */}
        <nav className="admin-cms-nav" aria-label="CMS Sections">
          <button
            type="button"
            className={`admin-cms-tab-btn ${activeTab === 'overview' ? 'is-active' : ''}`}
            onClick={() => setActiveTab('overview')}
          >
            <Activity size={15} /> Overview
          </button>
          <button
            type="button"
            className={`admin-cms-tab-btn ${activeTab === 'users' ? 'is-active' : ''}`}
            onClick={() => setActiveTab('users')}
          >
            <Users size={15} /> Registered Users
            <span className="admin-cms-badge">{data?.accounts?.length ?? 0}</span>
          </button>
          <button
            type="button"
            className={`admin-cms-tab-btn ${activeTab === 'keys' ? 'is-active' : ''}`}
            onClick={() => setActiveTab('keys')}
          >
            <Key size={15} /> License Keys
            {unusedKeysCount > 0 && (
              <span className="admin-cms-badge">{unusedKeysCount} unused</span>
            )}
          </button>
          <button
            type="button"
            className={`admin-cms-tab-btn ${activeTab === 'support' ? 'is-active' : ''}`}
            onClick={() => setActiveTab('support')}
          >
            <LifeBuoy size={15} /> Support Desk
            {pendingTicketsCount > 0 && (
              <span className="admin-cms-badge" style={{ background: '#ef4444', color: '#fff' }}>
                {pendingTicketsCount} pending
              </span>
            )}
          </button>
          <button
            type="button"
            className={`admin-cms-tab-btn ${activeTab === 'diagnostics' ? 'is-active' : ''}`}
            onClick={() => setActiveTab('diagnostics')}
          >
            <ShieldCheck size={15} /> System Diagnostics
          </button>
        </nav>

        {/* 1. OVERVIEW TAB */}
        {activeTab === 'overview' && (
          <>
            <section className="admin-stats-grid">
              <div className="admin-stat-card">
                <div className="admin-stat-card-header">
                  <span className="admin-stat-card-title">Registered Users</span>
                  <div className="admin-stat-icon-wrap users">
                    <Users size={18} />
                  </div>
                </div>
                <div className="admin-stat-card-value">{data?.accounts?.length ?? 0}</div>
                <span className="admin-stat-card-sub">Active spreadsheet accounts</span>
              </div>

              <div className="admin-stat-card">
                <div className="admin-stat-card-header">
                  <span className="admin-stat-card-title">Active Licenses</span>
                  <div className="admin-stat-icon-wrap licenses">
                    <ShieldCheck size={18} />
                  </div>
                </div>
                <div className="admin-stat-card-value">
                  {data?.keys?.filter((k) => k.status === 'active').length ?? 0}
                </div>
                <span className="admin-stat-card-sub">Activated entitlements</span>
              </div>

              <div className="admin-stat-card">
                <div className="admin-stat-card-header">
                  <span className="admin-stat-card-title">Unused Keys</span>
                  <div className="admin-stat-icon-wrap keys">
                    <Key size={18} />
                  </div>
                </div>
                <div className="admin-stat-card-value">{unusedKeysCount}</div>
                <span className="admin-stat-card-sub">Available for distribution</span>
              </div>

              <div className="admin-stat-card">
                <div className="admin-stat-card-header">
                  <span className="admin-stat-card-title">Support Desk</span>
                  <div className="admin-stat-icon-wrap tickets">
                    <LifeBuoy size={18} />
                  </div>
                </div>
                <div className="admin-stat-card-value">{pendingTicketsCount}</div>
                <span className="admin-stat-card-sub">Pending user inquiries</span>
              </div>
            </section>

            {/* Quick Actions Strip */}
            <section className="license-card" style={{ marginBottom: '20px' }}>
              <div className="license-card-heading">
                <div>
                  <span className="studio-eyebrow">CONTROL ACCELERATORS</span>
                  <h2>Quick Management Operations</h2>
                </div>
              </div>
              <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveTab('keys')}
                >
                  <Plus size={15} /> Generate New Keys
                </button>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveTab('users')}
                >
                  <Users size={15} /> Extend User Trial / Access
                </button>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveTab('support')}
                >
                  <LifeBuoy size={15} /> Review Support Tickets ({pendingTicketsCount} pending)
                </button>
                {unusedKeysCount > 0 && (
                  <button
                    type="button"
                    className="btn btn-ghost danger-text"
                    onClick={handleDeleteAllUnusedKeys}
                  >
                    <Trash2 size={15} /> Clean Up All Unused Keys ({unusedKeysCount})
                  </button>
                )}
              </div>
            </section>

            {/* Recent Audit trail */}
            <section className="license-card admin-audit-card">
              <div className="license-card-heading">
                <div>
                  <span className="studio-eyebrow">AUDIT TRAIL</span>
                  <h2>Recent administrative events</h2>
                </div>
              </div>
              <div className="admin-event-list">
                {(data?.events ?? []).slice(0, 10).map((event) => (
                  <div key={event.id}>
                    <code>{event.event_type}</code>
                    <span>{formatDate(event.created_at)}</span>
                    <small>{event.target_user_id || event.license_id || 'system'}</small>
                  </div>
                ))}
                {(data?.events ?? []).length === 0 && (
                  <div className="admin-empty-state">No recent administrative events.</div>
                )}
              </div>
            </section>
          </>
        )}

        {/* 2. USERS MANAGEMENT TAB */}
        {activeTab === 'users' && (
          <section className="license-card admin-table-card">
            <div className="admin-table-toolbar">
              <div>
                <span className="studio-eyebrow">USER ACCOUNTS &amp; EXTENSIONS</span>
                <h2>Registered Accounts ({data?.accounts?.length ?? 0})</h2>
              </div>
              <input
                className="form-input admin-search"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search user email or ID…"
              />
            </div>

            <div className="admin-table-wrap">
              <table className="admin-license-table">
                <thead>
                  <tr>
                    <th>User Email / ID</th>
                    <th>Status</th>
                    <th>Days Remaining</th>
                    <th>Expiry Timestamp</th>
                    <th>Extend Access (Add Days)</th>
                    <th>Account Controls</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAccounts.map((account: AdminAccountRecord) => {
                    const days = account.days_remaining ?? 0;
                    const customDays = customDaysMap[account.user_id] || '30';
                    const isOwner = isSystemOwner(account);

                    if (isOwner) {
                      return (
                        <tr key={account.user_id} className="admin-owner-row">
                          <td>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <strong>{account.email}</strong>
                                <span className="admin-owner-pill">
                                  <Crown size={11} /> Root Owner
                                </span>
                              </div>
                              <small>{account.user_id}</small>
                            </div>
                          </td>
                          <td>
                            <span className="license-state-badge owner">
                              <Crown size={12} /> SUPER ADMIN
                            </span>
                          </td>
                          <td>
                            <span className="days-remaining-pill lifetime">
                              <InfinityIcon size={13} /> Lifetime Access
                            </span>
                          </td>
                          <td>
                            <span className="admin-permanent-text">
                              Never Expires (Permanent ∞)
                            </span>
                          </td>
                          <td>
                            <div className="admin-root-privilege-chip">
                              <ShieldCheck size={14} /> Full Root Privilege (Immortal)
                            </div>
                          </td>
                          <td>
                            <div className="admin-protected-btn">
                              <Lock size={12} /> Root Protected
                            </div>
                          </td>
                        </tr>
                      );
                    }

                    return (
                      <tr key={account.user_id}>
                        <td>
                          <strong>{account.email}</strong>
                          <small>{account.user_id}</small>
                        </td>
                        <td>
                          <span
                            className={`license-state-badge ${
                              account.status === 'licensed'
                                ? 'success'
                                : account.status === 'trial'
                                  ? 'info'
                                  : account.status === 'banned'
                                    ? 'danger'
                                    : 'warning'
                            }`}
                          >
                            {account.status}
                          </span>
                          {account.ban_reason && (
                            <small className="admin-banned-label">{account.ban_reason}</small>
                          )}
                        </td>
                        <td>
                          <span
                            className={`days-remaining-pill ${
                              account.status === 'banned' || days <= 0
                                ? 'expired'
                                : days <= 7
                                  ? 'warning'
                                  : ''
                            }`}
                          >
                            {days > 0 ? `${days} days left` : 'Expired'}
                          </span>
                        </td>
                        <td>{formatDate(account.trial_expires_at)}</td>
                        <td>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <div className="admin-actions">
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustUserDays(account.user_id, 7, account.email)
                                }
                                title="Add 7 days"
                              >
                                +7d
                              </button>
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustUserDays(account.user_id, 30, account.email)
                                }
                                title="Add 30 days"
                              >
                                +30d
                              </button>
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustUserDays(account.user_id, 60, account.email)
                                }
                                title="Add 60 days"
                              >
                                +60d
                              </button>
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustUserDays(account.user_id, 365, account.email)
                                }
                                title="Add 1 year"
                              >
                                +365d
                              </button>
                            </div>

                            {/* Custom days input */}
                            <div className="inline-day-adjust">
                              <input
                                className="form-input"
                                type="number"
                                min="1"
                                max="3650"
                                value={customDays}
                                onChange={(e) =>
                                  setCustomDaysMap({
                                    ...customDaysMap,
                                    [account.user_id]: e.target.value,
                                  })
                                }
                                placeholder="Days"
                              />
                              <button
                                type="button"
                                className="btn btn-ghost btn-sm"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustUserDays(
                                    account.user_id,
                                    Number(customDays) || 30,
                                    account.email,
                                  )
                                }
                              >
                                Apply +{customDays}d
                              </button>
                            </div>
                          </div>
                        </td>
                        <td>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            disabled={loading}
                            onClick={() =>
                              void run(
                                account.status === 'banned' ? 'unban_user' : 'ban_user',
                                {
                                  userId: account.user_id,
                                  reason: 'Administrative suspension review.',
                                },
                                account.status === 'banned'
                                  ? `Account ${account.email} unbanned.`
                                  : `Account ${account.email} banned.`,
                              )
                            }
                          >
                            {account.status === 'banned' ? 'Unban Account' : 'Ban Account'}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>

              {filteredAccounts.length === 0 && (
                <div className="admin-empty-state">
                  <UserRound size={22} /> No registered accounts match your query.
                </div>
              )}
            </div>
          </section>
        )}

        {/* 3. LICENSE KEYS CMS TAB */}
        {activeTab === 'keys' && (
          <>
            {/* Key Generator Tool */}
            <section className="admin-command-grid" style={{ marginBottom: '20px' }}>
              <form className="license-card admin-generate-card" onSubmit={(e) => void generate(e)}>
                <div className="license-card-heading">
                  <div>
                    <span className="studio-eyebrow">KEY ISSUANCE</span>
                    <h2>Generate Activation Keys</h2>
                  </div>
                  <Clipboard size={21} />
                </div>
                <div className="admin-form-row">
                  <label>
                    Count
                    <input
                      className="form-input"
                      type="number"
                      min="1"
                      max="100"
                      value={count}
                      onChange={(e) => setCount(e.target.value)}
                    />
                  </label>
                  <label>
                    Validity (Days)
                    <input
                      className="form-input"
                      type="number"
                      min="1"
                      max="3650"
                      value={durationDays}
                      onChange={(e) => setDurationDays(e.target.value)}
                    />
                  </label>
                  <button type="submit" className="btn btn-primary" disabled={loading}>
                    <Plus size={15} /> Generate
                  </button>
                </div>
                <p className="license-muted">
                  Cryptographically generated activation keys. Plaintext keys can be copied
                  immediately upon creation.
                </p>
              </form>

              {generatedKeys.length > 0 && (
                <section className="license-card admin-generated-card">
                  <div className="license-card-heading">
                    <div>
                      <span className="studio-eyebrow">SHOW ONCE</span>
                      <h2>New Generated Batch ({generatedKeys.length})</h2>
                    </div>
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm"
                      onClick={() => setGeneratedKeys([])}
                    >
                      Dismiss
                    </button>
                  </div>
                  <textarea
                    className="admin-key-output"
                    readOnly
                    value={generatedKeys.join('\n')}
                    aria-label="Generated activation keys"
                  />
                  <button
                    type="button"
                    className="btn btn-secondary btn-sm"
                    onClick={() => {
                      void navigator.clipboard?.writeText(generatedKeys.join('\n'));
                      showFeedback('Copied key batch to clipboard!');
                    }}
                  >
                    <Copy size={14} /> Copy batch
                  </button>
                </section>
              )}
            </section>

            {/* Keys Table */}
            <section className="license-card admin-table-card">
              <div className="admin-table-toolbar">
                <div>
                  <span className="studio-eyebrow">ENTITLEMENTS CMS</span>
                  <h2>License Keys ({data?.keys?.length ?? 0})</h2>
                </div>
                <div className="admin-toolbar-inputs">
                  <input
                    className="form-input admin-search"
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search key, email, status…"
                  />
                  {unusedKeysCount > 0 && (
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm danger-text"
                      disabled={loading}
                      onClick={handleDeleteAllUnusedKeys}
                    >
                      <Trash2 size={14} /> Delete All Unused ({unusedKeysCount})
                    </button>
                  )}
                </div>
              </div>

              {/* Filter tabs */}
              <div className="ticket-filter-bar">
                <button
                  type="button"
                  className={`ticket-filter-btn ${keyFilter === 'all' ? 'active' : ''}`}
                  onClick={() => setKeyFilter('all')}
                >
                  All Keys ({data?.keys?.length ?? 0})
                </button>
                <button
                  type="button"
                  className={`ticket-filter-btn ${keyFilter === 'unused' ? 'active' : ''}`}
                  onClick={() => setKeyFilter('unused')}
                >
                  Unused Keys ({unusedKeysCount})
                </button>
                <button
                  type="button"
                  className={`ticket-filter-btn ${keyFilter === 'active' ? 'active' : ''}`}
                  onClick={() => setKeyFilter('active')}
                >
                  Active Keys ({data?.keys?.filter((k) => k.status === 'active').length ?? 0})
                </button>
                <button
                  type="button"
                  className={`ticket-filter-btn ${keyFilter === 'revoked' ? 'active' : ''}`}
                  onClick={() => setKeyFilter('revoked')}
                >
                  Revoked / Expired
                </button>
              </div>

              <div className="admin-table-wrap">
                <table className="admin-license-table">
                  <thead>
                    <tr>
                      <th>Key / Plaintext</th>
                      <th>Account</th>
                      <th>Status</th>
                      <th>Days Left / Expiry</th>
                      <th>Adjust Days</th>
                      <th>Controls &amp; Deletion</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredKeys.map((license: AdminLicenseRecord) => {
                      const account = license.bound_user_id
                        ? accountByUser.get(license.bound_user_id)
                        : undefined;
                      const plaintextKey = resolvePlaintextLicenseKey(license, generatedKeys);
                      const daysLeft = license.days_remaining ?? license.duration_days;

                      return (
                        <tr key={license.id}>
                          <td>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '0.45rem' }}>
                              <code>{plaintextKey || license.key_hint}</code>
                              <button
                                type="button"
                                className="btn btn-ghost btn-icon btn-sm"
                                onClick={() => {
                                  if (plaintextKey) {
                                    void navigator.clipboard?.writeText(plaintextKey);
                                    showFeedback('Copied key to clipboard!');
                                  } else {
                                    setError(
                                      'Older hashed key cannot be recovered in plaintext. Generate a replacement key to copy.',
                                    );
                                  }
                                }}
                                title={plaintextKey ? 'Copy activation key' : 'Key hash only'}
                                disabled={!plaintextKey}
                                style={{ padding: '2px', height: 'auto' }}
                              >
                                <Copy size={13} />
                              </button>
                            </div>
                            <small>{license.duration_days} days validity</small>
                          </td>
                          <td>
                            {license.bound_email || (
                              <span style={{ color: 'var(--text-muted)' }}>
                                Unassigned (Available)
                              </span>
                            )}
                            {license.bound_user_id && <small>{license.bound_user_id}</small>}
                          </td>
                          <td>
                            <span
                              className={`license-state-badge ${
                                license.status === 'active'
                                  ? 'success'
                                  : license.status === 'unused'
                                    ? 'info'
                                    : license.status === 'revoked'
                                      ? 'danger'
                                      : 'warning'
                              }`}
                            >
                              {license.status}
                            </span>
                            {account?.status === 'banned' && (
                              <small className="admin-banned-label">Account Banned</small>
                            )}
                          </td>
                          <td>
                            <span
                              className={`days-remaining-pill ${
                                license.status === 'unused'
                                  ? 'info'
                                  : daysLeft <= 0
                                    ? 'expired'
                                    : daysLeft <= 7
                                      ? 'warning'
                                      : ''
                              }`}
                            >
                              {license.status === 'unused'
                                ? `${license.duration_days}d upon activation`
                                : `${daysLeft} days remaining`}
                            </span>
                            <small>{formatDate(license.expires_at)}</small>
                          </td>
                          <td>
                            <div className="admin-actions">
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustKeyDays(license.id, 7, license.key_hint)
                                }
                                title="Add 7 days"
                              >
                                +7d
                              </button>
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustKeyDays(license.id, 30, license.key_hint)
                                }
                                title="Add 30 days"
                              >
                                +30d
                              </button>
                              <button
                                type="button"
                                className="btn-chip"
                                disabled={loading}
                                onClick={() =>
                                  void handleAdjustKeyDays(license.id, -7, license.key_hint)
                                }
                                title="Subtract 7 days"
                              >
                                −7d
                              </button>
                            </div>
                          </td>
                          <td>
                            <div className="admin-actions">
                              {/* Reset Device */}
                              {license.bound_device_id && (
                                <button
                                  type="button"
                                  className="btn btn-ghost btn-sm"
                                  onClick={() =>
                                    void run(
                                      'reset_device',
                                      { licenseId: license.id },
                                      'Device binding reset.',
                                    )
                                  }
                                >
                                  Reset device
                                </button>
                              )}

                              {/* Revoke */}
                              {license.status === 'active' && (
                                <button
                                  type="button"
                                  className="btn btn-ghost btn-sm danger-text"
                                  onClick={() =>
                                    void run(
                                      'revoke',
                                      { licenseId: license.id },
                                      'License revoked.',
                                    )
                                  }
                                >
                                  Revoke
                                </button>
                              )}

                              {/* Delete Key */}
                              {(license.status === 'unused' ||
                                license.status === 'revoked' ||
                                license.status === 'expired') && (
                                <button
                                  type="button"
                                  className="btn btn-ghost btn-sm danger-text"
                                  disabled={loading}
                                  onClick={() =>
                                    void handleDeleteSingleKey(
                                      license.id,
                                      plaintextKey || license.key_hint,
                                    )
                                  }
                                  title="Delete key permanently"
                                >
                                  <Trash2 size={13} /> Delete
                                </button>
                              )}

                              {/* Transfer */}
                              {license.bound_user_id && (
                                <button
                                  type="button"
                                  className="btn btn-ghost btn-sm"
                                  onClick={() =>
                                    setTransfer({
                                      id: license.id,
                                      email: '',
                                      note: '',
                                      installId: '',
                                      preserveDevice: false,
                                    })
                                  }
                                >
                                  Transfer
                                </button>
                              )}
                            </div>

                            {transfer?.id === license.id && (
                              <form
                                className="admin-transfer-form"
                                onSubmit={(e) => void transferLicense(e)}
                              >
                                <input
                                  className="form-input"
                                  type="email"
                                  required
                                  placeholder="New account email"
                                  value={transfer.email}
                                  onChange={(e) =>
                                    setTransfer({ ...transfer, email: e.target.value })
                                  }
                                />
                                <input
                                  className="form-input"
                                  required={!transfer.preserveDevice}
                                  minLength={16}
                                  placeholder="Destination install ID"
                                  value={transfer.installId}
                                  onChange={(e) =>
                                    setTransfer({ ...transfer, installId: e.target.value })
                                  }
                                  disabled={transfer.preserveDevice}
                                />
                                <input
                                  className="form-input"
                                  required
                                  minLength={12}
                                  placeholder="Ownership verification note"
                                  value={transfer.note}
                                  onChange={(e) =>
                                    setTransfer({ ...transfer, note: e.target.value })
                                  }
                                />
                                <label>
                                  <input
                                    type="checkbox"
                                    checked={transfer.preserveDevice}
                                    onChange={(e) =>
                                      setTransfer({ ...transfer, preserveDevice: e.target.checked })
                                    }
                                  />{' '}
                                  Keep verified device binding
                                </label>
                                <button
                                  className="btn btn-primary btn-sm"
                                  type="submit"
                                  disabled={loading}
                                >
                                  Confirm transfer
                                </button>
                                <button
                                  className="btn btn-ghost btn-sm"
                                  type="button"
                                  onClick={() => setTransfer(null)}
                                >
                                  Cancel
                                </button>
                              </form>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>

                {filteredKeys.length === 0 && (
                  <div className="admin-empty-state">
                    <Key size={22} /> No license records match this filter.
                  </div>
                )}
              </div>
            </section>
          </>
        )}

        {/* 4. SUPPORT & HELPDESK CMS TAB */}
        {activeTab === 'support' && (
          <section className="license-card admin-table-card">
            <div className="admin-table-toolbar">
              <div>
                <span className="studio-eyebrow">CUSTOMER SUPPORT INQUIRIES</span>
                <h2>Support Tickets ({data?.tickets?.length ?? 0})</h2>
              </div>
              <input
                className="form-input admin-search"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search sender, email, subject, message…"
              />
            </div>

            {/* Filter Buttons */}
            <div className="ticket-filter-bar">
              <button
                type="button"
                className={`ticket-filter-btn ${ticketFilter === 'all' ? 'active' : ''}`}
                onClick={() => setTicketFilter('all')}
              >
                All Tickets ({data?.tickets?.length ?? 0})
              </button>
              <button
                type="button"
                className={`ticket-filter-btn ${ticketFilter === 'pending' ? 'active' : ''}`}
                onClick={() => setTicketFilter('pending')}
              >
                Pending ({pendingTicketsCount})
              </button>
              <button
                type="button"
                className={`ticket-filter-btn ${ticketFilter === 'in_progress' ? 'active' : ''}`}
                onClick={() => setTicketFilter('in_progress')}
              >
                In Progress ({data?.tickets?.filter((t) => t.status === 'in_progress').length ?? 0})
              </button>
              <button
                type="button"
                className={`ticket-filter-btn ${ticketFilter === 'resolved' ? 'active' : ''}`}
                onClick={() => setTicketFilter('resolved')}
              >
                Resolved ({data?.tickets?.filter((t) => t.status === 'resolved').length ?? 0})
              </button>
            </div>

            {/* Tickets List */}
            <div className="ticket-list">
              {filteredTickets.map((ticket: SupportTicketRecord) => (
                <div key={ticket.id} className="ticket-card">
                  <div className="ticket-card-top">
                    <div className="ticket-meta-info">
                      <span className="ticket-id-tag">
                        #TICK-{ticket.id.slice(0, 8).toUpperCase()}
                      </span>
                      <span className="ticket-category-tag">{ticket.category}</span>
                      <span className="ticket-user-line">
                        <strong>{ticket.name}</strong>
                        <span>({ticket.email})</span>
                      </span>
                    </div>
                    <span className="text-sm text-muted">{formatDate(ticket.created_at)}</span>
                  </div>

                  <h3 className="ticket-subject">{ticket.subject}</h3>
                  <div className="ticket-message-body">{ticket.message}</div>

                  <div className="ticket-card-bottom">
                    <div className="ticket-status-selector">
                      <span className="ticket-status-label">Status:</span>
                      <button
                        type="button"
                        className={`status-badge-btn pending ${ticket.status === 'pending' ? 'active' : ''}`}
                        onClick={() => void handleUpdateTicketStatus(ticket.id, 'pending')}
                      >
                        Pending
                      </button>
                      <button
                        type="button"
                        className={`status-badge-btn in_progress ${ticket.status === 'in_progress' ? 'active' : ''}`}
                        onClick={() => void handleUpdateTicketStatus(ticket.id, 'in_progress')}
                      >
                        In Progress
                      </button>
                      <button
                        type="button"
                        className={`status-badge-btn resolved ${ticket.status === 'resolved' ? 'active' : ''}`}
                        onClick={() => void handleUpdateTicketStatus(ticket.id, 'resolved')}
                      >
                        Resolved
                      </button>
                    </div>

                    <div className="ticket-reply-actions">
                      <button
                        type="button"
                        className="btn-email-reply"
                        onClick={() => handleReplyMailto(ticket)}
                        title="Open email composer to reply directly to user"
                      >
                        <Mail size={14} /> Reply on User Email
                      </button>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm danger-text"
                        onClick={() => void handleDeleteSupportTicket(ticket.id)}
                        title="Delete ticket"
                      >
                        <Trash2 size={14} /> Delete
                      </button>
                    </div>
                  </div>
                </div>
              ))}

              {filteredTickets.length === 0 && (
                <div className="admin-empty-state">
                  <LifeBuoy size={24} /> No support tickets in this view.
                </div>
              )}
            </div>
          </section>
        )}

        {/* 5. SYSTEM DIAGNOSTICS TAB */}
        {activeTab === 'diagnostics' && (
          <>
            <section className="license-card admin-table-card" style={{ marginBottom: '20px' }}>
              <div className="license-card-heading">
                <div>
                  <span className="studio-eyebrow">HARDWARE BINDINGS</span>
                  <h2>Device Installations Registry</h2>
                </div>
              </div>
              <p className="license-muted" style={{ marginTop: 0, marginBottom: '16px' }}>
                Hardware hash identifiers bound to user workstations for anti-abuse and license
                containment.
              </p>
              <div className="admin-device-list">
                {(data?.devices ?? []).map((device) => (
                  <div className="admin-device-row" key={device.id}>
                    <code>{device.install_id_hint}</code>
                    <span>{accountByUser.get(device.user_id)?.email || device.user_id}</span>
                    <span
                      className={`license-state-badge ${
                        device.status === 'active'
                          ? 'success'
                          : device.status === 'banned'
                            ? 'danger'
                            : 'warning'
                      }`}
                    >
                      {device.status}
                    </span>
                    <span className="admin-device-seen">
                      Last seen {formatDate(device.last_seen_at)}
                    </span>
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm"
                      onClick={() =>
                        void run(
                          device.status === 'banned' ? 'unban_device' : 'ban_device',
                          {
                            deviceId: device.id,
                            reason: 'Device access modified by administrator.',
                          },
                          `Device ${device.install_id_hint} ${device.status === 'banned' ? 'unbanned' : 'banned'}.`,
                        )
                      }
                    >
                      {device.status === 'banned' ? 'Unban device' : 'Ban device'}
                    </button>
                  </div>
                ))}
                {(data?.devices ?? []).length === 0 && (
                  <div className="admin-empty-state">
                    <UserRound size={22} /> No hardware device records.
                  </div>
                )}
              </div>
            </section>

            <section className="license-card admin-audit-card">
              <div className="license-card-heading">
                <div>
                  <span className="studio-eyebrow">AUDIT STREAM</span>
                  <h2>Recent administrative events</h2>
                </div>
              </div>
              <div className="admin-event-list">
                {(data?.events ?? []).map((event) => (
                  <div key={event.id}>
                    <code>{event.event_type}</code>
                    <span>{formatDate(event.created_at)}</span>
                    <small>{event.target_user_id || event.license_id || 'system'}</small>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}
      </main>
    </div>
  );
};
