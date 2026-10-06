import React, { useEffect, useMemo, useState } from 'react';
import { Clipboard, RefreshCw, ShieldCheck, UserRound, XCircle } from 'lucide-react';
import { useLicense } from '../lib/license-context.js';
import {
  adminLicenseAction,
  fetchAdminLicenseData,
  type AdminLicenseData,
  type AdminLicenseRecord,
} from '../lib/licensing.js';

interface AdminLicensePageProps {
  onBack?: () => void;
}

function formatDate(value: string | null | undefined): string {
  if (!value) return '—';
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '—' : parsed.toLocaleString();
}

export const AdminLicensePage: React.FC<AdminLicensePageProps> = ({ onBack }) => {
  const { status, loading: licenseLoading, refresh } = useLicense();
  const [data, setData] = useState<AdminLicenseData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [count, setCount] = useState('10');
  const [durationDays, setDurationDays] = useState('30');
  const [adjustmentDays, setAdjustmentDays] = useState('7');
  const [generatedKeys, setGeneratedKeys] = useState<string[]>([]);
  const [query, setQuery] = useState('');
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
      setData(await fetchAdminLicenseData());
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
    const normalizedQuery = query.trim().toLowerCase();
    if (!normalizedQuery) return data?.keys ?? [];
    return (data?.keys ?? []).filter((license) =>
      [license.key_hint, license.status, license.bound_email, license.bound_user_id].some((value) =>
        String(value ?? '')
          .toLowerCase()
          .includes(normalizedQuery),
      ),
    );
  }, [data, query]);

  const run = async (action: string, payload: Record<string, unknown> = {}) => {
    setLoading(true);
    setError('');
    try {
      await adminLicenseAction(action, payload);
      await load();
    } catch (requestError) {
      setError(
        requestError instanceof Error ? requestError.message : 'The administrator action failed.',
      );
      setLoading(false);
    }
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
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'Key generation failed.');
      setLoading(false);
    }
  };

  const transferLicense = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!transfer) return;
    await run('transfer', {
      licenseId: transfer.id,
      targetEmail: transfer.email,
      verificationNote: transfer.note,
      installId: transfer.installId,
      preserveDevice: transfer.preserveDevice,
    });
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
            <XCircle size={17} /> Administrator access is required.
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

  return (
    <div className="app-container app-page-scroll license-page">
      <main className="license-shell admin-license-shell" aria-labelledby="admin-license-title">
        <div className="license-topbar">
          <button type="button" className="btn btn-ghost btn-sm" onClick={onBack}>
            ← Back to workspace
          </button>
          <span className="studio-eyebrow">OPERATIONS · LICENSE ADMINISTRATION</span>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => void load()}
            disabled={loading}
          >
            <RefreshCw size={14} className={loading ? 'license-spin' : undefined} /> Refresh
          </button>
        </div>
        <section className="license-hero license-tone-success">
          <div className="license-hero-icon">
            <ShieldCheck size={26} />
          </div>
          <div>
            <span className="studio-eyebrow">SERVER-AUTHORITATIVE CONTROLS</span>
            <h1 id="admin-license-title">License operations</h1>
            <p>Generate, extend, revoke, transfer, and audit official ExcelAgento entitlements.</p>
          </div>
          <span className="license-state-badge success">Admin</span>
        </section>
        {error && (
          <div className="license-alert danger" role="alert">
            <XCircle size={17} />
            {error}
          </div>
        )}

        <section className="admin-command-grid">
          <form
            className="license-card admin-generate-card"
            onSubmit={(event) => void generate(event)}
          >
            <div className="license-card-heading">
              <div>
                <span className="studio-eyebrow">KEY ISSUANCE</span>
                <h2>Generate keys</h2>
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
                  max="1000"
                  value={count}
                  onChange={(event) => setCount(event.target.value)}
                />
              </label>
              <label>
                Days
                <input
                  className="form-input"
                  type="number"
                  min="1"
                  max="3650"
                  value={durationDays}
                  onChange={(event) => setDurationDays(event.target.value)}
                />
              </label>
              <button type="submit" className="btn btn-primary" disabled={loading}>
                Generate
              </button>
            </div>
            <p className="license-muted">
              Keys are generated with cryptographic randomness and stored only as hashes. Copy the
              plaintext batch now; it cannot be recovered later.
            </p>
          </form>
          {generatedKeys.length > 0 && (
            <section className="license-card admin-generated-card">
              <div className="license-card-heading">
                <div>
                  <span className="studio-eyebrow">SHOW ONCE</span>
                  <h2>New keys</h2>
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
                onClick={() => void navigator.clipboard?.writeText(generatedKeys.join('\n'))}
              >
                Copy batch
              </button>
            </section>
          )}
        </section>

        <section className="license-card admin-table-card">
          <div className="admin-table-toolbar">
            <div>
              <span className="studio-eyebrow">CONTROL PLANE</span>
              <h2>Entitlements</h2>
            </div>
            <div className="admin-toolbar-inputs">
              <label>
                Days to adjust
                <input
                  className="form-input"
                  type="number"
                  min="1"
                  max="3650"
                  value={adjustmentDays}
                  onChange={(event) => setAdjustmentDays(event.target.value)}
                />
              </label>
              <input
                className="form-input admin-search"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="Search email, status, key…"
              />
            </div>
          </div>
          <div className="admin-table-wrap">
            <table className="admin-license-table">
              <thead>
                <tr>
                  <th>Key</th>
                  <th>Account</th>
                  <th>Status</th>
                  <th>Device</th>
                  <th>Expiry</th>
                  <th>Controls</th>
                </tr>
              </thead>
              <tbody>
                {filteredKeys.map((license: AdminLicenseRecord) => {
                  const account = license.bound_user_id
                    ? accountByUser.get(license.bound_user_id)
                    : undefined;
                  return (
                    <tr key={license.id}>
                      <td>
                        <code>{license.key_hint}</code>
                        <small>{license.duration_days} days</small>
                      </td>
                      <td>
                        {license.bound_email || 'Unassigned'}
                        {license.bound_user_id && <small>{license.bound_user_id}</small>}
                      </td>
                      <td>
                        <span
                          className={`license-state-badge ${license.status === 'active' ? 'success' : license.status === 'revoked' ? 'danger' : 'warning'}`}
                        >
                          {license.status}
                        </span>
                        {account?.status === 'banned' && (
                          <small className="admin-banned-label">Banned</small>
                        )}
                      </td>
                      <td>
                        <code>
                          {license.bound_device_id
                            ? (data?.devices.find((device) => device.id === license.bound_device_id)
                                ?.install_id_hint ?? 'Bound')
                            : 'Unbound'}
                        </code>
                      </td>
                      <td>{formatDate(license.expires_at)}</td>
                      <td>
                        <div className="admin-actions">
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            onClick={() =>
                              void run('adjust_days', {
                                licenseId: license.id,
                                deltaDays: Number(adjustmentDays),
                              })
                            }
                          >
                            +{adjustmentDays}d
                          </button>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            onClick={() =>
                              void run('adjust_days', {
                                licenseId: license.id,
                                deltaDays: -Number(adjustmentDays),
                              })
                            }
                          >
                            −{adjustmentDays}d
                          </button>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            onClick={() => void run('reset_device', { licenseId: license.id })}
                          >
                            Reset device
                          </button>
                          {license.status !== 'revoked' && (
                            <button
                              type="button"
                              className="btn btn-ghost btn-sm danger-text"
                              onClick={() => void run('revoke', { licenseId: license.id })}
                            >
                              Revoke
                            </button>
                          )}
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
                          {license.bound_user_id && (
                            <button
                              type="button"
                              className="btn btn-ghost btn-sm"
                              onClick={() =>
                                void run(account?.status === 'banned' ? 'unban_user' : 'ban_user', {
                                  licenseId: license.id,
                                  userId: license.bound_user_id,
                                  reason: 'Account access suspended by administrator.',
                                })
                              }
                            >
                              {account?.status === 'banned' ? 'Unban' : 'Ban'}
                            </button>
                          )}
                        </div>
                        {transfer?.id === license.id && (
                          <form
                            className="admin-transfer-form"
                            onSubmit={(event) => void transferLicense(event)}
                          >
                            <input
                              className="form-input"
                              type="email"
                              required
                              placeholder="New account email"
                              value={transfer.email}
                              onChange={(event) =>
                                setTransfer({ ...transfer, email: event.target.value })
                              }
                            />
                            <input
                              className="form-input"
                              required={!transfer.preserveDevice}
                              minLength={16}
                              placeholder="Destination installation ID"
                              value={transfer.installId}
                              onChange={(event) =>
                                setTransfer({ ...transfer, installId: event.target.value })
                              }
                              disabled={transfer.preserveDevice}
                            />
                            <input
                              className="form-input"
                              required
                              minLength={12}
                              placeholder="Ownership verification note"
                              value={transfer.note}
                              onChange={(event) =>
                                setTransfer({ ...transfer, note: event.target.value })
                              }
                            />
                            <label>
                              <input
                                type="checkbox"
                                checked={transfer.preserveDevice}
                                onChange={(event) =>
                                  setTransfer({ ...transfer, preserveDevice: event.target.checked })
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
                <UserRound size={22} /> No license records match this search.
              </div>
            )}
          </div>
        </section>

        <section className="license-card admin-table-card">
          <div className="license-card-heading">
            <div>
              <span className="studio-eyebrow">DEVICE REGISTRY</span>
              <h2>Bound installations</h2>
            </div>
          </div>
          <div className="admin-device-list">
            {(data?.devices ?? []).map((device) => (
              <div className="admin-device-row" key={device.id}>
                <code>{device.install_id_hint}</code>
                <span>{accountByUser.get(device.user_id)?.email || device.user_id}</span>
                <span
                  className={`license-state-badge ${device.status === 'active' ? 'success' : device.status === 'banned' ? 'danger' : 'warning'}`}
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
                    void run(device.status === 'banned' ? 'unban_device' : 'ban_device', {
                      deviceId: device.id,
                      reason: 'Device access suspended by administrator.',
                    })
                  }
                >
                  {device.status === 'banned' ? 'Unban device' : 'Ban device'}
                </button>
              </div>
            ))}
            {(data?.devices ?? []).length === 0 && (
              <div className="admin-empty-state">
                <UserRound size={22} /> No device records yet.
              </div>
            )}
          </div>
        </section>

        <section className="license-card admin-audit-card">
          <div className="license-card-heading">
            <div>
              <span className="studio-eyebrow">AUDIT TRAIL</span>
              <h2>Recent administrative events</h2>
            </div>
          </div>
          <div className="admin-event-list">
            {(data?.events ?? []).slice(0, 12).map((event) => (
              <div key={event.id}>
                <code>{event.event_type}</code>
                <span>{formatDate(event.created_at)}</span>
                <small>{event.target_user_id || event.license_id || 'system'}</small>
              </div>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
};
