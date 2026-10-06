import React from 'react';
import { BarChart2, ArrowLeft, Sparkles, ExternalLink, Sliders } from 'lucide-react';

import {
  PROVIDER_HOSTS,
  formatTokenCount,
  maskApiKey,
  summarizeUsage,
  type UsageEntry,
} from '../lib/usage.js';
import {
  PROVIDER_LABELS,
  defaultModelFor,
  isDemoKey,
  type AgentSettings,
} from '../lib/settings.js';
import { ResourcePageChrome, ResourcePageFooter } from './ResourcePageChrome.js';

export interface ModelUsagePageProps {
  settings: AgentSettings;
  entries: UsageEntry[];
  learnedActions: number;
  engineOperations: number;
  toolCount: number;
  workbookSummary: {
    fileName: string;
    sheetName: string;
    sheets: number;
    rows: number;
    cols: number;
  };
  onBack: () => void;
  onOpenSettings: () => void;
  onClearUsage: () => void;
  onForgetLearned: () => void;
}

const StatCard: React.FC<{ label: string; value: string; hint?: string; testId?: string }> = ({
  label,
  value,
  hint,
  testId,
}) => (
  <div className="usage-stat-card">
    <span className="usage-stat-label">{label}</span>
    <span className="usage-stat-value" data-testid={testId}>
      {value}
    </span>
    {hint ? <span className="usage-stat-hint">{hint}</span> : null}
  </div>
);

function formatTimestamp(timestamp: number): string {
  try {
    return new Date(timestamp).toLocaleString();
  } catch {
    return '—';
  }
}

function formatLatency(ms: number | undefined | null): string {
  if (typeof ms !== 'number' || isNaN(ms) || ms <= 0) return '—';
  return ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${Math.round(ms)} ms`;
}

export const ModelUsagePage: React.FC<ModelUsagePageProps> = ({
  settings,
  entries,
  learnedActions,
  engineOperations,
  toolCount,
  workbookSummary,
  onBack,
  onOpenSettings,
  onClearUsage,
  onForgetLearned,
}) => {
  const summary = summarizeUsage(entries);
  const keyConfigured = settings.apiKey.trim().length > 0 && !isDemoKey(settings.apiKey);
  const effectiveModel = settings.model?.trim() || defaultModelFor(settings.provider);
  const recent = [...entries].reverse().slice(0, 50);

  return (
    <div
      className="usage-page docs-page resource-page resource-page-usage"
      data-testid="usage-page"
    >
      <ResourcePageChrome current="usage" onBack={onBack} />
      {/* Unified Top Bar with Zenith OS Branding */}
      <div className="docs-top-bar">
        <button className="btn btn-secondary btn-sm" onClick={onBack} title="Return to Workspace">
          <ArrowLeft size={14} /> Back to workspace
        </button>

        <a
          href="https://zenithopensourceprojects.vercel.app/os"
          target="_blank"
          rel="noopener noreferrer"
          className="zenith-brand-badge"
          title="Visit Zenith Open Source Projects Official Hub"
        >
          <Sparkles size={13} className="zenith-sparkle-icon" />
          <span>Zenith Open Source Projects (zenithopensourceprojects.vercel.app/os)</span>
          <ExternalLink size={12} />
        </a>
      </div>

      <header className="docs-header usage-header">
        <div className="docs-badge">
          <BarChart2 size={14} />
          <span>TELEMETRY &amp; REAL-TIME TOKEN AUDIT LEDGER</span>
        </div>
        <h1 className="docs-title usage-title">
          Model &amp; <span className="text-emerald">Usage</span>
        </h1>
        <p className="docs-subtitle usage-subtitle">
          Every inference call this workspace records locally, with provider-reported token counts
          when available. This is a local diagnostic ledger, not remote telemetry; missing usage is
          shown as unknown and connected providers receive the context described on the Privacy
          page.
        </p>
        <div
          className="usage-header-actions"
          style={{ justifyContent: 'center', marginTop: '16px' }}
        >
          <button type="button" className="btn btn-secondary btn-sm" onClick={onOpenSettings}>
            <Sliders size={13} /> Model settings
          </button>
        </div>
      </header>

      <section className="usage-panel" aria-label="Active model configuration">
        <h2 className="usage-panel-title">Active configuration</h2>
        <div className="usage-grid">
          <StatCard
            label="Provider"
            value={PROVIDER_LABELS[settings.provider]}
            hint={PROVIDER_HOSTS[settings.provider]}
            testId="usage-provider"
          />
          <StatCard
            label="Active model"
            value={effectiveModel}
            hint={settings.model?.trim() ? 'Manual override' : 'Provider default'}
            testId="usage-model"
          />
          <StatCard
            label="API key"
            value={keyConfigured ? 'Active' : 'Not set'}
            hint={maskApiKey(settings.apiKey)}
            testId="usage-key-status"
          />
          <StatCard
            label="Inference mode"
            value={keyConfigured ? 'BYOK model + guardrail' : 'Local deterministic engine'}
            hint={
              keyConfigured
                ? 'Prompt, profiles and requested read results may be sent'
                : 'Zero tokens, offline'
            }
            testId="usage-mode"
          />
        </div>
      </section>

      <section className="usage-panel" aria-label="Workspace and runtime">
        <h2 className="usage-panel-title">Workspace &amp; runtime</h2>
        <div className="usage-grid">
          <StatCard
            label="Workbook"
            value={workbookSummary.fileName}
            hint={`${workbookSummary.sheets} sheet(s) · ${workbookSummary.sheetName} · ${workbookSummary.rows} rows × ${workbookSummary.cols} cols`}
          />
          <StatCard
            label="Engine operations"
            value={`${engineOperations}`}
            hint="Registered, schema-validated operations"
          />
          <StatCard
            label="Model tool contracts"
            value={`${toolCount}`}
            hint="Derived live from the engine registry"
          />
          <StatCard
            label="Verified learned patterns"
            value={`${learnedActions}`}
            hint="Local associations; optional cloud synchronization is deployment-configured"
            testId="usage-learned-actions"
          />
        </div>
        <div className="usage-actions">
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={onForgetLearned}
            disabled={learnedActions === 0}
          >
            Forget local actions
          </button>
        </div>
      </section>

      <section className="usage-panel" aria-label="Token usage">
        <h2 className="usage-panel-title">Token usage</h2>
        <div className="usage-grid">
          <StatCard
            label="Requests handled"
            value={formatTokenCount(summary.requests)}
            hint={`${summary.llmRequests} via model · ${summary.localRequests} local`}
            testId="usage-requests"
          />
          <StatCard
            label="Total tokens"
            value={formatTokenCount(summary.totalTokens)}
            hint={
              summary.unreportedUsageRequests
                ? `Reported subtotal · ${summary.unreportedUsageRequests} turn(s) have unknown usage`
                : 'Prompt + completion, provider-reported'
            }
            testId="usage-total-tokens"
          />
          <StatCard
            label="Prompt tokens"
            value={formatTokenCount(summary.promptTokens)}
            testId="usage-prompt-tokens"
          />
          <StatCard
            label="Completion tokens"
            value={formatTokenCount(summary.completionTokens)}
            testId="usage-completion-tokens"
          />
          <StatCard
            label="Failed calls"
            value={formatTokenCount(summary.failures)}
            testId="usage-failures"
          />
          <StatCard
            label="Average latency"
            value={formatLatency(summary.avgLatencyMs)}
            hint="Model calls only"
            testId="usage-avg-latency"
          />
        </div>
        {summary.lastRequestAt ? (
          <p className="usage-meta">Last request: {formatTimestamp(summary.lastRequestAt)}</p>
        ) : null}
        <div className="usage-actions">
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={onClearUsage}
            disabled={summary.requests === 0}
            data-testid="usage-clear"
          >
            Clear usage history
          </button>
        </div>
      </section>

      {summary.byProvider.length > 0 ? (
        <section className="usage-panel" aria-label="Usage breakdown">
          <h2 className="usage-panel-title">Breakdown</h2>
          <div className="usage-columns">
            <table className="usage-table">
              <caption>By provider</caption>
              <thead>
                <tr>
                  <th scope="col">Provider</th>
                  <th scope="col">Requests</th>
                  <th scope="col">Tokens</th>
                </tr>
              </thead>
              <tbody>
                {summary.byProvider.map((bucket) => (
                  <tr key={`provider-${bucket.label}`}>
                    <td>{bucket.label}</td>
                    <td>{bucket.requests}</td>
                    <td>{formatTokenCount(bucket.totalTokens)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <table className="usage-table">
              <caption>By model</caption>
              <thead>
                <tr>
                  <th scope="col">Model</th>
                  <th scope="col">Requests</th>
                  <th scope="col">Tokens</th>
                </tr>
              </thead>
              <tbody>
                {summary.byModel.map((bucket) => (
                  <tr key={`model-${bucket.label}`}>
                    <td>{bucket.label}</td>
                    <td>{bucket.requests}</td>
                    <td>{formatTokenCount(bucket.totalTokens)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : null}

      <section className="usage-panel" aria-label="Recent requests">
        <h2 className="usage-panel-title">Recent requests</h2>
        {recent.length === 0 ? (
          <p className="usage-empty" data-testid="usage-empty">
            No inference calls recorded yet. Send a message in the workspace and its token counts
            will appear here.
          </p>
        ) : (
          <div className="usage-table-wrap">
            <table className="usage-table">
              <thead>
                <tr>
                  <th scope="col">Time</th>
                  <th scope="col">Request</th>
                  <th scope="col">Source</th>
                  <th scope="col">Provider</th>
                  <th scope="col">Model</th>
                  <th scope="col">Prompt</th>
                  <th scope="col">Completion</th>
                  <th scope="col">Total</th>
                  <th scope="col">Latency</th>
                  <th scope="col">Status</th>
                </tr>
              </thead>
              <tbody>
                {recent.map((entry) => (
                  <tr key={entry.id} data-testid="usage-log-row">
                    <td>{formatTimestamp(entry.timestamp)}</td>
                    <td className="usage-query" title={entry.query}>
                      {entry.query || '—'}
                    </td>
                    <td>{entry.source}</td>
                    <td>{entry.provider}</td>
                    <td>{entry.model}</td>
                    <td>
                      {entry.tokenUsageReported === false
                        ? 'Unknown'
                        : formatTokenCount(entry.promptTokens)}
                    </td>
                    <td>
                      {entry.tokenUsageReported === false
                        ? 'Unknown'
                        : formatTokenCount(entry.completionTokens)}
                    </td>
                    <td>
                      {entry.tokenUsageReported === false
                        ? 'Unknown'
                        : formatTokenCount(entry.totalTokens)}
                    </td>
                    <td>{formatLatency(entry.latencyMs)}</td>
                    <td>
                      <span
                        className={
                          entry.ok ? 'usage-badge usage-badge-ok' : 'usage-badge usage-badge-fail'
                        }
                      >
                        {entry.ok ? 'OK' : 'Failed'}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      <section className="usage-note">
        <strong>Data &amp; usage:</strong> local mode needs no model request. Connected models
        receive your prompt, recent chat context, workbook profiles and examples, and requested
        read-tool results. Optional cloud memory can also send queries and operation arguments to
        the configured backend. Reported token totals include completed inference calls across the
        turn; missing usage is shown as unknown, not free. Failed upstream calls may incur costs the
        provider does not report. See the Privacy page before working with sensitive data.
      </section>
      <ResourcePageFooter current="usage" />
    </div>
  );
};
