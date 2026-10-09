import React, { useEffect, useRef, useState } from 'react';

import {
  forgetLearnedActions,
  getMemoryCloudStatus,
  learnedActionCount,
} from '../lib/agent-runtime.js';
import { useDialogA11y } from '../lib/use-dialog-a11y.js';
import {
  DEFAULT_OPENROUTER_MODEL,
  SUGGESTED_OPENROUTER_MODELS,
  type AgentSettings,
} from '../lib/settings.js';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AgentSettings;
  onSave: (settings: AgentSettings) => void;
  onClear: () => void;
}

interface DiagnosticDetails {
  authOk: boolean;
  modelReachable: boolean;
  toolEngagementOk: boolean;
  latencyMs?: number;
  tokensUsed?: number;
  toolCallSummary?: string;
  advice?: string;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  settings,
  onSave,
  onClear,
}) => {
  const [apiKey, setApiKey] = useState(settings.apiKey);
  const [model, setModel] = useState(settings.model?.trim() || DEFAULT_OPENROUTER_MODEL);
  const [showKey, setShowKey] = useState(false);
  const [savedSuccess, setSavedSuccess] = useState(false);
  const [testStatus, setTestStatus] = useState<{
    testing: boolean;
    message?: string;
    ok?: boolean;
    diagnostics?: DiagnosticDetails;
  }>({
    testing: false,
  });
  const [learnedCount, setLearnedCount] = useState(0);
  const cardRef = useRef<HTMLDivElement>(null);
  const connectionAbortRef = useRef<AbortController | null>(null);
  const closeTimerRef = useRef<number | null>(null);

  useDialogA11y(isOpen, cardRef, onClose);

  useEffect(() => {
    if (!isOpen) return;
    setApiKey(settings.apiKey);
    setModel(settings.model?.trim() || DEFAULT_OPENROUTER_MODEL);
    setShowKey(false);
    setSavedSuccess(false);
    setTestStatus({ testing: false });
    setLearnedCount(learnedActionCount());
    return () => {
      connectionAbortRef.current?.abort();
      connectionAbortRef.current = null;
    };
  }, [isOpen, settings]);

  useEffect(() => {
    if (isOpen) return;
    if (closeTimerRef.current !== null) {
      window.clearTimeout(closeTimerRef.current);
      closeTimerRef.current = null;
    }
  }, [isOpen]);

  useEffect(
    () => () => {
      if (closeTimerRef.current !== null) window.clearTimeout(closeTimerRef.current);
    },
    [],
  );

  if (!isOpen) return null;
  const cloudMemory = getMemoryCloudStatus();

  const handleSave = () => {
    const trimmedModel = model.trim() || DEFAULT_OPENROUTER_MODEL;
    onSave({
      provider: 'openrouter',
      apiKey: apiKey.trim(),
      model: trimmedModel,
    });
    setSavedSuccess(true);
    closeTimerRef.current = window.setTimeout(() => {
      closeTimerRef.current = null;
      onClose();
    }, 600);
  };

  const handleClear = () => {
    onClear();
    setApiKey('');
    setModel(DEFAULT_OPENROUTER_MODEL);
    setSavedSuccess(false);
    setTestStatus({ testing: false });
  };

  const handleForgetMemory = () => {
    forgetLearnedActions();
    setLearnedCount(0);
  };

  const handleTestConnection = async () => {
    const trimmedKey = apiKey.trim();
    const trimmedModel = model.trim() || DEFAULT_OPENROUTER_MODEL;

    if (!trimmedKey) {
      setTestStatus({
        testing: false,
        ok: false,
        message: 'Please enter your OpenRouter API key first.',
        diagnostics: {
          authOk: false,
          modelReachable: false,
          toolEngagementOk: false,
          advice: 'Obtain a key at https://openrouter.ai/keys',
        },
      });
      return;
    }

    setTestStatus({
      testing: true,
      message: 'Testing authentication, model access, and Excel spreadsheet tool engagement…',
    });

    const startTime = performance.now();
    const controller = new AbortController();
    connectionAbortRef.current?.abort();
    connectionAbortRef.current = controller;
    const timeout = window.setTimeout(() => controller.abort(), 30_000);

    try {
      // Execute a real spreadsheet tool-calling probe against OpenRouter
      const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        signal: controller.signal,
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${trimmedKey}`,
          'HTTP-Referer':
            typeof window !== 'undefined' ? window.location.origin : 'https://excelagento.local',
          'X-Title': 'ExcelAgento Connection Test',
        },
        body: JSON.stringify({
          model: trimmedModel,
          messages: [
            {
              role: 'system',
              content:
                'You are ExcelAgento, an autonomous Excel spreadsheet AI agent. When given a data formatting request, you MUST invoke the appropriate spreadsheet tool with correct arguments.',
            },
            {
              role: 'user',
              content:
                "In sheet 'Sales', column C contains dates in format 'MM/DD/YYYY'. Standardize them to 'YYYY-MM-DD' ISO format using the available tool.",
            },
          ],
          tools: [
            {
              type: 'function',
              function: {
                name: 'format_dates',
                description: 'Standardize dates in a spreadsheet column to a specified format.',
                parameters: {
                  type: 'object',
                  properties: {
                    sheet: { type: 'string', description: 'The target worksheet name' },
                    column: {
                      type: 'string',
                      description: 'The target column letter, for example C',
                    },
                    format: {
                      type: 'string',
                      description: 'Target date format, e.g. YYYY-MM-DD',
                      enum: ['YYYY-MM-DD', 'MM/DD/YYYY', 'DD/MM/YYYY'],
                    },
                  },
                  required: ['sheet', 'column', 'format'],
                },
              },
            },
          ],
          tool_choice: 'auto',
        }),
      });

      if (connectionAbortRef.current !== controller) return;
      const latencyMs = Math.round(performance.now() - startTime);

      if (!res.ok) {
        let errorDetail = '';
        try {
          const errJson = await res.json();
          errorDetail = errJson?.error?.message || JSON.stringify(errJson);
        } catch {
          errorDetail = (await res.text()).slice(0, 150);
        }

        let advice = 'Check your OpenRouter account dashboard.';
        if (res.status === 401) {
          advice = 'Authentication failed (HTTP 401). Verify your API key at openrouter.ai/keys.';
        } else if (res.status === 402) {
          advice =
            'Insufficient credits (HTTP 402). Your OpenRouter account requires credits. Top up at openrouter.ai/credits.';
        } else if (res.status === 404) {
          advice = `Model not found (HTTP 404). OpenRouter cannot find "${trimmedModel}". Verify the slug at openrouter.ai/models.`;
        } else if (res.status === 429) {
          advice =
            'Rate limit or quota reached (HTTP 429). The model is temporarily throttled or credits exhausted.';
        }

        setTestStatus({
          testing: false,
          ok: false,
          message: `HTTP ${res.status}: ${errorDetail || res.statusText}`,
          diagnostics: {
            authOk: res.status !== 401,
            modelReachable: res.status !== 404,
            toolEngagementOk: false,
            latencyMs,
            advice,
          },
        });
        return;
      }

      const data = await res.json();
      if (connectionAbortRef.current !== controller) return;
      const choice = data.choices?.[0];
      const message = choice?.message;
      const toolCalls = message?.tool_calls;
      const totalTokens = data.usage?.total_tokens;

      const formatDatesCall = Array.isArray(toolCalls)
        ? toolCalls.find(
            (c: { function?: { name?: string } }) => c?.function?.name === 'format_dates',
          )
        : undefined;

      let parsedArgs: Record<string, unknown> | null = null;
      if (formatDatesCall?.function?.arguments) {
        try {
          parsedArgs = JSON.parse(formatDatesCall.function.arguments);
        } catch {
          parsedArgs = null;
        }
      }

      const toolEngagementOk = Boolean(
        formatDatesCall &&
        parsedArgs &&
        typeof parsedArgs.sheet === 'string' &&
        parsedArgs.sheet.trim() !== '' &&
        typeof parsedArgs.column === 'string' &&
        /^[A-Za-z]+$/.test(parsedArgs.column) &&
        typeof parsedArgs.format === 'string' &&
        ['YYYY-MM-DD', 'MM/DD/YYYY', 'DD/MM/YYYY'].includes(parsedArgs.format),
      );

      if (toolEngagementOk) {
        const argSummary = parsedArgs
          ? `format_dates(${JSON.stringify(parsedArgs)})`
          : 'format_dates()';

        setTestStatus({
          testing: false,
          ok: true,
          message:
            'Full Excel Agent Capability Verified! Model successfully engaged with Excel function calling.',
          diagnostics: {
            authOk: true,
            modelReachable: true,
            toolEngagementOk: true,
            latencyMs,
            tokensUsed: totalTokens,
            toolCallSummary: argSummary,
            advice:
              'Model is responsive and supports tool calling for autonomous spreadsheet tasks.',
          },
        });
      } else {
        const textPreview = message?.content
          ? `"${message.content.slice(0, 100)}…"`
          : 'No response text';
        setTestStatus({
          testing: false,
          ok: false,
          message:
            'Connected to OpenRouter, but model did not call the Excel tool (returned plain text). ExcelAgento requires function calling.',
          diagnostics: {
            authOk: true,
            modelReachable: true,
            toolEngagementOk: false,
            latencyMs,
            tokensUsed: totalTokens,
            advice: `Model returned text instead of calling format_dates: ${textPreview}. For best spreadsheet performance, consider anthropic/claude-3.5-sonnet, deepseek/deepseek-chat, or google/gemini-2.5-pro.`,
          },
        });
      }
    } catch (err) {
      if (connectionAbortRef.current !== controller) return;
      if (controller.signal.aborted) {
        if (isOpen) {
          setTestStatus({
            testing: false,
            ok: false,
            message: 'Connection test timed out or was cancelled.',
            diagnostics: {
              authOk: false,
              modelReachable: false,
              toolEngagementOk: false,
              advice: 'Check your network connection and try again.',
            },
          });
        }
        return;
      }
      const raw = err instanceof Error ? err.message : String(err);
      const isNetworkBlock = /failed to fetch|networkerror|load failed/i.test(raw);
      setTestStatus({
        testing: false,
        ok: false,
        message: isNetworkBlock
          ? 'Could not reach the provider. Possible causes: the provider or its firewall blocked your network (try OpenRouter or Gemini), an ad-blocker is interfering, or you are offline.'
          : raw || 'Network connection failed.',
        diagnostics: {
          authOk: false,
          modelReachable: false,
          toolEngagementOk: false,
          advice: 'Check your internet connection, firewall settings, or ad-blocker extensions.',
        },
      });
    } finally {
      window.clearTimeout(timeout);
      if (connectionAbortRef.current === controller) connectionAbortRef.current = null;
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal-card"
        ref={cardRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-modal-title"
        data-dialog-open="true"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <div className="modal-title" id="settings-modal-title">
            Settings - Model Keys &amp; Providers (BYOK)
          </div>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={onClose}
            aria-label="Close settings"
          >
            ✕
          </button>
        </div>

        <div className="modal-body">
          <div className="settings-note">
            <strong>Browser-local engine with explicit provider context</strong>
            Excel operations run deterministically inside your browser via{' '}
            <code>@excel-agent/engine</code>. When a key is active, the prompt and the bounded
            workbook context needed for that request may be sent directly to the provider you
            selected. ExcelAgento does not use a shared model key or remote workbook storage.
          </div>

          <div className="form-group">
            <label className="form-label" htmlFor="settings-provider">
              AI Provider
            </label>
            <select
              id="settings-provider"
              data-autofocus
              className="select-input"
              value="openrouter"
              onChange={() => {}}
            >
              <option value="openrouter">OpenRouter (multi-model gateway)</option>
            </select>
            <div style={{ marginTop: '4px', fontSize: '11px', color: 'var(--text-muted)' }}>
              OpenRouter is the exclusive provider for ExcelAgento. Access Claude 3.5 Sonnet,
              DeepSeek V3, Gemini 2.5 Pro, Llama 3.3, and 300+ models with one key.
            </div>
          </div>

          <div className="form-group">
            <label className="form-label" htmlFor="settings-api-key">
              OpenRouter API Key
            </label>
            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
              <input
                id="settings-api-key"
                type={showKey ? 'text' : 'password'}
                className="form-input"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder="sk-or-v1-..."
              />
              <button
                type="button"
                className="btn btn-ghost btn-sm"
                aria-pressed={showKey}
                onClick={() => setShowKey((prev) => !prev)}
              >
                {showKey ? 'Hide' : 'Show'}
              </button>
            </div>
          </div>

          <div className="form-group">
            <label className="form-label" htmlFor="settings-model">
              Model ID (Manual Entry)
            </label>
            <input
              id="settings-model"
              type="text"
              className="form-input"
              value={model}
              onChange={(e) => setModel(e.target.value)}
              placeholder="e.g. anthropic/claude-3.5-sonnet"
            />
            <div style={{ marginTop: '6px', fontSize: '12px', color: 'var(--text-muted)' }}>
              Enter any OpenRouter model slug. Model fallback is disabled; your specified model is
              strictly used, and any errors or limit issues will be surfaced directly so you can
              address them.
            </div>
            <div
              style={{
                display: 'flex',
                flexWrap: 'wrap',
                gap: '6px',
                marginTop: '8px',
                alignItems: 'center',
              }}
            >
              <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>Suggestions:</span>
              {SUGGESTED_OPENROUTER_MODELS.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  className="btn btn-ghost btn-sm"
                  style={{
                    fontSize: '11px',
                    padding: '2px 8px',
                    height: 'auto',
                    borderRadius: '12px',
                    background:
                      model === item.id
                        ? 'var(--primary-subtle, rgba(16, 185, 129, 0.15))'
                        : 'rgba(255, 255, 255, 0.05)',
                    border:
                      model === item.id
                        ? '1px solid var(--primary)'
                        : '1px solid rgba(255, 255, 255, 0.1)',
                  }}
                  onClick={() => setModel(item.id)}
                >
                  {item.label}
                </button>
              ))}
            </div>
          </div>

          <div style={{ marginTop: '12px' }}>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
              }}
            >
              <button
                type="button"
                className="btn btn-secondary btn-sm"
                onClick={handleTestConnection}
                disabled={testStatus.testing}
              >
                {testStatus.testing ? 'Testing Excel Capability…' : 'Test Connection'}
              </button>
              {testStatus.message && (
                <span
                  style={{
                    fontSize: '12px',
                    fontWeight: 500,
                    color: testStatus.ok ? 'var(--accent-emerald)' : 'var(--accent-rose)',
                    maxWidth: '300px',
                    textAlign: 'right',
                  }}
                >
                  {testStatus.message}
                </span>
              )}
            </div>

            {testStatus.diagnostics && (
              <div
                style={{
                  marginTop: '10px',
                  padding: '10px 12px',
                  borderRadius: '6px',
                  fontSize: '12px',
                  background: testStatus.ok
                    ? 'rgba(16, 185, 129, 0.08)'
                    : 'rgba(244, 63, 94, 0.08)',
                  border: `1px solid ${testStatus.ok ? 'rgba(16, 185, 129, 0.25)' : 'rgba(244, 63, 94, 0.25)'}`,
                }}
              >
                <div
                  style={{
                    fontWeight: 600,
                    marginBottom: '6px',
                    display: 'flex',
                    justifyContent: 'space-between',
                  }}
                >
                  <span>Diagnostic Health Check</span>
                  {testStatus.diagnostics.latencyMs !== undefined && (
                    <span style={{ fontWeight: 400, color: 'var(--text-muted)' }}>
                      Latency: {testStatus.diagnostics.latencyMs} ms
                    </span>
                  )}
                </div>
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
                    gap: '6px',
                    marginBottom: '8px',
                  }}
                >
                  <div>
                    API Auth:{' '}
                    <strong>{testStatus.diagnostics.authOk ? '✓ Passed' : '✗ Failed'}</strong>
                  </div>
                  <div>
                    Model Reachable:{' '}
                    <strong>
                      {testStatus.diagnostics.modelReachable ? '✓ Passed' : '✗ Failed'}
                    </strong>
                  </div>
                  <div>
                    Excel Tool Calling:{' '}
                    <strong
                      style={{
                        color: testStatus.diagnostics.toolEngagementOk
                          ? 'var(--accent-emerald)'
                          : 'var(--accent-rose)',
                      }}
                    >
                      {testStatus.diagnostics.toolEngagementOk ? '✓ Verified' : '✗ Not Engaging'}
                    </strong>
                  </div>
                </div>
                {testStatus.diagnostics.toolCallSummary && (
                  <div
                    style={{ fontSize: '11px', color: 'var(--text-muted)', marginBottom: '4px' }}
                  >
                    Invoked Tool: <code>{testStatus.diagnostics.toolCallSummary}</code>
                  </div>
                )}
                {testStatus.diagnostics.advice && (
                  <div
                    style={{
                      fontSize: '11px',
                      color: testStatus.ok ? 'var(--text-muted)' : 'var(--accent-rose)',
                    }}
                  >
                    {testStatus.diagnostics.advice}
                  </div>
                )}
              </div>
            )}
          </div>

          <div
            className="settings-memory-box"
            style={{
              marginTop: '16px',
              padding: '12px',
              background: 'rgba(255, 255, 255, 0.04)',
              borderRadius: '6px',
              border: '1px solid rgba(255, 255, 255, 0.08)',
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                marginBottom: '6px',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span
                  style={{
                    display: 'inline-block',
                    width: '8px',
                    height: '8px',
                    borderRadius: '50%',
                    background: 'var(--brand-emerald)',
                    boxShadow: '0 0 8px rgba(16, 185, 129, 0.6)',
                  }}
                />
                <span style={{ fontSize: '13px', fontWeight: 600 }}>Learned action memory</span>
              </div>
              <span
                style={{
                  fontSize: '11px',
                  color: 'var(--brand-emerald)',
                  background: 'rgba(16, 185, 129, 0.12)',
                  padding: '2px 8px',
                  borderRadius: '4px',
                  border: '1px solid rgba(16, 185, 129, 0.25)',
                }}
              >
                {cloudMemory.enabled ? 'Cloud configured · not connection-tested' : 'Local only'}
              </span>
            </div>
            <div
              style={{
                fontSize: '12px',
                color: 'var(--text-muted)',
                marginBottom: '8px',
                lineHeight: 1.4,
              }}
            >
              {cloudMemory.enabled
                ? 'Optional synchronization sends queries, learned operation arguments, and working-step context to the deployment-configured backend. Clearing local actions does not delete cloud records.'
                : 'Verified actions stay in this browser. Cloud synchronization is disabled for this deployment.'}
            </div>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
              }}
            >
              <span style={{ fontSize: '12px' }}>
                Learned actions: <strong>{learnedCount}</strong> active pattern
                {learnedCount === 1 ? '' : 's'}
              </span>
              <button
                type="button"
                className="btn btn-ghost btn-sm"
                onClick={handleForgetMemory}
                disabled={learnedCount === 0}
              >
                Forget local actions
              </button>
            </div>
          </div>

          {savedSuccess && (
            <div
              style={{
                color: 'var(--primary)',
                fontSize: '13px',
                fontWeight: 600,
                marginTop: '8px',
              }}
            >
              ✓ Settings saved to this browser
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button type="button" className="btn btn-ghost btn-sm" onClick={handleClear}>
            Clear Key
          </button>
          <button type="button" className="btn btn-primary btn-sm" onClick={handleSave}>
            Save Preferences
          </button>
        </div>
      </div>
    </div>
  );
};
