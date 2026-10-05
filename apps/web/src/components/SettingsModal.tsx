import React, { useEffect, useRef, useState } from 'react';

import {
  forgetLearnedActions,
  getMemoryCloudStatus,
  learnedActionCount,
} from '../lib/agent-runtime.js';
import { useDialogA11y } from '../lib/use-dialog-a11y.js';
import {
  AVAILABLE_MODELS,
  PROVIDER_LABELS,
  defaultModelFor,
  type AgentSettings,
  type ProviderName,
} from '../lib/settings.js';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AgentSettings;
  onSave: (settings: AgentSettings) => void;
  onClear: () => void;
}

const PROVIDERS: ProviderName[] = ['groq', 'gemini', 'openrouter', 'openai', 'custom'];

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  settings,
  onSave,
  onClear,
}) => {
  const [provider, setProvider] = useState<ProviderName>(settings.provider);
  const [apiKey, setApiKey] = useState(settings.apiKey);
  const [model, setModel] = useState(settings.model ?? '');
  const [baseUrl, setBaseUrl] = useState(settings.baseUrl ?? '');
  const [showKey, setShowKey] = useState(false);
  const [savedSuccess, setSavedSuccess] = useState(false);
  const [testStatus, setTestStatus] = useState<{
    testing: boolean;
    message?: string;
    ok?: boolean;
  }>({
    testing: false,
  });
  const [learnedCount, setLearnedCount] = useState(0);
  const cardRef = useRef<HTMLDivElement>(null);

  useDialogA11y(isOpen, cardRef, onClose);

  useEffect(() => {
    if (!isOpen) return;
    setProvider(settings.provider);
    setApiKey(settings.apiKey);
    setModel(settings.model ?? '');
    setBaseUrl(settings.baseUrl ?? '');
    setShowKey(false);
    setSavedSuccess(false);
    setTestStatus({ testing: false });
    setLearnedCount(learnedActionCount());
  }, [isOpen, settings]);

  if (!isOpen) return null;
  const cloudMemory = getMemoryCloudStatus();

  const handleProviderChange = (newProvider: ProviderName) => {
    setProvider(newProvider);
    const available = AVAILABLE_MODELS[newProvider];
    if (available && available[0]) {
      setModel(available[0].id);
    } else {
      setModel('');
    }
  };

  const handleSave = () => {
    const trimmedModel = model.trim();
    const trimmedBaseUrl = baseUrl.trim();
    onSave({
      provider,
      apiKey: apiKey.trim(),
      ...(trimmedModel ? { model: trimmedModel } : {}),
      ...(trimmedBaseUrl ? { baseUrl: trimmedBaseUrl } : {}),
    });
    setSavedSuccess(true);
    setTimeout(onClose, 600);
  };

  const handleClear = () => {
    onClear();
    setApiKey('');
    setModel('');
    setBaseUrl('');
    setSavedSuccess(false);
  };

  const handleForgetMemory = () => {
    forgetLearnedActions();
    setLearnedCount(0);
  };

  const handleTestConnection = async () => {
    if (!apiKey.trim() && provider !== 'custom') {
      setTestStatus({ testing: false, ok: false, message: 'Please enter an API key first.' });
      return;
    }
    setTestStatus({ testing: true, message: 'Testing connection…' });
    try {
      const endpoint =
        provider === 'groq'
          ? 'https://api.groq.com/openai/v1/models'
          : provider === 'openai'
            ? 'https://api.openai.com/v1/models'
            : provider === 'openrouter'
              ? 'https://openrouter.ai/api/v1/models'
              : provider === 'custom'
                ? `${baseUrl.replace(/\/+$/, '') || 'http://localhost:11434/v1'}/models`
                : 'https://generativelanguage.googleapis.com/v1beta/models';

      const headers: Record<string, string> = {};
      if (provider === 'gemini') {
        headers['x-goog-api-key'] = apiKey.trim();
      } else {
        headers.Authorization = `Bearer ${apiKey.trim()}`;
      }

      const res = await fetch(endpoint, { method: 'GET', headers });
      if (res.ok) {
        setTestStatus({
          testing: false,
          ok: true,
          message: 'Connection successful! Models verified.',
        });
      } else {
        const text = await res.text();
        setTestStatus({
          testing: false,
          ok: false,
          message: `HTTP ${res.status}: ${text.slice(0, 100) || res.statusText}`,
        });
      }
    } catch (err) {
      const raw = err instanceof Error ? err.message : '';
      const isNetworkBlock = /failed to fetch|networkerror|load failed/i.test(raw);
      setTestStatus({
        testing: false,
        ok: false,
        message: isNetworkBlock
          ? 'Could not reach the provider. Possible causes: the provider or its firewall blocked your network (try OpenRouter or Gemini), an ad-blocker is interfering, or you are offline.'
          : raw || 'Network connection failed.',
      });
    }
  };

  const modelsForCurrentProvider = AVAILABLE_MODELS[provider] || [];

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
            Settings - Model Keys & Providers (BYOK)
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
            <strong>Enterprise-Grade, Privacy-Preserving Architecture</strong>
            Excel operations run deterministically inside your browser via{' '}
            <code>@excel-agent/engine</code>. Only minimal structural column profiles are sent to
            the AI reasoning model. Your spreadsheet data stays on your machine.
          </div>

          <div className="form-group">
            <label className="form-label" htmlFor="settings-provider">
              AI Provider
            </label>
            <select
              id="settings-provider"
              data-autofocus
              className="select-input"
              value={provider}
              onChange={(e) => handleProviderChange(e.target.value as ProviderName)}
            >
              {PROVIDERS.map((name) => (
                <option key={name} value={name}>
                  {PROVIDER_LABELS[name]}
                </option>
              ))}
            </select>
          </div>

          {provider === 'custom' && (
            <div className="form-group">
              <label className="form-label" htmlFor="settings-base-url">
                Custom API Base URL (Ollama, LM Studio, vLLM)
              </label>
              <input
                id="settings-base-url"
                type="text"
                className="form-input"
                value={baseUrl}
                onChange={(e) => setBaseUrl(e.target.value)}
                placeholder="http://localhost:11434/v1"
              />
            </div>
          )}

          <div className="form-group">
            <label className="form-label" htmlFor="settings-api-key">
              API Key {provider === 'custom' ? '(optional for local)' : ''}
            </label>
            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
              <input
                id="settings-api-key"
                type={showKey ? 'text' : 'password'}
                className="form-input"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder={
                  provider === 'groq'
                    ? 'gsk_...'
                    : provider === 'openrouter'
                      ? 'sk-or-...'
                      : provider === 'openai'
                        ? 'sk-proj-...'
                        : provider === 'custom'
                          ? 'optional or sk-...'
                          : 'AIzaSy...'
                }
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
              Model Selection
            </label>
            {modelsForCurrentProvider.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <select
                  id="settings-model"
                  className="select-input"
                  value={
                    modelsForCurrentProvider.some((m) => m.id === model) ? model : 'custom_override'
                  }
                  onChange={(e) => {
                    if (e.target.value !== 'custom_override') {
                      setModel(e.target.value);
                    }
                  }}
                >
                  {modelsForCurrentProvider.map((m) => (
                    <option key={m.id} value={m.id}>
                      {m.label}
                    </option>
                  ))}
                  <option value="custom_override">Custom Model ID…</option>
                </select>

                {(!modelsForCurrentProvider.some((m) => m.id === model) ||
                  model === 'custom_override') && (
                  <>
                    <label className="sr-only" htmlFor="settings-model-override">
                      Custom model ID
                    </label>
                    <input
                      id="settings-model-override"
                      type="text"
                      className="form-input"
                      value={model === 'custom_override' ? '' : model}
                      onChange={(e) => setModel(e.target.value)}
                      placeholder="Enter custom model ID (e.g. qwen/qwen3.8-27b)"
                    />
                  </>
                )}
              </div>
            ) : (
              <input
                id="settings-model"
                type="text"
                className="form-input"
                value={model}
                onChange={(e) => setModel(e.target.value)}
                placeholder={defaultModelFor(provider)}
              />
            )}
          </div>

          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              marginTop: '8px',
            }}
          >
            <button
              type="button"
              className="btn btn-secondary btn-sm"
              onClick={handleTestConnection}
              disabled={testStatus.testing}
            >
              {testStatus.testing ? 'Testing…' : 'Test Connection'}
            </button>
            {testStatus.message && (
              <span
                style={{
                  fontSize: '12px',
                  fontWeight: 500,
                  color: testStatus.ok ? 'var(--accent-emerald)' : 'var(--accent-rose)',
                }}
              >
                {testStatus.message}
              </span>
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
