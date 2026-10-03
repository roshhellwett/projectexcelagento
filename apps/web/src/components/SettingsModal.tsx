import React, { useEffect, useState } from 'react';

import { forgetLearnedActions, learnedActionCount } from '../lib/agent-runtime.js';
import {
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

const PROVIDERS: ProviderName[] = ['groq', 'openrouter', 'gemini'];

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
  const [showKey, setShowKey] = useState(false);
  const [savedSuccess, setSavedSuccess] = useState(false);
  const [learnedCount, setLearnedCount] = useState(0);

  useEffect(() => {
    if (!isOpen) return;
    setProvider(settings.provider);
    setApiKey(settings.apiKey);
    setModel(settings.model ?? '');
    setShowKey(false);
    setSavedSuccess(false);
    setLearnedCount(learnedActionCount());
  }, [isOpen, settings]);

  if (!isOpen) return null;

  const handleSave = () => {
    const trimmedModel = model.trim();
    onSave({
      provider,
      apiKey: apiKey.trim(),
      ...(trimmedModel ? { model: trimmedModel } : {}),
    });
    setSavedSuccess(true);
    setTimeout(onClose, 600);
  };

  const handleClear = () => {
    onClear();
    setApiKey('');
    setModel('');
    setSavedSuccess(false);
  };

  const handleForgetMemory = () => {
    forgetLearnedActions();
    setLearnedCount(0);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <div className="modal-title">Settings - Model Keys (BYOK)</div>
          <button className="btn btn-ghost btn-sm" onClick={onClose} aria-label="Close settings">
            ✕
          </button>
        </div>

        <div className="modal-body">
          <div className="settings-note">
            <strong>Local, privacy-preserving architecture</strong>
            Every transformation runs 100% deterministically in your browser via{' '}
            <code>@excel-agent/engine</code>. Full spreadsheets never leave your machine - only a
            compact column profile is sent to your BYOK model. Keys are stored only in this browser.
          </div>

          <div className="form-group">
            <label className="form-label">AI Provider</label>
            <select
              className="select-input"
              value={provider}
              onChange={(e) => setProvider(e.target.value as ProviderName)}
            >
              {PROVIDERS.map((name) => (
                <option key={name} value={name}>
                  {PROVIDER_LABELS[name]}
                </option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label className="form-label">API Key</label>
            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
              <input
                type={showKey ? 'text' : 'password'}
                className="form-input"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder={
                  provider === 'groq'
                    ? 'gsk_...'
                    : provider === 'openrouter'
                      ? 'sk-or-...'
                      : 'AIzaSy...'
                }
              />
              <button
                type="button"
                className="btn btn-ghost btn-sm"
                onClick={() => setShowKey((prev) => !prev)}
              >
                {showKey ? 'Hide' : 'Show'}
              </button>
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Model Override (optional)</label>
            <input
              type="text"
              className="form-input"
              value={model}
              onChange={(e) => setModel(e.target.value)}
              placeholder={defaultModelFor(provider)}
            />
          </div>

          <div className="settings-memory-row">
            <span>
              Self-learning memory: <strong>{learnedCount}</strong> verified action
              {learnedCount === 1 ? '' : 's'}
            </span>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              onClick={handleForgetMemory}
              disabled={learnedCount === 0}
            >
              Forget learned actions
            </button>
          </div>

          {savedSuccess && (
            <div style={{ color: 'var(--primary)', fontSize: '13px', fontWeight: 600 }}>
              ✓ Settings saved to this browser
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button className="btn btn-ghost btn-sm" onClick={handleClear}>
            Clear Key
          </button>
          <button className="btn btn-primary btn-sm" onClick={handleSave}>
            Save Preferences
          </button>
        </div>
      </div>
    </div>
  );
};
