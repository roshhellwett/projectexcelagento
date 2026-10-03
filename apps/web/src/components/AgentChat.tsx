import React, { useState } from 'react';
import { type ProposedAction, type ProviderName, type SheetAudit } from '../lib/agent-helper.js';
import type { Preview } from '@excel-agent/engine';
import { TypewriterText } from './TypewriterText.js';

export interface ChatMessage {
  id: string;
  sender: 'user' | 'assistant';
  text: string;
  /** The user request that produced this message, used for self-learning feedback. */
  sourceQuery?: string;
  proposedAction?: ProposedAction;
  preview?: Preview;
  status?: 'pending' | 'applied' | 'error';
  errorMessage?: string;
  isStreaming?: boolean;
}

interface AgentChatProps {
  audit: SheetAudit;
  messages: ChatMessage[];
  isProcessing: boolean;
  hasApiKey: boolean;
  apiKeyProvider: ProviderName;
  onSaveApiKey: (provider: ProviderName, key: string) => void;
  onClearApiKey: () => void;
  onSendMessage: (query: string) => void;
  onApplyAction: (messageId: string, action: ProposedAction) => void;
  onUndoLast: () => void;
  canUndo: boolean;
}

export const AgentChat: React.FC<AgentChatProps> = ({
  audit,
  messages,
  isProcessing,
  hasApiKey,
  apiKeyProvider,
  onSaveApiKey,
  onClearApiKey,
  onSendMessage,
  onApplyAction,
  onUndoLast,
  canUndo,
}) => {
  const [inputText, setInputText] = useState('');

  // BYOK setup state
  const [setupProvider, setSetupProvider] = useState<ProviderName>('groq');
  const [setupKey, setSetupKey] = useState('');
  const [showKeyText, setShowKeyText] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!inputText.trim() || isProcessing) return;
    onSendMessage(inputText);
    setInputText('');
  };

  const handleSuggestionClick = (prompt: string) => {
    if (isProcessing) return;
    onSendMessage(prompt);
  };

  const handleActivateKey = (e: React.FormEvent) => {
    e.preventDefault();
    if (!setupKey.trim()) return;
    onSaveApiKey(setupProvider, setupKey.trim());
  };

  const handleDemoKey = () => {
    onSaveApiKey('groq', 'demo-local-mode');
  };

  return (
    <aside className="agent-panel">
      {/* Agent Header */}
      <div className="agent-header">
        <div className="agent-title-group">
          <div className={`agent-status-dot ${hasApiKey ? 'active' : 'inactive'}`} />
          <div>
            <div className="agent-title">ExcelAgento</div>
            <div className="agent-subtitle">
              {hasApiKey ? (
                <span className="byok-active-tag">{apiKeyProvider.toUpperCase()} ACTIVE</span>
              ) : (
                'BYOK Key Required'
              )}
            </div>
          </div>
        </div>

        {hasApiKey && (
          <button
            className="btn btn-ghost btn-sm"
            style={{ fontSize: '11px', padding: '2px 8px' }}
            onClick={onClearApiKey}
            title="Change or disconnect API Key"
          >
            Change Key
          </button>
        )}
      </div>

      {/* GATED BYOK SETUP STATE */}
      {!hasApiKey ? (
        <div className="byok-gate-container">
          <div className="byok-gate-card">
            <div className="byok-avatar-icon">
              <svg
                width="20"
                height="20"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                <path d="M7 11V7a5 5 0 0 1 10 0v4" />
              </svg>
            </div>

            <h3 className="byok-gate-title">Activate Excel Agent</h3>
            <p className="byok-gate-desc">
              Enter your BYOK (Bring Your Own Key) to unlock intelligent conversational data
              cleaning, natural language Excel transformations, and deterministic engine edits.
            </p>

            <form onSubmit={handleActivateKey} className="byok-gate-form">
              <div className="form-group">
                <label className="form-label">Select AI Provider</label>
                <div className="provider-chips">
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'groq' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('groq')}
                  >
                    Groq (Ultra-Fast)
                  </button>
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'openrouter' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('openrouter')}
                  >
                    OpenRouter
                  </button>
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'gemini' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('gemini')}
                  >
                    Google Gemini
                  </button>
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">
                  {setupProvider === 'groq'
                    ? 'Groq API Key (starts with gsk_)'
                    : setupProvider === 'openrouter'
                      ? 'OpenRouter Key (starts with sk-or-)'
                      : 'Google Gemini Key'}
                </label>
                <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                  <input
                    type={showKeyText ? 'text' : 'password'}
                    className="form-input"
                    style={{ width: '100%', paddingRight: '40px' }}
                    placeholder={
                      setupProvider === 'groq'
                        ? 'gsk_...'
                        : setupProvider === 'openrouter'
                          ? 'sk-or-...'
                          : 'AIzaSy...'
                    }
                    value={setupKey}
                    onChange={(e) => setSetupKey(e.target.value)}
                    required
                  />
                  <button
                    type="button"
                    className="btn btn-ghost btn-sm"
                    style={{ position: 'absolute', right: '4px', height: '26px', padding: '0 6px' }}
                    onClick={() => setShowKeyText(!showKeyText)}
                  >
                    {showKeyText ? 'Hide' : 'Show'}
                  </button>
                </div>
              </div>

              <button
                type="submit"
                className="btn btn-primary"
                style={{ width: '100%', justifyContent: 'center', marginTop: '4px' }}
                disabled={!setupKey.trim()}
              >
                Activate Excel Agent
              </button>

              <div
                style={{
                  textAlign: 'center',
                  margin: '4px 0',
                  fontSize: '11px',
                  color: 'var(--text-dim)',
                }}
              >
                or
              </div>

              <button
                type="button"
                className="btn btn-secondary btn-sm"
                style={{ width: '100%', justifyContent: 'center' }}
                onClick={handleDemoKey}
              >
                Launch with Demo Mode (Instant)
              </button>
            </form>

            <div className="byok-privacy-callout">
              <strong>Local Privacy Guaranteed:</strong> Keys and spreadsheets are processed
              directly in your browser memory and never stored on remote servers.
            </div>
          </div>
        </div>
      ) : (
        /* UNLOCKED ACTIVE CHAT STATE */
        <>
          {/* Chat Messages */}
          <div className="chat-messages">
            {messages.map((msg, index) => {
              const isLatestAssistant = msg.sender === 'assistant' && index === messages.length - 1;

              return (
                <div key={msg.id} className={`chat-bubble ${msg.sender}`}>
                  {msg.sender === 'user' ? (
                    <div className="user-text-pill">{msg.text}</div>
                  ) : (
                    <div className="assistant-card">
                      <div className="assistant-header">
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span
                            style={{
                              width: '8px',
                              height: '8px',
                              borderRadius: '50%',
                              background: 'var(--accent-emerald)',
                              display: 'inline-block',
                            }}
                          />
                          <span
                            style={{ fontWeight: 600, color: 'var(--text-main)', fontSize: '13px' }}
                          >
                            Excel Agent
                          </span>
                        </div>
                        {msg.proposedAction ? (
                          <span className="op-badge">{msg.proposedAction.name}</span>
                        ) : msg.status === 'applied' ? (
                          <span
                            style={{ fontSize: '11px', color: 'var(--primary)', fontWeight: 600 }}
                          >
                            ✓ Applied
                          </span>
                        ) : null}
                      </div>

                      <div className="assistant-text">
                        <TypewriterText text={msg.text} animate={isLatestAssistant} speed={10} />
                      </div>

                      {/* Operation Preview & Diff Card */}
                      {msg.proposedAction && msg.preview && (
                        <div className="preview-summary-box">
                          <div className="preview-stat-row">
                            <span>Affected Cells:</span>
                            <strong>{msg.preview.affectedCells}</strong>
                          </div>

                          {msg.preview.warnings.length > 0 && (
                            <div
                              style={{
                                color: 'var(--accent-amber)',
                                fontSize: '11px',
                                marginTop: '4px',
                              }}
                            >
                              Notice: {msg.preview.warnings.map((w) => w.message).join('; ')}
                            </div>
                          )}

                          {msg.preview.changes.length > 0 && (
                            <table className="preview-diff-table">
                              <thead>
                                <tr>
                                  <th>Cell</th>
                                  <th>Before</th>
                                  <th>After</th>
                                </tr>
                              </thead>
                              <tbody>
                                {msg.preview.changes.slice(0, 4).map((ch, idx) => (
                                  <tr key={idx}>
                                    <td>
                                      {ch.location.column}
                                      {ch.location.row}
                                    </td>
                                    <td className="diff-del">{String(ch.before.value ?? '')}</td>
                                    <td className="diff-ins">{String(ch.after.value ?? '')}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>
                      )}

                      {msg.errorMessage && (
                        <div style={{ color: 'var(--accent-rose)', fontSize: '12px' }}>
                          ✕ Error: {msg.errorMessage}
                        </div>
                      )}

                      {/* Action CTA */}
                      {msg.proposedAction && msg.status === 'pending' && (
                        <div className="action-buttons-group">
                          <button
                            className="btn btn-primary btn-sm"
                            onClick={() => onApplyAction(msg.id, msg.proposedAction!)}
                          >
                            Apply Changes
                          </button>
                        </div>
                      )}

                      {msg.status === 'applied' && canUndo && (
                        <div className="action-buttons-group">
                          <button className="btn btn-ghost btn-sm" onClick={onUndoLast}>
                            Undo This Step
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Suggestions Drawer */}
          {audit.suggestions.length > 0 && (
            <div className="suggestions-drawer">
              <div className="suggestions-title">
                <svg
                  width="12"
                  height="12"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83" />
                </svg>
                Suggested Actions for {audit.sheetName}
              </div>
              <div className="chips-container">
                {audit.suggestions.map((s, idx) => (
                  <button
                    key={idx}
                    className="suggestion-chip"
                    onClick={() => handleSuggestionClick(s.prompt)}
                  >
                    {s.prompt}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Chat Input Bar */}
          <form className="chat-input-bar" onSubmit={handleSubmit}>
            <input
              type="text"
              className="chat-input"
              placeholder="Ask ExcelAgento (e.g. 'Format dates in col C to YYYY-MM-DD')..."
              value={inputText}
              onChange={(e) => setInputText(e.target.value)}
              disabled={isProcessing}
            />
            <button
              type="submit"
              className="btn btn-primary btn-sm"
              disabled={!inputText.trim() || isProcessing}
            >
              Send
            </button>
          </form>
        </>
      )}
    </aside>
  );
};
