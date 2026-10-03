import React, { useState } from 'react';
import {
  type ProposedAction,
  type ProviderName,
  type SheetAudit,
  type AgentActivityEvent,
  type ExecutionPlan,
} from '../lib/agent-helper.js';
import type { Preview } from '@excel-agent/engine';
import { TypewriterText } from './TypewriterText.js';
import { MarkdownText } from './MarkdownText.js';

export interface ChatMessage {
  id: string;
  sender: 'user' | 'assistant';
  text: string;
  thought?: string;
  activities?: AgentActivityEvent[];
  plan?: ExecutionPlan;
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
  onSaveApiKey: (provider: ProviderName, key: string, baseUrl?: string) => void;
  onClearApiKey: () => void;
  onSendMessage: (query: string) => void;
  onApplyAction: (messageId: string, action: ProposedAction) => void;
  onApplyPlan?: (messageId: string, plan: ExecutionPlan) => void;
  onUndoLast: () => void;
  canUndo: boolean;
  onStop?: () => void;
  selectionContext?: import('../lib/selection-context.js').CellSelection | null;
  onClearSelectionContext?: () => void;
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
  onApplyPlan,
  onUndoLast,
  canUndo,
  onStop,
  selectionContext,
  onClearSelectionContext,
}) => {
  const [inputText, setInputText] = useState('');
  const messagesEndRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    messagesEndRef.current?.scrollIntoView?.({ behavior: 'smooth', block: 'end' });
  }, [messages.length, messages[messages.length - 1]?.text, isProcessing]);

  // BYOK setup state
  const [setupProvider, setSetupProvider] = useState<ProviderName>('groq');
  const [setupKey, setSetupKey] = useState('');
  const [setupBaseUrl, setSetupBaseUrl] = useState('');
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
    if (!setupKey.trim() && setupProvider !== 'custom') return;
    onSaveApiKey(
      setupProvider,
      setupKey.trim() || 'local-no-key',
      setupBaseUrl.trim() || undefined,
    );
  };

  const handleDemoKey = () => {
    onSaveApiKey('groq', 'demo-local-mode');
  };

  const getActivityIcon = (type: AgentActivityEvent['type']) => {
    switch (type) {
      case 'inspecting':
        return '🔍';
      case 'planning':
        return '📋';
      case 'guardrail_check':
        return '🛡️';
      case 'tool_call':
        return '⚡';
      case 'thinking':
        return '💭';
      default:
        return '✨';
    }
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
              Connect your AI provider to unlock conversational data engineering, multi-step
              planning, and deterministic Excel transformations directly in your browser.
            </p>

            <form onSubmit={handleActivateKey} className="byok-gate-form">
              <div className="form-group">
                <label className="form-label">Select AI Provider</label>
                <div
                  className="provider-chips"
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(80px, 1fr))',
                    gap: '4px',
                  }}
                >
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'groq' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('groq')}
                  >
                    Groq
                  </button>
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'gemini' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('gemini')}
                  >
                    Gemini
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
                    className={`provider-chip ${setupProvider === 'openai' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('openai')}
                  >
                    OpenAI
                  </button>
                  <button
                    type="button"
                    className={`provider-chip ${setupProvider === 'custom' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('custom')}
                  >
                    Ollama/Local
                  </button>
                </div>
              </div>

              {setupProvider === 'custom' && (
                <div className="form-group">
                  <label className="form-label">Endpoint Base URL</label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="http://localhost:11434/v1"
                    value={setupBaseUrl}
                    onChange={(e) => setSetupBaseUrl(e.target.value)}
                  />
                </div>
              )}

              <div className="form-group">
                <label className="form-label">
                  {setupProvider === 'groq'
                    ? 'Groq API Key (starts with gsk_)'
                    : setupProvider === 'gemini'
                      ? 'Google Gemini Key (AIzaSy...)'
                      : setupProvider === 'openrouter'
                        ? 'OpenRouter Key (sk-or-...)'
                        : setupProvider === 'openai'
                          ? 'OpenAI Key (sk-...)'
                          : 'API Key (Optional for Ollama)'}
                </label>
                <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                  <input
                    type={showKeyText ? 'text' : 'password'}
                    className="form-input"
                    style={{ width: '100%', paddingRight: '40px' }}
                    placeholder={
                      setupProvider === 'groq'
                        ? 'gsk_...'
                        : setupProvider === 'gemini'
                          ? 'AIzaSy...'
                          : setupProvider === 'openrouter'
                            ? 'sk-or-...'
                            : setupProvider === 'openai'
                              ? 'sk-...'
                              : 'ollama / none'
                    }
                    value={setupKey}
                    onChange={(e) => setSetupKey(e.target.value)}
                    required={setupProvider !== 'custom'}
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
                disabled={!setupKey.trim() && setupProvider !== 'custom'}
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
                        {msg.plan ? (
                          <span
                            className="op-badge"
                            style={{
                              background: '#faf5ff',
                              color: '#7e22ce',
                              borderColor: '#e9d5ff',
                            }}
                          >
                            {msg.plan.steps.length} Steps Plan
                          </span>
                        ) : msg.proposedAction ? (
                          <span className="op-badge">{msg.proposedAction.name}</span>
                        ) : msg.status === 'applied' ? (
                          <span
                            style={{ fontSize: '11px', color: 'var(--primary)', fontWeight: 600 }}
                          >
                            ✓ Applied
                          </span>
                        ) : null}
                      </div>

                      {/* Agent Activity Timeline */}
                      {msg.activities && msg.activities.length > 0 && (
                        <div className="activity-timeline">
                          {msg.activities.map((act, actIdx) => (
                            <div
                              key={act.id || actIdx}
                              className={`activity-pill activity-${act.type}`}
                            >
                              <span>{getActivityIcon(act.type)}</span>
                              <span>{act.summary}</span>
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Thought / CoT Accordion */}
                      {msg.thought && (
                        <details className="thought-box">
                          <summary className="thought-summary">
                            <span className="thought-icon">💡</span>
                            <span>Reasoning Process</span>
                          </summary>
                          <div className="thought-content">{msg.thought}</div>
                        </details>
                      )}

                      {/* Message Content */}
                      <div className="assistant-text">
                        {msg.isStreaming ? (
                          <div>
                            <MarkdownText text={msg.text} />
                            <span
                              className="activity-pulse-dot"
                              style={{
                                display: 'inline-block',
                                marginLeft: '4px',
                                verticalAlign: 'middle',
                              }}
                            />
                          </div>
                        ) : (
                          <TypewriterText text={msg.text} animate={isLatestAssistant} speed={10} />
                        )}
                      </div>

                      {/* Multi-Step Execution Plan Card */}
                      {msg.plan && (
                        <div className="plan-card">
                          <div className="plan-header">
                            <div>
                              <div className="plan-title">{msg.plan.title}</div>
                              <div className="plan-summary">{msg.plan.description}</div>
                            </div>
                            <span className={`plan-status-badge ${msg.plan.status}`}>
                              {msg.plan.status}
                            </span>
                          </div>

                          <div className="plan-steps-list">
                            {msg.plan.steps.map((step, sIdx) => (
                              <div key={step.id || sIdx} className="plan-step-item">
                                <div className="plan-step-num">{sIdx + 1}</div>
                                <div className="plan-step-body">
                                  <div className="plan-step-desc">{step.description}</div>
                                  <div className="plan-step-meta">
                                    <span
                                      className="op-badge"
                                      style={{ fontSize: '9.5px', padding: '1px 5px' }}
                                    >
                                      {step.operation}
                                    </span>
                                    <span
                                      className="plan-step-status"
                                      style={{
                                        color:
                                          step.status === 'completed'
                                            ? 'var(--accent-emerald)'
                                            : step.status === 'error'
                                              ? 'var(--accent-rose)'
                                              : 'var(--text-dim)',
                                      }}
                                    >
                                      {step.status === 'completed'
                                        ? '✓ Applied'
                                        : step.status === 'error'
                                          ? '✕ Failed'
                                          : 'Pending'}
                                    </span>
                                  </div>

                                  {step.preview?.changes && step.preview.changes.length > 0 && (
                                    <div className="plan-diff-preview">
                                      {step.preview.changes.slice(0, 2).map((ch, chIdx) => (
                                        <div key={chIdx}>
                                          <strong>
                                            {ch.location.column}
                                            {ch.location.row}:
                                          </strong>{' '}
                                          <span className="diff-del">
                                            {String(ch.before.value ?? '')}
                                          </span>{' '}
                                          →{' '}
                                          <span className="diff-ins">
                                            {String(ch.after.value ?? '')}
                                          </span>
                                        </div>
                                      ))}
                                    </div>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>

                          {msg.status === 'pending' && onApplyPlan && (
                            <div className="action-buttons-group">
                              <button
                                className="btn btn-primary btn-sm"
                                onClick={() => onApplyPlan(msg.id, msg.plan!)}
                                disabled={isProcessing}
                              >
                                Apply All {msg.plan.steps.length} Steps
                              </button>
                            </div>
                          )}
                        </div>
                      )}

                      {/* Single Operation Preview & Diff Card */}
                      {msg.proposedAction && msg.preview && !msg.plan && (
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
                      {msg.proposedAction && msg.status === 'pending' && !msg.plan && (
                        <div className="action-buttons-group">
                          <button
                            className="btn btn-primary btn-sm"
                            onClick={() => onApplyAction(msg.id, msg.proposedAction!)}
                            disabled={isProcessing}
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
            <div ref={messagesEndRef} />
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
            {selectionContext && (
              <div className="selection-context-chip">
                <span className="selection-context-label">{selectionContext.label}</span>
                <span className="selection-context-sheet">{selectionContext.sheetName}</span>
                <button
                  type="button"
                  className="selection-context-clear"
                  aria-label="Clear selection"
                  onClick={onClearSelectionContext}
                >
                  ×
                </button>
              </div>
            )}
            <input
              type="text"
              className="chat-input"
              placeholder="Ask ExcelAgento (e.g. 'Format dates in col C to YYYY-MM-DD')…"
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
            {isProcessing && onStop && (
              <button type="button" className="btn btn-secondary btn-sm" onClick={onStop}>
                Stop
              </button>
            )}
          </form>
        </>
      )}
    </aside>
  );
};
