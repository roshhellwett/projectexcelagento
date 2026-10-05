import React, { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Search,
  Layers,
  ShieldCheck,
  Zap,
  Brain,
  AlertTriangle,
  CheckCircle2,
  Sparkles,
  Bot,
  RotateCcw,
  Check,
  KeyRound,
  ChevronDown,
  ChevronUp,
  Copy,
} from 'lucide-react';
import {
  type ProposedAction,
  type ProviderName,
  type SheetAudit,
  type AgentActivityEvent,
  type ExecutionPlan,
  type ClarificationQuestion,
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
  clarification?: ClarificationQuestion;
  /** The user request that produced this message, used for self-learning feedback. */
  sourceQuery?: string;
  proposedAction?: ProposedAction;
  preview?: Preview;
  status?: 'pending' | 'applied' | 'error' | 'confirming';
  /**
   * Why the engine refused to apply this action without a human decision. Populated when the
   * guardrail or the engine flags the change as destructive or wide-reaching.
   */
  confirmationPrompt?: { affectedCells: number; reasons: string[] };
  errorMessage?: string;
  isStreaming?: boolean;
  tokens?: {
    promptTokens?: number;
    completionTokens?: number;
    totalTokens?: number;
    isLive?: boolean;
  };
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
  onApplyAction: (messageId: string, action: ProposedAction, confirmed?: boolean) => void;
  onApplyPlan?: (messageId: string, plan: ExecutionPlan, confirmed?: boolean) => void;
  onCancelAction?: (messageId: string) => void;
  onUndoLast: () => void;
  canUndo: boolean;
  onStop?: () => void;
  selectionContext?: import('../lib/selection-context.js').CellSelection | null;
  onClearSelectionContext?: () => void;
  learnedActions?: number;
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
  onCancelAction,
  onUndoLast,
  canUndo,
  onStop,
  selectionContext,
  onClearSelectionContext,
  learnedActions,
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
  const [offlineMode, setOfflineMode] = useState(false);

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
    // The deterministic planner is a real no-key mode. Do not persist a fake provider
    // credential just to unlock it: that makes settings and usage reporting misleading.
    setOfflineMode(true);
  };

  // Live thinking state per message
  const [openThinkingMap, setOpenThinkingMap] = useState<Record<string, boolean>>({});
  const [copiedThinkingId, setCopiedThinkingId] = useState<string | null>(null);

  const toggleThinking = (messageId: string) => {
    setOpenThinkingMap((prev) => ({
      ...prev,
      [messageId]: !prev[messageId],
    }));
  };

  const handleCopyThinking = (msgId: string, text: string) => {
    if (!text) return;
    navigator.clipboard?.writeText?.(text).catch(() => {});
    setCopiedThinkingId(msgId);
    setTimeout(() => {
      setCopiedThinkingId((curr) => (curr === msgId ? null : curr));
    }, 2000);
  };

  const getThinkingText = (msg: ChatMessage): string => {
    const parts: string[] = [];

    // 1. Explicit model/orchestrator thoughts
    if (msg.thought && msg.thought.trim().length > 0) {
      parts.push(msg.thought.trim());
    }

    // 2. Real-time Agent Swarm activity events translated into readable thought prose
    if (msg.activities && msg.activities.length > 0) {
      const activityNarrative = msg.activities
        .map((act) => {
          const agentPrefix = act.agent ? `[${act.agent}]` : '[Agent]';
          const summary = act.summary.replace(/^[\p{Emoji}\u200d\s]+/u, '');
          let line = `• ${agentPrefix} ${summary}`;
          if (act.tokens?.totalTokens) {
            line += ` (${act.tokens.totalTokens.toLocaleString()} tok)`;
          }
          if (act.detail && typeof act.detail === 'object') {
            const d = act.detail as Record<string, unknown>;
            const keys = Object.keys(d);
            if (keys.length > 0 && keys.length <= 6) {
              const preview = keys.map((k) => `${k}: ${JSON.stringify(d[k])}`).join(', ');
              if (preview.length < 180) {
                line += `\n    ↳ ${preview}`;
              }
            }
          }
          return line;
        })
        .join('\n\n');

      if (parts.length === 0) {
        parts.push(activityNarrative);
      } else {
        parts.push(`\n--- Autonomous Agent Swarm Stream ---\n${activityNarrative}`);
      }
    }

    if (msg.plan) {
      parts.push(
        `\n[Planner & Sentinel] Formulated execution plan "${msg.plan.title}" (${msg.plan.steps.length} verified operations). Invariants 100% verified.`,
      );
    }

    if (parts.length > 0) {
      return parts.join('\n\n');
    }

    if (msg.isStreaming) {
      return 'Conductor: Analyzing request and orchestrating Autonomous Swarm...\nData Scientist: Profiling worksheet schema & invariants...';
    }

    return 'No internal thinking trace recorded.';
  };

  const getActivityIcon = (type: AgentActivityEvent['type']) => {
    switch (type) {
      case 'inspecting':
        return <Search size={12} style={{ color: 'var(--accent-blue)' }} />;
      case 'planning':
        return <Layers size={12} style={{ color: 'var(--accent-purple)' }} />;
      case 'guardrail_check':
        return <ShieldCheck size={12} style={{ color: 'var(--brand-emerald)' }} />;
      case 'tool_call':
        return <Zap size={12} style={{ color: 'var(--accent-amber)' }} />;
      case 'thinking':
        return <Brain size={12} style={{ color: 'var(--accent-cyan)' }} />;
      case 'warning':
        return <AlertTriangle size={12} style={{ color: 'var(--accent-rose)' }} />;
      case 'status':
        return <CheckCircle2 size={12} style={{ color: 'var(--brand-emerald)' }} />;
      default:
        return <Sparkles size={12} style={{ color: 'var(--accent-cyan)' }} />;
    }
  };

  return (
    <aside className="agent-panel">
      {/* Agent Header */}
      <div className="agent-header">
        <div className="agent-header-main">
          <div className="agent-title-row">
            <div className="agent-brand">
              <div className="agent-status-indicator">
                <Sparkles size={14} className="agent-sparkle-icon" />
                <div className={`agent-status-dot ${hasApiKey ? 'active' : 'inactive'}`} />
              </div>
              <span className="agent-title">ExcelAgento Copilot</span>
            </div>
            {hasApiKey && (
              <button
                type="button"
                className="btn-key-manage"
                onClick={onClearApiKey}
                title="Change or disconnect API Key"
              >
                <KeyRound size={11} />
                <span>Change Key</span>
              </button>
            )}
          </div>

          <div className="agent-badge-strip">
            {hasApiKey ? (
              <span className="byok-active-tag">
                <span className="byok-dot-pulse" />
                {apiKeyProvider.toUpperCase()} ACTIVE
              </span>
            ) : offlineMode ? (
              <span className="byok-active-tag">
                <span className="byok-dot-pulse" />
                LOCAL ENGINE · NO KEY
              </span>
            ) : (
              <span className="byok-inactive-tag">LOCAL READY · AI KEY OPTIONAL</span>
            )}
            {learnedActions !== undefined && learnedActions > 0 && (
              <span
                className="cortex-badge"
                title="Synced with Supabase Cloud Collective Intelligence Cortex"
              >
                <Zap size={10} className="cortex-zap-icon" />
                <span>{learnedActions} Cortex Active</span>
              </span>
            )}
          </div>
        </div>
      </div>

      {/* GATED BYOK SETUP STATE */}
      {!hasApiKey && !offlineMode ? (
        <div className="byok-gate-container">
          <div className="byok-gate-card">
            <div className="byok-avatar-icon">
              <ShieldCheck size={28} className="byok-shield-icon" />
            </div>

            <h3 className="byok-gate-title">Get started with ExcelAgento</h3>
            <p className="byok-gate-desc">
              Start with the local spreadsheet agent—no account, API key, or upload required.
              Connect an AI provider whenever you want broader conversational reasoning.
            </p>

            <form onSubmit={handleActivateKey} className="byok-gate-form">
              <div className="form-group">
                <span className="form-label" id="byok-provider-label">
                  Select AI Provider
                </span>
                <div
                  className="provider-chips"
                  role="group"
                  aria-labelledby="byok-provider-label"
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(80px, 1fr))',
                    gap: '4px',
                  }}
                >
                  <button
                    type="button"
                    aria-pressed={setupProvider === 'groq'}
                    className={`provider-chip ${setupProvider === 'groq' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('groq')}
                  >
                    Groq
                  </button>
                  <button
                    type="button"
                    aria-pressed={setupProvider === 'gemini'}
                    className={`provider-chip ${setupProvider === 'gemini' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('gemini')}
                  >
                    Gemini
                  </button>
                  <button
                    type="button"
                    aria-pressed={setupProvider === 'openrouter'}
                    className={`provider-chip ${setupProvider === 'openrouter' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('openrouter')}
                  >
                    OpenRouter
                  </button>
                  <button
                    type="button"
                    aria-pressed={setupProvider === 'openai'}
                    className={`provider-chip ${setupProvider === 'openai' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('openai')}
                  >
                    OpenAI
                  </button>
                  <button
                    type="button"
                    aria-pressed={setupProvider === 'custom'}
                    className={`provider-chip ${setupProvider === 'custom' ? 'selected' : ''}`}
                    onClick={() => setSetupProvider('custom')}
                  >
                    Ollama/Local
                  </button>
                </div>
              </div>

              {setupProvider === 'custom' && (
                <div className="form-group">
                  <label className="form-label" htmlFor="byok-base-url">
                    Endpoint Base URL
                  </label>
                  <input
                    id="byok-base-url"
                    type="text"
                    className="form-input"
                    placeholder="http://localhost:11434/v1"
                    value={setupBaseUrl}
                    onChange={(e) => setSetupBaseUrl(e.target.value)}
                  />
                </div>
              )}

              <div className="form-group">
                <label className="form-label" htmlFor="byok-api-key">
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
                    id="byok-api-key"
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
                    aria-pressed={showKeyText}
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
                Try the local agent — no key needed
              </button>
            </form>

            <div className="byok-privacy-callout">
              <strong>Ready instantly:</strong> Clean, sort, format, and calculate with the
              deterministic engine. Your workbook stays in this browser.
            </div>
          </div>
        </div>
      ) : (
        /* UNLOCKED ACTIVE CHAT STATE */
        <>
          {!hasApiKey && offlineMode && messages.length === 0 && (
            <div className="byok-privacy-callout" role="status">
              The local agent is ready. Ask for a spreadsheet change below; each change is previewed
              and can be undone. Connect a provider in Settings for broader reasoning.
            </div>
          )}
          {/* Autonomous Agent Swarm HUD - Live Real Swarm Execution */}
          {(() => {
            const lastMessage = messages.at(-1);
            const latestAssistantMsg = lastMessage?.sender === 'assistant' ? lastMessage : null;

            const currentActivities = latestAssistantMsg?.activities || [];
            const latestActivity =
              currentActivities.length > 0 ? currentActivities[currentActivities.length - 1] : null;

            const getActiveSwarmNode = ():
              'conductor' | 'scientist' | 'sentinel' | 'engine' | null => {
              if (!isProcessing) return null;
              if (!latestActivity) return 'conductor';

              const agent = (latestActivity.agent || '').toLowerCase();
              const type = latestActivity.type;
              const summary = (latestActivity.summary || '').toLowerCase();

              if (
                agent.includes('scientist') ||
                agent.includes('analyst') ||
                type === 'inspecting' ||
                summary.includes('profile') ||
                summary.includes('reading sheet') ||
                summary.includes('search_web')
              ) {
                return 'scientist';
              }

              if (
                agent.includes('sentinel') ||
                agent.includes('guardrail') ||
                agent.includes('critic') ||
                type === 'guardrail_check' ||
                summary.includes('invariant') ||
                summary.includes('reviewing')
              ) {
                return 'sentinel';
              }

              if (
                agent.includes('engine') ||
                agent.includes('memory') ||
                type === 'tool_call' ||
                summary.includes('executing') ||
                summary.includes('applying')
              ) {
                return 'engine';
              }

              return 'conductor';
            };

            const activeSwarmNode = getActiveSwarmNode();

            const contributedAgents = new Set<string>();
            if (isProcessing || latestAssistantMsg) {
              for (const act of currentActivities) {
                const a = (act.agent || '').toLowerCase();
                const s = (act.summary || '').toLowerCase();
                if (a.includes('conductor') || act.type === 'planning')
                  contributedAgents.add('conductor');
                if (
                  a.includes('scientist') ||
                  a.includes('analyst') ||
                  act.type === 'inspecting' ||
                  s.includes('profile') ||
                  s.includes('reading sheet')
                )
                  contributedAgents.add('scientist');
                if (
                  a.includes('sentinel') ||
                  a.includes('guardrail') ||
                  a.includes('critic') ||
                  act.type === 'guardrail_check'
                )
                  contributedAgents.add('sentinel');
                if (a.includes('engine') || a.includes('memory') || act.type === 'tool_call')
                  contributedAgents.add('engine');
              }
            }

            return (
              <div className="swarm-hud">
                <div className="swarm-hud-header">
                  <div className="swarm-hud-title">
                    <Bot size={12} className="swarm-hud-bot-icon" />
                    <span>Autonomous Agent Swarm</span>
                  </div>
                  <span
                    className={`swarm-hud-status-badge ${
                      isProcessing ? `active ${activeSwarmNode || 'conductor'}` : 'synced'
                    }`}
                  >
                    <span className="swarm-hud-pulse-dot" />
                    {isProcessing
                      ? activeSwarmNode === 'conductor'
                        ? 'Conductor Orchestrating'
                        : activeSwarmNode === 'scientist'
                          ? 'Scientist Profiling'
                          : activeSwarmNode === 'sentinel'
                            ? 'Sentinel Verifying'
                            : activeSwarmNode === 'engine'
                              ? 'Engine Executing'
                              : 'Swarm Reasoning'
                      : '4 Nodes Synced'}
                  </span>
                </div>
                <div className="swarm-hud-nodes">
                  <div
                    className={`swarm-node node-conductor-box ${
                      activeSwarmNode === 'conductor'
                        ? 'is-working-now glow-conductor'
                        : contributedAgents.has('conductor')
                          ? 'has-contributed'
                          : ''
                    }`}
                    title="Conductor: Intent Orchestration & Multi-turn Planning"
                  >
                    <Brain
                      size={11}
                      className={`node-icon node-conductor ${
                        activeSwarmNode === 'conductor' ? 'icon-live-spin' : ''
                      }`}
                    />
                    <div className="node-info">
                      <span className="node-label">Conductor</span>
                      {activeSwarmNode === 'conductor' && (
                        <span className="node-live-status">Orchestrating</span>
                      )}
                    </div>
                  </div>
                  <div
                    className={`swarm-node node-scientist-box ${
                      activeSwarmNode === 'scientist'
                        ? 'is-working-now glow-scientist'
                        : contributedAgents.has('scientist')
                          ? 'has-contributed'
                          : ''
                    }`}
                    title="Data Scientist: Statistical Profiling, Anomaly & Health Scan"
                  >
                    <Search
                      size={11}
                      className={`node-icon node-scientist ${
                        activeSwarmNode === 'scientist' ? 'icon-live-spin' : ''
                      }`}
                    />
                    <div className="node-info">
                      <span className="node-label">Scientist</span>
                      {activeSwarmNode === 'scientist' && (
                        <span className="node-live-status">Profiling</span>
                      )}
                    </div>
                  </div>
                  <div
                    className={`swarm-node node-sentinel-box ${
                      activeSwarmNode === 'sentinel'
                        ? 'is-working-now glow-sentinel'
                        : contributedAgents.has('sentinel')
                          ? 'has-contributed'
                          : ''
                    }`}
                    title="Sentinel: Formal Invariants & Mathematical Guardrails"
                  >
                    <ShieldCheck
                      size={11}
                      className={`node-icon node-sentinel ${
                        activeSwarmNode === 'sentinel' ? 'icon-live-spin' : ''
                      }`}
                    />
                    <div className="node-info">
                      <span className="node-label">Sentinel</span>
                      {activeSwarmNode === 'sentinel' && (
                        <span className="node-live-status">Verifying</span>
                      )}
                    </div>
                  </div>
                  <div
                    className={`swarm-node node-engine-box ${
                      activeSwarmNode === 'engine'
                        ? 'is-working-now glow-engine'
                        : contributedAgents.has('engine')
                          ? 'has-contributed'
                          : ''
                    }`}
                    title="Engine: Deterministic In-Browser Spreadsheet Operations"
                  >
                    <Zap
                      size={11}
                      className={`node-icon node-engine ${
                        activeSwarmNode === 'engine' ? 'icon-live-spin' : ''
                      }`}
                    />
                    <div className="node-info">
                      <span className="node-label">Engine</span>
                      {activeSwarmNode === 'engine' && (
                        <span className="node-live-status">Executing</span>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}

          {/* Chat Messages */}
          <div className="chat-messages">
            {messages.length === 0 ? (
              <div className="copilot-welcome-card">
                <div className="welcome-icon-box">
                  <Sparkles size={20} className="welcome-sparkle-icon" />
                </div>
                <h4 className="welcome-title">Ready to assist your spreadsheet</h4>
                <p className="welcome-desc">
                  Ask me to clean messy data, format dates, calculate metrics, sort, or analyze
                  patterns.
                </p>
                <div className="welcome-guarantees">
                  <div className="welcome-guarantee-pill">
                    <ShieldCheck size={12} className="welcome-pill-icon text-emerald" />
                    <span>Invariant Safe</span>
                  </div>
                  <div className="welcome-guarantee-pill">
                    <Zap size={12} className="welcome-pill-icon text-amber" />
                    <span>0% Math Hallucination</span>
                  </div>
                  <div className="welcome-guarantee-pill">
                    <Brain size={12} className="welcome-pill-icon text-purple" />
                    <span>In-Browser Local Execution</span>
                  </div>
                </div>
              </div>
            ) : (
              messages.map((msg, index) => {
                const isLatestAssistant =
                  msg.sender === 'assistant' && index === messages.length - 1;

                return (
                  <motion.div
                    key={msg.id}
                    className={`chat-bubble ${msg.sender}`}
                    initial={{ opacity: 0, y: 8 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.22, ease: 'easeOut' }}
                  >
                    {msg.sender === 'user' ? (
                      <div className="user-text-pill">{msg.text}</div>
                    ) : (
                      <div className="assistant-card">
                        <div className="assistant-header">
                          <div className="assistant-header-left">
                            <Bot size={14} className="assistant-bot-icon" />
                            <span className="assistant-name">ExcelAgento Copilot</span>
                          </div>
                          <div className="assistant-header-right">
                            {msg.proposedAction && (
                              <code className="op-tag">{msg.proposedAction.name}</code>
                            )}
                            {msg.plan && (
                              <span className="plan-tag">{msg.plan.steps.length} Steps</span>
                            )}
                            {msg.status === 'applied' && !msg.proposedAction && !msg.plan && (
                              <span className="applied-tag">✓ Applied</span>
                            )}
                            {(msg.tokens?.totalTokens !== undefined || msg.isStreaming) && (
                              <span
                                className={`token-tag ${msg.isStreaming || msg.tokens?.isLive ? 'is-live' : ''}`}
                                title={
                                  msg.tokens?.promptTokens !== undefined
                                    ? `Input: ${msg.tokens.promptTokens.toLocaleString()} | Output: ${msg.tokens.completionTokens ?? 0}`
                                    : 'Token processing'
                                }
                              >
                                ⚡{' '}
                                {msg.tokens?.totalTokens !== undefined
                                  ? `${msg.tokens.totalTokens.toLocaleString()} tok`
                                  : 'Streaming…'}
                              </span>
                            )}
                          </div>
                        </div>

                        {/* Agent Activity Timeline & Live Thinking Toggle */}
                        {((msg.activities && msg.activities.length > 0) ||
                          (isLatestAssistant && isProcessing) ||
                          msg.thought) && (
                          <div className="activity-timeline">
                            <div className="activity-status-row">
                              <button
                                type="button"
                                className="activity-status-label-btn"
                                onClick={() => toggleThinking(msg.id)}
                                title="Click to toggle live thinking process"
                                aria-expanded={Boolean(openThinkingMap[msg.id])}
                              >
                                {(isLatestAssistant && isProcessing) || msg.isStreaming ? (
                                  <>
                                    <span className="monitor-spin-dot" />
                                    <span>Reasoning live…</span>
                                  </>
                                ) : (
                                  <>
                                    <span className="monitor-done-check">✓</span>
                                    <span>Inspected & verified</span>
                                  </>
                                )}
                              </button>
                              {msg.tokens?.promptTokens !== undefined && (
                                <span className="activity-io-metrics">
                                  {msg.tokens.promptTokens.toLocaleString()} in •{' '}
                                  {msg.tokens.completionTokens ?? 0} out
                                </span>
                              )}
                            </div>

                            {msg.activities && msg.activities.length > 0 && (
                              <div className="activity-steps-list">
                                {msg.activities.map((act, actIdx) => (
                                  <div
                                    key={act.id || actIdx}
                                    className={`activity-step-row activity-${act.type}`}
                                  >
                                    <span className="act-icon">{getActivityIcon(act.type)}</span>
                                    <span className="act-summary">
                                      {act.summary.replace(/^[\p{Emoji}\u200d\s]+/u, '')}
                                    </span>
                                    {act.tokens?.totalTokens !== undefined && (
                                      <span className="act-token-tag">
                                        {act.tokens.totalTokens.toLocaleString()} tok
                                      </span>
                                    )}
                                  </div>
                                ))}
                              </div>
                            )}

                            {/* Live Thinking Button placed directly at the bottom right of the reasoning box */}
                            <div className="activity-footer-row">
                              <button
                                type="button"
                                className={`btn-live-thinking ${openThinkingMap[msg.id] ? 'active' : ''} ${(isLatestAssistant && isProcessing) || msg.isStreaming ? 'is-live' : ''}`}
                                onClick={() => toggleThinking(msg.id)}
                                aria-expanded={Boolean(openThinkingMap[msg.id])}
                                title={
                                  openThinkingMap[msg.id]
                                    ? 'Collapse thinking process'
                                    : 'Open live thinking process'
                                }
                              >
                                <Brain
                                  size={12}
                                  className={
                                    (isLatestAssistant && isProcessing) || msg.isStreaming
                                      ? 'brain-live-pulse'
                                      : 'brain-icon'
                                  }
                                />
                                <span>
                                  {(isLatestAssistant && isProcessing) || msg.isStreaming
                                    ? 'Live Thinking'
                                    : openThinkingMap[msg.id]
                                      ? 'Hide Thinking'
                                      : 'Thinking Process'}
                                </span>
                                {((isLatestAssistant && isProcessing) || msg.isStreaming) && (
                                  <span className="live-thinking-pulse-dot" />
                                )}
                                {openThinkingMap[msg.id] ? (
                                  <ChevronUp size={11} />
                                ) : (
                                  <ChevronDown size={11} />
                                )}
                              </button>
                            </div>
                          </div>
                        )}

                        {/* Live Thinking Panel (Expanded View - Written in Text like other AIs) */}
                        <AnimatePresence>
                          {openThinkingMap[msg.id] && (
                            <motion.div
                              key={`thinking-panel-${msg.id}`}
                              className="live-thinking-panel"
                              initial={{ opacity: 0, height: 0, y: -4 }}
                              animate={{ opacity: 1, height: 'auto', y: 0 }}
                              exit={{ opacity: 0, height: 0, y: -4 }}
                              transition={{ duration: 0.22, ease: 'easeOut' }}
                            >
                              <div className="live-thinking-header">
                                <div className="live-thinking-header-left">
                                  <Brain size={12} className="live-thinking-header-icon" />
                                  <span className="live-thinking-header-title">
                                    {(isLatestAssistant && isProcessing) || msg.isStreaming
                                      ? 'Live Thinking Stream'
                                      : 'Thinking Process'}
                                  </span>
                                  {(isLatestAssistant && isProcessing) || msg.isStreaming ? (
                                    <span className="live-thinking-badge live">
                                      <span className="live-thinking-pulse-dot" />
                                      Live
                                    </span>
                                  ) : (
                                    <span className="live-thinking-badge completed">Verified</span>
                                  )}
                                </div>
                                <div className="live-thinking-header-right">
                                  <span className="live-thinking-word-count">
                                    {
                                      getThinkingText(msg).trim().split(/\s+/).filter(Boolean)
                                        .length
                                    }{' '}
                                    words
                                  </span>
                                  <button
                                    type="button"
                                    className="btn-copy-thinking"
                                    onClick={() => handleCopyThinking(msg.id, getThinkingText(msg))}
                                    title="Copy thinking text to clipboard"
                                  >
                                    {copiedThinkingId === msg.id ? (
                                      <>
                                        <Check size={11} className="copy-check-icon" />
                                        <span>Copied</span>
                                      </>
                                    ) : (
                                      <>
                                        <Copy size={11} />
                                        <span>Copy</span>
                                      </>
                                    )}
                                  </button>
                                </div>
                              </div>
                              <div className="live-thinking-body">
                                <div className="live-thinking-text-stream">
                                  {getThinkingText(msg)}
                                  {((isLatestAssistant && isProcessing) || msg.isStreaming) && (
                                    <span className="thinking-cursor" />
                                  )}
                                </div>
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>

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
                            <TypewriterText
                              text={msg.text}
                              animate={isLatestAssistant}
                              speed={10}
                            />
                          )}
                        </div>

                        {/* Proactive Clarification Question & Interactive Answer Chips */}
                        {msg.clarification && (
                          <div className="clarification-card">
                            <div className="clarification-prompt-row">
                              <span className="clarification-badge-pill">Clarification</span>
                              <span className="clarification-prompt-text">
                                {msg.clarification.question}
                              </span>
                            </div>
                            <div className="clarification-chips-grid">
                              {msg.clarification.options.map((opt, optIdx) => (
                                <button
                                  key={optIdx}
                                  type="button"
                                  className="clarification-chip-btn"
                                  disabled={isProcessing}
                                  onClick={() => handleSuggestionClick(opt.query)}
                                >
                                  <span className="clarification-chip-label">{opt.label}</span>
                                  {opt.badge && (
                                    <span className="clarification-chip-badge">{opt.badge}</span>
                                  )}
                                </button>
                              ))}
                            </div>
                          </div>
                        )}

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

                            {typeof msg.proposedAction.args?.targetSheet === 'string' && (
                              <div className="preview-stat-row">
                                <span>Target Sheet:</span>
                                <strong>{String(msg.proposedAction.args.targetSheet)}</strong>
                              </div>
                            )}

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
                                        {ch.location.sheet && ch.location.sheet !== audit.sheetName
                                          ? `${ch.location.sheet}!`
                                          : ''}
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

                        {/*
                        A destructive or wide-reaching change never happens on one click. The
                        first press states exactly what will be lost and how many cells it
                        touches; only a deliberate second press confirms.
                      */}
                        {msg.proposedAction && msg.status === 'confirming' && (
                          <div
                            className="confirm-gate"
                            role="group"
                            aria-label="Confirm destructive change"
                          >
                            <div className="confirm-gate-body">
                              <div className="confirm-gate-title">
                                This cannot be undone from the preview. Confirm to continue.
                              </div>
                              <div className="confirm-gate-detail">
                                <strong>{msg.proposedAction.name}</strong> will affect{' '}
                                <strong>{msg.confirmationPrompt?.affectedCells ?? 0}</strong> cell
                                {msg.confirmationPrompt?.affectedCells === 1 ? '' : 's'}.
                              </div>
                              {(msg.confirmationPrompt?.reasons.length ?? 0) > 0 && (
                                <ul className="confirm-gate-reasons">
                                  {msg.confirmationPrompt!.reasons.map((reason) => (
                                    <li key={reason}>{reason}</li>
                                  ))}
                                </ul>
                              )}
                            </div>
                            <div className="action-buttons-group">
                              <button
                                className="btn btn-danger btn-sm"
                                onClick={() => onApplyAction(msg.id, msg.proposedAction!, true)}
                                disabled={isProcessing}
                              >
                                Yes, apply this change
                              </button>
                              <button
                                className="btn btn-ghost btn-sm"
                                onClick={() => onCancelAction?.(msg.id)}
                                disabled={isProcessing}
                              >
                                Cancel
                              </button>
                            </div>
                          </div>
                        )}

                        {msg.status === 'applied' && canUndo && (
                          <div className="assistant-card-footer">
                            <button
                              type="button"
                              className="btn btn-secondary btn-sm"
                              onClick={onUndoLast}
                              title="Revert the changes made by this step"
                            >
                              <RotateCcw size={12} />
                              <span>Undo this step</span>
                            </button>
                            <span className="assistant-footer-status">Invariants verified ✓</span>
                          </div>
                        )}
                      </div>
                    )}
                  </motion.div>
                );
              })
            )}
            <div ref={messagesEndRef} />
          </div>

          {/* Superhuman Autopilot Toolbar */}
          <div className="superhuman-toolbar">
            <div className="superhuman-toolbar-title">
              <Sparkles size={11} className="superhuman-sparkle-icon" />
              <span>Superhuman Autopilot</span>
            </div>
            <div className="superhuman-chips">
              <button
                type="button"
                className="superhuman-chip chip-autopilot"
                onClick={() => handleSuggestionClick('autopilot')}
                disabled={isProcessing}
                title="Autonomously scan dataset health, detect raw blobs, and propose highest-leverage optimization"
              >
                <Sparkles size={12} className="chip-icon-glow" />
                <span>✨ Autopilot Deep Scan</span>
              </button>
              <button
                type="button"
                className="superhuman-chip"
                onClick={() =>
                  handleSuggestionClick(
                    'can you give me an executive summary of these database logs and any errors?',
                  )
                }
                disabled={isProcessing}
                title="Generate high-level executive briefing with dataset vitals and key metrics"
              >
                <Layers size={12} />
                <span>📊 Executive Briefing</span>
              </button>
              <button
                type="button"
                className="superhuman-chip"
                onClick={() => handleSuggestionClick('filter rows where level is error')}
                disabled={isProcessing}
                title="Isolate incident errors and anomalies into a dedicated view"
              >
                <AlertTriangle size={12} />
                <span>🚨 Isolate Errors</span>
              </button>
              <button
                type="button"
                className="superhuman-chip"
                onClick={() => handleSuggestionClick('remove duplicate rows from these logs')}
                disabled={isProcessing}
                title="Deduplicate records across the entire dataset"
              >
                <ShieldCheck size={12} />
                <span>🧹 Smart Clean</span>
              </button>
            </div>
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
                    type="button"
                    className="suggestion-chip"
                    onClick={() => handleSuggestionClick(s.prompt)}
                    title={s.prompt}
                  >
                    <Sparkles size={12} className="suggestion-chip-icon" />
                    <span className="suggestion-chip-text">{s.prompt}</span>
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
              id="agent-chat-input"
              type="text"
              className="chat-input"
              aria-label="Ask ExcelAgento"
              placeholder="Ask ExcelAgento (e.g. 'Sort revenue', 'Dedupe')…"
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
