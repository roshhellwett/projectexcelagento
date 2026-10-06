import React, { useEffect, useState } from 'react';
import { Brain, Sparkles, Square, Zap } from 'lucide-react';
import type { AgentActivityEvent } from '@excel-agent/agent';

interface LiveThinkingProgressBarProps {
  isProcessing: boolean;
  onStop?: () => void;
  latestActivity?: AgentActivityEvent;
  tokens?: { totalTokens?: number; completionTokens?: number; promptTokens?: number };
  thoughtText?: string;
}

export const LiveThinkingProgressBar: React.FC<LiveThinkingProgressBarProps> = ({
  isProcessing,
  onStop,
  latestActivity,
  tokens,
  thoughtText,
}) => {
  const [elapsedSeconds, setElapsedSeconds] = useState(0);

  useEffect(() => {
    if (!isProcessing) {
      setElapsedSeconds(0);
      return;
    }
    const timer = setInterval(() => {
      setElapsedSeconds((s) => s + 1);
    }, 1000);
    return () => clearInterval(timer);
  }, [isProcessing]);

  if (!isProcessing) return null;

  // Asymptotic progress calculation from 10% to 96%
  const progressPercent = Math.min(
    96,
    Math.round(10 + Math.atan(elapsedSeconds / 10) * (85 / (Math.PI / 2))),
  );

  const formatTimer = (totalSeconds: number) => {
    const mins = Math.floor(totalSeconds / 60);
    const secs = totalSeconds % 60;
    return `${mins}:${secs.toString().padStart(2, '0')}`;
  };

  // Determine active agent phase
  let phaseTitle = '🧠 Conductor: Grounding Context';
  let phaseDetail = 'Grounding context with workbook schema & rules...';

  if (latestActivity) {
    const agentLabel = latestActivity.agent ? `[${latestActivity.agent}]` : 'Agent';
    const icon =
      latestActivity.type === 'inspecting'
        ? '🔬 Inspecting Data'
        : latestActivity.type === 'planning'
          ? '🛡️ Synthesizing Plan'
          : latestActivity.type === 'guardrail_check'
            ? '⚖️ Verifying Invariants'
            : '⚡ Processing Operations';
    phaseTitle = `${agentLabel} ${icon}`;
    phaseDetail = latestActivity.summary;
  } else if (elapsedSeconds < 4) {
    phaseTitle = '🧠 Conductor: Grounding Context';
    phaseDetail = 'Grounding context with workbook schema & rules...';
  } else if (elapsedSeconds < 14) {
    phaseTitle = '🔬 Data Scientist: Deep Inspection';
    phaseDetail = 'Reading records, testing invariants & cross-checking data...';
  } else if (elapsedSeconds < 30) {
    phaseTitle = '🛡️ Sentinel: Execution Strategy';
    phaseDetail = 'Formulating verified operations & mathematical checks...';
  } else {
    phaseTitle = '⚡ Conductor: Synthesizing Answer';
    phaseDetail = 'Compiling deep analytical reasoning & deliverables...';
  }

  const wordCount = thoughtText ? thoughtText.trim().split(/\s+/).filter(Boolean).length : 0;

  return (
    <div className="live-thinking-progress-card" role="status" aria-live="polite">
      {/* Header Row */}
      <div className="live-thinking-progress-top">
        <div className="live-thinking-progress-left">
          <span className="live-thinking-pulse-ring">
            <span className="live-thinking-pulse-inner" />
          </span>
          <Brain size={13} className="live-thinking-icon" />
          <span className="live-thinking-phase-title">{phaseTitle}</span>
          <span className="live-thinking-timer-badge">{formatTimer(elapsedSeconds)}</span>
        </div>

        <div className="live-thinking-progress-right">
          {tokens?.totalTokens ? (
            <span className="live-thinking-token-pill">
              <Zap size={10} />
              {tokens.totalTokens.toLocaleString()} tok
            </span>
          ) : wordCount > 0 ? (
            <span className="live-thinking-token-pill">
              <Sparkles size={10} />
              {wordCount} words
            </span>
          ) : null}

          {onStop && (
            <button
              type="button"
              className="btn-stop-thinking-sm"
              onClick={onStop}
              title="Stop agent thinking and cancel request"
              aria-label="Cancel thinking"
            >
              <Square size={9} fill="currentColor" />
              <span>Cancel</span>
            </button>
          )}
        </div>
      </div>

      {/* Progress Bar Track & Shimmer Fill */}
      <div className="live-thinking-progress-track">
        <div className="live-thinking-progress-fill" style={{ width: `${progressPercent}%` }} />
      </div>

      {/* Sub-detail description */}
      <div className="live-thinking-progress-detail">
        <span className="live-thinking-detail-text">{phaseDetail}</span>
        <span className="live-thinking-pct">{progressPercent}%</span>
      </div>
    </div>
  );
};
