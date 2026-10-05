import React, { useRef } from 'react';
import type { HistoryEntry } from '@excel-agent/engine';

import { useDialogA11y } from '../lib/use-dialog-a11y.js';

interface HistoryDrawerProps {
  isOpen: boolean;
  onClose: () => void;
  entries: HistoryEntry[];
  currentPosition: number;
  onStepBack: (position: number) => void;
}

export const HistoryDrawer: React.FC<HistoryDrawerProps> = ({
  isOpen,
  onClose,
  entries,
  currentPosition,
  onStepBack,
}) => {
  const drawerRef = useRef<HTMLElement>(null);
  // The drawer is not a modal - the grid stays usable behind it - but Escape still closes it and
  // focus still returns to the control that opened it.
  useDialogA11y(isOpen, drawerRef, onClose);

  if (!isOpen) return null;

  return (
    <aside
      className="history-drawer"
      ref={drawerRef}
      role="dialog"
      aria-modal="false"
      aria-label="Operation audit log"
      data-dialog-open="true"
    >
      <div className="history-header">
        <div className="history-title-group">
          <svg
            className="history-icon"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <circle cx="12" cy="12" r="10" />
            <polyline points="12 6 12 12 14 14" />
          </svg>
          <div>
            <div className="history-title">Operation Audit Log</div>
            <div className="history-subtitle">
              {currentPosition} of {entries.length} steps applied
            </div>
          </div>
        </div>

        <button
          type="button"
          className="btn btn-ghost btn-sm"
          aria-label="Close history"
          onClick={onClose}
        >
          ✕
        </button>
      </div>

      <div className="history-list">
        {/* Initial Baseline Item */}
        <div className={`history-item ${currentPosition === 0 ? 'active' : ''}`}>
          <div className="history-item-row">
            <span className="history-item-title">0. Initial State</span>
            {currentPosition === 0 && <span className="history-current-tag">● Current</span>}
          </div>
          <p className="history-item-desc">Original loaded workbook</p>
          {currentPosition !== 0 && (
            <button
              type="button"
              className="btn btn-secondary btn-sm history-step-btn"
              onClick={() => onStepBack(0)}
            >
              Restore to Origin
            </button>
          )}
        </div>

        {entries.map((entry, idx) => {
          const stepNum = idx + 1;
          const isCurrent = currentPosition === stepNum;

          return (
            <div key={idx} className={`history-item ${isCurrent ? 'active' : ''}`}>
              <div className="history-item-row">
                <span className="op-badge">{entry.operationName}</span>
                {isCurrent && <span className="history-current-tag">● Current</span>}
              </div>

              <div className="history-item-meta">
                Step {stepNum} • {entry.patch.length} patch entries
              </div>

              {!isCurrent && (
                <button
                  type="button"
                  className="btn btn-secondary btn-sm history-step-btn"
                  onClick={() => onStepBack(stepNum)}
                >
                  Step to this State
                </button>
              )}
            </div>
          );
        })}
      </div>
    </aside>
  );
};
