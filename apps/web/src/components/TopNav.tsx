import React, { useRef, useState } from 'react';
import { getActiveTheme, toggleTheme } from '../lib/theme.js';

/**
 * Persistent light/dark switch. The state is owned by the document element (see theme.ts), so
 * this button does not need to be wired into the workspace's state tree.
 */
export function ThemeToggle() {
  const [theme, setTheme] = useState<'light' | 'dark'>(getActiveTheme());
  return (
    <button
      className="btn btn-ghost btn-sm"
      onClick={() => setTheme(toggleTheme())}
      title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
      aria-pressed={theme === 'dark'}
      aria-label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
      data-testid="theme-toggle"
    >
      {theme === 'dark' ? '☀' : '☾'}
    </button>
  );
}

interface TopNavProps {
  fileName: string;
  activeSheetName: string;
  rowCount: number;
  colCount: number;
  canUndo: boolean;
  canRedo: boolean;
  historyLength: number;
  historyPosition: number;
  searchQuery: string;
  searchMatchCount: number;
  onSearchChange: (q: string) => void;
  onUndo: () => void;
  onRedo: () => void;
  onReset: () => void;
  onFileUpload: (file: File) => void;
  onExport: () => void;
  onSelectFixture: (fixtureName: string) => void;
  onOpenOperationModal: () => void;
  onToggleHistory: () => void;
  onOpenSettings: () => void;
  onOpenUsage?: () => void;
  onOpenCommandPalette?: () => void;
}

const FIXTURES = [
  { label: 'Sample: Messy Orders (Built-in)', value: 'sample' },
  { label: '01: Mixed Date Formats', value: '01-mixed-date-formats.xlsx' },
  { label: '02: Merged Cells', value: '02-merged-cells.xlsx' },
  { label: '03: Blank Rows', value: '03-blank-rows.xlsx' },
  { label: '04: Numbers as Text', value: '04-numbers-stored-as-text.xlsx' },
  { label: '05: Multiple Sheets', value: '05-multiple-sheets.xlsx' },
  { label: '06: Chart File', value: '06-chart.xlsx' },
  { label: '07: Conditional Formatting', value: '07-conditional-formatting.xlsx' },
  { label: '08: Leap Days', value: '08-leap-days.xlsx' },
  { label: '09: Formulas & Empty Cells', value: '09-formulas-and-empty-cells.xlsx' },
  { label: '10: Combined Messy Orders', value: '10-combined-messy-orders.xlsx' },
];

export const TopNav: React.FC<TopNavProps> = ({
  fileName,
  activeSheetName,
  rowCount,
  colCount,
  canUndo,
  canRedo,
  historyLength,
  historyPosition,
  searchQuery,
  searchMatchCount,
  onSearchChange,
  onUndo,
  onRedo,
  onReset,
  onFileUpload,
  onExport,
  onSelectFixture,
  onOpenOperationModal,
  onToggleHistory,
  onOpenSettings,
  onOpenUsage,
  onOpenCommandPalette,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileUpload(file);
      e.target.value = '';
    }
  };

  return (
    <header className="top-nav">
      <div className="top-nav-left">
        <div className="logo-badge">
          <div className="logo-icon">X</div>
          <h1
            className="logo-title"
            style={{ fontSize: 'inherit', fontWeight: 'inherit', margin: 0 }}
          >
            Excel Agent
          </h1>
        </div>

        <div className="file-meta-pill">
          <strong>{fileName}</strong>
          <span>•</span>
          <span>{activeSheetName}</span>
          <span>•</span>
          <span>
            {rowCount} rows × {colCount} cols
          </span>
        </div>

        <label className="sr-only" htmlFor="top-nav-fixture">
          Load test fixtures
        </label>
        <select
          id="top-nav-fixture"
          className="select-input"
          title="Load Test Fixtures"
          onChange={(e) => {
            onSelectFixture(e.target.value);
            // Reset so the same fixture can be re-selected.
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>
            Load fixture...
          </option>
          {FIXTURES.map((f) => (
            <option key={f.value} value={f.value}>
              {f.label}
            </option>
          ))}
        </select>
      </div>

      <div className="top-nav-center">
        {/* Quick In-Sheet Search */}
        <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
          <label className="sr-only" htmlFor="top-nav-search">
            Find in sheet
          </label>
          <input
            id="top-nav-search"
            type="search"
            className="form-input"
            style={{
              height: '30px',
              padding: '4px 28px 4px 28px',
              fontSize: '12px',
              width: '160px',
              background: 'var(--bg-elevated)',
            }}
            placeholder="Find in sheet…"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
          />
          <svg
            style={{
              position: 'absolute',
              left: '8px',
              color: 'var(--text-dim)',
              pointerEvents: 'none',
            }}
            width="13"
            height="13"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <circle cx="11" cy="11" r="8" />
            <line x1="21" y1="21" x2="16.65" y2="16.65" />
          </svg>
          {searchQuery && (
            <span
              style={{
                position: 'absolute',
                right: '8px',
                fontSize: '10px',
                color: searchMatchCount > 0 ? 'var(--accent-amber)' : 'var(--text-dim)',
                fontWeight: 600,
              }}
            >
              {searchMatchCount}
            </span>
          )}
        </div>

        {onOpenCommandPalette && (
          <button
            className="btn btn-secondary btn-sm"
            onClick={onOpenCommandPalette}
            title="Open Command Palette (Cmd+K / Ctrl+K)"
            style={{
              padding: '3px 8px',
              height: '30px',
              fontFamily: 'var(--font-mono)',
              fontSize: '11px',
              gap: '6px',
            }}
          >
            <svg
              width="12"
              height="12"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M18 3a3 3 0 0 0-3 3v12a3 3 0 0 0 3 3 3 3 0 0 0 3-3 3 3 0 0 0-3-3H6a3 3 0 0 0-3 3 3 3 0 0 0 3 3 3 3 0 0 0 3-3V6a3 3 0 0 0-3-3 3 3 0 0 0-3 3 3 3 0 0 0 3 3h12a3 3 0 0 0 3-3 3 3 0 0 0-3-3z" />
            </svg>
            <span>HUD</span>
            <span
              style={{
                background: 'var(--bg-elevated)',
                border: '1px solid var(--border-main)',
                padding: '1px 4px',
                borderRadius: '3px',
                fontSize: '9.5px',
                color: 'var(--text-dim)',
              }}
            >
              ⌘K
            </span>
          </button>
        )}

        <label className="sr-only" htmlFor="top-nav-file">
          Upload Excel or CSV file
        </label>
        <input
          type="file"
          id="top-nav-file"
          ref={fileInputRef}
          style={{ display: 'none' }}
          // Reached through the Upload button, so it is kept out of the tab order and the tree.
          tabIndex={-1}
          aria-hidden="true"
          accept=".xlsx,.xls,.csv"
          onChange={handleFileChange}
        />
        <button
          type="button"
          className="btn btn-secondary btn-sm"
          onClick={() => fileInputRef.current?.click()}
          title="Upload Excel or CSV file"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
            <polyline points="17 8 12 3 7 8" />
            <line x1="12" y1="3" x2="12" y2="15" />
          </svg>
          Upload File
        </button>

        <button
          className="btn btn-primary btn-sm"
          onClick={onExport}
          title="Export current workbook to .xlsx"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
            <polyline points="7 10 12 15 17 10" />
            <line x1="12" y1="15" x2="12" y2="3" />
          </svg>
          Export .xlsx
        </button>

        <button
          className="btn btn-secondary btn-sm"
          onClick={onOpenOperationModal}
          title="Manually configure an engine operation"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <polygon points="12 2 2 7 12 12 22 7 12 2" />
            <polyline points="2 17 12 22 22 17" />
            <polyline points="2 12 12 17 22 12" />
          </svg>
          Run Operation
        </button>
      </div>

      <div className="top-nav-right">
        <button
          className="btn btn-ghost btn-sm"
          onClick={onUndo}
          disabled={!canUndo}
          title="Undo (Ctrl+Z)"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M3 7v6h6" />
            <path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13" />
          </svg>
          Undo
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onRedo}
          disabled={!canRedo}
          title="Redo (Ctrl+Y)"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 7v6h-6" />
            <path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3l3 2.7" />
          </svg>
          Redo
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onToggleHistory}
          title="View Operation Audit History"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <circle cx="12" cy="12" r="10" />
            <polyline points="12 6 12 12 14 14" />
          </svg>
          History ({historyPosition}/{historyLength})
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onOpenUsage}
          title="Model & token usage"
          data-testid="open-usage"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <line x1="18" y1="20" x2="18" y2="10" />
            <line x1="12" y1="20" x2="12" y2="4" />
            <line x1="6" y1="20" x2="6" y2="14" />
          </svg>
          Usage
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onReset}
          title="Reset to initial workbook"
        >
          Reset
        </button>

        <ThemeToggle />

        <button
          className="btn btn-ghost btn-sm"
          onClick={onOpenSettings}
          title="API Keys & Settings"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <circle cx="12" cy="12" r="3" />
            <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z" />
          </svg>
        </button>
      </div>
    </header>
  );
};
