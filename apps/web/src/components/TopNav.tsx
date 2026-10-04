import React, { useRef, useState } from 'react';
import {
  Sun,
  Moon,
  Search,
  Command,
  Upload,
  Download,
  PlayCircle,
  Undo2,
  Redo2,
  History,
  BarChart2,
  RotateCcw,
  Settings,
  Sparkles,
  FileSpreadsheet,
  Bot,
  BookOpen,
} from 'lucide-react';
import { getActiveTheme, toggleTheme } from '../lib/theme.js';

/**
 * Persistent light/dark switch. The state is owned by the document element (see theme.ts), so
 * this button does not need to be wired into the workspace's state tree.
 */
export function ThemeToggle() {
  const [theme, setTheme] = useState<'light' | 'dark'>(getActiveTheme());
  return (
    <button
      className="btn btn-ghost btn-sm btn-icon"
      onClick={() => setTheme(toggleTheme())}
      title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
      aria-pressed={theme === 'dark'}
      aria-label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
      data-testid="theme-toggle"
    >
      {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
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
  onOpenAgents?: () => void;
  onOpenDocs?: () => void;
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
  onOpenAgents,
  onOpenDocs,
  onOpenCommandPalette,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMac =
    typeof navigator !== 'undefined' &&
    /(Mac|iPhone|iPod|iPad)/i.test(navigator.platform || navigator.userAgent);

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
          <div className="logo-icon" aria-hidden="true">
            <img
              src="/excel-agent-logo.svg"
              alt="ExcelAgento logo"
              style={{ width: '24px', height: '24px', objectFit: 'contain' }}
            />
          </div>
          <span className="logo-title">
            Excel<span className="logo-title-accent">Agento</span>
          </span>
          <span className="logo-version-tag">PRO</span>
        </div>

        <a
          href="https://zenithopensourceprojects.vercel.app/os"
          target="_blank"
          rel="noopener noreferrer"
          className="zenith-nav-badge"
          title="Zenith Open Source Projects Hub (https://zenithopensourceprojects.vercel.app/os)"
        >
          <span className="zenith-nav-dot" />
          <span>Zenith OS</span>
        </a>

        <span className="nav-vertical-divider" aria-hidden="true" />

        <div className="file-meta-pill" title={`${fileName} • ${activeSheetName}`}>
          <FileSpreadsheet size={13} className="file-meta-icon" />
          <strong className="file-meta-name">{fileName}</strong>
          <span className="file-meta-dot">/</span>
          <span className="file-meta-sheet">{activeSheetName}</span>
          <span className="file-meta-dims">
            {rowCount} rows • {colCount} cols
          </span>
        </div>

        <label className="sr-only" htmlFor="top-nav-fixture">
          Load test fixtures
        </label>
        <select
          id="top-nav-fixture"
          className="select-input select-fixture-pill"
          title="Load Test Fixtures"
          onChange={(e) => {
            onSelectFixture(e.target.value);
            // Reset so the same fixture can be re-selected.
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>
            Load fixture…
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
              height: '28px',
              padding: '3px 24px 3px 26px',
              fontSize: '11.5px',
              width: '125px',
              background: 'var(--bg-elevated)',
            }}
            placeholder="Find in sheet…"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
          />
          <Search
            size={13}
            style={{
              position: 'absolute',
              left: '8px',
              color: 'var(--text-dim)',
              pointerEvents: 'none',
            }}
          />
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
            type="button"
            className="btn btn-secondary btn-sm nav-hud-btn"
            onClick={onOpenCommandPalette}
            title={`Open Command Palette (${isMac ? '⌘K' : 'Ctrl+K'})`}
          >
            <Command size={12} />
            <span>Commands</span>
            <kbd className="nav-kbd">{isMac ? '⌘K' : 'Ctrl+K'}</kbd>
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
          <Upload size={14} />
          Upload File
        </button>

        <button
          className="btn btn-primary btn-sm"
          onClick={onExport}
          title="Export current workbook to .xlsx"
        >
          <Download size={14} />
          Export .xlsx
        </button>

        <button
          className="btn btn-secondary btn-sm"
          onClick={onOpenOperationModal}
          title="Manually configure an engine operation"
        >
          <PlayCircle size={14} />
          Run Operation
        </button>
      </div>

      <div className="top-nav-right">
        <button
          className="btn btn-ghost btn-sm btn-icon"
          onClick={onUndo}
          disabled={!canUndo}
          title="Undo (Ctrl+Z)"
          aria-label="Undo"
        >
          <Undo2 size={14} />
        </button>

        <button
          className="btn btn-ghost btn-sm btn-icon"
          onClick={onRedo}
          disabled={!canRedo}
          title="Redo (Ctrl+Y)"
          aria-label="Redo"
        >
          <Redo2 size={14} />
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onToggleHistory}
          title="View Operation Audit History"
        >
          <History size={14} />
          History ({historyPosition}/{historyLength})
        </button>

        <button
          className="btn btn-ghost btn-sm nav-agents-btn"
          onClick={onOpenAgents}
          title="Meet our autonomous multi-agent workforce"
          data-testid="open-agents"
        >
          <Bot size={14} className="text-emerald" />
          <span>Agents</span>
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onOpenDocs}
          title="Architecture & Developer Docs"
          data-testid="open-docs"
        >
          <BookOpen size={14} />
          <span>Docs</span>
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onOpenUsage}
          title="Model & token usage"
          data-testid="open-usage"
        >
          <BarChart2 size={14} />
          Usage
        </button>

        <button
          className="btn btn-ghost btn-sm"
          onClick={onReset}
          title="Reset to initial workbook"
        >
          <RotateCcw size={14} />
          Reset
        </button>

        <ThemeToggle />

        <button
          className="btn btn-ghost btn-sm btn-icon"
          onClick={onOpenSettings}
          title="API Keys & Settings"
          aria-label="API Keys & Settings"
        >
          <Settings size={14} />
        </button>
      </div>
    </header>
  );
};
