import React, { useRef, useState } from 'react';
import {
  Sun,
  Moon,
  Search,
  Command,
  Upload,
  Download,
  SlidersHorizontal,
  Undo2,
  Redo2,
  History,
  BarChart2,
  RotateCcw,
  Settings,
  FileSpreadsheet,
  Bot,
  ClipboardCheck,
  BookOpen,
  ChevronRight,
  Volume2,
  VolumeX,
} from 'lucide-react';
import { getActiveTheme, toggleTheme } from '../lib/theme.js';
import { isSoundEnabled, setSoundEnabled } from '../lib/sound-effects.js';
import type { CheckpointStatus } from '../lib/workspace-recovery.js';

export function ThemeToggle() {
  const [theme, setTheme] = useState<'light' | 'dark'>(getActiveTheme());
  const label = theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme';
  return (
    <button
      type="button"
      className="btn btn-ghost btn-sm btn-icon"
      onClick={() => setTheme(toggleTheme())}
      title={label}
      aria-label={label}
      aria-pressed={theme === 'dark'}
      data-testid="theme-toggle"
      data-sound="toggle"
    >
      {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
    </button>
  );
}

export function SoundToggle() {
  const [enabled, setEnabled] = useState(isSoundEnabled);
  const label = enabled ? 'Mute interface sounds' : 'Enable interface sounds';
  return (
    <button
      type="button"
      className="btn btn-ghost btn-sm btn-icon"
      onClick={() => {
        const next = !enabled;
        setSoundEnabled(next);
        setEnabled(next);
      }}
      title={label}
      aria-label={label}
      aria-pressed={enabled}
      data-sound="toggle"
      data-testid="sound-toggle"
    >
      {enabled ? <Volume2 size={16} /> : <VolumeX size={16} />}
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
  onOpenMissions?: () => void;
  onOpenAgents?: () => void;
  onOpenDocs?: () => void;
  onOpenCommandPalette?: () => void;
  checkpointStatus?: CheckpointStatus;
  checkpointLabel?: string;
  checkpointDetail?: string;
  onClearCheckpoint?: () => void;
  hasApiKey?: boolean;
}

const FIXTURES = [
  ['Sample: Messy Orders (Built-in)', 'sample'],
  ['01: Mixed Date Formats', '01-mixed-date-formats.xlsx'],
  ['02: Merged Cells', '02-merged-cells.xlsx'],
  ['03: Blank Rows', '03-blank-rows.xlsx'],
  ['04: Numbers as Text', '04-numbers-stored-as-text.xlsx'],
  ['05: Multiple Sheets', '05-multiple-sheets.xlsx'],
  ['06: Chart File', '06-chart.xlsx'],
  ['07: Conditional Formatting', '07-conditional-formatting.xlsx'],
  ['08: Leap Days', '08-leap-days.xlsx'],
  ['09: Formulas & Empty Cells', '09-formulas-and-empty-cells.xlsx'],
  ['10: Combined Messy Orders', '10-combined-messy-orders.xlsx'],
] as const;

export const TopNav: React.FC<TopNavProps> = (props) => {
  const fileInput = useRef<HTMLInputElement>(null);
  const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);
  return (
    <header className="studio-header">
      <div className="studio-header-main">
        <a className="studio-brand" href="#" aria-label="ExcelAgento workspace">
          <span className="studio-brand-mark" aria-hidden="true">
            <img src="/excel-agent-logo.svg" alt="" />
          </span>
          <span>
            Excel<span className="studio-brand-accent">Agento</span>
            <small>THE AGENT WORKSPACE</small>
          </span>
        </a>
        <div className="file-meta-pill" title={`${props.fileName} • ${props.activeSheetName}`}>
          <FileSpreadsheet size={17} aria-hidden="true" />
          <div className="studio-file-copy">
            <strong className="file-meta-name">{props.fileName}</strong>
            <span className="file-meta-dims">
              {props.rowCount} rows • {props.colCount} cols
            </span>
            <span className="file-meta-sheet">{props.activeSheetName}</span>
          </div>
          <span className="studio-file-badge">Workbook</span>
        </div>
        <button
          type="button"
          className="studio-command-trigger"
          onClick={props.onOpenCommandPalette}
          title={`Open Command Palette (${isMac ? '⌘K' : 'Ctrl+K'})`}
        >
          <Search size={15} aria-hidden="true" />
          <span>Search commands & workflows</span>
          <kbd>{isMac ? '⌘ K' : 'Ctrl K'}</kbd>
        </button>
        <div className="studio-header-actions">
          <input
            type="file"
            ref={fileInput}
            hidden
            tabIndex={-1}
            aria-hidden="true"
            accept=".xlsx,.xls,.csv"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) props.onFileUpload(file);
              event.target.value = '';
            }}
          />
          <button
            type="button"
            className="btn btn-secondary"
            onClick={() => fileInput.current?.click()}
            title="Upload Excel or CSV file"
            data-sound="click"
          >
            <Upload size={15} />
            <span>Upload File</span>
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={props.onExport}
            title="Export current workbook to .xlsx"
            data-sound="click"
          >
            <Download size={15} />
            <span>Export .xlsx</span>
          </button>
          <ThemeToggle />
          <SoundToggle />
          <button
            type="button"
            className="btn btn-ghost btn-icon"
            onClick={props.onOpenSettings}
            aria-label="API Keys & Settings"
            title="API Keys & Settings"
            data-sound="click"
          >
            <Settings size={17} />
          </button>
        </div>
      </div>
      <div className="studio-toolbar">
        <div className="studio-breadcrumb">
          <span>Workspace</span>
          <ChevronRight size={12} />
          <strong>{props.activeSheetName}</strong>
        </div>
        <div className="studio-undo-group" aria-label="Workbook history controls">
          <button
            type="button"
            className="btn btn-ghost btn-icon"
            onClick={props.onUndo}
            disabled={!props.canUndo}
            data-sound="click"
            aria-label="Undo"
            title="Undo (Ctrl+Z)"
          >
            <Undo2 size={16} />
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-icon"
            onClick={props.onRedo}
            disabled={!props.canRedo}
            data-sound="click"
            aria-label="Redo"
            title="Redo (Ctrl+Y)"
          >
            <Redo2 size={16} />
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onToggleHistory}
            title="View Operation Audit History"
            aria-label={`History (${props.historyPosition}/${props.historyLength})`}
          >
            <History size={14} />
            <span>
              History{' '}
              <span className="studio-count">
                {props.historyPosition}/{props.historyLength}
              </span>
            </span>
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-icon"
            onClick={props.onReset}
            aria-label="Reset workbook"
            title="Reset to initial workbook"
            data-sound="click"
          >
            <RotateCcw size={14} />
          </button>
        </div>
        <div className="studio-toolbar-tools">
          <label className="studio-find">
            <Search size={13} aria-hidden="true" />
            <span className="sr-only">Find in sheet</span>
            <input
              type="search"
              placeholder="Find in sheet…"
              value={props.searchQuery}
              onChange={(event) => props.onSearchChange(event.target.value)}
            />
            {props.searchQuery && (
              <span className="studio-search-count">{props.searchMatchCount}</span>
            )}
          </label>
          <label className="studio-examples">
            <span className="sr-only">Load test fixtures</span>
            <select
              title="Load Test Fixtures"
              defaultValue=""
              onChange={(event) => {
                props.onSelectFixture(event.target.value);
                event.target.value = '';
              }}
            >
              <option value="" disabled>
                Sample workbooks
              </option>
              {FIXTURES.map(([label, value]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onOpenOperationModal}
          >
            <SlidersHorizontal size={14} />
            <span>Run Operation</span>
          </button>
        </div>
        <div className={`studio-checkpoint-status is-${props.checkpointStatus ?? 'idle'}`}>
          <span
            role="status"
            aria-live="polite"
            aria-label="Local checkpoint status"
            title={props.checkpointDetail}
          >
            <span className="checkpoint-status-dot" aria-hidden="true" />
            {props.checkpointLabel ?? 'Not checkpointed'}
          </span>
          {props.onClearCheckpoint && (
            <button
              type="button"
              className="studio-text-button"
              onClick={props.onClearCheckpoint}
              title="Delete this browser’s stored workbook checkpoint"
            >
              Clear checkpoint
            </button>
          )}
        </div>
        <nav className="studio-resource-nav" aria-label="Resources">
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onOpenAgents}
            data-testid="open-agents"
            title="Meet the agents"
          >
            <Bot size={15} />
            <span>Agents</span>
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onOpenDocs}
            data-testid="open-docs"
            title="Architecture & Developer Docs"
          >
            <BookOpen size={15} />
            <span>Docs</span>
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onOpenMissions}
            data-testid="open-missions"
            title="Mission control"
          >
            <ClipboardCheck size={15} />
            <span>Missions</span>
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={props.onOpenUsage}
            data-testid="open-usage"
            title="Model & token usage"
          >
            <BarChart2 size={15} />
            <span>Usage</span>
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-icon studio-mobile-command"
            onClick={props.onOpenCommandPalette}
            aria-label="Open command palette"
          >
            <Command size={15} />
          </button>
        </nav>
      </div>
    </header>
  );
};
