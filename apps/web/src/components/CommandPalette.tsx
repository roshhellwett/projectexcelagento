import React, { useState, useEffect, useRef } from 'react';

import { useDialogA11y } from '../lib/use-dialog-a11y.js';

interface CommandPaletteProps {
  isOpen: boolean;
  onClose: () => void;
  onExecutePrompt: (prompt: string) => void;
  onExport: () => void;
  onToggleHistory: () => void;
  onOpenSettings: () => void;
  activeSheetName: string;
}

interface CommandItem {
  id: string;
  category: 'Transform' | 'Audit' | 'Navigation' | 'System';
  title: string;
  subtitle: string;
  shortcut?: string;
  action: () => void;
}

export const CommandPalette: React.FC<CommandPaletteProps> = ({
  isOpen,
  onClose,
  onExecutePrompt,
  onExport,
  onToggleHistory,
  onOpenSettings,
  activeSheetName,
}) => {
  const [search, setSearch] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const cardRef = useRef<HTMLDivElement>(null);

  // Escape closes, focus lands in the search field, and focus goes back to whatever opened it.
  useDialogA11y(isOpen, cardRef, onClose);

  useEffect(() => {
    if (isOpen) {
      setSearch('');
      setSelectedIndex(0);
    }
  }, [isOpen]);

  const commands: CommandItem[] = [
    {
      id: 'cmd-format-dates',
      category: 'Transform',
      title: 'Normalize All Dates to ISO 8601',
      subtitle: `Standardize irregular dates across ${activeSheetName} to YYYY-MM-DD`,
      shortcut: 'D',
      action: () => onExecutePrompt('Format all date columns to YYYY-MM-DD'),
    },
    {
      id: 'cmd-dedup',
      category: 'Transform',
      title: 'Deduplicate Rows',
      subtitle: 'Identify and remove redundant identical row records',
      shortcut: 'R',
      action: () => onExecutePrompt('Find and delete duplicate rows in the active sheet'),
    },
    {
      id: 'cmd-missing',
      category: 'Audit',
      title: 'Audit Missing & Blank Values',
      subtitle: 'Scan all columns and highlight missing fields or null values',
      shortcut: 'M',
      action: () => onExecutePrompt('Audit missing values and empty cells in this sheet'),
    },
    {
      id: 'cmd-stats',
      category: 'Audit',
      title: 'Compute Statistical Overview',
      subtitle: 'Calculate total sums, means, ranges, and cardinality for numeric columns',
      shortcut: 'S',
      action: () => onExecutePrompt('Give me a statistical summary of the numbers in this sheet'),
    },
    {
      id: 'cmd-history',
      category: 'Navigation',
      title: 'Inspect Invariant Audit History',
      subtitle: 'Open the undo stack and atomic patch history log',
      shortcut: 'H',
      action: () => onToggleHistory(),
    },
    {
      id: 'cmd-export',
      category: 'System',
      title: 'Export Verified .xlsx File',
      subtitle: 'Compile and download the workbook as a clean Excel document',
      shortcut: 'E',
      action: () => onExport(),
    },
    {
      id: 'cmd-settings',
      category: 'System',
      title: 'Configure BYOK Keys & Model Providers',
      subtitle: 'Switch between Groq, OpenRouter, and Gemini inference engines',
      action: () => onOpenSettings(),
    },
  ];

  const filtered = commands.filter(
    (c) =>
      c.title.toLowerCase().includes(search.toLowerCase()) ||
      c.subtitle.toLowerCase().includes(search.toLowerCase()) ||
      c.category.toLowerCase().includes(search.toLowerCase()),
  );

  useEffect(() => {
    setSelectedIndex(0);
  }, [search]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!isOpen) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex((prev) => (prev + 1) % Math.max(1, filtered.length));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex((prev) => (prev - 1 + filtered.length) % Math.max(1, filtered.length));
      } else if (e.key === 'Enter') {
        e.preventDefault();
        const selected = filtered[selectedIndex];
        if (selected) {
          selected.action();
          onClose();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, filtered, selectedIndex, onClose]);

  if (!isOpen) return null;

  return (
    <div className="modal-overlay command-palette-overlay" onClick={onClose}>
      <div
        className="modal-card command-palette-card"
        ref={cardRef}
        role="dialog"
        aria-modal="true"
        aria-label="Command palette"
        data-dialog-open="true"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Search Header */}
        <div className="command-palette-header">
          <svg
            className="command-palette-search-icon"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <circle cx="11" cy="11" r="8" />
            <line x1="21" y1="21" x2="16.65" y2="16.65" />
          </svg>
          <label className="sr-only" htmlFor="command-palette-input">
            Type a command, transformation, or audit query
          </label>
          <input
            ref={inputRef}
            id="command-palette-input"
            data-autofocus
            type="text"
            className="command-palette-input"
            placeholder="Type a command, transformation, or audit query… (Esc to close)"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <span className="command-palette-esc">ESC</span>
        </div>

        {/* Command List */}
        <div
          className="command-list command-palette-list"
          role="listbox"
          aria-label="Commands"
          aria-activedescendant={
            filtered[selectedIndex] ? `command-${filtered[selectedIndex]!.id}` : undefined
          }
        >
          {filtered.length === 0 ? (
            <div className="command-empty">No commands matching &quot;{search}&quot;</div>
          ) : (
            filtered.map((item, index) => {
              const isSelected = index === selectedIndex;
              const badgeClass =
                item.category === 'Transform'
                  ? 'command-badge-transform'
                  : item.category === 'Audit'
                    ? 'command-badge-audit'
                    : item.category === 'Navigation'
                      ? 'command-badge-nav'
                      : 'command-badge-system';

              return (
                <div
                  key={item.id}
                  id={`command-${item.id}`}
                  role="option"
                  aria-selected={isSelected}
                  className={`command-item ${isSelected ? 'selected' : ''}`}
                  onClick={() => {
                    item.action();
                    onClose();
                  }}
                  onMouseEnter={() => setSelectedIndex(index)}
                >
                  <div className="command-item-left">
                    <div className="command-item-header">
                      <span className={`command-badge ${badgeClass}`}>{item.category}</span>
                      <span className="command-item-title">{item.title}</span>
                    </div>
                    <span className="command-item-subtitle">{item.subtitle}</span>
                  </div>

                  {item.shortcut && <span className="command-shortcut">{item.shortcut}</span>}
                </div>
              );
            })
          )}
        </div>

        {/* Footer HUD info */}
        <div className="command-palette-footer">
          <span>Use ↑↓ to navigate • Enter to run</span>
          <span>Deterministic Invariant Sandbox Active</span>
        </div>
      </div>
    </div>
  );
};
