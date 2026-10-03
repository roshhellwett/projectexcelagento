import React, { useState, useEffect, useRef } from 'react';

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

  useEffect(() => {
    if (isOpen) {
      setSearch('');
      setSelectedIndex(0);
      setTimeout(() => inputRef.current?.focus(), 50);
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

      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      } else if (e.key === 'ArrowDown') {
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
    <div
      className="modal-overlay"
      onClick={onClose}
      style={{ alignItems: 'flex-start', paddingTop: '12vh' }}
    >
      <div
        className="modal-card"
        onClick={(e) => e.stopPropagation()}
        style={{
          maxWidth: '620px',
          borderRadius: '8px',
          border: '1px solid var(--border-strong)',
          boxShadow:
            '0 20px 25px -5px rgba(15, 23, 42, 0.15), 0 8px 10px -6px rgba(15, 23, 42, 0.1)',
        }}
      >
        {/* Search Header */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            padding: '12px 16px',
            borderBottom: '1px solid var(--border-main)',
            background: 'var(--bg-surface)',
          }}
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="var(--text-dim)"
            strokeWidth="2"
          >
            <circle cx="11" cy="11" r="8" />
            <line x1="21" y1="21" x2="16.65" y2="16.65" />
          </svg>
          <input
            ref={inputRef}
            type="text"
            placeholder="Type a command, transformation, or audit query… (Esc to close)"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{
              flex: 1,
              background: 'transparent',
              border: 'none',
              outline: 'none',
              fontSize: '13.5px',
              fontFamily: 'var(--font-sans)',
              color: 'var(--text-main)',
            }}
          />
          <span
            style={{
              fontSize: '10px',
              fontFamily: 'var(--font-mono)',
              padding: '2px 6px',
              background: 'var(--bg-elevated)',
              border: '1px solid var(--border-main)',
              borderRadius: '4px',
              color: 'var(--text-dim)',
            }}
          >
            ESC
          </span>
        </div>

        {/* Command List */}
        <div style={{ maxHeight: '340px', overflowY: 'auto', padding: '6px' }}>
          {filtered.length === 0 ? (
            <div
              style={{
                padding: '24px',
                textAlign: 'center',
                color: 'var(--text-dim)',
                fontSize: '13px',
              }}
            >
              No commands matching &quot;{search}&quot;
            </div>
          ) : (
            filtered.map((item, index) => {
              const isSelected = index === selectedIndex;
              return (
                <div
                  key={item.id}
                  onClick={() => {
                    item.action();
                    onClose();
                  }}
                  onMouseEnter={() => setSelectedIndex(index)}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '8px 12px',
                    borderRadius: '4px',
                    cursor: 'pointer',
                    background: isSelected ? 'var(--bg-elevated)' : 'transparent',
                    border: isSelected ? '1px solid var(--border-main)' : '1px solid transparent',
                    transition: 'all 120ms ease',
                  }}
                >
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span
                        style={{
                          fontSize: '9.5px',
                          textTransform: 'uppercase',
                          letterSpacing: '0.06em',
                          fontFamily: 'var(--font-mono)',
                          padding: '1px 5px',
                          borderRadius: '3px',
                          background:
                            item.category === 'Transform'
                              ? '#eff6ff'
                              : item.category === 'Audit'
                                ? '#f0fdf4'
                                : 'var(--bg-canvas)',
                          color:
                            item.category === 'Transform'
                              ? '#1d4ed8'
                              : item.category === 'Audit'
                                ? '#166534'
                                : 'var(--text-dim)',
                          border: '1px solid var(--border-main)',
                          fontWeight: 600,
                        }}
                      >
                        {item.category}
                      </span>
                      <span
                        style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-main)' }}
                      >
                        {item.title}
                      </span>
                    </div>
                    <span
                      style={{ fontSize: '11.5px', color: 'var(--text-muted)', paddingLeft: '2px' }}
                    >
                      {item.subtitle}
                    </span>
                  </div>

                  {item.shortcut && (
                    <span
                      style={{
                        fontSize: '10.5px',
                        fontFamily: 'var(--font-mono)',
                        padding: '2px 6px',
                        borderRadius: '3px',
                        background: 'var(--bg-canvas)',
                        border: '1px solid var(--border-main)',
                        color: 'var(--text-dim)',
                      }}
                    >
                      {item.shortcut}
                    </span>
                  )}
                </div>
              );
            })
          )}
        </div>

        {/* Footer HUD info */}
        <div
          style={{
            padding: '8px 16px',
            borderTop: '1px solid var(--border-main)',
            background: 'var(--bg-canvas)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            fontSize: '11px',
            color: 'var(--text-dim)',
            fontFamily: 'var(--font-mono)',
          }}
        >
          <span>Use ↑↓ to navigate • Enter to run</span>
          <span>Deterministic Invariant Sandbox Active</span>
        </div>
      </div>
    </div>
  );
};
