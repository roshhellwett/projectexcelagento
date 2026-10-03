import React, { useState, useEffect, useMemo, useCallback } from 'react';
import {
  type Workbook,
  type Sheet,
  type ApplyOperationResult,
  type Preview,
  applyOperation,
  HistoryStack,
  cloneWorkbook,
} from '@excel-agent/engine';

import { TopNav } from './components/TopNav.js';
import { SpreadsheetGrid } from './components/SpreadsheetGrid.js';
import { AgentChat, type ChatMessage } from './components/AgentChat.js';
import { OperationModal } from './components/OperationModal.js';
import { HistoryDrawer } from './components/HistoryDrawer.js';
import { SettingsModal } from './components/SettingsModal.js';
import { CommandPalette } from './components/CommandPalette.js';
import { ToastHost, useToasts } from './components/Toaster.js';

import {
  createSampleWorkbook,
  xlsxToWorkbook,
  downloadWorkbookAsXlsx,
} from './lib/engine-adapter.js';

import { auditSheet, searchCellsInSheet, type ProposedAction } from './lib/agent-helper.js';

import { askExcelAgent } from './lib/llm-service.js';
import { orchestrator, persistMemory, registry } from './lib/agent-runtime.js';
import {
  clearSettings,
  loadSettings,
  saveSettings,
  type AgentSettings,
  type ProviderName,
} from './lib/settings.js';

const initialWorkbook = createSampleWorkbook();

export const App: React.FC = () => {
  const [workbook, setWorkbook] = useState<Workbook>(initialWorkbook);
  const [activeSheetName, setActiveSheetName] = useState<string>(
    initialWorkbook.sheets[0]?.name || 'Sheet1',
  );
  const [fileName, setFileName] = useState('sample-orders.xlsx');
  const [hasUserUploadedFile, setHasUserUploadedFile] = useState(false);

  // Search in sheet
  const [searchQuery, setSearchQuery] = useState('');
  const [isCommandPaletteOpen, setIsCommandPaletteOpen] = useState(false);

  // BYOK settings: a single source of truth shared by the chat gate and Settings modal.
  const [settings, setSettings] = useState<AgentSettings>(() => loadSettings());
  const hasApiKey = settings.apiKey.trim().length > 0;

  const handleSaveApiKey = (provider: ProviderName, key: string) => {
    const next: AgentSettings = { provider, apiKey: key };
    setSettings(next);
    saveSettings(next);

    // Greeting from ExcelAgento when the key is unlocked.
    setMessages((prev) => [
      ...prev,
      {
        id: `unlocked-${Date.now()}`,
        sender: 'assistant',
        text: `Welcome to **ExcelAgento**.\n\nUpload an Excel or CSV file to begin, or select a sample fixture from the navigation bar.\n\nOnce loaded you can ask for transformations, date formatting, deduplication, sorting, filtering, and deterministic calculations.`,
      },
    ]);
  };

  const handleClearApiKey = () => {
    setSettings({ provider: settings.provider, apiKey: '' });
    clearSettings();
  };

  const handleSettingsChange = (next: AgentSettings) => {
    setSettings(next);
    saveSettings(next);
  };

  // Toasts replace blocking alert() dialogs.
  const { toasts, pushToast, dismissToast } = useToasts();

  // Engine registry (shared runtime) and history stack
  const [historyStack, setHistoryStack] = useState<HistoryStack>(
    () => new HistoryStack(initialWorkbook, { snapshotEvery: 5 }),
  );
  const [historyRevision, setHistoryRevision] = useState(0);

  // Recently changed cells for diff highlighting in the grid
  const [recentChangedCells, setRecentChangedCells] = useState<Set<string>>(new Set());

  // UI Modals
  const [isOpModalOpen, setIsOpModalOpen] = useState(false);
  const [isHistoryDrawerOpen, setIsHistoryDrawerOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  // Cmd+K / Ctrl+K keyboard shortcut for Command Palette HUD
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setIsCommandPaletteOpen((prev) => !prev);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Current active sheet
  const currentSheet: Sheet = useMemo(
    () =>
      workbook.sheets.find((s) => s.name === activeSheetName) ||
      workbook.sheets[0] || { name: 'Sheet1', rows: [] },
    [workbook, activeSheetName],
  );

  const sheetAudit = useMemo(() => auditSheet(currentSheet), [currentSheet]);

  // Search results and highlighted cells
  const searchMatches = useMemo(() => {
    if (!searchQuery.trim()) return [];
    return searchCellsInSheet(currentSheet, searchQuery);
  }, [currentSheet, searchQuery]);

  const searchHighlightCells = useMemo(() => {
    return new Set(searchMatches.map((m) => `${m.sheet}:${m.row}:${m.column}`));
  }, [searchMatches]);

  // Initial Chat Messages
  const [messages, setMessages] = useState<ChatMessage[]>([]);

  // Execute an engine operation transactionally
  const executeOperation = useCallback(
    (name: string, input: unknown, learnQuery?: string): ApplyOperationResult => {
      setIsProcessing(true);
      try {
        const result = applyOperation(workbook, name, input, {
          registry,
          history: historyStack,
        });

        const args =
          typeof input === 'object' && input !== null ? (input as Record<string, unknown>) : {};
        const sheetName = typeof args.sheet === 'string' ? args.sheet : activeSheetName;

        if (result.ok) {
          setWorkbook(result.workbook);
          setHistoryRevision((r) => r + 1);

          // Extract changed cell coordinates for visual diff highlighting.
          const changedKeys = new Set<string>();
          for (const patchEntry of result.patch) {
            if (patchEntry.kind === 'cell') {
              const { sheet, row, column } = patchEntry.address;
              changedKeys.add(`${sheet}:${row}:${column}`);
            }
          }
          setRecentChangedCells(changedKeys);
          pushToast(
            'success',
            `${name} applied - ${result.report.affectedCells} cell(s) updated, invariants verified.`,
          );
        } else {
          pushToast('error', result.error.messages.join(' '));
        }

        // Self-learning: reinforce or decay the association for this request.
        if (learnQuery) {
          orchestrator.learn({
            query: learnQuery,
            sheetName,
            operation: name,
            args,
            success: result.ok,
          });
          persistMemory();
        }

        return result;
      } finally {
        setIsProcessing(false);
      }
    },
    [workbook, historyStack, activeSheetName, pushToast],
  );

  // Undo / Redo handlers
  const handleUndo = useCallback(() => {
    const prev = historyStack.undo();
    if (prev) {
      setWorkbook(prev);
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
    }
  }, [historyStack]);

  const handleRedo = useCallback(() => {
    const next = historyStack.redo();
    if (next) {
      setWorkbook(next);
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
    }
  }, [historyStack]);

  const handleStepBack = useCallback(
    (position: number) => {
      const restored = historyStack.stepBack(position);
      setWorkbook(restored);
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
    },
    [historyStack],
  );

  const handleReset = useCallback(() => {
    handleStepBack(0);
  }, [handleStepBack]);

  // Keyboard shortcuts (Ctrl+Z / Ctrl+Y)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) {
          handleRedo();
        } else {
          handleUndo();
        }
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleUndo, handleRedo]);

  // Load new workbook
  const loadNewWorkbook = (wb: Workbook, newFileName: string, isUserUpload = false) => {
    setWorkbook(wb);
    setActiveSheetName(wb.sheets[0]?.name || 'Sheet1');
    setFileName(newFileName);
    setHasUserUploadedFile(isUserUpload);
    const newStack = new HistoryStack(wb, { snapshotEvery: 5 });
    setHistoryStack(newStack);
    setHistoryRevision(0);
    setRecentChangedCells(new Set());
    setSearchQuery('');

    if (hasApiKey) {
      setMessages((prev) => [
        ...prev,
        {
          id: `loaded-${Date.now()}`,
          sender: 'assistant',
          text: `Loaded **"${newFileName}"** with ${wb.sheets.length} sheet(s) (${wb.sheets[0]?.rows.length ?? 0} rows). Ready to inspect, clean, sort, or transform.`,
        },
      ]);
    }
  };

  // Upload handler
  const handleFileUpload = (file: File) => {
    const MAX_BYTES = 50 * 1024 * 1024;
    if (file.size > MAX_BYTES) {
      pushToast('error', `"${file.name}" is larger than 50 MB. Please split the workbook.`);
      return;
    }

    const reader = new FileReader();
    reader.onerror = () => pushToast('error', `Could not read "${file.name}".`);
    reader.onload = (e) => {
      const buffer = e.target?.result;
      if (!(buffer instanceof ArrayBuffer)) {
        pushToast('error', `Could not read "${file.name}".`);
        return;
      }
      try {
        const wb = xlsxToWorkbook(buffer);
        const totalRows = wb.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
        if (totalRows === 0) {
          pushToast('info', `"${file.name}" loaded but contains no rows.`);
        } else {
          pushToast(
            'success',
            `Loaded "${file.name}" - ${wb.sheets.length} sheet(s), ${totalRows} rows.`,
          );
        }
        loadNewWorkbook(wb, file.name, true);
      } catch (error) {
        pushToast(
          'error',
          error instanceof Error ? error.message : `Could not parse "${file.name}".`,
        );
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Fixture switcher
  const handleSelectFixture = async (fixtureName: string) => {
    if (fixtureName === 'sample') {
      loadNewWorkbook(createSampleWorkbook(), 'sample-orders.xlsx', false);
      return;
    }

    try {
      const response = await fetch(`/fixtures/${fixtureName}`);
      if (!response.ok) throw new Error('Fixture file not found');
      const buffer = await response.arrayBuffer();
      const wb = xlsxToWorkbook(buffer);
      loadNewWorkbook(wb, fixtureName, true);
    } catch (error) {
      pushToast('error', error instanceof Error ? error.message : `Could not load ${fixtureName}`);
    }
  };

  // Export current workbook
  const handleExport = () => {
    const target = `${fileName.replace(/\.[^.]+$/, '')}-cleaned.xlsx`;
    if (downloadWorkbookAsXlsx(workbook, target)) {
      pushToast('success', `Exported "${target}".`);
    } else {
      pushToast('error', 'Export failed. Please try again.');
    }
  };

  // Chat message send handler
  const handleSendMessage = async (query: string) => {
    const userMsgId = `user-${Date.now()}`;
    const newMsg: ChatMessage = {
      id: userMsgId,
      sender: 'user',
      text: query,
    };

    setMessages((prev) => [...prev, newMsg]);
    setIsProcessing(true);

    try {
      const agentRes = await askExcelAgent(
        query,
        workbook,
        activeSheetName,
        hasApiKey ? { provider: settings.provider, apiKey: settings.apiKey } : null,
        hasUserUploadedFile,
      );

      const proposed = agentRes.proposedAction;
      let previewResult: Preview | undefined;

      if (proposed) {
        const op = registry.get(proposed.name);
        if (op) {
          try {
            const validated = op.validate(cloneWorkbook(workbook), proposed.args);
            if (validated.valid) {
              previewResult = op.preview(cloneWorkbook(workbook), proposed.args);
            }
          } catch {
            // Preview is best-effort; apply re-validates transactionally.
          }
        }
      }

      if (agentRes.source === 'fallback' && agentRes.trace) {
        const blocked = agentRes.trace.find(
          (step) => step.layer === 'llm' && step.summary.startsWith('Provider unavailable'),
        );
        if (blocked) pushToast('info', blocked.summary);
      }

      setMessages((prev) => [
        ...prev,
        {
          id: `assist-${Date.now()}`,
          sender: 'assistant',
          text: agentRes.message,
          sourceQuery: query,
          ...(proposed ? { proposedAction: proposed } : {}),
          ...(previewResult ? { preview: previewResult } : {}),
          status: 'pending',
          isStreaming: true,
        } satisfies ChatMessage,
      ]);
    } catch (error) {
      pushToast('error', error instanceof Error ? error.message : 'The agent could not respond.');
    } finally {
      setIsProcessing(false);
    }
  };

  // Apply proposed action from chat
  const handleApplyAction = (messageId: string, action: ProposedAction) => {
    const sourceQuery = messages.find((message) => message.id === messageId)?.sourceQuery;
    const result = executeOperation(action.name, action.args, sourceQuery);
    setMessages((prev) =>
      prev.map((m) => {
        if (m.id === messageId) {
          if (result.ok) {
            const affectedCount =
              result.report.affectedCells ??
              result.report.removedRows ??
              result.report.deletedColumns ??
              result.report.addedColumns ??
              0;
            return {
              ...m,
              status: 'applied',
              text: `Successfully executed **${action.name}**. Applied update (${affectedCount} changes). Invariants verified ✓`,
            };
          } else {
            return {
              ...m,
              status: 'error',
              errorMessage: result.error.messages.join(', '),
            };
          }
        }
        return m;
      }),
    );
  };

  // Quick sort action from grid column headers
  const handleQuickSort = (columnLetter: string, direction: 'asc' | 'desc') => {
    executeOperation('sort_range', {
      sheet: currentSheet.name,
      column: columnLetter,
      direction,
      startRow: 2,
      startColumn: 'A',
    });
  };

  return (
    <div className="app-container">
      {/* Top Navigation */}
      <TopNav
        key={historyRevision}
        fileName={fileName}
        activeSheetName={activeSheetName}
        rowCount={currentSheet.rows.length}
        colCount={Math.max(...currentSheet.rows.map((r) => r.length), 0)}
        canUndo={historyStack.canUndo}
        canRedo={historyStack.canRedo}
        historyLength={historyStack.length}
        historyPosition={historyStack.position}
        searchQuery={searchQuery}
        searchMatchCount={searchMatches.length}
        onSearchChange={setSearchQuery}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onReset={handleReset}
        onFileUpload={handleFileUpload}
        onExport={handleExport}
        onSelectFixture={handleSelectFixture}
        onOpenOperationModal={() => setIsOpModalOpen(true)}
        onToggleHistory={() => setIsHistoryDrawerOpen((prev) => !prev)}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onOpenCommandPalette={() => setIsCommandPaletteOpen(true)}
      />

      {/* Main Workspace Body */}
      <main className="workspace-body">
        {/* Spreadsheet Grid with Drop Zone */}
        <SpreadsheetGrid
          workbook={workbook}
          activeSheetName={activeSheetName}
          onSelectSheet={(sheet) => setActiveSheetName(sheet)}
          recentChangedCells={recentChangedCells}
          searchHighlightCells={searchHighlightCells}
          onQuickSort={handleQuickSort}
          onFileDrop={handleFileUpload}
        />

        {/* AI Agent Chat Panel */}
        <AgentChat
          audit={sheetAudit}
          messages={messages}
          isProcessing={isProcessing}
          hasApiKey={hasApiKey}
          apiKeyProvider={settings.provider}
          onSaveApiKey={handleSaveApiKey}
          onClearApiKey={handleClearApiKey}
          onSendMessage={handleSendMessage}
          onApplyAction={handleApplyAction}
          onUndoLast={handleUndo}
          canUndo={historyStack.canUndo}
        />
      </main>

      {/* Manual Operation Modal */}
      <OperationModal
        isOpen={isOpModalOpen}
        onClose={() => setIsOpModalOpen(false)}
        workbook={workbook}
        activeSheetName={activeSheetName}
        onExecute={(name, args) => executeOperation(name, args)}
      />

      {/* Operation Audit History Drawer */}
      <HistoryDrawer
        isOpen={isHistoryDrawerOpen}
        onClose={() => setIsHistoryDrawerOpen(false)}
        entries={historyStack.history}
        currentPosition={historyStack.position}
        onStepBack={handleStepBack}
      />

      {/* Settings / BYOK Modal */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        settings={settings}
        onSave={handleSettingsChange}
        onClear={handleClearApiKey}
      />

      {/* Command Palette HUD (Cmd+K) */}
      <CommandPalette
        isOpen={isCommandPaletteOpen}
        onClose={() => setIsCommandPaletteOpen(false)}
        onExecutePrompt={(prompt) => handleSendMessage(prompt)}
        onExport={handleExport}
        onToggleHistory={() => setIsHistoryDrawerOpen((prev) => !prev)}
        onOpenSettings={() => setIsSettingsOpen(true)}
        activeSheetName={activeSheetName}
      />

      {/* Non-blocking notifications */}
      <ToastHost toasts={toasts} onDismiss={dismissToast} />
    </div>
  );
};
