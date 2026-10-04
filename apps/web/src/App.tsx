import React, { useState, useEffect, useMemo, useCallback, useRef, useDeferredValue } from 'react';
import {
  type Workbook,
  type Sheet,
  type ApplyOperationResult,
  type DateSystem,
  type Preview,
  applyOperation,
  HistoryStack,
  cloneWorkbook,
  maxColumnCount,
} from '@excel-agent/engine';

import { TopNav } from './components/TopNav.js';
import { SpreadsheetGrid } from './components/SpreadsheetGrid.js';
import type { CellSelection } from './lib/selection-context.js';
import type { CellEdit } from './lib/grid-edit.js';
import { AgentChat, type ChatMessage } from './components/AgentChat.js';
import { OperationModal } from './components/OperationModal.js';
import { HistoryDrawer } from './components/HistoryDrawer.js';
import { SettingsModal } from './components/SettingsModal.js';
import { CommandPalette } from './components/CommandPalette.js';
import { ModelUsagePage } from './components/ModelUsagePage.js';
import { ErrorBoundary } from './components/ErrorBoundary.js';
import { ToastHost, useToasts } from './components/Toaster.js';

import {
  createSampleWorkbook,
  xlsxToWorkbook,
  downloadWorkbookAsXlsx,
  type ImportReport,
} from './lib/engine-adapter.js';

import {
  auditSheet,
  searchCellsInSheet,
  type ProposedAction,
  type AgentActivityEvent,
  type ExecutionPlan,
  type ChatMessage as LLMChatMessage,
} from './lib/agent-helper.js';

import { askExcelAgent } from './lib/llm-service.js';
import {
  forgetLearnedActions,
  learnedActionCount,
  orchestrator,
  persistMemory,
  registry,
} from './lib/agent-runtime.js';
import {
  appendUsageEntry,
  clearUsageLog,
  createUsageEntry,
  loadUsageLog,
  type UsageEntry,
} from './lib/usage.js';
import {
  clearSettings,
  loadSettings,
  saveSettings,
  type AgentSettings,
  type ProviderName,
} from './lib/settings.js';

const initialWorkbook = createSampleWorkbook();

/** Top-level pages. The usage view is URL-addressable via `#/usage`. */
export type WorkspaceView = 'workspace' | 'usage';

/** Renders a detected delimiter in words, since a raw tab character is invisible in a toast. */
function describeDelimiter(delimiter: string): string {
  if (delimiter === 'tab') return 'tab';
  if (delimiter === ',') return 'comma';
  if (delimiter === ';') return 'semicolon';
  if (delimiter === '|') return 'pipe';
  return delimiter;
}

function readViewFromHash(): WorkspaceView {
  try {
    if (typeof window === 'undefined') return 'workspace';
    return window.location.hash.replace(/^#\/?/, '') === 'usage' ? 'usage' : 'workspace';
  } catch {
    return 'workspace';
  }
}

export const App: React.FC = () => {
  const [workbook, setWorkbook] = useState<Workbook>(initialWorkbook);
  const [activeSheetName, setActiveSheetName] = useState<string>(
    initialWorkbook.sheets[0]?.name || 'Sheet1',
  );
  const [fileName, setFileName] = useState('sample-orders.xlsx');
  const [hasUserUploadedFile, setHasUserUploadedFile] = useState(false);
  // Which epoch the loaded workbook's serials count from; a 1904 file read as 1900 would show
  // every date four years and a day early.
  const [dateSystem, setDateSystem] = useState<DateSystem>('1900');

  // Search in sheet
  const [searchQuery, setSearchQuery] = useState('');
  const [isCommandPaletteOpen, setIsCommandPaletteOpen] = useState(false);

  // BYOK settings: a single source of truth shared by the chat gate and Settings modal.
  const [settings, setSettings] = useState<AgentSettings>(() => loadSettings());
  const hasApiKey = settings.apiKey.trim().length > 0;

  // Separate Model & Usage page, addressable at #/usage.
  const [view, setView] = useState<WorkspaceView>(() => readViewFromHash());
  const [usageEntries, setUsageEntries] = useState<UsageEntry[]>(() => loadUsageLog());
  const [learnedActions, setLearnedActions] = useState(() => learnedActionCount());

  useEffect(() => {
    const syncView = () => setView(readViewFromHash());
    window.addEventListener('hashchange', syncView);
    return () => window.removeEventListener('hashchange', syncView);
  }, []);

  const navigate = useCallback((next: WorkspaceView) => {
    setView(next);
    try {
      window.location.hash = next === 'usage' ? '#/usage' : '';
    } catch {
      // Hash updates are best-effort; the in-memory view state still switches.
    }
  }, []);

  const handleForgetLearned = () => {
    forgetLearnedActions();
    setLearnedActions(0);
    pushToast('info', 'Learned actions cleared from this browser.');
  };

  const handleClearUsage = () => {
    clearUsageLog();
    setUsageEntries([]);
    pushToast('info', 'Usage history cleared.');
  };

  const handleSaveApiKey = (provider: ProviderName, key: string, baseUrl?: string) => {
    const next: AgentSettings = { provider, apiKey: key, baseUrl };
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

  // Any uncaught error / rejected promise surfaces as a floating toast — never
  // a silent console error or a dead workspace (panel boundaries stay alive).
  useEffect(() => {
    const onWindowError = (event: ErrorEvent) => {
      pushToast('error', event.message || 'An unexpected error occurred.');
    };
    const onUnhandledRejection = (event: PromiseRejectionEvent) => {
      const reason = event.reason;
      pushToast(
        'error',
        reason instanceof Error ? reason.message : 'An operation failed unexpectedly.',
      );
    };
    window.addEventListener('error', onWindowError);
    window.addEventListener('unhandledrejection', onUnhandledRejection);
    return () => {
      window.removeEventListener('error', onWindowError);
      window.removeEventListener('unhandledrejection', onUnhandledRejection);
    };
  }, [pushToast]);

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

  // Search results and highlighted cells.
  // The matcher walks every cell, so it runs against a deferred copy of the query: typing stays
  // responsive on a large sheet because React is free to keep the previous highlight on screen
  // while the new one is computed.
  const deferredSearchQuery = useDeferredValue(searchQuery);
  const searchMatches = useMemo(() => {
    if (!deferredSearchQuery.trim()) return [];
    return searchCellsInSheet(currentSheet, deferredSearchQuery);
  }, [currentSheet, deferredSearchQuery]);

  const searchHighlightCells = useMemo(() => {
    return new Set(searchMatches.map((m) => `${m.sheet}:${m.row}:${m.column}`));
  }, [searchMatches]);

  // Initial Chat Messages
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const turnAbortRef = useRef<AbortController | null>(null);
  const [selectionContext, setSelectionContext] = useState<CellSelection | null>(null);

  // Execute an engine operation transactionally
  const executeOperation = useCallback(
    (
      name: string,
      input: unknown,
      learnQuery?: string,
      confirmed = false,
      // A grid keystroke must not raise a toast per character, but a refusal still must be loud.
      options?: { quiet?: boolean },
    ): ApplyOperationResult => {
      setIsProcessing(true);
      try {
        const result = applyOperation(workbook, name, input, {
          registry,
          history: historyStack,
          confirmed,
        });

        const args =
          typeof input === 'object' && input !== null ? (input as Record<string, unknown>) : {};
        const sheetName = typeof args.sheet === 'string' ? args.sheet : activeSheetName;

        if (result.ok) {
          setWorkbook(result.workbook);
          setHistoryRevision((r) => r + 1);

          // Automatically switch active view to newly created sheet, or fallback if active sheet was deleted
          const newlyAddedSheet = result.workbook.sheets.find(
            (s) => !workbook.sheets.some((old) => old.name === s.name),
          );
          if (newlyAddedSheet) {
            setActiveSheetName(newlyAddedSheet.name);
          } else if (!result.workbook.sheets.some((s) => s.name === activeSheetName)) {
            setActiveSheetName(result.workbook.sheets[0]?.name ?? '');
          }

          // Extract changed cell coordinates for visual diff highlighting.
          const changedKeys = new Set<string>();
          for (const patchEntry of result.patch) {
            if (patchEntry.kind === 'cell') {
              const { sheet, row, column } = patchEntry.address;
              changedKeys.add(`${sheet}:${row}:${column}`);
            }
          }
          setRecentChangedCells(changedKeys);
          if (!options?.quiet) {
            const isNoChange = result.report.affectedCells === 0;
            pushToast(
              isNoChange ? 'info' : 'success',
              isNoChange
                ? `${name} completed (0 cells changed).`
                : `${name} applied - ${result.report.affectedCells} cell(s) updated, invariants verified.`,
            );
          }
        } else {
          pushToast('error', result.error.messages.join(' '));
        }

        // Self-learning: reinforce or decay the association for this request. A change that is
        // merely awaiting confirmation is neither a success nor a rejection, so it must not
        // decay what was learned - the user has not said it was wrong, only that they have not
        // agreed to it yet.
        const awaitingConfirmation = !result.ok && result.error.code === 'confirmation-required';
        if (learnQuery && !awaitingConfirmation) {
          orchestrator.learn({
            query: learnQuery,
            sheetName,
            operation: name,
            args,
            success: result.ok,
            // Ties the learned arguments to the column layout they were correct for, so a
            // later replay can be refused if the sheet has since been reshaped.
            workbook: result.ok ? result.workbook : workbook,
          });
          persistMemory();
          setLearnedActions(learnedActionCount());
        }

        return result;
      } finally {
        setIsProcessing(false);
      }
    },
    [workbook, historyStack, activeSheetName, pushToast],
  );

  /**
   * Commits a grid edit - a typed cell, a cleared range, a paste, a fill - through the engine, so it
   * lands in the same undo stack and the same history log as every other change. It never demands
   * confirmation: the person typing is the decision.
   */
  const handleEditCells = useCallback(
    (sheet: string, edits: CellEdit[]) => {
      if (edits.length === 0) return;
      executeOperation('edit_cells', { sheet, edits }, undefined, false, { quiet: true });
    },
    [executeOperation],
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
  const loadNewWorkbook = (
    wb: Workbook,
    newFileName: string,
    isUserUpload = false,
    system: DateSystem = '1900',
  ) => {
    setWorkbook(wb);
    setActiveSheetName(wb.sheets[0]?.name || 'Sheet1');
    setFileName(newFileName);
    setHasUserUploadedFile(isUserUpload);
    setDateSystem(system);
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
  /**
   * Tells the user how their file was interpreted and what could not be kept.
   *
   * A silent import is the worst outcome: the user opens a formatted report with merged
   * titles and colour-coded status columns, is told it "loaded", works on it, exports, and
   * hands back a flattened file with no idea anything was dropped.
   */
  const reportImport = (report: ImportReport, fileName: string, totalRows: number) => {
    if (totalRows === 0) {
      pushToast('info', `"${fileName}" loaded but contains no rows.`);
      return;
    }

    const details: string[] = [];
    if (report.source === 'csv') {
      details.push(
        `read as ${report.encoding.toUpperCase()}, ${describeDelimiter(report.delimiter)}-delimited`,
      );
    }
    if (report.dateSystem === '1904') details.push('1904 date system');
    if (report.cellsConvertedToDates > 0) {
      details.push(`${report.cellsConvertedToDates} date cells`);
    }

    const summary = `Loaded "${fileName}" - ${totalRows} rows${details.length > 0 ? ` (${details.join(', ')})` : ''}.`;
    if (report.dropped.length === 0) {
      pushToast('success', summary);
      return;
    }
    pushToast(
      'warning',
      `${summary} Not carried over: ${report.dropped.join(', ')}. Exporting will not restore these.`,
    );
  };

  const handleFileUpload = (file: File) => {
    const MAX_BYTES = 50 * 1024 * 1024;
    if (file.size > MAX_BYTES) {
      pushToast('error', `"${file.name}" is larger than 50 MB. Please split the workbook.`);
      return;
    }

    const reader = new FileReader();
    reader.onerror = () => pushToast('error', `Could not read "${file.name}".`);
    reader.onload = async (e) => {
      const buffer = e.target?.result;
      if (!(buffer instanceof ArrayBuffer)) {
        pushToast('error', `Could not read "${file.name}".`);
        return;
      }
      try {
        const { workbook: wb, report } = await xlsxToWorkbook(buffer);
        const totalRows = wb.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
        loadNewWorkbook(wb, file.name, true, report.dateSystem);
        reportImport(report, file.name, totalRows);
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
      const { workbook: wb, report } = await xlsxToWorkbook(buffer);
      loadNewWorkbook(wb, fixtureName, true, report.dateSystem);
      const fixtureRows = wb.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
      reportImport(report, fixtureName, fixtureRows);
    } catch (error) {
      pushToast('error', error instanceof Error ? error.message : `Could not load ${fixtureName}`);
    }
  };

  // Export current workbook
  const handleExport = async () => {
    const target = `${fileName.replace(/\.[^.]+$/, '')}-cleaned.xlsx`;
    if (await downloadWorkbookAsXlsx(workbook, target)) {
      pushToast('success', `Exported "${target}".`);
    } else {
      pushToast('error', 'Export failed. Please try again.');
    }
  };

  // Chat message send handler
  const handleSendMessage = async (query: string) => {
    const userMsgId = `user-${Date.now()}`;
    const assistMsgId = `assist-${Date.now() + 1}`;

    const userMsg: ChatMessage = {
      id: userMsgId,
      sender: 'user',
      text: query,
    };

    const initialAssistMsg: ChatMessage = {
      id: assistMsgId,
      sender: 'assistant',
      text: '',
      sourceQuery: query,
      status: 'pending',
      isStreaming: true,
      activities: [],
    };

    setMessages((prev) => [...prev, userMsg, initialAssistMsg]);
    setIsProcessing(true);
    const turnAbort = new AbortController();
    turnAbortRef.current = turnAbort;

    try {
      // Build conversation history for multi-turn reasoning context
      const conversationHistory: LLMChatMessage[] = messages.slice(-10).map((m) => ({
        role: m.sender === 'user' ? 'user' : 'assistant',
        content: m.text,
      }));

      // When the user pins a selection chip, scope the agent's request to it.
      const agentQuery = selectionContext
        ? `Context: ${selectionContext.summary}\n\nUser request (spreadsheet operations only): ${query}`
        : query;

      const callbacks = {
        onToken: (chunk: string) => {
          setMessages((prev) =>
            prev.map((m) => (m.id === assistMsgId ? { ...m, text: m.text + chunk } : m)),
          );
        },
        onThinking: (thoughtChunk: string) => {
          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistMsgId ? { ...m, thought: (m.thought || '') + thoughtChunk } : m,
            ),
          );
        },
      };

      const onActivity = (activity: AgentActivityEvent) => {
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistMsgId ? { ...m, activities: [...(m.activities || []), activity] } : m,
          ),
        );
      };

      const agentRes = await askExcelAgent(
        agentQuery,
        workbook,
        activeSheetName,
        hasApiKey
          ? {
              provider: settings.provider,
              apiKey: settings.apiKey,
              model: settings.model,
              baseUrl: settings.baseUrl,
            }
          : null,
        hasUserUploadedFile,
        {
          conversationHistory,
          callbacks,
          onActivity,
          signal: turnAbort.signal,
        },
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

      // Telemetry ledger: real provider-reported token counts when the model ran
      setUsageEntries(
        appendUsageEntry(
          createUsageEntry({
            query,
            source: agentRes.source ?? 'heuristic',
            ...(agentRes.telemetry ? { telemetry: agentRes.telemetry } : {}),
          }),
        ),
      );

      setMessages((prev) =>
        prev.map((m) => {
          if (m.id === assistMsgId) {
            return {
              ...m,
              text: agentRes.message || m.text,
              thought: agentRes.thought || m.thought,
              activities: agentRes.activities ?? m.activities,
              proposedAction: proposed,
              plan: agentRes.plan,
              preview: previewResult,
              isStreaming: false,
              status: 'pending',
            };
          }
          return m;
        }),
      );
    } catch (error) {
      const aborted =
        turnAbort.signal.aborted || (error instanceof DOMException && error.name === 'AbortError');
      setMessages((prev) =>
        prev.map((m) =>
          m.id === assistMsgId
            ? {
                ...m,
                text:
                  m.text ||
                  (aborted ? 'Stopped.' : 'An error occurred while answering your request.'),
                isStreaming: false,
                status: aborted ? 'pending' : 'error',
                errorMessage: aborted
                  ? undefined
                  : error instanceof Error
                    ? error.message
                    : 'The agent could not respond.',
              }
            : m,
        ),
      );
      if (!aborted) {
        pushToast('error', error instanceof Error ? error.message : 'The agent could not respond.');
      }
    } finally {
      turnAbortRef.current = null;
      setIsProcessing(false);
    }
  };

  // Apply multi-step execution plan from chat
  const handleApplyPlan = (messageId: string, plan: ExecutionPlan, confirmed = false) => {
    setIsProcessing(true);
    let currentWb = workbook;
    let appliedCount = 0;
    const updatedSteps = [...plan.steps];
    let planFailed = false;
    // Remembered so a partial run can be rewound through history rather than guessed at.
    const historyPositionBefore = historyStack.position;

    try {
      for (let i = 0; i < updatedSteps.length; i++) {
        const step = updatedSteps[i];
        if (!step) continue;

        // A step the guardrail already rejected must never be executed just because it is part
        // of a batch. Applying the rest is still worth offering, so it is skipped, not fatal.
        if (step.status === 'error' && !confirmed) {
          pushToast(
            'warning',
            `Skipped Step ${i + 1} (${step.operation}): ${step.error ?? 'this step failed validation'}.`,
          );
          continue;
        }

        const result = applyOperation(currentWb, step.operation, step.args, {
          registry,
          history: historyStack,
          confirmed,
        });

        if (result.ok) {
          currentWb = result.workbook;
          appliedCount++;
          updatedSteps[i] = { ...step, status: 'completed' };
        } else if (result.error.code === 'confirmation-required') {
          planFailed = true;
          updatedSteps[i] = {
            ...step,
            status: 'error',
            error: result.error.messages.join(', '),
          };
          pushToast(
            'warning',
            `Plan needs confirmation at Step ${i + 1} (${step.operation}). Nothing was applied - review the plan and apply again to confirm.`,
          );
          break;
        } else {
          planFailed = true;
          updatedSteps[i] = {
            ...step,
            status: 'error',
            error: result.error.messages.join(', '),
          };
          pushToast(
            'error',
            `Plan stopped at Step ${i + 1} (${step.operation}): ${result.error.messages.join(', ')}`,
          );
          break;
        }
      }

      // A plan is one unit of work. If any step failed, the workbook is left exactly as it was
      // rather than half-transformed, because there is no "undo the plan" affordance for a user
      // to discover and pressing Ctrl+Z a step at a time is not a recovery strategy.
      const rolledBack = planFailed && appliedCount > 0;
      if (rolledBack) {
        pushToast(
          'info',
          `Rolled back ${appliedCount} applied step(s) so the workbook stays consistent.`,
        );
        // Rewind history to where the plan started, so the partial run leaves neither the
        // workbook nor the undo stack in a half-transformed state.
        historyStack.restore(historyPositionBefore);
      }

      setWorkbook(rolledBack ? workbook : currentWb);
      setHistoryRevision((r) => r + 1);

      if (!rolledBack) {
        const newlyAddedSheetInPlan = currentWb.sheets.find(
          (s) => !workbook.sheets.some((old) => old.name === s.name),
        );
        if (newlyAddedSheetInPlan) {
          setActiveSheetName(newlyAddedSheetInPlan.name);
        } else if (!currentWb.sheets.some((s) => s.name === activeSheetName)) {
          setActiveSheetName(currentWb.sheets[0]?.name ?? '');
        }
      }

      const finalStatus: 'applied' | 'error' = planFailed ? 'error' : 'applied';
      const updatedPlan: ExecutionPlan = {
        ...plan,
        steps: updatedSteps,
        status: finalStatus,
      };

      setMessages((prev) =>
        prev.map((m) =>
          m.id === messageId
            ? {
                ...m,
                plan: updatedPlan,
                status: finalStatus,
              }
            : m,
        ),
      );

      if (!planFailed) {
        pushToast(
          'success',
          `Plan "${plan.title}" executed (${appliedCount} steps applied). Invariants verified ✓`,
        );
      }
    } finally {
      setIsProcessing(false);
    }
  };

  // Apply proposed action from chat
  const handleApplyAction = (messageId: string, action: ProposedAction, confirmed = false) => {
    const sourceQuery = messages.find((message) => message.id === messageId)?.sourceQuery;
    const result = executeOperation(action.name, action.args, sourceQuery, confirmed);
    setMessages((prev) =>
      prev.map((m) => {
        if (m.id !== messageId) return m;
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
            confirmationPrompt: undefined,
            text: `Successfully executed **${action.name}**. Applied update (${affectedCount} changes). Invariants verified ✓`,
          };
        }
        // The engine refused pending a human decision. Move the card into its confirm state
        // instead of showing an error, so the user can approve or dismiss it.
        if (result.error.code === 'confirmation-required') {
          return {
            ...m,
            status: 'confirming',
            confirmationPrompt: {
              affectedCells: result.preview?.affectedCells ?? 0,
              reasons: result.preview?.warnings.map((warning) => warning.message) ?? [],
            },
          };
        }
        return {
          ...m,
          status: 'error',
          confirmationPrompt: undefined,
          errorMessage: result.error.messages.join(', '),
        };
      }),
    );
  };

  const handleCancelAction = (messageId: string) => {
    // A cancellation is a real signal: the agent proposed the wrong thing.
    const message = messages.find((m) => m.id === messageId);
    if (message?.proposedAction && message.sourceQuery) {
      orchestrator.learn({
        query: message.sourceQuery,
        sheetName: activeSheetName,
        operation: message.proposedAction.name,
        args: message.proposedAction.args,
        success: false,
      });
      persistMemory();
      setLearnedActions(learnedActionCount());
    }
    setMessages((prev) =>
      prev.map((m) =>
        m.id === messageId
          ? { ...m, status: 'error', confirmationPrompt: undefined, errorMessage: 'Cancelled.' }
          : m,
      ),
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

  // Dedicated Model & Usage page (kept as a separate route-like view).
  if (view === 'usage') {
    return (
      <div className="app-container">
        <ErrorBoundary variant="panel" label="Model & Usage">
          <ModelUsagePage
            settings={settings}
            entries={usageEntries}
            learnedActions={learnedActions}
            engineOperations={registry.names.length}
            toolCount={orchestrator.tools.length}
            workbookSummary={{
              fileName,
              sheetName: activeSheetName,
              sheets: workbook.sheets.length,
              rows: currentSheet.rows.length,
              cols: maxColumnCount(currentSheet.rows),
            }}
            onBack={() => navigate('workspace')}
            onOpenSettings={() => setIsSettingsOpen(true)}
            onClearUsage={handleClearUsage}
            onForgetLearned={handleForgetLearned}
          />
        </ErrorBoundary>

        <SettingsModal
          isOpen={isSettingsOpen}
          onClose={() => setIsSettingsOpen(false)}
          settings={settings}
          onSave={handleSettingsChange}
          onClear={handleClearApiKey}
        />

        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Top Navigation */}
      <TopNav
        key={historyRevision}
        fileName={fileName}
        activeSheetName={activeSheetName}
        rowCount={currentSheet.rows.length}
        colCount={maxColumnCount(currentSheet.rows)}
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
        onOpenUsage={() => navigate('usage')}
        onOpenCommandPalette={() => setIsCommandPaletteOpen(true)}
      />

      {/* Main Workspace Body */}
      <main className="workspace-body">
        {/* Spreadsheet Grid with Drop Zone */}
        <ErrorBoundary variant="panel" label="the grid">
          <SpreadsheetGrid
            workbook={workbook}
            activeSheetName={activeSheetName}
            dateSystem={dateSystem}
            onSelectSheet={(sheet) => setActiveSheetName(sheet)}
            recentChangedCells={recentChangedCells}
            searchHighlightCells={searchHighlightCells}
            onQuickSort={handleQuickSort}
            onFileDrop={handleFileUpload}
            onAddSelectionContext={setSelectionContext}
            onEditCells={handleEditCells}
          />
        </ErrorBoundary>

        {/* AI Agent Chat Panel */}
        <ErrorBoundary variant="panel" label="the agent chat">
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
            onApplyPlan={handleApplyPlan}
            onCancelAction={handleCancelAction}
            onUndoLast={handleUndo}
            canUndo={historyStack.canUndo}
            onStop={() => turnAbortRef.current?.abort()}
            selectionContext={selectionContext}
            onClearSelectionContext={() => setSelectionContext(null)}
          />
        </ErrorBoundary>
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
