import React, { useState, useEffect, useMemo, useCallback, useRef, useDeferredValue } from 'react';
import {
  type Workbook,
  type Sheet,
  type ApplyOperationResult,
  type DateSystem,
  type Preview,
  applyOperation,
  applyOperationPlan,
  HistoryStack,
  cloneWorkbook,
  maxColumnCount,
  createWorkbookValueReader,
} from '@excel-agent/engine';
import { getCompactColumnProfiles } from '@excel-agent/agent';
import { MotionConfig } from 'framer-motion';

import { TopNav } from './components/TopNav.js';
import { WorkspaceShell, type StudioView } from './components/WorkspaceShell.js';
import { WorkbookInsights } from './components/WorkbookInsights.js';
import { WorkflowLibrary } from './components/WorkflowLibrary.js';
import { SpreadsheetGrid } from './components/SpreadsheetGrid.js';
import type { CellSelection } from './lib/selection-context.js';
import type { CellEdit } from './lib/grid-edit.js';
import { AgentChat, type ChatMessage } from './components/AgentChat.js';
import { OperationModal } from './components/OperationModal.js';
import { HistoryDrawer } from './components/HistoryDrawer.js';
import { SettingsModal } from './components/SettingsModal.js';
import { CommandPalette } from './components/CommandPalette.js';
import { ModelUsagePage } from './components/ModelUsagePage.js';
import { AgentsPage } from './components/AgentsPage.js';
import { PrivacyPolicyPage } from './components/PrivacyPolicyPage.js';
import { TermsPage } from './components/TermsPage.js';
import { DocsPage } from './components/DocsPage.js';
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
  initCloudMemory,
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

/** Top-level pages. The views are URL-addressable via `#/usage`, `#/agents`, `#/privacy`, `#/terms`, `#/docs`. */
export type WorkspaceView = 'workspace' | 'usage' | 'agents' | 'privacy' | 'terms' | 'docs';

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
    const clean = window.location.hash.replace(/^#\/?/, '').toLowerCase();
    if (clean === 'usage') return 'usage';
    if (clean === 'agents') return 'agents';
    if (clean === 'privacy') return 'privacy';
    if (clean === 'terms') return 'terms';
    if (clean === 'docs') return 'docs';
    return 'workspace';
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

  useEffect(() => {
    initCloudMemory((count) => {
      setLearnedActions(count);
    }).catch(() => {});
  }, []);

  const navigate = useCallback((next: WorkspaceView) => {
    setView(next);
    try {
      window.location.hash = next === 'workspace' ? '' : `#/${next}`;
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
  const [, setHistoryRevision] = useState(0);

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
  const [studioView, setStudioView] = useState<StudioView>('sheet');
  const [agentDraft, setAgentDraft] = useState<{ text: string; revision: number }>();
  const [revealAgentRevision, setRevealAgentRevision] = useState(0);
  const studioProfiles = useMemo(() => {
    const read = createWorkbookValueReader(workbook);
    return getCompactColumnProfiles(currentSheet, 8, (column, row) => {
      const cell = currentSheet.rows[row - 1]?.[column];
      return cell?.formula ? read(currentSheet.name, column, row) : cell?.value ?? null;
    });
  }, [workbook, currentSheet]);

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
    const turnStarted = Date.now();

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
        onTokenCount: (counts: {
          promptTokens?: number;
          completionTokens?: number;
          totalTokens?: number;
        }) => {
          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistMsgId
                ? {
                    ...m,
                    tokens: {
                      promptTokens: counts.promptTokens ?? m.tokens?.promptTokens,
                      completionTokens: counts.completionTokens ?? m.tokens?.completionTokens,
                      totalTokens: counts.totalTokens ?? m.tokens?.totalTokens,
                      isLive: true,
                    },
                  }
                : m,
            ),
          );
        },
      };

      const onActivity = (activity: AgentActivityEvent) => {
        setMessages((prev) =>
          prev.map((m) => {
            if (m.id !== assistMsgId) return m;
            const updatedTokens = activity.tokens
              ? {
                  promptTokens: activity.tokens.promptTokens ?? m.tokens?.promptTokens,
                  completionTokens: activity.tokens.completionTokens ?? m.tokens?.completionTokens,
                  totalTokens: activity.tokens.totalTokens ?? m.tokens?.totalTokens,
                  isLive: true,
                }
              : m.tokens;
            return {
              ...m,
              tokens: updatedTokens,
              activities: [...(m.activities || []), activity],
            };
          }),
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
              clarification: agentRes.clarification,
              preview: previewResult,
              tokens: agentRes.telemetry?.totalTokens
                ? {
                    promptTokens: agentRes.telemetry.promptTokens,
                    completionTokens: agentRes.telemetry.completionTokens,
                    totalTokens: agentRes.telemetry.totalTokens,
                    isLive: false,
                  }
                : m.tokens
                  ? { ...m.tokens, isLive: false }
                  : undefined,
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
      if (!aborted && hasApiKey) {
        setUsageEntries(
          appendUsageEntry(
            createUsageEntry({
              query,
              source: 'llm',
              telemetry: {
                provider: settings.provider,
                model: settings.model || 'unknown',
                latencyMs: Math.max(1, Date.now() - turnStarted),
                ok: false,
                error: error instanceof Error ? error.message : 'The agent could not respond.',
              },
            }),
          ),
        );
      }
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
    try {
      const result = applyOperationPlan(workbook, plan.steps, {
        registry, history: historyStack, confirmed, operationName: `plan: ${plan.title}`,
      });
      if (result.ok) {
        setWorkbook(result.workbook);
        setHistoryRevision((r) => r + 1);
        setRecentChangedCells(new Set());
        const newlyAddedSheetInPlan = result.workbook.sheets.find(
          (s) => !workbook.sheets.some((old) => old.name === s.name),
        );
        if (newlyAddedSheetInPlan) {
          setActiveSheetName(newlyAddedSheetInPlan.name);
        } else if (!result.workbook.sheets.some((s) => s.name === activeSheetName)) {
          setActiveSheetName(result.workbook.sheets[0]?.name ?? '');
        }
        pushToast('success', `Plan "${plan.title}" executed (${plan.steps.length} steps applied). Invariants verified ✓`);
      } else if (result.error.code !== 'confirmation-required') {
        pushToast('error', `Plan stopped at Step ${result.failedStep + 1}: ${result.error.messages.join(', ')} Nothing was applied.`);
      }
      const awaitingConfirmation = !result.ok && result.error.code === 'confirmation-required';
      setMessages((prev) => prev.map((m) => m.id !== messageId ? m : {
        ...m,
        status: result.ok ? 'applied' : awaitingConfirmation ? 'confirming' : 'error',
        confirmationPrompt: awaitingConfirmation ? {
          affectedCells: result.preview?.affectedCells ?? 0,
          reasons: result.preview?.warnings.map((warning) => warning.message) ?? [],
        } : undefined,
        errorMessage: !result.ok && !awaitingConfirmation ? result.error.messages.join(', ') : undefined,
        plan: {
          ...plan,
          status: result.ok ? 'applied' : awaitingConfirmation ? 'pending' : 'error',
          steps: plan.steps.map((step, index) => ({
            ...step,
            status: result.ok ? 'completed' : !awaitingConfirmation && index === result.failedStep ? 'error' : 'pending',
            error: !result.ok && !awaitingConfirmation && index === result.failedStep ? result.error.messages.join(', ') : undefined,
          })),
        },
      }));
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
    // Cancelling a confirmation is not evidence that the proposed operation was incorrect.
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

  const handleWorkflow = (prompt: string) => {
    if (isProcessing) return;
    setRevealAgentRevision((revision) => revision + 1);
    void handleSendMessage(prompt);
  };

  const handleDraft = (prompt: string) => {
    setRevealAgentRevision((revision) => revision + 1);
    setAgentDraft((previous) => ({ text: prompt, revision: (previous?.revision ?? 0) + 1 }));
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

  // Dedicated Autonomous Multi-Agent Workforce Showcase
  if (view === 'agents') {
    return (
      <div className="app-container">
        <ErrorBoundary variant="panel" label="Agents Showcase">
          <AgentsPage
            onBack={() => navigate('workspace')}
            onSelectPrompt={(prompt) => {
              navigate('workspace');
              setTimeout(() => handleSendMessage(prompt), 100);
            }}
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

  // Dedicated Privacy Policy & DPDP Act 2023 Statutory Notice
  if (view === 'privacy') {
    return (
      <div className="app-container">
        <ErrorBoundary variant="panel" label="Privacy Policy">
          <PrivacyPolicyPage
            onBack={() => navigate('workspace')}
            onOpenSettings={() => setIsSettingsOpen(true)}
            onForgetLearned={handleForgetLearned}
            onClearUsage={handleClearUsage}
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

  // Dedicated Terms of Service & Open Source Governance
  if (view === 'terms') {
    return (
      <div className="app-container">
        <ErrorBoundary variant="panel" label="Terms of Service">
          <TermsPage onBack={() => navigate('workspace')} />
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

  // Dedicated Developer Architecture & Documentation
  if (view === 'docs') {
    return (
      <div className="app-container">
        <ErrorBoundary variant="panel" label="Documentation">
          <DocsPage
            onBack={() => navigate('workspace')}
            onOpenSettings={() => setIsSettingsOpen(true)}
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
    <MotionConfig reducedMotion="user">
    <div className="app-container workbench">
      {/* Top Navigation */}
      <TopNav
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
        onOpenAgents={() => navigate('agents')}
        onOpenDocs={() => navigate('docs')}
        onOpenCommandPalette={() => setIsCommandPaletteOpen(true)}
      />

      {/* Main Workspace Body */}
      <WorkspaceShell view={studioView} onViewChange={setStudioView} sheetName={currentSheet.name}
        isProcessing={isProcessing} revealAgentRevision={revealAgentRevision}
        agent={
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
              learnedActions={learnedActions}
              draftPrompt={agentDraft}
              workflowProfiles={studioProfiles}
              onOpenWorkflows={() => setStudioView('workflows')}
            />
          </ErrorBoundary>
        }>
        <div className="studio-sheet-view" hidden={studioView !== 'sheet'}>
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

        </div>
        {studioView === 'insights' && <ErrorBoundary variant="panel" label="workbook insights"><WorkbookInsights profiles={studioProfiles} audit={sheetAudit} onRun={handleWorkflow} isProcessing={isProcessing} /></ErrorBoundary>}
        {studioView === 'workflows' && <ErrorBoundary variant="panel" label="workflow library"><WorkflowLibrary profiles={studioProfiles} tools={orchestrator.tools} onRun={handleWorkflow} onDraft={handleDraft} isProcessing={isProcessing} /></ErrorBoundary>}
      </WorkspaceShell>

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
        onExecutePrompt={handleWorkflow}
        onExport={handleExport}
        onToggleHistory={() => setIsHistoryDrawerOpen((prev) => !prev)}
        onOpenSettings={() => setIsSettingsOpen(true)}
        activeSheetName={activeSheetName}
      />

      {/* Workspace Footer with Zenith OS Branding and Statutory Links */}
      <footer className="workspace-footer">
        <div className="workspace-footer-left">
          <a
            href="https://zenithopensourceprojects.vercel.app/os"
            target="_blank"
            rel="noopener noreferrer"
            className="footer-brand-link"
            title="Zenith Open Source Projects Official Hub"
          >
            <span className="zenith-sparkle-dot" />
            <strong>Zenith Open Source Projects</strong>
          </a>
          <span className="footer-sep">•</span>
          <span className="footer-tagline">{isProcessing ? 'Agent working…' : 'All changes run through the verified engine'}</span>
        </div>
        <div className="workspace-footer-right">
          <button className="footer-link-btn" onClick={() => navigate('agents')}>
            Autonomous Agents
          </button>
          <span className="footer-sep">•</span>
          <button className="footer-link-btn" onClick={() => navigate('docs')}>
            Documentation
          </button>
          <span className="footer-sep">•</span>
          <button className="footer-link-btn" onClick={() => navigate('privacy')}>
            Privacy (DPDP 2023)
          </button>
          <span className="footer-sep">•</span>
          <button className="footer-link-btn" onClick={() => navigate('terms')}>
            Terms
          </button>
          <span className="footer-sep">•</span>
          <button className="footer-link-btn" onClick={() => navigate('usage')}>
            Token Ledger
          </button>
        </div>
      </footer>

      {/* Non-blocking notifications */}
      <ToastHost toasts={toasts} onDismiss={dismissToast} />
    </div>
    </MotionConfig>
  );
};
