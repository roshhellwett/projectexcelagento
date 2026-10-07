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
  indexToColumn,
  createWorkbookValueReader,
} from '@excel-agent/engine';
import { getCompactColumnProfiles } from '@excel-agent/agent';
import { MotionConfig } from 'framer-motion';

import { TopNav } from './components/TopNav.js';
import { WorkspaceShell, type StudioView } from './components/WorkspaceShell.js';
import { WorkspaceRibbon } from './components/WorkspaceRibbon.js';
import { WorkbookInsights } from './components/WorkbookInsights.js';
import { WorkflowLibrary } from './components/WorkflowLibrary.js';
import { SpreadsheetGrid } from './components/SpreadsheetGrid.js';
import type { CellSelection } from './lib/selection-context.js';
import {
  clipboardTextForRect,
  editsForClipboardBlock,
  parseClipboardGrid,
  type CellEdit,
  type CellRect,
} from './lib/grid-edit.js';
import { AgentChat, type ChatMessage, type TaskReceipt } from './components/AgentChat.js';
import { OperationModal } from './components/OperationModal.js';
import { HistoryDrawer } from './components/HistoryDrawer.js';
import { SettingsModal } from './components/SettingsModal.js';
import { CommandPalette } from './components/CommandPalette.js';
import { ModelUsagePage } from './components/ModelUsagePage.js';
import { AgentsPage } from './components/AgentsPage.js';
import { PrivacyPolicyPage } from './components/PrivacyPolicyPage.js';
import { TermsPage } from './components/TermsPage.js';
import { DocsPage } from './components/DocsPage.js';
import { MissionControlPage } from './components/MissionControlPage.js';
import { AnalystBriefing } from './components/AnalystBriefing.js';
import { ErrorBoundary } from './components/ErrorBoundary.js';
import { ToastHost, useToasts } from './components/Toaster.js';
import { generateWorkbookStudy } from './lib/workbook-study.js';
import { AuthProvider, useAuth } from './lib/auth-context.js';
import { LandingPage } from './components/LandingPage.js';
import { AuthPage } from './components/AuthPage.js';
import { AuthCallbackPage } from './components/AuthCallbackPage.js';
import { LicensePanel } from './components/LicensePanel.js';
import { AdminLicensePage } from './components/AdminLicensePage.js';

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
  cloudMemoryConfig,
  getTabSessionId,
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
  defaultModelFor,
  DEFAULT_OPENROUTER_MODEL,
  isDemoKey,
  loadSettings,
  saveSettings,
  type AgentSettings,
  type ProviderName,
} from './lib/settings.js';

import { useDialogA11y } from './lib/use-dialog-a11y.js';
import {
  missionStore,
  missionMatchesWorkbook,
  workbookSignature,
  type MissionRecord,
  type MissionStore,
} from './lib/missions.js';
import { useMissionLedger } from './lib/use-mission-ledger.js';
import { LicenseProvider, useLicense } from './lib/license-context.js';
import { executeMissionMutation } from './lib/mission-execution.js';
import {
  workspaceRecoveryStore,
  type CheckpointStatus,
  type WorkspaceCheckpoint,
  type WorkspaceRecoveryStore,
} from './lib/workspace-recovery.js';

const initialWorkbook = createSampleWorkbook();

interface WorkbookReplacement {
  workbook: Workbook;
  fileName: string;
  isUserUpload: boolean;
  dateSystem: DateSystem;
  checkpoint?: WorkspaceCheckpoint;
  report?: ImportReport;
}

type WorkspaceDecision = { kind: 'replace'; replacement: WorkbookReplacement } | { kind: 'clear' };

/** Top-level pages. The views are URL-addressable via hash routes. */
export type WorkspaceView =
  | 'workspace'
  | 'landing'
  | 'auth'
  | 'account'
  | 'admin'
  | 'missions'
  | 'usage'
  | 'agents'
  | 'privacy'
  | 'terms'
  | 'docs';

/** Renders a detected delimiter in words, since a raw tab character is invisible in a toast. */
function describeDelimiter(delimiter: string): string {
  if (delimiter === 'tab') return 'tab';
  if (delimiter === ',') return 'comma';
  if (delimiter === ';') return 'semicolon';
  if (delimiter === '|') return 'pipe';
  return delimiter;
}

function readViewFromHash(fallback: WorkspaceView = 'landing'): {
  view: WorkspaceView;
  authMode?: 'signin' | 'signup' | 'reset' | 'verify';
} {
  try {
    if (typeof window === 'undefined') return { view: fallback };
    const clean = window.location.hash.replace(/^#\/?/, '').toLowerCase();
    if (clean === 'workspace') return { view: 'workspace' };
    if (clean === 'landing') return { view: 'landing' };
    if (clean === 'auth' || clean === 'login' || clean === 'signin')
      return { view: 'auth', authMode: 'signin' };
    if (clean === 'signup' || clean === 'register') return { view: 'auth', authMode: 'signup' };
    if (clean === 'forgot-password') return { view: 'auth', authMode: 'reset' };
    if (clean === 'verify-email') return { view: 'auth', authMode: 'verify' };
    if (clean === 'account' || clean === 'activation') return { view: 'account' };
    if (clean === 'admin' || clean === 'license-admin') return { view: 'admin' };
    if (clean === 'missions') return { view: 'missions' };
    if (clean === 'usage') return { view: 'usage' };
    if (clean === 'agents') return { view: 'agents' };
    if (clean === 'privacy') return { view: 'privacy' };
    if (clean === 'terms') return { view: 'terms' };
    if (clean === 'docs') return { view: 'docs' };
    if (clean === '') {
      if (typeof import.meta !== 'undefined' && import.meta.env?.MODE === 'test') {
        return { view: 'workspace' };
      }
      return { view: 'landing' };
    }
    return { view: fallback };
  } catch {
    return { view: fallback };
  }
}

const AppWorkspace: React.FC<{
  recoveryStore?: WorkspaceRecoveryStore;
  missionRepository?: MissionStore;
  initialView?: WorkspaceView;
}> = ({
  recoveryStore = workspaceRecoveryStore,
  missionRepository = missionStore,
  initialView,
}) => {
  const { user, loading: authLoading, callback: authCallback, dismissCallback } = useAuth();
  const { status: licenseStatus, loading: licenseLoading, error: licenseError } = useLicense();
  const [workbook, setWorkbook] = useState<Workbook>(initialWorkbook);
  const [baselineWorkbook, setBaselineWorkbook] = useState<Workbook>(initialWorkbook);
  const [activeSheetName, setActiveSheetName] = useState<string>(
    initialWorkbook.sheets[0]?.name || 'Sheet1',
  );
  const [fileName, setFileName] = useState('sample-orders.xlsx');
  const [hasUserUploadedFile, setHasUserUploadedFile] = useState(false);
  // Which epoch the loaded workbook's serials count from; a 1904 file read as 1900 would show
  // every date four years and a day early.
  const [dateSystem, setDateSystem] = useState<DateSystem>('1900');

  // Generation changes on replacement; revision changes on every edit, undo and redo.
  // Refs update synchronously so even a callback arriving before React renders is fenced out.
  const documentGeneration = useRef(0);
  const workbookRevision = useRef(0);
  const [generation, setGeneration] = useState(0);
  const unexportedRef = useRef(false);
  const [hasUnexportedChanges, setHasUnexportedChanges] = useState(false);
  const importRequest = useRef(0);
  const [workspaceDecision, setWorkspaceDecision] = useState<WorkspaceDecision | null>(null);
  const decisionRef = useRef<HTMLDivElement>(null);
  useDialogA11y(workspaceDecision !== null, decisionRef, () => setWorkspaceDecision(null));

  const [checkpointStatus, setCheckpointStatus] = useState<CheckpointStatus>('checking');
  const [checkpointError, setCheckpointError] = useState('');
  const [checkpointDeleteFailed, setCheckpointDeleteFailed] = useState(false);
  const [checkpointClearing, setCheckpointClearing] = useState(false);
  const checkpointEpoch = useRef(0);
  const checkpointMounted = useRef(true);
  useEffect(() => {
    checkpointMounted.current = true;
    return () => {
      checkpointMounted.current = false;
    };
  }, []);
  const [recoveryCandidate, setRecoveryCandidate] = useState<WorkspaceCheckpoint | null>(null);
  const [recoveryResolved, setRecoveryResolved] = useState(false);
  const [checkpointEnabled, setCheckpointEnabled] = useState(true);
  const [checkpointEligible, setCheckpointEligible] = useState(false);
  const [hasStoredCheckpoint, setHasStoredCheckpoint] = useState(false);
  const [lastCheckpointAt, setLastCheckpointAt] = useState<number | null>(null);
  const [readAttempt, setReadAttempt] = useState(0);
  const [saveAttempt, setSaveAttempt] = useState(0);
  const [recoveryNotice, setRecoveryNotice] = useState(false);
  // Also order injected storage implementations: clear must follow any already-dispatched save.
  const checkpointQueue = useRef<Promise<unknown>>(Promise.resolve());
  const queueCheckpoint = useCallback(<T,>(task: () => Promise<T>): Promise<T> => {
    const next = checkpointQueue.current.then(task, task);
    checkpointQueue.current = next.catch(() => undefined);
    return next;
  }, []);

  useEffect(() => {
    let current = true;
    setCheckpointStatus('checking');
    void queueCheckpoint(() => recoveryStore.load())
      .then((checkpoint) => {
        if (!current) return;
        setRecoveryCandidate(checkpoint);
        setHasStoredCheckpoint(checkpoint !== null);
        setLastCheckpointAt(checkpoint?.savedAt ?? null);
        setRecoveryResolved(checkpoint === null);
        setCheckpointError('');
        setCheckpointStatus(checkpoint ? 'recovery' : 'idle');
      })
      .catch((error: unknown) => {
        if (!current) return;
        // Never overwrite a checkpoint we could not read. The user can retry or clear it.
        setRecoveryResolved(false);
        setCheckpointStatus('unavailable');
        setCheckpointError(
          error instanceof Error ? error.message : 'Local checkpoint storage could not be read.',
        );
      });
    return () => {
      current = false;
    };
  }, [recoveryStore, queueCheckpoint, readAttempt]);

  useEffect(() => {
    if (!recoveryResolved || !checkpointEnabled || !checkpointEligible) return;
    let current = true;
    const savingGeneration = documentGeneration.current;
    const savingRevision = workbookRevision.current;
    const savingEpoch = checkpointEpoch.current;
    setCheckpointStatus('saving');
    const timer = setTimeout(() => {
      const checkpoint: WorkspaceCheckpoint = {
        version: 1,
        workbook,
        fileName,
        activeSheetName,
        dateSystem,
        hasUserUploadedFile,
        savedAt: Date.now(),
      };
      void queueCheckpoint(() => recoveryStore.save(checkpoint))
        .then(() => {
          if (!checkpointMounted.current || savingEpoch !== checkpointEpoch.current) return;
          // A superseded save is still the last durable checkpoint if the next write fails.
          setLastCheckpointAt(checkpoint.savedAt);
          setHasStoredCheckpoint(true);
          if (
            !current ||
            savingGeneration !== documentGeneration.current ||
            savingRevision !== workbookRevision.current
          )
            return;
          setCheckpointStatus('saved');
          setCheckpointError('');
        })
        .catch((error: unknown) => {
          if (!current) return;
          setCheckpointStatus('unavailable');
          setCheckpointError(
            error instanceof Error ? error.message : 'The workbook checkpoint was not saved.',
          );
        });
    }, 450);
    return () => {
      current = false;
      clearTimeout(timer);
    };
  }, [
    workbook,
    fileName,
    activeSheetName,
    dateSystem,
    hasUserUploadedFile,
    recoveryResolved,
    checkpointEnabled,
    checkpointEligible,
    recoveryStore,
    queueCheckpoint,
    saveAttempt,
  ]);

  useEffect(() => {
    const warnOnLeave = (event: BeforeUnloadEvent) => {
      if (!unexportedRef.current) return;
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', warnOnLeave);
    return () => window.removeEventListener('beforeunload', warnOnLeave);
  }, []);

  // Search in sheet
  const [searchQuery, setSearchQuery] = useState('');
  const [isCommandPaletteOpen, setIsCommandPaletteOpen] = useState(false);

  // BYOK settings: a single source of truth shared by the chat gate and Settings modal.
  const [settings, setSettings] = useState<AgentSettings>(() => loadSettings());
  const hasApiKey = settings.apiKey.trim().length > 0;

  // Separate Model & Usage page, addressable at #/usage.
  const [authInitialMode, setAuthInitialMode] = useState<'signin' | 'signup' | 'reset' | 'verify'>(
    () => {
      return readViewFromHash().authMode ?? 'signin';
    },
  );
  const [view, setView] = useState<WorkspaceView>(() => {
    if (initialView) return initialView;
    return readViewFromHash().view;
  });
  const [usageEntries, setUsageEntries] = useState<UsageEntry[]>(() => loadUsageLog());
  const [learnedActions, setLearnedActions] = useState(() => learnedActionCount());
  const {
    missions,
    records: missionRecords,
    storageStatus: missionStorageStatus,
    storageError: missionStorageError,
    unsavedIds: unsavedMissionIds,
    upsert: persistMission,
    update: updateMission,
    remove: removeMission,
    clear: clearMissionLedger,
    retry: retryMissionStorage,
  } = useMissionLedger(missionRepository);
  const [missionBusy, setMissionBusy] = useState(false);

  useEffect(() => {
    const syncView = () => {
      const parsed = readViewFromHash();
      setView(parsed.view);
      if (parsed.authMode) setAuthInitialMode(parsed.authMode);
    };
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
      window.location.hash = `#/${next}`;
    } catch {
      // Hash updates are best-effort; the in-memory view state still switches.
    }
  }, []);

  const openAuth = useCallback((mode: 'signin' | 'signup' | 'reset' | 'verify' = 'signin') => {
    setAuthInitialMode(mode);
    setView('auth');
    const route = {
      signin: 'auth',
      signup: 'signup',
      reset: 'forgot-password',
      verify: 'verify-email',
    }[mode];
    window.location.hash = `#/${route}`;
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

  const handleSaveApiKey = (
    provider: ProviderName,
    key: string,
    baseUrl?: string,
    model?: string,
  ) => {
    const next: AgentSettings = {
      provider,
      apiKey: key,
      baseUrl,
      model: model?.trim() || settings.model || DEFAULT_OPENROUTER_MODEL,
    };
    setSettings(next);
    saveSettings(next);

    // Greeting from ExcelAgento when the key is unlocked.
    setMessages((prev) => [
      ...prev,
      {
        id: `unlocked-${Date.now()}`,
        sender: 'assistant',
        text: `Welcome to **ExcelAgento**.\n\nOpenRouter workspace unlocked (Model: \`${next.model}\`).\n\nUpload an Excel or CSV file to begin, or select a sample fixture from the navigation bar.\n\nOnce loaded you can ask for transformations, date formatting, deduplication, sorting, filtering, and deterministic calculations.`,
      },
    ]);
  };

  const handleClearApiKey = () => {
    setSettings({ provider: settings.provider, apiKey: '', model: '' });
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
  const missionHistory = useRef(
    new Map<string, { stack: HistoryStack; position: number; messageId: string }>(),
  );

  // Recently changed cells for diff highlighting in the grid
  const [recentChangedCells, setRecentChangedCells] = useState<Set<string>>(new Set());

  // UI Modals
  const [isOpModalOpen, setIsOpModalOpen] = useState(false);
  const [operationModalOperation, setOperationModalOperation] = useState<string | undefined>();
  const [isHistoryDrawerOpen, setIsHistoryDrawerOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [gridSelection, setGridSelection] = useState<CellRect>({
    startRow: 1,
    endRow: 1,
    startColIdx: 0,
    endColIdx: 0,
  });

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
  const mutationAbortRef = useRef<AbortController | null>(null);
  const activeMissionId = useRef<string | null>(null);
  const turnSequence = useRef(0);
  useEffect(
    () => () => {
      documentGeneration.current += 1;
      turnAbortRef.current?.abort();
      mutationAbortRef.current?.abort();
    },
    [],
  );

  const markWorkbookEdited = useCallback(
    (committingMissionId?: string) => {
      if (mutationAbortRef.current && committingMissionId !== activeMissionId.current)
        mutationAbortRef.current.abort();
      for (const mission of missionRecords.current) {
        if (
          mission.id !== committingMissionId &&
          mission.status === 'prepared' &&
          mission.workbook.generation === documentGeneration.current
        ) {
          updateMission(mission.id, {
            status: 'stale',
            error: 'The workbook changed after inspection. Replan this request before applying.',
          });
        }
      }
      workbookRevision.current += 1;
      unexportedRef.current = true;
      setHasUnexportedChanges(true);
      setCheckpointEligible(true);
      setMessages((previous) =>
        previous.map((message) =>
          (message.proposedAction || message.plan) &&
          (message.status === 'pending' || message.status === 'confirming')
            ? {
                ...message,
                status: 'stale',
                confirmationPrompt: undefined,
                staleReason:
                  'The workbook changed after this preview. Preview again before applying; previous confirmation no longer applies.',
              }
            : message,
        ),
      );
    },
    [missionRecords, updateMission],
  );
  const trackMissionCommit = useCallback(
    (beforePosition: number, beforeCompacted: number, missionId?: string, messageId?: string) => {
      const shift = Math.max(0, historyStack.compactedCount - beforeCompacted);
      for (const [id, binding] of missionHistory.current) {
        if (binding.stack !== historyStack) continue;
        if (binding.position > beforePosition || (shift > 0 && binding.position <= shift + 1))
          missionHistory.current.delete(id);
        else binding.position -= shift;
      }
      if (missionId && messageId)
        missionHistory.current.set(missionId, {
          stack: historyStack,
          position: historyStack.position,
          messageId,
        });
    },
    [historyStack],
  );
  const syncMissionHistory = useCallback(() => {
    const affected = new Map<string, 'applied' | 'undone'>();
    for (const [id, binding] of missionHistory.current) {
      if (binding.stack !== historyStack) continue;
      const record = missionRecords.current.find((item) => item.id === id);
      const status = binding.position > historyStack.position ? 'undone' : 'applied';
      affected.set(binding.messageId, status);
      // Clearing persisted history must not detach a still-visible chat receipt from real undo.
      if (
        record &&
        (record.status === 'applied' || record.status === 'undone') &&
        record.status !== status
      )
        updateMission(id, {
          status,
          receipt: record.receipt ? { ...record.receipt, status } : undefined,
        });
    }
    if (affected.size)
      setMessages((previous) =>
        previous.map((message) => {
          const status = affected.get(message.id);
          return status && message.receipt && message.receipt.status !== status
            ? { ...message, receipt: { ...message.receipt, status } }
            : message;
        }),
      );
  }, [historyStack, missionRecords, updateMission]);
  const undoableMessageId = useMemo(() => {
    // Old task buttons must not undo a different task or a later manual edit.
    void historyRevision;
    return (
      [...missionHistory.current.values()].find(
        (binding) => binding.stack === historyStack && binding.position === historyStack.position,
      )?.messageId ?? ''
    );
  }, [historyRevision, historyStack]);
  const [selectionContext, setSelectionContext] = useState<CellSelection | null>(null);
  const [studioView, setStudioView] = useState<StudioView>('sheet');
  const [agentDraft, setAgentDraft] = useState<{ text: string; revision: number }>();
  const [revealAgentRevision, setRevealAgentRevision] = useState(0);
  const studioProfiles = useMemo(() => {
    const read = createWorkbookValueReader(workbook);
    return getCompactColumnProfiles(currentSheet, 8, (column, row) => {
      const cell = currentSheet.rows[row - 1]?.[column];
      return cell?.formula ? read(currentSheet.name, column, row) : (cell?.value ?? null);
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
      options?: { quiet?: boolean; missionId?: string; messageId?: string },
    ): ApplyOperationResult => {
      const beforePosition = historyStack.position;
      const beforeCompacted = historyStack.compactedCount;
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
          markWorkbookEdited(options?.missionId);
          trackMissionCommit(
            beforePosition,
            beforeCompacted,
            options?.missionId,
            options?.messageId,
          );
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
    [workbook, historyStack, activeSheetName, pushToast, markWorkbookEdited, trackMissionCommit],
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

  const openOperationModal = useCallback((operation?: string) => {
    setOperationModalOperation(operation);
    setIsOpModalOpen(true);
  }, []);

  /** Ribbon controls are local Excel-like actions. They never create a chat turn. */
  const handleRibbonManualAction = useCallback(
    (actionId: string) => {
      if (isProcessing) return;
      const sheet = currentSheet.name;
      const startColumn = indexToColumn(gridSelection.startColIdx);
      const endColumn = indexToColumn(gridSelection.endColIdx);
      const range = {
        sheet,
        startRow: gridSelection.startRow,
        endRow: gridSelection.endRow,
        startColumn,
        endColumn,
      };
      const cellsInSelection = (): CellEdit[] => {
        const edits: CellEdit[] = [];
        for (let row = gridSelection.startRow; row <= gridSelection.endRow; row += 1) {
          for (let col = gridSelection.startColIdx; col <= gridSelection.endColIdx; col += 1) {
            edits.push({ row, column: indexToColumn(col), value: null });
          }
        }
        return edits.slice(0, 20_000);
      };
      const toggleStyle = (key: 'bold' | 'italic' | 'underline' | 'wrapText') => {
        const selected = [];
        for (let row = gridSelection.startRow; row <= gridSelection.endRow; row += 1) {
          for (let col = gridSelection.startColIdx; col <= gridSelection.endColIdx; col += 1) {
            selected.push(currentSheet.rows[row - 1]?.[col]?.style?.[key] === true);
          }
        }
        return selected.length > 0 && selected.every(Boolean) ? false : true;
      };
      const styleActions: Record<string, Record<string, unknown>> = {
        bold: { bold: toggleStyle('bold') },
        italic: { italic: toggleStyle('italic') },
        underline: { underline: toggleStyle('underline') },
        wrap: { wrapText: toggleStyle('wrapText') },
        fill: { fillColor: '#FFF2CC' },
        'font-color': { fontColor: '#20342B' },
        'align-left': { horizontalAlignment: 'left' },
        'align-center': { horizontalAlignment: 'center' },
        'align-right': { horizontalAlignment: 'right' },
      };
      if (styleActions[actionId]) {
        executeOperation('format_cells', { ...range, style: styleActions[actionId] });
        return;
      }
      const numberFormats: Record<string, string> = {
        general: 'General',
        currency: '$#,##0.00',
        percent: '0.00%',
        comma: '#,##0.00',
        decimals: '0.00',
      };
      if (numberFormats[actionId]) {
        executeOperation('format_cells', {
          ...range,
          style: {},
          numberFormat: numberFormats[actionId],
        });
        return;
      }
      if (actionId === 'clear') {
        executeOperation('edit_cells', { sheet, edits: cellsInSelection() }, undefined, false, {
          quiet: true,
        });
        return;
      }
      if (actionId === 'copy' || actionId === 'cut') {
        const text = clipboardTextForRect(
          currentSheet,
          gridSelection,
          (cell) => cell?.formula ?? String(cell?.value ?? ''),
        );
        void navigator.clipboard?.writeText?.(text).catch?.(() => undefined);
        if (actionId === 'cut') {
          executeOperation('edit_cells', { sheet, edits: cellsInSelection() }, undefined, false, {
            quiet: true,
          });
        }
        pushToast('success', actionId === 'copy' ? 'Selection copied.' : 'Selection cut.');
        return;
      }
      if (actionId === 'paste') {
        void navigator.clipboard
          ?.readText?.()
          .then((text) => {
            const block = parseClipboardGrid(text);
            const edits = editsForClipboardBlock(
              block,
              gridSelection.startRow,
              gridSelection.startColIdx,
            );
            if (edits.length > 0)
              executeOperation('edit_cells', { sheet, edits }, undefined, false, { quiet: true });
          })
          .catch(() =>
            pushToast('warning', 'Clipboard access was denied. Use Ctrl+V in the grid.'),
          );
        return;
      }
      if (actionId === 'sort' || actionId === 'sort-filter') {
        openOperationModal('sort_range');
        return;
      }
      if (actionId === 'filter') {
        openOperationModal('filter_rows');
        return;
      }
      if (actionId === 'remove-duplicates') {
        executeOperation('delete_duplicates', {
          sheet,
          columns: Array.from({ length: maxColumnCount(currentSheet.rows) }, (_, index) =>
            indexToColumn(index),
          ),
          headerRow: 1,
          keep: 'first',
        });
        return;
      }
      if (actionId === 'find-select') {
        openOperationModal('find_replace');
        return;
      }
      if (actionId === 'search-help' || actionId === 'open-tools' || actionId === 'agent-tools') {
        openOperationModal();
        return;
      }
      pushToast(
        'info',
        `${actionId.replace(/-/g, ' ')} is available as a manual operation or from chat. No chat request was created.`,
      );
    },
    [currentSheet, executeOperation, gridSelection, isProcessing, openOperationModal, pushToast],
  );

  // Undo / Redo handlers
  const handleUndo = useCallback(() => {
    const prev = historyStack.undo();
    if (prev) {
      markWorkbookEdited();
      setWorkbook(prev);
      setActiveSheetName((curr) =>
        prev.sheets.some((s) => s.name === curr) ? curr : prev.sheets[0]?.name || '',
      );
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
      syncMissionHistory();
    }
  }, [historyStack, markWorkbookEdited, syncMissionHistory]);

  const handleRedo = useCallback(() => {
    const next = historyStack.redo();
    if (next) {
      markWorkbookEdited();
      setWorkbook(next);
      setActiveSheetName((curr) =>
        next.sheets.some((s) => s.name === curr) ? curr : next.sheets[0]?.name || '',
      );
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
      syncMissionHistory();
    }
  }, [historyStack, markWorkbookEdited, syncMissionHistory]);

  const handleStepBack = useCallback(
    (position: number) => {
      const restored = historyStack.stepBack(position);
      markWorkbookEdited();
      setWorkbook(restored);
      setActiveSheetName((curr) =>
        restored.sheets.some((s) => s.name === curr) ? curr : restored.sheets[0]?.name || '',
      );
      setRecentChangedCells(new Set());
      setHistoryRevision((r) => r + 1);
      syncMissionHistory();
    },
    [historyStack, markWorkbookEdited, syncMissionHistory],
  );

  const handleReset = useCallback(() => {
    handleStepBack(0);
  }, [handleStepBack]);

  // Keyboard shortcuts (Ctrl+Z / Ctrl+Y)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (document.querySelector('[data-dialog-open="true"]')) return;
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

  // Replacement is a new document, never a continuation of the previous conversation.
  const loadNewWorkbook = (replacement: WorkbookReplacement) => {
    const {
      workbook: wb,
      fileName: newFileName,
      isUserUpload,
      dateSystem: system,
      checkpoint,
      report,
    } = replacement;
    documentGeneration.current += 1;
    workbookRevision.current = 0;
    importRequest.current += 1;
    turnAbortRef.current?.abort();
    mutationAbortRef.current?.abort();
    mutationAbortRef.current = null;
    if (activeMissionId.current)
      updateMission(activeMissionId.current, {
        status: 'cancelled',
        stage: undefined,
        error: 'Workbook replaced during planning. Nothing was applied.',
      });
    activeMissionId.current = null;
    turnAbortRef.current = null;
    setIsProcessing(false);
    setGeneration(documentGeneration.current);
    missionHistory.current.clear();
    setWorkbook(wb);
    setBaselineWorkbook(wb);
    setActiveSheetName(
      checkpoint && wb.sheets.some((sheet) => sheet.name === checkpoint.activeSheetName)
        ? checkpoint.activeSheetName
        : wb.sheets[0]?.name || 'Sheet1',
    );
    setFileName(newFileName);
    setHasUserUploadedFile(isUserUpload);
    setDateSystem(system);
    setHistoryStack(new HistoryStack(wb, { snapshotEvery: 5 }));
    setHistoryRevision(0);
    setRecentChangedCells(new Set());
    setSearchQuery('');
    setSelectionContext(null);
    setAgentDraft(undefined);
    setIsOpModalOpen(false);
    setIsHistoryDrawerOpen(false);
    setIsCommandPaletteOpen(false);
    setWorkspaceDecision(null);
    // A restored checkpoint has not been exported in this new session.
    unexportedRef.current = Boolean(checkpoint);
    setHasUnexportedChanges(Boolean(checkpoint));
    setCheckpointEligible(true);
    setRecoveryNotice(Boolean(checkpoint));
    if (checkpoint) {
      setRecoveryCandidate(null);
      setRecoveryResolved(true);
      setCheckpointEnabled(true);
    }
    const initialStudyText = isUserUpload
      ? generateWorkbookStudy(wb, newFileName)
      : `${checkpoint ? 'Restored' : 'Loaded'} **"${newFileName}"** with ${wb.sheets.length} sheet(s). New workbook context started; previous proposals and conversation are cleared.${checkpoint ? ' Undo history starts fresh from this checkpoint.' : ''}`;

    setMessages([
      {
        id: `loaded-${documentGeneration.current}`,
        sender: 'assistant',
        text: initialStudyText,
      },
    ]);
    if (report)
      reportImport(
        report,
        newFileName,
        wb.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0),
      );
  };

  const requestWorkbookReplacement = (replacement: WorkbookReplacement) => {
    if (unexportedRef.current) {
      setWorkspaceDecision({ kind: 'replace', replacement });
    } else {
      loadNewWorkbook(replacement);
    }
  };

  const restoreCheckpoint = () => {
    if (!recoveryCandidate || checkpointClearing) return;
    requestWorkbookReplacement({
      workbook: recoveryCandidate.workbook,
      fileName: recoveryCandidate.fileName,
      isUserUpload: recoveryCandidate.hasUserUploadedFile,
      dateSystem: recoveryCandidate.dateSystem,
      checkpoint: recoveryCandidate,
    });
  };

  const clearCheckpoint = async () => {
    checkpointEpoch.current += 1;
    setWorkspaceDecision(null);
    setCheckpointEnabled(false);
    setCheckpointDeleteFailed(false);
    setCheckpointClearing(true);
    setCheckpointStatus('saving');
    try {
      await queueCheckpoint(() => recoveryStore.clear());
      setRecoveryCandidate(null);
      setRecoveryResolved(true);
      setHasStoredCheckpoint(false);
      setLastCheckpointAt(null);
      setCheckpointError('');
      setCheckpointStatus('off');
      pushToast(
        'info',
        'Local workbook checkpoint deleted. Checkpoints are off for this session; export to keep your work.',
      );
    } catch {
      setCheckpointDeleteFailed(true);
      setCheckpointStatus('unavailable');
      setCheckpointError(
        'Could not delete the local checkpoint. Stored workbook data may remain in this browser. Retry clearing it.',
      );
    } finally {
      setCheckpointClearing(false);
    }
  };

  const retryCheckpoint = () => {
    if (checkpointDeleteFailed) {
      void clearCheckpoint();
      return;
    }
    setCheckpointEnabled(true);
    if (!recoveryResolved) setReadAttempt((attempt) => attempt + 1);
    else {
      setCheckpointEligible(true);
      setSaveAttempt((attempt) => attempt + 1);
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

    const request = ++importRequest.current;
    const importingGeneration = documentGeneration.current;
    const isCurrentImport = () =>
      request === importRequest.current && importingGeneration === documentGeneration.current;
    const reader = new FileReader();
    reader.onerror = () => {
      if (isCurrentImport()) pushToast('error', `Could not read "${file.name}".`);
    };
    reader.onload = async (e) => {
      if (!isCurrentImport()) return;
      const buffer = e.target?.result;
      if (!(buffer instanceof ArrayBuffer)) {
        pushToast('error', `Could not read "${file.name}".`);
        return;
      }
      try {
        const { workbook: wb, report } = await xlsxToWorkbook(buffer);
        if (!isCurrentImport()) return;
        requestWorkbookReplacement({
          workbook: wb,
          fileName: file.name,
          isUserUpload: true,
          dateSystem: report.dateSystem,
          report,
        });
      } catch (error) {
        if (!isCurrentImport()) return;
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
    const request = ++importRequest.current;
    const importingGeneration = documentGeneration.current;
    const isCurrentImport = () =>
      request === importRequest.current && importingGeneration === documentGeneration.current;
    if (fixtureName === 'sample') {
      requestWorkbookReplacement({
        workbook: createSampleWorkbook(),
        fileName: 'sample-orders.xlsx',
        isUserUpload: false,
        dateSystem: '1900',
      });
      return;
    }

    try {
      const response = await fetch(`/fixtures/${fixtureName}`);
      if (!response.ok) throw new Error('Fixture file not found');
      const buffer = await response.arrayBuffer();
      const { workbook: wb, report } = await xlsxToWorkbook(buffer);
      if (!isCurrentImport()) return;
      requestWorkbookReplacement({
        workbook: wb,
        fileName: fixtureName,
        isUserUpload: true,
        dateSystem: report.dateSystem,
        report,
      });
    } catch (error) {
      if (!isCurrentImport()) return;
      pushToast('error', error instanceof Error ? error.message : `Could not load ${fixtureName}`);
    }
  };

  // Export current workbook
  const handleExport = async () => {
    const exportingGeneration = documentGeneration.current;
    const exportingRevision = workbookRevision.current;
    const target = `${fileName.replace(/\.[^.]+$/, '')}-cleaned.xlsx`;
    if (await downloadWorkbookAsXlsx(workbook, target)) {
      if (
        exportingGeneration === documentGeneration.current &&
        exportingRevision === workbookRevision.current
      ) {
        unexportedRef.current = false;
        setHasUnexportedChanges(false);
      }
      pushToast('success', `Exported "${target}".`);
    } else {
      pushToast('error', 'Export failed. Please try again.');
    }
  };

  // Chat message send handler
  const handleSendMessage = async (query: string) => {
    if (turnAbortRef.current || mutationAbortRef.current || missionBusy) return;
    const context = { generation: documentGeneration.current, revision: workbookRevision.current };
    const turnId = ++turnSequence.current;
    const userMsgId = `user-${context.generation}-${turnId}`;
    const assistMsgId = `assist-${context.generation}-${turnId}`;
    const missionId = `mission-${globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`}`;
    const startedAt = Date.now();
    const signaturePromise = workbookSignature(workbook).catch(() => null);
    persistMission({
      version: 1,
      id: missionId,
      kind: 'analysis',
      status: 'planning',
      title: query.slice(0, 110),
      request: query.slice(0, 10000),
      createdAt: startedAt,
      updatedAt: startedAt,
      workbook: { ...context, fileName, sheetName: activeSheetName, signature: null },
      stage: 'Inspecting the workbook and preparing a grounded response…',
    });
    activeMissionId.current = missionId;

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
      workbookContext: context,
      missionId,
      isStreaming: true,
      activities: [],
    };

    setMessages((prev) => [...prev, userMsg, initialAssistMsg]);
    setIsProcessing(true);
    const turnAbort = new AbortController();
    turnAbortRef.current = turnAbort;
    const belongsToDocument = () =>
      context.generation === documentGeneration.current && turnAbortRef.current === turnAbort;
    const acceptsCallback = () => belongsToDocument() && !turnAbort.signal.aborted;
    const turnStarted = Date.now();

    try {
      // Build conversation history for multi-turn reasoning context
      const conversationHistory: LLMChatMessage[] = messages.slice(-30).map((m) => ({
        role: m.sender === 'user' ? 'user' : 'assistant',
        content: m.text,
      }));

      // When the user pins a selection chip, scope the agent's request to it.
      const agentQuery = selectionContext
        ? `Context: ${selectionContext.summary}\n\nUser request (spreadsheet operations only): ${query}`
        : query;

      const callbacks = {
        onToken: (chunk: string) => {
          if (!acceptsCallback()) return;
          setMessages((prev) =>
            prev.map((m) => (m.id === assistMsgId ? { ...m, text: m.text + chunk } : m)),
          );
        },
        onThinking: (thoughtChunk: string) => {
          if (!acceptsCallback()) return;
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
          if (!acceptsCallback()) return;
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
        if (!acceptsCallback()) return;
        updateMission(missionId, { stage: activity.summary.slice(0, 1000) });
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

      const config = hasApiKey
        ? {
            provider: settings.provider,
            apiKey: settings.apiKey,
            model: settings.model,
            baseUrl: settings.baseUrl,
          }
        : null;
      const options = { conversationHistory, callbacks, onActivity, signal: turnAbort.signal };
      // The existing service scopes cloud scratchpad memory to the tab. For a connected memory
      // deployment, scope it more tightly to this workbook generation so even a late old write
      // cannot enter the next document's context. No shared settings/session state is rotated.
      const agentRes = cloudMemoryConfig.enabled
        ? await (async () => {
            const decision = await orchestrator.decide({
              query: agentQuery,
              workbook,
              sheetName: activeSheetName,
              hasUserFile: hasUserUploadedFile,
              sessionId: `${getTabSessionId()}:workbook:${context.generation}`,
              config:
                config && !isDemoKey(config.apiKey)
                  ? { ...config, model: config.model ?? defaultModelFor(config.provider) }
                  : null,
              ...options,
            });
            return {
              message: decision.message,
              thought: decision.thought,
              proposedAction: decision.action,
              plan: decision.plan,
              clarification: decision.clarification,
              source: decision.source,
              guardrail: decision.guardrail,
              trace: decision.trace,
              activities: decision.activities,
              evidence: decision.evidence,
              telemetry: decision.telemetry,
            };
          })()
        : await askExcelAgent(
            agentQuery,
            workbook,
            activeSheetName,
            config,
            hasUserUploadedFile,
            options,
          );

      if (!acceptsCallback()) return;
      const signature = await signaturePromise;
      if (!acceptsCallback()) return;
      const proposalIsStale = context.revision !== workbookRevision.current;
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

      const planBlocked = agentRes.plan?.status === 'error';
      updateMission(missionId, {
        kind: agentRes.plan ? 'plan' : proposed ? 'action' : 'analysis',
        status: proposalIsStale
          ? 'stale'
          : planBlocked
            ? 'failed'
            : proposed || agentRes.plan
              ? 'prepared'
              : 'analyzed',
        title: agentRes.plan?.title?.slice(0, 1000) ?? query.slice(0, 110),
        answer: agentRes.message?.slice(0, 20000),
        stage: undefined,
        workbook: { ...context, fileName, sheetName: activeSheetName, signature },
        action: proposed,
        plan: agentRes.plan,
        preview: previewResult,
        evidence: agentRes.evidence?.slice(0, 12),
        error: proposalIsStale
          ? 'The workbook changed during planning. Replan before applying.'
          : planBlocked
            ? 'The engine could not verify this plan. Nothing was applied.'
            : undefined,
      });

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
              evidence: agentRes.evidence ?? m.evidence,
              missionId,
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
              status: planBlocked
                ? 'error'
                : proposalIsStale && (proposed || agentRes.plan)
                  ? 'stale'
                  : 'pending',
              staleReason: proposalIsStale
                ? 'The workbook changed while this response was being prepared. Preview again against the current workbook.'
                : undefined,
            };
          }
          return m;
        }),
      );
    } catch (error) {
      if (!belongsToDocument()) return;
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
      updateMission(missionId, {
        status: aborted ? 'cancelled' : 'failed',
        stage: undefined,
        error: aborted
          ? 'Stopped by the user. No action was applied.'
          : error instanceof Error
            ? error.message.slice(0, 20000)
            : 'The agent could not respond.',
      });
      if (!aborted) {
        pushToast('error', error instanceof Error ? error.message : 'The agent could not respond.');
      }
    } finally {
      if (belongsToDocument()) {
        turnAbortRef.current = null;
        activeMissionId.current = null;
        setIsProcessing(false);
      }
    }
  };

  // Never trust a card's rendered props: recheck the live document and revision at commit time.
  const canApplyProposal = (messageId: string, confirmed: boolean): boolean => {
    if (mutationAbortRef.current || turnAbortRef.current) return false;
    const message = messages.find((candidate) => candidate.id === messageId);
    const context = message?.workbookContext;
    const mission = message?.missionId
      ? missionRecords.current.find((item) => item.id === message.missionId)
      : undefined;
    if (
      (message?.missionId && mission?.status !== 'prepared') ||
      !message ||
      !context ||
      context.generation !== documentGeneration.current ||
      context.revision !== workbookRevision.current ||
      (message.status !== 'pending' && message.status !== 'confirming') ||
      (confirmed && message.status !== 'confirming')
    ) {
      setMessages((previous) =>
        previous.map((candidate) =>
          candidate.id === messageId && candidate.status !== 'applied'
            ? {
                ...candidate,
                status: 'stale',
                confirmationPrompt: undefined,
                staleReason:
                  'This preview belongs to an earlier workbook revision. Preview again before applying.',
              }
            : candidate,
        ),
      );
      pushToast(
        'warning',
        'The workbook changed. Preview the proposal again; nothing was applied.',
      );
      return false;
    }
    return true;
  };

  const executeChatMutation = async (
    messageId: string,
    steps: { operation: string; args: Record<string, unknown> }[],
    title: string,
    confirmed: boolean,
  ) => {
    const source = messages.find((message) => message.id === messageId)!;
    const context = source.workbookContext!;
    const controller = new AbortController();
    const beforePosition = historyStack.position;
    const beforeCompacted = historyStack.compactedCount;
    mutationAbortRef.current = controller;
    activeMissionId.current = source.missionId ?? null;
    setIsProcessing(true);
    if (source.missionId)
      updateMission(source.missionId, {
        status: 'executing',
        stage: 'Staging verified changes in a private workbook…',
      });
    const engineExecId = `act-${Date.now()}-${Math.random().toString(36).slice(2, 7)}-engine`;
    const engineStartActivity: AgentActivityEvent = {
      id: engineExecId,
      type: 'tool_call',
      agent: 'Engine',
      summary: `Executing ${steps.length} operation(s) in deterministic engine...`,
      timestamp: Date.now(),
    };
    setMessages((prev) =>
      prev.map((m) =>
        m.id === messageId
          ? {
              ...m,
              activities: [...(m.activities || []), engineStartActivity],
            }
          : m,
      ),
    );
    try {
      const result = await executeMissionMutation(
        { workbook, steps, confirmed },
        {
          signal: controller.signal,
          onProgress: (event) => {
            if (!controller.signal.aborted) {
              if (source.missionId)
                updateMission(source.missionId, {
                  stage: `Step ${event.index + 1}/${steps.length}: ${event.phase} ${steps[event.index]?.operation ?? 'operation'}`,
                });
              const stepOp = steps[event.index]?.operation ?? 'operation';
              setMessages((prev) =>
                prev.map((m) => {
                  if (m.id !== messageId) return m;
                  const acts = m.activities || [];
                  const lastAct = acts[acts.length - 1];
                  if (lastAct && lastAct.id === engineExecId) {
                    return {
                      ...m,
                      activities: [
                        ...acts.slice(0, -1),
                        {
                          ...lastAct,
                          summary: `Step ${event.index + 1}/${steps.length}: Applying ${stepOp}...`,
                        },
                      ],
                    };
                  }
                  return m;
                }),
              );
            }
          },
        },
      );
      if (
        controller.signal.aborted ||
        context.generation !== documentGeneration.current ||
        context.revision !== workbookRevision.current
      )
        throw new DOMException(
          'Execution was cancelled or the workbook changed. Nothing was committed.',
          'AbortError',
        );
      if (result.ok && result.patch.length) {
        historyStack.commit(title, result.workbook, result.patch, result.inverse);
        markWorkbookEdited(source.missionId);
        trackMissionCommit(beforePosition, beforeCompacted, source.missionId, messageId);
        setWorkbook(result.workbook);
        setHistoryRevision((revision) => revision + 1);
        const added = result.workbook.sheets.find(
          (sheet) => !workbook.sheets.some((old) => old.name === sheet.name),
        );
        if (added) setActiveSheetName(added.name);
        else if (!result.workbook.sheets.some((sheet) => sheet.name === activeSheetName))
          setActiveSheetName(result.workbook.sheets[0]?.name ?? '');
        setRecentChangedCells(
          new Set(
            result.patch.flatMap((entry) =>
              entry.kind === 'cell'
                ? [`${entry.address.sheet}:${entry.address.row}:${entry.address.column}`]
                : [],
            ),
          ),
        );
        setMessages((prev) =>
          prev.map((m) => {
            if (m.id !== messageId) return m;
            const acts = m.activities || [];
            const lastAct = acts[acts.length - 1];
            if (lastAct && lastAct.id === engineExecId) {
              return {
                ...m,
                activities: [
                  ...acts.slice(0, -1),
                  {
                    ...lastAct,
                    summary: `Successfully applied ${result.steps.length} operation(s). Invariants verified ✓`,
                  },
                ],
              };
            }
            return m;
          }),
        );
      }
      if (!result.ok && result.error.code === 'confirmation-required')
        pushToast(
          'warning',
          'This change was not confirmed. Review the affected cells and confirm before applying.',
        );
      if (source.missionId)
        updateMission(source.missionId, {
          stage: undefined,
          ...(!result.ok && result.error.code === 'confirmation-required'
            ? { status: 'prepared' as const }
            : {}),
        });
      return result;
    } finally {
      if (mutationAbortRef.current === controller) {
        mutationAbortRef.current = null;
        activeMissionId.current = null;
        setIsProcessing(false);
      }
    }
  };
  const failChatMutation = (messageId: string, error: unknown) => {
    const source = messages.find((message) => message.id === messageId);
    if (source?.workbookContext?.generation !== documentGeneration.current) return;
    const stale = source.workbookContext.revision !== workbookRevision.current;
    const cancelled = error instanceof Error && error.name === 'AbortError';
    const reason = stale
      ? 'The workbook changed during execution. Staged changes were discarded; replan before applying.'
      : cancelled
        ? 'Stopped by the user. Nothing was committed.'
        : error instanceof Error
          ? error.message
          : 'Mission execution failed. Nothing was committed.';
    if (source.missionId)
      updateMission(source.missionId, {
        status: stale ? 'stale' : cancelled ? 'cancelled' : 'failed',
        stage: undefined,
        error: reason,
      });
    setMessages((previous) =>
      previous.map((message) =>
        message.id === messageId
          ? {
              ...message,
              status: stale ? 'stale' : 'error',
              errorMessage: reason,
              staleReason: stale ? reason : undefined,
              confirmationPrompt: undefined,
            }
          : message,
      ),
    );
    pushToast(cancelled || stale ? 'warning' : 'error', reason);
  };

  // Apply multi-step execution plan from chat, atomically and off the UI thread when supported.
  const handleApplyPlan = async (messageId: string, plan: ExecutionPlan, confirmed = false) => {
    if (!canApplyProposal(messageId, confirmed)) return;
    const missionId = messages.find((message) => message.id === messageId)?.missionId;
    try {
      const result = await executeChatMutation(
        messageId,
        plan.steps,
        `plan: ${plan.title}`,
        confirmed,
      );
      if (result.ok) {
        const newlyAddedSheetInPlan = result.workbook.sheets.find(
          (s) => !workbook.sheets.some((old) => old.name === s.name),
        );
        if (newlyAddedSheetInPlan) {
          setActiveSheetName(newlyAddedSheetInPlan.name);
        } else if (!result.workbook.sheets.some((s) => s.name === activeSheetName)) {
          setActiveSheetName(result.workbook.sheets[0]?.name ?? '');
        }
        pushToast(
          'success',
          `Plan "${plan.title}" executed (${plan.steps.length} steps applied). Invariants verified ✓`,
        );
      } else if (result.error.code !== 'confirmation-required') {
        pushToast(
          'error',
          `Plan stopped at Step ${result.failedStep + 1}: ${result.error.messages.join(', ')} Nothing was applied.`,
        );
      }
      const awaitingConfirmation = !result.ok && result.error.code === 'confirmation-required';
      if (missionId && !awaitingConfirmation) {
        updateMission(
          missionId,
          result.ok
            ? {
                status: 'applied',
                receipt: {
                  id: `receipt-${Date.now()}`,
                  status: 'applied',
                  createdAt: Date.now(),
                  completedAt: Date.now(),
                  operations: result.steps.map((step, index) => ({
                    name: plan.steps[index]?.operation ?? 'step',
                    affectedCells: step.report.affectedCells,
                    warnings: step.report.warnings.map((warning) => warning.message),
                  })),
                },
              }
            : { status: 'failed', error: result.error.messages.join(' ') },
        );
      }
      setMessages((prev) =>
        prev.map((m) =>
          m.id !== messageId
            ? m
            : {
                ...m,
                status: result.ok ? 'applied' : awaitingConfirmation ? 'confirming' : 'error',
                receipt: result.ok
                  ? ({
                      id: `receipt-${Date.now()}`,
                      status: 'applied',
                      request: m.sourceQuery ?? plan.title,
                      createdAt: Date.now(),
                      completedAt: Date.now(),
                      workbookContext: {
                        generation: documentGeneration.current,
                        revision: workbookRevision.current,
                        fileName,
                        sheetName: activeSheetName,
                      },
                      operations: result.steps.map((step, index) => ({
                        name: plan.steps[index]?.operation ?? 'step',
                        affectedCells: step.report.affectedCells,
                        warnings: step.report.warnings.map((warning) => warning.message),
                      })),
                    } satisfies TaskReceipt)
                  : !awaitingConfirmation
                    ? ({
                        id: `receipt-${Date.now()}`,
                        status: 'failed',
                        request: m.sourceQuery ?? plan.title,
                        createdAt: Date.now(),
                        completedAt: Date.now(),
                        operations: [
                          {
                            name: plan.title,
                            affectedCells: 0,
                            warnings: result.error.messages,
                          },
                        ],
                      } satisfies TaskReceipt)
                    : undefined,
                confirmationPrompt: awaitingConfirmation
                  ? {
                      affectedCells: result.preview?.affectedCells ?? 0,
                      reasons: result.preview?.warnings.map((warning) => warning.message) ?? [],
                    }
                  : undefined,
                errorMessage:
                  !result.ok && !awaitingConfirmation
                    ? result.error.messages.join(', ')
                    : undefined,
                plan: {
                  ...plan,
                  status: result.ok ? 'applied' : awaitingConfirmation ? 'pending' : 'error',
                  steps: plan.steps.map((step, index) => ({
                    ...step,
                    status: result.ok
                      ? 'completed'
                      : !awaitingConfirmation && index === result.failedStep
                        ? 'error'
                        : 'pending',
                    error:
                      !result.ok && !awaitingConfirmation && index === result.failedStep
                        ? result.error.messages.join(', ')
                        : undefined,
                  })),
                },
              },
        ),
      );
    } catch (error) {
      failChatMutation(messageId, error);
    }
  };

  // Apply proposed action through the same staged worker/confirmation contract.
  const handleApplyAction = async (
    messageId: string,
    action: ProposedAction,
    confirmed = false,
  ) => {
    if (!canApplyProposal(messageId, confirmed)) return;
    const sourceMessage = messages.find((message) => message.id === messageId);
    const missionId = sourceMessage?.missionId;
    const sourceQuery = sourceMessage?.sourceQuery;
    try {
      const outcome = await executeChatMutation(
        messageId,
        [{ operation: action.name, args: action.args }],
        action.name,
        confirmed,
      );
      const result = outcome.ok
        ? { ...outcome, report: outcome.steps[0]!.report, preview: outcome.steps[0]!.preview }
        : outcome;
      if (result.ok)
        pushToast(
          result.report.affectedCells === 0 ? 'info' : 'success',
          `${action.name} applied - ${result.report.affectedCells} cell(s) updated, invariants verified.`,
        );
      else if (result.error.code !== 'confirmation-required')
        pushToast('error', result.error.messages.join(' '));
      if (sourceQuery && (result.ok || result.error.code !== 'confirmation-required')) {
        orchestrator.learn({
          query: sourceQuery,
          sheetName: activeSheetName,
          operation: action.name,
          args: action.args,
          success: result.ok,
          workbook: result.ok ? result.workbook : workbook,
        });
        persistMemory();
        setLearnedActions(learnedActionCount());
      }
      if (missionId && (result.ok || result.error.code !== 'confirmation-required')) {
        updateMission(
          missionId,
          result.ok
            ? {
                status: 'applied',
                receipt: {
                  id: `receipt-${Date.now()}`,
                  status: 'applied',
                  createdAt: Date.now(),
                  completedAt: Date.now(),
                  operations: [
                    {
                      name: action.name,
                      affectedCells: result.report.affectedCells,
                      warnings: result.report.warnings.map((warning) => warning.message),
                    },
                  ],
                },
              }
            : { status: 'failed', error: result.error.messages.join(' ') },
        );
      }
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
              receipt: {
                id: `receipt-${Date.now()}`,
                status: 'applied',
                request: m.sourceQuery ?? action.name,
                createdAt: Date.now(),
                completedAt: Date.now(),
                workbookContext: {
                  generation: documentGeneration.current,
                  revision: workbookRevision.current,
                  fileName,
                  sheetName: activeSheetName,
                },
                operations: [
                  {
                    name: action.name,
                    affectedCells: result.report.affectedCells,
                    warnings: result.report.warnings.map((warning) => warning.message),
                  },
                ],
              } satisfies TaskReceipt,
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
            receipt: {
              id: `receipt-${Date.now()}`,
              status: 'failed',
              request: m.sourceQuery ?? action.name,
              createdAt: Date.now(),
              completedAt: Date.now(),
              workbookContext: {
                generation: documentGeneration.current,
                revision: workbookRevision.current,
                fileName,
                sheetName: activeSheetName,
              },
              operations: [
                {
                  name: action.name,
                  affectedCells: 0,
                  warnings: result.error.messages,
                },
              ],
            } satisfies TaskReceipt,
            confirmationPrompt: undefined,
            errorMessage: result.error.messages.join(', '),
          };
        }),
      );
    } catch (error) {
      failChatMutation(messageId, error);
    }
  };

  const handleCancelAction = (messageId: string) => {
    const missionId = messages.find((message) => message.id === messageId)?.missionId;
    if (missionId)
      updateMission(missionId, {
        status: 'cancelled',
        error: 'Cancelled before execution. Nothing was applied.',
      });
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

  const handleResumeMission = async (mission: MissionRecord) => {
    if (
      isProcessing ||
      missionBusy ||
      missionRecords.current.find((item) => item.id === mission.id)?.status !== 'prepared'
    )
      return;
    const context = { generation: documentGeneration.current, revision: workbookRevision.current };
    setMissionBusy(true);
    try {
      const matches = await missionMatchesWorkbook(mission, workbook);
      if (
        context.generation !== documentGeneration.current ||
        context.revision !== workbookRevision.current
      )
        throw new Error(
          'The workbook changed while verifying this mission. Replan before applying.',
        );
      if (!matches)
        throw new Error(
          'This workbook does not match the content inspected by the mission. Restore the matching checkpoint, or replan the request.',
        );
      let action = mission.action;
      let plan = mission.plan;
      let preview: Preview | undefined;
      if (action) {
        const op = registry.get(action.name);
        if (!op) throw new Error('The saved operation is no longer available.');
        const parsed = op.schema.safeParse(action.args);
        if (!parsed.success) throw new Error('The saved operation arguments are invalid.');
        action = { ...action, args: parsed.data as Record<string, unknown> };
        const guardrail = orchestrator.guardrail(workbook, action);
        if (!guardrail.passed) throw new Error(guardrail.errors.join(' '));
        preview = guardrail.preview;
      } else if (plan) {
        const simulation = await executeMissionMutation({
          workbook,
          steps: plan.steps,
          confirmed: true,
        });
        if (!simulation.ok) throw new Error(simulation.error.messages.join(' '));
        plan = {
          ...plan,
          status: 'pending',
          steps: plan.steps.map((step, index) => ({
            ...step,
            status: 'pending',
            error: undefined,
            preview: simulation.steps[index]?.preview,
          })),
        };
      } else throw new Error('The saved mission has no complete operation to review.');
      if (
        context.generation !== documentGeneration.current ||
        context.revision !== workbookRevision.current
      )
        throw new Error(
          'The workbook changed while rebuilding the preview. Replan before applying.',
        );
      if (missionRecords.current.find((item) => item.id === mission.id)?.status !== 'prepared')
        return;
      updateMission(mission.id, {
        action,
        plan,
        preview,
        workbook: { ...mission.workbook, ...context, fileName, sheetName: activeSheetName },
        error: undefined,
      });
      const resumed: ChatMessage = {
        id: `mission-resume-${mission.id}-${turnSequence.current++}`,
        sender: 'assistant',
        text: `Resumed mission **${mission.title}**. Its workbook content matched and the engine generated a fresh preview. Review it before applying; previous confirmation is not restored.`,
        sourceQuery: mission.request,
        missionId: mission.id,
        proposedAction: action,
        plan,
        preview,
        evidence: mission.evidence,
        status: 'pending',
        workbookContext: context,
      };
      setMessages((previous) => [
        ...previous.map((message) =>
          message.missionId === mission.id && message.status !== 'applied'
            ? { ...message, status: 'stale' as const, confirmationPrompt: undefined }
            : message,
        ),
        resumed,
      ]);
      setRevealAgentRevision((revision) => revision + 1);
      navigate('workspace');
    } catch (error) {
      const reason = error instanceof Error ? error.message : 'Mission verification failed.';
      updateMission(mission.id, { status: 'stale', error: reason });
      pushToast('warning', `${reason} Nothing was applied.`);
    } finally {
      setMissionBusy(false);
    }
  };
  const handleReplanMission = (mission: MissionRecord) => {
    if (isProcessing || missionBusy) return;
    navigate('workspace');
    handleDraft(mission.request);
    pushToast(
      'info',
      'Request loaded into the composer. Review it and send to create a new mission for this workbook.',
    );
  };
  const handleDownloadMission = (mission: MissionRecord) => {
    const blob = new Blob(
      [JSON.stringify({ ...mission, exportedAt: new Date().toISOString() }, null, 2)],
      { type: 'application/json' },
    );
    const url = URL.createObjectURL(blob);
    try {
      const link = document.createElement('a');
      link.href = url;
      link.download = `${mission.id}.json`;
      link.click();
    } finally {
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  };
  const handleDeleteMission = async (mission: MissionRecord) => {
    await removeMission(mission.id);
    setMessages((previous) =>
      previous.map((message) =>
        message.missionId === mission.id
          ? {
              ...message,
              status: 'stale',
              confirmationPrompt: undefined,
              staleReason: 'This mission was removed. Replan the request before applying.',
            }
          : message,
      ),
    );
  };
  const handleClearMissions = async () => {
    await clearMissionLedger();
    setMessages((previous) =>
      previous.map((message) =>
        message.missionId && message.status !== 'applied'
          ? {
              ...message,
              status: 'stale',
              confirmationPrompt: undefined,
              staleReason: 'Mission history was cleared. Replan before applying.',
            }
          : message,
      ),
    );
  };

  const handleStopTask = () => {
    turnAbortRef.current?.abort();
    mutationAbortRef.current?.abort();
    mutationAbortRef.current = null;
    if (activeMissionId.current)
      updateMission(activeMissionId.current, {
        status: 'cancelled',
        stage: undefined,
        error: 'Stopped by the user. Nothing was applied.',
      });
    activeMissionId.current = null;
    turnAbortRef.current = null;
    setIsProcessing(false);
    setMessages((previous) =>
      previous.map((message) =>
        message.isStreaming
          ? { ...message, isStreaming: false, text: message.text || 'Stopped.', status: 'error' }
          : message,
      ),
    );
  };

  const checkpointLabel: Record<CheckpointStatus, string> = {
    checking: 'Checking local checkpoint…',
    idle: 'Not checkpointed yet',
    recovery: 'Recovery available',
    saving: 'Saving checkpoint…',
    saved: 'Checkpoint saved locally',
    unavailable: 'Checkpoint unavailable',
    off: 'Checkpoints off',
  };
  const checkpointDetail =
    checkpointStatus === 'saved'
      ? 'Stored in this browser only. Export an .xlsx for a portable copy. Undo history is not checkpointed.'
      : checkpointStatus === 'unavailable'
        ? 'Your current edits are in memory. Export to keep them; only the last successful checkpoint can be restored.'
        : 'Local checkpoints stay in this browser. They do not replace an exported workbook.';

  // The public overview and documentation remain readable without an account. Workbook access,
  // mission history, and model usage are account-gated so there is one unambiguous entry point.
  // The test bootstrap intentionally opens the workspace directly so the existing engine/UI
  // coverage can exercise workbook behavior without depending on a remote auth service.
  const protectedView =
    view === 'workspace' ||
    view === 'missions' ||
    view === 'usage' ||
    view === 'account' ||
    view === 'admin';
  const licenseGatedView = view === 'workspace' || view === 'missions' || view === 'usage';
  const isTestMode = import.meta.env.MODE === 'test';
  if (authCallback) {
    return (
      <div className="app-container app-page-scroll">
        <ErrorBoundary variant="panel" label="Email confirmation">
          <AuthCallbackPage
            callback={authCallback}
            onContinue={() => {
              dismissCallback();
              navigate('workspace');
            }}
            onSignIn={() => {
              dismissCallback();
              openAuth('signin');
            }}
            onRequestRecovery={() => {
              dismissCallback();
              openAuth('reset');
            }}
            onRequestVerification={() => {
              dismissCallback();
              openAuth('verify');
            }}
          />
        </ErrorBoundary>
        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }
  if (protectedView && !isTestMode && authLoading) {
    return (
      <div className="app-container app-page-scroll auth-gate-loading">
        <div className="auth-gate-loading-card" role="status" aria-live="polite">
          <span className="auth-gate-loading-mark">
            <img src="/excel-agent-logo.svg" alt="" />
          </span>
          <strong>Checking workspace access</strong>
          <span>Preparing your account session…</span>
        </div>
      </div>
    );
  }
  if (protectedView && !isTestMode && !user) {
    return (
      <div className="app-container app-page-scroll">
        <ErrorBoundary variant="panel" label="Authentication Page">
          <AuthPage
            initialMode="signin"
            onSuccess={() => navigate(view)}
            onNavigateBack={() => navigate('landing')}
          />
        </ErrorBoundary>
        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }

  if (protectedView && !isTestMode && user && licenseLoading) {
    return (
      <div className="app-container app-page-scroll auth-gate-loading">
        <div className="auth-gate-loading-card" role="status" aria-live="polite">
          <span className="auth-gate-loading-mark">
            <img src="/excel-agent-logo.svg" alt="" />
          </span>
          <strong>Checking activation access</strong>
          <span>Verifying your account and installation…</span>
        </div>
      </div>
    );
  }

  if (licenseGatedView && !isTestMode && user && (!licenseStatus?.canUse || licenseError)) {
    return (
      <LicensePanel
        locked
        onBack={() => navigate('landing')}
        onOpenAdmin={licenseStatus?.isAdmin ? () => navigate('admin') : undefined}
      />
    );
  }

  if (view === 'account') {
    return (
      <LicensePanel
        onBack={() => navigate('workspace')}
        onOpenAdmin={licenseStatus?.isAdmin ? () => navigate('admin') : undefined}
      />
    );
  }

  if (view === 'admin') {
    return <AdminLicensePage onBack={() => navigate('workspace')} />;
  }

  // Durable mission control: prepared tasks and applied receipts survive navigation and refresh.
  if (view === 'missions') {
    return (
      <div className="app-container app-page-scroll">
        <ErrorBoundary variant="panel" label="Mission Control">
          <MissionControlPage
            missions={missions}
            storageStatus={missionStorageStatus}
            storageError={missionStorageError}
            unsavedIds={unsavedMissionIds}
            busy={isProcessing || missionBusy}
            activeMissionId={activeMissionId.current}
            onStop={(mission) => {
              if (activeMissionId.current === mission.id) handleStopTask();
            }}
            onBack={() => navigate('workspace')}
            onResume={(mission) => void handleResumeMission(mission)}
            onReplan={handleReplanMission}
            onDelete={handleDeleteMission}
            onClear={handleClearMissions}
            onRetryStorage={() => void retryMissionStorage()}
            onDownload={handleDownloadMission}
          />
        </ErrorBoundary>
        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }

  // Dedicated Model & Usage page (kept as a separate route-like view).
  if (view === 'usage') {
    return (
      <div className="app-container app-page-scroll">
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

  // Dedicated Precision Landing Page
  if (view === 'landing') {
    return (
      <div className="app-container app-page-scroll">
        <ErrorBoundary variant="panel" label="Landing Page">
          <LandingPage
            onLaunchWorkspace={() => {
              if (user) {
                navigate('workspace');
              } else {
                openAuth('signup');
              }
            }}
            onOpenAuth={openAuth}
            onOpenAgents={() => navigate('agents')}
            onOpenDocs={() => navigate('docs')}
            onOpenPrivacy={() => navigate('privacy')}
          />
        </ErrorBoundary>

        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }

  // Dedicated Supabase Auth Page (Login / Signup)
  if (view === 'auth') {
    return (
      <div className="app-container app-page-scroll">
        <ErrorBoundary variant="panel" label="Authentication Page">
          <AuthPage
            initialMode={authInitialMode}
            onSuccess={() => navigate('workspace')}
            onNavigateBack={() => navigate('landing')}
          />
        </ErrorBoundary>

        <ToastHost toasts={toasts} onDismiss={dismissToast} />
      </div>
    );
  }

  // Dedicated Autonomous Multi-Agent Workforce Showcase
  if (view === 'agents') {
    return (
      <div className="app-container app-page-scroll">
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
      <div className="app-container app-page-scroll">
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
      <div className="app-container app-page-scroll">
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
      <div className="app-container app-page-scroll">
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
          hasApiKey={hasApiKey}
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
          onOpenOperationModal={() => openOperationModal()}
          onToggleHistory={() => setIsHistoryDrawerOpen((prev) => !prev)}
          onOpenSettings={() => setIsSettingsOpen(true)}
          onOpenUsage={() => navigate('usage')}
          onOpenMissions={() => navigate('missions')}
          onOpenAgents={() => navigate('agents')}
          onOpenDocs={() => navigate('docs')}
          onOpenCommandPalette={() => setIsCommandPaletteOpen(true)}
          onOpenLanding={() => navigate('landing')}
          onOpenAuth={openAuth}
          onOpenAccount={() => navigate('account')}
          onOpenAdmin={() => navigate('admin')}
          isAdmin={licenseStatus?.isAdmin}
          checkpointStatus={checkpointStatus}
          checkpointLabel={
            checkpointClearing ? 'Deleting checkpoint…' : checkpointLabel[checkpointStatus]
          }
          checkpointDetail={checkpointDetail}
          onClearCheckpoint={
            hasStoredCheckpoint || checkpointStatus === 'unavailable'
              ? () => setWorkspaceDecision({ kind: 'clear' })
              : undefined
          }
        />

        {recoveryCandidate && (
          <section
            className="workspace-checkpoint-banner"
            aria-labelledby="checkpoint-recovery-title"
          >
            <div className="workspace-checkpoint-copy">
              <span className="studio-eyebrow">LOCAL RECOVERY · THIS BROWSER ONLY</span>
              <h2 id="checkpoint-recovery-title">Resume your last successful checkpoint</h2>
              <p>
                <strong>{recoveryCandidate.fileName}</strong> ·{' '}
                {new Date(recoveryCandidate.savedAt).toLocaleString()}. Restore the saved workbook
                with fresh undo history and a new conversation. Changes made after this checkpoint
                are not recoverable.
              </p>
            </div>
            <div className="workspace-checkpoint-actions">
              <button
                type="button"
                className="btn btn-primary"
                onClick={restoreCheckpoint}
                disabled={checkpointClearing}
              >
                Restore checkpoint
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setWorkspaceDecision({ kind: 'clear' })}
                disabled={checkpointClearing}
              >
                Discard checkpoint
              </button>
            </div>
          </section>
        )}
        {checkpointStatus === 'unavailable' && (
          <section
            className="workspace-checkpoint-banner is-warning"
            role="alert"
            aria-label="Checkpoint storage warning"
          >
            <div className="workspace-checkpoint-copy">
              <h2>Your work is still editable, but not checkpointed</h2>
              <p>
                {checkpointError} Export to keep current changes.
                {lastCheckpointAt !== null &&
                  ` Only the last successful checkpoint (${new Date(lastCheckpointAt).toLocaleString()}) is recoverable.`}
              </p>
            </div>
            <div className="workspace-checkpoint-actions">
              <button type="button" className="btn btn-secondary" onClick={retryCheckpoint}>
                {checkpointDeleteFailed ? 'Retry checkpoint deletion' : 'Retry local storage'}
              </button>
              <button type="button" className="btn btn-primary" onClick={handleExport}>
                Export workbook
              </button>
            </div>
          </section>
        )}
        {checkpointStatus === 'off' && (
          <section
            className="workspace-checkpoint-banner is-compact"
            aria-label="Local checkpoints disabled"
          >
            <p>
              Local checkpoint deleted. Automatic checkpoints are off for this session. Export to
              keep your work.
            </p>
            <button type="button" className="btn btn-secondary btn-sm" onClick={retryCheckpoint}>
              Enable local checkpoints
            </button>
          </section>
        )}
        {recoveryNotice && (
          <section className="workspace-checkpoint-banner is-compact" role="status">
            <p>
              Restored the last successful checkpoint. Undo history starts fresh; previous
              conversations and proposals were not restored.
            </p>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              onClick={() => setRecoveryNotice(false)}
            >
              Dismiss recovery notice
            </button>
          </section>
        )}

        {/* Main Workspace Body */}
        <WorkspaceShell
          view={studioView}
          onViewChange={setStudioView}
          sheetName={currentSheet.name}
          isProcessing={isProcessing}
          revealAgentRevision={revealAgentRevision}
          ribbon={
            <WorkspaceRibbon
              activeSheetName={currentSheet.name}
              isProcessing={isProcessing}
              onManualAction={handleRibbonManualAction}
              onOpenOperationModal={() => openOperationModal()}
              onUploadFile={handleFileUpload}
              onExport={handleExport}
              onUndo={handleUndo}
              onRedo={handleRedo}
              canUndo={historyStack.canUndo}
              canRedo={historyStack.canRedo}
            />
          }
          agent={
            <ErrorBoundary variant="panel" label="the agent chat">
              <AgentChat
                key={generation}
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
                undoableMessageId={undoableMessageId}
                onStop={handleStopTask}
                selectionContext={selectionContext}
                onClearSelectionContext={() => setSelectionContext(null)}
                learnedActions={learnedActions}
                draftPrompt={agentDraft}
                workflowProfiles={studioProfiles}
                onOpenWorkflows={() => setStudioView('workflows')}
              />
            </ErrorBoundary>
          }
        >
          <div className="studio-sheet-view" hidden={studioView !== 'sheet'}>
            {/* Spreadsheet Grid with Drop Zone */}
            <ErrorBoundary variant="panel" label="the grid">
              <SpreadsheetGrid
                key={generation}
                workbook={workbook}
                activeSheetName={activeSheetName}
                dateSystem={dateSystem}
                onSelectSheet={(sheet) => setActiveSheetName(sheet)}
                recentChangedCells={recentChangedCells}
                searchHighlightCells={searchHighlightCells}
                onQuickSort={handleQuickSort}
                onFileDrop={handleFileUpload}
                onAddSelectionContext={setSelectionContext}
                onSelectionChange={setGridSelection}
                onEditCells={handleEditCells}
              />
            </ErrorBoundary>
          </div>
          {studioView === 'insights' && (
            <ErrorBoundary variant="panel" label="workbook insights">
              <div className="studio-insights-stack">
                <div className="briefing-baseline-controls">
                  <p>
                    Comparison uses this document’s opening snapshot, or the checkpoint you
                    restored. Baselines stay in memory and reset on refresh or file replacement.
                  </p>
                  <button
                    type="button"
                    className="btn btn-secondary btn-sm"
                    disabled={isProcessing || baselineWorkbook === workbook}
                    onClick={() => setBaselineWorkbook(workbook)}
                  >
                    Use current workbook as baseline
                  </button>
                </div>
                <AnalystBriefing
                  workbook={workbook}
                  sheetName={activeSheetName}
                  baseline={baselineWorkbook}
                />
                <details className="briefing-workflow-suggestions">
                  <summary>Suggested analyst workflows</summary>
                  <WorkbookInsights
                    profiles={studioProfiles}
                    audit={sheetAudit}
                    onRun={handleWorkflow}
                    isProcessing={isProcessing}
                  />
                </details>
              </div>
            </ErrorBoundary>
          )}
          {studioView === 'workflows' && (
            <ErrorBoundary variant="panel" label="workflow library">
              <WorkflowLibrary
                profiles={studioProfiles}
                tools={orchestrator.tools}
                onRun={handleWorkflow}
                onDraft={handleDraft}
                isProcessing={isProcessing}
              />
            </ErrorBoundary>
          )}
        </WorkspaceShell>

        {workspaceDecision && (
          <div className="workspace-decision-overlay">
            <div
              ref={decisionRef}
              className="workspace-decision-card"
              role="dialog"
              aria-modal="true"
              aria-labelledby="workspace-decision-title"
              aria-describedby="workspace-decision-description"
              data-dialog-open="true"
            >
              <span className="studio-eyebrow">YOU’RE IN CONTROL</span>
              <h2 id="workspace-decision-title">
                {workspaceDecision.kind === 'clear'
                  ? 'Delete local checkpoint?'
                  : 'Replace workbook with unexported changes?'}
              </h2>
              <p id="workspace-decision-description">
                {workspaceDecision.kind === 'clear'
                  ? 'This deletes the stored workbook from this browser and turns automatic checkpoints off for this session. Your open workbook remains editable. Export first if you need a copy.'
                  : `“${fileName}” has changes you have not exported. Replacing it with “${workspaceDecision.replacement.fileName}” clears its undo history, conversation and proposals. A checkpoint is not an exported copy. Export first, or explicitly replace it.`}
              </p>
              <div className="workspace-checkpoint-actions">
                <button
                  type="button"
                  className="btn btn-secondary"
                  data-autofocus
                  onClick={() => setWorkspaceDecision(null)}
                >
                  Keep working
                </button>
                <button type="button" className="btn btn-secondary" onClick={handleExport}>
                  Export current workbook
                </button>
                <button
                  type="button"
                  className="btn btn-danger"
                  onClick={() =>
                    workspaceDecision.kind === 'clear'
                      ? void clearCheckpoint()
                      : loadNewWorkbook(workspaceDecision.replacement)
                  }
                >
                  {workspaceDecision.kind === 'clear' ? 'Delete checkpoint' : 'Replace workbook'}
                </button>
              </div>
              <small>Tab to move between choices · Esc to keep working</small>
            </div>
          </div>
        )}

        {/* Manual Operation Modal */}
        <OperationModal
          isOpen={isOpModalOpen}
          onClose={() => setIsOpModalOpen(false)}
          workbook={workbook}
          activeSheetName={activeSheetName}
          initialOperation={operationModalOperation}
          operationCatalog={orchestrator.tools}
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
            <span className="footer-tagline">
              {isProcessing
                ? 'Agent working…'
                : hasUnexportedChanges
                  ? 'Unexported changes · Export for a portable copy'
                  : 'All changes run through the verified engine'}
            </span>
          </div>
          <div className="workspace-footer-right">
            <button className="footer-link-btn" onClick={() => navigate('landing')}>
              Product Overview
            </button>
            <span className="footer-sep">•</span>
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

export const App: React.FC<{
  recoveryStore?: WorkspaceRecoveryStore;
  missionRepository?: MissionStore;
  initialView?: WorkspaceView;
}> = (props) => {
  return (
    <AuthProvider>
      <LicenseProvider>
        <AppWorkspace {...props} />
      </LicenseProvider>
    </AuthProvider>
  );
};
