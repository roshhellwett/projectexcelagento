import {
  applyOperation,
  cloneWorkbook,
  createOperationRegistry,
  indexToColumn,
  maxColumnCount,
  type OperationRegistry,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';

import {
  getColumnProfiles,
  isUnstructuredSourceSheet,
  type ColumnMetadata,
  type ProposedAction,
} from './analysis.js';
import { analyzeReconciliationIntent } from './reconciliation-intent.js';
import type { ExecutionPlan, ExecutionPlanStep } from './types.js';

export const WORKFLOW_TEMPLATE_IDS = [
  'clean_weekly_export',
  'normalize_dates',
  'remove_duplicates',
  'reconcile_sheets',
] as const;

export type WorkflowTemplateId = (typeof WORKFLOW_TEMPLATE_IDS)[number];

export interface WorkflowTemplateDefinition {
  id: WorkflowTemplateId;
  title: string;
  description: string;
  category: 'Clean' | 'Compare';
  assumptions: string[];
}

export const WORKFLOW_TEMPLATES: readonly WorkflowTemplateDefinition[] = [
  {
    id: 'clean_weekly_export',
    title: 'Clean weekly export',
    description:
      'Create a cleaned copy for handoff while preserving the original export for comparison.',
    category: 'Clean',
    assumptions: [
      'The source sheet stays unchanged; cleaning is written to a new worksheet.',
      'The engine trims text, collapses whitespace, drops blank padding, and conservatively coerces numeric text.',
    ],
  },
  {
    id: 'normalize_dates',
    title: 'Normalize dates',
    description: 'Standardize detected date columns to the unambiguous ISO YYYY-MM-DD format.',
    category: 'Clean',
    assumptions: [
      'Only columns with date-like headers or values are selected.',
      'Ambiguous or invalid date cells remain unchanged and are disclosed as engine warnings.',
    ],
  },
  {
    id: 'remove_duplicates',
    title: 'Remove duplicates',
    description: 'Remove repeated records while keeping the first occurrence after review.',
    category: 'Clean',
    assumptions: [
      'A single unambiguous identifier column is used when one is detected; otherwise the full row is the key.',
      'The engine confirmation gate is required because rows are removed.',
    ],
  },
  {
    id: 'reconcile_sheets',
    title: 'Reconcile & compare sheets',
    description:
      'Compare two source tables and create matched, exception, summary, and methodology sheets.',
    category: 'Compare',
    assumptions: [
      'Both source sheets must have an unambiguous shared record key.',
      'Amount comparison is used only when one comparable amount column exists on each side; otherwise the template asks for input.',
      'Source sheets remain unchanged and all report sheets are value-only outputs.',
    ],
  },
] as const;

export type WorkflowTemplateResult =
  | {
      ok: true;
      template: WorkflowTemplateDefinition;
      plan: ExecutionPlan;
      assumptions: string[];
    }
  | {
      ok: false;
      template: WorkflowTemplateDefinition;
      kind: 'unsupported' | 'needs-input' | 'no-op';
      reason: string;
    };

type WorkflowTemplateFailureKind = Extract<WorkflowTemplateResult, { ok: false }>['kind'];

export interface WorkflowTemplateDiscovery {
  template: WorkflowTemplateDefinition;
  result: WorkflowTemplateResult;
}

interface TemplateStep {
  operation: string;
  args: Record<string, unknown>;
  description: string;
}

function template(id: WorkflowTemplateId): WorkflowTemplateDefinition {
  return WORKFLOW_TEMPLATES.find((candidate) => candidate.id === id)!;
}

function activeSheet(workbook: Workbook, activeSheetName: string): Sheet | undefined {
  return workbook.sheets.find((sheet) => sheet.name === activeSheetName) ?? workbook.sheets[0];
}

function hasTabularData(sheet: Sheet | undefined): boolean {
  return Boolean(
    sheet &&
    sheet.rows.length > 1 &&
    maxColumnCount(sheet.rows) > 0 &&
    !isUnstructuredSourceSheet(sheet),
  );
}

function safeSheetName(workbook: Workbook, desired: string): string {
  const cleaned =
    desired
      .replace(/[\\/:?*\u005b\u005d]/g, '_')
      .replace(/^'+|'+$/g, '')
      .trim()
      .slice(0, 31) || 'Workflow Output';
  const names = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  if (!names.has(cleaned.toLowerCase())) return cleaned;
  for (let suffix = 2; suffix < 1000; suffix += 1) {
    const suffixText = ` (${suffix})`;
    const candidate = `${cleaned.slice(0, 31 - suffixText.length)}${suffixText}`;
    if (!names.has(candidate.toLowerCase())) return candidate;
  }
  return `${cleaned.slice(0, 25)} (${Date.now() % 100000})`;
}

function unsupported(
  id: WorkflowTemplateId,
  kind: WorkflowTemplateFailureKind,
  reason: string,
): WorkflowTemplateResult {
  return { ok: false, template: template(id), kind, reason };
}

function buildVerifiedPlan(
  workbook: Workbook,
  registry: OperationRegistry,
  id: WorkflowTemplateId,
  title: string,
  description: string,
  assumptions: string[],
  steps: TemplateStep[],
): WorkflowTemplateResult {
  if (steps.length === 0)
    return unsupported(id, 'no-op', 'There is no supported change to review.');

  let simulated = cloneWorkbook(workbook);
  const planSteps: ExecutionPlanStep[] = [];
  const noOpReasons: string[] = [];

  for (const [index, step] of steps.entries()) {
    const result = applyOperation(simulated, step.operation, step.args, {
      registry,
      // This is private planning only. The actual workspace still applies the resulting plan
      // through its normal confirmation gate.
      confirmed: true,
    });
    if (!result.ok) {
      return unsupported(
        id,
        'unsupported',
        `The engine could not prepare the ${step.operation} step: ${result.error.messages.join(' ')}`,
      );
    }
    if (result.preview.affectedCells === 0) {
      const warning = result.preview.warnings[0]?.message;
      noOpReasons.push(
        `${step.operation} found no cells to change${warning ? ` (${warning})` : ''}.`,
      );
      continue;
    }
    simulated = result.workbook;
    planSteps.push({
      id: `template-${id}-${index + 1}`,
      operation: step.operation,
      args: step.args,
      description: step.description,
      category: 'transform',
      status: 'pending',
      preview: result.preview,
    });
  }

  if (planSteps.length === 0) {
    return unsupported(
      id,
      'no-op',
      noOpReasons.length > 0
        ? noOpReasons.join(' ')
        : 'The workbook already satisfies this workflow; nothing needs review.',
    );
  }

  const completeDescription = [description, ...assumptions, ...noOpReasons].join(' ');
  return {
    ok: true,
    template: template(id),
    assumptions,
    plan: {
      id: `template-plan-${id}`,
      title,
      description: completeDescription,
      assumptions,
      steps: planSteps,
      status: 'pending',
      totalAffectedCells: planSteps.reduce(
        (total, planStep) => total + (planStep.preview?.affectedCells ?? 0),
        0,
      ),
    },
  };
}

function weeklyExport(
  workbook: Workbook,
  activeSheetName: string,
  registry: OperationRegistry,
): WorkflowTemplateResult {
  const sheet = activeSheet(workbook, activeSheetName);
  if (!sheet)
    return unsupported('clean_weekly_export', 'unsupported', 'The workbook has no worksheet.');
  if (!hasTabularData(sheet))
    return unsupported(
      'clean_weekly_export',
      'unsupported',
      `"${sheet.name}" needs a header row and at least one tabular data row. Code, logs, and empty sheets need a confirmed table shape first.`,
    );

  const targetSheet = safeSheetName(workbook, 'Weekly Clean Export');
  return buildVerifiedPlan(
    workbook,
    registry,
    'clean_weekly_export',
    `Clean weekly export — ${sheet.name}`,
    `Prepare a reviewable cleaned copy of "${sheet.name}" as "${targetSheet}".`,
    [
      ...template('clean_weekly_export').assumptions,
      `The output name is "${targetSheet}"; existing sheets are never overwritten.`,
    ],
    [
      {
        operation: 'clean_to_new_sheet',
        args: {
          sheet: sheet.name,
          targetSheet,
          trim: true,
          collapseWhitespace: true,
          dropEmptyRows: true,
          dropEmptyColumns: true,
          coerceNumbers: true,
        },
        description: `Create "${targetSheet}" with trimmed, structured values from "${sheet.name}".`,
      },
    ],
  );
}

function dateColumns(sheet: Sheet): ColumnMetadata[] {
  return getColumnProfiles(sheet).filter(
    (column) =>
      column.nonBlankCount > 0 &&
      (column.isDate || /date|time|timestamp|day|created|updated/i.test(column.rawName)),
  );
}

function normalizeDates(
  workbook: Workbook,
  activeSheetName: string,
  registry: OperationRegistry,
): WorkflowTemplateResult {
  const sheet = activeSheet(workbook, activeSheetName);
  if (!sheet || !hasTabularData(sheet))
    return unsupported(
      'normalize_dates',
      'unsupported',
      'Date normalization needs a tabular worksheet with a header row and data rows.',
    );
  const candidates = dateColumns(sheet);
  if (candidates.length === 0)
    return unsupported(
      'normalize_dates',
      'unsupported',
      `No date-like column was detected in "${sheet.name}". Specify a date column explicitly before using this template.`,
    );

  const steps = candidates.map((column) => ({
    operation: 'format_dates',
    args: {
      sheet: sheet.name,
      column: column.letter,
      format: 'YYYY-MM-DD',
      headerRow: 1,
    },
    description: `Normalize "${column.rawName}" (${column.letter}) to YYYY-MM-DD.`,
  }));
  return buildVerifiedPlan(
    workbook,
    registry,
    'normalize_dates',
    `Normalize dates — ${sheet.name}`,
    `Standardize ${candidates.length} detected date column${candidates.length === 1 ? '' : 's'} in "${sheet.name}".`,
    template('normalize_dates').assumptions,
    steps,
  );
}

function normalizedHeader(value: unknown): string {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function duplicateKeyColumns(sheet: Sheet): { columns: string[]; explanation: string } {
  const width = maxColumnCount(sheet.rows);
  const headers = Array.from({ length: width }, (_, index) =>
    normalizedHeader(sheet.rows[0]?.[index]?.value),
  );
  const identifierIndexes = headers.flatMap((header, index) =>
    /^(?:id|key|sku|code|reference|ref)$/.test(header) ||
    /(?:order|invoice|transaction|record|account|customer|product|item).*(?:id|number|no|code|reference|ref)$/.test(
      header,
    )
      ? [index]
      : [],
  );
  if (identifierIndexes.length === 1) {
    const index = identifierIndexes[0]!;
    return {
      columns: [indexToColumn(index)],
      explanation: `Use the unambiguous identifier column ${indexToColumn(index)} as the duplicate key.`,
    };
  }
  const columns = Array.from({ length: width }, (_, index) => indexToColumn(index));
  return {
    columns,
    explanation:
      identifierIndexes.length > 1
        ? 'Several identifier-like columns were detected, so the complete row is used instead of guessing a key.'
        : 'No unambiguous identifier column was detected, so the complete row is used as the key.',
  };
}

function removeDuplicates(
  workbook: Workbook,
  activeSheetName: string,
  registry: OperationRegistry,
): WorkflowTemplateResult {
  const sheet = activeSheet(workbook, activeSheetName);
  if (!sheet || !hasTabularData(sheet))
    return unsupported(
      'remove_duplicates',
      'unsupported',
      'Duplicate removal needs a tabular worksheet with a header row and data rows.',
    );
  const key = duplicateKeyColumns(sheet);
  return buildVerifiedPlan(
    workbook,
    registry,
    'remove_duplicates',
    `Remove duplicates — ${sheet.name}`,
    `Review repeated records in "${sheet.name}" and keep the first occurrence.`,
    [...template('remove_duplicates').assumptions, key.explanation],
    [
      {
        operation: 'delete_duplicates',
        args: {
          sheet: sheet.name,
          columns: key.columns,
          headerRow: 1,
          keep: 'first',
        },
        description: `Remove duplicate rows in "${sheet.name}" using ${key.columns.join(', ')}.`,
      },
    ],
  );
}

function reconcileSheets(workbook: Workbook, registry: OperationRegistry): WorkflowTemplateResult {
  if (workbook.sheets.length < 2)
    return unsupported(
      'reconcile_sheets',
      'unsupported',
      'Reconciliation needs two source worksheets in the same workbook; the source sheet will not be modified.',
    );

  const intent = analyzeReconciliationIntent('reconcile these sheets', workbook, {
    allowComplex: true,
  });
  if (!intent?.proposedAction) {
    return unsupported(
      'reconcile_sheets',
      intent?.clarification ? 'needs-input' : 'unsupported',
      intent?.clarification?.question ??
        intent?.message ??
        'The two source sheets do not have enough unambiguous bindings for a safe comparison.',
    );
  }
  const action = intent.proposedAction;
  const leftSheet = String(action.args.leftSheet ?? 'left source');
  const rightSheet = String(action.args.rightSheet ?? 'right source');
  const amountNote = action.args.leftAmount
    ? `Compare amounts ${String(action.args.leftAmount)} ↔ ${String(action.args.rightAmount)} with absolute tolerance 0.`
    : 'Compare keys only because no unambiguous amount pair was detected.';
  return buildVerifiedPlan(
    workbook,
    registry,
    'reconcile_sheets',
    `Reconcile & compare — ${leftSheet} ↔ ${rightSheet}`,
    `Prepare a source-preserving comparison report for "${leftSheet}" and "${rightSheet}".`,
    [...template('reconcile_sheets').assumptions, amountNote],
    [
      {
        operation: 'reconcile_sheets',
        args: action.args,
        description: `Compare ${leftSheet} and ${rightSheet}; append matched and exception reports without rewriting either source.`,
      },
    ],
  );
}

export function planWorkflowTemplate(
  id: WorkflowTemplateId,
  workbook: Workbook,
  activeSheetName: string,
  registry: OperationRegistry = createOperationRegistry(),
): WorkflowTemplateResult {
  switch (id) {
    case 'clean_weekly_export':
      return weeklyExport(workbook, activeSheetName, registry);
    case 'normalize_dates':
      return normalizeDates(workbook, activeSheetName, registry);
    case 'remove_duplicates':
      return removeDuplicates(workbook, activeSheetName, registry);
    case 'reconcile_sheets':
      return reconcileSheets(workbook, registry);
  }
}

export function discoverWorkflowTemplates(
  workbook: Workbook,
  activeSheetName: string,
  registry: OperationRegistry = createOperationRegistry(),
): WorkflowTemplateDiscovery[] {
  return WORKFLOW_TEMPLATE_IDS.map((id) => {
    const result = planWorkflowTemplate(id, workbook, activeSheetName, registry);
    return { template: result.template, result };
  });
}

export function workflowActionFromPlan(
  discovery: Extract<WorkflowTemplateResult, { ok: true }>,
): ProposedAction {
  const first = discovery.plan.steps[0];
  if (!first) throw new Error('A workflow plan must contain at least one step.');
  return {
    name: first.operation,
    args: first.args,
    explanation: first.description,
    category: first.category ?? 'transform',
  };
}
