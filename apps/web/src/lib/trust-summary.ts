import type { ExecutionPlan, GuardrailReport, ProposedAction } from '@excel-agent/agent';
import type { Preview } from '@excel-agent/engine';

/**
 * The small, durable subset of a guardrail report that is useful in mission history.
 * Preview cell values are intentionally not copied here: the preview itself remains attached to
 * the in-tab chat card, while this summary keeps the audit record bounded and readable.
 */
export interface MissionSafetySummary {
  passed: boolean;
  schemaValid: boolean;
  engineValid: boolean;
  requiresConfirmation: boolean;
  affectedCells: number;
  errors: string[];
  warnings: string[];
}

export function summarizeGuardrail(guardrail: GuardrailReport): MissionSafetySummary {
  return {
    passed: guardrail.passed,
    schemaValid: guardrail.schemaValid,
    engineValid: guardrail.engineValid,
    requiresConfirmation: guardrail.requiresConfirmation,
    affectedCells: guardrail.preview?.affectedCells ?? 0,
    errors: guardrail.errors.slice(0, 12),
    warnings: guardrail.warnings.slice(0, 12),
  };
}

export function planAffectedCells(plan: ExecutionPlan): number {
  return (
    plan.totalAffectedCells ??
    plan.steps.reduce((total, step) => total + (step.preview?.affectedCells ?? 0), 0)
  );
}

export function planRequiresConfirmation(plan: ExecutionPlan): boolean {
  return plan.steps.some((step) => step.preview?.requiresConfirmation === true);
}

function unique(values: string[]): string[] {
  return [...new Set(values.filter((value) => value.trim().length > 0))];
}

function stringArg(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

/** Make the operation arguments legible without dumping raw JSON into the review card. */
export function planAssumptions(plan: ExecutionPlan): string[] {
  const derived: string[] = [];
  plan.steps.forEach((step, index) => {
    const args = step.args;
    const stepLabel = `Step ${index + 1}`;
    const sheet = stringArg(args, 'sheet');
    const targetSheet = stringArg(args, 'targetSheet');
    const leftSheet = stringArg(args, 'leftSheet');
    const rightSheet = stringArg(args, 'rightSheet');
    const headerRow = args.headerRow;
    const columns = Array.isArray(args.columns)
      ? args.columns.filter((column): column is string => typeof column === 'string')
      : [];

    if (leftSheet && rightSheet) {
      derived.push(`${stepLabel} compares “${leftSheet}” with “${rightSheet}”.`);
    } else if (sheet) {
      derived.push(`${stepLabel} targets “${sheet}”.`);
    }
    if (targetSheet) derived.push(`${stepLabel} writes to “${targetSheet}”.`);
    if (typeof headerRow === 'number') {
      derived.push(
        `${stepLabel} treats row ${headerRow} as the header; data begins below that row.`,
      );
    }
    if (columns.length > 0) {
      derived.push(`${stepLabel} uses the explicit column scope: ${columns.join(', ')}.`);
    }
  });

  return unique([...(plan.assumptions ?? []), ...derived]).slice(0, 12);
}

/**
 * These are statements about checks that are actually enforced by the engine's plan path, not
 * claims about the quality or business meaning of the workbook data.
 */
export function planSafeguards(plan: ExecutionPlan): string[] {
  const previewed = plan.steps.filter((step) => step.preview !== undefined).length;
  const failed = plan.steps.filter((step) => step.status === 'error').length;
  const safeguards = [
    previewed === plan.steps.length
      ? `${plan.steps.length} step${plan.steps.length === 1 ? '' : 's'} has a deterministic engine preview.`
      : `Every step is revalidated by the deterministic engine before commit (${previewed}/${plan.steps.length} preview${previewed === 1 ? '' : 's'} attached now).`,
    'Execution is staged as one plan; a failed step leaves the workbook unchanged.',
    'A successful run returns a forward patch and inverse patch for reversible history.',
  ];
  if (planRequiresConfirmation(plan)) {
    safeguards.push(
      'A destructive or wide-reaching step requires a second confirmation before commit.',
    );
  } else {
    safeguards.push('The workbook stays unchanged until you approve the plan.');
  }
  if (failed > 0)
    safeguards.push(`${failed} step${failed === 1 ? '' : 's'} is blocked by engine validation.`);
  return safeguards;
}

export function singleActionSafeguards(
  guardrail: GuardrailReport | undefined,
  preview: Preview | undefined,
): string[] {
  if (!guardrail) {
    return [
      'The registered operation is previewed before it can be applied.',
      'The workbook is committed transactionally and recorded in reversible history.',
      'Destructive or wide-reaching changes require an explicit confirmation.',
    ];
  }
  return [
    `Schema validation ${guardrail.schemaValid ? 'passed' : 'failed'}.`,
    `Engine validation ${guardrail.engineValid ? 'passed' : 'failed'}.`,
    `Deterministic preview covers ${preview?.affectedCells ?? guardrail.preview?.affectedCells ?? 0} affected cell${(preview?.affectedCells ?? guardrail.preview?.affectedCells ?? 0) === 1 ? '' : 's'}.`,
    guardrail.requiresConfirmation
      ? 'The engine requires a second confirmation before this change can be committed.'
      : 'The workbook remains unchanged until you approve the proposal.',
    'A successful run returns a forward patch and inverse patch for reversible history.',
  ];
}

export function actionScope(action: ProposedAction): string[] {
  const args = action.args;
  const scope: string[] = [];
  const sheet = stringArg(args, 'sheet');
  const targetSheet = stringArg(args, 'targetSheet');
  const leftSheet = stringArg(args, 'leftSheet');
  const rightSheet = stringArg(args, 'rightSheet');
  if (sheet) scope.push(`Source sheet: ${sheet}`);
  if (targetSheet) scope.push(`Target sheet: ${targetSheet}`);
  if (leftSheet && rightSheet) scope.push(`Sources: ${leftSheet} ↔ ${rightSheet}`);
  if (typeof args.headerRow === 'number') scope.push(`Header row: ${args.headerRow}`);
  if (typeof args.column === 'string') scope.push(`Column: ${args.column}`);
  if (Array.isArray(args.columns)) {
    const columns = args.columns.filter((column): column is string => typeof column === 'string');
    if (columns.length > 0) scope.push(`Columns: ${columns.join(', ')}`);
  }
  return scope;
}
