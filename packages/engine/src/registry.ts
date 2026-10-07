import type { ZodType } from 'zod';

import { formatDatesOperation } from './format-dates.js';
import { analyticsOperations } from './analytics-operations.js';
import { initialOperations } from './operations.js';
import { advancedOperations } from './advanced-operations.js';
import { reconcileSheetsOperation } from './reconciliation-operation.js';
import { formatCellsOperation } from './style-operation.js';
import { invariantNoCellsOutsideTargetRange } from './invariants.js';
import type { HistoryStack } from './history.js';
import type { Operation, OperationResult, Preview, Workbook } from './types.js';
import {
  applyPatch,
  cloneWorkbook,
  invertPatch,
  patchBetween,
  workbookEquals,
} from './workbook.js';

export class OperationRegistry {
  private readonly operations = new Map<string, Operation<unknown>>();

  register<Args>(operation: Operation<Args>): this {
    if (this.operations.has(operation.name)) {
      throw new Error(`Operation "${operation.name}" is already registered.`);
    }
    this.operations.set(operation.name, operation as Operation<unknown>);
    return this;
  }

  get(name: string): Operation<unknown> | undefined {
    return this.operations.get(name);
  }

  getSchema(name: string): ZodType | undefined {
    return this.operations.get(name)?.schema;
  }

  applyOperation(
    workbook: Workbook,
    name: string,
    input: unknown,
    options: Omit<ApplyOperationOptions, 'registry'> = {},
  ): ApplyOperationResult {
    return applyOperation(workbook, name, input, { ...options, registry: this });
  }

  get schemas(): ReadonlyMap<string, ZodType> {
    return new Map([...this.operations].map(([name, operation]) => [name, operation.schema]));
  }

  get names(): string[] {
    return [...this.operations.keys()];
  }
}

export function createOperationRegistry(): OperationRegistry {
  const registry = new OperationRegistry().register(formatDatesOperation);
  for (const operation of initialOperations) registry.register(operation);
  for (const operation of advancedOperations) registry.register(operation);
  for (const operation of analyticsOperations) registry.register(operation);
  registry.register(reconcileSheetsOperation);
  registry.register(formatCellsOperation);
  return registry;
}

export type OperationErrorCode =
  | 'unknown-operation'
  | 'schema-error'
  | 'validation-error'
  | 'preview-error'
  | 'confirmation-required'
  | 'execution-error'
  | 'invariant-error';

export type ApplyOperationResult =
  | ({ ok: true; preview: Preview } & OperationResult)
  | {
      ok: false;
      workbook: Workbook;
      error: { code: OperationErrorCode; messages: string[]; rolledBack: boolean };
      preview?: Preview;
    };

export interface ApplyOperationOptions {
  registry?: OperationRegistry;
  history?: HistoryStack;
  /**
   * Required when the operation's preview declares `requiresConfirmation`. Until a caller sets
   * this, such an operation is refused rather than applied - which is what makes the flag a
   * guarantee instead of a suggestion.
   */
  confirmed?: boolean;
}

export interface OperationPlanStep {
  operation: string;
  args: unknown;
}

export type ApplyOperationPlanResult =
  | { ok: true; workbook: Workbook; steps: Extract<ApplyOperationResult, { ok: true }>[] }
  | (Extract<ApplyOperationResult, { ok: false }> & { failedStep: number });

/**
 * Stage and verify every step against the preceding result, then commit once. A refused plan
 * cannot discard an existing redo branch, leak partial steps into history, or trigger history
 * compaction. One undo restores the entire plan.
 */
export function applyOperationPlan(
  workbook: Workbook,
  steps: OperationPlanStep[],
  options: ApplyOperationOptions & {
    operationName?: string;
    onStep?: (index: number, phase: 'verifying' | 'verified' | 'failed') => void;
  } = {},
): ApplyOperationPlanResult {
  if (steps.length === 0) {
    return {
      ok: false,
      workbook: cloneWorkbook(workbook),
      failedStep: 0,
      error: {
        code: 'validation-error',
        messages: ['A plan must contain at least one step.'],
        rolledBack: false,
      },
    };
  }
  let current = cloneWorkbook(workbook);
  const results: Extract<ApplyOperationResult, { ok: true }>[] = [];
  let confirmationStep = -1;
  for (const [index, step] of steps.entries()) {
    options.onStep?.(index, 'verifying');
    // Confirmation here only authorizes the private simulation; no caller-visible state is committed.
    const result = applyOperation(current, step.operation, step.args, {
      registry: options.registry,
      confirmed: true,
    });
    options.onStep?.(index, result.ok ? 'verified' : 'failed');
    if (!result.ok) return { ...result, workbook: cloneWorkbook(workbook), failedStep: index };
    if (result.preview.requiresConfirmation && confirmationStep === -1) confirmationStep = index;
    current = result.workbook;
    results.push(result);
  }
  if (confirmationStep !== -1 && !options.confirmed) {
    const previews = results.map((result) => result.preview);
    return {
      ok: false,
      workbook: cloneWorkbook(workbook),
      failedStep: confirmationStep,
      error: {
        code: 'confirmation-required',
        messages: ['This plan requires confirmation. Nothing was applied.'],
        rolledBack: false,
      },
      preview: {
        valid: true,
        affectedCells: previews.reduce((count, preview) => count + preview.affectedCells, 0),
        changes: previews.flatMap((preview) => preview.changes).slice(0, 20),
        warnings: previews.flatMap((preview) => preview.warnings),
        errors: [],
        requiresConfirmation: true,
      },
    };
  }
  const patch = patchBetween(workbook, current);
  if (patch.length > 0) {
    options.history?.commit(options.operationName ?? 'plan', current, patch, invertPatch(patch));
  }
  return { ok: true, workbook: current, steps: results };
}

/** Single transactional entry point. Failures never modify the caller's workbook. */
export function applyOperation(
  workbook: Workbook,
  name: string,
  input: unknown,
  options: ApplyOperationOptions = {},
): ApplyOperationResult {
  const before = cloneWorkbook(workbook);
  let preview: Preview | undefined;
  const failure = (
    code: OperationErrorCode,
    messages: string[],
    rolledBack = false,
  ): ApplyOperationResult => ({
    ok: false,
    workbook: cloneWorkbook(before),
    error: { code, messages, rolledBack },
    ...(preview ? { preview } : {}),
  });
  const operation = (options.registry ?? createOperationRegistry()).get(name);
  if (!operation) return failure('unknown-operation', [`Unknown operation: ${name}`]);
  const parsed = operation.schema.safeParse(input);
  if (!parsed.success) {
    return failure(
      'schema-error',
      parsed.error.issues.map((issue) => issue.message),
    );
  }

  try {
    const args = parsed.data;
    const validation = operation.validate(cloneWorkbook(before), args);
    if (!validation.valid) {
      return failure(
        'validation-error',
        validation.errors.map((issue) => issue.message),
      );
    }
    preview = operation.preview(cloneWorkbook(before), args);
    if (!preview.valid) {
      return failure(
        'preview-error',
        preview.errors.map((issue) => issue.message),
      );
    }

    // A destructive or wide-reaching operation stops here and asks to be confirmed. This is the
    // hard wall: an operation that declares it needs confirmation can never be applied by
    // accident, because the only way through is a caller that explicitly confirms.
    if (preview.requiresConfirmation && !options.confirmed) {
      const message =
        preview.affectedCells === 1
          ? 'This change affects 1 cell and was not confirmed.'
          : `This change affects ${preview.affectedCells} cells and was not confirmed.`;
      return failure('confirmation-required', [
        message,
        ...preview.warnings.map((issue) => issue.message),
      ]);
    }

    const result = operation.apply(cloneWorkbook(before), args);
    const check = operation.invariants(cloneWorkbook(before), cloneWorkbook(result.workbook), args);
    const errors = [
      ...check.errors,
      ...invariantNoCellsOutsideTargetRange(
        before,
        result.workbook,
        operation.targetRanges(cloneWorkbook(before), args),
      ),
    ];
    if (!check.valid && errors.length === 0) errors.push('An operation invariant failed.');
    try {
      if (!workbookEquals(applyPatch(before, result.patch), result.workbook)) {
        errors.push('The forward patch does not reproduce the result.');
      }
      if (!workbookEquals(applyPatch(result.workbook, result.inverse), before)) {
        errors.push('The inverse patch does not restore the original workbook.');
      }
    } catch {
      errors.push('The operation returned an invalid patch or inverse.');
    }
    if (errors.length > 0) {
      // Try the inverse first. The immutable pre-operation snapshot remains the
      // authoritative fallback even if a faulty operation supplied a bad inverse.
      try {
        applyPatch(result.workbook, result.inverse);
      } catch {
        // Snapshot rollback still succeeds without changing history.
      }
      return failure('invariant-error', [...new Set(errors)], true);
    }
    options.history?.commit(name, result.workbook, result.patch, result.inverse);
    return { ok: true, ...result, preview };
  } catch (error) {
    return failure('execution-error', [
      error instanceof Error ? error.message : 'Operation failed.',
    ]);
  }
}
