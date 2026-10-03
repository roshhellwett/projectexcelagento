import type { ZodType } from 'zod';

import { formatDatesOperation } from './format-dates.js';
import { initialOperations } from './operations.js';
import { invariantNoCellsOutsideTargetRange } from './invariants.js';
import type { HistoryStack } from './history.js';
import type { Operation, OperationResult, Preview, Workbook } from './types.js';
import { applyPatch, cloneWorkbook, workbookEquals } from './workbook.js';

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
  confirmed?: boolean;
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
