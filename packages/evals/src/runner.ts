import type { Workbook } from '@excel-agent/engine';

import { evalCases, type EvalCase } from './cases.js';
import { EVAL_SHEET, evalWorkbook } from './fixtures.js';

export interface ResolvedAction {
  operation: string | null;
  args: Record<string, unknown>;
  guardrailPassed: boolean;
  source: string;
}

export interface OrchestratorLike {
  decide(input: { query: string; workbook: Workbook; sheetName: string }): Promise<{
    action?: { name: string; args: Record<string, unknown> };
    guardrail?: { passed: boolean };
    source: string;
  }>;
}

/**
 * Adapt an orchestrator into an eval resolver against the canonical fixture.
 * Kept separate so evals stay decoupled from any single orchestration shape.
 */
export function createOrchestratorResolver(
  orchestrator: OrchestratorLike,
  options: { workbook?: Workbook; sheetName?: string } = {},
): (prompt: string) => Promise<ResolvedAction> {
  const workbook = options.workbook ?? evalWorkbook();
  const sheetName = options.sheetName ?? EVAL_SHEET;
  return async (prompt) => {
    const decision = await orchestrator.decide({ query: prompt, workbook, sheetName });
    return {
      operation: decision.action?.name ?? null,
      args: decision.action?.args ?? {},
      guardrailPassed: decision.guardrail?.passed ?? decision.action === undefined,
      source: decision.source,
    };
  };
}

export interface EvalResult {
  id: string;
  prompt: string;
  expected: string | null;
  actual: string | null;
  passed: boolean;
  guardrailPassed: boolean;
  source: string;
  reason: string;
}

export interface EvalSummary {
  total: number;
  passed: number;
  failed: number;
  accuracy: number;
  /** Fraction of proposed actions that passed the guardrail. */
  actionPrecision: number;
  results: EvalResult[];
}

function matchesArgs(
  expected: Record<string, unknown> | undefined,
  actual: Record<string, unknown>,
): { ok: boolean; detail: string } {
  if (!expected) return { ok: true, detail: '' };
  for (const [key, value] of Object.entries(expected)) {
    if (JSON.stringify(actual[key]) !== JSON.stringify(value)) {
      return {
        ok: false,
        detail: `arg ${key}: expected ${JSON.stringify(value)}, got ${JSON.stringify(actual[key])}`,
      };
    }
  }
  return { ok: true, detail: '' };
}

export async function runEvals(
  resolve: (prompt: string) => Promise<ResolvedAction>,
  cases: EvalCase[] = evalCases,
): Promise<EvalSummary> {
  const results: EvalResult[] = [];
  let proposed = 0;
  let proposedValid = 0;

  for (const testCase of cases) {
    const resolved = await resolve(testCase.prompt);
    if (resolved.operation) {
      proposed += 1;
      if (resolved.guardrailPassed) proposedValid += 1;
    }

    const operationOk = resolved.operation === testCase.expected.operation;
    const args = operationOk
      ? matchesArgs(testCase.expected.args, resolved.args)
      : { ok: false, detail: '' };
    const passed = operationOk && args.ok;

    results.push({
      id: testCase.id,
      prompt: testCase.prompt,
      expected: testCase.expected.operation,
      actual: resolved.operation,
      passed,
      guardrailPassed: resolved.guardrailPassed,
      source: resolved.source,
      reason: !operationOk
        ? `expected ${testCase.expected.operation ?? 'no action'}, got ${resolved.operation ?? 'no action'}`
        : args.detail,
    });
  }

  const passed = results.filter((result) => result.passed).length;
  return {
    total: results.length,
    passed,
    failed: results.length - passed,
    accuracy: results.length === 0 ? 0 : passed / results.length,
    actionPrecision: proposed === 0 ? 1 : proposedValid / proposed,
    results,
  };
}

export function formatSummary(summary: EvalSummary): string {
  const lines = [
    `Evals: ${summary.passed}/${summary.total} passed (accuracy ${(summary.accuracy * 100).toFixed(1)}%)`,
    `Action precision (guardrail-passing proposals): ${(summary.actionPrecision * 100).toFixed(1)}%`,
  ];
  for (const result of summary.results) {
    lines.push(`  ${result.passed ? 'PASS' : 'FAIL'} ${result.id} - ${result.reason || 'ok'}`);
  }
  return lines.join('\n');
}
