import { describe, expect, it } from 'vitest';

import { createOperationRegistry } from '@excel-agent/engine';
import { createMemoryStore, createOrchestrator } from '@excel-agent/agent';

import {
  EVAL_SHEET,
  createOrchestratorResolver,
  evalWorkbook,
  formatSummary,
  runEvals,
} from '../src/index.js';

describe('deterministic planner evals', () => {
  it('meets the baseline accuracy and never proposes an unguardrailed action', async () => {
    const orchestrator = createOrchestrator({ registry: createOperationRegistry() });
    const summary = await runEvals(createOrchestratorResolver(orchestrator));

    console.log(`\n${formatSummary(summary)}`);
    expect(summary.actionPrecision).toBe(1);
    expect(summary.accuracy).toBeGreaterThanOrEqual(0.9);
  });

  it('resolves a paraphrase through self-learning memory after verified successes', async () => {
    const registry = createOperationRegistry();
    const memory = createMemoryStore();
    const orchestrator = createOrchestrator({ registry, memory });
    const resolver = createOrchestratorResolver(orchestrator, {
      workbook: evalWorkbook(),
      sheetName: EVAL_SHEET,
    });

    const first = await resolver('remove duplicate rows');
    expect(first.operation).toBe('delete_duplicates');
    expect(first.guardrailPassed).toBe(true);

    // Simulate two successful applications confirming the learned association.
    for (let round = 0; round < 2; round += 1) {
      orchestrator.learn({
        query: 'remove duplicate rows',
        sheetName: EVAL_SHEET,
        operation: 'delete_duplicates',
        args: first.args,
        success: true,
      });
    }

    const second = await resolver('please remove duplicate rows');
    expect(second.source).toBe('memory');
    expect(second.operation).toBe('delete_duplicates');
  });
});
