import { describe, expect, it } from 'vitest';

import { createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';

import { createMemoryStore, createOrchestrator } from '../src/index.js';

function ordersWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Orders',
        rows: [
          [createCell('Order ID'), createCell('Customer'), createCell('Amount')],
          [createCell('ORD-1'), createCell('acme corp'), createCell(10)],
          [createCell('ORD-2'), createCell('bmc ltd'), createCell(20)],
          [createCell('ORD-1'), createCell('acme corp'), createCell(10)],
        ],
      },
    ],
  };
}

describe('ExcelAgentOrchestrator', () => {
  const registry = createOperationRegistry();

  it('answers a greeting without proposing an operation', async () => {
    const orchestrator = createOrchestrator({ registry });
    const decision = await orchestrator.decide({
      query: 'hello there',
      workbook: ordersWorkbook(),
      sheetName: 'Orders',
    });

    expect(decision.action).toBeUndefined();
    expect(decision.trace[0]?.layer).toBe('intent');
  });

  it('proposes a deterministic dedupe action that passes the guardrail', async () => {
    const orchestrator = createOrchestrator({ registry });
    const decision = await orchestrator.decide({
      query: 'remove duplicate rows',
      workbook: ordersWorkbook(),
      sheetName: 'Orders',
    });

    expect(decision.action?.name).toBe('delete_duplicates');
    expect(decision.guardrail?.passed).toBe(true);
    expect(decision.source).toBe('heuristic');
    expect(decision.trace.some((step) => step.layer === 'guardrail')).toBe(true);
  });

  it('resolves an explicitly referenced column for sorting', async () => {
    const orchestrator = createOrchestrator({ registry });
    const decision = await orchestrator.decide({
      query: 'sort by column A descending',
      workbook: ordersWorkbook(),
      sheetName: 'Orders',
    });

    expect(decision.action?.name).toBe('sort_range');
    expect(decision.action?.args.column).toBe('A');
    expect(decision.guardrail?.passed).toBe(true);
  });

  it('falls back to conversation when no operation is recognised', async () => {
    const orchestrator = createOrchestrator({ registry });
    const decision = await orchestrator.decide({
      query: 'zzz qqq nothing meaningful',
      workbook: ordersWorkbook(),
      sheetName: 'Orders',
    });

    expect(decision.action).toBeUndefined();
    expect(decision.source).toBe('fallback');
  });

  it('blocks an unknown operation in the guardrail', () => {
    const orchestrator = createOrchestrator({ registry });
    const report = orchestrator.guardrail(ordersWorkbook(), {
      name: 'not_a_real_operation',
      args: {},
      explanation: 'nope',
      category: 'transform',
    });

    expect(report.passed).toBe(false);
    expect(report.schemaValid).toBe(false);
    expect(report.errors[0]).toContain('Unknown operation');
  });

  it('blocks an action whose args fail schema validation', () => {
    const orchestrator = createOrchestrator({ registry });
    const report = orchestrator.guardrail(ordersWorkbook(), {
      name: 'format_dates',
      args: { sheet: 'Orders', column: 'C', format: 'NOT-A-FORMAT' },
      explanation: 'bad format',
      category: 'format',
    });

    expect(report.passed).toBe(false);
    expect(report.schemaValid).toBe(false);
    expect(report.errors.length).toBeGreaterThan(0);
  });

  it('replays a verified memory hit without touching the model', async () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'remove duplicate rows',
      rawQuery: 'remove duplicate rows',
      operation: 'delete_duplicates',
      args: { sheet: 'Orders', columns: ['A', 'B', 'C'], headerRow: 1, keep: 'first' },
      sheetName: 'Orders',
    });
    memory.recordOutcome('delete_duplicates', 'Orders', true);
    memory.recordOutcome('delete_duplicates', 'Orders', true);

    const orchestrator = createOrchestrator({ registry, memory });
    const decision = await orchestrator.decide({
      query: 'please remove the duplicate rows',
      workbook: ordersWorkbook(),
      sheetName: 'Orders',
    });

    expect(decision.source).toBe('memory');
    expect(decision.action?.name).toBe('delete_duplicates');
    expect(decision.guardrail?.passed).toBe(true);
  });

  it('exposes a tool catalog derived from the engine registry', () => {
    const orchestrator = createOrchestrator({ registry });
    expect(orchestrator.tools.map((tool) => tool.name)).toEqual(registry.names);
  });
});
