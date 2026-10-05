import { describe, expect, it } from 'vitest';

import { createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';

import { createMemoryStore, createOrchestrator, sheetFingerprint } from '../src/index.js';

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
      // The learned arguments are column letters, so they are only valid against the header
      // layout they were learned on.
      schemaFingerprint: sheetFingerprint(ordersWorkbook(), 'Orders'),
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

  it('refuses to replay a learned action after the sheet shape has changed', async () => {
    const memory = createMemoryStore();
    memory.remember({
      key: 'sort by revenue',
      rawQuery: 'sort by revenue',
      operation: 'sort_range',
      // "Revenue" is column C in this layout.
      args: { sheet: 'Orders', column: 'C', direction: 'desc', startRow: 2 },
      sheetName: 'Orders',
      schemaFingerprint: sheetFingerprint(ordersWorkbook(), 'Orders'),
    });
    for (let i = 0; i < 4; i += 1) memory.recordOutcome('sort_range', 'Orders', true);

    // A column is deleted, so every letter from D onward shifts left and column C is no longer
    // revenue. Replaying the frozen args would sort by the wrong field and pass every check.
    const reshaped = ordersWorkbook();
    reshaped.sheets[0]!.rows[0] = [createCell('Order ID'), createCell('Customer')];
    for (const row of reshaped.sheets[0]!.rows) row.splice(2, 1);

    const orchestrator = createOrchestrator({ registry, memory });
    const decision = await orchestrator.decide({
      query: 'sort by revenue',
      workbook: reshaped,
      sheetName: 'Orders',
    });

    // The stale record must not be replayed. Whatever happens next must be freshly derived.
    expect(decision.source).not.toBe('memory');
    expect(
      decision.trace.some((step) => step.layer === 'memory' && step.summary.includes('sort_range')),
    ).toBe(false);
  });

  it('proposes clean_to_new_sheet when asked to make sheet clean and structured on an unstructured code sheet', async () => {
    const orchestrator = createOrchestrator({ registry });
    const codeWorkbook: Workbook = {
      sheets: [
        {
          name: 'Worksheet',
          rows: [
            [createCell('#!/usr/bin/env python3')],
            [createCell('import os, sys')],
            [createCell('def main(): pass')],
          ],
        },
        {
          name: 'Zenith_Leads_Template',
          rows: [[createCell('Name'), createCell('Company'), createCell('Email')]],
        },
      ],
    };

    const decision = await orchestrator.decide({
      query: 'can you make this sheet clean and structured so that it can be understandable',
      workbook: codeWorkbook,
      sheetName: 'Worksheet',
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('clean_to_new_sheet');
    expect(decision.message).not.toContain('Here are common things I can do for you:');
  });

  it('resolves "perform then" conversational followup using history and proposes action', async () => {
    const orchestrator = createOrchestrator({ registry });
    const codeWorkbook: Workbook = {
      sheets: [
        {
          name: 'Worksheet',
          rows: [
            [createCell('#!/usr/bin/env python3')],
            [createCell('import os, sys')],
            [createCell('def main(): pass')],
          ],
        },
      ],
    };

    const decision = await orchestrator.decide({
      query: 'perform then',
      workbook: codeWorkbook,
      sheetName: 'Worksheet',
      history: [
        { role: 'user', content: 'can you make this sheet clean and structured so that it can be understandable' },
        { role: 'assistant', content: 'I am ready to perform...' },
      ],
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('clean_to_new_sheet');
    expect(decision.message).not.toContain('Here are common things I can do for you:');
  });

  it('exposes a tool catalog derived from the engine registry', () => {
    const orchestrator = createOrchestrator({ registry });
    expect(orchestrator.tools.map((tool) => tool.name)).toEqual(registry.names);
  });
});
