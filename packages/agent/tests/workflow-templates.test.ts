import { describe, expect, it } from 'vitest';

import {
  applyOperationPlan,
  createCell,
  createOperationRegistry,
  type Workbook,
} from '@excel-agent/engine';

import {
  discoverWorkflowTemplates,
  planWorkflowTemplate,
  WORKFLOW_TEMPLATE_IDS,
} from '../src/index.js';

function workbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Orders',
        rows: [
          [createCell('Order ID'), createCell('Order Date'), createCell('Amount')],
          [createCell('ORD-1'), createCell('2026-01-02'), createCell(100)],
          [createCell('ORD-2'), createCell('2026-01-03'), createCell(200)],
          [createCell('ORD-1'), createCell('2026-01-02'), createCell(100)],
        ],
      },
      {
        name: 'Invoices',
        rows: [
          [createCell('Order ID'), createCell('Invoice Date'), createCell('Amount')],
          [createCell('ORD-1'), createCell('2026-01-02'), createCell(100)],
          [createCell('ORD-2'), createCell('2026-01-03'), createCell(210)],
        ],
      },
    ],
  };
}

describe('workflow templates', () => {
  it('discovers the named templates and prepares engine-backed plans', () => {
    const registry = createOperationRegistry();
    const discoveries = discoverWorkflowTemplates(workbook(), 'Orders', registry);

    expect(discoveries.map((item) => item.template.id)).toEqual([...WORKFLOW_TEMPLATE_IDS]);
    for (const discovery of discoveries) {
      expect(discovery.result.ok, discovery.template.id).toBe(true);
      if (!discovery.result.ok) continue;
      expect(discovery.result.plan.steps.length).toBeGreaterThan(0);
      expect(
        discovery.result.plan.steps.every((step) => registry.get(step.operation) !== undefined),
      ).toBe(true);
      expect(discovery.result.plan.steps.every((step) => step.preview?.valid === true)).toBe(true);
      expect(discovery.result.plan.assumptions?.length).toBeGreaterThan(0);
    }
  });

  it('executes the weekly export as a real reviewable plan without changing its source', () => {
    const before = workbook();
    const registry = createOperationRegistry();
    const result = planWorkflowTemplate('clean_weekly_export', before, 'Orders', registry);

    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(result.reason);
    const staged = applyOperationPlan(before, result.plan.steps, {
      registry,
      confirmed: true,
    });

    expect(staged.ok).toBe(true);
    if (!staged.ok) throw new Error(staged.error.messages.join(' '));
    expect(staged.workbook.sheets).toHaveLength(3);
    expect(staged.workbook.sheets[0]).toEqual(before.sheets[0]);
    expect(staged.workbook.sheets[2]?.name).toBe('Weekly Clean Export');
  });

  it('plans date normalization and duplicate removal through the existing operations', () => {
    const registry = createOperationRegistry();
    const dates = planWorkflowTemplate('normalize_dates', workbook(), 'Orders', registry);
    const duplicates = planWorkflowTemplate('remove_duplicates', workbook(), 'Orders', registry);

    expect(dates.ok).toBe(true);
    if (dates.ok) {
      expect(dates.plan.steps.map((step) => step.operation)).toEqual(['format_dates']);
      expect(dates.plan.steps[0]?.args).toMatchObject({
        sheet: 'Orders',
        column: 'B',
        format: 'YYYY-MM-DD',
      });
    }
    expect(duplicates.ok).toBe(true);
    if (duplicates.ok) {
      expect(duplicates.plan.steps[0]?.operation).toBe('delete_duplicates');
      expect(duplicates.plan.steps[0]?.args).toMatchObject({
        sheet: 'Orders',
        columns: ['A'],
        keep: 'first',
      });
      expect(duplicates.plan.steps[0]?.preview?.requiresConfirmation).toBe(true);
    }
  });

  it('explains unsupported or ambiguous inputs instead of exposing a fake action', () => {
    const empty: Workbook = {
      sheets: [{ name: 'Notes', rows: [[createCell('A script or note')]] }],
    };
    const registry = createOperationRegistry();

    const emptyDiscoveries = discoverWorkflowTemplates(empty, 'Notes', registry);
    expect(emptyDiscoveries.every((item) => !item.result.ok)).toBe(true);
    expect(
      emptyDiscoveries.map((item) => (item.result.ok ? '' : item.result.reason)).join(' '),
    ).toMatch(/needs|no|date|worksheet/i);

    const threeSheets: Workbook = {
      sheets: [
        ...workbook().sheets,
        { name: 'Archive', rows: [[createCell('Order ID')], [createCell('ORD-1')]] },
      ],
    };
    const reconciliation = planWorkflowTemplate(
      'reconcile_sheets',
      threeSheets,
      'Orders',
      registry,
    );
    expect(reconciliation.ok).toBe(false);
    if (!reconciliation.ok) {
      expect(reconciliation.kind).toBe('needs-input');
      expect(reconciliation.reason).toMatch(/two source sheets|which two/i);
    }
  });

  it('reports a no-op when there are no duplicate rows to remove', () => {
    const clean = workbook();
    clean.sheets[0]!.rows.pop();
    const result = planWorkflowTemplate('remove_duplicates', clean, 'Orders');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.kind).toBe('no-op');
      expect(result.reason).toMatch(/no|nothing/i);
    }
  });
});
