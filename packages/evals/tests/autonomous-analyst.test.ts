import { describe, expect, it, vi, afterEach } from 'vitest';
import {
  applyOperationPlan,
  createCell,
  createOperationRegistry,
  type Workbook,
} from '@excel-agent/engine';
import { createOrchestrator } from '@excel-agent/agent';

function sampleEnterpriseWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Orders',
        rows: [
          [
            createCell('Order ID'),
            createCell('Date'),
            createCell('Customer'),
            createCell('Region'),
            createCell('Amount'),
          ],
          [
            createCell('ORD-101'),
            createCell('2026-10-01'),
            createCell('Acme Corp'),
            createCell('North'),
            createCell(1200),
          ],
          [
            createCell('ORD-102'),
            createCell('2026-10-02'),
            createCell('Beta LLC'),
            createCell('South'),
            createCell(850),
          ],
          [
            createCell('ORD-103'),
            createCell('2026-10-02'),
            createCell('Gamma Inc'),
            createCell('North'),
            createCell(430),
          ],
          [
            createCell('ORD-104'),
            createCell('2026-10-03'),
            createCell('Delta Co'),
            createCell('West'),
            createCell(950),
          ],
          // Duplicate row to test cleaning
          [
            createCell('ORD-104'),
            createCell('2026-10-03'),
            createCell('Delta Co'),
            createCell('West'),
            createCell(950),
          ],
          [
            createCell('ORD-105'),
            createCell('2026-10-04'),
            createCell('Epsilon Ltd'),
            createCell('East'),
            createCell(2100),
          ],
        ],
      },
      {
        name: 'Invoices',
        rows: [
          [
            createCell('Invoice ID'),
            createCell('Date'),
            createCell('Amount'),
            createCell('Status'),
          ],
          [createCell('ORD-101'), createCell('2026-10-01'), createCell(1200), createCell('Paid')],
          [createCell('ORD-102'), createCell('2026-10-02'), createCell(850), createCell('Paid')],
          // Discrepancy: Amount mismatch ($400 instead of $430)
          [createCell('ORD-103'), createCell('2026-10-02'), createCell(400), createCell('Partial')],
          [createCell('ORD-104'), createCell('2026-10-03'), createCell(950), createCell('Paid')],
          // ORD-105 is missing from Invoices (another discrepancy!)
        ],
      },
    ],
  };
}

const OUTCOME_QUERY =
  'Prepare this month’s sales report. Clean the orders, reconcile them against invoices, flag discrepancies, summarize performance by region, and give me a presentation-ready workbook.';

describe('North Star: Autonomous AI Analyst delivering finished Excel work', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('formulates an end-to-end verified deliverable plan from plain english outcome (deterministic path)', async () => {
    const registry = createOperationRegistry();
    const orchestrator = createOrchestrator({ registry });
    const initialWorkbook = sampleEnterpriseWorkbook();

    const decision = await orchestrator.decide({
      query: OUTCOME_QUERY,
      workbook: initialWorkbook,
      sheetName: 'Orders',
    });

    // 1. Must produce an ExecutionPlan (not refuse or return a formula)
    expect(decision.plan).toBeDefined();
    const plan = decision.plan!;
    expect(plan.status).toBe('pending');
    expect(plan.steps.length).toBeGreaterThanOrEqual(3);

    // 2. Coordinated stages
    const opNames = plan.steps.map((s) => s.operation);
    expect(opNames).toContain('delete_duplicates');
    expect(opNames).toContain('reconcile_sheets');
    expect(opNames).toContain('group_and_summarize');

    // 3. Execute the full deliverable through the engine with Sentinel verification
    const planResult = applyOperationPlan(
      initialWorkbook,
      plan.steps.map((s) => ({ operation: s.operation, args: s.args })),
      { registry, confirmed: true },
    );

    expect(planResult.ok).toBe(true);
    if (!planResult.ok) return;

    const finishedWorkbook = planResult.workbook;

    // 4. Verify the deliverable workbook structure
    const sheetNames = finishedWorkbook.sheets.map((s) => s.name);
    // Source sheets are preserved
    expect(sheetNames).toContain('Orders');
    expect(sheetNames).toContain('Invoices');

    // Cleaned orders: duplicate ORD-104 was removed (6 rows including header instead of 7)
    const ordersSheet = finishedWorkbook.sheets.find((s) => s.name === 'Orders')!;
    expect(ordersSheet.rows.length).toBe(6);

    // Reconciled sheets with discrepancies flagged
    expect(sheetNames.some((n) => n.includes('Reconciliation') && n.includes('Summary'))).toBe(
      true,
    );
    expect(sheetNames.some((n) => n.includes('Reconciliation') && n.includes('Exceptions'))).toBe(
      true,
    );
    expect(sheetNames.some((n) => n.includes('Reconciliation') && n.includes('Matched'))).toBe(
      true,
    );

    // Regional performance summary generated
    expect(sheetNames).toContain('Regional Performance');
    const regionalSheet = finishedWorkbook.sheets.find((s) => s.name === 'Regional Performance')!;
    expect(regionalSheet.rows.length).toBeGreaterThan(1); // Header + region aggregations
  });

  it('coordinates multi-agent swarm (Analyst, Planner, Critic, Sentinel) on composite outcome', async () => {
    const registry = createOperationRegistry();
    const orchestrator = createOrchestrator({ registry });
    const initialWorkbook = sampleEnterpriseWorkbook();

    // Mock provider to simulate specialist LLM swarm
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url: RequestInfo | URL, init?: RequestInit) => {
        const body = JSON.parse(String(init?.body));
        const systemPrompt = body.messages[0]?.content ?? '';
        const isPlanner = systemPrompt.includes('You are the Planner');
        const isCritic = systemPrompt.includes('You are the Critic');

        if (isPlanner) {
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      title: 'Monthly Sales Report & Reconciliation Deliverable',
                      description:
                        'Clean orders, reconcile against invoices, flag discrepancies, and summarize regional performance.',
                      steps: [
                        {
                          operation: 'delete_duplicates',
                          args: { sheet: 'Orders', columns: ['A'], keep: 'first', headerRow: 1 },
                          description: 'Clean orders by removing duplicate Order IDs',
                        },
                        {
                          operation: 'reconcile_sheets',
                          args: {
                            leftSheet: 'Orders',
                            rightSheet: 'Invoices',
                            leftKeys: ['A'],
                            rightKeys: ['A'],
                            leftAmount: 'E',
                            rightAmount: 'C',
                            tolerance: 0,
                            reportPrefix: 'Sales Reconciliation',
                          },
                          description:
                            'Reconcile Orders against Invoices and extract discrepancy exceptions',
                        },
                        {
                          operation: 'group_and_summarize',
                          args: {
                            sheet: 'Orders',
                            groupBy: ['D'],
                            valueColumn: 'E',
                            aggregation: 'sum',
                            targetSheet: 'Regional Performance',
                          },
                          description: 'Summarize total sales revenue grouped by region',
                        },
                      ],
                    }),
                  },
                },
              ],
              model: 'llama-3.3-70b-versatile',
            }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }

        if (isCritic) {
          return new Response(
            JSON.stringify({
              choices: [{ message: { content: '{"approved": true, "issues": []}' } }],
              model: 'llama-3.3-70b-versatile',
            }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }

        // Analyst summary
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  content:
                    'Workbook contains Orders (6 orders, regions North/South/West/East) and Invoices (4 billed). One duplicate detected in Orders.',
                },
              },
            ],
            model: 'llama-3.3-70b-versatile',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } },
        );
      }),
    );

    const decision = await orchestrator.decide({
      query: OUTCOME_QUERY,
      workbook: initialWorkbook,
      sheetName: 'Orders',
      config: { provider: 'groq', apiKey: 'gsk_live_key' },
    });

    expect(decision.source).toBe('llm');
    expect(decision.plan).toBeDefined();
    expect(decision.plan?.steps).toHaveLength(3);
    expect(decision.plan?.status).toBe('pending');
    expect(decision.plan?.steps.every((s) => s.status === 'pending')).toBe(true);

    // Verify Sentinel invariant verification ran for every step
    expect(decision.trace.map((s) => s.layer)).toContain('verification');
    expect(decision.activities?.map((a) => a.agent)).toContain('Critic');
    expect(decision.activities?.map((a) => a.agent)).toContain('Planner');
    expect(decision.activities?.map((a) => a.agent)).toContain('Analyst');
  });
});
