import { describe, expect, it, vi, afterEach } from 'vitest';
import { createOperationRegistry, createCell, type Workbook } from '@excel-agent/engine';
import { ExcelAgentOrchestrator } from '../src/orchestrator.js';
import { gatherGroundedFacts, isComplexRequest } from '../src/multi-agent.js';

describe('multi-agent complexity routing', () => {
  it('stays on the fast path for simple single-verb requests', () => {
    expect(isComplexRequest('remove duplicate rows')).toBe(false);
    expect(isComplexRequest('sort by revenue descending')).toBe(false);
    expect(isComplexRequest('set cell A1 to 5')).toBe(false);
  });

  it('routes multi-step requests to the full pipeline', () => {
    expect(isComplexRequest('clean duplicates and then sort by revenue')).toBe(true);
    expect(isComplexRequest('analyse revenue by region and then group by state')).toBe(true);
    expect(isComplexRequest('summarize this sheet and remove empty rows')).toBe(true);
  });

  it('routes analytical requests', () => {
    expect(isComplexRequest('compare north and south regional sales')).toBe(true);
    expect(isComplexRequest('group orders by region and sum the totals')).toBe(true);
  });
});

describe('multi-agent large worksheet grounding', () => {
  it('profiles tall worksheets without spreading every row into a function call', async () => {
    const rows = Array.from({ length: 130_000 }, (_, index) => [
      createCell(index === 0 ? 'ID' : index),
    ]);
    const workbook: Workbook = { sheets: [{ name: 'Tall', rows }] };

    await expect(gatherGroundedFacts(workbook, 'Tall')).resolves.toMatchObject({
      columns: [{ name: 'ID', nonBlank: 129_999 }],
    });
  });

  it('grounds named columns beyond Z and data from non-active worksheets', async () => {
    const sales = sampleWorkbook().sheets[0]!;
    const archiveHeaders = Array.from({ length: 29 }, (_, index) =>
      createCell(index === 28 ? 'Region' : `Field ${index + 1}`),
    );
    const workbook: Workbook = {
      sheets: [
        sales,
        {
          name: 'Archive',
          rows: [
            archiveHeaders,
            Array.from({ length: 29 }, (_, index) => createCell(index === 28 ? 'West' : index)),
          ],
        },
      ],
    };

    const facts = await gatherGroundedFacts(
      workbook,
      'Sales',
      'summarize Region from Archive column AC alongside Revenue from Sales',
    );

    expect(facts.columns).toContainEqual(
      expect.objectContaining({ sheet: 'Archive', letter: 'AC', name: 'Region' }),
    );
    expect(facts.overview).toContain('Archive');
    expect(facts.overview).toContain('Sales');
  });
});

function sampleWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: 'Sales',
        rows: [
          [createCell('Region'), createCell('Revenue'), createCell('Quarter')],
          [createCell('North'), createCell(1200), createCell('Q1')],
          [createCell('South'), createCell(900), createCell('Q1')],
          [createCell('North'), createCell(1500), createCell('Q2')],
        ],
      },
    ],
  };
}

describe('runMultiAgentTurn', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('decomposes, runs specialists, and returns a verified plan shape', async () => {
    let roleSystemPrompt = '';
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const body = JSON.parse(String(init?.body));
        roleSystemPrompt = body.messages[0]?.content ?? '';
        const isPlanner = roleSystemPrompt.includes('You are the Planner');
        const isCritic = roleSystemPrompt.includes('You are the Critic');
        if (isPlanner) {
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      title: 'Revenue rollup',
                      description: 'Group revenue by region',
                      steps: [
                        {
                          operation: 'group_and_summarize',
                          args: {
                            sheet: 'Sales',
                            groupBy: ['A'],
                            valueColumn: 'B',
                            aggregation: 'sum',
                            targetSheet: 'Summary',
                          },
                          description: 'Aggregate revenue by region',
                        },
                      ],
                    }),
                  },
                },
              ],
              model: 'test',
            }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }
        if (isCritic) {
          return new Response(
            JSON.stringify({
              choices: [{ message: { content: '{"approved": true, "issues": []}' } }],
              model: 'test',
            }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }
        // Analyst: return plain prose so the plan payload stays the planner's job.
        return new Response(
          JSON.stringify({
            choices: [{ message: { content: 'Revenue by region looks concentrated in North.' } }],
            model: 'test',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } },
        );
      }),
    );

    const orchestrator = new ExcelAgentOrchestrator({ registry: createOperationRegistry() });
    const result = await orchestrator.decide({
      query: 'analyse revenue by region and then group the results by region',
      workbook: sampleWorkbook(),
      sheetName: 'Sales',
      config: { provider: 'groq', apiKey: 'gsk_test' },
    });

    expect(result.source).toBe('llm');
    expect(result.plan).toBeDefined();
    expect(result.plan?.steps).toHaveLength(1);
    expect(result.plan?.steps[0]?.operation).toBe('group_and_summarize');
    expect(result.trace.map((step) => step.layer)).toContain('planner');
    expect(result.trace.map((step) => step.layer)).toContain('verification');
    expect(result.activities?.map((activity) => activity.agent)).toContain('Critic');
  });

  it('falls back honestly when the provider fails', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response('err', { status: 500 })),
    );
    const orchestrator = new ExcelAgentOrchestrator({ registry: createOperationRegistry() });
    const result = await orchestrator.decide({
      query: 'summarize this and then remove blank rows',
      workbook: sampleWorkbook(),
      sheetName: 'Sales',
      config: { provider: 'groq', apiKey: 'gsk_test' },
    });
    // Multi-agent errors must surface; they must not throw because the orchestrator owns the UX.
    expect(['llm', 'fallback']).toContain(result.source);
    expect(result.trace.length).toBeGreaterThan(0);
  });

  it('delivers grounded analytical findings directly without requiring a mutation plan', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const body = JSON.parse(String(init?.body));
        const roleSystemPrompt = body.messages[0]?.content ?? '';
        const isPlanner = roleSystemPrompt.includes('You are the Planner');
        if (isPlanner) {
          // Planner produces no mutation steps because request was informational
          return new Response(
            JSON.stringify({
              choices: [{ message: { content: '{"title": "Analysis Only", "steps": []}' } }],
              model: 'test',
            }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          );
        }
        // Analyst delivers grounded analysis
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  content:
                    'Row 21 total across all recorded periods is 73,456 with an average of 7,345.60.',
                },
              },
            ],
            model: 'test',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } },
        );
      }),
    );

    const orchestrator = new ExcelAgentOrchestrator({ registry: createOperationRegistry() });
    const result = await orchestrator.decide({
      query: 'compare and analyze row 21 total',
      workbook: sampleWorkbook(),
      sheetName: 'Sales',
      config: { provider: 'groq', apiKey: 'gsk_test' },
    });

    expect(result.source).toBe('llm');
    expect(result.message).toContain('73,456');
    expect(result.message).not.toContain('I could not produce a safe execution plan');
  });
});
