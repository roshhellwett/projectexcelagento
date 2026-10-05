import { afterEach, describe, expect, it, vi } from 'vitest';

import { createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';

import { createOrchestrator, type ProviderConfig } from '../src/index.js';

const SHEET = 'Orders';

/**
 * A workbook with a duplicate row so the deterministic layers have something real
 * to find, which lets the tests prove the LLM is genuinely optional.
 */
function workbook(): Workbook {
  return {
    sheets: [
      {
        name: SHEET,
        rows: [
          [createCell('Order ID'), createCell('Customer'), createCell('Amount')],
          [createCell('ORD-1'), createCell('Acme'), createCell(100)],
          [createCell('ORD-2'), createCell('Beta'), createCell(250)],
          [createCell('ORD-2'), createCell('Beta'), createCell(250)],
        ],
      },
    ],
  };
}

/** Wrap an action exactly as the system prompt instructs the model to reply. */
function llmResponse(action: Record<string, unknown>, usage?: Record<string, number>): Response {
  const content = `Here is the plan.\n\n\`\`\`json\n${JSON.stringify(action)}\n\`\`\``;
  return new Response(
    JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      choices: [{ index: 0, message: { role: 'assistant', content }, finish_reason: 'stop' }],
      usage: usage ?? { prompt_tokens: 900, completion_tokens: 80, total_tokens: 980 },
    }),
    { status: 200, headers: { 'content-type': 'application/json' } },
  );
}

type FetchImpl = (url: string, init: RequestInit) => Promise<Response>;

function stubFetch(impl: FetchImpl) {
  const mock = vi.fn(impl);
  vi.stubGlobal('fetch', mock);
  return mock;
}

const LIVE_CONFIG: ProviderConfig = { provider: 'groq', apiKey: 'gsk_live_key' };

function orchestrator() {
  return createOrchestrator({ registry: createOperationRegistry() });
}

function systemPromptFrom(init: RequestInit): string {
  const body = JSON.parse(String(init.body)) as {
    messages: { role: string; content: string }[];
  };
  return body.messages.find((message) => message.role === 'system')?.content ?? '';
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('LLM layer drives operations the deterministic planner cannot express', () => {
  it('executes an LLM-proposed set_cells action after the guardrail approves it', async () => {
    stubFetch(async () =>
      llmResponse({
        name: 'set_cells',
        args: { sheet: SHEET, cells: [{ row: 2, column: 'C', value: 999 }] },
        explanation: 'Set C2 to 999.',
      }),
    );

    const decision = await orchestrator().decide({
      query: 'set C2 to 999',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.action?.name).toBe('set_cells');
    expect(decision.source).toBe('llm');
    expect(decision.guardrail?.passed).toBe(true);
    expect(decision.guardrail?.requiresConfirmation).toBe(false);
  });

  it('executes an LLM-proposed add_column action after the guardrail approves it', async () => {
    stubFetch(async () =>
      llmResponse({
        name: 'add_column',
        args: { sheet: SHEET, column: 'D', headerName: 'Notes', headerRow: 1 },
        explanation: 'Add a Notes column.',
      }),
    );

    const decision = await orchestrator().decide({
      query: 'add a notes column',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.action?.name).toBe('add_column');
    expect(decision.source).toBe('llm');
    expect(decision.guardrail?.passed).toBe(true);
  });
});

describe('the guardrail is the hard wall', () => {
  it('blocks a hallucinated operation and reports the reason', async () => {
    stubFetch(async () =>
      llmResponse({ name: 'drop_database', args: {}, explanation: 'Not a real operation.' }),
    );

    const decision = await orchestrator().decide({
      query: 'please wipe everything',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.action).toBeUndefined();
    expect(decision.source).toBe('fallback');
    const blocked = decision.trace.filter(
      (step) => step.layer === 'guardrail' && step.summary.includes('blocked'),
    );
    expect(blocked.length).toBeGreaterThan(0);
    expect(JSON.stringify(blocked)).toContain('drop_database');
  });

  it('blocks a real operation whose arguments the engine rejects', async () => {
    stubFetch(async () =>
      llmResponse({
        name: 'format_dates',
        args: { sheet: SHEET, column: 'ZZ', format: 'YYYY-MM-DD', headerRow: 1 },
        explanation: 'Format a column that does not exist.',
      }),
    );

    // A query the deterministic planner cannot interpret, so the model's action
    // is the only candidate and its rejection is unambiguous.
    const decision = await orchestrator().decide({
      query: 'apply the change I mentioned',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.action).toBeUndefined();
    // The engine's own column-existence check is what rejected it.
    expect(JSON.stringify(decision.trace)).toContain('does not exist');
  });

  it('blocks a set_cells action that targets the same address twice', async () => {
    stubFetch(async () =>
      llmResponse({
        name: 'set_cells',
        args: {
          sheet: SHEET,
          cells: [
            { row: 2, column: 'C', value: 1 },
            { row: 2, column: 'C', value: 2 },
          ],
        },
        explanation: 'Duplicate address.',
      }),
    );

    const decision = await orchestrator().decide({
      query: 'set the same cell twice',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.action).toBeUndefined();
  });

  it('surfaces the error to the user when the model call fails instead of silent fallback', async () => {
    stubFetch(async () => new Response('{"error":{"message":"Invalid API Key"}}', { status: 401 }));

    const decision = await orchestrator().decide({
      query: 'remove duplicate rows',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.message).toContain('OpenRouter / Model Error');
    expect(decision.telemetry?.ok).toBe(false);
    expect(decision.action).toBeUndefined();
  });
});

describe('telemetry', () => {
  it('records provider-reported tokens and latency for a model call', async () => {
    stubFetch(async () =>
      llmResponse({
        name: 'set_cells',
        args: { sheet: SHEET, cells: [{ row: 2, column: 'C', value: 5 }] },
        explanation: 'Set C2.',
      }),
    );

    const decision = await orchestrator().decide({
      query: 'set C2 to 5',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.telemetry).toMatchObject({
      provider: 'groq',
      model: 'llama-3.3-70b-versatile',
      promptTokens: 900,
      completionTokens: 80,
      totalTokens: 980,
      ok: true,
    });
    expect(decision.telemetry?.latencyMs).toBeGreaterThanOrEqual(0);
  });

  it('derives a total when the provider omits one', async () => {
    stubFetch(async () =>
      llmResponse(
        {
          name: 'set_cells',
          args: { sheet: SHEET, cells: [{ row: 2, column: 'C', value: 5 }] },
          explanation: 'Set C2.',
        },
        { prompt_tokens: 100, completion_tokens: 25 },
      ),
    );

    const decision = await orchestrator().decide({
      query: 'set C2 to 5',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.telemetry?.totalTokens).toBe(125);
  });

  it('records a failed call with its error instead of throwing', async () => {
    stubFetch(async () => new Response('{"error":{"message":"Invalid API Key"}}', { status: 401 }));

    const decision = await orchestrator().decide({
      query: 'remove duplicate rows',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(decision.telemetry?.ok).toBe(false);
    expect(decision.telemetry?.error).toContain('Invalid API Key');
    expect(decision.telemetry?.provider).toBe('groq');
    expect(decision.telemetry?.model).toBe('unknown');
  });

  it('omits telemetry entirely when no provider is configured', async () => {
    const fetchMock = stubFetch(async () => llmResponse({ name: 'set_cells', args: {} }));

    const decision = await orchestrator().decide({
      query: 'remove duplicate rows',
      workbook: workbook(),
      sheetName: SHEET,
    });

    expect(decision.telemetry).toBeUndefined();
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe('privacy and token discipline', () => {
  it('never contacts the provider in demo mode', async () => {
    const fetchMock = stubFetch(async () => llmResponse({ name: 'set_cells', args: {} }));

    const decision = await orchestrator().decide({
      query: 'remove duplicate rows',
      workbook: workbook(),
      sheetName: SHEET,
      config: { provider: 'groq', apiKey: 'demo-local-mode' },
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(decision.telemetry).toBeUndefined();
    expect(decision.action?.name).toBe('delete_duplicates');
  });

  it('answers greetings without spending a single token', async () => {
    const fetchMock = stubFetch(async () => llmResponse({ name: 'set_cells', args: {} }));

    const decision = await orchestrator().decide({
      query: 'hello there',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(decision.source).toBe('heuristic');
    expect(decision.message).toContain('ExcelAgento');
  });

  it('sends only a compact sheet profile plus the tool catalogue', async () => {
    const fetchMock = stubFetch(async () =>
      llmResponse({
        name: 'set_cells',
        args: { sheet: SHEET, cells: [{ row: 2, column: 'C', value: 1 }] },
        explanation: 'Set C2.',
      }),
    );

    await orchestrator().decide({
      query: 'set C2 to 1',
      workbook: workbook(),
      sheetName: SHEET,
      config: LIVE_CONFIG,
    });

    const init = fetchMock.mock.calls[0]![1];
    const system = systemPromptFrom(init);

    // The engine contract is present, derived from the live registry.
    expect(system).toContain('format_dates');
    expect(system).toContain('set_cells');
    expect(system).toContain(SHEET);
    // The profile carries headers and aggregates, and stays token-efficient.
    expect(system).toContain('Order ID');
    // The ceiling tracks the size of the operation catalogue, which grows when the engine gains an
    // operation. It was 16,000 for the original 17 tools; the business and sheet operations added
    // to the catalogue grow the budget proportionally to maintain the live schema contract.
    expect(system.length).toBeLessThan(35000);

    // The user turn is exactly what the user typed, nothing added.
    const body = JSON.parse(String(init.body)) as {
      messages: { role: string; content: string }[];
    };
    expect(body.messages.at(-1)).toEqual({ role: 'user', content: 'set C2 to 1' });
  });

  it('delivers conversational LLM response directly without heuristic hijacking on user feedback', async () => {
    const singleColumnCodeWorkbook: Workbook = {
      sheets: [
        {
          name: 'Worksheet_Cleaned',
          rows: [
            [createCell('#!/usr/bin/env python3')],
            [createCell('"""zenith_leads.py - Kolkata lead pipeline"""')],
            [createCell('python zenith_leads.py discover --max-queries 20')],
          ],
        },
      ],
    };

    stubFetch(async () =>
      new Response(
        JSON.stringify({
          model: 'llama-3.3-70b-versatile',
          choices: [
            {
              index: 0,
              message: {
                role: 'assistant',
                content:
                  "You're completely right. This worksheet contains Python source code for `zenith_leads.py` rather than a tabular dataset, so running a cleaning operation simply duplicated the script lines.",
              },
              finish_reason: 'stop',
            },
          ],
          usage: { prompt_tokens: 500, completion_tokens: 45, total_tokens: 545 },
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      ),
    );

    const decision = await orchestrator().decide({
      query: 'you have just copy pasted the sheet',
      workbook: singleColumnCodeWorkbook,
      sheetName: 'Worksheet_Cleaned',
      config: LIVE_CONFIG,
    });

    expect(decision.action).toBeUndefined();
    expect(decision.plan).toBeUndefined();
    expect(decision.source).toBe('llm');
    expect(decision.message).toContain("You're completely right");
    expect(decision.message).not.toContain('clean_to_new_sheet');
  });

  it('triggers data population workflow when user asks in Hindi/Hinglish ("data daal isme") on empty structured table', async () => {
    let capturedBody: any;
    stubFetch(async (_url, init) => {
      capturedBody = JSON.parse(init?.body as string);
      return new Response(
        JSON.stringify({
          model: 'llama-3.3-70b-versatile',
          choices: [
            {
              index: 0,
              message: {
                role: 'assistant',
                content:
                  'I analyzed the source records and structured the leads into Worksheet_Structured.',
                tool_calls: [
                  {
                    id: 'call-append-1',
                    type: 'function',
                    function: {
                      name: 'append_rows',
                      arguments: JSON.stringify({
                        sheet: 'Worksheet_Structured',
                        rows: [
                          [1, 'Zenith Realty', 'Kolkata', 'Real Estate', '+91 9876543210', 'info@zenith.com', 'zenith.com', 'High', 'Active', 'Lead generation system'],
                        ],
                      }),
                    },
                  },
                ],
              },
              finish_reason: 'stop',
            },
          ],
          usage: { prompt_tokens: 600, completion_tokens: 80, total_tokens: 680 },
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    });

    const multiSheetWorkbook: Workbook = {
      sheets: [
        {
          name: 'Worksheet',
          rows: [
            [createCell('#!/usr/bin/env python3')],
            [createCell('zenith_leads.py - Kolkata lead pipeline')],
            [createCell('Company: Zenith Realty, Kolkata, Phone: +91 9876543210')],
            [createCell('Email: info@zenith.com, Website: zenith.com')],
          ],
        },
        {
          name: 'Worksheet_Structured',
          rows: [
            [
              createCell('#'),
              createCell('Company'),
              createCell('Area'),
              createCell('Business Type'),
              createCell('Phone'),
              createCell('Email'),
              createCell('Website'),
            ],
          ],
        },
      ],
    };

    const decision = await orchestrator().decide({
      query: 'data daal isme',
      workbook: multiSheetWorkbook,
      sheetName: 'Worksheet_Structured',
      config: LIVE_CONFIG,
    });

    // Check system prompt included Data Population directive
    const systemMsg = capturedBody.messages.find((m: any) =>
      typeof m.content === 'string' && m.content.includes('Data Population & Extraction Directive'),
    );
    expect(systemMsg).toBeDefined();
    expect(systemMsg.content).toContain('Source Sheet: "Worksheet"');
    expect(systemMsg.content).toContain('Destination Sheet: "Worksheet_Structured"');

    // Model action is append_rows
    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('append_rows');
    expect(decision.action?.args.sheet).toBe('Worksheet_Structured');
    expect(decision.message).not.toContain('I analyzed Worksheet_Structured (0 rows)');
  });

  it('never outputs the robotic heuristic prompt recommendations when model responds', async () => {
    stubFetch(async () =>
      new Response(
        JSON.stringify({
          model: 'llama-3.3-70b-versatile',
          choices: [
            {
              index: 0,
              message: {
                role: 'assistant',
                content: 'I have inspected your sheet and I am ready to process your query.',
              },
              finish_reason: 'stop',
            },
          ],
          usage: { prompt_tokens: 300, completion_tokens: 20, total_tokens: 320 },
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      ),
    );

    const emptySheetWorkbook: Workbook = {
      sheets: [
        {
          name: 'Worksheet_Structured',
          rows: [[createCell('#'), createCell('Company')]],
        },
      ],
    };

    const decision = await orchestrator().decide({
      query: 'now fill the data',
      workbook: emptySheetWorkbook,
      sheetName: 'Worksheet_Structured',
      config: LIVE_CONFIG,
    });

    expect(decision.message).not.toContain('Tell me what transformation or analysis you would like to run');
    expect(decision.message).toContain('I have inspected your sheet');
  });
});


