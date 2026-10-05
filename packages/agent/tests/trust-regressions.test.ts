import { afterEach, describe, expect, it, vi } from 'vitest';
import { createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';
import { createOrchestrator, geminiAdapter, groqAdapter, type ChatMessage } from '../src/index.js';
import { runMultiAgentTurn } from '../src/multi-agent.js';

const workbook: Workbook = {
  sheets: [
    {
      name: 'Sales',
      rows: [
        [createCell('Region'), createCell('Income')],
        [createCell('North'), createCell(-100)],
        [createCell('South'), createCell(-200)],
      ],
    },
  ],
};
const config = { provider: 'groq' as const, apiKey: 'test-key', model: 'test-model', retries: 0 };
const json = (content: string, usage?: Record<string, number>, toolCalls?: unknown[]) =>
  new Response(
    JSON.stringify({
      model: 'served-model',
      choices: [
        {
          message: {
            content,
            ...(toolCalls ? { tool_calls: toolCalls } : {}),
          },
        },
      ],
      usage,
    }),
    { headers: { 'Content-Type': 'application/json' } },
  );
const call = (name: string, args: Record<string, unknown> = {}) => ({
  id: `call-${name}`,
  type: 'function',
  function: { name, arguments: JSON.stringify(args) },
});
const candidatePlan = {
  title: 'Income summary',
  description: 'Summarize income',
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
      description: 'Group income by region',
    },
  ],
};
const turn = () =>
  runMultiAgentTurn(
    {
      query: 'compare regional income and group the results',
      workbook,
      sheetName: 'Sales',
      config,
      emit: () => {},
    },
    { registry: createOperationRegistry() },
  );

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('analytical integrity', () => {
  it('never manufactures profit probabilities to complete a provider preamble', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => json("Here's what the data shows:")),
    );
    const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({
      query: 'what is in column B?',
      workbook,
      sheetName: 'Sales',
      config,
    });
    expect(decision.message).not.toContain('100%');
    expect(decision.message).not.toContain('consistently positive');
    expect(decision.message).toMatch(/incomplete|could not|couldn't|did not/i);
    expect(decision.action).toBeUndefined();
  });

  it('accounts for all successful inference calls in a read-tool turn', async () => {
    let count = 0;
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        ++count === 1
          ? json('', { prompt_tokens: 20, completion_tokens: 10, total_tokens: 30 }, [
              call('get_workbook_overview'),
            ])
          : json('Sales contains three rows.', {
              prompt_tokens: 50,
              completion_tokens: 20,
              total_tokens: 70,
            }),
      ),
    );
    const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({
      query: 'what is in this workbook?',
      workbook,
      sheetName: 'Sales',
      config,
    });
    expect(count).toBe(2);
    expect(decision.telemetry).toMatchObject({
      promptTokens: 70,
      completionTokens: 30,
      totalTokens: 100,
      model: 'served-model',
    });
  });
});

describe('critic approval is mandatory', () => {
  it.each([
    '{"approved":false,"issues":["Wrong request"],"revisions":[]}',
    'not json',
    '{"approved":"false"}',
    '{"approved":true,"issues":["Deletes unrequested data"]}',
  ])('never offers a plan after an unsuccessful review: %s', async (critique) => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url, init: RequestInit) => {
        const prompt = JSON.parse(String(init.body)).messages[0].content;
        if (prompt.includes('You are the Planner')) return json(JSON.stringify(candidatePlan));
        if (prompt.includes('You are the Critic')) return json(critique);
        return json('The full Income column contains two negative observations.');
      }),
    );
    const decision = await turn();
    expect(decision.plan).toBeUndefined();
    expect(decision.message).toMatch(/review|critic|approve/i);
  });

  it('repairs once, reviews the revised plan, and accounts for the complete pipeline', async () => {
    let plans = 0;
    let reviews = 0;
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url, init: RequestInit) => {
        const body = JSON.parse(String(init.body));
        const prompt = body.messages[0].content;
        const usage = { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 };
        if (prompt.includes('You are the Planner')) {
          plans += 1;
          if (plans === 2)
            expect(body.messages.map((m: ChatMessage) => m.content).join(' ')).toContain(
              'Use a summary sheet',
            );
          return json(
            JSON.stringify({
              ...candidatePlan,
              title: plans === 1 ? 'First draft' : 'Reviewed revision',
            }),
            usage,
          );
        }
        if (prompt.includes('You are the Critic')) {
          reviews += 1;
          return json(
            JSON.stringify(
              reviews === 1
                ? {
                    approved: false,
                    issues: ['Use a summary sheet'],
                    revisions: [{ stepIndex: 0, reason: 'Use a summary sheet' }],
                  }
                : { approved: true, issues: [] },
            ),
            usage,
          );
        }
        return json('Both regions have negative income.', usage);
      }),
    );
    const decision = await turn();
    expect(plans).toBe(2);
    expect(reviews).toBe(2);
    expect(decision.plan?.title).toBe('Reviewed revision');
    expect(decision.plan?.status).toBe('pending');
    expect(decision.telemetry?.totalTokens).toBe(75);
    expect(decision.telemetry?.model).toBe('served-model');
  });

  it('does not silently drop malformed requested steps from a plan', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url, init: RequestInit) => {
        const prompt = JSON.parse(String(init.body)).messages[0].content;
        if (prompt.includes('You are the Planner'))
          return json(
            JSON.stringify({
              ...candidatePlan,
              steps: [
                ...candidatePlan.steps,
                { operation: 'delete_column', args: null, description: 'Invalid' },
              ],
            }),
          );
        if (prompt.includes('You are the Critic')) return json('{"approved":true,"issues":[]}');
        return json('Two income observations are negative.');
      }),
    );
    expect((await turn()).plan).toBeUndefined();
  });

  it('allows the analyst to inspect a requested range through read tools', async () => {
    let analystCalls = 0;
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url, init: RequestInit) => {
        const body = JSON.parse(String(init.body));
        const prompt = body.messages[0].content;
        if (prompt.includes('You are the Analyst')) {
          analystCalls += 1;
          expect(
            body.tools.some(
              (tool: { function: { name: string } }) => tool.function.name === 'describe_column',
            ),
          ).toBe(true);
          if (analystCalls === 1)
            return json('', undefined, [call('describe_column', { sheet: 'Sales', column: 'B' })]);
          const observation = body.messages.find((m: ChatMessage) => m.role === 'tool');
          expect(observation.content).toContain('UNTRUSTED_SPREADSHEET_CONTENT');
          expect(observation.content).toContain('-300');
          return json('Total income is -300 for the full column.');
        }
        if (prompt.includes('You are the Planner'))
          return json('{"title":"Analysis only","steps":[]}');
        return json('{"approved":true,"issues":[]}');
      }),
    );
    expect((await turn()).message).toContain('-300');
    expect(analystCalls).toBe(2);
  });
});

describe('provider contracts and cancellation', () => {
  it('sends native Gemini function calls and results, not unstructured user text', async () => {
    const fetch = vi.fn(async (_url, init: RequestInit) => {
      const body = JSON.parse(String(init.body));
      expect(body.contents[1]).toEqual({
        role: 'model',
        parts: [{ functionCall: { name: 'get_workbook_overview', args: {} } }],
      });
      expect(body.contents[2]).toEqual({
        role: 'user',
        parts: [
          {
            functionResponse: {
              name: 'get_workbook_overview',
              response: { output: { sheets: ['Sales'] } },
            },
          },
        ],
      });
      return new Response(
        JSON.stringify({ candidates: [{ content: { parts: [{ text: 'Sales' }] } }] }),
      );
    });
    vi.stubGlobal('fetch', fetch);
    await geminiAdapter.complete(
      [
        { role: 'user', content: 'Inspect workbook' },
        {
          role: 'assistant',
          content: '',
          tool_calls: [call('get_workbook_overview')] as ChatMessage['tool_calls'],
        },
        {
          role: 'tool',
          name: 'get_workbook_overview',
          tool_call_id: 'call-get_workbook_overview',
          content: '{"sheets":["Sales"]}',
        },
      ],
      { provider: 'gemini', apiKey: 'test-key', retries: 0 },
    );
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it('does not contact a provider with an already cancelled signal', async () => {
    const fetch = vi.fn(async () => json('Should not be sent'));
    vi.stubGlobal('fetch', fetch);
    const controller = new AbortController();
    controller.abort();
    await expect(
      groqAdapter.complete([{ role: 'user', content: 'stop' }], {
        ...config,
        signal: controller.signal,
      }),
    ).rejects.toThrow(/cancel/i);
    expect(fetch).not.toHaveBeenCalled();
  });

  it('does not start another model or offer an action after cancellation during read tools', async () => {
    const controller = new AbortController();
    let count = 0;
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        count += 1;
        if (count === 1) return json('', undefined, [call('get_workbook_overview')]);
        controller.abort();
        throw new DOMException('Aborted', 'AbortError');
      }),
    );
    const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({
      query: 'remove duplicate rows',
      workbook,
      sheetName: 'Sales',
      config,
      signal: controller.signal,
    });
    expect(count).toBe(2);
    expect(decision.action).toBeUndefined();
    expect(decision.plan).toBeUndefined();
    expect(decision.message).toMatch(/cancel/i);
  });
});
