import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  ProviderError,
  complete,
  geminiAdapter,
  getAdapter,
  groqAdapter,
  listProviders,
  openRouterAdapter,
  sleep,
  type ChatMessage,
  type ProviderConfig,
} from '../src/index.js';

/**
 * Recorded provider responses. These are the exact payload shapes the live APIs
 * return, so the adapters are exercised against reality without a network call.
 */
const GROQ_FIXTURE = {
  id: 'chatcmpl-9xQ2',
  object: 'chat.completion',
  created: 1745000000,
  model: 'llama-3.3-70b-versatile',
  choices: [
    {
      index: 0,
      message: { role: 'assistant', content: 'I normalised the dates in column C.' },
      finish_reason: 'stop',
    },
  ],
  usage: { prompt_tokens: 512, completion_tokens: 64, total_tokens: 576 },
};

/** OpenRouter is OpenAI-compatible but reports its own served model. */
const OPENROUTER_FIXTURE = {
  id: 'gen-1745000000-abc',
  model: 'google/gemini-2.0-flash-001',
  choices: [{ index: 0, message: { role: 'assistant', content: 'Done.' }, finish_reason: 'stop' }],
  usage: { prompt_tokens: 700, completion_tokens: 20, total_tokens: 720 },
};

const GEMINI_FIXTURE = {
  candidates: [
    {
      content: { role: 'model', parts: [{ text: 'Here is the summary of the worksheet.' }] },
      finishReason: 'STOP',
    },
  ],
  usageMetadata: { promptTokenCount: 300, candidatesTokenCount: 40, totalTokenCount: 340 },
};

const MESSAGES: ChatMessage[] = [
  { role: 'system', content: 'SYS' },
  { role: 'user', content: 'clean the dates' },
];

const BASE_CONFIG: ProviderConfig = { provider: 'groq', apiKey: 'gsk_test_key' };

type FetchImpl = (url: string, init: RequestInit) => Promise<Response>;

function stubFetch(impl: FetchImpl) {
  const mock = vi.fn(impl);
  vi.stubGlobal('fetch', mock);
  return mock;
}

function jsonResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

/** A fetch that never resolves until its signal aborts, to model a hung upstream. */
function hangingFetch(): FetchImpl {
  return (_url, init) =>
    new Promise<Response>((_resolve, reject) => {
      init.signal?.addEventListener('abort', () =>
        reject(new DOMException('The operation was aborted.', 'AbortError')),
      );
    });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('groq (OpenAI-compatible) adapter', () => {
  it('sends the documented request shape and maps usage into camelCase', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(GROQ_FIXTURE));

    const result = await groqAdapter.complete(MESSAGES, {
      ...BASE_CONFIG,
      apiKey: '  gsk_test_key  ',
    });

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(url).toBe('https://api.groq.com/openai/v1/chat/completions');
    expect(init.method).toBe('POST');

    const headers = init.headers as Record<string, string>;
    expect(headers.Authorization).toBe('Bearer gsk_test_key');
    expect(headers['Content-Type']).toBe('application/json');

    const body = JSON.parse(String(init.body)) as Record<string, unknown>;
    expect(body.model).toBe('llama-3.3-70b-versatile');
    expect(body.messages).toEqual(MESSAGES);
    expect(body.temperature).toBe(0.2);
    expect(body.max_tokens).toBe(8192);

    expect(result).toEqual({
      content: 'I normalised the dates in column C.',
      provider: 'groq',
      model: 'llama-3.3-70b-versatile',
      usage: { promptTokens: 512, completionTokens: 64, totalTokens: 576 },
    });
  });

  it('passes a model override but reports the model actually served', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(GROQ_FIXTURE));

    const result = await groqAdapter.complete(MESSAGES, {
      ...BASE_CONFIG,
      model: 'llama-3.1-8b-instant',
      temperature: 0.7,
      maxTokens: 256,
    });

    const body = JSON.parse(String(fetchMock.mock.calls[0]![1].body)) as Record<string, unknown>;
    expect(body.model).toBe('llama-3.1-8b-instant');
    expect(body.temperature).toBe(0.7);
    expect(body.max_tokens).toBe(256);
    // The provider's own report wins: that is what the usage page displays.
    expect(result.model).toBe('llama-3.3-70b-versatile');
  });

  it('falls back to the adapter default model when none is configured', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(GROQ_FIXTURE));
    await groqAdapter.complete(MESSAGES, BASE_CONFIG);

    const body = JSON.parse(String(fetchMock.mock.calls[0]![1].body)) as Record<string, unknown>;
    expect(body.model).toBe('llama-3.3-70b-versatile');
  });

  it('tolerates a payload with no choices and no usage block', async () => {
    stubFetch(async () => jsonResponse({ model: 'llama-3.3-70b-versatile' }));

    const result = await groqAdapter.complete(MESSAGES, BASE_CONFIG);

    expect(result.content).toBe('');
    expect(result.usage).toEqual({
      promptTokens: undefined,
      completionTokens: undefined,
      totalTokens: undefined,
    });
  });
});

describe('openrouter adapter', () => {
  it('adds attribution headers and hits the vendor endpoint', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(OPENROUTER_FIXTURE));

    const result = await openRouterAdapter.complete(MESSAGES, {
      provider: 'openrouter',
      apiKey: 'sk-or-v1-test',
    });

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(url).toBe('https://openrouter.ai/api/v1/chat/completions');
    const headers = init.headers as Record<string, string>;
    expect(headers['X-Title']).toBe('ExcelAgento');
    expect(headers['HTTP-Referer']).toBe('https://excel-agent.app');
    expect(headers.Authorization).toBe('Bearer sk-or-v1-test');
    expect(result.model).toBe('google/gemini-2.0-flash-001');
    expect(result.usage?.totalTokens).toBe(720);
  });
});

describe('gemini adapter', () => {
  it('sends the key in a header (never the URL), splits the system instruction, and reads usageMetadata', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(GEMINI_FIXTURE));

    const result = await geminiAdapter.complete(MESSAGES, {
      provider: 'gemini',
      apiKey: ' AIza test ',
    });

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(url).toBe(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent',
    );
    expect(url).not.toContain('key=');
    // The key travels in a header so it never lands in URLs/logs.
    expect((init.headers as Record<string, string>)['x-goog-api-key']).toBe('AIza test');

    const body = JSON.parse(String(init.body)) as Record<string, unknown>;
    expect(body.systemInstruction).toEqual({ parts: [{ text: 'SYS' }] });
    expect(body.contents).toEqual([{ role: 'user', parts: [{ text: 'clean the dates' }] }]);
    expect(body.generationConfig).toEqual({ temperature: 0.2, maxOutputTokens: 8192 });

    expect(result.content).toBe('Here is the summary of the worksheet.');
    expect(result.usage).toEqual({ promptTokens: 300, completionTokens: 40, totalTokens: 340 });
  });

  it('uses the configured model in the URL and omits an empty system instruction', async () => {
    const fetchMock = stubFetch(async () => jsonResponse(GEMINI_FIXTURE));

    await geminiAdapter.complete([{ role: 'user', content: 'hi' }], {
      provider: 'gemini',
      apiKey: 'k',
      model: 'gemini-2.5-flash',
    });

    const url = String(fetchMock.mock.calls[0]![0]);
    expect(url).toContain('/models/gemini-2.5-flash:generateContent');
    const body = JSON.parse(String(fetchMock.mock.calls[0]![1].body)) as Record<string, unknown>;
    expect(body.systemInstruction).toBeUndefined();
  });
});

describe('failure handling and retries', () => {
  it('does not retry a 401 and surfaces the provider message', async () => {
    const fetchMock = stubFetch(
      async () => new Response('{"error":{"message":"Invalid API Key"}}', { status: 401 }),
    );

    const error = await groqAdapter.complete(MESSAGES, BASE_CONFIG).catch((e: unknown) => e);

    expect(error).toBeInstanceOf(ProviderError);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const providerError = error as ProviderError;
    expect(providerError.status).toBe(401);
    expect(providerError.retryable).toBe(false);
    expect(providerError.message).toContain('Invalid API Key');
  });

  it('retries a transient 429 and then succeeds', async () => {
    let calls = 0;
    const fetchMock = stubFetch(async () => {
      calls += 1;
      return calls === 1
        ? new Response('rate limited', { status: 429 })
        : jsonResponse(GROQ_FIXTURE);
    });

    const result = await groqAdapter.complete(MESSAGES, BASE_CONFIG);

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(result.content).toBe('I normalised the dates in column C.');
  });

  it('gives up after the configured retries and reports a retryable failure', async () => {
    const fetchMock = stubFetch(async () => new Response('upstream boom', { status: 503 }));

    const error = await groqAdapter
      .complete(MESSAGES, { ...BASE_CONFIG, retries: 1 })
      .catch((e: unknown) => e);

    const providerError = error as ProviderError;
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(providerError).toBeInstanceOf(ProviderError);
    expect(providerError.status).toBe(503);
    expect(providerError.retryable).toBe(true);
  });

  it('retries network failures and reports them as retryable', async () => {
    let calls = 0;
    stubFetch(async () => {
      calls += 1;
      throw new TypeError('network down');
    });

    const error = await groqAdapter
      .complete(MESSAGES, { ...BASE_CONFIG, retries: 1 })
      .catch((e: unknown) => e);

    expect(calls).toBe(2);
    expect(error).toBeInstanceOf(ProviderError);
    expect((error as ProviderError).retryable).toBe(true);
    expect((error as ProviderError).message).toContain('network down');
  });

  it('honours a caller abort immediately and does not retry', async () => {
    const fetchMock = stubFetch(hangingFetch());
    const controller = new AbortController();

    const promise = groqAdapter.complete(MESSAGES, {
      ...BASE_CONFIG,
      retries: 3,
      signal: controller.signal,
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    controller.abort();

    const error = await promise.catch((e: unknown) => e);

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(error).toBeInstanceOf(ProviderError);
    expect((error as ProviderError).message).toContain('cancelled');
    expect((error as ProviderError).retryable).toBe(false);
  });

  it('aborts a hung request using the configured timeout', async () => {
    const fetchMock = stubFetch(hangingFetch());

    const error = await groqAdapter
      .complete(MESSAGES, { ...BASE_CONFIG, timeoutMs: 20, retries: 0 })
      .catch((e: unknown) => e);

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(error).toBeInstanceOf(ProviderError);
  });

  it('rejects a 200 that is not valid JSON with a clear message', async () => {
    stubFetch(async () => new Response('<html>proxy error</html>', { status: 200 }));

    const error = await groqAdapter.complete(MESSAGES, BASE_CONFIG).catch((e: unknown) => e);

    expect(error).toBeInstanceOf(ProviderError);
    expect((error as ProviderError).message).toContain('not valid JSON');
  });
});

describe('provider catalogue', () => {
  it('lists every supported adapter and routes complete() to it', async () => {
    expect(
      listProviders()
        .map((adapter) => adapter.name)
        .sort(),
    ).toEqual(['custom', 'gemini', 'groq', 'openai', 'openrouter']);
    expect(getAdapter('gemini')).toBe(geminiAdapter);

    const fetchMock = stubFetch(async () => jsonResponse(OPENROUTER_FIXTURE));
    await complete(MESSAGES, { provider: 'openrouter', apiKey: 'sk-or-test' });

    expect(String(fetchMock.mock.calls[0]![0])).toContain('openrouter.ai');
  });

  it('exposes a sleep helper', async () => {
    await expect(sleep(1)).resolves.toBeUndefined();
  });
});
