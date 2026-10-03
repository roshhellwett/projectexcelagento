import type {
  ChatMessage,
  ProviderAdapter,
  ProviderConfig,
  ProviderName,
  ProviderResponse,
} from './types.js';

export class ProviderError extends Error {
  readonly provider: ProviderName;
  readonly status: number | undefined;
  readonly retryable: boolean;

  constructor(
    provider: ProviderName,
    message: string,
    options: { status?: number; retryable?: boolean } = {},
  ) {
    super(message);
    this.name = 'ProviderError';
    this.provider = provider;
    this.status = options.status;
    this.retryable = options.retryable ?? false;
  }
}

const DEFAULT_TIMEOUT_MS = 30_000;
const DEFAULT_RETRIES = 2;

function isTransientStatus(status: number): boolean {
  return status === 408 || status === 409 || status === 425 || status === 429 || status >= 500;
}

export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

/**
 * Fetch with an abort-based timeout, merged with an optional caller signal, plus
 * bounded exponential backoff for transient upstream failures. This is the single
 * reliability choke point shared by every provider.
 */
async function requestWithRetry(
  provider: ProviderName,
  url: string,
  init: RequestInit,
  config: ProviderConfig,
): Promise<Response> {
  const retries = Math.max(0, config.retries ?? DEFAULT_RETRIES);
  const timeoutMs = config.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  let lastError: unknown;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const onAbort = () => controller.abort();
    config.signal?.addEventListener('abort', onAbort, { once: true });

    try {
      const response = await fetch(url, { ...init, signal: controller.signal });
      if (!response.ok && isTransientStatus(response.status) && attempt < retries) {
        const body = await response.text().catch(() => '');
        lastError = new ProviderError(
          provider,
          `HTTP ${response.status}: ${body || response.statusText}`,
          { status: response.status, retryable: true },
        );
        await sleep(250 * 2 ** attempt);
        continue;
      }
      return response;
    } catch (error) {
      lastError = error;
      const aborted = error instanceof DOMException && error.name === 'AbortError';
      if (aborted && config.signal?.aborted) {
        throw new ProviderError(provider, 'Request cancelled by caller.', { retryable: false });
      }
      if (attempt < retries) {
        await sleep(250 * 2 ** attempt);
        continue;
      }
    } finally {
      clearTimeout(timer);
      config.signal?.removeEventListener('abort', onAbort);
    }
  }

  if (lastError instanceof ProviderError) {
    throw lastError;
  }
  throw new ProviderError(
    provider,
    lastError instanceof Error ? lastError.message : 'Network request failed.',
    { retryable: true },
  );
}

async function readError(provider: ProviderName, response: Response): Promise<ProviderError> {
  const body = await response.text().catch(() => '');
  const retryable = isTransientStatus(response.status);
  return new ProviderError(
    provider,
    `HTTP ${response.status}: ${body.slice(0, 400) || response.statusText}`,
    { status: response.status, retryable },
  );
}

/**
 * Parse a success response body. A 200 that is not JSON (captive portal, proxy
 * error page, truncated body) must surface as a ProviderError with a usable
 * message rather than a raw SyntaxError.
 */
async function readJson<T>(provider: ProviderName, response: Response): Promise<T> {
  try {
    return (await response.json()) as T;
  } catch {
    throw new ProviderError(provider, 'Provider returned a response that was not valid JSON.', {
      status: response.status,
      retryable: false,
    });
  }
}

type OpenAICompatibleShape = {
  choices?: { message?: { content?: string } }[];
  model?: string;
  usage?: { prompt_tokens?: number; completion_tokens?: number; total_tokens?: number };
};

function openAiCompatibleAdapter(
  name: ProviderName,
  defaultModel: string,
  endpoint: string,
  extraHeaders: () => Record<string, string> = () => ({}),
): ProviderAdapter {
  return {
    name,
    defaultModel,
    async complete(messages: ChatMessage[], config: ProviderConfig): Promise<ProviderResponse> {
      const model = config.model || defaultModel;
      const response = await requestWithRetry(
        name,
        endpoint,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${config.apiKey.trim()}`,
            'Content-Type': 'application/json',
            ...extraHeaders(),
          },
          body: JSON.stringify({
            model,
            messages,
            temperature: config.temperature ?? 0.2,
            max_tokens: config.maxTokens ?? 1200,
          }),
        },
        config,
      );

      if (!response.ok) throw await readError(name, response);
      const data = await readJson<OpenAICompatibleShape>(name, response);
      return {
        content: data.choices?.[0]?.message?.content ?? '',
        provider: name,
        model: data.model ?? model,
        usage: {
          promptTokens: data.usage?.prompt_tokens,
          completionTokens: data.usage?.completion_tokens,
          totalTokens: data.usage?.total_tokens,
        },
      };
    },
  };
}

export const groqAdapter: ProviderAdapter = openAiCompatibleAdapter(
  'groq',
  'llama-3.3-70b-versatile',
  'https://api.groq.com/openai/v1/chat/completions',
);

export const openRouterAdapter: ProviderAdapter = openAiCompatibleAdapter(
  'openrouter',
  'google/gemini-2.0-flash-001',
  'https://openrouter.ai/api/v1/chat/completions',
  () => ({
    'HTTP-Referer':
      typeof window === 'undefined' ? 'https://excel-agent.app' : window.location.origin,
    'X-Title': 'ExcelAgento',
  }),
);

type GeminiShape = {
  candidates?: { content?: { parts?: { text?: string }[] } }[];
  usageMetadata?: {
    promptTokenCount?: number;
    candidatesTokenCount?: number;
    totalTokenCount?: number;
  };
};

export const geminiAdapter: ProviderAdapter = {
  name: 'gemini',
  defaultModel: 'gemini-2.0-flash',
  async complete(messages, config): Promise<ProviderResponse> {
    const model = config.model || this.defaultModel;
    const system = messages
      .filter((message) => message.role === 'system')
      .map((message) => message.content)
      .join('\n\n');
    const conversation = messages
      .filter((message) => message.role !== 'system')
      .map((message) => message.content)
      .join('\n\n');
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(
      config.apiKey.trim(),
    )}`;

    const response = await requestWithRetry(
      'gemini',
      url,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          systemInstruction: system ? { parts: [{ text: system }] } : undefined,
          contents: [{ role: 'user', parts: [{ text: conversation }] }],
          generationConfig: {
            temperature: config.temperature ?? 0.2,
            maxOutputTokens: config.maxTokens ?? 1200,
          },
        }),
      },
      config,
    );

    if (!response.ok) throw await readError('gemini', response);
    const data = await readJson<GeminiShape>('gemini', response);
    return {
      content: data.candidates?.[0]?.content?.parts?.[0]?.text ?? '',
      provider: 'gemini',
      model,
      usage: {
        promptTokens: data.usageMetadata?.promptTokenCount,
        completionTokens: data.usageMetadata?.candidatesTokenCount,
        totalTokens: data.usageMetadata?.totalTokenCount,
      },
    };
  },
};

const ADAPTERS: Record<ProviderName, ProviderAdapter> = {
  groq: groqAdapter,
  openrouter: openRouterAdapter,
  gemini: geminiAdapter,
};

export function getAdapter(provider: ProviderName): ProviderAdapter {
  return ADAPTERS[provider];
}

export function listProviders(): ProviderAdapter[] {
  return Object.values(ADAPTERS);
}

/** Convenience wrapper that resolves the adapter and performs one completion. */
export function complete(
  messages: ChatMessage[],
  config: ProviderConfig,
): Promise<ProviderResponse> {
  return getAdapter(config.provider).complete(messages, config);
}
