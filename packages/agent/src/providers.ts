import type {
  ChatMessage,
  ProviderAdapter,
  ProviderConfig,
  ProviderName,
  ProviderResponse,
  StreamCallbacks,
  ToolCall,
  ToolDefinition,
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

const DEFAULT_TIMEOUT_MS = 45_000;
const DEFAULT_RETRIES = 2;

export const FALLBACK_MODELS: Record<ProviderName, string[]> = {
  groq: [
    'openai/gpt-oss-120b',
    'openai/gpt-oss-20b',
    'qwen/qwen3.8-27b',
    'allam-2-7b',
    'llama-3.3-70b-versatile',
  ],
  gemini: ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash'],
  openrouter: [
    'google/gemini-2.0-flash-001',
    'meta-llama/llama-3.3-70b-instruct',
    'deepseek/deepseek-chat',
  ],
  openai: ['gpt-4o-mini', 'gpt-4o'],
  custom: ['default'],
};

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
 * bounded exponential backoff for transient upstream failures.
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

type OpenAIToolCallShape = {
  id?: string;
  type?: 'function';
  function?: { name?: string; arguments?: string };
};

type OpenAICompatibleShape = {
  choices?: {
    message?: {
      content?: string;
      reasoning_content?: string;
      thought?: string;
      tool_calls?: OpenAIToolCallShape[];
    };
  }[];
  model?: string;
  usage?: { prompt_tokens?: number; completion_tokens?: number; total_tokens?: number };
};

function formatMessagesForOpenAI(messages: ChatMessage[]) {
  return messages.map((m) => {
    const msg: Record<string, unknown> = {
      role: m.role,
      content: m.content,
    };
    if (m.name) msg.name = m.name;
    if (m.tool_call_id) msg.tool_call_id = m.tool_call_id;
    if (m.tool_calls && m.tool_calls.length > 0) msg.tool_calls = m.tool_calls;
    return msg;
  });
}

function openAiCompatibleAdapter(
  name: ProviderName,
  defaultModel: string,
  endpointOrResolver: string | ((config: ProviderConfig) => string),
  extraHeaders: () => Record<string, string> = () => ({}),
): ProviderAdapter {
  const resolveEndpoint = (config: ProviderConfig) =>
    typeof endpointOrResolver === 'function' ? endpointOrResolver(config) : endpointOrResolver;

  return {
    name,
    defaultModel,
    async complete(
      messages: ChatMessage[],
      config: ProviderConfig,
      tools?: ToolDefinition[],
    ): Promise<ProviderResponse> {
      const model = config.model || defaultModel;
      const endpoint = resolveEndpoint(config);

      const requestBody: Record<string, unknown> = {
        model,
        messages: formatMessagesForOpenAI(messages),
        temperature: config.temperature ?? 0.2,
        max_tokens: config.maxTokens ?? 1200,
      };
      if (tools && tools.length > 0) {
        requestBody.tools = tools;
      }

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
          body: JSON.stringify(requestBody),
        },
        config,
      );

      if (!response.ok) throw await readError(name, response);
      const data = await readJson<OpenAICompatibleShape>(name, response);

      const messageObj = data.choices?.[0]?.message;
      const content = messageObj?.content ?? '';
      const thought = messageObj?.reasoning_content ?? messageObj?.thought ?? undefined;

      const toolCalls: ToolCall[] | undefined = messageObj?.tool_calls?.map((tc, idx) => ({
        id: tc.id || `call-${Date.now()}-${idx}`,
        type: 'function',
        function: {
          name: tc.function?.name ?? '',
          arguments: tc.function?.arguments ?? '{}',
        },
      }));

      return {
        content,
        thought,
        toolCalls: toolCalls && toolCalls.length > 0 ? toolCalls : undefined,
        provider: name,
        model: data.model ?? model,
        usage: {
          promptTokens: data.usage?.prompt_tokens,
          completionTokens: data.usage?.completion_tokens,
          totalTokens: data.usage?.total_tokens,
        },
      };
    },

    async completeStream(
      messages: ChatMessage[],
      config: ProviderConfig,
      callbacks: StreamCallbacks,
      tools?: ToolDefinition[],
    ): Promise<ProviderResponse> {
      const model = config.model || defaultModel;
      const endpoint = resolveEndpoint(config);

      const requestBody: Record<string, unknown> = {
        model,
        messages: formatMessagesForOpenAI(messages),
        temperature: config.temperature ?? 0.2,
        max_tokens: config.maxTokens ?? 1200,
        stream: true,
        stream_options: { include_usage: true },
      };
      if (tools && tools.length > 0) {
        requestBody.tools = tools;
      }

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
          body: JSON.stringify(requestBody),
        },
        config,
      );

      if (!response.ok) throw await readError(name, response);

      const contentType = response.headers.get('content-type') || '';
      if (!contentType.includes('text/event-stream') && contentType.includes('application/json')) {
        const data = await readJson<OpenAICompatibleShape>(name, response);
        const messageObj = data.choices?.[0]?.message;
        const content = messageObj?.content ?? '';
        const thought = messageObj?.reasoning_content ?? messageObj?.thought ?? undefined;
        if (content) callbacks.onToken?.(content);
        if (thought) callbacks.onThinking?.(thought);

        const toolCalls: ToolCall[] | undefined = messageObj?.tool_calls?.map((tc, idx) => ({
          id: tc.id || `call-${Date.now()}-${idx}`,
          type: 'function',
          function: {
            name: tc.function?.name ?? '',
            arguments: tc.function?.arguments ?? '{}',
          },
        }));

        if (toolCalls) {
          for (const tc of toolCalls) callbacks.onToolCall?.(tc);
        }

        return {
          content,
          thought,
          toolCalls: toolCalls && toolCalls.length > 0 ? toolCalls : undefined,
          provider: name,
          model: data.model ?? model,
          usage: {
            promptTokens: data.usage?.prompt_tokens,
            completionTokens: data.usage?.completion_tokens,
            totalTokens: data.usage?.total_tokens,
          },
        };
      }

      if (!response.body) {
        return this.complete(messages, config, tools);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      let fullContent = '';
      let fullThought = '';
      let reportedModel = model;
      let reportedUsage:
        | { promptTokens?: number; completionTokens?: number; totalTokens?: number }
        | undefined;
      const toolCallMap = new Map<number, { id: string; name: string; args: string }>();

      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });

          const lines = buffer.split('\n');
          buffer = lines.pop() ?? '';

          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed.startsWith('data:')) continue;
            const payload = trimmed.slice(5).trim();
            if (payload === '[DONE]') continue;

            try {
              const parsed = JSON.parse(payload) as {
                model?: string;
                choices?: {
                  delta?: {
                    content?: string;
                    reasoning_content?: string;
                    thought?: string;
                    tool_calls?: {
                      index?: number;
                      id?: string;
                      function?: { name?: string; arguments?: string };
                    }[];
                  };
                }[];
                usage?: { prompt_tokens?: number; completion_tokens?: number; total_tokens?: number };
              };

              if (parsed.model) reportedModel = parsed.model;
              if (parsed.usage) {
                reportedUsage = {
                  promptTokens: parsed.usage.prompt_tokens,
                  completionTokens: parsed.usage.completion_tokens,
                  totalTokens: parsed.usage.total_tokens,
                };
              }
              const delta = parsed.choices?.[0]?.delta;
              if (delta) {
                if (delta.content) {
                  fullContent += delta.content;
                  callbacks.onToken?.(delta.content);
                }
                const thoughtToken = delta.reasoning_content ?? delta.thought;
                if (thoughtToken) {
                  fullThought += thoughtToken;
                  callbacks.onThinking?.(thoughtToken);
                }
                if (delta.tool_calls) {
                  for (const tc of delta.tool_calls) {
                    const idx = tc.index ?? 0;
                    const existing = toolCallMap.get(idx) ?? {
                      id: tc.id || `call-${Date.now()}-${idx}`,
                      name: '',
                      args: '',
                    };
                    if (tc.id) existing.id = tc.id;
                    if (tc.function?.name) existing.name += tc.function.name;
                    if (tc.function?.arguments) existing.args += tc.function.arguments;
                    toolCallMap.set(idx, existing);
                  }
                }
              }
            } catch {
              // Ignore non-json sse chunks
            }
          }
        }
      } finally {
        reader.releaseLock();
      }

      const toolCalls: ToolCall[] = Array.from(toolCallMap.entries())
        .sort(([a], [b]) => a - b)
        .map(([_, tc]) => ({
          id: tc.id,
          type: 'function',
          function: {
            name: tc.name,
            arguments: tc.args || '{}',
          },
        }));

      for (const tc of toolCalls) {
        callbacks.onToolCall?.(tc);
      }

      return {
        content: fullContent,
        thought: fullThought || undefined,
        toolCalls: toolCalls.length > 0 ? toolCalls : undefined,
        provider: name,
        model: reportedModel,
        usage: reportedUsage,
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

export const openAiAdapter: ProviderAdapter = openAiCompatibleAdapter(
  'openai',
  'gpt-4o-mini',
  'https://api.openai.com/v1/chat/completions',
);

export const customAdapter: ProviderAdapter = openAiCompatibleAdapter(
  'custom',
  'default',
  (config) => {
    const base = config.baseUrl?.replace(/\/+$/, '') || 'http://localhost:11434/v1';
    return `${base}/chat/completions`;
  },
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

  async completeStream(messages, config, callbacks, tools): Promise<ProviderResponse> {
    const res = await this.complete(messages, config, tools);
    if (res.content) {
      callbacks.onToken?.(res.content);
    }
    return res;
  },
};

const ADAPTERS: Record<ProviderName, ProviderAdapter> = {
  groq: groqAdapter,
  openrouter: openRouterAdapter,
  gemini: geminiAdapter,
  openai: openAiAdapter,
  custom: customAdapter,
};

export function getAdapter(provider: ProviderName): ProviderAdapter {
  return ADAPTERS[provider] ?? groqAdapter;
}

export function listProviders(): ProviderAdapter[] {
  return Object.values(ADAPTERS);
}

/** Convenience wrapper that resolves the adapter and performs one completion. */
export function complete(
  messages: ChatMessage[],
  config: ProviderConfig,
  tools?: ToolDefinition[],
): Promise<ProviderResponse> {
  return getAdapter(config.provider).complete(messages, config, tools);
}

/** Streaming completion wrapper. */
export function completeStream(
  messages: ChatMessage[],
  config: ProviderConfig,
  callbacks: StreamCallbacks,
  tools?: ToolDefinition[],
): Promise<ProviderResponse> {
  const adapter = getAdapter(config.provider);
  if (adapter.completeStream) {
    return adapter.completeStream(messages, config, callbacks, tools);
  }
  return adapter.complete(messages, config, tools).then((res) => {
    if (res.content) callbacks.onToken?.(res.content);
    return res;
  });
}
