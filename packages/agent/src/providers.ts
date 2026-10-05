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

const DEFAULT_TIMEOUT_MS = 120_000;
const DEFAULT_RETRIES = 2;
export const DEFAULT_MAX_TOKENS = 8192;

export function resolveMaxTokens(provider: ProviderName, requested?: number): number {
  if (provider === 'groq') {
    return Math.min(requested ?? 8192, 8192);
  }
  return requested ?? DEFAULT_MAX_TOKENS;
}

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

export function throwIfCancelled(provider: ProviderName, signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw new ProviderError(provider, 'Request cancelled by caller.', { retryable: false });
  }
}

export function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException('Request cancelled by caller.', 'AbortError'));
      return;
    }
    const onAbort = () => {
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      reject(new DOMException('Request cancelled by caller.', 'AbortError'));
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort);
      resolve();
    }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
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
    throwIfCancelled(provider, config.signal);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const onAbort = () => controller.abort();
    config.signal?.addEventListener('abort', onAbort, { once: true });

    try {
      const response = await fetch(url, { ...init, signal: controller.signal });
      throwIfCancelled(provider, config.signal);
      if (!response.ok && isTransientStatus(response.status) && attempt < retries) {
        const body = await response.text().catch(() => '');
        lastError = new ProviderError(
          provider,
          `HTTP ${response.status}: ${body || response.statusText}`,
          { status: response.status, retryable: true },
        );
        await sleep(250 * 2 ** attempt, config.signal);
        continue;
      }
      return response;
    } catch (error) {
      lastError = error;
      if (config.signal?.aborted) {
        throw new ProviderError(provider, 'Request cancelled by caller.', { retryable: false });
      }
      if (attempt < retries) {
        await sleep(250 * 2 ** attempt, config.signal);
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

/** Guards an SSE read against a stalled upstream stream and caller cancellation. */
async function readWithTimeout(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  config: ProviderConfig,
): Promise<ReadableStreamReadResult<Uint8Array>> {
  const timeoutMs = config.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  throwIfCancelled(config.provider, config.signal);
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  try {
    const result = await Promise.race([
      reader.read(),
      new Promise<never>((_resolve, reject) => {
        timeoutId = setTimeout(() => {
          reject(
            new ProviderError(config.provider, `Stream read timed out after ${timeoutMs}ms.`, {
              retryable: true,
            }),
          );
          void reader.cancel().catch(() => {});
        }, timeoutMs);
        onAbort = () => {
          reject(
            new ProviderError(config.provider, 'Request cancelled by caller.', {
              retryable: false,
            }),
          );
          void reader.cancel().catch(() => {});
        };
        config.signal?.addEventListener('abort', onAbort, { once: true });
      }),
    ]);
    throwIfCancelled(config.provider, config.signal);
    return result;
  } finally {
    if (timeoutId !== undefined) clearTimeout(timeoutId);
    if (onAbort) config.signal?.removeEventListener('abort', onAbort);
  }
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

export function extractThoughtAndCleanContent(
  rawContent: string,
  existingThought?: string,
): { content: string; thought?: string } {
  let content = rawContent || '';
  let thought = existingThought || '';

  // Extract <think>...</think>
  const thinkMatch = content.match(/<think>([\s\S]*?)<\/think>/i);
  if (thinkMatch && thinkMatch[1]) {
    thought = (thought ? `${thought}\n` : '') + thinkMatch[1].trim();
    content = content.replace(/<think>[\s\S]*?<\/think>/gi, '').trim();
  }

  // Extract |<|minimax|>| or |<|...|>| tags
  if (/\|<\|[a-zA-Z0-9_-]+\|>\|/.test(content)) {
    const parts = content.split(/\|<\|[a-zA-Z0-9_-]+\|>\|/);
    const reasoningText = parts
      .filter((p) => p.trim().length > 0)
      .join('\n')
      .trim();
    thought = (thought ? `${thought}\n` : '') + reasoningText;
    content = '';
  }

  content = content.replace(/<\/?think>/gi, '').trim();

  return {
    content,
    thought: thought || undefined,
  };
}

function formatMessagesForOpenAI(messages: ChatMessage[]) {
  return messages.map((m) => {
    let cleanContent = m.content;
    if (typeof cleanContent === 'string') {
      cleanContent = cleanContent
        .replace(/\|<\|[a-zA-Z0-9_-]+\|>\|/g, '')
        .replace(/<think>[\s\S]*?<\/think>/g, '')
        .replace(/<\/?think>/g, '')
        .trim();
    }
    const msg: Record<string, unknown> = {
      role: m.role,
      content: cleanContent,
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
        max_tokens: resolveMaxTokens(name, config.maxTokens),
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
      const rawContent = messageObj?.content ?? '';
      const rawThought = messageObj?.reasoning_content ?? messageObj?.thought ?? undefined;
      const cleaned = extractThoughtAndCleanContent(rawContent, rawThought);

      const toolCalls: ToolCall[] | undefined = messageObj?.tool_calls?.map((tc, idx) => ({
        id: tc.id || `call-${Date.now()}-${idx}`,
        type: 'function',
        function: {
          name: tc.function?.name ?? '',
          arguments: tc.function?.arguments ?? '{}',
        },
      }));

      return {
        content: cleaned.content,
        thought: cleaned.thought,
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
        max_tokens: resolveMaxTokens(name, config.maxTokens),
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
        { promptTokens?: number; completionTokens?: number; totalTokens?: number } | undefined;
      const toolCallMap = new Map<number, { id: string; name: string; args: string }>();

      // Real-time live token tracking for prompt and streaming tokens
      const estimatedPromptTokens = Math.max(
        15,
        Math.round(messages.reduce((acc, m) => acc + (m.content?.length || 0), 0) / 3.8),
      );
      let streamedCompletionTokens = 0;
      callbacks.onTokenCount?.({
        promptTokens: estimatedPromptTokens,
        completionTokens: 0,
        totalTokens: estimatedPromptTokens,
      });

      try {
        while (true) {
          const { done, value } = await readWithTimeout(reader, config);
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
                usage?: {
                  prompt_tokens?: number;
                  completion_tokens?: number;
                  total_tokens?: number;
                };
              };

              if (parsed.model) reportedModel = parsed.model;
              if (parsed.usage) {
                reportedUsage = {
                  promptTokens: parsed.usage.prompt_tokens,
                  completionTokens: parsed.usage.completion_tokens,
                  totalTokens: parsed.usage.total_tokens,
                };
                callbacks.onTokenCount?.(reportedUsage);
              }
              const delta = parsed.choices?.[0]?.delta;
              if (delta) {
                let chunkChars = 0;
                if (delta.content) {
                  fullContent += delta.content;
                  callbacks.onToken?.(delta.content);
                  chunkChars += delta.content.length;
                }
                const thoughtToken = delta.reasoning_content ?? delta.thought;
                if (thoughtToken) {
                  fullThought += thoughtToken;
                  callbacks.onThinking?.(thoughtToken);
                  chunkChars += thoughtToken.length;
                }
                if (chunkChars > 0 && !reportedUsage) {
                  streamedCompletionTokens += Math.max(1, Math.round(chunkChars / 3.8));
                  callbacks.onTokenCount?.({
                    promptTokens: estimatedPromptTokens,
                    completionTokens: streamedCompletionTokens,
                    totalTokens: estimatedPromptTokens + streamedCompletionTokens,
                  });
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
        .map(([, tc]) => ({
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

      const cleaned = extractThoughtAndCleanContent(fullContent, fullThought);

      return {
        content: cleaned.content,
        thought: cleaned.thought,
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

function formatMessagesForGemini(messages: ChatMessage[]) {
  const names = new Map<string, string>();
  const contents: { role: 'model' | 'user'; parts: Record<string, unknown>[] }[] = [];
  let previousWasTool = false;
  for (const message of messages) {
    if (message.role === 'system') continue;
    if (message.role === 'tool') {
      const name = message.name || names.get(message.tool_call_id ?? '');
      if (!name) throw new ProviderError('gemini', 'Tool result is missing its function name.');
      let output: unknown = message.content;
      try {
        output = JSON.parse(message.content);
      } catch {
        /* Keep the untrusted-data wrapper intact. */
      }
      const part = { functionResponse: { name, response: { output } } };
      if (previousWasTool) contents[contents.length - 1]!.parts.push(part);
      else contents.push({ role: 'user', parts: [part] });
      previousWasTool = true;
      continue;
    }
    const parts: Record<string, unknown>[] = [];
    if (message.content) parts.push({ text: message.content });
    for (const call of message.tool_calls ?? []) {
      let args: unknown;
      try {
        args = JSON.parse(call.function.arguments || '{}');
      } catch {
        throw new ProviderError('gemini', 'Function-call arguments are not valid JSON.');
      }
      names.set(call.id, call.function.name);
      parts.push({
        functionCall: { name: call.function.name, args },
        ...(call.thoughtSignature ? { thoughtSignature: call.thoughtSignature } : {}),
      });
    }
    contents.push({
      role: message.role === 'assistant' ? 'model' : 'user',
      parts: parts.length ? parts : [{ text: '' }],
    });
    previousWasTool = false;
  }
  return contents;
}

type GeminiShape = {
  candidates?: {
    content?: {
      parts?: {
        text?: string;
        thoughtSignature?: string;
        functionCall?: { name?: string; args?: unknown };
      }[];
    };
  }[];
  usageMetadata?: {
    promptTokenCount?: number;
    candidatesTokenCount?: number;
    totalTokenCount?: number;
  };
};

export const geminiAdapter: ProviderAdapter = {
  name: 'gemini',
  defaultModel: 'gemini-2.0-flash',
  async complete(messages, config, tools): Promise<ProviderResponse> {
    const model = config.model || this.defaultModel;
    const system = messages
      .filter((message) => message.role === 'system')
      .map((message) => message.content)
      .join('\n\n');
    const contents = formatMessagesForGemini(messages);
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;

    const geminiTools =
      tools && tools.length > 0
        ? [
            {
              functionDeclarations: tools.map((tool) => ({
                name: tool.function.name,
                description: tool.function.description,
                parameters: tool.function.parameters,
              })),
            },
          ]
        : undefined;

    const response = await requestWithRetry(
      'gemini',
      url,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          // Key in header, never in the URL (logged by proxies/tooling)
          'x-goog-api-key': config.apiKey.trim(),
        },
        body: JSON.stringify({
          systemInstruction: system ? { parts: [{ text: system }] } : undefined,
          contents: contents.length > 0 ? contents : [{ role: 'user', parts: [{ text: '' }] }],
          tools: geminiTools,
          generationConfig: {
            temperature: config.temperature ?? 0.2,
            maxOutputTokens: resolveMaxTokens('gemini', config.maxTokens),
          },
        }),
      },
      config,
    );

    if (!response.ok) throw await readError('gemini', response);
    const data = await readJson<GeminiShape>('gemini', response);
    const parts = data.candidates?.[0]?.content?.parts ?? [];
    const text = parts.map((part) => (part as { text?: string }).text ?? '').join('');
    const toolCalls: ToolCall[] = [];
    for (const part of parts) {
      const fc = (part as { functionCall?: { name?: string; args?: unknown } }).functionCall;
      if (fc?.name) {
        toolCalls.push({
          id: `gemini-call-${toolCalls.length}`,
          type: 'function',
          function: { name: fc.name, arguments: JSON.stringify(fc.args ?? {}) },
          ...(part.thoughtSignature ? { thoughtSignature: part.thoughtSignature } : {}),
        });
      }
    }
    return {
      content: text,
      provider: 'gemini',
      model,
      toolCalls: toolCalls.length > 0 ? toolCalls : undefined,
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
