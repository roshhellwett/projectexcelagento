import type { ProviderName } from '@excel-agent/agent';

export type { ProviderName };

export interface AgentSettings {
  provider: ProviderName;
  apiKey: string;
  /** Optional explicit model override per provider. */
  model?: string;
  /** Optional custom endpoint base URL for Ollama / LM Studio */
  baseUrl?: string;
}

const SETTINGS_KEY = 'excel_agent_settings_v2';
const LEGACY_KEY = 'excel_agent_byok_key';
const LEGACY_PROVIDER = 'excel_agent_byok_provider';

export const DEFAULT_OPENROUTER_MODEL = 'anthropic/claude-3.5-sonnet';

const DEFAULT_MODELS = {
  openrouter: DEFAULT_OPENROUTER_MODEL,
  groq: 'llama-3.3-70b-versatile',
  gemini: 'gemini-2.0-flash',
  openai: 'gpt-4o-mini',
  custom: 'default',
} as const satisfies Record<ProviderName, string>;

export function defaultModelFor(provider: ProviderName): string {
  return provider === 'openrouter'
    ? DEFAULT_OPENROUTER_MODEL
    : (DEFAULT_MODELS[provider] ?? DEFAULT_OPENROUTER_MODEL);
}

export const PROVIDER_LABELS = {
  openrouter: 'OpenRouter (multi-model)',
  groq: 'Groq (ultra-fast)',
  gemini: 'Google Gemini',
  openai: 'OpenAI (GPT-4o / o3)',
  custom: 'Custom / Local (Ollama, LM Studio)',
} as const satisfies Record<ProviderName, string>;

export const SUGGESTED_OPENROUTER_MODELS = [
  { id: 'anthropic/claude-3.5-sonnet', label: 'Claude 3.5 Sonnet (Recommended)' },
  { id: 'deepseek/deepseek-chat', label: 'DeepSeek V3' },
  { id: 'google/gemini-2.5-pro', label: 'Gemini 2.5 Pro' },
  { id: 'google/gemini-2.0-flash-001', label: 'Gemini 2.0 Flash' },
  { id: 'meta-llama/llama-3.3-70b-instruct', label: 'Llama 3.3 70B' },
];

export const AVAILABLE_MODELS: Record<ProviderName, { id: string; label: string }[]> = {
  openrouter: SUGGESTED_OPENROUTER_MODELS,
  groq: [
    { id: 'openai/gpt-oss-120b', label: 'GPT OSS 120B (OpenAI / Groq) - Recommended' },
    { id: 'llama-3.3-70b-versatile', label: 'Llama 3.3 70B (Legacy)' },
  ],
  gemini: [
    { id: 'gemini-2.5-flash', label: 'Gemini 2.5 Flash' },
    { id: 'gemini-2.0-flash', label: 'Gemini 2.0 Flash' },
  ],
  openai: [
    { id: 'gpt-4o-mini', label: 'GPT-4o mini' },
    { id: 'gpt-4o', label: 'GPT-4o' },
  ],
  custom: [{ id: 'default', label: 'Default / Configured on host' }],
};

export function isDemoKey(apiKey: string | undefined | null): boolean {
  if (!apiKey) return true;
  const normalized = apiKey.trim().toLowerCase();
  return normalized.length === 0 || normalized.startsWith('demo');
}

function storage(): Storage | undefined {
  try {
    return typeof localStorage === 'undefined' ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

/** Load settings, migrating the legacy single-key layout and enforcing OpenRouter as the AI provider. */
export function loadSettings(): AgentSettings {
  const store = storage();
  if (!store) return { provider: 'openrouter', apiKey: '', model: DEFAULT_OPENROUTER_MODEL };
  try {
    const raw = store.getItem(SETTINGS_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<AgentSettings>;
      if (typeof parsed.apiKey === 'string') {
        return {
          provider: 'openrouter',
          apiKey: parsed.apiKey,
          model: parsed.model?.trim() || DEFAULT_OPENROUTER_MODEL,
          ...(parsed.baseUrl ? { baseUrl: parsed.baseUrl } : {}),
        };
      }
    }
    // Legacy keys used by the chat gate: excel_agent_byok_key / _provider.
    const legacyKey = store.getItem(LEGACY_KEY) ?? '';
    return {
      provider: 'openrouter',
      apiKey: legacyKey,
      model: DEFAULT_OPENROUTER_MODEL,
    };
  } catch {
    // Corrupt storage is ignored rather than crashing the workspace.
  }
  return { provider: 'openrouter', apiKey: '', model: DEFAULT_OPENROUTER_MODEL };
}

export function saveSettings(settings: AgentSettings): void {
  const store = storage();
  if (!store) return;
  const sanitized: AgentSettings = {
    ...settings,
    provider: 'openrouter',
    model: settings.model?.trim() || DEFAULT_OPENROUTER_MODEL,
  };
  try {
    store.setItem(SETTINGS_KEY, JSON.stringify(sanitized));
    // Legacy mirrors are migrated on read, so they are cleared rather than duplicated.
    store.removeItem(LEGACY_KEY);
    store.removeItem(LEGACY_PROVIDER);
  } catch {
    // Quota/private-mode failures must not make connecting the local agent crash the workspace.
    // React state still keeps the current session usable; persistence can be retried later.
  }
}

export function clearSettings(): void {
  const store = storage();
  if (!store) return;
  store.removeItem(SETTINGS_KEY);
  store.removeItem(LEGACY_KEY);
  store.removeItem(LEGACY_PROVIDER);
}
