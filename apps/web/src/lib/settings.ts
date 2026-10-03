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

const DEFAULT_MODELS = {
  groq: 'llama-3.3-70b-versatile',
  openrouter: 'google/gemini-2.0-flash-001',
  gemini: 'gemini-2.0-flash',
  openai: 'gpt-4o-mini',
  custom: 'default',
} as const satisfies Record<ProviderName, string>;

export function defaultModelFor(provider: ProviderName): string {
  return DEFAULT_MODELS[provider] ?? 'default';
}

export const PROVIDER_LABELS = {
  groq: 'Groq (ultra-fast)',
  gemini: 'Google Gemini',
  openrouter: 'OpenRouter (multi-model)',
  openai: 'OpenAI (GPT-4o / o3)',
  custom: 'Custom / Local (Ollama, LM Studio)',
} as const satisfies Record<ProviderName, string>;

export const AVAILABLE_MODELS: Record<ProviderName, { id: string; label: string }[]> = {
  groq: [
    { id: 'openai/gpt-oss-120b', label: 'GPT OSS 120B (OpenAI / Groq) - Recommended' },
    { id: 'openai/gpt-oss-20b', label: 'GPT OSS 20B (High Speed)' },
    { id: 'qwen/qwen3.8-27b', label: 'Qwen 3.8 27B' },
    { id: 'allam-2-7b', label: 'Allam 2 7B' },
    { id: 'llama-3.3-70b-versatile', label: 'Llama 3.3 70B (Legacy)' },
  ],
  gemini: [
    { id: 'gemini-2.5-flash', label: 'Gemini 2.5 Flash (Ultra Fast & Smart)' },
    { id: 'gemini-2.5-pro', label: 'Gemini 2.5 Pro (Deep Reasoning)' },
    { id: 'gemini-2.0-flash', label: 'Gemini 2.0 Flash' },
  ],
  openrouter: [
    { id: 'google/gemini-2.0-flash-001', label: 'Gemini 2.0 Flash' },
    { id: 'meta-llama/llama-3.3-70b-instruct', label: 'Llama 3.3 70B Instruct' },
    { id: 'anthropic/claude-3.5-sonnet', label: 'Claude 3.5 Sonnet' },
    { id: 'deepseek/deepseek-chat', label: 'DeepSeek V3' },
  ],
  openai: [
    { id: 'gpt-4o-mini', label: 'GPT-4o mini (Fast & Cost Effective)' },
    { id: 'gpt-4o', label: 'GPT-4o (Flagship Omni)' },
    { id: 'o3-mini', label: 'o3-mini (High Reasoning)' },
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

/** Load settings, migrating the legacy single-key layout when present. */
export function loadSettings(): AgentSettings {
  const store = storage();
  if (!store) return { provider: 'groq', apiKey: '' };
  try {
    const raw = store.getItem(SETTINGS_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<AgentSettings>;
      if (parsed.provider && typeof parsed.apiKey === 'string') {
        return {
          provider: parsed.provider,
          apiKey: parsed.apiKey,
          ...(parsed.model ? { model: parsed.model } : {}),
          ...(parsed.baseUrl ? { baseUrl: parsed.baseUrl } : {}),
        };
      }
    }
    // Legacy keys used by the chat gate: excel_agent_byok_key / _provider.
    const legacyKey = store.getItem(LEGACY_KEY) ?? '';
    const legacyProvider = store.getItem(LEGACY_PROVIDER);
    if (
      legacyProvider === 'groq' ||
      legacyProvider === 'openrouter' ||
      legacyProvider === 'gemini' ||
      legacyProvider === 'openai' ||
      legacyProvider === 'custom'
    ) {
      return { provider: legacyProvider, apiKey: legacyKey };
    }
  } catch {
    // Corrupt storage is ignored rather than crashing the workspace.
  }
  return { provider: 'groq', apiKey: '' };
}

export function saveSettings(settings: AgentSettings): void {
  const store = storage();
  if (!store) return;
  store.setItem(SETTINGS_KEY, JSON.stringify(settings));
  // Mirror the legacy keys so older builds and the chat gate stay in sync.
  store.setItem(LEGACY_KEY, settings.apiKey);
  store.setItem(LEGACY_PROVIDER, settings.provider);
}

export function clearSettings(): void {
  const store = storage();
  if (!store) return;
  store.removeItem(SETTINGS_KEY);
  store.removeItem(LEGACY_KEY);
  store.removeItem(LEGACY_PROVIDER);
}
