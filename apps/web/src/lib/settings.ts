import type { ProviderName } from '@excel-agent/agent';

export type { ProviderName };

export interface AgentSettings {
  provider: ProviderName;
  apiKey: string;
  /** Optional explicit model override per provider. */
  model?: string;
}

const SETTINGS_KEY = 'excel_agent_settings_v2';
const LEGACY_KEY = 'excel_agent_byok_key';
const LEGACY_PROVIDER = 'excel_agent_byok_provider';

const DEFAULT_MODELS = {
  groq: 'llama-3.3-70b-versatile',
  openrouter: 'google/gemini-2.0-flash-001',
  gemini: 'gemini-2.0-flash',
} as const satisfies Record<ProviderName, string>;

export function defaultModelFor(provider: ProviderName): string {
  return DEFAULT_MODELS[provider];
}

export const PROVIDER_LABELS = {
  groq: 'Groq (ultra-fast)',
  openrouter: 'OpenRouter (multi-model)',
  gemini: 'Google Gemini',
} as const satisfies Record<ProviderName, string>;

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
        };
      }
    }
    // Legacy keys used by the chat gate: excel_agent_byok_key / _provider.
    const legacyKey = store.getItem(LEGACY_KEY) ?? '';
    const legacyProvider = store.getItem(LEGACY_PROVIDER);
    if (
      legacyProvider === 'groq' ||
      legacyProvider === 'openrouter' ||
      legacyProvider === 'gemini'
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
