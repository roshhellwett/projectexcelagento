// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';

import {
  DEFAULT_OPENROUTER_MODEL,
  PROVIDER_LABELS,
  clearSettings,
  defaultModelFor,
  isDemoKey,
  loadSettings,
  saveSettings,
} from './settings.js';

/** Storage contract: these keys are read by older builds and by the settings UI. */
const SETTINGS_KEY = 'excel_agent_settings_v2';
const LEGACY_KEY = 'excel_agent_byok_key';
const LEGACY_PROVIDER = 'excel_agent_byok_provider';

beforeEach(() => {
  localStorage.clear();
});

describe('settings store', () => {
  it('returns OpenRouter with default model when nothing is stored', () => {
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: '',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });

  it('round-trips a saved configuration, including a model override', () => {
    saveSettings({
      provider: 'openrouter',
      apiKey: 'sk-or-key',
      model: 'deepseek/deepseek-chat',
    });

    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: 'sk-or-key',
      model: 'deepseek/deepseek-chat',
    });
  });

  it('defaults to Claude 3.5 Sonnet when no model is explicitly provided', () => {
    saveSettings({ provider: 'openrouter', apiKey: 'sk-or-x' });

    const loaded = loadSettings();
    expect(loaded.model).toBe(DEFAULT_OPENROUTER_MODEL);
    expect(loaded).toEqual({
      provider: 'openrouter',
      apiKey: 'sk-or-x',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });

  it('clears mirrored legacy keys instead of duplicating the secret', () => {
    localStorage.setItem(LEGACY_KEY, 'sk-or-legacy');
    localStorage.setItem(LEGACY_PROVIDER, 'openrouter');

    saveSettings({ provider: 'openrouter', apiKey: 'sk-or-new' });

    expect(localStorage.getItem(LEGACY_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_PROVIDER)).toBeNull();
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: 'sk-or-new',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });

  it('migrates a legacy single-key configuration to openrouter', () => {
    localStorage.setItem(LEGACY_KEY, 'sk-or-legacy');
    localStorage.setItem(LEGACY_PROVIDER, 'groq');

    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: 'sk-or-legacy',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });

  it('survives corrupt or structurally invalid stored settings', () => {
    localStorage.setItem(SETTINGS_KEY, '{broken json');
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: '',
      model: DEFAULT_OPENROUTER_MODEL,
    });

    // Present provider but no apiKey is not a usable configuration; fall through.
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ provider: 'gemini' }));
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: '',
      model: DEFAULT_OPENROUTER_MODEL,
    });

    localStorage.setItem(SETTINGS_KEY, 'null');
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: '',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });

  it('clears every key it owns', () => {
    saveSettings({ provider: 'openrouter', apiKey: 'sk-or-x' });
    clearSettings();

    expect(localStorage.getItem(SETTINGS_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_PROVIDER)).toBeNull();
    expect(loadSettings()).toEqual({
      provider: 'openrouter',
      apiKey: '',
      model: DEFAULT_OPENROUTER_MODEL,
    });
  });
});

describe('key and provider helpers', () => {
  it('treats missing, empty, and demo keys as demo mode', () => {
    expect(isDemoKey(undefined)).toBe(true);
    expect(isDemoKey(null)).toBe(true);
    expect(isDemoKey('')).toBe(true);
    expect(isDemoKey('   ')).toBe(true);
    expect(isDemoKey('demo-local-mode')).toBe(true);
    expect(isDemoKey('DEMO-anything')).toBe(true);
  });

  it('treats a real key as active', () => {
    expect(isDemoKey('sk-or-1234567890')).toBe(false);
    expect(isDemoKey('  sk-or-real  ')).toBe(false);
  });

  it('exposes a default model and label for openrouter', () => {
    expect(defaultModelFor('openrouter')).toBe(DEFAULT_OPENROUTER_MODEL);
    expect(PROVIDER_LABELS['openrouter'].length).toBeGreaterThan(0);
  });
});
