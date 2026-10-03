// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';

import {
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
  it('returns Groq with no key when nothing is stored', () => {
    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });
  });

  it('round-trips a saved configuration, including a model override', () => {
    saveSettings({ provider: 'gemini', apiKey: 'AIza-key', model: 'gemini-2.5-flash' });

    expect(loadSettings()).toEqual({
      provider: 'gemini',
      apiKey: 'AIza-key',
      model: 'gemini-2.5-flash',
    });
  });

  it('omits the model field when no override is configured', () => {
    saveSettings({ provider: 'groq', apiKey: 'gsk_x' });

    const loaded = loadSettings();
    expect(loaded.model).toBeUndefined();
    expect(loaded).toEqual({ provider: 'groq', apiKey: 'gsk_x' });
  });

  it('clears mirrored legacy keys instead of duplicating the secret', () => {
    localStorage.setItem(LEGACY_KEY, 'gsk_legacy');
    localStorage.setItem(LEGACY_PROVIDER, 'groq');

    saveSettings({ provider: 'openrouter', apiKey: 'sk-or-x' });

    expect(localStorage.getItem(LEGACY_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_PROVIDER)).toBeNull();
    expect(loadSettings()).toEqual({ provider: 'openrouter', apiKey: 'sk-or-x' });
  });

  it('migrates a legacy single-key configuration', () => {
    localStorage.setItem(LEGACY_KEY, 'gsk_legacy');
    localStorage.setItem(LEGACY_PROVIDER, 'gemini');

    expect(loadSettings()).toEqual({ provider: 'gemini', apiKey: 'gsk_legacy' });
  });

  it('ignores an unknown legacy provider instead of trusting it', () => {
    localStorage.setItem(LEGACY_KEY, 'gsk_legacy');
    localStorage.setItem(LEGACY_PROVIDER, 'not-a-provider');

    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });
  });

  it('survives corrupt or structurally invalid stored settings', () => {
    localStorage.setItem(SETTINGS_KEY, '{broken json');
    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });

    // Present provider but no apiKey is not a usable configuration; fall through.
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ provider: 'gemini' }));
    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });

    localStorage.setItem(SETTINGS_KEY, 'null');
    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });
  });

  it('clears every key it owns', () => {
    saveSettings({ provider: 'groq', apiKey: 'gsk_x' });
    clearSettings();

    expect(localStorage.getItem(SETTINGS_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_KEY)).toBeNull();
    expect(localStorage.getItem(LEGACY_PROVIDER)).toBeNull();
    expect(loadSettings()).toEqual({ provider: 'groq', apiKey: '' });
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
    expect(isDemoKey('gsk_1234567890')).toBe(false);
    expect(isDemoKey('  AIzaSyReal  ')).toBe(false);
  });

  it('exposes a default model and label for every provider', () => {
    for (const provider of ['groq', 'openrouter', 'gemini'] as const) {
      expect(defaultModelFor(provider).length).toBeGreaterThan(0);
      expect(PROVIDER_LABELS[provider].length).toBeGreaterThan(0);
    }
    expect(defaultModelFor('groq')).toBe('llama-3.3-70b-versatile');
  });
});
