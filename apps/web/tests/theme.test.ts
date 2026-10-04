// @vitest-environment jsdom
import { describe, expect, it, beforeEach } from 'vitest';
import { getActiveTheme, toggleTheme, applyTheme } from '../src/lib/theme.js';

describe('theme', () => {
  beforeEach(() => {
    window.localStorage.clear();
    document.documentElement.dataset.theme = 'light';
  });

  it('starts from the document theme', () => {
    expect(getActiveTheme()).toBe('light');
  });

  it('toggles to dark and back, persisting each step', () => {
    toggleTheme();
    expect(getActiveTheme()).toBe('dark');
    expect(window.localStorage.getItem('excel_agent_theme')).toBe('dark');

    toggleTheme();
    expect(getActiveTheme()).toBe('light');
    expect(window.localStorage.getItem('excel_agent_theme')).toBe('light');
  });

  it('is settable explicitly', () => {
    applyTheme('dark');
    expect(document.documentElement.dataset.theme).toBe('dark');
    expect(document.documentElement.style.colorScheme).toBe('dark');
  });
});
