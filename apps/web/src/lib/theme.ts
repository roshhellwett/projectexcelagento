/**
 * Workspace theme state.
 *
 * The theme is read from `localStorage` once at start, persisted on every change, and exposed
 * on `<html data-theme>`, where the stylesheet overrides the design tokens. It defaults to the
 * operating system's preference so a first-time user is not forced into a light screen.
 */

export type Theme = 'light' | 'dark';

const STORAGE_KEY = 'excel_agent_theme';

export function getStoredTheme(): Theme | null {
  try {
    if (typeof window === 'undefined') return null;
    const stored = window.localStorage.getItem(STORAGE_KEY);
    return stored === 'dark' || stored === 'light' ? stored : null;
  } catch {
    return null;
  }
}

export function getActiveTheme(): Theme {
  if (typeof document === 'undefined') return 'light';
  return document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
}

export function applyTheme(theme: Theme): void {
  document.documentElement.dataset.theme = theme;
  document.documentElement.style.colorScheme = theme;
  try {
    window.localStorage.setItem(STORAGE_KEY, theme);
  } catch {
    // Private browsing keeps the theme for the session only; that is fine.
  }
}

export function toggleTheme(): Theme {
  const next: Theme = getActiveTheme() === 'dark' ? 'light' : 'dark';
  applyTheme(next);
  return next;
}

/** Called once by the bootstrap script so a saved choice applies on the very first paint. */
export function initThemeFromStorageOrPreference(): Theme {
  const stored = getStoredTheme();
  if (stored) {
    applyTheme(stored);
    return stored;
  }
  const prefersDark =
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-color-scheme: dark)').matches;
  const theme: Theme = prefersDark ? 'dark' : 'light';
  applyTheme(theme);
  return theme;
}
