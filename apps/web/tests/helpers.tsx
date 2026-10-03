import { render, screen, type RenderResult } from '@testing-library/react';
import type { UserEvent } from '@testing-library/user-event';

import { App } from '../src/App.js';
import { forgetLearnedActions } from '../src/lib/agent-runtime.js';
import { saveSettings } from '../src/lib/settings.js';
import { clearUsageLog } from '../src/lib/usage.js';

let createdBlobs: Blob[] = [];
const revokedUrls: string[] = [];
let installed = false;

/**
 * jsdom implements neither object URLs nor anchor navigation, so both are stubbed
 * once per test file. Captured blobs let the export test assert real bytes.
 */
export function installBrowserStubs(): void {
  if (installed) return;
  installed = true;
  URL.createObjectURL = (blob: Blob) => {
    createdBlobs.push(blob);
    return `blob:excel-agent/${createdBlobs.length}`;
  };
  URL.revokeObjectURL = (url: string) => {
    revokedUrls.push(url);
  };
  HTMLAnchorElement.prototype.click = function click() {
    // Navigation is meaningless in jsdom; the download path stops here.
  };
}

export function createdBlobsSnapshot(): Blob[] {
  return [...createdBlobs];
}

export function revokedUrlsSnapshot(): string[] {
  return [...revokedUrls];
}

/** Reset every piece of persisted and module-level state between tests. */
export function resetAppState(): void {
  localStorage.clear();
  forgetLearnedActions();
  clearUsageLog();
  createdBlobs = [];
  revokedUrls.length = 0;
}

export function renderApp(): RenderResult {
  return render(<App />);
}

export function metaPillText(container: HTMLElement): string {
  return container.querySelector('.file-meta-pill')?.textContent ?? '';
}

export function sheetTabTexts(container: HTMLElement): string[] {
  return [...container.querySelectorAll('.sheet-tab')].map((tab) => tab.textContent ?? '');
}

export function chatInput(): HTMLInputElement {
  return screen.getByPlaceholderText(/Ask ExcelAgento/i) as HTMLInputElement;
}

export async function askAgent(user: UserEvent, prompt: string): Promise<void> {
  await user.type(chatInput(), prompt);
  await user.click(screen.getByRole('button', { name: /^Send$/i }));
}

export function enterDemoMode(): void {
  saveSettings({ provider: 'groq', apiKey: 'demo-local-mode' });
}

export function enterLiveMode(): void {
  saveSettings({ provider: 'groq', apiKey: 'gsk_live_test_key' });
}

export function fileInput(container: HTMLElement): HTMLInputElement {
  const input = container.querySelector('input[type="file"]');
  if (!input) throw new Error('Expected a file input in the workspace shell.');
  return input as HTMLInputElement;
}

/** A cheap File whose reported size is spoofed, for the upload guard test. */
export function fileWithSize(name: string, size: number): File {
  const file = new File(['x'], name, { type: 'application/vnd.ms-excel' });
  Object.defineProperty(file, 'size', { value: size });
  return file;
}

export function toastTexts(container: HTMLElement): string[] {
  return [...container.querySelectorAll('.toast')].map((toast) => toast.textContent ?? '');
}
