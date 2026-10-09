// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { writeClipboardText } from './clipboard.js';

afterEach(() => {
  vi.restoreAllMocks();
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
});

describe('writeClipboardText', () => {
  it('reports provider success only after the clipboard promise resolves', async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
    await expect(writeClipboardText('copied')).resolves.toBe(true);
    expect(writeText).toHaveBeenCalledWith('copied');
  });

  it('uses the document fallback when the async clipboard API is unavailable', async () => {
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: vi.fn().mockReturnValue(true),
    });
    await expect(writeClipboardText('fallback')).resolves.toBe(true);
  });

  it('returns false instead of throwing when browser copy is blocked', async () => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: vi.fn().mockRejectedValue(new Error('blocked')) },
    });
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: vi.fn().mockReturnValue(false),
    });
    await expect(writeClipboardText('blocked')).resolves.toBe(false);
  });

  it('tries the legacy fallback after an async clipboard rejection', async () => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: vi.fn().mockRejectedValue(new Error('permission changed')) },
    });
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: vi.fn().mockReturnValue(true),
    });
    await expect(writeClipboardText('fallback-after-rejection')).resolves.toBe(true);
  });
});
