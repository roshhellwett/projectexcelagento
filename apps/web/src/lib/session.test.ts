import { beforeEach, describe, expect, it, vi } from 'vitest';
import { clearCurrentTabWorkingMemory, getTabSessionId } from './session';
import { memory } from './agent-runtime';

describe('Tab Session Isolation', () => {
  beforeEach(() => {
    sessionStorage.clear();
    vi.restoreAllMocks();
  });

  it('generates a stable cryptographic UUID for the current tab session', () => {
    const id1 = getTabSessionId();
    expect(id1).toBeDefined();
    expect(typeof id1).toBe('string');
    expect(id1.length).toBeGreaterThanOrEqual(16);

    // Subsequent calls within the same tab return the identical session ID
    const id2 = getTabSessionId();
    expect(id2).toBe(id1);
  });

  it('generates distinct session IDs across different browser sessions', () => {
    const id1 = getTabSessionId();
    sessionStorage.clear();
    const id2 = getTabSessionId();

    expect(id1).not.toBe(id2);
  });

  it('invokes working memory clear on tab session cleanup', async () => {
    const clearSpy = vi.spyOn(memory, 'clearWorkingMemory').mockResolvedValue(true);
    const success = await clearCurrentTabWorkingMemory();
    expect(success).toBe(true);
    expect(clearSpy).toHaveBeenCalledWith(getTabSessionId());
  });
});
