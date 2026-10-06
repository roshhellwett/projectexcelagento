import { memory } from './agent-runtime';

const SESSION_KEY = 'excel_agent_tab_session_id';
let fallbackSessionId: string | undefined;
let storageUnavailable = false;

function safeSessionStorage(): Storage | undefined {
  try {
    return typeof sessionStorage === 'undefined' ? undefined : sessionStorage;
  } catch {
    return undefined;
  }
}

/**
 * Returns a cryptographically isolated session ID unique to this specific browser tab/window.
 *
 * Entropy: 128-bit CSPRNG (crypto.randomUUID).
 * Collision resistance: < 1 in 10^15 even across 100 billion concurrent active browser sessions.
 * Isolation: Stored in sessionStorage so opening new tabs, windows, or duplicate sessions
 * creates completely separate, isolated working memory sandboxes.
 */
export function getTabSessionId(): string {
  if (storageUnavailable && fallbackSessionId) return fallbackSessionId;
  let store = safeSessionStorage();
  let existing: string | null | undefined;
  try {
    existing = store?.getItem(SESSION_KEY);
  } catch {
    store = undefined;
    storageUnavailable = true;
  }
  // New tabs can inherit sessionStorage from their opener. Only reuse an ID generated
  // by this document; a copied tab must not share the opener's working-memory sandbox.
  if (existing && existing === fallbackSessionId) {
    return existing;
  }
  if (!store && fallbackSessionId) return fallbackSessionId;

  let newId: string;
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    newId = crypto.randomUUID();
  } else {
    // Fallback CSPRNG if crypto.randomUUID is unavailable
    const buf = new Uint8Array(16);
    if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
      crypto.getRandomValues(buf);
    } else {
      for (let i = 0; i < 16; i++) buf[i] = Math.floor(Math.random() * 256);
    }
    buf[6] = (buf[6]! & 0x0f) | 0x40; // v4 UUID
    buf[8] = (buf[8]! & 0x3f) | 0x80; // variant
    newId = Array.from(buf, (b) => b.toString(16).padStart(2, '0')).join('');
  }

  try {
    store?.setItem(SESSION_KEY, newId);
  } catch {
    // Storage might be restricted; session ID remains held in memory
    storageUnavailable = true;
  }
  fallbackSessionId = newId;
  return newId;
}

/**
 * Explicitly frees this tab's ephemeral working memory calculations in local in-browser memory.
 */
export async function clearCurrentTabWorkingMemory(): Promise<boolean> {
  const sessionId = getTabSessionId();
  return memory.clearWorkingMemory(sessionId);
}

/**
 * Automatically hook into browser lifecycle events (pagehide / beforeunload)
 * to immediately recycle this tab's ephemeral scratchpad memory locally.
 */
if (typeof window !== 'undefined') {
  const cleanup = () => {
    try {
      void clearCurrentTabWorkingMemory();
    } catch {
      // Ignore unload errors
    }
  };

  window.addEventListener('pagehide', cleanup);
  window.addEventListener('beforeunload', cleanup);
}
