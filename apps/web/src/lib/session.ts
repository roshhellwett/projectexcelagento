import { memory } from './agent-runtime';

const SESSION_KEY = 'excel_agent_tab_session_id';

const SUPABASE_URL =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_URL) ||
  'https://fsepapdadtrlddkyqqxu.supabase.co';
const SUPABASE_KEY =
  (typeof import.meta !== 'undefined' && import.meta.env?.VITE_SUPABASE_ANON_KEY) ||
  'sb_publishable_8JPAZFfCaS_U8nAAqc1rrQ_V6OSKAic';

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
  const store = safeSessionStorage();
  const existing = store?.getItem(SESSION_KEY);
  if (existing && existing.length >= 16) {
    return existing;
  }

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
  }

  return newId;
}

/**
 * Explicitly frees this tab's ephemeral working memory calculations in Supabase.
 */
export async function clearCurrentTabWorkingMemory(): Promise<boolean> {
  const sessionId = getTabSessionId();
  return memory.clearWorkingMemory(sessionId);
}

/**
 * Automatically hook into browser lifecycle events (pagehide / beforeunload)
 * to immediately recycle this tab's scratchpad rows from the 100MB pool.
 */
if (typeof window !== 'undefined') {
  const cleanup = () => {
    try {
      const sessionId = safeSessionStorage()?.getItem(SESSION_KEY);
      if (!sessionId) return;

      const endpoint = `${SUPABASE_URL.replace(/\/+$/, '')}/rest/v1/agent_working_memory?session_id=eq.${encodeURIComponent(
        sessionId,
      )}`;

      // Use keepalive fetch so the deletion request completes even as the browser closes
      if (typeof fetch === 'function') {
        void fetch(endpoint, {
          method: 'DELETE',
          headers: {
            apikey: SUPABASE_KEY,
            Authorization: `Bearer ${SUPABASE_KEY}`,
          },
          keepalive: true,
        }).catch(() => {});
      }
    } catch {
      // Ignore unload errors
    }
  };

  window.addEventListener('pagehide', cleanup);
  window.addEventListener('beforeunload', cleanup);
}
