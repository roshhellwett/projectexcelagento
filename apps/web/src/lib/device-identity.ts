const INSTALL_ID_STORAGE_KEY = 'excelagento_install_id_v1';

let memoryInstallId: string | null = null;

function fallbackInstallId(): string {
  const bytes = new Uint8Array(16);
  if (typeof globalThis.crypto?.getRandomValues === 'function') {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = Math.floor(Math.random() * 256);
    }
  }
  const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

function createInstallId(): string {
  return globalThis.crypto?.randomUUID?.() ?? fallbackInstallId();
}

/**
 * Returns a stable app-install identifier. It deliberately does not collect a
 * WebGL, canvas, font, or hardware fingerprint: browsers cannot provide an
 * immutable hardware ID, and those fingerprints are unnecessarily invasive.
 */
export function getInstallId(): string {
  if (memoryInstallId) return memoryInstallId;
  try {
    const existing = window.localStorage.getItem(INSTALL_ID_STORAGE_KEY)?.trim();
    if (existing) {
      memoryInstallId = existing;
      return existing;
    }
    const created = createInstallId();
    window.localStorage.setItem(INSTALL_ID_STORAGE_KEY, created);
    memoryInstallId = created;
    return created;
  } catch {
    memoryInstallId ??= createInstallId();
    return memoryInstallId;
  }
}

export function getInstallIdHint(installId = getInstallId()): string {
  return `${installId.slice(0, 8)}…${installId.slice(-6)}`;
}
