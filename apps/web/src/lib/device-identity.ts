const INSTALL_ID_STORAGE_KEY = 'excelagento_unique_id_v2';

let cachedUniqueId: string | null = null;

function hashStringToUuid(str: string): string {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  let h3 = 0x811c9dc5;
  let h4 = 0x9e3779b9;
  for (let i = 0; i < str.length; i += 1) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ (ch << 3), 1597334677);
    h3 = Math.imul(h3 ^ (ch << 7), 3812015801);
    h4 = Math.imul(h4 ^ (ch << 11), 2246822507);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h2 = Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h3 = Math.imul(h3 ^ (h3 >>> 16), 2654435761);
  h4 = Math.imul(h4 ^ (h4 >>> 15), 1597334677);
  const part1 = (h1 >>> 0).toString(16).padStart(8, '0');
  const part2 = (h2 >>> 0).toString(16).padStart(8, '0');
  const part3 = (h3 >>> 0).toString(16).padStart(8, '0');
  const part4 = (h4 >>> 0).toString(16).padStart(8, '0');
  const hex = `${part1}${part2}${part3}${part4}`;
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
}

/**
 * Derives a hardware-level raw signature from WebGL GPU vendor/renderer,
 * CPU cores, and screen characteristics (matching projectcortex HWID architecture).
 * This remains consistent across different browsers on the exact same device.
 */
export function getRawHWID(): string {
  let renderer = 'unknown_renderer';
  let vendor = 'unknown_vendor';
  try {
    let gl: WebGLRenderingContext | null = null;
    if (typeof document !== 'undefined') {
      const canvas = document.createElement('canvas');
      gl =
        (canvas.getContext('webgl') as WebGLRenderingContext | null) ||
        (canvas.getContext('experimental-webgl' as 'webgl') as WebGLRenderingContext | null);
    } else if (typeof OffscreenCanvas !== 'undefined') {
      const canvas = new OffscreenCanvas(1, 1);
      gl =
        (canvas.getContext('webgl') as WebGLRenderingContext | null) ||
        (canvas.getContext as (id: string) => WebGLRenderingContext | null)('experimental-webgl');
    }
    if (gl) {
      const debugInfo = gl.getExtension('WEBGL_debug_renderer_info');
      if (debugInfo) {
        renderer = String(gl.getParameter(debugInfo.UNMASKED_RENDERER_WEBGL) || 'unknown');
        vendor = String(gl.getParameter(debugInfo.UNMASKED_VENDOR_WEBGL) || 'unknown');
      }
    }
  } catch {
    // Canvas / WebGL blocked or unavailable
  }

  const cores = typeof navigator !== 'undefined' ? navigator.hardwareConcurrency || 2 : 2;
  const colorDepth = typeof screen !== 'undefined' ? screen.colorDepth || 24 : 24;
  return `${vendor}||${renderer}||${cores}||${colorDepth}`;
}

/**
 * Returns a deterministic Unique Device ID formatted as a standard UUID.
 * The ID is derived from physical hardware characteristics so that opening different
 * browsers or incognito tabs on the same machine resolves to the same device identifier.
 */
export function getInstallId(): string {
  if (cachedUniqueId) return cachedUniqueId;
  const hwid = getRawHWID();
  const generated = hashStringToUuid(hwid);

  try {
    window.localStorage.setItem(INSTALL_ID_STORAGE_KEY, generated);
  } catch {
    // LocalStorage may be unavailable in some sandboxes
  }

  cachedUniqueId = generated;
  return generated;
}

export function getInstallIdHint(installId = getInstallId()): string {
  return `${installId.slice(0, 8)}…${installId.slice(-6)}`;
}
