/**
 * Best-effort clipboard write used by non-essential copy buttons. Clipboard permissions vary by
 * browser, private mode, and whether the page is in a secure context; callers should only show a
 * success state after this promise resolves true.
 */
export async function writeClipboardText(text: string): Promise<boolean> {
  try {
    if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    if (typeof document === 'undefined' || typeof document.execCommand !== 'function') return false;
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    const parent = document.body;
    if (!parent) return false;
    parent.appendChild(textarea);
    try {
      textarea.select();
      return document.execCommand('copy');
    } finally {
      textarea.remove();
    }
  } catch {
    return false;
  }
}
