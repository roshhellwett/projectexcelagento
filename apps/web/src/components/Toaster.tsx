import React, { useCallback, useState } from 'react';

export type ToastKind = 'success' | 'error' | 'info';

export interface Toast {
  id: string;
  kind: ToastKind;
  message: string;
}

let toastSequence = 0;

/** Minimal, dependency-free toast queue used instead of blocking `alert()` calls. */
export function useToasts() {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((toast) => toast.id !== id));
  }, []);

  const pushToast = useCallback(
    (kind: ToastKind, message: string, ttlMs = kind === 'error' ? 6000 : 3500) => {
      toastSequence += 1;
      const id = `toast-${Date.now().toString(36)}-${toastSequence.toString(36)}`;
      setToasts((prev) => [...prev.slice(-3), { id, kind, message }]);
      setTimeout(() => dismissToast(id), ttlMs);
      return id;
    },
    [dismissToast],
  );

  return { toasts, pushToast, dismissToast };
}

const ICONS: Record<ToastKind, string> = { success: '✓', error: '✕', info: 'ℹ' };

export const ToastHost: React.FC<{ toasts: Toast[]; onDismiss: (id: string) => void }> = ({
  toasts,
  onDismiss,
}) => {
  if (toasts.length === 0) return null;

  return (
    <div className="toast-host" role="status" aria-live="polite">
      {toasts.map((toast) => (
        <div key={toast.id} className={`toast toast-${toast.kind}`}>
          <span className="toast-icon">{ICONS[toast.kind]}</span>
          <span className="toast-message">{toast.message}</span>
          <button
            type="button"
            className="toast-close"
            onClick={() => onDismiss(toast.id)}
            aria-label="Dismiss notification"
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  );
};
