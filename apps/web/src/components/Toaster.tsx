import React, { useCallback, useEffect, useRef, useState } from 'react';
import { playUiSound } from '../lib/sound-effects.js';

export type ToastKind = 'success' | 'error' | 'info' | 'warning';

export interface Toast {
  id: string;
  kind: ToastKind;
  message: string;
}

let toastSequence = 0;

/** Minimal, dependency-free toast queue used instead of blocking `alert()` calls. */
export function useToasts() {
  const [toasts, setToasts] = useState<Toast[]>([]);
  const timersRef = useRef(new Map<string, ReturnType<typeof setTimeout>>());

  useEffect(
    () => () => {
      for (const timer of timersRef.current.values()) clearTimeout(timer);
      timersRef.current.clear();
    },
    [],
  );

  const dismissToast = useCallback((id: string) => {
    const timer = timersRef.current.get(id);
    if (timer !== undefined) clearTimeout(timer);
    timersRef.current.delete(id);
    setToasts((prev) => prev.filter((toast) => toast.id !== id));
  }, []);

  const pushToast = useCallback(
    (
      kind: ToastKind,
      message: string,
      ttlMs = kind === 'error' ? 6000 : kind === 'warning' ? 9000 : 3500,
    ) => {
      toastSequence += 1;
      const id = `toast-${Date.now().toString(36)}-${toastSequence.toString(36)}`;
      setToasts((prev) => [...prev.slice(-3), { id, kind, message }]);
      playUiSound(kind === 'success' ? 'success' : kind === 'error' ? 'error' : 'click');
      timersRef.current.set(
        id,
        setTimeout(() => {
          timersRef.current.delete(id);
          dismissToast(id);
        }, ttlMs),
      );
      return id;
    },
    [dismissToast],
  );

  return { toasts, pushToast, dismissToast };
}

// A warning is the one state that must not be missed: it is how a partial import reports that
// something in the user's file could not be kept.
const ICONS: Record<ToastKind, string> = {
  success: '✓',
  error: '✕',
  info: 'ℹ',
  warning: '⚠',
};

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
