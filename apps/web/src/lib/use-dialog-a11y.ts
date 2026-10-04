import { useEffect, useRef, type RefObject } from 'react';

/**
 * Modal behaviour for the dialogs and drawers in this app.
 *
 * Opening one moves focus inside it, keeps Tab inside it, closes it on Escape, and gives focus back
 * to whatever opened it. Without any of that a dialog is a visual convention only: a keyboard user
 * tabs straight out of it into the page behind, and Escape does nothing.
 *
 * The panel is marked `data-dialog-open` so the grid's keyboard handling can stand down while
 * something is on top of it - a grid that swallows Escape or arrow keys while a modal is open is
 * worse than no keyboard support at all.
 */
export function useDialogA11y<T extends HTMLElement>(
  isOpen: boolean,
  panelRef: RefObject<T | null>,
  onClose: () => void,
): void {
  const closeRef = useRef(onClose);
  closeRef.current = onClose;

  useEffect(() => {
    if (!isOpen) return;
    const panel = panelRef.current;
    if (!panel) return;

    const previouslyFocused = document.activeElement as HTMLElement | null;
    const focusables = (): HTMLElement[] =>
      [
        ...panel.querySelectorAll<HTMLElement>(
          'button, [href], input, select, textarea, [tabindex]',
        ),
      ].filter((el) => !el.hasAttribute('disabled') && el.getAttribute('tabindex') !== '-1');

    // Prefer the first real control; fall back to the panel itself so focus never stays behind.
    const initial =
      panel.querySelector<HTMLElement>('[data-autofocus]') ?? focusables()[0] ?? panel;
    if (initial === panel && !panel.hasAttribute('tabindex')) {
      panel.setAttribute('tabindex', '-1');
    }
    initial.focus();

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        closeRef.current();
        return;
      }
      if (event.key !== 'Tab') return;

      const items = focusables();
      if (items.length === 0) {
        event.preventDefault();
        return;
      }
      const first = items[0] as HTMLElement;
      const last = items[items.length - 1] as HTMLElement;
      const active = document.activeElement;
      if (event.shiftKey && (active === first || active === panel)) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && active === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener('keydown', onKeyDown, true);
    return () => {
      document.removeEventListener('keydown', onKeyDown, true);
      // Focus goes back where it came from, so closing a dialog does not dump the user at the top
      // of the document.
      if (previouslyFocused && document.contains(previouslyFocused)) previouslyFocused.focus();
    };
  }, [isOpen, panelRef]);
}
