import { useEffect, useRef } from 'react';

/** Elements that can hold focus inside a dialog. */
const FOCUSABLE =
  'a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * Gives a hand-rolled modal the behaviour Radix provides in the reference.
 *
 * Three things, all of which a native `<div>` overlay lacks:
 *
 * - **Escape closes.** Bound on the document so it works wherever focus sits.
 * - **Focus is trapped.** Tab and Shift+Tab cycle within the dialog instead of
 *   escaping to the page behind it.
 * - **Focus is restored.** Whatever was focused before the dialog opened is
 *   refocused on unmount, so keyboard users are not dropped at the top of the
 *   document.
 *
 * On mount, focus moves to the first focusable control. Pair the returned ref
 * with `role="dialog"` and `aria-modal="true"` on the dialog element.
 *
 * @param onClose - Called when Escape is pressed.
 * @returns Ref to attach to the dialog's outermost element.
 */
export const useDialogBehavior = (onClose: () => void): React.RefObject<HTMLDivElement> => {
  const dialogRef = useRef<HTMLDivElement>(null);
  // Held in a ref so the focus/Escape effect does not re-run when the caller
  // passes a new inline function on every render.
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;

  useEffect(() => {
    const previouslyFocused = document.activeElement as HTMLElement | null;

    const focusable = (): HTMLElement[] =>
      dialogRef.current
        ? Array.prototype.slice.call(dialogRef.current.querySelectorAll(FOCUSABLE))
        : [];

    const first = focusable()[0];
    if (first) first.focus();

    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onCloseRef.current();
        return;
      }
      if (e.key !== 'Tab') return;

      const items = focusable();
      if (items.length === 0) return;

      const firstItem = items[0];
      const lastItem = items[items.length - 1];
      const active = document.activeElement;

      // Wrap at both ends. Focus sitting outside the dialog is pulled back in.
      if (e.shiftKey && (active === firstItem || !dialogRef.current?.contains(active))) {
        e.preventDefault();
        lastItem.focus();
      } else if (!e.shiftKey && active === lastItem) {
        e.preventDefault();
        firstItem.focus();
      }
    };

    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('keydown', onKeyDown);
      if (previouslyFocused && typeof previouslyFocused.focus === 'function') {
        previouslyFocused.focus();
      }
    };
  }, []);

  return dialogRef;
};
