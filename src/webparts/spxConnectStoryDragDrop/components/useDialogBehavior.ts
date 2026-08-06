import { useEffect, useRef } from 'react';

/** Elements that can hold focus inside a dialog. */
const FOCUSABLE =
  'a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * How many dialogs are currently mounted, and what `body.style.overflow` was
 * before the first of them locked it.
 *
 * Module scope rather than component state because the lock is a property of
 * the page, not of any one dialog: a dialog that opens a nested confirmation
 * must not restore scrolling when the confirmation closes.
 */
let openDialogCount = 0;
let previousBodyOverflow = '';

/** Options for {@link useDialogBehavior}. */
interface IUseDialogBehaviorOptions {
  /**
   * Whether this dialog currently owns the keyboard. Defaults to `true`.
   *
   * Set to `false` while a nested dialog is open so Escape and Tab are handled
   * by the dialog on top rather than by the one behind it — Radix does the
   * same by only listening on the topmost layer.
   */
  enabled?: boolean;
}

/**
 * Gives a hand-rolled modal the behaviour Radix provides in the reference.
 *
 * Four things, all of which a native `<div>` overlay lacks:
 *
 * - **Escape closes.** Bound on the document so it works wherever focus sits.
 * - **Focus is trapped.** Tab and Shift+Tab cycle within the dialog instead of
 *   escaping to the page behind it.
 * - **Focus is restored.** Whatever was focused before the dialog opened is
 *   refocused on unmount, so keyboard users are not dropped at the top of the
 *   document.
 * - **The page behind is scroll-locked**, refcounted so nested dialogs unlock
 *   only once the last one closes.
 *
 * On mount, focus moves to the first focusable control. Pair the returned ref
 * with `role="dialog"` and `aria-modal="true"` on the dialog element.
 *
 * @param onClose - Called when Escape is pressed.
 * @param options - See {@link IUseDialogBehaviorOptions}.
 * @returns Ref to attach to the dialog's outermost element.
 */
export const useDialogBehavior = (
  onClose: () => void,
  options?: IUseDialogBehaviorOptions
): React.RefObject<HTMLDivElement> => {
  const enabled = options?.enabled !== false;
  const dialogRef = useRef<HTMLDivElement>(null);
  // Held in a ref so the focus/Escape effect does not re-run when the caller
  // passes a new inline function on every render.
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  // Autofocus is a once-per-mount thing. Without this guard, closing a nested
  // confirmation would re-enable this dialog and yank focus to its close
  // button, overriding the focus the confirmation just restored.
  const hasAutoFocused = useRef(false);

  // Mount-scoped concerns. Deliberately independent of `enabled`: a dialog that
  // has handed the keyboard to a nested confirmation is still on screen, so the
  // page behind it must stay locked and its focus-restore target must survive.
  useEffect(() => {
    const previouslyFocused = document.activeElement as HTMLElement | null;

    if (openDialogCount === 0) {
      previousBodyOverflow = document.body.style.overflow;
      document.body.style.overflow = 'hidden';
    }
    openDialogCount += 1;

    return () => {
      openDialogCount -= 1;
      if (openDialogCount === 0) {
        document.body.style.overflow = previousBodyOverflow;
      }
      if (previouslyFocused && typeof previouslyFocused.focus === 'function') {
        previouslyFocused.focus();
      }
    };
  }, []);

  useEffect(() => {
    if (!enabled) return undefined;

    const focusable = (): HTMLElement[] =>
      dialogRef.current
        ? Array.prototype.slice.call(dialogRef.current.querySelectorAll(FOCUSABLE))
        : [];

    if (!hasAutoFocused.current) {
      const first = focusable()[0];
      if (first) {
        first.focus();
        hasAutoFocused.current = true;
      }
    }

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
      } else if (!e.shiftKey && (active === lastItem || !dialogRef.current?.contains(active))) {
        e.preventDefault();
        firstItem.focus();
      }
    };

    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [enabled]);

  return dialogRef;
};
