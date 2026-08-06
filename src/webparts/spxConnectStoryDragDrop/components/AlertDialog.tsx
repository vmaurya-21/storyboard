import * as React from 'react';
import styles from './AddStoryModal.module.scss';
import { XIcon } from './icons';
import { useDialogBehavior } from './useDialogBehavior';

/** Props for {@link AlertDialog}. */
interface IAlertDialogProps {
  /** Question the dialog is asking. Rendered as the accessible name. */
  title: string;
  /** Id for the title element, so `aria-labelledby` can point at it. */
  titleId: string;
  /** Supporting copy beneath the title. */
  children: React.ReactNode;
  /** Label for the destructive action. e.g. `Yes, Delete Group`. */
  actionLabel: string;
  /** Called when the dialog is dismissed without acting. */
  onCancel: () => void;
  /** Called when the destructive action is confirmed. */
  onAction: () => void;
  /** Stacking order, for a confirmation opened on top of another dialog. */
  zIndex?: number;
}

/**
 * Confirmation dialog, modelled on the reference's `AlertDialog`.
 *
 * The distinction from {@link AddStoryModal}'s shell matters: `AlertDialogContent`
 * is `grid gap-4 p-6` — a flat 24px at every breakpoint, where `DialogContent`
 * carries `p-4 md:p-8` — and `AlertDialogFooter` is `flex-col-reverse` below
 * 640px, stacking the buttons with the action above Cancel. Reusing the story
 * dialog's classes got both wrong.
 *
 * Radix `AlertDialog` deliberately ignores backdrop clicks: a destructive
 * confirmation should not be dismissible by a stray click. Escape still closes,
 * via {@link useDialogBehavior}.
 */
const AlertDialog: React.FC<IAlertDialogProps> = ({
  title,
  titleId,
  children,
  actionLabel,
  onCancel,
  onAction,
  zIndex,
}) => {
  const dialogRef = useDialogBehavior(onCancel);

  return (
    // No `onClick` on the overlay — see the note above.
    <div className={styles.modalOverlay} style={zIndex ? { zIndex } : undefined}>
      <div
        ref={dialogRef}
        className={styles.modalContent}
        role="alertdialog"
        aria-modal="true"
        aria-labelledby={titleId}
      >
        <div className={styles.alertHeader}>
          <h2 id={titleId} className={styles.alertTitle}>{title}</h2>
          {/* Reference `AlertDialogContent` renders no close button. The `×` is
              kept as a deliberate deviation: SharePoint editors reach for it,
              and Cancel plus Escape are otherwise the only ways out. */}
          <button
            type="button"
            className={styles.closeBtn}
            onClick={onCancel}
            title="Close"
            aria-label="Close"
          >
            <XIcon />
          </button>
        </div>

        {/* Reference: a single AlertDialogDescription. */}
        <div className={styles.alertBody}>
          <p className={styles.modalDescription}>{children}</p>
        </div>

        <div className={styles.alertFooter}>
          <button className={styles.btnCancel} onClick={onCancel}>
            Cancel
          </button>
          <button className={styles.btnAlertAction} onClick={onAction}>
            {actionLabel}
          </button>
        </div>
      </div>
    </div>
  );
};

export default AlertDialog;
