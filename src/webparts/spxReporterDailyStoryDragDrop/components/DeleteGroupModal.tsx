import * as React from 'react';
import styles from './AddStoryModal.module.scss';
import { ScheduledStoryGroup } from './types';
import { formatTime12 } from './DateTimePicker';

/** Props for {@link DeleteGroupModal}. */
interface IDeleteGroupModalProps {
  /** Group the dialog is asking about. */
  group: ScheduledStoryGroup;
  /** Called when the dialog is dismissed without deleting. */
  onClose: () => void;
  /** Called when deletion is confirmed. */
  onConfirm: () => void;
}

/**
 * Confirmation dialog for removing a scheduled group.
 *
 * Names the group's date, time and story count so the editor can tell which
 * schedule they are about to discard.
 */
const DeleteGroupModal: React.FC<IDeleteGroupModalProps> = ({ group, onClose, onConfirm }) => {
  const formattedDate = group.date.toLocaleDateString('en-US', {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  });

  const formattedTime = formatTime12(group.time);

  /**
   * Dismisses the dialog when the backdrop itself is clicked.
   *
   * The target check keeps clicks inside the dialog from bubbling up and
   * closing it.
   *
   * @param e - Click event from the overlay element.
   */
  const handleOverlayClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) {
      onClose();
    }
  };

  return (
    <div className={styles.modalOverlay} onClick={handleOverlayClick}>
      <div className={styles.modalContent}>
        <div className={styles.modalHeader}>
          <h2>Delete Scheduled Group?</h2>
          <button className={styles.closeBtn} onClick={onClose} title="Close">
            ×
          </button>
        </div>

        <div className={styles.modalBody}>
          <p className={styles.modalDescription} style={{ textAlign: 'left', marginBottom: '16px' }}>
            Are you sure you want to delete this scheduled group?
          </p>
          <p className={styles.modalTip} style={{ textAlign: 'left', color: '#4b5563', marginBottom: '8px' }}>
            This will remove the scheduled group for <strong>{formattedDate}</strong> at <strong>{formattedTime}</strong>.
          </p>
          <p className={styles.modalTip} style={{ textAlign: 'left', color: '#dc2626', fontWeight: 500 }}>
            ⚠️ All stories in this group will be returned to the available stories list.
          </p>
        </div>

        <div className={styles.modalFooter}>
          <button className={styles.btnCancel} onClick={onClose}>
            Cancel
          </button>
          <button className={styles.btnDelete} onClick={onConfirm}>
            Delete
          </button>
        </div>
      </div>
    </div>
  );
};

export default DeleteGroupModal;
