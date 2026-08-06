import * as React from 'react';
import { ScheduledStoryGroup } from './types';
import AlertDialog from './AlertDialog';

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
 * Names the group's date and story count so the editor can tell which schedule
 * they are about to discard. Escape, the focus trap and the refusal to close on
 * a backdrop click all come from {@link AlertDialog}.
 */
const DeleteGroupModal: React.FC<IDeleteGroupModalProps> = ({ group, onClose, onConfirm }) => {
  // Reference: `{ weekday: 'long', month: 'short', day: 'numeric' }`.
  const formattedDate = group.date.toLocaleDateString('en-US', {
    weekday: 'long',
    month: 'short',
    day: 'numeric'
  });

  // Reference: `groupToDelete.stories.length`. This board keeps the
  // arrangement as a slot-to-story map, so the count is the filled slots —
  // empty slots hold `undefined`.
  const storyCount = Object.keys(group.slotStories).filter(
    (slotId) => group.slotStories[slotId]
  ).length;

  return (
    <AlertDialog
      title="Are you sure you want to delete this scheduled group?"
      titleId="delete-group-heading"
      actionLabel="Yes, Delete Group"
      onCancel={onClose}
      onAction={onConfirm}
    >
      {/* Reference: the date sits in a `font-medium` span — 500, not
          `<strong>`'s 700. */}
      This will remove the scheduled group for{' '}
      <span style={{ fontWeight: 500 }}>{formattedDate}</span>
      {' '}and return all {storyCount} {storyCount === 1 ? 'story' : 'stories'} to the available list.
    </AlertDialog>
  );
};

export default DeleteGroupModal;
