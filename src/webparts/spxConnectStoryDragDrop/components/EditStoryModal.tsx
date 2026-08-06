import * as React from 'react';
import { useState } from 'react';
import styles from './AddStoryModal.module.scss';
import { IStory } from './types';
import { XIcon } from './icons';
import { useDialogBehavior } from './useDialogBehavior';
import AlertDialog from './AlertDialog';
import StoryFormFields from './StoryFormFields';
import { useStoryForm } from './useStoryForm';
import { TITLE_MAX_LENGTH } from './storyFormValidation';

/** Props for {@link EditStoryModal}. */
interface IEditStoryModalProps {
  /** Story being edited; seeds the form fields. */
  story: IStory;
  /** Called when the dialog is dismissed without saving. */
  onClose: () => void;
  /** Called with the edited story when the form passes validation. */
  onUpdate: (story: IStory) => void;
  /** Called with the story id once deletion is confirmed. */
  onDelete: (storyId: string) => void;
  /** Raises a notification. Used to report validation failures. */
  showToast?: (message: string, type: 'success' | 'error' | 'info', description?: string) => void;
}

/**
 * Dialog for editing or deleting an existing story.
 *
 * Mirrors `AddStoryModal`'s validation, and adds a two-step delete: the
 * destructive action opens a confirmation rather than firing at once.
 */
const EditStoryModal: React.FC<IEditStoryModalProps> = ({ story, onClose, onUpdate, onDelete, showToast }) => {
  const {
    values,
    errors,
    setDescription,
    handleTitleChange,
    handleImageUrlChange,
    handleLinkToPostChange,
    validateForm,
    getFieldClass,
  } = useStoryForm({
    title: story.title,
    description: story.description,
    imageUrl: story.imageUrl,
    linkToPost: story.linkToPost,
  });
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);

  /** Clears validation errors and dismisses the dialog. */
  const handleClose = (): void => {
    onClose();
  };

  // Escape-to-close, focus trap, focus restore and the body scroll lock — the
  // behaviour Radix gives the reference dialog for free. Keyboard ownership is
  // handed to the confirmation while it is open, so Escape there dismisses the
  // confirmation rather than this whole dialog.
  const dialogRef = useDialogBehavior(() => handleClose(), { enabled: !isDeleteConfirmOpen });

  /**
   * Submits the edits.
   *
   * Passes the original story through with the edited fields applied, so
   * properties the form does not expose survive the round trip. On failure the
   * dialog stays open and a toast names the problem.
   */
  const handleSubmit = (): void => {
    if (validateForm()) {
      onUpdate({
        ...story,
        title: values.title,
        description: values.description.trim(),
        imageUrl: values.imageUrl,
        linkToPost: values.linkToPost
      });
      handleClose();
    } else {
      if (!values.title.trim() || !values.imageUrl.trim() || !values.linkToPost.trim()) {
        showToast?.(
          'Missing Required Fields',
          'error',
          'Please fill in Title, Image URL, and Link to Post.'
        );
      } else if (values.title.length > TITLE_MAX_LENGTH) {
        showToast?.(
          'Title Too Long',
          'error',
          `Title must be ${TITLE_MAX_LENGTH} characters or less.`
        );
      } else {
        showToast?.(
          'Invalid URL',
          'error',
          'Please enter valid URLs for Image URL and Link to Post.'
        );
      }
    }
  };

  /** Opens the inline delete confirmation. Deletion happens only once confirmed. */
  const handleDelete = (): void => {
    setIsDeleteConfirmOpen(true);
  };

  /**
   * Dismisses the dialog when the backdrop itself is clicked.
   *
   * @param e - Click event from the overlay element.
   */
  const handleOverlayClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) {
      handleClose();
    }
  };

  return (
    <>
      <div className={styles.modalOverlay} onClick={handleOverlayClick}>
        <div
          ref={dialogRef}
          className={styles.modalContent}
          role="dialog"
          aria-modal="true"
          aria-labelledby="edit-story-heading"
        >
          <div className={styles.modalHeader}>
            <h2 id="edit-story-heading" className={styles.modalTitle}>Edit Story</h2>
            <button
              type="button"
              className={styles.closeBtn}
              onClick={handleClose}
              title="Close"
              aria-label="Close"
            >
              <XIcon />
            </button>
          </div>

          <div className={styles.modalBody}>
            <p className={styles.modalDescription}>
              Update the story details below. Fields marked with <span className={styles.required}>*</span> are required.
            </p>
            <p className={styles.modalTip}>
              💡 You can add stories from LinkedIn, external websites, or internal sources. The source will be automatically detected from the URL.
            </p>

            <StoryFormFields
              titleId="edit-story-title"
              descriptionId="edit-story-description"
              imageUrlId="edit-story-image"
              linkToPostId="edit-story-link"
              title={values.title}
              description={values.description}
              imageUrl={values.imageUrl}
              linkToPost={values.linkToPost}
              errors={errors}
              getFieldClass={(field) => getFieldClass(field, styles.inputError)}
              onTitleChange={handleTitleChange}
              onDescriptionChange={setDescription}
              onImageUrlChange={handleImageUrlChange}
              onLinkToPostChange={handleLinkToPostChange}
            />
          </div>

          <div className={styles.modalFooterWithDelete}>
            <button className={styles.btnDelete} onClick={handleDelete} title="Delete story">
              Delete Story
            </button>
            <div className={styles.buttonGroup}>
              <button className={styles.btnCancel} onClick={handleClose}>
                Cancel
              </button>
              {/* Reference: `{editingStory ? 'Update' : 'Save'}` — the edit
                  dialog's primary action reads "Update". */}
              <button className={styles.btnSave} onClick={handleSubmit}>
                Update
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* A sibling rather than a child of the overlay above, so its own focus
          trap encloses it. Nested inside, `useDialogBehavior`'s containment
          check treated it as outside the dialog and Shift+Tab pulled focus back
          into the edit form behind it.

          Reference: the AlertDialogAction reads "Yes, Delete Story". */}
      {isDeleteConfirmOpen && (
        <AlertDialog
          title="Are you sure you want to delete this story?"
          titleId="delete-story-heading"
          actionLabel="Yes, Delete Story"
          zIndex={1100}
          onCancel={() => setIsDeleteConfirmOpen(false)}
          onAction={() => {
            onDelete(story.id);
            setIsDeleteConfirmOpen(false);
            handleClose();
          }}
        >
          This action cannot be undone. The story will be permanently removed from all locations.
        </AlertDialog>
      )}
    </>
  );
};

export default EditStoryModal;
