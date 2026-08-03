import * as React from 'react';
import styles from './AddStoryModal.module.scss';
import { IStory } from './types';
import { XIcon } from './icons';
import { useDialogBehavior } from './useDialogBehavior';
import StoryFormFields from './StoryFormFields';
import { useStoryForm } from './useStoryForm';
import { TITLE_MAX_LENGTH } from './storyFormValidation';

/** Props for {@link AddStoryModal}. */
interface IAddStoryModalProps {
  /** Called when the dialog is dismissed without saving. */
  onClose: () => void;
  /**
   * Called with the new story when the form passes validation. The id and
   * display date are assigned by SharePoint, so they are omitted here.
   */
  onAdd: (story: Omit<IStory, 'id' | 'date'>) => void;
  /** Raises a notification. Used to report validation failures. */
  showToast?: (message: string, type: 'success' | 'error' | 'info', description?: string) => void;
}

/**
 * Dialog for creating a story.
 *
 * Title, image URL and post link are required; the description is optional.
 * Fields validate as they are edited and again on submit. The dialog is mounted
 * only while open, so a reopened dialog starts clean without an explicit reset.
 */
const AddStoryModal: React.FC<IAddStoryModalProps> = ({ onClose, onAdd, showToast }) => {
  const {
    values,
    errors,
    setDescription,
    handleTitleChange,
    handleImageUrlChange,
    handleLinkToPostChange,
    validateForm,
    getFieldClass,
  } = useStoryForm();

  /** Dismisses the dialog. */
  const handleClose = (): void => {
    onClose();
  };

  // Escape-to-close, focus trap and focus restore — the behaviour Radix gives
  // the reference dialog for free. Wrapped so the latest `handleClose` is used
  // without re-running the effect.
  const dialogRef = useDialogBehavior(() => handleClose());

  /**
   * Submits the form.
   *
   * On success the story is handed to `onAdd` and the dialog closes. On
   * failure the dialog stays open and a toast names which of the three
   * problems it is: something required is missing, the title is too long, or
   * a URL is malformed.
   */
  const handleSubmit = (): void => {
    if (validateForm()) {
      onAdd({
        title: values.title,
        description: values.description.trim(),
        imageUrl: values.imageUrl,
        linkToPost: values.linkToPost
      });
      handleClose();
      return;
    }

    if (!values.title.trim() || !values.imageUrl.trim() || !values.linkToPost.trim()) {
      showToast?.(
        'Missing Required Fields',
        'error',
        'Please fill in Title, Image URL, and Link to Post.'
      );
    } else if (values.title.length > TITLE_MAX_LENGTH) {
      // Own branch: an over-long title is neither missing nor malformed, and
      // reporting it as "Invalid URL" contradicted the inline field error.
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
    <div className={styles.modalOverlay} onClick={handleOverlayClick}>
      <div
        ref={dialogRef}
        className={styles.modalContent}
        role="dialog"
        aria-modal="true"
        aria-labelledby="add-story-heading"
      >
        <div className={styles.modalHeader}>
          <h2 id="add-story-heading" className={styles.modalTitle}>Add New Story</h2>
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
            Create a new story by filling out the form below. Fields marked with <span className={styles.required}>*</span> are required.
          </p>
          <p className={styles.modalTip}>
            💡 You can add stories from LinkedIn, external websites, or internal sources. The source will be automatically detected from the URL.
          </p>

          <StoryFormFields
            titleId="story-title"
            descriptionId="story-description"
            imageUrlId="story-image"
            linkToPostId="story-link"
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

        <div className={styles.modalFooter}>
          <button className={styles.btnCancel} onClick={handleClose}>
            Cancel
          </button>
          <button className={styles.btnSave} onClick={handleSubmit}>
            Save
          </button>
        </div>
      </div>
    </div>
  );
};

export default AddStoryModal;