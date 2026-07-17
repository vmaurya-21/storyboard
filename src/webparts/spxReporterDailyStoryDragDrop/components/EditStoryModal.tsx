import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './AddStoryModal.module.scss';
import { IStory } from './types';

interface IEditStoryModalProps {
  story: IStory;
  onClose: () => void;
  onUpdate: (story: IStory) => void;
  onDelete: (storyId: string) => void;
}

const EditStoryModal: React.FC<IEditStoryModalProps> = ({ story, onClose, onUpdate, onDelete }) => {
  const [title, setTitle] = useState(story.title);
  const [description, setDescription] = useState(story.description);
  const [imageUrl, setImageUrl] = useState(story.imageUrl);
  const [linkToPost, setLinkToPost] = useState(story.linkToPost);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);

  // Helper function to validate URLs
  const isValidUrl = (url: string): boolean => {
    try {
      // eslint-disable-next-line no-new
      new URL(url);
      return true;
    } catch {
      return false;
    }
  };

  // Validate form inputs
  const validateForm = (): boolean => {
    const newErrors: { [key: string]: string } = {};

    if (!title.trim()) {
      newErrors.title = 'Title is required';
    }

    if (!imageUrl.trim()) {
      newErrors.imageUrl = 'Image URL is required';
    } else if (!isValidUrl(imageUrl)) {
      newErrors.imageUrl = 'Please enter a valid URL';
    }

    if (!linkToPost.trim()) {
      newErrors.linkToPost = 'Link to Post is required';
    } else if (!isValidUrl(linkToPost)) {
      newErrors.linkToPost = 'Please enter a valid URL';
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  // Close modal
  const handleClose = (): void => {
    onClose();
  };

  // Submit form
  const handleSubmit = (): void => {
    if (validateForm()) {
      onUpdate({
        ...story,
        title,
        description: description || '(optional)',
        imageUrl: imageUrl || 'https://www.spxdaily.com/images-bg/extra-solar-flares-patch-bg.jpg',
        linkToPost: linkToPost || '#'
      });
      handleClose();
    }
  };

  // Delete story
  const handleDelete = (): void => {
    setIsDeleteConfirmOpen(true);
  };

  const handleOverlayClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) {
      handleClose();
    }
  };

  return (
    <div className={styles.modalOverlay} onClick={handleOverlayClick}>
      <div className={styles.modalContent}>
        <div className={styles.modalHeader}>
          <h2>Edit Story</h2>
          <button className={styles.closeBtn} onClick={handleClose} title="Close">
            ×
          </button>
        </div>

        <div className={styles.modalBody}>
          <p className={styles.modalDescription}>
            Update the story details below. Fields marked with <span className={styles.required}>*</span> are required.
          </p>
          <p className={styles.modalTip}>
            💡 You can add stories from LinkedIn, external websites, or internal sources. The source will be automatically detected from the URL.
          </p>

          <div className={styles.formGroup}>
            <label htmlFor="edit-story-title">
              Title <span className={styles.required}>*</span>
            </label>
            <input
              id="edit-story-title"
              type="text"
              placeholder="Enter story title"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className={errors.title ? styles.inputError : ''}
            />
            {errors.title && <span className={styles.errorMessage}>{errors.title}</span>}
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="edit-story-description">
              Description
            </label>
            <textarea
              id="edit-story-description"
              placeholder="Enter story description (optional)"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              rows={4}
            />
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="edit-story-image">
              Image URL <span className={styles.required}>*</span>
            </label>
            <input
              id="edit-story-image"
              type="text"
              placeholder="Enter image URL"
              value={imageUrl}
              onChange={(e) => setImageUrl(e.target.value)}
              className={errors.imageUrl ? styles.inputError : ''}
            />
            {errors.imageUrl && <span className={styles.errorMessage}>{errors.imageUrl}</span>}
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="edit-story-link">
              Link to Post <span className={styles.required}>*</span>
            </label>
            <input
              id="edit-story-link"
              type="text"
              placeholder="Enter story URL"
              value={linkToPost}
              onChange={(e) => setLinkToPost(e.target.value)}
              className={errors.linkToPost ? styles.inputError : ''}
            />
            {errors.linkToPost && <span className={styles.errorMessage}>{errors.linkToPost}</span>}
          </div>
        </div>

        <div className={styles.modalFooterWithDelete}>
          <button className={styles.btnDelete} onClick={handleDelete} title="Delete story">
            Delete Story
          </button>
          <div className={styles.buttonGroup}>
            <button className={styles.btnCancel} onClick={handleClose}>
              Cancel
            </button>
            <button className={styles.btnSave} onClick={handleSubmit}>
              Save
            </button>
          </div>
        </div>
      </div>

      {isDeleteConfirmOpen && (
        <div className={styles.modalOverlay} style={{ zIndex: 1100 }}>
          <div className={styles.modalContent} style={{ maxWidth: '400px' }}>
            <div className={styles.modalHeader}>
              <h2>Delete Story?</h2>
              <button 
                className={styles.closeBtn} 
                onClick={() => setIsDeleteConfirmOpen(false)} 
                title="Close"
              >
                ×
              </button>
            </div>
            <div className={styles.modalBody}>
              <p className={styles.modalDescription} style={{ textAlign: 'left', marginBottom: '16px' }}>
                Are you sure you want to delete the story <strong>{story.title}</strong>?
              </p>
              <p className={styles.modalTip} style={{ textAlign: 'left', color: '#dc2626', fontWeight: 500 }}>
                ⚠️ This action cannot be undone and will permanently remove this story from SharePoint.
              </p>
            </div>
            <div className={styles.modalFooter}>
              <button className={styles.btnCancel} onClick={() => setIsDeleteConfirmOpen(false)}>
                Cancel
              </button>
              <button 
                className={styles.btnDelete} 
                onClick={() => {
                  onDelete(story.id);
                  setIsDeleteConfirmOpen(false);
                  handleClose();
                }}
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default EditStoryModal;
