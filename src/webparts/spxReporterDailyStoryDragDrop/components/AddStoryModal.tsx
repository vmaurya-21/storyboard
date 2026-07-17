import * as React from 'react';
import { useState } from 'react';
import styles from './AddStoryModal.module.scss';
import { IStory } from './types';

interface IAddStoryModalProps {
  onClose: () => void;
  onAdd: (story: Omit<IStory, 'id' | 'date'>) => void;
}

const AddStoryModal: React.FC<IAddStoryModalProps> = ({ onClose, onAdd }) => {
  const [title, setTitle] = useState('');
  const [description, setDescription] = useState('');
  const [imageUrl, setImageUrl] = useState('');
  const [linkToPost, setLinkToPost] = useState('');
  const [errors, setErrors] = useState<{ [key: string]: string }>({});

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

  // Close modal and reset form
  const handleClose = (): void => {
    setTitle('');
    setDescription('');
    setImageUrl('');
    setLinkToPost('');
    setErrors({});
    onClose();
  };

  // Submit form
  const handleSubmit = (): void => {
    if (validateForm()) {
      onAdd({
        title,
        description: description || '(optional)',
        imageUrl,
        linkToPost
      });
      handleClose();
    }
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
          <h2>Add New Story</h2>
          <button className={styles.closeBtn} onClick={handleClose} title="Close">
            ×
          </button>
        </div>

        <div className={styles.modalBody}>
          <p className={styles.modalDescription}>
            Create a new story by filling out the form below. Fields marked with <span className={styles.required}>*</span> are required.
          </p>
          <p className={styles.modalTip}>
            💡 You can add stories from LinkedIn, external websites, or internal sources. The source will be automatically detected from the URL.
          </p>

          <div className={styles.formGroup}>
            <label htmlFor="story-title">
              Title <span className={styles.required}>*</span>
            </label>
            <input
              id="story-title"
              type="text"
              placeholder="Enter story title"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className={errors.title ? styles.inputError : ''}
            />
            {errors.title && <span className={styles.errorMessage}>{errors.title}</span>}
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="story-description">
              Description
            </label>
            <textarea
              id="story-description"
              placeholder="Enter story description (optional)"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              rows={4}
            />
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="story-image">
              Image URL <span className={styles.required}>*</span>
            </label>
            <input
              id="story-image"
              type="text"
              placeholder="Enter image URL"
              value={imageUrl}
              onChange={(e) => setImageUrl(e.target.value)}
              className={errors.imageUrl ? styles.inputError : ''}
            />
            {errors.imageUrl && <span className={styles.errorMessage}>{errors.imageUrl}</span>}
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="story-link">
              Link to Post <span className={styles.required}>*</span>
            </label>
            <input
              id="story-link"
              type="text"
              placeholder="Enter story URL"
              value={linkToPost}
              onChange={(e) => setLinkToPost(e.target.value)}
              className={errors.linkToPost ? styles.inputError : ''}
            />
            {errors.linkToPost && <span className={styles.errorMessage}>{errors.linkToPost}</span>}
          </div>
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