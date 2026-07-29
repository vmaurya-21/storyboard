import * as React from 'react';
import styles from './AddStoryModal.module.scss';

interface IStoryFormFieldsProps {
  titleId: string;
  descriptionId: string;
  imageUrlId: string;
  linkToPostId: string;
  title: string;
  description: string;
  imageUrl: string;
  linkToPost: string;
  errors: { [key: string]: string };
  getFieldClass: (field: 'title' | 'imageUrl' | 'linkToPost') => string;
  onTitleChange: (value: string) => void;
  onDescriptionChange: (value: string) => void;
  onImageUrlChange: (value: string) => void;
  onLinkToPostChange: (value: string) => void;
}

/** Shared input block used by Add and Edit story modals. */
const StoryFormFields: React.FC<IStoryFormFieldsProps> = ({
  titleId,
  descriptionId,
  imageUrlId,
  linkToPostId,
  title,
  description,
  imageUrl,
  linkToPost,
  errors,
  getFieldClass,
  onTitleChange,
  onDescriptionChange,
  onImageUrlChange,
  onLinkToPostChange,
}) => {
  return (
    <>
      <div className={styles.formGroup}>
        <label htmlFor={titleId}>
          Title <span className={styles.required}>*</span>
        </label>
        <input
          id={titleId}
          type="text"
          placeholder="Enter story title"
          value={title}
          onChange={(e) => onTitleChange(e.target.value)}
          className={getFieldClass('title')}
        />
        {errors.title && <span className={styles.errorMessage}>{errors.title}</span>}
      </div>

      <div className={styles.formGroup}>
        <label htmlFor={descriptionId}>
          Description
        </label>
        <textarea
          id={descriptionId}
          placeholder="Enter story description (optional)"
          value={description}
          onChange={(e) => onDescriptionChange(e.target.value)}
          rows={3}
        />
      </div>

      <div className={styles.formGroup}>
        <label htmlFor={imageUrlId}>
          Image URL <span className={styles.required}>*</span>
        </label>
        <input
          id={imageUrlId}
          type="text"
          placeholder="Enter image URL"
          value={imageUrl}
          onChange={(e) => onImageUrlChange(e.target.value)}
          className={getFieldClass('imageUrl')}
        />
        {errors.imageUrl && <span className={styles.errorMessage}>{errors.imageUrl}</span>}
      </div>

      <div className={styles.formGroup}>
        <label htmlFor={linkToPostId}>
          Link to Post <span className={styles.required}>*</span>
        </label>
        <input
          id={linkToPostId}
          type="text"
          placeholder="Enter story URL"
          value={linkToPost}
          onChange={(e) => onLinkToPostChange(e.target.value)}
          className={getFieldClass('linkToPost')}
        />
        {errors.linkToPost && <span className={styles.errorMessage}>{errors.linkToPost}</span>}
      </div>
    </>
  );
};

export default StoryFormFields;
