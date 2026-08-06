import { useState } from 'react';
import { StoryRequiredField, validateStoryField } from './storyFormValidation';

/** Seed values for {@link useStoryForm}. Any omitted field starts empty. */
interface IUseStoryFormInitialValues {
  /** Initial headline. */
  title?: string;
  /** Initial summary body. */
  description?: string;
  /** Initial image URL. */
  imageUrl?: string;
  /** Initial post URL. */
  linkToPost?: string;
}

/** Current value of every field the story form owns. */
interface IStoryFormValues {
  /** Headline. Required. */
  title: string;
  /** Summary body. Optional. */
  description: string;
  /** Image URL. Required, and must be absolute. */
  imageUrl: string;
  /** Post URL. Required, and must be absolute. */
  linkToPost: string;
}

/** What {@link useStoryForm} hands back to a story modal. */
interface IUseStoryForm {
  /** Current field values. */
  values: IStoryFormValues;
  /** Validation messages, keyed by field. Only failing fields appear. */
  errors: { [key: string]: string };
  /** Sets the description. Unvalidated, since it is optional. */
  setDescription: (value: string) => void;
  /** Sets the title and refreshes its error. */
  handleTitleChange: (value: string) => void;
  /** Sets the image URL and refreshes its error. */
  handleImageUrlChange: (value: string) => void;
  /** Sets the post URL and refreshes its error. */
  handleLinkToPostChange: (value: string) => void;
  /** Validates every required field and publishes the results. Returns `true` when submittable. */
  validateForm: () => boolean;
  /** Returns `errorClassName` when the field is failing, otherwise `''`. */
  getFieldClass: (field: StoryRequiredField, errorClassName: string) => string;
}

/**
 * Shared state and validation behavior for Add/Edit story modals.
 *
 * Errors are stored only for fields that currently fail, so `validateForm` can
 * treat a non-empty `errors` object as "not submittable".
 *
 * @param initialValues - Seed values; omitted fields start empty.
 * @returns The form's values, errors and handlers.
 */
export const useStoryForm = (initialValues?: IUseStoryFormInitialValues): IUseStoryForm => {
  const [title, setTitle] = useState(initialValues?.title || '');
  const [description, setDescription] = useState(initialValues?.description || '');
  const [imageUrl, setImageUrl] = useState(initialValues?.imageUrl || '');
  const [linkToPost, setLinkToPost] = useState(initialValues?.linkToPost || '');
  const [errors, setErrors] = useState<{ [key: string]: string }>({});

  const updateFieldError = (field: StoryRequiredField, value: string): void => {
    setErrors(prev => {
      const next = { ...prev };
      const err = validateStoryField(field, value);
      if (err) {
        next[field] = err;
      } else {
        delete next[field];
      }
      return next;
    });
  };

  const handleTitleChange = (value: string): void => {
    setTitle(value);
    updateFieldError('title', value);
  };

  const handleImageUrlChange = (value: string): void => {
    setImageUrl(value);
    updateFieldError('imageUrl', value);
  };

  const handleLinkToPostChange = (value: string): void => {
    setLinkToPost(value);
    updateFieldError('linkToPost', value);
  };

  const validateForm = (): boolean => {
    const newErrors: { [key: string]: string } = {};

    const requiredValues: Record<StoryRequiredField, string> = {
      title,
      imageUrl,
      linkToPost,
    };

    (Object.keys(requiredValues) as StoryRequiredField[]).forEach((field) => {
      const err = validateStoryField(field, requiredValues[field]);
      if (err) {
        newErrors[field] = err;
      }
    });

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const getFieldClass = (field: StoryRequiredField, errorClassName: string): string => {
    return errors[field] ? errorClassName : '';
  };

  const values: IStoryFormValues = {
    title,
    description,
    imageUrl,
    linkToPost,
  };

  return {
    values,
    errors,
    setDescription,
    handleTitleChange,
    handleImageUrlChange,
    handleLinkToPostChange,
    validateForm,
    getFieldClass,
  };
};
