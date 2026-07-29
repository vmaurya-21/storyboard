import { useState } from 'react';
import { StoryRequiredField, validateStoryField } from './storyFormValidation';

interface IUseStoryFormInitialValues {
  title?: string;
  description?: string;
  imageUrl?: string;
  linkToPost?: string;
}

interface IStoryFormValues {
  title: string;
  description: string;
  imageUrl: string;
  linkToPost: string;
}

/**
 * Shared state and validation behavior for Add/Edit story modals.
 */
export const useStoryForm = (initialValues?: IUseStoryFormInitialValues) => {
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

  const reset = (values?: IUseStoryFormInitialValues): void => {
    setTitle(values?.title || '');
    setDescription(values?.description || '');
    setImageUrl(values?.imageUrl || '');
    setLinkToPost(values?.linkToPost || '');
    setErrors({});
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
    reset,
  };
};
