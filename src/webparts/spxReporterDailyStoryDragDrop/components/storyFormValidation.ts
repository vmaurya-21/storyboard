/** Longest title SharePoint's single-line text column accepts. */
export const TITLE_MAX_LENGTH = 255;

/** Required fields for story forms. */
export type StoryRequiredField = 'title' | 'imageUrl' | 'linkToPost';

/**
 * Tests whether a string parses as an absolute URL.
 */
export const isValidUrl = (url: string): boolean => {
  try {
    // eslint-disable-next-line no-new
    new URL(url);
    return true;
  } catch {
    return false;
  }
};

/**
 * Validates one story form field.
 */
export const validateStoryField = (field: StoryRequiredField, value: string): string => {
  switch (field) {
    case 'title':
      if (!value.trim()) return 'Title is required';
      if (value.length > TITLE_MAX_LENGTH) return `Title must be ${TITLE_MAX_LENGTH} characters or less`;
      return '';
    case 'imageUrl':
      if (!value.trim()) return 'Image URL is required';
      if (!isValidUrl(value)) return 'Please enter a valid URL';
      return '';
    case 'linkToPost':
      if (!value.trim()) return 'Link to Post is required';
      if (!isValidUrl(value)) return 'Please enter a valid URL';
      return '';
    default:
      return '';
  }
};
