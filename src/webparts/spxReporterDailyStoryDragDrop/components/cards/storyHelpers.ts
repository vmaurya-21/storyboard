// Small presentational helpers shared by the board cards, mirroring the
// reference app's source detection + read-time chrome.

export type StorySource = 'internal' | 'linkedin' | 'external';

export interface IStorySourceInfo {
  source: StorySource;
  domain?: string;
}

// Derive a story's "source" from its post link, matching the reference
// detectStorySource(): LinkedIn / external / internal.
export const detectStorySource = (url?: string): IStorySourceInfo => {
  if (!url || url === '#') {
    return { source: 'internal' };
  }
  try {
    const hostname = new URL(url).hostname.toLowerCase();

    if (hostname.indexOf('linkedin.com') !== -1) {
      return { source: 'linkedin', domain: 'linkedin.com' };
    }
    if (hostname.indexOf('whitecase.com') !== -1) {
      if (hostname.indexOf('external.') === 0 || hostname.indexOf('external-') !== -1) {
        return { source: 'external', domain: hostname };
      }
      return { source: 'internal', domain: hostname };
    }
    if (hostname.indexOf('unsplash.com') !== -1) {
      return { source: 'external', domain: 'unsplash.com' };
    }
    return { source: 'external', domain: hostname };
  } catch {
    return { source: 'internal' };
  }
};

// The reference shows a constant "2 min read" estimate on every card.
export const READ_TIME_LABEL = '2 min read';

// Format date into reference style (e.g., Nov 22nd, Sep 14th)
export const formatPublishDate = (dateVal?: string | Date): string => {
  if (!dateVal) return '';
  
  // Try to parse the string to a date. 
  // If dateVal is already "Nov 22nd" or similar, new Date(dateVal) will return an invalid date.
  const parsed = new Date(dateVal);
  if (isNaN(parsed.getTime())) {
    return typeof dateVal === 'string' ? dateVal : '';
  }
  
  const month = parsed.toLocaleDateString('en-US', { month: 'short' });
  const day = parsed.getDate();
  
  let suffix = 'th';
  if (day === 1 || day === 21 || day === 31) suffix = 'st';
  else if (day === 2 || day === 22) suffix = 'nd';
  else if (day === 3 || day === 23) suffix = 'rd';
  
  return `${month} ${day}${suffix}`;
};
