import { detectStorySource, formatPublishDate, READ_TIME_LABEL } from '../../cards/storyHelpers';

describe('storyHelpers', () => {
  it('classifies LinkedIn and external/internal domains', () => {
    expect(detectStorySource('https://www.linkedin.com/posts/x')).toEqual({ source: 'linkedin', domain: 'linkedin.com' });
    expect(detectStorySource('https://whitecase.sharepoint.com/sites/comms')).toEqual({ source: 'internal', domain: 'whitecase.sharepoint.com' });
    expect(detectStorySource('https://contoso.sharepoint.com/sites/comms')).toEqual({ source: 'external', domain: 'contoso.sharepoint.com' });
    expect(detectStorySource('https://whitecase.com/a')).toEqual({ source: 'external', domain: 'whitecase.com' });
    expect(detectStorySource('https://news.example.com/a')).toEqual({ source: 'external', domain: 'news.example.com' });
  });

  it('falls back to internal for blank or invalid links', () => {
    expect(detectStorySource()).toEqual({ source: 'internal' });
    expect(detectStorySource('')).toEqual({ source: 'internal' });
    expect(detectStorySource('#')).toEqual({ source: 'internal' });
    expect(detectStorySource('not-a-url')).toEqual({ source: 'internal' });
  });

  it('formats publish dates and keeps invalid strings unchanged', () => {
    expect(formatPublishDate('2026-03-01')).toBe('Mar 1st');
    expect(formatPublishDate('2026-03-02')).toBe('Mar 2nd');
    expect(formatPublishDate('2026-03-03')).toBe('Mar 3rd');
    expect(formatPublishDate('2026-03-04')).toBe('Mar 4th');
    expect(formatPublishDate('bad-date')).toBe('bad-date');
    expect(formatPublishDate(new Date('invalid'))).toBe('');
    expect(formatPublishDate(undefined)).toBe('');
  });

  it('exposes the fixed read-time label', () => {
    expect(READ_TIME_LABEL).toBe('2 min read');
  });
});
