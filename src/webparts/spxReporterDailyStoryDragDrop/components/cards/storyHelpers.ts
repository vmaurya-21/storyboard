/** Where a story's post link points. */
export type StorySource = 'internal' | 'linkedin' | 'external';

/** Result of classifying a post link with {@link detectStorySource}. */
export interface IStorySourceInfo {
  /** The classification. */
  source: StorySource;
  /** Hostname the classification was derived from, when the URL parsed. */
  domain?: string;
}

/**
 * Classifies a story's post link as internal, LinkedIn or external.
 *
 * The rules mirror the Reporter Daily reference:
 *
 * - any `*.linkedin.com` host is `linkedin`, reported with the canonical
 *   domain `linkedin.com` rather than the subdomain;
 * - `*.whitecase.com` is `internal`, unless the host starts with `external.`
 *   or contains `external-`, which marks a partner-facing site;
 * - anything else is `external`.
 *
 * A missing, blank or placeholder (`'#'`) link, or one that fails to parse,
 * is treated as `internal` with no domain — the safe default, since those
 * links never leave the tenant.
 *
 * @param url - Absolute post URL. Relative URLs do not parse and fall back to
 * `internal`.
 * @returns The classification and, when available, the hostname behind it.
 *
 * @example
 * ```ts
 * detectStorySource('https://www.linkedin.com/posts/x');
 * // => { source: 'linkedin', domain: 'linkedin.com' }
 *
 * detectStorySource('https://external.whitecase.com/a');
 * // => { source: 'external', domain: 'external.whitecase.com' }
 * ```
 */
export const detectStorySource = (url?: string): IStorySourceInfo => {
  if (!url || url === '#' || url.trim() === '') {
    return { source: 'internal' };
  }
  try {
    const hostname = new URL(url).hostname.toLowerCase();

    if (hostname === 'linkedin.com' || hostname.endsWith('.linkedin.com')) {
      return { source: 'linkedin', domain: 'linkedin.com' };
    }

    if (hostname === 'whitecase.com' || hostname.endsWith('.whitecase.com')) {
      if (hostname.startsWith('external.') || hostname.indexOf('external-') !== -1) {
        return { source: 'external', domain: hostname };
      }
      return { source: 'internal', domain: hostname };
    }

    return { source: 'external', domain: hostname };
  } catch {
    return { source: 'internal' };
  }
};

/**
 * Fixed reading-time label shown on cards.
 *
 * The reference design displays a constant rather than estimating per story,
 * because the list holds links rather than article bodies.
 */
export const READ_TIME_LABEL = '2 min read';

/**
 * Formats a publish date the way the cards display it: short month plus an
 * ordinal day, e.g. `Mar 3rd`.
 *
 * @param dateVal - A `Date`, or any string `Date` can parse.
 * @returns The formatted label. An empty string when `dateVal` is absent;
 * the original string when it was given but could not be parsed, so a
 * pre-formatted value from SharePoint passes through untouched rather than
 * rendering as `Invalid Date`.
 *
 * @example
 * ```ts
 * formatPublishDate('2026-03-03'); // => 'Mar 3rd'
 * formatPublishDate('sometime');   // => 'sometime'
 * ```
 */
export const formatPublishDate = (dateVal?: string | Date): string => {
  if (!dateVal) return '';

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
