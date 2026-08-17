import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory, CardVariant, CardLayout } from '../types';
import { detectStorySource, READ_TIME_LABEL, StorySource, formatPublishDate } from './storyHelpers';
import styles from './StoryCard.module.scss';

/** Props for {@link StoryCard}. */
export interface IStoryCardProps {
  /** Story to render. */
  story: IStory;
  /** Size family, applied as a modifier class. */
  variant: CardVariant;
  /**
   * Content layout. When omitted it is derived from `showImage`:
   * `text-only` if images are suppressed, otherwise `full-image`.
   */
  cardLayout?: CardLayout;
  /** Whether the image may render. Defaults to `true`. */
  showImage?: boolean;
  /** Invoked when the remove control is used. Omit to hide the control. */
  onRemove?: () => void;
  /** Whether the card is mid-drag, which dims it. Defaults to `false`. */
  isDragging?: boolean;
}

/**
 * Presentational card for a single story.
 *
 * Renders one of three content layouts and derives its source badge from the
 * story's post link via {@link detectStorySource}. Holds no state and knows
 * nothing about drag and drop — `CardSlot` supplies that.
 */
const StoryCard: React.FC<IStoryCardProps> = ({
  story,
  variant,
  cardLayout,
  showImage = true,
  onRemove,
  isDragging = false
}) => {
  const layout: CardLayout =
    cardLayout || (showImage === false ? 'text-only' : 'full-image');

  const { source, domain } = detectStorySource(story.linkToPost);

  const layoutClass =
    layout === 'full-image'
      ? styles.fullImage
      : layout === 'thumbnail-text'
      ? styles.thumbnailText
      : styles.textOnlyLayout;

  const cardClassName = [
    styles.storyCard,
    styles[variant],
    layoutClass,
    isDragging ? styles.dragging : ''
  ]
    .filter(Boolean)
    .join(' ');

  const rawSource = (story.source || source || '').toLowerCase();
  const canOpenExternalStory =
    !onRemove &&
    rawSource !== 'internal' &&
    !!story.linkToPost &&
    story.linkToPost !== '#';

  const handleOpenStory = (): void => {
    if (!canOpenExternalStory || typeof window === 'undefined') return;
    window.open(story.linkToPost, '_blank', 'noopener,noreferrer');
  };

  const handleCardKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (!canOpenExternalStory) return;
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      handleOpenStory();
    }
  };

  const handleRemove = (e: React.MouseEvent): void => {
    e.stopPropagation();
    if (onRemove) onRemove();
  };

  const grip = (
    <span className={styles.grip} aria-hidden="true">
      <Icon iconName="GripperDotsVertical" />
    </span>
  );

  const removeButton = onRemove ? (
    <button
      className={styles.removeBtn}
      onClick={handleRemove}
      aria-label="Remove story"
      title="Remove story"
    >
      <Icon iconName="Cancel" />
    </button>
  ) : null;

  const sourceBadge = (tone: 'onImage' | 'onLight'): React.ReactNode => {
    if (rawSource === 'internal') return null;

    const toneClass = tone === 'onImage' ? styles.badgeOnImage : styles.badgeOnLight;
    const isLinkedin = rawSource === 'linkedin';
    const typeClass = isLinkedin ? styles.badgeLinkedin : styles.badgeExternal;

    const displayDomain = domain ? domain.replace(/^www\./, '') : '';
    const badgeLabel = isLinkedin ? 'LinkedIn' : (displayDomain || 'External');

    return (
      <span className={`${styles.sourceBadge} ${toneClass} ${typeClass}`}>
        <Icon iconName={isLinkedin ? 'LinkedInLogo' : 'Globe'} />
        {badgeLabel}
      </span>
    );
  };

  const metaRow = (tone: 'onImage' | 'onLight'): React.ReactNode => (
    <div className={styles.metaRow}>
      <span className={styles.metaDate}>{formatPublishDate(story.date)}</span>
      <span className={styles.metaRead}>{'• ' + READ_TIME_LABEL}</span>
      {sourceBadge(tone)}
    </div>
  );

  if (layout === 'full-image') {
    return (
      <div
        className={cardClassName}
        onClick={canOpenExternalStory ? handleOpenStory : undefined}
        onKeyDown={canOpenExternalStory ? handleCardKeyDown : undefined}
        role={canOpenExternalStory ? 'link' : undefined}
        tabIndex={canOpenExternalStory ? 0 : undefined}
      >
        <div className={styles.overlayTopLeft}>
          {grip}
        </div>
        <div className={styles.overlayTopRight}>{removeButton}</div>

        <div className={styles.imageContainer}>
          {story.imageUrl && (
            <img src={story.imageUrl} alt={story.title} className={styles.storyImage} />
          )}
        </div>

        <div className={styles.caption}>
          <h4 className={styles.captionTitle}>{story.title}</h4>
          {metaRow('onImage')}
        </div>
      </div>
    );
  }

  if (layout === 'thumbnail-text') {
    return (
      <div
        className={cardClassName}
        onClick={canOpenExternalStory ? handleOpenStory : undefined}
        onKeyDown={canOpenExternalStory ? handleCardKeyDown : undefined}
        role={canOpenExternalStory ? 'link' : undefined}
        tabIndex={canOpenExternalStory ? 0 : undefined}
      >
        <div className={styles.overlayTopLeft}>
          {grip}
        </div>
        <div className={styles.overlayTopRight}>{removeButton}</div>
        <div className={styles.thumbRow}>
          <div className={styles.thumb}>
            {story.imageUrl && (
              <img src={story.imageUrl} alt={story.title} className={styles.storyImage} />
            )}
          </div>
          <div className={styles.thumbText}>
            <h4 className={styles.storyTitle}>{story.title}</h4>
            {story.description && (
              <p className={styles.storyDescription}>{story.description}</p>
            )}
          </div>
        </div>
        {metaRow('onLight')}
      </div>
    );
  }

  return (
    <div
      className={cardClassName}
      onClick={canOpenExternalStory ? handleOpenStory : undefined}
      onKeyDown={canOpenExternalStory ? handleCardKeyDown : undefined}
      role={canOpenExternalStory ? 'link' : undefined}
      tabIndex={canOpenExternalStory ? 0 : undefined}
    >
      <div className={styles.overlayTopLeft}>
        {grip}
      </div>
      <div className={styles.overlayTopRight}>{removeButton}</div>

      <div className={styles.textOnlyInner}>
        <div className={styles.textOnlyContent}>
          <h4 className={styles.storyTitle}>{story.title}</h4>
          {story.description && (
            <p className={styles.storyDescription}>{story.description}</p>
          )}
        </div>
        {metaRow('onLight')}
      </div>
    </div>
  );
};

export default StoryCard;

export type { StorySource };
