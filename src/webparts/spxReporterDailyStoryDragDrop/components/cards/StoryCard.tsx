import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory, CardVariant, CardLayout } from '../types';
import { detectStorySource, READ_TIME_LABEL, StorySource, formatPublishDate } from './storyHelpers';
import styles from './StoryCard.module.scss';

export interface IStoryCardProps {
  story: IStory;
  variant: CardVariant;
  cardLayout?: CardLayout;
  showImage?: boolean;
  onRemove?: () => void;
  isDragging?: boolean;
  featured?: boolean;
}

const StoryCard: React.FC<IStoryCardProps> = ({
  story,
  variant,
  cardLayout,
  showImage = true,
  onRemove,
  isDragging = false,
  featured = false
}) => {
  // Resolve the render layout; fall back to legacy behaviour so existing
  // callers / tests (which only pass `variant`) still render sensibly.
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
    isDragging ? styles.dragging : '',
    featured ? styles.featured : ''
  ]
    .filter(Boolean)
    .join(' ');

  const handleRemove = (e: React.MouseEvent): void => {
    e.stopPropagation();
    if (onRemove) onRemove();
  };

  // --- shared fragments ------------------------------------------------
  const grip = (
    <span className={styles.grip} aria-hidden="true">
      <Icon iconName="GripperDotsVertical" />
    </span>
  );

  const featuredBadge = featured ? (
    <span className={styles.featuredBadge}>
      <Icon iconName="FavoriteStarFill" /> Featured
    </span>
  ) : null;

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
    // Commented out source badge (e.g. google.com / LinkedIn / External)
    return null;
  };

  const metaRow = (tone: 'onImage' | 'onLight'): React.ReactNode => (
    <div className={styles.metaRow}>
      <span className={styles.metaDate}>{formatPublishDate(story.date)}</span>
      <span className={styles.metaRead}>{'• ' + READ_TIME_LABEL}</span>
      {sourceBadge(tone)}
    </div>
  );

  // --- full-image ------------------------------------------------------
  if (layout === 'full-image') {
    return (
      <div className={cardClassName}>
        <div className={styles.overlayTopLeft}>
          {grip}
          {featuredBadge}
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

  // --- thumbnail-text (horizontal) ------------------------------------
  if (layout === 'thumbnail-text') {
    return (
      <div className={cardClassName}>
        <div className={styles.overlayTopLeft}>
          {grip}
          {featuredBadge}
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

  // --- text-only -------------------------------------------------------
  return (
    <div className={cardClassName}>
      <div className={styles.overlayTopLeft}>
        {grip}
        {featuredBadge}
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

// Re-export so callers that only need the union don't reach into types.
export type { StorySource };
