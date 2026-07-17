import * as React from 'react';
import { useState, useEffect } from 'react';
import { useDroppable } from '@dnd-kit/core';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory } from './types';
import styles from './StoryCarousel.module.scss';
import { formatPublishDate } from './cards/storyHelpers';

export type CarouselDisplayMode = 'full-image' | 'thumbnail-text';

export interface IStoryCarouselProps {
  slotIds: string[];
  slotStories: { [key: string]: IStory | undefined };
  slotLayoutPreferences?: { [key: string]: CarouselDisplayMode };
  onSlotLayoutChange?: (slotId: string, mode: CarouselDisplayMode) => void;
  onRemoveStory?: (slotId: string) => void;
  isAdminMode?: boolean;
  featuredStoryId?: string;
}

// ---- Filled card content (full-image / split thumbnail-text) ----------
interface ICardContentProps {
  story: IStory;
  displayMode: CarouselDisplayMode;
}

const CardContent: React.FC<ICardContentProps> = ({ story, displayMode }) => {
  if (displayMode === 'thumbnail-text') {
    return (
      <div className={styles.split}>
        <div className={styles.splitText}>
          <h3 className={styles.splitTitle}>{story.title}</h3>
          {story.description && <p className={styles.splitDesc}>{story.description}</p>}
          <div className={styles.splitMeta}>
            <span>{formatPublishDate(story.date)}</span>
          </div>
          <a
            href={story.linkToPost || '#'}
            className={styles.readMore}
            onClick={(e) => e.stopPropagation()}
          >
            Read more →
          </a>
        </div>
        <div className={styles.splitImageWrap}>
          {story.imageUrl && (
            <img src={story.imageUrl} alt={story.title} className={styles.splitImage} />
          )}
        </div>
      </div>
    );
  }

  // full-image
  return (
    <div className={styles.fullImage}>
      {story.imageUrl && (
        <img src={story.imageUrl} alt={story.title} className={styles.image} />
      )}
      <div className={styles.gradient} />
      <div className={styles.content}>
        <h3 className={styles.title}>{story.title}</h3>
        {story.description && <p className={styles.desc}>{story.description}</p>}
        <div className={styles.meta}>
          <span>{formatPublishDate(story.date)}</span>
        </div>
      </div>
    </div>
  );
};

// ---- Admin controls (layout menu + remove) ----------------------------
interface IControlsProps {
  slotId: string;
  displayMode: CarouselDisplayMode;
  menuOpen: boolean;
  onToggleMenu: () => void;
  onSelectMode: (mode: CarouselDisplayMode) => void;
  onRemove?: () => void;
  canChangeLayout: boolean;
}

const Controls: React.FC<IControlsProps> = ({
  displayMode,
  menuOpen,
  onToggleMenu,
  onSelectMode,
  onRemove,
  canChangeLayout
}) => (
  <div
    className={styles.controls}
    onPointerDown={(e) => e.stopPropagation()}
  >
    {canChangeLayout && (
      <button
        type="button"
        className={styles.ctrlBtn}
        onClick={(e) => { e.stopPropagation(); onToggleMenu(); }}
        aria-label="Change slide layout"
        title="Change slide layout"
      >
        <Icon iconName="More" />
      </button>
    )}
    {onRemove && (
      <button
        type="button"
        className={`${styles.ctrlBtn} ${styles.ctrlRemove}`}
        onClick={(e) => { e.stopPropagation(); onRemove(); }}
        aria-label="Remove story"
        title="Remove story"
      >
        <Icon iconName="Cancel" />
      </button>
    )}

    {menuOpen && canChangeLayout && (
      <div className={styles.menu} role="menu">
        <button
          type="button"
          className={`${styles.menuItem} ${displayMode === 'full-image' ? styles.menuItemActive : ''}`}
          onClick={(e) => { e.stopPropagation(); onSelectMode('full-image'); }}
        >
          Full Image
        </button>
        <button
          type="button"
          className={`${styles.menuItem} ${displayMode === 'thumbnail-text' ? styles.menuItemActive : ''}`}
          onClick={(e) => { e.stopPropagation(); onSelectMode('thumbnail-text'); }}
        >
          Thumbnail with Text
        </button>
      </div>
    )}
  </div>
);

// ---- Filled card ------------------------------------------------------
interface ICardProps {
  story: IStory;
  displayMode: CarouselDisplayMode;
  controls: React.ReactNode;
  featured?: boolean;
}

// Cards are droppable targets (via the slide) but not themselves draggable:
// dnd-kit moves the node in place, which the overflow-hidden viewport would
// clip. Removal back to the Available list is handled by the X control.
const CarouselCard: React.FC<ICardProps> = ({ story, displayMode, controls, featured }) => (
  <div className={styles.card}>
    {featured && (
      <div className={styles.featuredBadge}>
        <Icon iconName="FavoriteStarFill" />
        <span>Featured</span>
      </div>
    )}
    {controls}
    <CardContent story={story} displayMode={displayMode} />
  </div>
);

// ---- One slide (droppable) -------------------------------------------
interface ISlideProps {
  slotId: string;
  story: IStory | undefined;
  displayMode: CarouselDisplayMode;
  isAdminMode: boolean;
  menuOpen: boolean;
  onToggleMenu: () => void;
  onSelectMode: (mode: CarouselDisplayMode) => void;
  onRemove?: () => void;
  canChangeLayout: boolean;
  featured?: boolean;
}

const CarouselSlide: React.FC<ISlideProps> = ({
  slotId,
  story,
  displayMode,
  isAdminMode,
  menuOpen,
  onToggleMenu,
  onSelectMode,
  onRemove,
  canChangeLayout,
  featured
}) => {
  const { setNodeRef, isOver } = useDroppable({ id: slotId });

  return (
    <div ref={setNodeRef} className={`${styles.slide} ${isOver ? styles.isOver : ''}`}>
      {story ? (
        <CarouselCard
          story={story}
          displayMode={displayMode}
          featured={featured}
          controls={
            isAdminMode ? (
              <Controls
                slotId={slotId}
                displayMode={displayMode}
                menuOpen={menuOpen}
                onToggleMenu={onToggleMenu}
                onSelectMode={onSelectMode}
                onRemove={onRemove}
                canChangeLayout={canChangeLayout}
              />
            ) : null
          }
        />
      ) : (
        <div className={styles.empty}>
          <p className={styles.emptyText}>
            {isAdminMode ? 'Drop a story here' : 'No story'}
          </p>
        </div>
      )}
    </div>
  );
};

// ---- Carousel ---------------------------------------------------------
const StoryCarousel: React.FC<IStoryCarouselProps> = ({
  slotIds,
  slotStories,
  slotLayoutPreferences = {},
  onSlotLayoutChange,
  onRemoveStory,
  isAdminMode = true,
  featuredStoryId
}) => {
  const count = slotIds.length;
  const [index, setIndex] = useState(0);
  const [menuOpenSlot, setMenuOpenSlot] = useState<string | null>(null);

  // Keep the active index valid if the slot count ever changes.
  useEffect(() => {
    if (index > count - 1) setIndex(count > 0 ? count - 1 : 0);
  }, [count, index]);

  // Close the layout menu on any outside pointer.
  useEffect(() => {
    if (!menuOpenSlot) return undefined;
    const onOutside = (e: Event): void => {
      const target = e.target as HTMLElement;
      if (!target.closest(`.${styles.controls}`)) {
        setMenuOpenSlot(null);
      }
    };
    document.addEventListener('mousedown', onOutside);
    document.addEventListener('touchstart', onOutside);
    return () => {
      document.removeEventListener('mousedown', onOutside);
      document.removeEventListener('touchstart', onOutside);
    };
  }, [menuOpenSlot]);

  if (count === 0) return null;

  const goPrev = (): void => setIndex(i => (i > 0 ? i - 1 : count - 1));
  const goNext = (): void => setIndex(i => (i < count - 1 ? i + 1 : 0));

  return (
    <div className={styles.carousel}>
      <div className={styles.viewport}>
        <div
          className={styles.track}
          style={{ transform: `translateX(-${index * 100}%)` }}
        >
          {slotIds.map(slotId => {
            const mode: CarouselDisplayMode = slotLayoutPreferences[slotId] || 'full-image';
            const story = slotStories[slotId];
            return (
              <CarouselSlide
                key={slotId}
                slotId={slotId}
                story={story}
                displayMode={mode}
                featured={!!story && !!featuredStoryId && story.id === featuredStoryId}
                isAdminMode={isAdminMode}
                menuOpen={menuOpenSlot === slotId}
                onToggleMenu={() =>
                  setMenuOpenSlot(prev => (prev === slotId ? null : slotId))
                }
                onSelectMode={(m) => {
                  if (onSlotLayoutChange) onSlotLayoutChange(slotId, m);
                  setMenuOpenSlot(null);
                }}
                onRemove={onRemoveStory ? () => onRemoveStory(slotId) : undefined}
                canChangeLayout={!!onSlotLayoutChange}
              />
            );
          })}
        </div>
      </div>

      <button
        type="button"
        className={styles.navPrev}
        onClick={goPrev}
        aria-label="Previous slide"
      >
        <Icon iconName="ChevronLeft" />
      </button>
      <button
        type="button"
        className={styles.navNext}
        onClick={goNext}
        aria-label="Next slide"
      >
        <Icon iconName="ChevronRight" />
      </button>
    </div>
  );
};

export default StoryCarousel;
