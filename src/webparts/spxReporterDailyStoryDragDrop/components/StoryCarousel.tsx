import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { useDroppable } from '@dnd-kit/core';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory } from './types';
import styles from './StoryCarousel.module.scss';
import { formatPublishDate } from './cards/storyHelpers';

/**
 * How a carousel slide renders its story.
 *
 * A narrower set than `CardLayout` — slides are always tall enough for
 * an image, so `text-only` is not offered.
 */
export type CarouselDisplayMode = 'full-image' | 'thumbnail-text';

/** Props for {@link StoryCarousel}. */
export interface IStoryCarouselProps {
  /** Slot ids to render, in order. Each becomes one slide. */
  slotIds: string[];
  /** Board state: which story sits in which slot. */
  slotStories: { [key: string]: IStory | undefined };
  /** Per-slot display mode overrides, keyed by slot id. */
  slotLayoutPreferences?: { [key: string]: CarouselDisplayMode };
  /** Called when the editor picks a different display mode for a slide. */
  onSlotLayoutChange?: (slotId: string, mode: CarouselDisplayMode) => void;
  /** Called with a slot id when the editor clears that slide. */
  onRemoveStory?: (slotId: string) => void;
  /** Whether to show editing controls. Defaults to `true`. */
  isAdminMode?: boolean;
}

/** Props for {@link CardContent}. */
interface ICardContentProps {
  /** Story to render. */
  story: IStory;
  /** Which of the two content layouts to use. */
  displayMode: CarouselDisplayMode;
}

/**
 * Inner content of a filled slide.
 *
 * `thumbnail-text` splits text and image side by side; `full-image` overlays
 * the text on the image behind a gradient scrim.
 */
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

/** Props for {@link Controls}. */
interface IControlsProps {
  /** Slot the controls act on. */
  slotId: string;
  /** Current display mode, used to mark the active menu item. */
  displayMode: CarouselDisplayMode;
  /** Whether this slide's layout menu is open. */
  menuOpen: boolean;
  /** Toggles the layout menu. */
  onToggleMenu: () => void;
  /** Called with the chosen display mode. */
  onSelectMode: (mode: CarouselDisplayMode) => void;
  /** Called when the slide is cleared. Omit to hide the remove button. */
  onRemove?: () => void;
  /** Whether the layout menu is offered at all. */
  canChangeLayout: boolean;
}

/**
 * Editor overlay for a slide: a layout menu and a remove button.
 *
 * Pointer events are stopped at the root so using a control never starts a
 * drag on the card beneath it.
 */
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

/** Props for {@link CarouselCard}. */
interface ICardProps {
  /** Story to render. */
  story: IStory;
  /** Which content layout to use. */
  displayMode: CarouselDisplayMode;
  /** Editor overlay to render above the content, or `null` for read-only. */
  controls: React.ReactNode;
}

/**
 * Card shell for a filled slide.
 *
 * Cards are drop targets by way of their slide, but are not themselves
 * draggable: dnd-kit moves a dragged node in place, and the viewport's
 * `overflow: hidden` would clip it. Removal is handled by the overlay's X
 * control instead.
 */
const CarouselCard: React.FC<ICardProps> = ({ story, displayMode, controls }) => (
  <div className={styles.card}>
    {controls}
    <CardContent story={story} displayMode={displayMode} />
  </div>
);

/** Props for {@link CarouselSlide}. */
interface ISlideProps {
  /** Slot id; also the dnd-kit droppable id for this slide. */
  slotId: string;
  /** Story in the slot, or `undefined` for the empty state. */
  story: IStory | undefined;
  /** Which content layout to use. */
  displayMode: CarouselDisplayMode;
  /** Whether to render editing controls. */
  isAdminMode: boolean;
  /** Whether this slide's layout menu is open. */
  menuOpen: boolean;
  /** Toggles the layout menu. */
  onToggleMenu: () => void;
  /** Called with the chosen display mode. */
  onSelectMode: (mode: CarouselDisplayMode) => void;
  /** Called when the slide is cleared. */
  onRemove?: () => void;
  /** Whether the layout menu is offered. */
  canChangeLayout: boolean;
}

/**
 * One full-width slide, and a drop target for the slot it represents.
 *
 * Highlights while a dragged story hovers over it.
 */
const CarouselSlide: React.FC<ISlideProps> = ({
  slotId,
  story,
  displayMode,
  isAdminMode,
  menuOpen,
  onToggleMenu,
  onSelectMode,
  onRemove,
  canChangeLayout
}) => {
  const { setNodeRef, isOver } = useDroppable({ id: slotId });

  return (
    <div ref={setNodeRef} className={`${styles.slide} ${isOver ? styles.isOver : ''}`}>
      {story ? (
        <CarouselCard
          story={story}
          displayMode={displayMode}
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

/**
 * Full-width carousel used by the Connect Homepage layout.
 *
 * Looping is seamless in both directions: the track carries a clone of the
 * last slide before the first and a clone of the first after the last, so
 * advancing past either end keeps moving the same way instead of rewinding.
 * Once the transition onto a clone finishes, the track jumps to the real
 * slide with animation suppressed — see the effect on `animate` for why that
 * hand-off has to wait a painted frame.
 *
 * Slides are drop targets, so a story can be dragged onto whichever slide is
 * on screen; the clones are inert and never receive drops.
 */
const StoryCarousel: React.FC<IStoryCarouselProps> = ({
  slotIds,
  slotStories,
  slotLayoutPreferences = {},
  onSlotLayoutChange,
  onRemoveStory,
  isAdminMode = true
}) => {
  const count = slotIds.length;

  /**
   * Index into the track, including clones: real slides occupy `1..count`,
   * with the clone of the last slide at `0` and the clone of the first at
   * `count + 1`.
   */
  const [pos, setPos] = useState(1);
  /** Whether the track's CSS transition is active. Cleared for seamless jumps. */
  const [animate, setAnimate] = useState(true);
  /** Slot id whose layout menu is open, or `null` when none is. */
  const [menuOpenSlot, setMenuOpenSlot] = useState<string | null>(null);
  /** The track element, read to force a style flush before re-enabling animation. */
  const trackRef = useRef<HTMLDivElement>(null);

  // Return to the first real slide whenever the slot set changes, without
  // animating the jump.
  useEffect(() => {
    setAnimate(false);
    setPos(1);
  }, [count]);

  // Restore the transition only after the un-animated snap has been painted.
  // Reading `offsetHeight` forces a synchronous style and layout recalc so the
  // new transform commits while the transition is still off; the double
  // `requestAnimationFrame` then waits a full painted frame. Re-enabling any
  // earlier lets the browser interpolate from the clone back to the real
  // slide, which reads as the whole track rewinding.
  useEffect(() => {
    if (animate) return undefined;
    if (trackRef.current) void trackRef.current.offsetHeight;
    let inner = 0;
    const outer = requestAnimationFrame(() => {
      inner = requestAnimationFrame(() => setAnimate(true));
    });
    return () => {
      cancelAnimationFrame(outer);
      if (inner) cancelAnimationFrame(inner);
    };
  }, [animate]);

  // Close the layout menu on any pointer landing outside a controls overlay.
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

  /** Advances one slide backwards, onto the leading clone when already first. */
  const goPrev = (): void => { setAnimate(true); setPos(p => p - 1); };

  /** Advances one slide forwards, onto the trailing clone when already last. */
  const goNext = (): void => { setAnimate(true); setPos(p => p + 1); };

  /**
   * Completes a wrap-around once a slide transition ends.
   *
   * When the track has landed on a clone, it jumps to that clone's real
   * counterpart with animation suppressed, which is what makes the loop
   * seamless. Events from descendants and from non-`transform` properties are
   * ignored so an unrelated transition cannot trigger a jump.
   *
   * @param e - The transition event from the track.
   */
  const handleTransitionEnd = (e: React.TransitionEvent<HTMLDivElement>): void => {
    if (e.target !== e.currentTarget || e.propertyName !== 'transform') return;
    if (pos === count + 1) {
      setAnimate(false);
      setPos(1);
    } else if (pos === 0) {
      setAnimate(false);
      setPos(count);
    }
  };

  /**
   * Renders one of the two loop clones.
   *
   * Clones are presentational only: they carry no droppable slot id and no
   * editing controls, so dnd-kit never sees a duplicate slot and a drop can
   * never land on one. They are hidden from assistive technology and opt out
   * of pointer events.
   *
   * @param slotId - Slot whose content to mirror.
   * @param key - React key, since both clones mirror an existing slot id.
   * @returns The clone slide.
   */
  const renderClone = (slotId: string, key: string): JSX.Element => {
    const mode: CarouselDisplayMode = slotLayoutPreferences[slotId] || 'full-image';
    const story = slotStories[slotId];
    return (
      <div key={key} className={styles.slide} aria-hidden="true" style={{ pointerEvents: 'none' }}>
        {story ? (
          <CarouselCard story={story} displayMode={mode} controls={null} />
        ) : (
          <div className={styles.empty}>
            <p className={styles.emptyText}>{isAdminMode ? 'Drop a story here' : 'No story'}</p>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={styles.carousel}>
      <div className={styles.viewport}>
        <div
          ref={trackRef}
          className={styles.track}
          style={{
            transform: `translateX(-${pos * 100}%)`,
            transition: animate ? undefined : 'none'
          }}
          onTransitionEnd={handleTransitionEnd}
        >
          {renderClone(slotIds[count - 1], 'clone-last')}
          {slotIds.map(slotId => {
            const mode: CarouselDisplayMode = slotLayoutPreferences[slotId] || 'full-image';
            const story = slotStories[slotId];
            return (
              <CarouselSlide
                key={slotId}
                slotId={slotId}
                story={story}
                displayMode={mode}
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
          {renderClone(slotIds[0], 'clone-first')}
        </div>
      </div>

      <button
        type="button"
        className={styles.navPrev}
        onClick={goPrev}
        aria-label="Previous slide"
      >
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="m12 19-7-7 7-7"/>
          <path d="M19 12H5"/>
        </svg>
      </button>
      <button
        type="button"
        className={styles.navNext}
        onClick={goNext}
        aria-label="Next slide"
      >
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M5 12h14"/>
          <path d="m12 5 7 7-7 7"/>
        </svg>
      </button>
    </div>
  );
};

export default StoryCarousel;
