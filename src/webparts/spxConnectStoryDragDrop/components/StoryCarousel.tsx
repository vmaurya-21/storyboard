import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { useDroppable } from '@dnd-kit/core';
import { Icon } from '@fluentui/react/lib/Icon';
import { CardLayout, IStory } from './types';
import styles from './StoryCarousel.module.scss';
import { formatPublishDate } from './cards/storyHelpers';

/**
 * How a carousel slide renders its story.
 *
 * A narrower set than `CardLayout` — slides are always tall enough for
 * an image, so `text-only` is not offered.
 */
export type CarouselDisplayMode = 'full-image' | 'thumbnail-text';

/**
 * Narrows a stored card layout to one the carousel can render.
 *
 * `slotLayoutPreferences` is a flat map keyed by bare slot id, and the grid
 * presets share ids `slot-1`..`slot-5` with this one — so a `text-only`
 * preference set under another preset can reach a slide. Anything that is not
 * `thumbnail-text` renders as `full-image`, which is also the default.
 *
 * @param mode - Stored preference, if any.
 * @returns The display mode to render.
 */
export const toCarouselMode = (mode: CardLayout | undefined): CarouselDisplayMode =>
  mode === 'thumbnail-text' ? 'thumbnail-text' : 'full-image';

/** Minimum horizontal travel, in px, before a touch counts as a swipe. */
const SWIPE_THRESHOLD = 40;

/** Props for {@link StoryCarousel}. */
export interface IStoryCarouselProps {
  /** Slot ids to render, in order. Each becomes one slide. */
  slotIds: string[];
  /** Board state: which story sits in which slot. */
  slotStories: { [key: string]: IStory | undefined };
  /**
   * Per-slot card layout overrides, keyed by slot id.
   *
   * Typed as the full `CardLayout` because the board stores one map across
   * every preset; {@link toCarouselMode} narrows each value on read.
   */
  slotLayoutPreferences?: { [key: string]: CardLayout };
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
 * Byline row: author and publish date, separated by a bullet.
 *
 * Reference: `<span>{author}</span><span>•</span><span>{date}</span>`. The
 * author is optional here — stories added before Created By was projected, and
 * board states published before that, carry no byline — so the separator is
 * dropped along with it rather than leaving a stray bullet.
 */
const Meta: React.FC<{ story: IStory; className: string }> = ({ story, className }) => (
  <div className={className}>
    {story.author && (
      <>
        <span>{story.author}</span>
        <span aria-hidden="true">•</span>
      </>
    )}
    <span>{formatPublishDate(story.date)}</span>
  </div>
);

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
          <Meta story={story} className={styles.splitMeta} />
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
        <Meta story={story} className={styles.meta} />
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
  /** 1-based position of this slide, for the accessible label. */
  index: number;
  /** Total number of real slides, for the accessible label. */
  total: number;
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
  canChangeLayout,
  index,
  total
}) => {
  const { setNodeRef, isOver } = useDroppable({ id: slotId });

  return (
    <div
      ref={setNodeRef}
      className={`${styles.slide} ${isOver ? styles.isOver : ''}`}
      // Reference: CarouselItem is `role="group" aria-roledescription="slide"`.
      role="group"
      aria-roledescription="slide"
      aria-label={`${index} of ${total}`}
    >
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
  /** Clientside X of the touch that may become a swipe, or `null` when idle. */
  const touchStartXRef = useRef<number | null>(null);

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

  /**
   * Moves one slide, onto a clone when stepping off either end.
   *
   * The clamp is what keeps a fast clicker out of trouble: the track only has
   * positions `0..count + 1`, and those two ends are the only ones
   * `handleTransitionEnd` knows how to wrap. Without it a second click landing
   * while the first transition was still running pushed `pos` past the clone,
   * translating the track into empty space with no pending event to recover
   * it. Clamping rather than locking keeps rapid stepping responsive — the
   * extra click is absorbed, not queued behind a 400ms animation.
   *
   * @param delta - `1` to advance, `-1` to go back.
   */
  const step = (delta: 1 | -1): void => {
    setAnimate(true);
    setPos(p => Math.min(count + 1, Math.max(0, p + delta)));
  };

  /** Advances one slide backwards, onto the leading clone when already first. */
  const goPrev = (): void => step(-1);

  /** Advances one slide forwards, onto the trailing clone when already last. */
  const goNext = (): void => step(1);

  /**
   * Arrow-key navigation.
   *
   * Reference: the Carousel root binds this on capture, so it works whenever
   * focus is anywhere inside — in practice the nav buttons.
   *
   * @param e - Keyboard event from the carousel root.
   */
  const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'ArrowLeft') {
      e.preventDefault();
      goPrev();
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      goNext();
    }
  };

  /** Records where a touch began, so `handleTouchEnd` can measure the travel. */
  const handleTouchStart = (e: React.TouchEvent<HTMLDivElement>): void => {
    touchStartXRef.current = e.touches[0].clientX;
  };

  /**
   * Turns a horizontal drag into a slide change.
   *
   * The nav buttons are hidden below the `md` breakpoint, matching the
   * reference — which can afford that because Embla handles pointer dragging
   * for it. This is the equivalent: without it the carousel has no reachable
   * control at all on a phone, leaving slides 2..n unreachable.
   *
   * Nothing is prevented or captured, so a vertical scroll that happens to
   * start on a slide still scrolls the page.
   *
   * @param e - Touch event from the viewport.
   */
  const handleTouchEnd = (e: React.TouchEvent<HTMLDivElement>): void => {
    const startX = touchStartXRef.current;
    touchStartXRef.current = null;
    if (startX === null) return;

    const deltaX = e.changedTouches[0].clientX - startX;
    if (Math.abs(deltaX) < SWIPE_THRESHOLD) return;

    step(deltaX < 0 ? 1 : -1);
  };

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
    const mode = toCarouselMode(slotLayoutPreferences[slotId]);
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
    // Reference: the Carousel root is `role="region"
    // aria-roledescription="carousel"` with arrow keys bound on capture. The
    // label is added here because an unnamed region is not exposed as a
    // landmark.
    <div
      className={styles.carousel}
      role="region"
      aria-roledescription="carousel"
      aria-label="Story carousel"
      onKeyDownCapture={handleKeyDown}
    >
      <div
        className={styles.viewport}
        onTouchStart={handleTouchStart}
        onTouchEnd={handleTouchEnd}
      >
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
          {slotIds.map((slotId, i) => {
            const mode = toCarouselMode(slotLayoutPreferences[slotId]);
            const story = slotStories[slotId];
            return (
              <CarouselSlide
                key={slotId}
                slotId={slotId}
                story={story}
                displayMode={mode}
                index={i + 1}
                total={count}
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
