import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { LayoutType, SlotStoryMap, CardLayout } from '../types';
import { getLayoutConfig } from './layoutConfig';
import CardSlot from '../cards/CardSlot';
import StoryCarousel from '../StoryCarousel';
import styles from './LayoutRenderer.module.scss';

/** Props for {@link LayoutRenderer}. */
export interface ILayoutRendererProps {
  /** Preset to render. */
  layoutType: LayoutType;
  /** Current board state: which story sits in which slot. */
  slotStories: SlotStoryMap;
  /** Called with a slot id when the editor clears that slot. */
  onRemoveStory: (slotId: string) => void;
  /** Per-slot card layout overrides, keyed by slot id. */
  slotLayoutPreferences?: { [key: string]: CardLayout };
  /** Called when the editor picks a different card layout for a slot. */
  onSlotLayoutChange?: (slotId: string, mode: CardLayout) => void;
}

/**
 * Renders the board for the selected layout preset.
 *
 * Three cases, in order:
 *
 * 1. `connectHomepage` delegates to {@link StoryCarousel} — its five slots are
 *    presented as a carousel, not a grid.
 * 2. A preset with no slots renders a "not configured" placeholder.
 * 3. Everything else renders a CSS grid of {@link CardSlot} elements, with
 *    `reporterDaily` and `general` using hand-built column structures rather
 *    than the preset's named grid areas.
 */
const LayoutRenderer: React.FC<ILayoutRendererProps> = ({
  layoutType,
  slotStories,
  onRemoveStory,
  slotLayoutPreferences = {},
  onSlotLayoutChange
}) => {
  const layoutConfig = getLayoutConfig(layoutType);

  if (layoutType === 'connectHomepage') {
    const slotIds = layoutConfig.slots.map(s => s.id);
    return (
      // No casts: the carousel takes the full `CardLayout` map and narrows each
      // value itself, and its `onSlotLayoutChange` only ever emits the two
      // modes it can render — which `handleSlotLayoutChange` already accepts.
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        slotLayoutPreferences={slotLayoutPreferences}
        onSlotLayoutChange={onSlotLayoutChange}
        onRemoveStory={onRemoveStory}
        isAdminMode
      />
    );
  }

  if (layoutConfig.slots.length === 0) {
    return (
      <div className={styles.emptyLayout}>
        <p>This layout is not yet configured.</p>
        <p className={styles.hint}>Please select a different layout.</p>
      </div>
    );
  }

  const gridStyle: React.CSSProperties = {
    gridTemplateColumns: layoutConfig.gridTemplate.columns,
    gridTemplateRows: layoutConfig.gridTemplate.rows,
    gridTemplateAreas: layoutConfig.gridTemplate.areas
  };

  const layoutClassName = layoutType === 'reporterDaily' ? styles.reporterDaily : 
                          layoutType === 'general' ? styles.general : 
                          layoutType === 'highlight' ? styles.highlight : '';

  if (layoutType === 'reporterDaily') {
    return (
      <div className={`${styles.layoutGrid} ${styles.reporterDaily}`}>
        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-1"
              story={slotStories['slot-1'] || undefined}
              variant="medium"
              cardLayout="full-image"
              showImage={true}
              onRemove={() => onRemoveStory('slot-1')}            />
            <CardSlot
              slotId="slot-2"
              story={slotStories['slot-2'] || undefined}
              variant="medium"
              cardLayout="thumbnail-text"
              showImage={true}
              onRemove={() => onRemoveStory('slot-2')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Split Column
          </div>
          <div className={styles.columnHeader}>Split Column</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-3"
              story={slotStories['slot-3'] || undefined}
              variant="small"
              cardLayout="text-only"
              showImage={false}
              onRemove={() => onRemoveStory('slot-3')}            />
            <CardSlot
              slotId="slot-4"
              story={slotStories['slot-4'] || undefined}
              variant="tall"
              cardLayout="full-image"
              showImage={true}
              onRemove={() => onRemoveStory('slot-4')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Full Slot
          </div>
          <div className={styles.columnHeader}>Full Slot</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-5"
              story={slotStories['slot-5'] || undefined}
              variant="large"
              cardLayout="full-image"
              showImage={true}
              onRemove={() => onRemoveStory('slot-5')}            />
          </div>
        </div>
      </div>
    );
  }

  if (layoutType === 'general') {
    return (
      <div className={`${styles.layoutGrid} ${styles.general}`}>
        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="gen-slot-1"
              story={slotStories['gen-slot-1'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-1'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-1')}            />
            <CardSlot
              slotId="gen-slot-4"
              story={slotStories['gen-slot-4'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-4'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-4')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="gen-slot-2"
              story={slotStories['gen-slot-2'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-2'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-2')}            />
            <CardSlot
              slotId="gen-slot-5"
              story={slotStories['gen-slot-5'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-5'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-5')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="gen-slot-3"
              story={slotStories['gen-slot-3'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-3'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-3')}            />
            <CardSlot
              slotId="gen-slot-6"
              story={slotStories['gen-slot-6'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-6'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-6')}            />
          </div>
        </div>
      </div>
    );
  }

  if (layoutType === 'highlight') {
    return (
      <div className={`${styles.layoutGrid} ${styles.highlight}`}>
        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-1"
              story={slotStories['slot-1'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-1'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-1')}            />
            <CardSlot
              slotId="slot-4"
              story={slotStories['slot-4'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-4'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-4')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Two Half Slots
          </div>
          <div className={styles.columnHeader}>Two Half Slots</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-2"
              story={slotStories['slot-2'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-2'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-2')}            />
            <CardSlot
              slotId="slot-5"
              story={slotStories['slot-5'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-5'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-5')}            />
          </div>
        </div>

        <div className={styles.column}>
          <div className={styles.columnDragHandle}>
            <Icon iconName="GripperDotsVertical" className={styles.gripIcon} />
            Full Slot
          </div>
          <div className={styles.columnHeader}>Full Slot</div>
          <div className={styles.cardsContainer}>
            <CardSlot
              slotId="slot-3"
              story={slotStories['slot-3'] || undefined}
              variant="large"
              cardLayout={slotLayoutPreferences['slot-3'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-3')}            />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div 
      className={`${styles.layoutGrid} ${layoutClassName}`}
      style={gridStyle}
    >
      {layoutConfig.slots.map((slot) => (
        <div
          key={slot.id}
          className={styles.gridItem}
          style={{ gridArea: slot.gridArea }}
        >
          <CardSlot
            slotId={slot.id}
            story={slotStories[slot.id] || undefined}
            variant={slot.variant}
            showImage={slot.showImage !== false}
            onRemove={() => onRemoveStory(slot.id)}          />
        </div>
      ))}
    </div>
  );
};

export default LayoutRenderer;
