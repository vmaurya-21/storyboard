import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { LayoutType, SlotStoryMap } from '../types';
import { getLayoutConfig } from './layoutConfig';
import CardSlot from '../cards/CardSlot';
import StoryCarousel, { CarouselDisplayMode } from '../StoryCarousel';
import styles from './LayoutRenderer.module.scss';

export interface ILayoutRendererProps {
  layoutType: LayoutType;
  slotStories: SlotStoryMap;
  onRemoveStory: (slotId: string) => void;
  slotLayoutPreferences?: { [key: string]: CarouselDisplayMode };
  onSlotLayoutChange?: (slotId: string, mode: CarouselDisplayMode) => void;
  featuredStoryId?: string;
}

const LayoutRenderer: React.FC<ILayoutRendererProps> = ({
  layoutType,
  slotStories,
  onRemoveStory,
  slotLayoutPreferences = {},
  onSlotLayoutChange,
  featuredStoryId
}) => {
  const layoutConfig = getLayoutConfig(layoutType);

  const isFeatured = (slotId: string): boolean => {
    const story = slotStories[slotId];
    return !!story && !!featuredStoryId && story.id === featuredStoryId;
  };

  // Connect Homepage — full-width carousel.
  if (layoutType === 'connectHomepage') {
    const slotIds = layoutConfig.slots.map(s => s.id);
    return (
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        slotLayoutPreferences={slotLayoutPreferences}
        onSlotLayoutChange={onSlotLayoutChange}
        onRemoveStory={onRemoveStory}
        featuredStoryId={featuredStoryId}
        isAdminMode
      />
    );
  }

  // Don't render if layout has no slots (like an unconfigured placeholder)
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

  // Reporter Daily — three flex columns matching the reference board.
  if (layoutType === 'reporterDaily') {
    return (
      <div className={`${styles.layoutGrid} ${styles.reporterDaily}`}>
        {/* Column 1: two-medium (full-image + thumbnail-text) */}
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
              onRemove={() => onRemoveStory('slot-1')}
              featured={isFeatured('slot-1')}
            />
            <CardSlot
              slotId="slot-2"
              story={slotStories['slot-2'] || undefined}
              variant="medium"
              cardLayout="thumbnail-text"
              showImage={true}
              onRemove={() => onRemoveStory('slot-2')}
              featured={isFeatured('slot-2')}
            />
          </div>
        </div>

        {/* Column 2: small-tall (text-only + full-image) */}
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
              onRemove={() => onRemoveStory('slot-3')}
              featured={isFeatured('slot-3')}
            />
            <CardSlot
              slotId="slot-4"
              story={slotStories['slot-4'] || undefined}
              variant="tall"
              cardLayout="full-image"
              showImage={true}
              onRemove={() => onRemoveStory('slot-4')}
              featured={isFeatured('slot-4')}
            />
          </div>
        </div>

        {/* Column 3: full (fills the column height) */}
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
              onRemove={() => onRemoveStory('slot-5')}
              featured={isFeatured('slot-5')}
            />
          </div>
        </div>
      </div>
    );
  }

  // General Layout - 3 columns of "Two Half Slots" each (total 6 slots)
  if (layoutType === 'general') {
    return (
      <div className={`${styles.layoutGrid} ${styles.general}`}>
        {/* Column 1: Two Half Slots (gen-slot-1 and gen-slot-4) */}
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
              onRemove={() => onRemoveStory('gen-slot-1')}
              featured={isFeatured('gen-slot-1')}
            />
            <CardSlot
              slotId="gen-slot-4"
              story={slotStories['gen-slot-4'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-4'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-4')}
              featured={isFeatured('gen-slot-4')}
            />
          </div>
        </div>

        {/* Column 2: Two Half Slots (gen-slot-2 and gen-slot-5) */}
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
              onRemove={() => onRemoveStory('gen-slot-2')}
              featured={isFeatured('gen-slot-2')}
            />
            <CardSlot
              slotId="gen-slot-5"
              story={slotStories['gen-slot-5'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-5'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-5')}
              featured={isFeatured('gen-slot-5')}
            />
          </div>
        </div>

        {/* Column 3: Two Half Slots (gen-slot-3 and gen-slot-6) */}
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
              onRemove={() => onRemoveStory('gen-slot-3')}
              featured={isFeatured('gen-slot-3')}
            />
            <CardSlot
              slotId="gen-slot-6"
              story={slotStories['gen-slot-6'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['gen-slot-6'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('gen-slot-6')}
              featured={isFeatured('gen-slot-6')}
            />
          </div>
        </div>
      </div>
    );
  }

  // Highlight Layout - Column 1: Two Half Slots, Column 2: Two Half Slots, Column 3: Full Slot
  if (layoutType === 'highlight') {
    return (
      <div className={`${styles.layoutGrid} ${styles.highlight}`}>
        {/* Column 1: Two Half Slots (slot-1 and slot-4) */}
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
              onRemove={() => onRemoveStory('slot-1')}
              featured={isFeatured('slot-1')}
            />
            <CardSlot
              slotId="slot-4"
              story={slotStories['slot-4'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-4'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-4')}
              featured={isFeatured('slot-4')}
            />
          </div>
        </div>

        {/* Column 2: Two Half Slots (slot-2 and slot-5) */}
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
              onRemove={() => onRemoveStory('slot-2')}
              featured={isFeatured('slot-2')}
            />
            <CardSlot
              slotId="slot-5"
              story={slotStories['slot-5'] || undefined}
              variant="medium"
              cardLayout={slotLayoutPreferences['slot-5'] || 'full-image'}
              showImage={true}
              onRemove={() => onRemoveStory('slot-5')}
              featured={isFeatured('slot-5')}
            />
          </div>
        </div>

        {/* Column 3: Full Slot (slot-3) */}
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
              onRemove={() => onRemoveStory('slot-3')}
              featured={isFeatured('slot-3')}
            />
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
            onRemove={() => onRemoveStory(slot.id)}
            featured={isFeatured(slot.id)}
          />
        </div>
      ))}
    </div>
  );
};

export default LayoutRenderer;
