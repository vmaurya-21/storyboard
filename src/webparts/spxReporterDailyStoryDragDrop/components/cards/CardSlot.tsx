import * as React from 'react';
import { useDroppable } from '@dnd-kit/core';
import { useSortable } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory, CardVariant, CardLayout } from '../types';
import StoryCard from './StoryCard';
import styles from './CardSlot.module.scss';

/** Props for {@link CardSlot}. */
export interface ICardSlotProps {
  /** Slot id from the layout preset; also the dnd-kit droppable id. */
  slotId: string;
  /** Story occupying the slot, or `undefined` to render the empty state. */
  story: IStory | undefined;
  /** Size family, applied as a modifier class on the slot. */
  variant: CardVariant;
  /** Content layout for the card. Falls back to the card's own default. */
  cardLayout?: CardLayout;
  /** Whether the card may render its image. Defaults to `true`. */
  showImage?: boolean;
  /** Invoked when the editor clears the slot. Omit to hide the remove control. */
  onRemove?: () => void;
}

/**
 * A placed story wrapped in a dnd-kit sortable handle.
 *
 * The whole card is the drag handle, so listeners are spread onto the wrapper
 * rather than a grip element. Only the translation is applied — scale and
 * rotation would distort the card against the fixed slot geometry.
 */
const DraggableStoryCard: React.FC<{
  story: IStory;
  variant: CardVariant;
  cardLayout?: CardLayout;
  showImage: boolean;
  onRemove?: () => void;
}> = ({ story, variant, cardLayout, showImage, onRemove }) => {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    isDragging,
  } = useSortable({ id: story.id });

  const style = {
    transform: CSS.Translate.toString(transform),
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      {...attributes}
      {...listeners}
      className={styles.cardWrapper}
    >
      <StoryCard
        story={story}
        variant={variant}
        cardLayout={cardLayout}
        showImage={showImage}
        onRemove={onRemove}
        isDragging={isDragging}
      />
    </div>
  );
};

/**
 * One position on the board: a drop target that holds at most one story.
 *
 * Empty slots show a plus affordance; filled slots render a
 * `DraggableStoryCard` so the story can be moved to another slot. The
 * slot highlights while a dragged story hovers over it.
 */
const CardSlot: React.FC<ICardSlotProps> = ({
  slotId,
  story,
  variant,
  cardLayout,
  showImage = true,
  onRemove
}) => {
  const { setNodeRef, isOver } = useDroppable({
    id: slotId,
  });

  return (
    <div
      ref={setNodeRef}
      className={`${styles.cardSlot} ${styles[variant]} ${isOver ? styles.draggingOver : ''}`}
    >
      {!story ? (
        <div className={styles.emptySlot}>
          <Icon iconName="Add" className={styles.plusIcon} />
        </div>
      ) : (
        <DraggableStoryCard
          story={story}
          variant={variant}
          cardLayout={cardLayout}
          showImage={showImage}
          onRemove={onRemove}
        />
      )}
    </div>
  );
};

export default CardSlot;
