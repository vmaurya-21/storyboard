import * as React from 'react';
import { useDroppable } from '@dnd-kit/core';
import { useSortable } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStory, CardVariant, CardLayout } from '../types';
import StoryCard from './StoryCard';
import styles from './CardSlot.module.scss';

export interface ICardSlotProps {
  slotId: string;
  story: IStory | undefined;
  variant: CardVariant;
  cardLayout?: CardLayout;
  showImage?: boolean;
  onRemove?: () => void;
  featured?: boolean;
}

// Separate component for draggable story cards in slots
const DraggableStoryCard: React.FC<{
  story: IStory;
  variant: CardVariant;
  cardLayout?: CardLayout;
  showImage: boolean;
  onRemove?: () => void;
  featured?: boolean;
}> = ({ story, variant, cardLayout, showImage, onRemove, featured }) => {
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
        featured={featured}
      />
    </div>
  );
};

const CardSlot: React.FC<ICardSlotProps> = ({
  slotId,
  story,
  variant,
  cardLayout,
  showImage = true,
  onRemove,
  featured
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
          featured={featured}
        />
      )}
    </div>
  );
};

export default CardSlot;
