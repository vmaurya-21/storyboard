import * as React from 'react';
import { render, screen } from '@testing-library/react';
import { DndContext } from '@dnd-kit/core';
import CardSlot from '../../cards/CardSlot';
import { IStory } from '../../types';

const DndWrapper: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <DndContext onDragEnd={() => {}}>{children}</DndContext>
);

describe('CardSlot Component', () => {
  const mockStory: IStory = {
    id: 'story-1',
    title: 'Test Story',
    description: 'Test Description',
    imageUrl: 'https://via.placeholder.com/300x200',
    linkToPost: 'https://example.com',
    date: '2024-01-15'
  };

  it('should render empty slot when no story provided', () => {
    render(
      <DndWrapper>
        <CardSlot slotId="slot-1" story={undefined} variant="large" />
      </DndWrapper>
    );
    
    expect(document.querySelector('[data-icon-name="Add"]')).toBeInTheDocument();
  });

  it('should render story when story is provided', () => {
    render(
      <DndWrapper>
        <CardSlot slotId="slot-1" story={mockStory} variant="large" />
      </DndWrapper>
    );
    
    expect(screen.getByText('Test Story')).toBeInTheDocument();
    expect(document.querySelector('[data-icon-name="Add"]')).not.toBeInTheDocument();
  });

  it('should pass variant to StoryCard', () => {
    const { container } = render(
      <DndWrapper>
        <CardSlot slotId="slot-1" story={mockStory} variant="small" />
      </DndWrapper>
    );
    
    expect(container.querySelector('[class*="small"]')).toBeInTheDocument();
  });

  it('should pass showImage prop to StoryCard', () => {
    render(
      <DndWrapper>
        <CardSlot slotId="slot-1" story={mockStory} variant="large" showImage={false} />
      </DndWrapper>
    );
    
    expect(screen.queryByAltText('Test Story')).not.toBeInTheDocument();
  });

  it('should show image by default', () => {
    render(
      <DndWrapper>
        <CardSlot slotId="slot-1" story={mockStory} variant="large" />
      </DndWrapper>
    );
    
    expect(screen.getByAltText('Test Story')).toBeInTheDocument();
  });
});
