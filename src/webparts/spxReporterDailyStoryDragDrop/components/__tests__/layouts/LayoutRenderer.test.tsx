import * as React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import { DndContext } from '@dnd-kit/core';
import LayoutRenderer from '../../layouts/LayoutRenderer';
import { SlotStoryMap } from '../../types';
import { mockStories } from '../mockData';

// Mock CardSlot to expose onRemove callbacks
jest.mock('../../cards/CardSlot', () => {
  return function MockCardSlot({ slotId, story, onRemove }: any) {
    return (
      <div data-testid={`card-slot-${slotId}`}>
        {story ? (
          <div>
            <span>{story.title}</span>
            <button onClick={onRemove} title="Remove story">Remove</button>
          </div>
        ) : (
          <span>+</span>
        )}
      </div>
    );
  };
});

// Wrapper component to provide DndContext
const DndWrapper: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <DndContext onDragEnd={() => {}}>{children}</DndContext>
);

describe('LayoutRenderer Component', () => {
  const mockOnRemoveStory = jest.fn();

  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('Reporter Daily Layout', () => {
    it('should render 5 card slots for reporter daily layout', () => {
      const slotStories: SlotStoryMap = {};
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      // Should render 5 empty slots (+ symbols)
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(5);
    });

    it('should render stories in correct slots', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-1': mockStories[0],
        'rd-slot-3': mockStories[1]
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
      
      // Should have 3 empty slots
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(3);
    });

    it('should call onRemoveStory for rd-slot-1', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-1': mockStories[0]
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      const removeButton = screen.getByTitle('Remove story');
      fireEvent.click(removeButton);
      
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-1');
    });

    it('should call onRemoveStory for all reporter daily slots', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-1': mockStories[0],
        'rd-slot-2': mockStories[1],
        'rd-slot-3': mockStories[2],
        'rd-slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'rd-slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' }
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      const removeButtons = screen.getAllByTitle('Remove story');
      expect(removeButtons).toHaveLength(5);
      
      // Test each slot's remove button
      fireEvent.click(removeButtons[0]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-1');
      
      fireEvent.click(removeButtons[1]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-2');
      
      fireEvent.click(removeButtons[2]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-3');
      
      fireEvent.click(removeButtons[3]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-4');
      
      fireEvent.click(removeButtons[4]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('rd-slot-5');
    });

    it('should render all reporter daily slots with stories', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-1': mockStories[0],
        'rd-slot-2': mockStories[1],
        'rd-slot-3': mockStories[2],
        'rd-slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'rd-slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' }
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[2].title)).toBeInTheDocument();
      expect(screen.getByText('Story 4')).toBeInTheDocument();
      expect(screen.getByText('Story 5')).toBeInTheDocument();
    });
  });

  describe('General Layout', () => {
    it('should render 6 card slots for general layout', () => {
      const slotStories: SlotStoryMap = {};
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      // Should render 6 empty slots
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(6);
    });

    it('should render stories in general layout', () => {
      const slotStories: SlotStoryMap = {
        'gen-slot-1': mockStories[0],
        'gen-slot-2': mockStories[1],
        'gen-slot-3': mockStories[2]
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[2].title)).toBeInTheDocument();
    });

    it('should call onRemoveStory for general layout slots', () => {
      const slotStories: SlotStoryMap = {
        'gen-slot-1': mockStories[0],
        'gen-slot-2': mockStories[1]
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      const removeButtons = screen.getAllByTitle('Remove story');
      
      fireEvent.click(removeButtons[0]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('gen-slot-1');
      
      fireEvent.click(removeButtons[1]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('gen-slot-2');
    });

    it('should render all 6 general layout slots with stories', () => {
      const slotStories: SlotStoryMap = {
        'gen-slot-1': mockStories[0],
        'gen-slot-2': mockStories[1],
        'gen-slot-3': mockStories[2],
        'gen-slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'gen-slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' },
        'gen-slot-6': { ...mockStories[2], id: 'story-6', title: 'Story 6' }
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[2].title)).toBeInTheDocument();
      expect(screen.getByText('Story 4')).toBeInTheDocument();
      expect(screen.getByText('Story 5')).toBeInTheDocument();
      expect(screen.getByText('Story 6')).toBeInTheDocument();
    });
  });

  describe('Highlight Layout (Empty)', () => {
    it('should show empty state for highlight layout', () => {
      const slotStories: SlotStoryMap = {};
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="highlight"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText('This layout is not yet configured.')).toBeInTheDocument();
      expect(screen.getByText('Please select a different layout.')).toBeInTheDocument();
    });
  });

  describe('Layout Switching', () => {
    it('should render different number of slots when layout changes', () => {
      const slotStories: SlotStoryMap = {};
      
      const { rerender } = render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getAllByText('+')).toHaveLength(5);
      
      rerender(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getAllByText('+')).toHaveLength(6);
    });

    it('should switch from general to reporter daily with stories', () => {
      const slotStories: SlotStoryMap = {
        'gen-slot-1': mockStories[0]
      };
      
      const { rerender } = render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      
      const rdSlotStories: SlotStoryMap = {
        'rd-slot-1': mockStories[1]
      };
      
      rerender(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={rdSlotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
    });
  });

  describe('Edge Cases', () => {
    it('should handle undefined story in slot', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-1': undefined as any
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      // Should render empty slots
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(5);
    });

    it('should handle mixed empty and filled slots', () => {
      const slotStories: SlotStoryMap = {
        'rd-slot-2': mockStories[0],
        'rd-slot-4': mockStories[1]
      };
      
      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="reporterDaily"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );
      
      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText(mockStories[1].title)).toBeInTheDocument();
      
      // Should have 3 empty slots
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(3);
    });
  });
});
