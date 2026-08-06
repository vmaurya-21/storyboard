import * as React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import { DndContext } from '@dnd-kit/core';
import LayoutRenderer from '../../layouts/LayoutRenderer';
import { SlotStoryMap } from '../../types';
import { mockStories } from '../mockData';
import * as layoutConfigModule from '../../layouts/layoutConfig';

const storyCarouselMock = jest.fn((_props: any) => <div data-testid="story-carousel" />);
jest.mock('../../StoryCarousel', () => ({
  __esModule: true,
  default: (props: any) => storyCarouselMock(props),
}));

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

const DndWrapper: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <DndContext onDragEnd={() => {}}>{children}</DndContext>
);

describe('LayoutRenderer Component', () => {
  const mockOnRemoveStory = jest.fn();

  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('Storyboard Layout', () => {
    it('should render 5 card slots for storyboard layout', () => {
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
      
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(5);
    });

    it('should render stories in correct slots', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': mockStories[0],
        'slot-3': mockStories[1]
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
      
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(3);
    });

    it('should call onRemoveStory for rd-slot-1', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': mockStories[0]
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
      
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-1');
    });

    it('should call onRemoveStory for all storyboard slots', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': mockStories[0],
        'slot-2': mockStories[1],
        'slot-3': mockStories[2],
        'slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' }
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
      
      fireEvent.click(removeButtons[0]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-1');
      
      fireEvent.click(removeButtons[1]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-2');
      
      fireEvent.click(removeButtons[2]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-3');
      
      fireEvent.click(removeButtons[3]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-4');
      
      fireEvent.click(removeButtons[4]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-5');
    });

    it('should render all storyboard slots with stories', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': mockStories[0],
        'slot-2': mockStories[1],
        'slot-3': mockStories[2],
        'slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' }
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

  describe('Connect Homepage Layout', () => {
    it('delegates rendering to StoryCarousel with slot ids and handlers', () => {
      const slotStories: SlotStoryMap = {
        'car-slot-1': mockStories[0]
      };

      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="connectHomepage"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
            onSlotLayoutChange={jest.fn()}
          />
        </DndWrapper>
      );

      expect(screen.getByTestId('story-carousel')).toBeInTheDocument();
      expect(storyCarouselMock).toHaveBeenCalled();
      expect(storyCarouselMock.mock.calls.length).toBeGreaterThan(0);
      const firstCallProps = storyCarouselMock.mock.calls[0][0] as any;
      expect(Array.isArray(firstCallProps.slotIds)).toBe(true);
      expect(firstCallProps.onRemoveStory).toBeDefined();
      expect(firstCallProps.isAdminMode).toBe(true);
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

  describe('Highlight Layout', () => {
    it('should render 5 card slots for highlight layout', () => {
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
      
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(5);
    });

    it('calls onRemoveStory for all highlight slots', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': mockStories[0],
        'slot-2': mockStories[1],
        'slot-3': mockStories[2],
        'slot-4': { ...mockStories[0], id: 'story-4', title: 'Story 4' },
        'slot-5': { ...mockStories[1], id: 'story-5', title: 'Story 5' }
      };

      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="highlight"
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );

      const removeButtons = screen.getAllByTitle('Remove story');
      expect(removeButtons).toHaveLength(5);

      fireEvent.click(removeButtons[0]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-1');

      fireEvent.click(removeButtons[1]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-4');

      fireEvent.click(removeButtons[2]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-2');

      fireEvent.click(removeButtons[3]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-5');

      fireEvent.click(removeButtons[4]);
      expect(mockOnRemoveStory).toHaveBeenCalledWith('slot-3');
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

    it('should switch from general to storyboard with stories', () => {
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
        'slot-1': mockStories[1]
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

  describe('Fallback Empty Layout Branch', () => {
    it('shows empty layout copy when layout config has no slots', () => {
      const spy = jest.spyOn(layoutConfigModule, 'getLayoutConfig').mockReturnValueOnce({
        id: 'general',
        name: 'General',
        description: 'General layout',
        slots: [],
        gridTemplate: {
          columns: '1fr',
          rows: '1fr',
          areas: '"a"'
        }
      });

      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType="general"
            slotStories={{}}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );

      expect(screen.getByText('This layout is not yet configured.')).toBeInTheDocument();
      expect(screen.getByText('Please select a different layout.')).toBeInTheDocument();
      spy.mockRestore();
    });
  });

  describe('Edge Cases', () => {
    it('should handle undefined story in slot', () => {
      const slotStories: SlotStoryMap = {
        'slot-1': undefined as any
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
      
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(5);
    });

    it('should handle mixed empty and filled slots', () => {
      const slotStories: SlotStoryMap = {
        'slot-2': mockStories[0],
        'slot-4': mockStories[1]
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
      
      const plusSigns = screen.getAllByText('+');
      expect(plusSigns).toHaveLength(3);
    });

    it('renders and removes slots through generic fallback branch', () => {
      const spy = jest.spyOn(layoutConfigModule, 'getLayoutConfig').mockReturnValueOnce({
        id: 'custom' as any,
        name: 'Custom',
        description: 'Custom fallback layout',
        slots: [
          { id: 'custom-slot-1', gridArea: 'a', variant: 'small' as any, showImage: true },
          { id: 'custom-slot-2', gridArea: 'b', variant: 'small' as any, showImage: false },
        ],
        gridTemplate: {
          columns: '1fr 1fr',
          rows: 'auto',
          areas: '"a b"'
        }
      });

      const slotStories: SlotStoryMap = {
        'custom-slot-1': mockStories[0],
      };

      render(
        <DndWrapper>
          <LayoutRenderer
            layoutType={'custom' as any}
            slotStories={slotStories}
            onRemoveStory={mockOnRemoveStory}
          />
        </DndWrapper>
      );

      expect(screen.getByText(mockStories[0].title)).toBeInTheDocument();
      expect(screen.getByText('+')).toBeInTheDocument();

      fireEvent.click(screen.getByTitle('Remove story'));
      expect(mockOnRemoveStory).toHaveBeenCalledWith('custom-slot-1');

      spy.mockRestore();
    });
  });
});
