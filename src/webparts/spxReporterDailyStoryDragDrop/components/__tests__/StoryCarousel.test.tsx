import * as React from 'react';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import StoryCarousel from '../StoryCarousel';
import { mockStories } from './mockData';

jest.mock('@dnd-kit/core', () => ({
  useDroppable: () => ({
    setNodeRef: jest.fn(),
    isOver: false,
  }),
}));

describe('StoryCarousel', () => {
  const slotIds = ['slot-1', 'slot-2'];
  const slotStories = {
    'slot-1': mockStories[0],
    'slot-2': mockStories[1],
  };

  const dispatchTransitionEnd = (element: HTMLElement, propertyName: string): void => {
    const event = new Event('transitionend', { bubbles: true });
    Object.defineProperty(event, 'propertyName', { value: propertyName });
    element.dispatchEvent(event);
  };

  it('renders nothing when there are no slots', () => {
    const { container } = render(
      <StoryCarousel
        slotIds={[]}
        slotStories={{}}
      />
    );

    expect(container.firstChild).toBeNull();
  });

  it('renders story cards and allows removing a story in admin mode', () => {
    const onRemoveStory = jest.fn();

    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        onRemoveStory={onRemoveStory}
        isAdminMode
      />
    );

    expect(screen.getAllByText('Test Story 1').length).toBeGreaterThan(0);

    const removeButtons = screen.getAllByTitle('Remove story');
    fireEvent.click(removeButtons[0]);

    expect(onRemoveStory).toHaveBeenCalledWith('slot-1');
  });

  it('changes slide layout through the controls menu', () => {
    const onSlotLayoutChange = jest.fn();

    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        onSlotLayoutChange={onSlotLayoutChange}
        isAdminMode
      />
    );

    const menuButtons = screen.getAllByTitle('Change slide layout');
    fireEvent.click(menuButtons[0]);

    fireEvent.click(screen.getByText('Thumbnail with Text'));
    expect(onSlotLayoutChange).toHaveBeenCalledWith('slot-1', 'thumbnail-text');
  });

  it('closes the layout menu when clicking outside controls', async () => {
    const onSlotLayoutChange = jest.fn();

    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        onSlotLayoutChange={onSlotLayoutChange}
        isAdminMode
      />
    );

    fireEvent.click(screen.getAllByTitle('Change slide layout')[0]);
    expect(screen.getByRole('menu')).toBeInTheDocument();

    fireEvent.mouseDown(document.body);

    await waitFor(() => {
      expect(screen.queryByRole('menu')).not.toBeInTheDocument();
    });
  });

  it('hides layout and remove controls when optional handlers are missing', () => {
    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        isAdminMode
      />
    );

    expect(screen.queryByTitle('Change slide layout')).not.toBeInTheDocument();
    expect(screen.queryByTitle('Remove story')).not.toBeInTheDocument();
  });

  it('shows read-only empty copy when not in admin mode', () => {
    render(
      <StoryCarousel
        slotIds={['slot-1']}
        slotStories={{ 'slot-1': undefined }}
        isAdminMode={false}
      />
    );

    expect(screen.getAllByText('No story').length).toBeGreaterThan(0);
    expect(screen.queryByTitle('Remove story')).not.toBeInTheDocument();
  });

  it('moves forward when next is clicked', async () => {
    const { container } = render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
      />
    );

    const next = screen.getByLabelText('Next slide');
    fireEvent.click(next);

    const track = container.querySelector('.track') as HTMLDivElement;

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-200%)');
    });
  });

  it('moves backward when previous is clicked', async () => {
    const { container } = render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
      />
    );

    const prev = screen.getByLabelText('Previous slide');
    fireEvent.click(prev);

    const track = container.querySelector('.track') as HTMLDivElement;

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-0%)');
    });
  });

  it('wraps forward from trailing clone back to the first real slide', async () => {
    const { container } = render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
      />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    fireEvent.click(screen.getByLabelText('Next slide'));
    fireEvent.click(screen.getByLabelText('Next slide'));

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-300%)');
    });

    dispatchTransitionEnd(track, 'transform');

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-100%)');
    });
  });

  it('wraps backward from leading clone to the last real slide', async () => {
    const { container } = render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
      />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    fireEvent.click(screen.getByLabelText('Previous slide'));

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-0%)');
    });

    dispatchTransitionEnd(track, 'transform');

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-200%)');
    });
  });

  it('ignores non-transform transition-end events', async () => {
    const { container } = render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
      />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    fireEvent.click(screen.getByLabelText('Next slide'));
    fireEvent.click(screen.getByLabelText('Next slide'));

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-300%)');
    });

    dispatchTransitionEnd(track, 'opacity');

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-300%)');
    });
  });
});
