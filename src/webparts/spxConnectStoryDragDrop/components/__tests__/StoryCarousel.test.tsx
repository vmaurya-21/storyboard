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

  // Regression: clicks used to increment without bound, so a third one pushed
  // the track past the trailing clone to -400%, where neither wrap branch
  // matches and the carousel showed empty space for good.
  it('clamps at the trailing clone when clicked past the end', async () => {
    const { container } = render(
      <StoryCarousel slotIds={slotIds} slotStories={slotStories} />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    const next = screen.getByLabelText('Next slide');
    fireEvent.click(next);
    fireEvent.click(next);
    fireEvent.click(next);
    fireEvent.click(next);

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-300%)');
    });

    // Still recoverable: the pending wrap lands on the first real slide.
    dispatchTransitionEnd(track, 'transform');

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-100%)');
    });
  });

  it('clamps at the leading clone when clicked past the start', async () => {
    const { container } = render(
      <StoryCarousel slotIds={slotIds} slotStories={slotStories} />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    const prev = screen.getByLabelText('Previous slide');
    fireEvent.click(prev);
    fireEvent.click(prev);

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-0%)');
    });
  });

  it('advances on a left swipe and goes back on a right swipe', async () => {
    const { container } = render(
      <StoryCarousel slotIds={slotIds} slotStories={slotStories} />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    const viewport = container.querySelector('.viewport') as HTMLDivElement;

    fireEvent.touchStart(viewport, { touches: [{ clientX: 200 }] });
    fireEvent.touchEnd(viewport, { changedTouches: [{ clientX: 100 }] });

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-200%)');
    });

    fireEvent.touchStart(viewport, { touches: [{ clientX: 100 }] });
    fireEvent.touchEnd(viewport, { changedTouches: [{ clientX: 200 }] });

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-100%)');
    });
  });

  it('ignores a touch that travels less than the swipe threshold', async () => {
    const { container } = render(
      <StoryCarousel slotIds={slotIds} slotStories={slotStories} />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    const viewport = container.querySelector('.viewport') as HTMLDivElement;

    fireEvent.touchStart(viewport, { touches: [{ clientX: 200 }] });
    fireEvent.touchEnd(viewport, { changedTouches: [{ clientX: 180 }] });

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-100%)');
    });
  });

  it('navigates with the arrow keys', async () => {
    const { container } = render(
      <StoryCarousel slotIds={slotIds} slotStories={slotStories} />
    );

    const track = container.querySelector('.track') as HTMLDivElement;
    const region = screen.getByRole('region');

    fireEvent.keyDown(region, { key: 'ArrowRight' });

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-200%)');
    });

    fireEvent.keyDown(region, { key: 'ArrowLeft' });

    await waitFor(() => {
      expect(track.style.transform).toBe('translateX(-100%)');
    });
  });

  it('exposes carousel and slide semantics', () => {
    render(<StoryCarousel slotIds={slotIds} slotStories={slotStories} />);

    expect(screen.getByRole('region', { name: 'Story carousel' })).toBeInTheDocument();
    // Two real slides; the loop clones are aria-hidden.
    expect(screen.getAllByRole('group')).toHaveLength(2);
    expect(screen.getByRole('group', { name: '1 of 2' })).toBeInTheDocument();
    expect(screen.getByRole('group', { name: '2 of 2' })).toBeInTheDocument();
  });

  it('renders the byline when the story has an author, and the date alone when not', () => {
    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={{
          'slot-1': { ...mockStories[0], author: 'Ada Lovelace' },
          'slot-2': mockStories[1],
        }}
      />
    );

    expect(screen.getAllByText('Ada Lovelace').length).toBeGreaterThan(0);
    expect(screen.getAllByText('•').length).toBeGreaterThan(0);
  });

  it('falls back to full-image for a layout the carousel cannot render', () => {
    render(
      <StoryCarousel
        slotIds={slotIds}
        slotStories={slotStories}
        slotLayoutPreferences={{ 'slot-1': 'text-only' }}
      />
    );

    // The split layout is the only one with a "Read more" link, so its absence
    // means the slide fell back to full-image rather than rendering nothing.
    expect(screen.queryByText('Read more →')).not.toBeInTheDocument();
  });
});
