import * as React from 'react';
import { render, screen, fireEvent, waitFor, within, act } from '@testing-library/react';
import StoryDragDrop from '../StoryDragDrop';
import { mockStories, mockContext } from './mockData';
import * as availableStoriesService from '../../services/availableStoriesService';
import * as scheduleStoriesService from '../../services/scheduleStoriesService';
import * as publishStoriesService from '../../services/publishStoriesService';

jest.mock('../../services/availableStoriesService');
jest.mock('../../services/scheduleStoriesService', () => {
  const actual = jest.requireActual('../../services/scheduleStoriesService');
  return {
    ...actual,
    createScheduledGroup: jest.fn(),
    getScheduledGroups: jest.fn(),
    deleteScheduledGroup: jest.fn(),
    deleteScheduledGroupByGroupId: jest.fn(),
    rescheduleGroup: jest.fn(),
  };
});
jest.mock('../../services/publishStoriesService', () => {
  const actual = jest.requireActual('../../services/publishStoriesService');
  return {
    ...actual,
    publishBoard: jest.fn(),
    getLiveBoard: jest.fn(),
  };
});
jest.mock('../AddStoryModal', () => ({ onClose, onAdd }: any) => (
  <div data-testid="add-story-modal">
    <button onClick={() => onAdd({ title: 'New Story', imageUrl: 'https://example.com/new.jpg', linkToPost: 'https://example.com/new' })}>Add</button>
    <button onClick={onClose}>Close</button>
  </div>
));
jest.mock('../EditStoryModal', () => ({ story, onClose, onUpdate, onDelete }: any) => (
  <div data-testid="edit-story-modal">
    <button onClick={() => onUpdate({ ...story, title: 'Updated' })}>Update</button>
    <button onClick={() => onDelete(story.id)}>Delete</button>
    <button onClick={onClose}>Close</button>
  </div>
));
// `data-slot1-layout` surfaces the shared `slotLayoutPreferences` map so tests
// can assert the live board's copy survives a trip through the editing modes.
jest.mock('../layouts/LayoutRenderer', () => ({ onRemoveStory, slotLayoutPreferences }: any) => (
  <div
    data-testid="layout-renderer"
    data-slot1-layout={(slotLayoutPreferences || {})['slot-1'] || ''}
  >
    <button onClick={() => onRemoveStory('slot-1')}>Remove</button>
  </div>
));
jest.mock('../DeleteGroupModal', () => ({ onClose, onConfirm }: any) => (
  <div data-testid="delete-group-modal">
    <button onClick={onConfirm}>Confirm Delete Group</button>
    <button onClick={onClose}>Cancel Delete Group</button>
  </div>
));

/** Latest `onDragEnd` handed to the DndContext, so tests can stage a drop. */
let mockOnDragEnd: ((event: any) => void) | undefined;

jest.mock('@dnd-kit/core', () => ({
  DndContext: ({ children, onDragEnd }: any) => {
    mockOnDragEnd = onDragEnd;
    return (
      <div data-testid="dnd-context" data-ondragend={onDragEnd ? 'true' : 'false'}>
        {children}
      </div>
    );
  },
  useSensor: jest.fn(),
  useSensors: jest.fn(() => []),
  PointerSensor: jest.fn(),
  pointerWithin: jest.fn(),
}));

jest.mock('@dnd-kit/sortable', () => ({
  SortableContext: ({ children }: any) => <div data-testid="sortable-context">{children}</div>,
  verticalListSortingStrategy: {},
  useSortable: () => ({
    attributes: {},
    listeners: {},
    setNodeRef: jest.fn(),
    transform: null,
    transition: null,
    isDragging: false,
  }),
}));

globalThis.confirm = jest.fn(() => true) as any;
globalThis.alert = jest.fn() as any;

describe('StoryDragDrop Component', () => {
  const buildScheduledGroup = (overrides: Record<string, any> = {}) => {
    const now = new Date();
    const future = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2, 9, 0, 0, 0);

    return {
      id: 'scheduled-1',
      date: future,
      time: '09:00',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {},
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      createdAt: future,
      ...overrides,
    };
  };

  beforeEach(() => {
    jest.clearAllMocks();
    (availableStoriesService.initializeSharePoint as jest.Mock).mockImplementation(() => {});
    (availableStoriesService.getStories as jest.Mock).mockResolvedValue(mockStories.map(s => ({
        id: parseInt(s.id.split('-')[1]),
        title: s.title,
        description: s.description,
        imageUrl: s.imageUrl,
        linkToPost: s.linkToPost
    })));
    (availableStoriesService.createStory as jest.Mock).mockResolvedValue({ data: { ID: 3 } });
    (availableStoriesService.updateStory as jest.Mock).mockResolvedValue(undefined);
    (availableStoriesService.deleteStory as jest.Mock).mockResolvedValue(undefined);
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValue([]);
    (scheduleStoriesService.createScheduledGroup as jest.Mock).mockImplementation(async (group: any) => group);
    (scheduleStoriesService.deleteScheduledGroup as jest.Mock).mockResolvedValue(undefined);
    (scheduleStoriesService.deleteScheduledGroupByGroupId as jest.Mock).mockResolvedValue(undefined);
    (scheduleStoriesService.rescheduleGroup as jest.Mock).mockResolvedValue(undefined);
    (publishStoriesService.publishBoard as jest.Mock).mockResolvedValue(undefined);
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue(null);
  });

  it('renders loading state initially', () => {
    render(<StoryDragDrop context={mockContext as any} />);
    expect(screen.getByRole('status')).toBeInTheDocument();
  });

  it('renders stories after loading', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.getByText('Test Story 1')).toBeInTheDocument();
    expect(screen.getByText('Test Story 2')).toBeInTheDocument();
  });

  it('sorts available stories by modified date desc, then title asc on ties', async () => {
    (availableStoriesService.getStories as jest.Mock).mockResolvedValue([
      {
        id: 1,
        title: 'Beta Story',
        description: '',
        imageUrl: 'https://example.com/1.jpg',
        linkToPost: 'https://example.com/1',
        date: '2026-08-10T10:00:00.000Z',
      },
      {
        id: 2,
        title: 'Alpha Story',
        description: '',
        imageUrl: 'https://example.com/2.jpg',
        linkToPost: 'https://example.com/2',
        date: '2026-08-10T10:00:00.000Z',
      },
      {
        id: 3,
        title: 'Newest Story',
        description: '',
        imageUrl: 'https://example.com/3.jpg',
        linkToPost: 'https://example.com/3',
        date: '2026-08-11T10:00:00.000Z',
      },
    ]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const titles = within(screen.getByTestId('sortable-context'))
      .getAllByRole('heading', { level: 4 })
      .map((el) => el.textContent);

    expect(titles).toEqual(['Newest Story', 'Alpha Story', 'Beta Story']);
  });

  it('keeps available-story sort order after refresh', async () => {
    const sortedInput = [
      {
        id: 1,
        title: 'Zulu Story',
        description: '',
        imageUrl: 'https://example.com/1.jpg',
        linkToPost: 'https://example.com/1',
        date: '2026-08-09T10:00:00.000Z',
      },
      {
        id: 2,
        title: 'Alpha Story',
        description: '',
        imageUrl: 'https://example.com/2.jpg',
        linkToPost: 'https://example.com/2',
        date: '2026-08-11T10:00:00.000Z',
      },
      {
        id: 3,
        title: 'Beta Story',
        description: '',
        imageUrl: 'https://example.com/3.jpg',
        linkToPost: 'https://example.com/3',
        date: '2026-08-10T10:00:00.000Z',
      },
    ];
    (availableStoriesService.getStories as jest.Mock).mockResolvedValue(sortedInput);

    const first = render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    let titles = within(screen.getByTestId('sortable-context'))
      .getAllByRole('heading', { level: 4 })
      .map((el) => el.textContent);
    expect(titles).toEqual(['Alpha Story', 'Beta Story', 'Zulu Story']);

    first.unmount();

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    titles = within(screen.getByTestId('sortable-context'))
      .getAllByRole('heading', { level: 4 })
      .map((el) => el.textContent);
    expect(titles).toEqual(['Alpha Story', 'Beta Story', 'Zulu Story']);
  });

  it('hydrates board from published state on initial load', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {
        'slot-1': 'thumbnail-text',
      },
    });

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.queryByText('Test Story 1')).not.toBeInTheDocument();
    expect(screen.getByText('Test Story 2')).toBeInTheDocument();
    expect(screen.getByTestId('layout-renderer')).toHaveAttribute('data-slot1-layout', 'thumbnail-text');
  });

  it('disables Reset Board when there are no unsaved changes', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.getByRole('button', { name: /Reset Board/i })).toBeDisabled();
  });

  it('disables Clear Board when board is already empty', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.getByRole('button', { name: /Clear Board/i })).toBeDisabled();
  });

  it('shows error message if SharePoint context is missing', async () => {
    const consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    render(<StoryDragDrop context={undefined} />);
    
    await waitFor(() => {
      expect(screen.getByText(/SharePoint context not available/i)).toBeInTheDocument();
    });
    consoleSpy.mockRestore();
  });

  it('shows error when getStories fails', async () => {
    const consoleSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
    (availableStoriesService.getStories as jest.Mock).mockRejectedValue(new Error('Failed to load'));
    
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.getByText(/Failed to load stories from SharePoint/i)).toBeInTheDocument();
    });
    consoleSpy.mockRestore();
  });

  it('does not render a layout selector control', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.queryByRole('combobox')).not.toBeInTheDocument();
  });

  it('opens and closes AddStoryModal', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const addButton = screen.getByText(/Add story/i);
    fireEvent.click(addButton);
    
    expect(screen.getByTestId('add-story-modal')).toBeInTheDocument();

    const closeButton = screen.getByText('Close');
    fireEvent.click(closeButton);
    
    await waitFor(() => {
      expect(screen.queryByTestId('add-story-modal')).not.toBeInTheDocument();
    });
  });

  it('adds a new story', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const addButton = screen.getByText(/Add story/i);
    fireEvent.click(addButton);
    
    const addStoryButton = within(screen.getByTestId('add-story-modal')).getByText('Add');
    fireEvent.click(addStoryButton);

    await waitFor(() => {
      expect(availableStoriesService.createStory).toHaveBeenCalled();
      expect(availableStoriesService.getStories).toHaveBeenCalled();
    });
  });

  it('handles error when adding story fails', async () => {
    const consoleSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
    (availableStoriesService.createStory as jest.Mock).mockRejectedValue(new Error('Create failed'));
    
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const addButton = screen.getByText(/Add story/i);
    fireEvent.click(addButton);
    
    const addStoryButton = within(screen.getByTestId('add-story-modal')).getByText('Add');
    fireEvent.click(addStoryButton);

    await waitFor(() => {
      expect(availableStoriesService.createStory).toHaveBeenCalled();
    });
    
    consoleSpy.mockRestore();
  });

  it('opens EditStoryModal when edit button clicked', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const editButtons = screen.getAllByTitle('Edit story');
    fireEvent.click(editButtons[0]);
    
    expect(screen.getByTestId('edit-story-modal')).toBeInTheDocument();
  });

  it('updates a story', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const editButtons = screen.getAllByTitle('Edit story');
    fireEvent.click(editButtons[0]);
    
    const updateButton = screen.getByText('Update');
    fireEvent.click(updateButton);

    await waitFor(() => {
      expect(availableStoriesService.updateStory).toHaveBeenCalled();
    });
  });

  it('handles error when updating story with invalid ID', async () => {
    globalThis.alert = jest.fn() as any;
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const editButtons = screen.getAllByTitle('Edit story');
    fireEvent.click(editButtons[0]);
    
    const updateButton = screen.getByText('Update');
    
    fireEvent.click(updateButton);

    await waitFor(() => {
      expect(availableStoriesService.updateStory).toHaveBeenCalled();
    });
  });

  it('deletes a story', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const editButtons = screen.getAllByTitle('Edit story');
    fireEvent.click(editButtons[0]);
    
    const deleteButton = screen.getByText('Delete');
    fireEvent.click(deleteButton);

    await waitFor(() => {
      expect(availableStoriesService.deleteStory).toHaveBeenCalled();
    });
  });

  it('closes EditStoryModal when close button clicked', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const editButtons = screen.getAllByTitle('Edit story');
    fireEvent.click(editButtons[0]);
    
    expect(screen.getByTestId('edit-story-modal')).toBeInTheDocument();
    
    const closeButton = screen.getByText('Close');
    fireEvent.click(closeButton);

    await waitFor(() => {
      expect(screen.queryByTestId('edit-story-modal')).not.toBeInTheDocument();
    });
  });

  it('clears the board', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {
        'slot-1': 'thumbnail-text',
      },
    });

    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const clearButton = screen.getByText(/Clear Board/i);
    fireEvent.click(clearButton);
    fireEvent.click(screen.getByText('Yes, Clear Board'));
    
    expect(screen.getByTestId('layout-renderer')).toBeInTheDocument();
    await waitFor(() => {
      expect(screen.getByText('Board Cleared')).toBeInTheDocument();
    });
    expect(screen.getByTestId('layout-renderer')).toHaveAttribute('data-slot1-layout', '');
  });

  it('enters and exits scheduling mode', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const scheduleBtn = screen.getByText(/Schedule New Group/i);
    fireEvent.click(scheduleBtn);
    
    expect(screen.getByText(/Exit Scheduling/i)).toBeInTheDocument();
    expect(screen.getByText(/Select date & time/i)).toBeInTheDocument();

    const exitBtn = screen.getByText(/Exit Scheduling/i);
    fireEvent.click(exitBtn);
    
    expect(screen.getByText(/Schedule New Group/i)).toBeInTheDocument();
  });

  it('returns stories placed during scheduling mode to the available list on exit', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Schedule New Group/i));

    // The one SortableContext wraps the available list, so scope to it: the
    // scheduled-group cards render the same titles.
    const availableList = (): HTMLElement => screen.getByTestId('sortable-context');

    // Drop a story into a slot: it leaves the available list.
    act(() => {
      mockOnDragEnd!({
        active: {
          id: 'story-1',
          data: { current: { sortable: { containerId: 'available-stories' } } },
        },
        over: { id: 'slot-1' },
      });
    });
    expect(within(availableList()).queryByText('Test Story 1')).not.toBeInTheDocument();

    // Abandoning the schedule hands it back — nothing else owns it now.
    fireEvent.click(screen.getByText(/Exit Scheduling/i));

    expect(within(availableList()).getByText('Test Story 1')).toBeInTheDocument();
  });

  it('keeps scheduled stories out of the available list when scheduling succeeds', async () => {
    (scheduleStoriesService.createScheduledGroup as jest.Mock).mockImplementation(
      (group: any) => Promise.resolve(group)
    );

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Schedule New Group/i));

    act(() => {
      mockOnDragEnd!({
        active: {
          id: 'story-1',
          data: { current: { sortable: { containerId: 'available-stories' } } },
        },
        over: { id: 'slot-1' },
      });
    });

    fireEvent.click(screen.getByText(/Select date & time/i));
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    fireEvent.click(screen.getByText(String(tomorrow.getDate()), { selector: 'button' }));
    fireEvent.click(screen.getByText('Schedule', { selector: 'button' }));

    await waitFor(() => {
      expect(screen.getByText('Board Scheduled Successfully!')).toBeInTheDocument();
    });

    // The group owns it now, so it must not reappear in the available list —
    // the group's own card still shows the title, hence the scoped query.
    expect(
      within(screen.getByTestId('sortable-context')).queryByText('Test Story 1')
    ).not.toBeInTheDocument();
  });

  it('restores the live board layout preferences when group edits are cancelled', async () => {
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({
        id: 'scheduled-layout-prefs',
        slotLayoutPreferences: { 'slot-1': 'thumbnail-text' },
      })
    ]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const board = screen.getByTestId('layout-renderer');
    expect(board).toHaveAttribute('data-slot1-layout', '');

    fireEvent.click(screen.getAllByTitle('Edit on main storyboard')[0]);
    expect(screen.getByTestId('layout-renderer')).toHaveAttribute(
      'data-slot1-layout',
      'thumbnail-text'
    );

    fireEvent.click(screen.getByText(/Cancel Edits/i));
    expect(screen.getByTestId('layout-renderer')).toHaveAttribute('data-slot1-layout', '');
  });

  it('opens and closes calendar overlay', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const scheduleBtn = screen.getByText(/Schedule New Group/i);
    fireEvent.click(scheduleBtn);
    
    const dateTimeBtn = screen.getByText(/Select date & time/i);
    fireEvent.click(dateTimeBtn);
    
    expect(screen.getByLabelText(/Time/i)).toBeInTheDocument();

    const exitBtn = screen.getByText(/Exit Scheduling/i);
    fireEvent.click(exitBtn);
    
    await waitFor(() => {
      expect(screen.queryByLabelText(/Time/i)).not.toBeInTheDocument();
    });
  });

  it('shows validation when scheduling with a date but empty board', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const scheduleBtn = screen.getByText(/Schedule New Group/i);
    fireEvent.click(scheduleBtn);
    
    const dateTimeBtn = screen.getByText(/Select date & time/i);
    fireEvent.click(dateTimeBtn);
    
    const dayButtons = screen
      .getAllByRole('button')
      .filter(btn => btn.textContent && /^\d+$/.test(btn.textContent) && !btn.hasAttribute('disabled'));
    fireEvent.click(dayButtons[0]);

    const timeInput = screen.getByLabelText(/^Time$/i);
    fireEvent.change(timeInput, { target: { value: '14:30' } });

    const scheduleNowBtn = screen.getByText('Schedule');
    fireEvent.click(scheduleNowBtn);

    await waitFor(() => {
      expect(screen.getByText('No Stories to Schedule')).toBeInTheDocument();
    });
  });

  it('shows validation when scheduling without selecting a date', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const scheduleBtn = screen.getByText(/Schedule New Group/i);
    fireEvent.click(scheduleBtn);
    
    const dateTimeBtn = screen.getByText(/Select date & time/i);
    fireEvent.click(dateTimeBtn);

    const scheduleNowBtn = screen.getByText('Schedule');
    
    fireEvent.click(scheduleNowBtn);

    await waitFor(() => {
      expect(screen.getByText('No Date Selected')).toBeInTheDocument();
    });
  });

  it('deletes a scheduled group', async () => {
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup()
    ]);

    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    await waitFor(() => {
      const deleteButtons = screen.getAllByTitle('Delete this scheduled group');
      expect(deleteButtons.length).toBeGreaterThan(0);
      
      fireEvent.click(deleteButtons[0]);
    });

    fireEvent.click(screen.getByText('Confirm Delete Group'));

    await waitFor(() => {
      expect(scheduleStoriesService.deleteScheduledGroupByGroupId).toHaveBeenCalledWith('scheduled-1');
    });
  });

  it('deletes a scheduled group by its SharePoint item id when it has one', async () => {
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({ id: 'scheduled-with-spid', spId: 42 })
    ]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getAllByTitle('Delete this scheduled group')[0]);
    fireEvent.click(screen.getByText('Confirm Delete Group'));

    await waitFor(() => {
      expect(scheduleStoriesService.deleteScheduledGroup).toHaveBeenCalledWith(42);
    });
    // The GroupId filter would match nothing for a row whose column is empty
    // and still resolve, reporting a success that never happened.
    expect(scheduleStoriesService.deleteScheduledGroupByGroupId).not.toHaveBeenCalled();
  });

  it('displays stories in scheduled groups', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const scheduleBtn = screen.getByText(/Schedule New Group/i);
    fireEvent.click(scheduleBtn);
    
    const dateTimeBtn = screen.getByText(/Select date & time/i);
    fireEvent.click(dateTimeBtn);
    
    const dayButtons = screen.getAllByRole('button').filter(btn => btn.textContent && /^\d+$/.test(btn.textContent));
    if (dayButtons.length > 0) {
      fireEvent.click(dayButtons[15]);
    }

    const scheduleNowBtn = screen.getByText('Schedule');
    fireEvent.click(scheduleNowBtn);

    await waitFor(() => {
      const scheduledGroups = screen.queryAllByText(/Test Story/i);
      expect(scheduledGroups.length).toBeGreaterThan(0);
    });
  });

  it('patches scheduled group to layout', async () => {
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup()
    ]);

    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    await waitFor(() => {
      const patchButtons = screen.getAllByTitle('Edit on main storyboard');
      expect(patchButtons.length).toBeGreaterThan(0);
      
      fireEvent.click(patchButtons[0]);
    });

    expect(screen.getByText(/Save Group Edits/i)).toBeInTheDocument();
  });

  it('prompts layout switch when patching group with different layout', async () => {
    (globalThis.confirm as jest.Mock).mockReturnValue(true);
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({
        id: 'scheduled-2',
        layoutType: 'general',
        layoutName: 'General'
      })
    ]);
    
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    await waitFor(() => {
      const patchButtons = screen.getAllByTitle('Edit on main storyboard');
      if (patchButtons.length > 0) {
        fireEvent.click(patchButtons[0]);
      }
    });

    expect(globalThis.confirm).toHaveBeenCalled();
  });

  it('filters stories based on search term', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const searchInput = screen.getByRole('textbox', { name: /Search available stories/i });
    fireEvent.change(searchInput, { target: { value: 'Story 1' } });
    
    expect(screen.getByText('Test Story 1')).toBeInTheDocument();
    expect(screen.queryByText('Test Story 2')).not.toBeInTheDocument();
  });

  it('does not affect stories already placed on board while searching', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    act(() => {
      mockOnDragEnd!({
        active: {
          id: 'story-1',
          data: { current: { sortable: { containerId: 'available-stories' } } },
        },
        over: { id: 'slot-1' },
      });
    });

    const searchInput = screen.getByRole('textbox', { name: /Search available stories/i });
    fireEvent.change(searchInput, { target: { value: 'Story 1' } });

    expect(screen.queryByText('Test Story 1')).not.toBeInTheDocument();
    expect(screen.getByText(/No stories found matching/i)).toBeInTheDocument();

    fireEvent.click(screen.getByText('Remove'));
    expect(screen.getByText('Test Story 1')).toBeInTheDocument();
  });

  it('does not match stories by description text', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const searchInput = screen.getByPlaceholderText(/Search stories by title/i);
    fireEvent.change(searchInput, { target: { value: 'Another test' } });
    
    expect(screen.queryByText('Test Story 1')).not.toBeInTheDocument();
    expect(screen.queryByText('Test Story 2')).not.toBeInTheDocument();
    expect(screen.queryByText('Test Story 3')).not.toBeInTheDocument();
    expect(screen.getByText(/No stories found matching/i)).toBeInTheDocument();
  });

  it('shows empty state when no stories match search', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const searchInput = screen.getByPlaceholderText(/Search stories by title/i);
    fireEvent.change(searchInput, { target: { value: 'NonexistentStory' } });
    
    expect(screen.getByText(/No stories found matching/i)).toBeInTheDocument();
  });

  it('removes story from slot', async () => {
    render(<StoryDragDrop context={mockContext as any} />);
    
    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const removeButton = screen.getByText('Remove');
    fireEvent.click(removeButton);
    
    expect(screen.getByTestId('layout-renderer')).toBeInTheDocument();
  });

  it('shows validation toast when publishing an empty board', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Publish Board/i));

    await waitFor(() => {
      expect(screen.getByText('Cannot Publish Empty Board')).toBeInTheDocument();
    });
    expect(publishStoriesService.publishBoard).not.toHaveBeenCalled();
  });

  it('resets board to previously published state when available', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {
        'slot-1': 'thumbnail-text',
      },
    });

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText('Remove'));

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));

    await waitFor(() => {
      expect(screen.getByText('Board Reset to Last Published State')).toBeInTheDocument();
    });
  });

  it('clears board when reset finds no published state', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValueOnce(null);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    act(() => {
      mockOnDragEnd!({
        active: {
          id: 'story-1',
          data: { current: { sortable: { containerId: 'available-stories' } } },
        },
        over: { id: 'slot-1' },
      });
    });

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));

    await waitFor(() => {
      expect(screen.getByText('Board Reset')).toBeInTheDocument();
      expect(screen.getByText('No previously published board state found. Board cleared.')).toBeInTheDocument();
    });
  });

  it('shows unsupported layout toast when reset state is not connectHomepage', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'general',
      layoutName: 'General',
      slotStories: {
        'gen-slot-1': mockStories[0],
      },
      slotLayoutPreferences: {},
    });

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Clear Board/i));
    fireEvent.click(screen.getByText('Yes, Clear Board'));

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));

    await waitFor(() => {
      expect(screen.getByText('Layout Type Does Not Exist')).toBeInTheDocument();
    });
  });

  it('shows reset failure toast when published board fetch fails', async () => {
    const consoleSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
    (publishStoriesService.getLiveBoard as jest.Mock).mockRejectedValue(new Error('reset failed'));

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    act(() => {
      mockOnDragEnd!({
        active: {
          id: 'story-1',
          data: { current: { sortable: { containerId: 'available-stories' } } },
        },
        over: { id: 'slot-1' },
      });
    });

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));

    await waitFor(() => {
      expect(screen.getByText('Reset failed')).toBeInTheDocument();
    });
    consoleSpy.mockRestore();
  });

  it('publishes successfully when board has stories after reset', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {},
    });

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText('Remove'));

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));
    await waitFor(() => {
      expect(screen.getByText('Board Reset to Last Published State')).toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Publish Board/i));

    await waitFor(() => {
      expect(publishStoriesService.publishBoard).toHaveBeenCalled();
      expect(screen.getByText('Storyboard Published!')).toBeInTheDocument();
    });
  });

  it('shows publish failure toast when publish request fails', async () => {
    (publishStoriesService.getLiveBoard as jest.Mock).mockResolvedValue({
      layoutType: 'connectHomepage',
      layoutName: 'Connect Homepage',
      slotStories: {
        'slot-1': mockStories[0],
      },
      slotLayoutPreferences: {},
    });
    (publishStoriesService.publishBoard as jest.Mock).mockRejectedValueOnce(new Error('publish failed'));

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getByText('Remove'));

    fireEvent.click(screen.getByText(/Reset Board/i));
    fireEvent.click(screen.getByText('Yes, Reset Board'));
    await waitFor(() => {
      expect(screen.getByText('Board Reset to Last Published State')).toBeInTheDocument();
    });

    fireEvent.click(screen.getByText(/Publish Board/i));

    await waitFor(() => {
      expect(screen.getByText('Publish Failed')).toBeInTheDocument();
    });
  });

  it('shows delete failure toast when deleting a scheduled group fails', async () => {
    const consoleSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({ id: 'scheduled-delete-fail' })
    ]);
    (scheduleStoriesService.deleteScheduledGroupByGroupId as jest.Mock).mockRejectedValueOnce(new Error('delete failed'));

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getAllByTitle('Delete this scheduled group')[0]);
    fireEvent.click(screen.getByText('Confirm Delete Group'));

    await waitFor(() => {
      expect(screen.getByText('Delete Failed')).toBeInTheDocument();
    });
    consoleSpy.mockRestore();
  });

  it('does not enter group editing when user declines layout switch', async () => {
    (globalThis.confirm as jest.Mock).mockReturnValueOnce(false);
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({
        id: 'scheduled-decline-switch',
        layoutType: 'general',
        layoutName: 'General',
      })
    ]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getAllByTitle('Edit on main storyboard')[0]);

    expect(globalThis.confirm).toHaveBeenCalled();
    expect(screen.queryByText(/Save Group Edits/i)).not.toBeInTheDocument();
  });

  it('saves group edits through SharePoint when group has a SharePoint id', async () => {
    (scheduleStoriesService.getScheduledGroups as jest.Mock).mockResolvedValueOnce([
      buildScheduledGroup({ id: 'scheduled-sp-id', spId: 17 })
    ]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    fireEvent.click(screen.getAllByTitle('Edit on main storyboard')[0]);
    fireEvent.click(screen.getByText(/Save Group Edits/i));

    await waitFor(() => {
      expect(scheduleStoriesService.rescheduleGroup).toHaveBeenCalled();
      expect(screen.getByText('Group Updated Successfully!')).toBeInTheDocument();
    });
  });

  it('clears search term when clear-search button is clicked', async () => {
    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    const searchInput = screen.getByPlaceholderText(/Search stories by title/i) as HTMLInputElement;
    fireEvent.change(searchInput, { target: { value: 'Story 1' } });
    expect(searchInput.value).toBe('Story 1');

    fireEvent.click(screen.getByLabelText(/Clear search/i));
    expect(searchInput.value).toBe('');
  });

  it('shows empty available-state and disabled scheduling when no stories are loaded', async () => {
    (availableStoriesService.getStories as jest.Mock).mockResolvedValueOnce([]);

    render(<StoryDragDrop context={mockContext as any} />);

    await waitFor(() => {
      expect(screen.queryByRole('status')).not.toBeInTheDocument();
    });

    expect(screen.getByText('No stories available')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Schedule New Group/i })).toBeDisabled();
  });
});
