import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  DndContext,
  DragEndEvent,
  MouseSensor,
  TouchSensor,
  useSensor,
  useSensors,
  pointerWithin,
} from '@dnd-kit/core';
import { SortableContext, verticalListSortingStrategy, useSortable } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './StoryDragDrop.module.scss';
import AddStoryModal from './AddStoryModal';
import EditStoryModal from './EditStoryModal';
import DeleteGroupModal from './DeleteGroupModal';
import DateTimePicker from './DateTimePicker';
import { IStory, LayoutType, SlotStoryMap, ScheduledStoryGroup } from './types';
import { getStories, createStory, updateStory, deleteStory, initializeSharePoint, IStoryItem } from '../services/sharePointService';
import {
  createScheduledGroup,
  getScheduledGroups,
  deleteScheduledGroupByGroupId,
  rescheduleGroup,
} from '../services/scheduleStoriesService';
import { publishBoard } from '../services/publishStoriesService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import LayoutRenderer from './layouts/LayoutRenderer';
import { AVAILABLE_LAYOUTS, getLayoutConfig } from './layouts/layoutConfig';
import LayoutSelect from './LayoutSelect';
import { detectStorySource, formatPublishDate } from './cards/storyHelpers';
import { ToastContainer, IToast } from './Toast';

// Reference cap: up to 10 scheduled groups (each holds max 5 stories — one per slot).
const MAX_SCHEDULED_GROUPS = 10;

const formatTime12 = (time: string): string => {
  const parts = time.split(':');
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  const suffix = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return `${pad(h12)}:${pad(m)} ${suffix}`;
};

// Map raw SharePoint story items into the UI's IStory shape.
const formatStories = (stories: IStoryItem[]): IStory[] =>
  stories.map((story, index) => ({
    id: `story-${story.id || index}`,
    title: story.title,
    description: story.description || '',
    imageUrl: story.imageUrl,
    linkToPost: story.linkToPost || '#',
    date: story.date || new Date().toLocaleDateString(),
    created: story.created,
  }));




interface StoryDragDropProps {
  context?: IWebPartContext;
}

const StoryDragDrop: React.FC<StoryDragDropProps> = ({ context }) => {
  const [availableStories, setAvailableStories] = useState<IStory[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedLayout, setSelectedLayout] = useState<LayoutType>('reporterDaily');
  const [slotStories, setSlotStories] = useState<SlotStoryMap>({});
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [editingStory, setEditingStory] = useState<IStory | null>(null);
  const [isSchedulingMode, setIsSchedulingMode] = useState(false);
  const [scheduledGroups, setScheduledGroups] = useState<ScheduledStoryGroup[]>([]);
  const [selectedScheduleDate, setSelectedScheduleDate] = useState<string>('');
  const [selectedScheduleTime, setSelectedScheduleTime] = useState<string>('09:00');
  const [showCalendarOverlay, setShowCalendarOverlay] = useState(false);
  const [slotLayoutPreferences, setSlotLayoutPreferences] = useState<{ [key: string]: 'full-image' | 'thumbnail-text' }>({});
  const [featuredStoryId, setFeaturedStoryId] = useState<string | null>(null);
  const [isGroupEditingMode, setIsGroupEditingMode] = useState(false);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);
  const [originalBoardState, setOriginalBoardState] = useState<SlotStoryMap>({});
  const [rescheduleDate, setRescheduleDate] = useState<string>('');
  const [rescheduleTime, setRescheduleTime] = useState<string>('09:00');
  const [showRescheduleCalendar, setShowRescheduleCalendar] = useState(false);
  const [toasts, setToasts] = useState<IToast[]>([]);
  const [groupToDelete, setGroupToDelete] = useState<ScheduledStoryGroup | null>(null);


  const showToast = (title: string, type: 'success' | 'error' | 'info' = 'success', description?: string): void => {
    const id = `toast-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
    setToasts(prev => [...prev, { id, type, title, description }]);
  };

  const handleCloseToast = (id: string): void => {
    setToasts(prev => prev.filter(t => t.id !== id));
  };

  // Load stories from SharePoint on component mount
  useEffect(() => {
    let isMounted = true; // Track if component is still mounted

    const loadStories = async (): Promise<void> => {
      try {
        if (!context) {
          if (isMounted) {
            setError('SharePoint context not available. Cannot load stories.');
            setIsLoading(false);
          }
          return;
        }
        
        initializeSharePoint(context);
        
        await new Promise(resolve => setTimeout(resolve, 200));
        
        try {
          const stories = await getStories();
          const formattedStories = formatStories(stories);
          
          // Only update state if component is still mounted
          if (isMounted) {
            setAvailableStories(formattedStories);
            setError(null);
          }
        } catch (err) {
          console.error('Failed to load stories from SharePoint:', err);
          if (isMounted) {
            setError('Failed to load stories from SharePoint');
            setAvailableStories([]);
          }
        }

        // Load persisted scheduled groups from the "Schedule Stories" list.
        // A failure here shouldn't block the stories UI, so it's logged only.
        try {
          const groups = await getScheduledGroups();
          if (isMounted) {
            setScheduledGroups(groups);
          }
        } catch (err) {
          console.error('Failed to load scheduled groups from SharePoint:', err);
        }
      } catch (err) {
        console.error('Failed to load stories:', err);
        if (isMounted) {
          setError('Failed to load stories from SharePoint');
        }
      } finally {
        if (isMounted) {
          setIsLoading(false);
        }
      }
    };

    loadStories().catch(() => {
      if (isMounted) {
        setIsLoading(false);
        setError('Failed to load stories');
      }
    });

    // Cleanup function to prevent state updates after unmount
    return () => {
      isMounted = false;
    };
  }, [context]);

  // Configure sensors for dnd-kit. Mouse and touch are handled by separate
  // sensors so each can have the right activation gesture:
  //  - Mouse: start dragging after 8px of movement.
  //  - Touch: press-and-hold (200ms) before dragging, so a quick swipe still
  //    scrolls the Available Stories list instead of being swallowed as a drag.
  //    Without a dedicated TouchSensor the browser claims the touch as a scroll
  //    and the drag never starts on mobile / touch emulation.
  const sensors = useSensors(
    useSensor(MouseSensor, {
      activationConstraint: {
        distance: 8, // Require 8px movement before drag starts
      },
    }),
    useSensor(TouchSensor, {
      activationConstraint: {
        delay: 200, // Hold for 200ms to begin a drag
        tolerance: 8, // Allow 8px of finger jitter during the hold
      },
    })
  );

  const handleDragEnd = (event: DragEndEvent): void => {
    const { active, over } = event;
    if (!over) return;

    const activeStoryId = active.id as string;
    const overId = over.id as string;
    const fromAvailable =
      active.data.current?.sortable?.containerId === 'available-stories';

    // Valid slot ids for the current layout.
    const slotIds = getLayoutConfig(selectedLayout).slots.map(s => s.id);

    // Resolve a drop target (which may be a slot droppable OR a card sitting
    // in a slot) to the owning slot id. Returns undefined if it's neither.
    const resolveDestSlot = (id: string): string | undefined => {
      if (slotIds.indexOf(id) !== -1) return id;
      return Object.keys(slotStories).find(
        key => slotStories[key]?.id === id
      );
    };

    // --- Dragging a card FROM the Available list ---------------------
    if (fromAvailable) {
      const story = availableStories.find(s => s.id === activeStoryId);
      if (!story) return;
      if (overId === 'available-stories') return; // dropped back on the list

      const destSlot = resolveDestSlot(overId);
      if (!destSlot) return; // not a real slot (e.g. reordering the list)

      const occupant = slotStories[destSlot];
      setSlotStories(prev => ({ ...prev, [destSlot]: story }));
      setAvailableStories(prev => {
        let next = prev.filter(s => s.id !== story.id);
        // Return whatever was already in that slot to the list.
        if (occupant && !next.some(s => s.id === occupant.id)) {
          next = [...next, occupant];
        }
        return next;
      });
      return;
    }

    // --- Dragging a card that is already in a slot ------------------
    const sourceSlot = Object.keys(slotStories).find(
      key => slotStories[key]?.id === activeStoryId
    );
    if (!sourceSlot) return;
    const sourceStory = slotStories[sourceSlot];

    // Slot -> back to the Available list
    if (overId === 'available-stories') {
      setSlotStories(prev => {
        const next = { ...prev };
        delete next[sourceSlot];
        return next;
      });
      setAvailableStories(prev =>
        sourceStory && !prev.some(s => s.id === sourceStory.id)
          ? [...prev, sourceStory]
          : prev
      );
      return;
    }

    // Slot -> another slot (move into empty, or swap with occupant)
    const destSlot = resolveDestSlot(overId);
    if (!destSlot || destSlot === sourceSlot) return;

    const destStory = slotStories[destSlot];
    setSlotStories(prev => {
      const next = { ...prev };
      if (destStory) {
        next[sourceSlot] = destStory;
        next[destSlot] = sourceStory;
      } else {
        next[destSlot] = sourceStory;
        delete next[sourceSlot];
      }
      return next;
    });
  };

  const handleAddStory = (story: Omit<IStory, 'id' | 'date'>): void => {
    createStory({
      title: story.title,
      description: story.description,
      imageUrl: story.imageUrl,
      linkToPost: story.linkToPost
    }).then(() => {
      return getStories();
    }).then((stories) => {
      const formattedStories: IStory[] = stories.map((s, index) => ({
        id: `story-${s.id || index}`,
        title: s.title,
        description: s.description || '',
        imageUrl: s.imageUrl,
        linkToPost: s.linkToPost || '#',
        date: s.date || new Date().toLocaleDateString(),
        created: s.created
      }));
      setAvailableStories(formattedStories);
      setIsModalOpen(false);
      showToast('Story Added Successfully!', 'success');
    }).catch((err) => {
      console.error('Failed to save or reload stories:', err);
      setIsModalOpen(false);
      showToast('Failed to add story', 'error', err.message || 'Unknown error');
    });
  };

  const handleEditStory = (storyId: string): void => {
    const story = availableStories.find(s => s.id === storyId);
    if (story) {
      setEditingStory(story);
    }
  };

  const handleUpdateStory = (updatedStory: IStory): void => {
    // Extract ID from story.id (format: "story-{id}")
    const storyIdMatch = updatedStory.id.match(/\d+/);
    const numericId = storyIdMatch ? parseInt(storyIdMatch[0], 10) : null;

    if (!numericId) {
      console.error('Invalid story ID format:', updatedStory.id);
      showToast('Error: Invalid story ID', 'error');
      return;
    }

    // Call API to update story
    updateStory(numericId, {
      title: updatedStory.title,
      description: updatedStory.description,
      imageUrl: updatedStory.imageUrl,
      linkToPost: updatedStory.linkToPost
    }).then(() => {
      // Refresh stories from API
      return getStories();
    }).then((stories) => {
      const formattedStories = formatStories(stories);
      setAvailableStories(formattedStories);
      
      // Update in slot stories
      const newSlotStories = { ...slotStories };
      Object.keys(newSlotStories).forEach(key => {
        const updatedSlotStory = formattedStories.find(s => s.id === newSlotStories[key]?.id);
        if (updatedSlotStory) {
          newSlotStories[key] = updatedSlotStory;
        }
      });
      setSlotStories(newSlotStories);
      setEditingStory(null);
      showToast('Story Updated Successfully!', 'success');
    }).catch((err) => {
      console.error('Failed to update story:', err);
      showToast('Failed to update story', 'error', err.message || 'Unknown error');
    });
  };

  const handleDeleteStory = (storyId: string): void => {
    // Extract ID from story.id (format: "story-{id}")
    const storyIdMatch = storyId.match(/\d+/);
    const numericId = storyIdMatch ? parseInt(storyIdMatch[0], 10) : null;

    if (!numericId) {
      console.error('Invalid story ID format:', storyId);
      showToast('Error: Invalid story ID', 'error');
      return;
    }

    // Call API to delete story
    deleteStory(numericId).then(() => {
      // Refresh stories from API
      return getStories();
    }).then((stories) => {
      setAvailableStories(formatStories(stories));
      
      // Remove from slot stories if it's there
      const newSlotStories = { ...slotStories };
      Object.keys(newSlotStories).forEach(key => {
        if (newSlotStories[key]?.id === storyId) {
          delete newSlotStories[key];
        }
      });
      setSlotStories(newSlotStories);
      // Clear featured state if the deleted story was featured.
      setFeaturedStoryId(prev => (prev === storyId ? null : prev));
      setEditingStory(null);
      showToast('Story deleted successfully!', 'success');
    }).catch((err) => {
      console.error('Failed to delete story:', err);
      showToast('Failed to delete story', 'error', err.message || 'Unknown error');
    });
  };

  const handleRemoveFromSlot = (slotId: string): void => {
    const story = slotStories[slotId];
    if (story) {
      // Add back to available stories
      setAvailableStories([...availableStories, story]);
      
      // Remove from slot
      const newSlotStories = { ...slotStories };
      delete newSlotStories[slotId];
      setSlotStories(newSlotStories);
    }
  };

  const handleClearBoard = (): void => {
    // Move all stories from slots back to available
    const storiesInSlots = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== null);
    setAvailableStories([...availableStories, ...storiesInSlots]);
    setSlotStories({});
    showToast('Board Reset', 'info', 'All stories returned to available list.');
  };

  const handleLayoutChange = (value: string): void => {
    setSelectedLayout(value as LayoutType);
    showToast('Layout Changed', 'info', `Switched layout to ${getLayoutConfig(value as LayoutType).name}`);
    // Optionally clear the board when switching layouts
    // handleClearBoard();
  };

  const layoutOptions = AVAILABLE_LAYOUTS.map(layoutType => ({
    value: layoutType,
    label: getLayoutConfig(layoutType).name
  }));

  const handleSlotLayoutChange = (
    slotId: string,
    mode: 'full-image' | 'thumbnail-text'
  ): void => {
    setSlotLayoutPreferences(prev => ({ ...prev, [slotId]: mode }));
  };

  // Only one story can be featured at a time (reference behaviour):
  // clicking the star on the featured story unfeatures it.
  const handleToggleFeatured = (storyId: string): void => {
    setFeaturedStoryId(prev => {
      const isFeature = prev !== storyId;
      if (isFeature) {
        showToast('Story marked as featured', 'success');
      } else {
        showToast('Story unfeatured', 'success');
      }
      return isFeature ? storyId : null;
    });
  };

  const handleScheduleGroup = (): void => {
    // Save current board state
    setOriginalBoardState({ ...slotStories });

    // Clear the board slots
    const clearedSlots: SlotStoryMap = {};
    Object.keys(slotStories).forEach(key => {
      clearedSlots[key] = undefined;
    });
    setSlotStories(clearedSlots);

    setIsSchedulingMode(true);
    setShowCalendarOverlay(false);
    setSelectedScheduleDate('');
    setSelectedScheduleTime('09:00');
  };

  const handleExitScheduling = (): void => {
    // Restore original board state
    setSlotStories({ ...originalBoardState });
    setOriginalBoardState({});

    setIsSchedulingMode(false);
    setShowCalendarOverlay(false);
    setSelectedScheduleDate('');
    setSelectedScheduleTime('09:00');
  };

  const handleOpenDateTimePicker = (): void => {
    setShowCalendarOverlay(true);
  };

  const handleScheduleNow = (): void => {
    if (!selectedScheduleDate) {
      showToast('Please select a date', 'error');
      return;
    }

    const boardStories = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== undefined);

    if (boardStories.length === 0) {
      showToast('No Stories to Schedule', 'error', 'Please add stories to the main board before scheduling.');
      return;
    }

    // Reference cap: at most 10 scheduled groups.
    if (scheduledGroups.length >= MAX_SCHEDULED_GROUPS) {
      showToast('Maximum Groups Reached', 'error', `You can only have up to ${MAX_SCHEDULED_GROUPS} scheduled groups.`);
      return;
    }

    // Reference behaviour: one group per date.
    const scheduleDate = new Date(selectedScheduleDate);
    const dateTaken = scheduledGroups.some(
      g => g.date.toDateString() === scheduleDate.toDateString()
    );
    if (dateTaken) {
      showToast('Date Already Scheduled', 'error', "There's already a scheduled group for this date.");
      return;
    }

    const layoutConfig = getLayoutConfig(selectedLayout);
    const scheduledGroup: ScheduledStoryGroup = {
      id: `scheduled-${Date.now()}`,
      date: new Date(selectedScheduleDate),
      time: selectedScheduleTime,
      slotStories: { ...slotStories }, // Store the exact slot-to-story mapping
      layoutType: selectedLayout, // Store the layout type for validation
      layoutName: layoutConfig.name,
      createdAt: new Date()
    };

    // Persist to the "Schedule Stories" list, then reflect the saved record
    // (which carries the SharePoint item id) in state.
    createScheduledGroup(scheduledGroup)
      .then((saved) => {
        setScheduledGroups(prev => [...prev, saved]);
        handleExitScheduling();

        const formattedDate = new Date(selectedScheduleDate).toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric'
        });
        showToast('Board Scheduled Successfully!', 'success', `Scheduled for ${formattedDate} at ${formatTime12(selectedScheduleTime)}.`);
      })
      .catch((err) => {
        console.error('Failed to save scheduled group:', err);
        showToast('Scheduling Failed', 'error', 'Could not save the scheduled group to SharePoint. Please try again.');
      });
  };

  const handleDeleteScheduledGroup = (groupId: string): void => {
    const groupToRemove = scheduledGroups.find(g => g.id === groupId);
    if (groupToRemove) {
      setGroupToDelete(groupToRemove);
    }
  };

  const handleDeleteScheduledGroupConfirmed = (groupId: string): void => {
    const groupToRemove = scheduledGroups.find(g => g.id === groupId);
    if (!groupToRemove) return;

    deleteScheduledGroupByGroupId(groupId)
      .then(() => {
        // Return stories to available list (preventing duplicates)
        const groupStories = Object.keys(groupToRemove.slotStories)
          .map(key => groupToRemove.slotStories[key])
          .filter((s): s is IStory => s !== undefined);
        
        setAvailableStories(prev => {
          const existingIds = new Set(prev.map(s => s.id));
          const toAdd = groupStories.filter(s => !existingIds.has(s.id));
          return [...toAdd, ...prev];
        });

        // Filter group from state and close overlay
        setScheduledGroups(prev => prev.filter(g => g.id !== groupId));
        setGroupToDelete(null);

        const formattedDate = groupToRemove.date.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric'
        });
        showToast('Scheduled Group Removed', 'success', `Stories scheduled for ${formattedDate} returned to available list.`);
      })
      .catch((err) => {
        console.error('Failed to delete scheduled group:', err);
        showToast('Delete Failed', 'error', 'Could not remove the scheduled group from SharePoint.');
        setGroupToDelete(null);
      });
  };

  const handleExitGroupEditingMode = (): void => {
    // 1. Move any stories currently on the board back to available stories
    const storiesFromBoard: IStory[] = [];
    Object.keys(slotStories).forEach(key => {
      const story = slotStories[key];
      if (story) storiesFromBoard.push(story);
    });
    
    if (storiesFromBoard.length > 0) {
      setAvailableStories(prev => {
        const existingIds = new Set(prev.map(s => s.id));
        const toAdd = storiesFromBoard.filter(s => !existingIds.has(s.id));
        return [...toAdd, ...prev];
      });
    }
    
    // 2. Restore the original board state
    setSlotStories({ ...originalBoardState });
    
    // 3. Clear editing state
    setIsGroupEditingMode(false);
    setEditingGroupId(null);
    setOriginalBoardState({});
    setShowRescheduleCalendar(false);
  };

  const handleEnterGroupEditingMode = (group: ScheduledStoryGroup): void => {
    // Check if the scheduled group's layout matches the current layout
    if (group.layoutType !== selectedLayout) {
      const confirmSwitch = window.confirm(
        `This group was created for "${group.layoutName}" layout, but you're currently viewing "${getLayoutConfig(selectedLayout).name}". ` +
        `Do you want to switch to the "${group.layoutName}" layout?`
      );
      
      if (confirmSwitch) {
        setSelectedLayout(group.layoutType);
        showToast('Layout Switched', 'info', `Switched layout to ${group.layoutName}`);
      } else {
        return; // User cancelled, don't patch
      }
    }

    // Save current board state
    setOriginalBoardState({ ...slotStories });

    // Load the group's stories onto the main board
    setSlotStories({ ...group.slotStories });
    
    // Set group editing mode state
    setIsGroupEditingMode(true);
    setEditingGroupId(group.id);
    
    // Format group date as YYYY-MM-DD
    const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
    const y = group.date.getFullYear();
    const m = pad2(group.date.getMonth() + 1);
    const d = pad2(group.date.getDate());
    setRescheduleDate(`${y}-${m}-${d}`);
    setRescheduleTime(group.time);

    // Remove scheduled stories from available list (if they are not already placed on board, but we want to make sure they are not duplicated)
    const scheduledStoryIds = Object.keys(group.slotStories)
      .map(key => group.slotStories[key])
      .filter((story): story is IStory => story !== undefined)
      .map(story => story.id);
    
    // Move any stories currently on the board back to available stories (except the ones in this group!)
    const storiesFromBoard: IStory[] = [];
    Object.keys(slotStories).forEach(key => {
      const story = slotStories[key];
      if (story) storiesFromBoard.push(story);
    });
    
    const newAvailable = [...availableStories];
    storiesFromBoard.forEach(story => {
      if (!newAvailable.some(s => s.id === story.id)) {
        newAvailable.push(story);
      }
    });

    const finalAvailable = newAvailable.filter(
      story => scheduledStoryIds.indexOf(story.id) === -1
    );

    setAvailableStories(finalAvailable);
    showToast('Group Editing Mode', 'info', 'Rearrange stories on the main board.');
  };

  const handleSaveGroupEdits = (): void => {
    if (!editingGroupId) return;
    const group = scheduledGroups.find(g => g.id === editingGroupId);
    if (!group) return;

    // Compute the new display date (local) from the picker, if the date changed.
    let newDate = group.date;
    if (rescheduleDate) {
      const parts = rescheduleDate.split('-');
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      const d = parseInt(parts[2], 10);
      const timeParts = rescheduleTime.split(':');
      const hr = parseInt(timeParts[0], 10);
      const min = parseInt(timeParts[1], 10);
      newDate = new Date(y, m, d, hr, min);
    }
    const newSlotStories = { ...slotStories };

    const applyLocal = (): void => {
      setScheduledGroups(prev => prev.map(g =>
        g.id === editingGroupId
          ? { ...g, slotStories: newSlotStories, date: newDate, time: rescheduleTime }
          : g
      ));
      handleExitGroupEditingMode();
      showToast('Group Updated Successfully!', 'success', 'Saved scheduled group with new arrangement.');
    };

    if (typeof group.spId === 'number') {
      // Persist to SharePoint. Use a UTC-midnight date built from the picker
      // string (timezone-proof) when the date changed, else the group's own.
      const serviceDate = rescheduleDate ? new Date(`${rescheduleDate}T00:00:00Z`) : group.date;
      rescheduleGroup(group.spId, serviceDate, rescheduleTime, newSlotStories)
        .then(applyLocal)
        .catch((err) => {
          console.error('Failed to reschedule group:', err);
          showToast('Update Failed', 'error', 'Could not update the scheduled group in SharePoint.');
        });
    } else {
      // Group was never persisted (no SharePoint id) — update locally only.
      applyLocal();
    }
  };

  const handleCancelGroupEdits = (): void => {
    if (!editingGroupId) return;
    handleExitGroupEditingMode();
    showToast('Group Edits Cancelled', 'info', 'No changes were saved.');
  };

  const handleConfirmReschedule = (): void => {
    setShowRescheduleCalendar(false);
    showToast('Group Rescheduled', 'success', 'New date/time confirmed. Click Save Group Edits to apply.');
  };

  // Publish the current board live: supersede the previous live board and write
  // an immutable snapshot to the "Publish Stories" list.
  const handlePublishBoard = (): void => {
    const boardStories = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== undefined);

    if (boardStories.length === 0) {
      showToast('Nothing to Publish', 'error', 'Add at least one story to the board before publishing.');
      return;
    }

    const layoutConfig = getLayoutConfig(selectedLayout);
    publishBoard({
      layoutType: selectedLayout,
      layoutName: layoutConfig.name,
      slotStories: { ...slotStories }
    })
      .then(() => {
        const count = boardStories.length;
        showToast('Board Published!', 'success', `${count} ${count === 1 ? 'story is' : 'stories are'} now live.`);
      })
      .catch((err) => {
        console.error('Failed to publish board:', err);
        showToast('Publish Failed', 'error', 'Could not publish the board to SharePoint. Please try again.');
      });
  };

  // Search matches on title only.
  const filteredStories = availableStories.filter(story =>
    story.title.toLowerCase().includes(searchTerm.toLowerCase())
  );



  // Row in the Available Stories list. Only the left region is a drag
  // handle (reference behaviour); star/edit actions live outside it.
  const AvailableStoryItem: React.FC<{
    story: IStory;
    isFeatured: boolean;
    onEdit: (storyId: string) => void;
    onToggleFeatured: (storyId: string) => void;
  }> = ({ story, isFeatured, onEdit, onToggleFeatured }) => {
    const {
      attributes,
      listeners,
      setNodeRef,
      transform,
      transition,
      isDragging,
    } = useSortable({
      id: story.id,
      data: {
        sortable: {
          containerId: 'available-stories',
        },
      },
    });

    const style = {
      transform: CSS.Translate.toString(transform),
      transition,
    };

    const { source, domain } = detectStorySource(story.linkToPost);

    return (
      <div
        ref={setNodeRef}
        style={style}
        {...attributes}
        className={`${styles.storyCard} ${isDragging ? styles.dragging : ''} ${isFeatured ? styles.storyCardFeatured : ''}`}
      >
        <div className={styles.storyDragArea} {...listeners}>
          <img src={story.imageUrl} alt={story.title} className={styles.storyImage} />
          <div className={styles.storyContent}>
            <div className={styles.storyTitleRow}>
              <h4>{story.title}</h4>
              {isFeatured && (
                <Icon iconName="FavoriteStarFill" className={styles.featuredStar} />
              )}
            </div>
            <div className={styles.storyMeta}>
              <span className={styles.storyDate}>{formatPublishDate(story.date)}</span>
              {/* Commented out source badge (e.g. google.com / LinkedIn / External)
              {source !== 'internal' && (
                <span
                  className={`${styles.rowBadge} ${source === 'linkedin' ? styles.rowBadgeLinkedin : styles.rowBadgeExternal}`}
                >
                  <Icon iconName={source === 'linkedin' ? 'LinkedInLogo' : 'Globe'} />
                  {source === 'linkedin' ? 'LinkedIn' : domain || 'External'}
                </span>
              )}
              */}
            </div>
          </div>
        </div>
        <div className={styles.storyActions}>
          {/* Commented out feature button
          <button
            type="button"
            className={`${styles.actionBtn} ${isFeatured ? styles.starBtnActive : ''}`}
            onClick={() => onToggleFeatured(story.id)}
            title={isFeatured ? 'Remove featured' : 'Mark as featured'}
            aria-label={isFeatured ? 'Remove featured' : 'Mark as featured'}
            aria-pressed={isFeatured}
          >
            <Icon iconName={isFeatured ? 'FavoriteStarFill' : 'FavoriteStar'} />
          </button>
          */}
          <button
            type="button"
            className={styles.actionBtn}
            onClick={() => onEdit(story.id)}
            title="Edit story"
            aria-label="Edit story"
          >
            <Icon iconName="Edit" />
          </button>
        </div>
      </div>
    );
  };

  return (
    <div className={styles.storyDragDrop}>
      {isLoading && (
        <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>
          Loading stories from SharePoint...
        </div>
      )}
      {error && (
        <div style={{ padding: '20px', backgroundColor: '#fee', color: '#c00', borderRadius: '4px', marginBottom: '20px' }}>
          {error}
        </div>
      )}
      <div className={styles.header}>
        <div className={styles.headerLeft}>
          <div className={styles.layoutSelectContainer}>
            <label className={styles.layoutLabel}>View:</label>
            <LayoutSelect
              className={styles.layoutSelect}
              value={selectedLayout}
              options={layoutOptions}
              onChange={handleLayoutChange}
              disabled={isSchedulingMode || isGroupEditingMode}
              ariaLabel="Select layout view"
            />
          </div>
          <button
            type="button"
            className={styles.btnPrimary}
            onClick={() => setIsModalOpen(true)}
            disabled={isSchedulingMode || isGroupEditingMode}
          >
            <Icon iconName="Add" />
            <span className={styles.btnLabelFull}>Add story</span>
            <span className={styles.btnLabelShort}>Add</span>
          </button>
          <button
            type="button"
            className={styles.btnPublish}
            onClick={handlePublishBoard}
            disabled={isSchedulingMode || isGroupEditingMode}
          >
            <Icon iconName="PublishContent" />
            <span className={styles.btnLabelFull}>Publish Board</span>
            <span className={styles.btnLabelShort}>Publish</span>
          </button>
        </div>
        <div className={styles.headerRight}>
          <button
            type="button"
            className={styles.btnReset}
            onClick={handleClearBoard}
            disabled={isSchedulingMode || isGroupEditingMode}
          >
            <Icon iconName="Refresh" />
            <span className={styles.btnLabelFull}>Reset Board</span>
            <span className={styles.btnLabelShort}>Reset</span>
          </button>
          <button
            type="button"
            className={styles.btnClear}
            onClick={handleClearBoard}
            disabled={isSchedulingMode || isGroupEditingMode}
          >
            <Icon iconName="Delete" />
            <span className={styles.btnLabelFull}>Clear Board</span>
            <span className={styles.btnLabelShort}>Clear</span>
          </button>
        </div>
      </div>

      <DndContext 
        sensors={sensors}
        collisionDetection={pointerWithin}
        onDragEnd={handleDragEnd}
      >
        <div className={styles.layoutContainer}>
          
          {/* Dynamic Layout Renderer */}
          <div className={`${styles.mainLayout} ${
            isSchedulingMode ? styles.schedulingActive : 
            isGroupEditingMode ? styles.groupEditingActive : ''
          }`}>
            {isSchedulingMode && (
              <div className={styles.modeBannerContainer}>
                <div className={styles.schedulingBanner}>
                  📅 <span>Scheduling Mode: Drag stories from the Available Stories section to the slots below</span>
                </div>
              </div>
            )}
            {isGroupEditingMode && (
              <div className={styles.modeBannerContainer}>
                <div className={styles.groupEditingBanner}>
                  ✏️ <span>Group Editing Mode: Rearrange stories as they will appear in the final layout</span>
                </div>
              </div>
            )}
            <LayoutRenderer
              layoutType={selectedLayout}
              slotStories={slotStories}
              onRemoveStory={handleRemoveFromSlot}
              slotLayoutPreferences={slotLayoutPreferences}
              onSlotLayoutChange={handleSlotLayoutChange}
              featuredStoryId={featuredStoryId || undefined}
            />
          </div>

          {/* Bottom Section */}
          <div className={styles.bottomSection}>
            {/* Scheduled Stories */}
            <div className={styles.scheduledSection}>
              <div className={styles.sectionHeader}>
                <h3>Scheduled Stories</h3>
                {isGroupEditingMode ? (
                  <div className={styles.schedulingControlsInline}>
                    <div className={styles.groupEditingBadge}>
                      ✏️ Group Editing Mode
                    </div>
                    <div className={styles.buttonWithDropdown}>
                      <button 
                        className={styles.schedulingBtnSmall}
                        onClick={() => setShowRescheduleCalendar(!showRescheduleCalendar)}
                      >
                        <Icon iconName="Calendar" /> Reschedule
                      </button>
                      {showRescheduleCalendar && (
                        <div className={styles.calendarOverlay}>
                          <DateTimePicker
                            selectedDate={rescheduleDate}
                            time={rescheduleTime}
                            onSelectDate={setRescheduleDate}
                            onTimeChange={setRescheduleTime}
                            onConfirm={handleConfirmReschedule}
                            confirmText="Reschedule"
                            isDateDisabled={(date) => {
                              const today = new Date();
                              today.setHours(0, 0, 0, 0);
                              
                              const compareDate = new Date(date);
                              compareDate.setHours(0, 0, 0, 0);
                              
                              if (compareDate < today) return true;
                              
                              // Disable if already scheduled (excluding the current group being edited)
                              return scheduledGroups.some(g => 
                                g.id !== editingGroupId && 
                                g.date.toDateString() === compareDate.toDateString()
                              );
                            }}
                          />
                        </div>
                      )}
                    </div>
                    <button 
                      className={`${styles.schedulingBtnSmall} ${styles.saveGroupEditsBtn}`}
                      onClick={handleSaveGroupEdits}
                    >
                      <Icon iconName="CheckMark" /> Save Group Edits
                    </button>
                    <button 
                      className={styles.schedulingBtnSmall}
                      onClick={handleCancelGroupEdits}
                    >
                      <Icon iconName="Cancel" /> Cancel Edits
                    </button>
                  </div>
                ) : !isSchedulingMode ? (
                  scheduledGroups.length < MAX_SCHEDULED_GROUPS && (
                    <button
                      className={styles.btnSchedule}
                      onClick={handleScheduleGroup}
                      disabled={availableStories.length === 0}
                    >
                      <Icon iconName="Add" /> Schedule New Group
                    </button>
                  )
                ) : (
                  <div className={styles.schedulingControlsInline}>
                    <div className={styles.schedulingBadge}>
                      📅 Scheduling Mode
                    </div>
                    <div className={styles.buttonWithDropdown}>
                      <button 
                        className={styles.schedulingBtnSmall}
                        onClick={handleOpenDateTimePicker}
                      >
                        <Icon iconName="Calendar" /> {selectedScheduleDate && selectedScheduleTime ? (() => {
                          const parts = selectedScheduleDate.split('-');
                          const y = parseInt(parts[0], 10);
                          const m = parseInt(parts[1], 10) - 1;
                          const d = parseInt(parts[2], 10);
                          const timeParts = selectedScheduleTime.split(':');
                          const hr = parseInt(timeParts[0], 10);
                          const min = parseInt(timeParts[1], 10);
                          const dt = new Date(y, m, d, hr, min);
                          return dt.toLocaleDateString('en-US', {
                            month: 'short',
                            day: 'numeric',
                            year: 'numeric'
                          }) + ' ' + dt.toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                          });
                        })() : 'Select date & time'}
                      </button>
                      {showCalendarOverlay && isSchedulingMode && (
                        <div className={styles.calendarOverlay}>
                          <DateTimePicker
                            selectedDate={selectedScheduleDate}
                            time={selectedScheduleTime}
                            onSelectDate={setSelectedScheduleDate}
                            onTimeChange={setSelectedScheduleTime}
                            onConfirm={handleScheduleNow}
                            confirmText="Schedule"
                            isDateDisabled={(date) => {
                              const today = new Date();
                              today.setHours(0, 0, 0, 0);
                              
                              const compareDate = new Date(date);
                              compareDate.setHours(0, 0, 0, 0);
                              
                              if (compareDate < today) return true;
                              
                              // Disable if already scheduled
                              return scheduledGroups.some(g => 
                                g.date.toDateString() === compareDate.toDateString()
                              );
                            }}
                          />
                        </div>
                      )}
                    </div>
                    <button 
                      className={styles.schedulingBtnSmall}
                      onClick={handleExitScheduling}
                    >
                      <Icon iconName="Cancel" /> Exit Scheduling
                    </button>
                  </div>
                )}
              </div>

              {scheduledGroups.length === 0 ? (
                <div className={styles.emptyScheduled}>
                  <p>No scheduled story groups yet.</p>
                  <p className={styles.hint}>Use scheduling mode to create groups of stories for specific dates.</p>
                </div>
              ) : (
                <div className={styles.scheduledList}>
                  {scheduledGroups.map(group => {
                    const isEditingThis = editingGroupId === group.id;
                    return (
                      <div key={group.id} className={`${styles.scheduledGroup} ${isEditingThis ? styles.editingGroup : ''} ${isGroupEditingMode && !isEditingThis ? styles.disabledGroup : ''}`}>
                        <div className={styles.groupHeader}>
                          <div className={styles.groupInfo}>
                            <div>
                              <h4>
                                {group.date.toLocaleDateString('en-US', {
                                  weekday: 'long',
                                  month: 'long',
                                  day: 'numeric',
                                  year: 'numeric'
                                })}
                              </h4>
                              <p className={styles.groupMetaText}>
                                Scheduled for {formatTime12(group.time)} • {Object.keys(group.slotStories).map(key => group.slotStories[key]).filter(story => story !== undefined).length} of 5 stories
                                {isEditingThis && <span className={styles.editingLabel}> (Editing on Main Board)</span>}
                              </p>
                            </div>
                            <span className={styles.layoutTag}>{group.layoutName}</span>
                          </div>
                          <div className={styles.groupActions}>
                            <button
                              className={`${styles.groupActionBtn} ${styles.patchBtn}`}
                              onClick={() => handleEnterGroupEditingMode(group)}
                              title="Edit on main storyboard"
                              disabled={isSchedulingMode || isGroupEditingMode}
                            >
                              <Icon iconName="System" />
                            </button>
                            <button
                              className={`${styles.groupActionBtn} ${styles.deleteBtn}`}
                              onClick={() => handleDeleteScheduledGroup(group.id)}
                              title="Delete this scheduled group"
                              disabled={isSchedulingMode || isGroupEditingMode || isEditingThis}
                            >
                              <Icon iconName="Cancel" />
                            </button>
                          </div>
                        </div>
                        <div className={styles.groupSlotsGrid}>
                        {(() => {
                          const storiesList = Object.keys(group.slotStories)
                            .map(key => group.slotStories[key])
                            .filter((s): s is IStory => s !== undefined);
                          return Array.from({ length: 5 }).map((_, index) => {
                            const story = storiesList[index];
                            return (
                              <div key={index} className={styles.groupSlotWrapper}>
                                {story ? (
                                  <div className={styles.groupStoryCard}>
                                    <div className={styles.groupStoryImageWrap}>
                                      <img src={story.imageUrl} alt={story.title} />
                                    </div>
                                    <p>{story.title}</p>
                                  </div>
                                ) : (
                                  <div className={styles.groupSlotEmpty}>
                                    <span>Empty</span>
                                  </div>
                                )}
                              </div>
                            );
                          });
                        })()}
                      </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Available Stories */}
            <div className={styles.availableContainer}>
              <div className={styles.sidebarHeader}>
                <h3>Available Stories</h3>
              </div>
              <div className={styles.availableSection}>
                <div className={styles.searchContainer}>
                   <Icon iconName="Search" className={styles.searchIcon} />
                   <input
                    type="text"
                    placeholder="Search stories by title..."
                    className={styles.searchInput}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                  {searchTerm && (
                    <button
                      type="button"
                      className={styles.clearSearchBtn}
                      onClick={() => setSearchTerm('')}
                      aria-label="Clear search"
                      title="Clear search"
                    >
                      <Icon iconName="Cancel" />
                    </button>
                  )}
                </div>

                <SortableContext
                  items={filteredStories.map(s => s.id)}
                  strategy={verticalListSortingStrategy}
                  id="available-stories"
                >
                  <div className={styles.availableStoriesList}>
                    {filteredStories.length === 0 ? (
                      <div className={styles.emptyState}>
                        {searchTerm ? (
                          <>
                            <Icon iconName="Search" className={styles.emptySearchIcon} />
                            <p>No stories found matching &quot;{searchTerm}&quot;</p>
                            <p className={styles.emptyHint}>Try adjusting your search terms</p>
                          </>
                        ) : (
                          <p>No stories available</p>
                        )}
                      </div>
                    ) : (
                      filteredStories.map((story) => (
                        <AvailableStoryItem
                          key={story.id}
                          story={story}
                          isFeatured={featuredStoryId === story.id}
                          onEdit={handleEditStory}
                          onToggleFeatured={handleToggleFeatured}
                        />
                      ))
                    )}
                  </div>
                </SortableContext>
              </div>
            </div>
          </div>
        </div>
      </DndContext>

      {isModalOpen && (
        <AddStoryModal
          onClose={() => setIsModalOpen(false)}
          onAdd={handleAddStory}
        />
      )}

      {editingStory && (
        <EditStoryModal
          story={editingStory}
          onClose={() => setEditingStory(null)}
          onUpdate={handleUpdateStory}
          onDelete={handleDeleteStory}
        />
      )}

      {groupToDelete && (
        <DeleteGroupModal
          group={groupToDelete}
          onClose={() => setGroupToDelete(null)}
          onConfirm={() => handleDeleteScheduledGroupConfirmed(groupToDelete.id)}
        />
      )}

      <ToastContainer toasts={toasts} onCloseToast={handleCloseToast} />
    </div>
  );
};

export default StoryDragDrop;
