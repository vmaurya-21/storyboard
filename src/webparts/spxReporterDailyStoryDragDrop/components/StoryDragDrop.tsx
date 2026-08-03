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
import { IStory, LayoutType, SlotStoryMap, ScheduledStoryGroup, CardLayout } from './types';
import { getStories, createStory, updateStory, deleteStory, initializeSharePoint, IStoryItem } from '../services/availableStoriesService';
import { LinkedinIcon, ExternalLinkIcon, EditIcon } from './icons';
import BoardSkeleton from './BoardSkeleton';
import {
  createScheduledGroup,
  getScheduledGroups,
  deleteScheduledGroupByGroupId,
  rescheduleGroup,
  fromDateKey,
} from '../services/scheduleStoriesService';
import { publishBoard, getLiveBoard } from '../services/publishStoriesService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import LayoutRenderer from './layouts/LayoutRenderer';
import { AVAILABLE_LAYOUTS, getLayoutConfig } from './layouts/layoutConfig';
import { detectStorySource, formatPublishDate } from './cards/storyHelpers';
import { ToastContainer, IToast } from './Toast';

/** Most scheduled groups that may exist at once, matching the reference cap. */
const MAX_SCHEDULED_GROUPS = 10;

/**
 * Converts a 24-hour `HH:mm` string to a padded 12-hour label.
 *
 * @param time - Time of day as `HH:mm`.
 * @returns The 12-hour label, e.g. `13:05` becomes `01:05 PM`. Unparseable
 * parts fall back to zero.
 */
const formatTime12 = (time: string): string => {
  const parts = time.split(':');
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  const suffix = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return `${pad(h12)}:${pad(m)} ${suffix}`;
};

/**
 * Maps raw SharePoint list items to the view model the UI consumes.
 *
 * Ids are namespaced as `story-{id}`, falling back to the array index when an
 * item has none, so React keys and dnd-kit ids stay unique. Missing optional
 * fields are filled with safe defaults — notably a `'#'` link, which
 * {@link detectStorySource} then classifies as internal.
 *
 * @param stories - Items as returned by the Available Stories list.
 * @returns The stories in view-model form, in the order supplied.
 */
const formatStories = (stories: IStoryItem[]): IStory[] =>
  stories.map((story, index) => ({
    id: `story-${story.id || index}`,
    title: story.title,
    description: story.description || '',
    imageUrl: story.imageUrl,
    linkToPost: story.linkToPost || '#',
    date: story.date || new Date().toLocaleDateString(),
    created: story.created,
    source: story.source || '',
    author: story.author || undefined,
  }));

/** Props for an available-story list item row. */
interface AvailableStoryItemProps {
  /** Story to render. */
  story: IStory;
  /** Called with the story id when edit is pressed. */
  onEdit: (storyId: string) => void;
}

/**
 * One row in the Available Stories list.
 *
 * Kept at module scope so React preserves row identity across parent renders,
 * which avoids repeatedly tearing down and re-registering each dnd-kit item.
 */
const AvailableStoryItem: React.FC<AvailableStoryItemProps> = ({ story, onEdit }) => {
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
  const rawSource = (story.source || source || '').toLowerCase();
  const isLinkedin = rawSource === 'linkedin';
  const isExternal = rawSource === 'external';

  const displayDomain = domain ? domain.replace(/^www\./, '') : '';
  const badgeLabel = isLinkedin ? 'LinkedIn' : (displayDomain || 'External');

  return (
    <div
      ref={setNodeRef}
      style={style}
      {...attributes}
      className={`${styles.storyCard} ${isDragging ? styles.dragging : ''}`}
    >
      <div className={styles.storyDragArea} {...listeners}>
        <img src={story.imageUrl} alt={story.title} className={styles.storyImage} />
        <div className={styles.storyContent}>
          <div className={styles.storyTitleRow}>
            <h4>{story.title}</h4>
          </div>
          <div className={styles.storyMeta}>
            <span className={styles.storyDate}>{formatPublishDate(story.date)}</span>
            {(isLinkedin || isExternal) && (
              <span
                className={`${styles.rowBadge} ${isLinkedin ? styles.rowBadgeLinkedin : styles.rowBadgeExternal}`}
              >
                {/* 12px, not the 10 the reference's `w-2.5` suggests: Badge's
                    own `[&>svg]:size-3` compiles to `.badge > svg` (0,1,1),
                    which outranks `.w-2\.5` (0,1,0) — so the icon renders at
                    `size-3` and the `w-2.5` never applies. */}
                {isLinkedin ? <LinkedinIcon size={12} /> : <ExternalLinkIcon size={12} />}
                {badgeLabel}
              </span>
            )}
          </div>
        </div>
      </div>
      <div className={styles.storyActions}>
        <button
          type="button"
          className={styles.actionBtn}
          onClick={() => onEdit(story.id)}
          title="Edit story"
          aria-label="Edit story"
        >
          <EditIcon />
        </button>
      </div>
    </div>
  );
};




/** Props for {@link StoryDragDrop}. */
interface StoryDragDropProps {
  /**
   * SharePoint context used to authenticate list calls. When absent the board
   * renders an error instead of loading, which is how tests exercise the
   * failure path.
   */
  context?: IWebPartContext;
}

/**
 * The storyboard editor: the component that owns all board state.
 *
 * Responsibilities:
 *
 * - loads stories, the live board and scheduled groups from SharePoint;
 * - hosts the drag-and-drop context that moves stories between the Available
 *   Stories list and the board's slots;
 * - publishes the board live, and schedules arrangements for future dates;
 * - runs two modal editing modes — scheduling and group editing — each of
 *   which parks the current board and restores it on exit;
 * - owns the toast queue every action reports through.
 */
const StoryDragDrop: React.FC<StoryDragDropProps> = ({ context }) => {
  const [availableStories, setAvailableStories] = useState<IStory[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedLayout, setSelectedLayout] = useState<LayoutType>('connectHomepage');
  const [slotStories, setSlotStories] = useState<SlotStoryMap>({});
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [editingStory, setEditingStory] = useState<IStory | null>(null);
  const [isSchedulingMode, setIsSchedulingMode] = useState(false);
  const [scheduledGroups, setScheduledGroups] = useState<ScheduledStoryGroup[]>([]);
  const [selectedScheduleDate, setSelectedScheduleDate] = useState<string>('');
  const [selectedScheduleTime, setSelectedScheduleTime] = useState<string>('09:00');
  const [showCalendarOverlay, setShowCalendarOverlay] = useState(false);
  const [slotLayoutPreferences, setSlotLayoutPreferences] = useState<Record<string, CardLayout>>({});
  const [isGroupEditingMode, setIsGroupEditingMode] = useState(false);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);
  const [originalBoardState, setOriginalBoardState] = useState<SlotStoryMap>({});
  const [rescheduleDate, setRescheduleDate] = useState<string>('');
  const [rescheduleTime, setRescheduleTime] = useState<string>('09:00');
  const [showRescheduleCalendar, setShowRescheduleCalendar] = useState(false);
  const [toasts, setToasts] = useState<IToast[]>([]);
  const [groupToDelete, setGroupToDelete] = useState<ScheduledStoryGroup | null>(null);


  /**
   * Queues a notification.
   *
   * @param title - Bold headline.
   * @param type - Severity. Defaults to `success`.
   * @param description - Optional supporting line.
   * @param duration - Auto-dismiss delay in ms. Defaults to 3000, matching the
   * reference; pass 4000 for publish and schedule confirmations.
   */
  const showToast = (
    title: string,
    type: 'success' | 'error' | 'info' = 'success',
    description?: string,
    duration = 3000
  ): void => {
    const id = `toast-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
    setToasts(prev => [...prev, { id, type, title, description, duration }]);
  };

  /**
   * Removes a notification from the queue.
   *
   * @param id - Id of the toast to drop.
   */
  const handleCloseToast = (id: string): void => {
    setToasts(prev => prev.filter(t => t.id !== id));
  };

  // Load stories and scheduled groups once the SharePoint context is known.
  // The `isMounted` guard keeps a late response from setting state after
  // unmount, which React would warn about.
  useEffect(() => {
    let isMounted = true;

    /**
     * Fetches everything the board needs for its first render.
     *
     * Story and scheduled-group loads are isolated from each other: a failure
     * loading groups is logged but leaves the stories usable. Either way the
     * loading flag clears, so the skeleton is always replaced.
     */
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

    return () => {
      isMounted = false;
    };
  }, [context]);

  const sensors = useSensors(
    useSensor(MouseSensor, {
      activationConstraint: {
        distance: 8,
      },
    }),
    useSensor(TouchSensor, {
      activationConstraint: {
        delay: 200,
        tolerance: 8,
      },
    })
  );

  /**
   * Resolves a completed drag into a board update.
   *
   * Four outcomes, by where the drag started and ended:
   *
   * - Available list to a slot: the story moves in, and any story it displaces
   *   returns to the available list.
   * - Slot to the available list: the slot is emptied and the story returns.
   * - Slot to slot: the two slots swap, or the story moves if the target is
   *   empty.
   * - Anything unrecognised — no drop target, an unknown story, or a drop back
   *   onto the available list from the available list — is ignored.
   *
   * Drops are accepted both on a slot itself and on the card occupying it,
   * since dnd-kit reports whichever is under the pointer.
   *
   * @param event - The dnd-kit drag-end event.
   */
  const handleDragEnd = (event: DragEndEvent): void => {
    const { active, over } = event;
    if (!over) return;

    const activeStoryId = active.id as string;
    const overId = over.id as string;
    const fromAvailable =
      active.data.current?.sortable?.containerId === 'available-stories';

    const slotIds = getLayoutConfig(selectedLayout).slots.map(s => s.id);

    /**
     * Resolves a drop target id to a slot id.
     *
     * @param id - Id reported by dnd-kit: either a slot, or a story sitting in one.
     * @returns The slot id, or `undefined` when the target is neither.
     */
    const resolveDestSlot = (id: string): string | undefined => {
      if (slotIds.indexOf(id) !== -1) return id;
      return Object.keys(slotStories).find(
        key => slotStories[key]?.id === id
      );
    };

    if (fromAvailable) {
      const story = availableStories.find(s => s.id === activeStoryId);
      if (!story) return;
      if (overId === 'available-stories') return;

      const destSlot = resolveDestSlot(overId);
      if (!destSlot) return;

      const occupant = slotStories[destSlot];
      setSlotStories(prev => ({ ...prev, [destSlot]: story }));
      setAvailableStories(prev => {
        let next = prev.filter(s => s.id !== story.id);
        if (occupant && !next.some(s => s.id === occupant.id)) {
          next = [...next, occupant];
        }
        return next;
      });
      return;
    }

    const sourceSlot = Object.keys(slotStories).find(
      key => slotStories[key]?.id === activeStoryId
    );
    if (!sourceSlot) return;
    const sourceStory = slotStories[sourceSlot];

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

  /**
   * Creates a story in SharePoint and refreshes the available list.
   *
   * The list is re-read rather than patched locally so the new row carries its
   * server-assigned id and created date. The confirmation names the detected
   * source and domain, matching the reference.
   *
   * @param story - Field values from the add dialog.
   */
  const handleAddStory = (story: Omit<IStory, 'id' | 'date'>): void => {
    createStory({
      title: story.title,
      description: story.description,
      imageUrl: story.imageUrl,
      linkToPost: story.linkToPost,
      source: story.source
    }).then(() => {
      return getStories();
    }).then((stories) => {
      const formattedStories = formatStories(stories);
      setAvailableStories(formattedStories);
      setIsModalOpen(false);

      const { source, domain } = detectStorySource(story.linkToPost);
      const sourceLabel = source === 'linkedin' ? 'LinkedIn' : source === 'external' ? 'External' : 'Internal';
      showToast(
        'Story Added Successfully!',
        'success',
        `New ${sourceLabel} story added${domain ? ` from ${domain}` : ''}.`
      );
    }).catch((err) => {
      console.error('Failed to save or reload stories:', err);
      setIsModalOpen(false);
      showToast('Failed to add story', 'error', err.message || 'Unknown error');
    });
  };

  /**
   * Opens the edit dialog for a story.
   *
   * @param storyId - Id of the story to edit. Unknown ids are ignored.
   */
  const handleEditStory = (storyId: string): void => {
    const story = availableStories.find(s => s.id === storyId);
    if (story) {
      setEditingStory(story);
    }
  };

  /**
   * Saves edits to a story and reflects them everywhere it appears.
   *
   * After the write the list is re-read, and any copy of the story already
   * placed on the board is swapped for the refreshed one so the two never
   * drift apart.
   *
   * @param updatedStory - The story with edits applied.
   */
  const handleUpdateStory = (updatedStory: IStory): void => {
    const storyIdMatch = updatedStory.id.match(/\d+/);
    const numericId = storyIdMatch ? parseInt(storyIdMatch[0], 10) : null;

    if (!numericId) {
      console.error('Invalid story ID format:', updatedStory.id);
      showToast('Error: Invalid story ID', 'error');
      return;
    }

    updateStory(numericId, {
      title: updatedStory.title,
      description: updatedStory.description,
      imageUrl: updatedStory.imageUrl,
      linkToPost: updatedStory.linkToPost,
      source: updatedStory.source
    }).then(() => {
      return getStories();
    }).then((stories) => {
      const formattedStories = formatStories(stories);
      setAvailableStories(formattedStories);
      
      const newSlotStories = { ...slotStories };
      Object.keys(newSlotStories).forEach(key => {
        const updatedSlotStory = formattedStories.find(s => s.id === newSlotStories[key]?.id);
        if (updatedSlotStory) {
          newSlotStories[key] = updatedSlotStory;
        }
      });
      setSlotStories(newSlotStories);
      setEditingStory(null);

      const { source, domain } = detectStorySource(updatedStory.linkToPost);
      const sourceLabel = source === 'linkedin' ? 'LinkedIn' : source === 'external' ? 'External' : 'Internal';
      showToast(
        'Story Updated Successfully!',
        'success',
        `Story updated with ${sourceLabel} source${domain ? ` (${domain})` : ''}.`
      );
    }).catch((err) => {
      console.error('Failed to update story:', err);
      showToast('Failed to update story', 'error', err.message || 'Unknown error');
    });
  };

  /**
   * Deletes a story and removes it from the board if it was placed.
   *
   * @param storyId - Id of the story to delete, in `story-{listItemId}` form.
   */
  const handleDeleteStory = (storyId: string): void => {
    const storyIdMatch = storyId.match(/\d+/);
    const numericId = storyIdMatch ? parseInt(storyIdMatch[0], 10) : null;

    if (!numericId) {
      console.error('Invalid story ID format:', storyId);
      showToast('Error: Invalid story ID', 'error');
      return;
    }

    deleteStory(numericId).then(() => {
      return getStories();
    }).then((stories) => {
      setAvailableStories(formatStories(stories));
      
      const newSlotStories = { ...slotStories };
      Object.keys(newSlotStories).forEach(key => {
        if (newSlotStories[key]?.id === storyId) {
          delete newSlotStories[key];
        }
      });
      setSlotStories(newSlotStories);
      setEditingStory(null);
      // Reference: `toast.success("Story deleted successfully!")` passes no
      // duration, so sonner's own 4000ms default applies rather than the 3000
      // the explicit calls use.
      showToast('Story deleted successfully!', 'success', undefined, 4000);
    }).catch((err) => {
      console.error('Failed to delete story:', err);
      showToast('Failed to delete story', 'error', err.message || 'Unknown error');
    });
  };

  /**
   * Clears one slot, returning its story to the available list.
   *
   * @param slotId - Slot to empty. Empty slots are ignored.
   */
  const handleRemoveFromSlot = (slotId: string): void => {
    const story = slotStories[slotId];
    if (story) {
      setAvailableStories([...availableStories, story]);
      
      const newSlotStories = { ...slotStories };
      delete newSlotStories[slotId];
      setSlotStories(newSlotStories);
    }
  };

  /**
   * Empties every slot, returning all placed stories to the available list.
   *
   * Deliberately silent — the reference's Clear Board raises no toast, and
   * Reset Board reports its own outcome.
   */
  const handleClearBoard = (): void => {
    const storiesInSlots = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== null);
    setAvailableStories([...availableStories, ...storiesInSlots]);
    setSlotStories({});
    showToast('Board Cleared', 'info', 'All stories moved to available stories.');
  };

  /**
   * Restores the board to the currently published state.
   *
   * The live snapshot supplies the slots, the per-slot layouts and the layout
   * preset. The available list is rebuilt from every story currently known,
   * minus those the snapshot places, so a restored story never appears both on
   * the board and in the list. With nothing published the board is cleared
   * instead.
   */
  const handleResetBoard = (): void => {
    getLiveBoard()
      .then((published) => {
        if (!published) {
          handleClearBoard();
          return;
        }

        if (published.layoutType !== 'connectHomepage') {
          showToast('Layout Type Does Not Exist', 'error', `Published layout type "${published.layoutType}" is not supported.`);
          return;
        }

        const restoredSlots: SlotStoryMap = { ...published.slotStories };
        const usedIds = new Set(
          Object.keys(restoredSlots)
            .map(k => restoredSlots[k])
            .filter((s): s is IStory => !!s)
            .map(s => s.id)
        );

        const universe: IStory[] = [
          ...availableStories,
          ...Object.keys(slotStories)
            .map(k => slotStories[k])
            .filter((s): s is IStory => !!s),
        ];
        const seen = new Set<string>();
        const nextAvailable = universe.filter(s => {
          if (usedIds.has(s.id) || seen.has(s.id)) return false;
          seen.add(s.id);
          return true;
        });

        setSlotStories(restoredSlots);
        setAvailableStories(nextAvailable);
        if (published.slotLayoutPreferences) {
          setSlotLayoutPreferences({ ...published.slotLayoutPreferences });
        }
        setSelectedLayout(published.layoutType);

        const restoredCount = Object.keys(restoredSlots)
          .map(k => restoredSlots[k])
          .filter((s): s is IStory => !!s).length;
        showToast(
          'Board Reset to Last Published State',
          'success',
          `Restored ${restoredCount} stories to their last published positions.`
        );
      })
      .catch((err) => {
        console.error('Reset Board failed:', err);
        showToast('Reset failed', 'error', 'Could not restore the published board.');
      });
  };

  /**
   * Switches the board to another layout preset.
   *
   * Placed stories keep their slots; slots the new preset does not define stop
   * rendering but are not discarded.
   *
   * @param value - Id of the preset to switch to.
   */
  const handleLayoutChange = (value: string): void => {
    setSelectedLayout(value as LayoutType);
    showToast('Layout Changed', 'info', `Switched to ${getLayoutConfig(value as LayoutType).name}. Your stories remain in place.`);
  };


  /**
   * Overrides how one slot renders its card.
   *
   * @param slotId - Slot to change.
   * @param mode - Card layout to apply.
   */
  const handleSlotLayoutChange = (
    slotId: string,
    mode: CardLayout
  ): void => {
    setSlotLayoutPreferences(prev => ({ ...prev, [slotId]: mode }));
  };

  /**
   * Enters scheduling mode.
   *
   * Parks the current board so it can be restored on exit, then clears the
   * slots to give the new schedule a blank arrangement to build.
   */
  const handleScheduleGroup = (): void => {
    setOriginalBoardState({ ...slotStories });

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

  /** Leaves scheduling mode and restores the board parked on entry. */
  const handleExitScheduling = (): void => {
    setSlotStories({ ...originalBoardState });
    setOriginalBoardState({});

    setIsSchedulingMode(false);
    setShowCalendarOverlay(false);
    setSelectedScheduleDate('');
    setSelectedScheduleTime('09:00');
  };

  /** Opens the calendar overlay used to pick the scheduled date and time. */
  const handleOpenDateTimePicker = (): void => {
    setShowCalendarOverlay(true);
  };

  /**
   * Saves the current arrangement as a scheduled group.
   *
   * Rejects the attempt, with an explanatory toast, when no date is chosen,
   * the board is empty, the group cap is reached, or that date already has a
   * group. On success the saved record — which carries the SharePoint item id —
   * is added to state and scheduling mode exits.
   */
  const handleScheduleNow = (): void => {
    if (!selectedScheduleDate) {
      showToast('No Date Selected', 'error', 'Please select a date from the calendar before scheduling.');
      return;
    }

    const boardStories = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== undefined);

    if (boardStories.length === 0) {
      showToast('No Stories to Schedule', 'error', 'Please add stories to the main board before scheduling.');
      return;
    }

    if (scheduledGroups.length >= MAX_SCHEDULED_GROUPS) {
      showToast('Maximum Groups Reached', 'error', `You can only have up to ${MAX_SCHEDULED_GROUPS} scheduled groups. Please remove a group first.`);
      return;
    }

    const scheduleDate = fromDateKey(selectedScheduleDate);
    const dateTaken = scheduledGroups.some(
      g => g.date.toDateString() === scheduleDate.toDateString()
    );
    if (dateTaken) {
      showToast('Date Already Scheduled', 'error', "There's already a scheduled group for this date. Please choose a different date.");
      return;
    }

    const layoutConfig = getLayoutConfig(selectedLayout);
    const scheduledGroup: ScheduledStoryGroup = {
      id: `scheduled-${Date.now()}`,
      date: scheduleDate,
      time: selectedScheduleTime,
      slotStories: { ...slotStories },
      slotLayoutPreferences: { ...slotLayoutPreferences },
      layoutType: selectedLayout,
      layoutName: layoutConfig.name,
      createdAt: new Date()
    };

    createScheduledGroup(scheduledGroup)
      .then((saved) => {
        setScheduledGroups(prev => [...prev, saved]);
        handleExitScheduling();

        const formattedDate = scheduleDate.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric',
          year: 'numeric'
        });
        showToast('Board Scheduled Successfully!', 'success', `${boardStories.length} stories scheduled for ${formattedDate}.`, 4000);
      })
      .catch((err) => {
        console.error('Failed to save scheduled group:', err);
        showToast('Scheduling Failed', 'error', 'Could not save the scheduled group to SharePoint. Please try again.');
      });
  };

  /**
   * Asks for confirmation before removing a scheduled group.
   *
   * @param groupId - Group to remove. Unknown ids are ignored.
   */
  const handleDeleteScheduledGroup = (groupId: string): void => {
    const groupToRemove = scheduledGroups.find(g => g.id === groupId);
    if (groupToRemove) {
      setGroupToDelete(groupToRemove);
    }
  };

  /**
   * Removes a scheduled group once deletion is confirmed.
   *
   * Its stories return to the available list, skipping any already present so
   * the list cannot gain duplicates.
   *
   * @param groupId - Group to remove.
   */
  const handleDeleteScheduledGroupConfirmed = (groupId: string): void => {
    const groupToRemove = scheduledGroups.find(g => g.id === groupId);
    if (!groupToRemove) return;

    deleteScheduledGroupByGroupId(groupId)
      .then(() => {
        const groupStories = Object.keys(groupToRemove.slotStories)
          .map(key => groupToRemove.slotStories[key])
          .filter((s): s is IStory => s !== undefined);
        
        setAvailableStories(prev => {
          const existingIds = new Set(prev.map(s => s.id));
          const toAdd = groupStories.filter(s => !existingIds.has(s.id));
          return [...toAdd, ...prev];
        });

        setScheduledGroups(prev => prev.filter(g => g.id !== groupId));
        setGroupToDelete(null);

        const formattedDate = groupToRemove.date.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric'
        });
        showToast('Scheduled Group Removed', 'success', `${groupStories.length} stories from ${formattedDate} returned to available list.`);
      })
      .catch((err) => {
        console.error('Failed to delete scheduled group:', err);
        showToast('Delete Failed', 'error', 'Could not remove the scheduled group from SharePoint.');
        setGroupToDelete(null);
      });
  };

  /**
   * Leaves group editing mode.
   *
   * Returns the group's stories to the available list and restores the board
   * parked when editing began. Shared by both the save and cancel paths.
   */
  const handleExitGroupEditingMode = (): void => {
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
    
    setSlotStories({ ...originalBoardState });
    
    setIsGroupEditingMode(false);
    setEditingGroupId(null);
    setOriginalBoardState({});
    setShowRescheduleCalendar(false);
  };

  /**
   * Opens a scheduled group for rearrangement.
   *
   * When the group was built for a different layout preset the editor is asked
   * whether to switch; declining aborts. The current board is parked, the
   * group's arrangement is loaded in its place, and the group's own stories
   * are withheld from the available list so they cannot be placed twice.
   *
   * @param group - Group to edit.
   */
  const handleEnterGroupEditingMode = (group: ScheduledStoryGroup): void => {
    if (group.layoutType !== selectedLayout) {
      const confirmSwitch = window.confirm(
        `This group was created for "${group.layoutName}" layout, but you're currently viewing "${getLayoutConfig(selectedLayout).name}". ` +
        `Do you want to switch to the "${group.layoutName}" layout?`
      );
      
      if (confirmSwitch) {
        setSelectedLayout(group.layoutType);
        showToast('Layout Switched', 'info', `Switched layout to ${group.layoutName}`);
      } else {
        return;
      }
    }

    setOriginalBoardState({ ...slotStories });

    setSlotStories({ ...group.slotStories });
    if (group.slotLayoutPreferences) {
      setSlotLayoutPreferences({ ...group.slotLayoutPreferences });
    } else {
      setSlotLayoutPreferences({});
    }
    
    setIsGroupEditingMode(true);
    setEditingGroupId(group.id);
    
    const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
    const y = group.date.getFullYear();
    const m = pad2(group.date.getMonth() + 1);
    const d = pad2(group.date.getDate());
    setRescheduleDate(`${y}-${m}-${d}`);
    setRescheduleTime(group.time);

    const scheduledStoryIds = Object.keys(group.slotStories)
      .map(key => group.slotStories[key])
      .filter((story): story is IStory => story !== undefined)
      .map(story => story.id);
    
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

    const editFormattedDate = group.date.toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'short',
      day: 'numeric'
    });
    showToast(
      'Group Editing Mode Active',
      'info',
      `Editing ${scheduledStoryIds.length} stories for ${editFormattedDate}. Rearrange on the main board.`
    );
  };

  /**
   * Saves the arrangement and date currently being edited back to the group.
   *
   * Persists to SharePoint when the group has an item id, and updates state
   * only when it does not. The date handed to the service is the same
   * locally-built `Date` shown in the UI, so the two can never disagree about
   * which calendar day was chosen.
   */
  const handleSaveGroupEdits = (): void => {
    if (!editingGroupId) return;
    const group = scheduledGroups.find(g => g.id === editingGroupId);
    if (!group) return;

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
          ? { ...g, slotStories: newSlotStories, slotLayoutPreferences: { ...slotLayoutPreferences }, date: newDate, time: rescheduleTime }
          : g
      ));
      handleExitGroupEditingMode();
      const savedCount = Object.keys(newSlotStories)
        .map(k => newSlotStories[k])
        .filter((s): s is IStory => !!s).length;
      showToast('Group Updated Successfully!', 'success', `Saved ${savedCount} stories with new arrangement.`);
    };

    if (typeof group.spId === 'number') {
      // `newDate` is already built from local parts, so the service and the
      // displayed state agree on the calendar day. Re-parsing the picker string
      // as UTC here is what used to make them disagree by up to a day.
      rescheduleGroup(group.spId, newDate, rescheduleTime, newSlotStories, slotLayoutPreferences)
        .then(applyLocal)
        .catch((err) => {
          console.error('Failed to reschedule group:', err);
          showToast('Update Failed', 'error', 'Could not update the scheduled group in SharePoint.');
        });
    } else {
      applyLocal();
    }
  };

  /** Abandons group edits and restores the board parked on entry. */
  const handleCancelGroupEdits = (): void => {
    if (!editingGroupId) return;
    handleExitGroupEditingMode();
    showToast('Group Editing Cancelled', 'info', 'No changes were saved. Returned to normal mode.');
  };

  /**
   * Accepts a new date and time for the group being edited.
   *
   * Only closes the picker and confirms the choice — the change reaches
   * SharePoint when the group edits are saved.
   */
  const handleConfirmReschedule = (): void => {
    setShowRescheduleCalendar(false);
    const parts = rescheduleDate.split('-');
    const rescheduledDate = rescheduleDate
      ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
      : new Date();
    const formattedDate = rescheduledDate.toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
    showToast('Group Rescheduled Successfully!', 'success', `Group has been rescheduled to ${formattedDate}.`);
  };

  /**
   * Publishes the board live.
   *
   * Supersedes the previous live board and writes an immutable snapshot of the
   * slots, their per-slot layouts and the layout preset. An empty board is
   * rejected with a toast rather than published.
   */
  const handlePublishBoard = (): void => {
    const boardStories = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== undefined);

    if (boardStories.length === 0) {
      showToast(
        'Cannot Publish Empty Board',
        'error',
        'Please add at least one story to the board before publishing. Drag stories from the Available Stories section to any slot.',
        4000
      );
      return;
    }

    const layoutConfig = getLayoutConfig(selectedLayout);
    publishBoard({
      layoutType: selectedLayout,
      layoutName: layoutConfig.name,
      slotStories: { ...slotStories },
      slotLayoutPreferences
    })
      .then(() => {
        showToast(
          'Storyboard Published!',
          'success',
          `Successfully published ${boardStories.length} stories to SharePoint. Board layout and content have been saved.`,
          4000
        );
      })
      .catch((err) => {
        console.error('Failed to publish board:', err);
        showToast('Publish Failed', 'error', 'Could not publish the board to SharePoint. Please try again.');
      });
  };

  const filteredStories = availableStories.filter(story =>
    story.title.toLowerCase().includes(searchTerm.toLowerCase())
  );

  if (isLoading) {
    return (
      <div className={styles.storyDragDrop}>
        <BoardSkeleton />
      </div>
    );
  }

  return (
    <div className={styles.storyDragDrop}>
      {error && (
        <div style={{ padding: '20px', backgroundColor: '#fee', color: '#c00', borderRadius: '4px', marginBottom: '20px' }}>
          {error}
        </div>
      )}
      <div className={styles.header}>
        <div className={styles.headerLeft}>
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
            onClick={handleResetBoard}
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
            />
          </div>

          <div className={styles.bottomSection}>
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
                        className={`${styles.schedulingBtnSmall} ${styles.datePickerBtn}`}
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
                      <span className={styles.labelFull}>✓ Save Group Edits</span>
                      <span className={styles.labelShort}>✓ Save</span>
                    </button>
                    <button
                      className={styles.schedulingBtnSmall}
                      onClick={handleCancelGroupEdits}
                    >
                      <span className={styles.labelFull}>✕ Cancel Edits</span>
                      <span className={styles.labelShort}>✕ Cancel</span>
                    </button>
                  </div>
                ) : !isSchedulingMode ? (
                  scheduledGroups.length < MAX_SCHEDULED_GROUPS && (
                    <button
                      className={styles.btnSchedule}
                      onClick={handleScheduleGroup}
                      disabled={availableStories.length === 0}
                    >
                      <Icon iconName="Add" />
                      <span className={styles.labelFull}>Schedule New Group</span>
                      <span className={styles.labelShort}>Schedule Group</span>
                    </button>
                  )
                ) : (
                  <div className={styles.schedulingControlsInline}>
                    <div className={styles.schedulingBadge}>
                      📅 Scheduling Mode
                    </div>
                    <div className={styles.buttonWithDropdown}>
                      <button
                        className={`${styles.schedulingBtnSmall} ${styles.datePickerBtn}`}
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
                      <span className={styles.labelFull}>Exit Scheduling</span>
                      <span className={styles.labelShort}>Exit</span>
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
                                {isEditingThis && <span className={styles.editingLabel}>(Editing on Main Board)</span>}
                              </p>
                            </div>
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
                            <p className={styles.emptyHint}>Try a different search term</p>
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
                          onEdit={handleEditStory}
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
          showToast={showToast}
        />
      )}

      {editingStory && (
        <EditStoryModal
          story={editingStory}
          onClose={() => setEditingStory(null)}
          onUpdate={handleUpdateStory}
          onDelete={handleDeleteStory}
          showToast={showToast}
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
