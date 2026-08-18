import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
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
import AlertDialog from './AlertDialog';
import DateTimePicker from './DateTimePicker';
import { IStory, LayoutType, SlotStoryMap, ScheduledStoryGroup, CardLayout } from './types';
import {
  getStories,
  createStory,
  updateStory,
  deleteStory,
  initializeSharePoint,
  setAvailableStoriesListName,
  IStoryItem
} from '../services/availableStoriesService';
import { LinkedinIcon, ExternalLinkIcon, EditIcon } from './icons';
import BoardSkeleton from './BoardSkeleton';
import {
  createScheduledGroup,
  getScheduledGroups,
  deleteScheduledGroup,
  deleteScheduledGroupByGroupId,
  rescheduleGroup,
  setScheduledStoriesListName,
  fromDateKey,
  toDateKey,
  toScheduledMoment,
} from '../services/scheduleStoriesService';
import { publishBoard, getLiveBoard, setPublishedStoriesListName } from '../services/publishStoriesService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import LayoutRenderer from './layouts/LayoutRenderer';
import { AVAILABLE_LAYOUTS, getLayoutConfig } from './layouts/layoutConfig';
import { detectStorySource, formatPublishDate } from './cards/storyHelpers';
import { ToastContainer, IToast } from './Toast';

/** Most scheduled groups that may exist at once, matching the reference cap. */
const MAX_SCHEDULED_GROUPS = 10;

const LOCAL_DRAFT_VERSION = 1;

interface IBoardDraft {
  version: number;
  layout: LayoutType;
  slotStories: SlotStoryMap;
  slotLayoutPreferences: Record<string, CardLayout>;
}

/**
 * Whether a calendar day is behind today.
 *
 * Day-granular on purpose: cells arrive as local midnight, and comparing them
 * against today's midnight rather than the current instant keeps *today*
 * selectable, so a board can still be scheduled for later the same day. The
 * hour is then policed separately by {@link isMomentPast}.
 *
 * @param date - Local midnight of the day being tested.
 * @returns `true` when the day must be disabled.
 */
const isDayPast = (date: Date): boolean => {
  const now = new Date();
  return date < new Date(now.getFullYear(), now.getMonth(), now.getDate());
};

/**
 * Whether a chosen day and time is already behind us.
 *
 * The companion to {@link isDayPast}: with today selectable, this is what stops
 * a group being scheduled into an hour that has already passed.
 *
 * @param date - Day being scheduled; only its date parts are read.
 * @param time - Time of day as `HH:mm`. Unparseable parts count as zero.
 * @returns `true` when the moment is now or earlier.
 */
const isMomentPast = (date: Date, time: string): boolean => {
  const parts = time.split(':');
  const hr = parseInt(parts[0], 10);
  const min = parseInt(parts[1], 10);
  const moment = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    isNaN(hr) ? 0 : hr,
    isNaN(min) ? 0 : min
  );

  return moment.getTime() <= new Date().getTime();
};

/**
 * Whether a scheduled group's moment has already gone by.
 *
 * Prefers `scheduledAtUtc`, the absolute instant the service persists, so the
 * verdict is the same for every viewer regardless of timezone. Groups still
 * held in memory before their first save have no instant yet, so those fall
 * back to the civil day and time the picker produced.
 *
 * @param group - Group to test.
 * @returns `true` when the group is due or overdue.
 */
const isGroupPast = (group: ScheduledStoryGroup): boolean => {
  if (group.scheduledAtUtc) {
    return group.scheduledAtUtc.getTime() <= new Date().getTime();
  }

  return isMomentPast(group.date, group.time);
};

/**
 * Orders scheduled groups soonest first, the order the service reads them in.
 *
 * Sorting the list rather than appending to it keeps a group created or
 * rescheduled mid-session in its rightful place instead of at the end.
 *
 * The day comes from {@link toDateKey} rather than the timestamp because the
 * two writers disagree about `date`: creating a group stores local midnight
 * and keeps the hour in `time`, while saving an edit folds the hour into the
 * `Date`. Reading local parts and pairing them with `time` sorts both alike.
 */
const byScheduledMoment = (a: ScheduledStoryGroup, b: ScheduledStoryGroup): number => {
  // Instants when both carry one: exact, and immune to the hour that repeats
  // itself when local clocks go back. The civil comparison below is the
  // fallback for a group not yet round-tripped through the service.
  if (a.scheduledAtUtc && b.scheduledAtUtc) {
    return a.scheduledAtUtc.getTime() - b.scheduledAtUtc.getTime();
  }

  return `${toDateKey(a.date)}T${a.time}`.localeCompare(`${toDateKey(b.date)}T${b.time}`);
};

/**
 * Whether a calendar day cannot be rescheduled onto: it has already begun, or
 * another group already holds it.
 *
 * Only the reschedule calendar carries the second rule, which is the split the
 * reference draws. Scheduling a *new* group leaves taken days enabled and
 * rejects them on confirm instead, where {@link StoryDragDrop}'s
 * `handleScheduleNow` raises the same "Date Already Scheduled" toast the
 * reference does; rescheduling has no confirm-time check, so its calendar is
 * the only thing standing between two groups and the same day.
 *
 * Groups are compared by calendar day, so their time of day does not enter
 * into it.
 *
 * @param date - Local midnight of the day being tested.
 * @param groups - Every scheduled group currently known.
 * @param exceptGroupId - Group to ignore, so the one being rescheduled does
 * not block its own date.
 * @returns `true` when the day must be disabled.
 */
const isDayBlocked = (
  date: Date,
  groups: ScheduledStoryGroup[],
  exceptGroupId?: string | null
): boolean => {
  if (isDayPast(date)) return true;

  return groups.some(
    g => g.id !== exceptGroupId && g.date.toDateString() === date.toDateString()
  );
};

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

/**
 * Drops the stories that are currently sitting in board slots.
 *
 * The re-read that follows a write returns the whole list, board occupants
 * included. Without this they reappear under Available while still on the
 * board, where they can be dragged into a second slot and published twice.
 *
 * @param stories - Freshly fetched list.
 * @param slots - Slot map whose occupants should be held back.
 */
const excludePlacedStories = (stories: IStory[], slots: SlotStoryMap): IStory[] => {
  const placedIds = new Set(
    Object.keys(slots)
      .map((slotId) => slots[slotId]?.id)
      .filter((id): id is string => !!id)
  );

  return stories.filter((story) => !placedIds.has(story.id));
};

const mergeSlotStoriesWithLatest = (
  slots: SlotStoryMap,
  latestStories: IStory[]
): SlotStoryMap => {
  const byId = new Map(latestStories.map((story) => [story.id, story]));
  const merged: SlotStoryMap = {};

  Object.keys(slots).forEach((slotId) => {
    const story = slots[slotId];
    if (!story) {
      merged[slotId] = story;
      return;
    }

    merged[slotId] = byId.get(story.id) || story;
  });

  return merged;
};

const loadDraft = (storageKey: string): IBoardDraft | null => {
  try {
    if (typeof window === 'undefined' || !window.localStorage) return null;
    const raw = window.localStorage.getItem(storageKey);
    if (!raw) return null;

    const parsed = JSON.parse(raw) as Partial<IBoardDraft>;
    if (!parsed || parsed.version !== LOCAL_DRAFT_VERSION) return null;
    if (!parsed.layout || !parsed.slotStories || !parsed.slotLayoutPreferences) return null;

    return {
      version: LOCAL_DRAFT_VERSION,
      layout: parsed.layout,
      slotStories: parsed.slotStories,
      slotLayoutPreferences: parsed.slotLayoutPreferences,
    };
  } catch {
    return null;
  }
};

const saveDraft = (storageKey: string, draft: IBoardDraft): void => {
  try {
    if (typeof window === 'undefined' || !window.localStorage) return;
    window.localStorage.setItem(storageKey, JSON.stringify(draft));
  } catch {
    // Ignore storage write failures.
  }
};

const clearDraft = (storageKey: string): void => {
  try {
    if (typeof window === 'undefined' || !window.localStorage) return;
    window.localStorage.removeItem(storageKey);
  } catch {
    // Ignore storage cleanup failures.
  }
};

const toDateOnlyTimestamp = (value?: string): number => {
  if (!value) return Number.NEGATIVE_INFINITY;

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return Number.NEGATIVE_INFINITY;

  // Compare by UTC calendar date only, ignoring time-of-day.
  return Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate());
};

const sortAvailableStories = (stories: IStory[]): IStory[] =>
  [...stories].sort((a, b) => {
    const modifiedDelta = toDateOnlyTimestamp(b.date) - toDateOnlyTimestamp(a.date);
    if (modifiedDelta !== 0) return modifiedDelta;
    return a.title.localeCompare(b.title, undefined, { sensitivity: 'base' });
  });

const serializeBoardStateForDirtyCheck = (
  layout: LayoutType,
  slots: SlotStoryMap,
  slotPrefs: Record<string, CardLayout>,
  available: IStory[]
): string => {
  const slotEntries = Object.keys(slots)
    .sort()
    .map((slotId) => `${slotId}:${slots[slotId]?.id || ''}`);

  const prefEntries = Object.keys(slotPrefs)
    .sort()
    .map((slotId) => `${slotId}:${slotPrefs[slotId]}`);

  const availableIds = Array.from(new Set(available.map((story) => story.id))).sort();

  return JSON.stringify({
    layout,
    slotEntries,
    prefEntries,
    availableIds,
  });
};

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
  /** SharePoint list title for the available stories source list. */
  availableStoriesListName?: string;
  /** SharePoint list title for scheduled board snapshots. */
  scheduledStoriesListName?: string;
  /** SharePoint list title for published board snapshots. */
  publishedStoriesListName?: string;
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
const StoryDragDrop: React.FC<StoryDragDropProps> = ({
  context,
  availableStoriesListName,
  scheduledStoriesListName,
  publishedStoriesListName,
}) => {
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
  // Parked alongside `originalBoardState`. `slotLayoutPreferences` is one flat
  // map keyed by slot id that the live board and every scheduled group share,
  // so entering either mode overwrites the live board's copy — without this it
  // never comes back, and the next publish writes the group's layouts.
  const [originalLayoutPreferences, setOriginalLayoutPreferences] = useState<Record<string, CardLayout>>({});
  const [rescheduleDate, setRescheduleDate] = useState<string>('');
  const [rescheduleTime, setRescheduleTime] = useState<string>('09:00');
  const [showRescheduleCalendar, setShowRescheduleCalendar] = useState(false);
  const [toasts, setToasts] = useState<IToast[]>([]);
  const [groupToDelete, setGroupToDelete] = useState<ScheduledStoryGroup | null>(null);
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState(false);
  const [isClearConfirmOpen, setIsClearConfirmOpen] = useState(false);
  const [isPublishConfirmOpen, setIsPublishConfirmOpen] = useState(false);
  /** Search field, so clearing it can hand focus back. */
  const searchInputRef = useRef<HTMLInputElement>(null);
  const [baselineSnapshot, setBaselineSnapshot] = useState<string | null>(null);

  const localDraftKey = useMemo(
    () => [
      'spx-storydragdrop-draft',
      availableStoriesListName || 'AvailableStories',
      publishedStoriesListName || 'PublishedStories',
    ].join(':'),
    [availableStoriesListName, publishedStoriesListName]
  );

  const currentSnapshot = useMemo(
    () => serializeBoardStateForDirtyCheck(selectedLayout, slotStories, slotLayoutPreferences, availableStories),
    [selectedLayout, slotStories, slotLayoutPreferences, availableStories]
  );

  const hasUnsavedChanges = baselineSnapshot !== null && currentSnapshot !== baselineSnapshot;
  const hasStoriesOnBoard = useMemo(
    () => Object.keys(slotStories).some((slotId) => !!slotStories[slotId]),
    [slotStories]
  );
  const isResetDisabled = isSchedulingMode || isGroupEditingMode || !hasUnsavedChanges;
  const isClearDisabled = isSchedulingMode || isGroupEditingMode || !hasStoriesOnBoard;
  const isAtScheduleLimit = scheduledGroups.length >= MAX_SCHEDULED_GROUPS;


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
        
        setAvailableStoriesListName(availableStoriesListName);
        setScheduledStoriesListName(scheduledStoriesListName);
        setPublishedStoriesListName(publishedStoriesListName);

        initializeSharePoint(context);
        
        await new Promise(resolve => setTimeout(resolve, 200));
        
        try {
          const stories = await getStories();
          const formattedStories = formatStories(stories);
          let nextAvailableStories = [...formattedStories];
          let nextSlotStories: SlotStoryMap = {};
          let nextSlotLayoutPreferences: Record<string, CardLayout> = {};
          let nextSelectedLayout: LayoutType | null = null;

          try {
            const published = await getLiveBoard();
            if (published) {
              nextSlotStories = mergeSlotStoriesWithLatest(
                { ...published.slotStories },
                formattedStories
              );
              nextSlotLayoutPreferences = published.slotLayoutPreferences
                ? { ...published.slotLayoutPreferences }
                : {};
              nextSelectedLayout = published.layoutType;

              const usedIds = new Set(
                Object.keys(nextSlotStories)
                  .map((key) => nextSlotStories[key])
                  .filter((story): story is IStory => !!story)
                  .map((story) => story.id)
              );

              const universe: IStory[] = [
                ...formattedStories,
                ...Object.keys(nextSlotStories)
                  .map((key) => nextSlotStories[key])
                  .filter((story): story is IStory => !!story),
              ];

              const seen = new Set<string>();
              nextAvailableStories = universe.filter((story) => {
                if (usedIds.has(story.id) || seen.has(story.id)) return false;
                seen.add(story.id);
                return true;
              });
            }
          } catch (err) {
            console.error('Failed to load published board from SharePoint:', err);
          }

          // The dirty baseline is the *published* board, captured before any
          // local draft is layered on top. Snapshotting after the draft would
          // make a restored draft compare equal to its own baseline, so
          // `hasUnsavedChanges` would read false and the persistence effect
          // below would immediately delete the draft it had just restored.
          const baselineLayout = nextSelectedLayout || 'connectHomepage';
          const baselineSlotStories = nextSlotStories;
          const baselineSlotLayoutPreferences = nextSlotLayoutPreferences;
          const baselineAvailableStories = nextAvailableStories;

          const draft = loadDraft(localDraftKey);
          if (draft) {
            const draftSlots = mergeSlotStoriesWithLatest(draft.slotStories, formattedStories);
            const draftHasStories = Object.keys(draftSlots).some((slotId) => !!draftSlots[slotId]);

            if (draftHasStories) {
              const usedIds = new Set(
                Object.keys(draftSlots)
                  .map((key) => draftSlots[key])
                  .filter((story): story is IStory => !!story)
                  .map((story) => story.id)
              );

              const universe: IStory[] = [
                ...formattedStories,
                ...Object.keys(draftSlots)
                  .map((key) => draftSlots[key])
                  .filter((story): story is IStory => !!story),
              ];

              const seen = new Set<string>();
              nextAvailableStories = universe.filter((story) => {
                if (usedIds.has(story.id) || seen.has(story.id)) return false;
                seen.add(story.id);
                return true;
              });
              nextSlotStories = draftSlots;
              nextSlotLayoutPreferences = { ...draft.slotLayoutPreferences };
              nextSelectedLayout = draft.layout;
            } else {
              clearDraft(localDraftKey);
            }
          }
          
          if (isMounted) {
            setAvailableStories(nextAvailableStories);
            setSlotStories(nextSlotStories);
            setSlotLayoutPreferences(nextSlotLayoutPreferences);
            if (nextSelectedLayout) {
              setSelectedLayout(nextSelectedLayout);
            }
            setBaselineSnapshot(
              serializeBoardStateForDirtyCheck(
                baselineLayout,
                baselineSlotStories,
                baselineSlotLayoutPreferences,
                baselineAvailableStories
              )
            );
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
            // Groups whose moment has gone by are dropped on load rather than
            // rendered as stale rows. They stay in SharePoint — nothing here
            // deletes them — so the list remains the audit trail while the
            // panel only ever shows what is still upcoming. Filtering into
            // state also keeps the schedule limit, the taken-date check and
            // the calendar's disabled days all counting the same groups.
            setScheduledGroups(groups.filter((group) => !isGroupPast(group)));
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
  }, [
    context,
    availableStoriesListName,
    scheduledStoriesListName,
    publishedStoriesListName,
    localDraftKey,
  ]);

  useEffect(() => {
    if (isLoading || isSchedulingMode || isGroupEditingMode) return;

    const hasStories = Object.keys(slotStories).some((slotId) => !!slotStories[slotId]);
    if (!hasStories) {
      clearDraft(localDraftKey);
      return;
    }

    if (!hasUnsavedChanges) {
      clearDraft(localDraftKey);
      return;
    }

    saveDraft(localDraftKey, {
      version: LOCAL_DRAFT_VERSION,
      layout: selectedLayout,
      slotStories,
      slotLayoutPreferences,
    });
  }, [
    isLoading,
    isSchedulingMode,
    isGroupEditingMode,
    slotStories,
    selectedLayout,
    slotLayoutPreferences,
    hasUnsavedChanges,
    localDraftKey,
  ]);

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

      // Belt and braces against the same story landing in two slots: nothing on
      // the board should be in Available to begin with, so this only fires if
      // some path lets a placed story back into the list.
      if (Object.keys(slotStories).some(key => slotStories[key]?.id === story.id)) {
        return;
      }

      const destSlot = resolveDestSlot(overId);
      if (!destSlot) return;

      const occupant = slotStories[destSlot];
      const hasEmptySlot = slotIds.some((slotId) => !slotStories[slotId]);
      if (!hasEmptySlot && occupant) {
        showToast(
          'No Available Slots',
          'error',
          'All board slots are already filled. Remove a story before adding another.'
        );
        return;
      }

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
      setAvailableStories(excludePlacedStories(formattedStories, slotStories));
      setIsModalOpen(false);
      // An active filter would otherwise hide the story that was just added,
      // against a toast saying it succeeded.
      setSearchTerm('');

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
      setAvailableStories(excludePlacedStories(formattedStories, slotStories));

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
      const newSlotStories = { ...slotStories };
      Object.keys(newSlotStories).forEach(key => {
        if (newSlotStories[key]?.id === storyId) {
          delete newSlotStories[key];
        }
      });

      // Filtered against the post-delete slots, so a story deleted while placed
      // is not held back from a list it is no longer in anyway.
      setAvailableStories(excludePlacedStories(formatStories(stories), newSlotStories));
      setSlotStories(newSlotStories);
      setEditingStory(null);
      // Wording is the one the spec calls for, in deliberate deviation from the
      // reference's "Story deleted successfully!". The reference passes no
      // duration, so sonner's own 4000ms default applies rather than the 3000
      // the explicit calls use.
      showToast('Story successfully deleted.', 'success', undefined, 4000);
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

      // A newly added story should start in the default full-image mode.
      setSlotLayoutPreferences((prev) => {
        const next = { ...prev };
        delete next[slotId];
        return next;
      });
    }
  };

  /**
   * Empties every slot, returning all placed stories to the available list.
   *
   * Deliberately silent — the reference's Clear Board raises no toast, and
   * Reset Board reports its own outcome.
   */
  const executeClearBoard = (): void => {
    const storiesInSlots = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== null);
    setAvailableStories([...availableStories, ...storiesInSlots]);
    setSlotStories({});
    setSlotLayoutPreferences({});
    clearDraft(localDraftKey);
    showToast('Board Cleared', 'info', 'All stories moved to available stories.');
  };

  /** Opens a confirmation before clearing the board. */
  const handleClearBoard = (): void => {
    if (isClearDisabled) return;
    setIsClearConfirmOpen(true);
  };

  /** Runs clear only after explicit confirmation. */
  const handleClearBoardConfirmed = (): void => {
    setIsClearConfirmOpen(false);
    executeClearBoard();
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
  const executeResetBoard = (): void => {
    getLiveBoard()
      .then((published) => {
        if (!published) {
          const storiesInSlots = Object.keys(slotStories)
            .map(key => slotStories[key])
            .filter((s): s is IStory => s !== null);
          const nextAvailable = [...availableStories, ...storiesInSlots];
          const nextSlots: SlotStoryMap = {};
          setAvailableStories(nextAvailable);
          setSlotStories(nextSlots);
          setBaselineSnapshot(
            serializeBoardStateForDirtyCheck(
              selectedLayout,
              nextSlots,
              slotLayoutPreferences,
              nextAvailable
            )
          );
          clearDraft(localDraftKey);
          showToast('Board Reset', 'info', 'No previously published board state found. Board cleared.');
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
        const nextPrefs = published.slotLayoutPreferences
          ? { ...published.slotLayoutPreferences }
          : slotLayoutPreferences;

        setSlotStories(restoredSlots);
        setAvailableStories(nextAvailable);
        if (published.slotLayoutPreferences) {
          setSlotLayoutPreferences(nextPrefs);
        }
        setSelectedLayout(published.layoutType);
        setBaselineSnapshot(
          serializeBoardStateForDirtyCheck(
            published.layoutType,
            restoredSlots,
            nextPrefs,
            nextAvailable
          )
        );
        clearDraft(localDraftKey);

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

  /** Opens a confirmation before restoring the published board snapshot. */
  const handleResetBoard = (): void => {
    setIsResetConfirmOpen(true);
  };

  /** Runs the reset flow only after explicit confirmation. */
  const handleResetBoardConfirmed = (): void => {
    setIsResetConfirmOpen(false);
    executeResetBoard();
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
    // The button is disabled at the cap, so this is the backstop that keeps the
    // limit explained rather than silent if that `disabled` is ever removed.
    if (isAtScheduleLimit) {
      showToast(
        'Maximum Groups Reached',
        'error',
        `You can only have up to ${MAX_SCHEDULED_GROUPS} scheduled groups. Please remove a group first.`
      );
      return;
    }

    const parkedCount = Object.keys(slotStories)
      .map((key) => slotStories[key])
      .filter((story): story is IStory => !!story).length;

    const parkedStories = Object.keys(slotStories)
      .map((key) => slotStories[key])
      .filter((story): story is IStory => !!story);

    setOriginalBoardState({ ...slotStories });
    setOriginalLayoutPreferences({ ...slotLayoutPreferences });

    const clearedSlots: SlotStoryMap = {};
    Object.keys(slotStories).forEach(key => {
      clearedSlots[key] = undefined;
    });

    // While scheduling, parked board stories are surfaced in Available so they
    // can be selected for the new scheduled group.
    if (parkedStories.length > 0) {
      setAvailableStories((prev) => {
        const existingIds = new Set(prev.map((story) => story.id));
        const toAdd = parkedStories.filter((story) => !existingIds.has(story.id));
        return [...toAdd, ...prev];
      });
    }

    setSlotStories(clearedSlots);

    setIsSchedulingMode(true);
    setShowCalendarOverlay(false);
    setSelectedScheduleDate('');
    setSelectedScheduleTime('09:00');

    showToast(
      'Scheduling Mode Active',
      'info',
      'Main board cleared. Select a date below to load stories for scheduling.'
    );
  };

  /**
   * Leaves scheduling mode and restores the board and layout preferences
   * parked on entry.
   *
   * Stories still sitting in the slots return to the available list: they left
   * it when they were dropped, and abandoning the schedule means nothing has
   * taken ownership of them.
   *
   * @param returnPlacedStories - Whether to hand the slots' stories back.
   * Pass `false` from the success path, where the new group owns them and a
   * story must live in exactly one place. Callers wired straight to an
   * `onClick` must go through a wrapper, or the click event arrives here as
   * a truthy first argument.
   */
  const handleExitScheduling = (returnPlacedStories: boolean = true): void => {
    const storiesFromBoard: IStory[] = [];
    if (returnPlacedStories) {
      Object.keys(slotStories).forEach(key => {
        const story = slotStories[key];
        if (story) storiesFromBoard.push(story);
      });
    }

    const restoredSlots: SlotStoryMap = { ...originalBoardState };
    const restoredIds = new Set(
      Object.keys(restoredSlots)
        .map((key) => restoredSlots[key])
        .filter((story): story is IStory => !!story)
        .map((story) => story.id)
    );

    // Restoring the parked board must remove those stories from Available so a
    // story is never present in both zones at once.
    setAvailableStories(prev => {
      const existingIds = new Set(prev.map(s => s.id));
      const toAdd = storiesFromBoard.filter(
        s => !existingIds.has(s.id) && !restoredIds.has(s.id)
      );
      const merged = [...toAdd, ...prev].filter(s => !restoredIds.has(s.id));
      const seen = new Set<string>();
      return merged.filter((story) => {
        if (seen.has(story.id)) return false;
        seen.add(story.id);
        return true;
      });
    });

    setSlotStories(restoredSlots);
    setOriginalBoardState({});
    setSlotLayoutPreferences({ ...originalLayoutPreferences });
    setOriginalLayoutPreferences({});

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
      const selectedDate = fromDateKey(selectedScheduleDate);
      const selectedLabel = selectedDate.toLocaleDateString('en-US', {
        weekday: 'long',
        month: 'short',
        day: 'numeric'
      });
      showToast(
        'Date Selected for Scheduling',
        'info',
        `Ready to schedule 0 stories for ${selectedLabel}. Add stories to the main board, then click "Schedule Board".`
      );
      return;
    }

    if (isAtScheduleLimit) {
      showToast('Maximum Groups Reached', 'error', `You can only have up to ${MAX_SCHEDULED_GROUPS} scheduled groups. Please remove a group first.`);
      return;
    }

    const scheduleDate = fromDateKey(selectedScheduleDate);

    // Today is a selectable day, so the hour has to be checked separately.
    if (isMomentPast(scheduleDate, selectedScheduleTime)) {
      showToast('Time Already Passed', 'error', 'Pick a date and time in the future.');
      return;
    }
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
        setScheduledGroups(prev => [...prev, saved].sort(byScheduledMoment));
        // The group now owns these stories, so they stay out of the available
        // list until the group is deleted.
        handleExitScheduling(false);

        const formattedDate = scheduleDate.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric',
          year: 'numeric'
        });
        showToast(
          'Board Scheduled Successfully!',
          'success',
          `${boardStories.length} stories scheduled for ${formattedDate}. Original board restored.`,
          4000
        );
        showToast('Scheduling Complete', 'info', 'Returned to normal mode. Original board layout restored.');
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

    // Delete by item id whenever the group carries one. Deleting by GroupId
    // matches nothing for a row whose column is empty — a group read back as
    // `sp-{id}` — and resolves quietly, which would report success here while
    // the row survived to reappear on the next load.
    const removal = typeof groupToRemove.spId === 'number'
      ? deleteScheduledGroup(groupToRemove.spId)
      : deleteScheduledGroupByGroupId(groupId);

    removal
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

    const restoredSlots: SlotStoryMap = { ...originalBoardState };

    // Entering editing mode pushed the parked board's occupants into the
    // available list, because the board was showing the group in their place.
    // Restoring the board has to take them back out: without this every parked
    // story sits in a slot and in the available list at once, where it can be
    // dragged into a second slot and published twice.
    const restoredIds = new Set(
      Object.keys(restoredSlots)
        .map(key => restoredSlots[key])
        .filter((story): story is IStory => !!story)
        .map(story => story.id)
    );

    setAvailableStories(prev => {
      const existingIds = new Set(prev.map(s => s.id));
      const toAdd = storiesFromBoard.filter(
        s => !existingIds.has(s.id) && !restoredIds.has(s.id)
      );
      return [...toAdd, ...prev].filter(s => !restoredIds.has(s.id));
    });

    setSlotStories(restoredSlots);
    setSlotLayoutPreferences({ ...originalLayoutPreferences });

    setIsGroupEditingMode(false);
    setEditingGroupId(null);
    setOriginalBoardState({});
    setOriginalLayoutPreferences({});
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
    setOriginalLayoutPreferences({ ...slotLayoutPreferences });

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

    // Only when the moment actually moved: a group whose slot has already
    // passed must still accept edits to its stories and layout.
    const movedMoment =
      `${toDateKey(newDate)}T${rescheduleTime}` !== `${toDateKey(group.date)}T${group.time}`;
    if (movedMoment && isMomentPast(newDate, rescheduleTime)) {
      showToast('Time Already Passed', 'error', 'Pick a date and time in the future.');
      return;
    }

    // The reschedule calendar already disables days another group holds, so
    // this is the backstop for a state that changed under it — the reference
    // runs the same check at its own commit point.
    const dateTaken = scheduledGroups.some(
      g => g.id !== editingGroupId && g.date.toDateString() === newDate.toDateString()
    );
    if (dateTaken) {
      showToast('Date Already Scheduled', 'error', "There's already a scheduled group for this date. Please choose a different date.");
      return;
    }

    const newSlotStories = { ...slotStories };

    // Recomputed rather than carried over: the stale instant would otherwise
    // outrank the new date in `byScheduledMoment` until the next reload.
    const nextMoment = toScheduledMoment(toDateKey(newDate), rescheduleTime);

    const applyLocal = (): void => {
      setScheduledGroups(prev => prev.map(g =>
        g.id === editingGroupId
          ? {
            ...g,
            slotStories: newSlotStories,
            slotLayoutPreferences: { ...slotLayoutPreferences },
            date: newDate,
            time: rescheduleTime,
            scheduledAtUtc: new Date(nextMoment.utcIso),
            authorTimeZone: nextMoment.timeZone || undefined,
            authorOffsetMinutes: nextMoment.offsetMinutes,
          }
          : g
      ).sort(byScheduledMoment));
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
   * Only closes the picker and holds the choice — nothing reaches the group or
   * SharePoint until the group edits are saved, and Cancel Edits discards it.
   * The toast says so rather than reporting a success that has not happened
   * yet: unlike the reference, which commits the date the moment the picker is
   * confirmed, this port carries date and arrangement to the server together.
   */
  const handleConfirmReschedule = (): void => {
    const parts = rescheduleDate.split('-');
    const rescheduledDate = rescheduleDate
      ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
      : new Date();

    // Left open, so the admin can correct the time rather than lose the picker.
    if (isMomentPast(rescheduledDate, rescheduleTime)) {
      showToast('Time Already Passed', 'error', 'Pick a date and time in the future.');
      return;
    }

    setShowRescheduleCalendar(false);
    const formattedDate = rescheduledDate.toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
    showToast('New Date Selected', 'info', `Group moves to ${formattedDate} at ${formatTime12(rescheduleTime)} once you save your edits.`);
  };

  /**
   * Publishes the board live, once {@link handlePublishBoardClick} has confirmed.
   *
   * Supersedes the previous live board and writes an immutable snapshot of the
   * slots, their per-slot layouts and the layout preset.
   */
  const handlePublishBoard = (): void => {
    const boardStories = Object.keys(slotStories)
      .map(key => slotStories[key])
      .filter((s): s is IStory => s !== undefined);

    const layoutConfig = getLayoutConfig(selectedLayout);
    publishBoard({
      layoutType: selectedLayout,
      layoutName: layoutConfig.name,
      slotStories: { ...slotStories },
      slotLayoutPreferences
    })
      .then(() => {
        setBaselineSnapshot(currentSnapshot);
        clearDraft(localDraftKey);
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

  /**
   * Opens a confirmation before publishing. The empty-board check runs first:
   * there is nothing to confirm when the publish would be rejected anyway.
   */
  const handlePublishBoardClick = (): void => {
    const placedCount = Object.keys(slotStories)
      .filter(key => slotStories[key] !== undefined).length;

    if (placedCount === 0) {
      showToast(
        'Cannot Publish Empty Board',
        'error',
        'Please add at least one story to the board before publishing. Drag stories from the Available Stories section to any slot.',
        4000
      );
      return;
    }

    setIsPublishConfirmOpen(true);
  };

  /** Runs the publish flow only after explicit confirmation. */
  const handlePublishBoardConfirmed = (): void => {
    setIsPublishConfirmOpen(false);
    handlePublishBoard();
  };

  const sortedAvailableStories = useMemo(
    () => sortAvailableStories(availableStories),
    [availableStories]
  );

  // Trimmed: a trailing space — routine when the term is pasted — would
  // otherwise match nothing and read as "no such story". A whitespace-only
  // term normalises to '' and so leaves the list unfiltered.
  const normalizedSearch = searchTerm.trim().toLowerCase();

  const filteredStories = normalizedSearch
    ? sortedAvailableStories.filter(story =>
      story.title.toLowerCase().includes(normalizedSearch)
    )
    : sortedAvailableStories;

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
            onClick={handlePublishBoardClick}
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
            disabled={isResetDisabled}
          >
            <Icon iconName="Refresh" />
            <span className={styles.btnLabelFull}>Reset Board</span>
            <span className={styles.btnLabelShort}>Reset</span>
          </button>
          <button
            type="button"
            className={styles.btnClear}
            onClick={handleClearBoard}
            disabled={isClearDisabled}
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
                        <div className={styles.calendarOverlay} onClick={() => setShowRescheduleCalendar(false)}>
                          <div onClick={(e) => e.stopPropagation()}>
                            <DateTimePicker
                              selectedDate={rescheduleDate}
                              time={rescheduleTime}
                              onSelectDate={setRescheduleDate}
                              onTimeChange={setRescheduleTime}
                              onConfirm={handleConfirmReschedule}
                              confirmText="Reschedule"
                              isDateDisabled={(date) => isDayBlocked(date, scheduledGroups, editingGroupId)}
                            />
                          </div>
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
                  // Kept rendered at the cap, disabled with a reason: a button
                  // that vanishes explains nothing. `handleScheduleGroup` still
                  // guards, so removing the `disabled` surfaces the toast.
                  <button
                    className={styles.btnSchedule}
                    onClick={handleScheduleGroup}
                    disabled={availableStories.length === 0 || isAtScheduleLimit}
                    title={
                      isAtScheduleLimit
                        ? `Maximum of ${MAX_SCHEDULED_GROUPS} scheduled groups reached. Delete a group to schedule another.`
                        : undefined
                    }
                  >
                    <Icon iconName="Add" />
                    <span className={styles.labelFull}>Schedule New Group</span>
                    <span className={styles.labelShort}>Schedule Group</span>
                  </button>
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
                        <div className={styles.calendarOverlay} onClick={() => setShowCalendarOverlay(false)}>
                          <div onClick={(e) => e.stopPropagation()}>
                            <DateTimePicker
                              selectedDate={selectedScheduleDate}
                              time={selectedScheduleTime}
                              onSelectDate={setSelectedScheduleDate}
                              onTimeChange={setSelectedScheduleTime}
                              onConfirm={handleScheduleNow}
                              confirmText="Schedule"
                              isDateDisabled={isDayPast}
                            />
                          </div>
                        </div>
                      )}
                    </div>
                    <button
                      className={styles.schedulingBtnSmall}
                      onClick={() => handleExitScheduling()}
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
                    ref={searchInputRef}
                    type="search"
                    placeholder="Search stories by title..."
                    aria-label="Search available stories"
                    className={styles.searchInput}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    // Escape is the expected way out of a search field, and the
                    // only one that does not require tabbing to the × button.
                    onKeyDown={(e) => {
                      if (e.key === 'Escape' && searchTerm) {
                        e.preventDefault();
                        setSearchTerm('');
                      }
                    }}
                  />
                  {searchTerm && (
                    <button
                      type="button"
                      className={styles.clearSearchBtn}
                      // This button unmounts with the term it clears, so a
                      // keyboard activation would drop focus to <body> and the
                      // next Tab would restart from the top of the page.
                      onClick={() => {
                        setSearchTerm('');
                        searchInputRef.current?.focus();
                      }}
                      aria-label="Clear search"
                      title="Clear search"
                    >
                      <Icon iconName="Cancel" />
                    </button>
                  )}
                </div>

                {/* Filtering is silent to a screen reader without this: the list
                    swaps under the cursor with no announcement. `aria-live`
                    rather than `role="status"`, which the loading skeleton owns
                    and the tests key off. */}
                <div className={styles.searchStatus} aria-live="polite" aria-atomic="true">
                  {normalizedSearch
                    ? `${filteredStories.length} ${filteredStories.length === 1 ? 'story' : 'stories'} found`
                    : ''}
                </div>

                <SortableContext
                  items={filteredStories.map(s => s.id)}
                  strategy={verticalListSortingStrategy}
                  id="available-stories"
                >
                  <div className={styles.availableStoriesList}>
                    {filteredStories.length === 0 ? (
                      <div className={styles.emptyState}>
                        {normalizedSearch ? (
                          <>
                            <Icon iconName="Search" className={styles.emptySearchIcon} />
                            <p>No stories found matching &quot;{searchTerm.trim()}&quot;</p>
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

      {isResetConfirmOpen && (
        <AlertDialog
          title="Reset board to last published state?"
          titleId="reset-board-heading"
          actionLabel="Yes, Reset Board"
          onCancel={() => setIsResetConfirmOpen(false)}
          onAction={handleResetBoardConfirmed}
        >
          This will discard unsaved board changes and restore the currently published board.
        </AlertDialog>
      )}

      {isClearConfirmOpen && (
        <AlertDialog
          title="Clear all stories from the board?"
          titleId="clear-board-heading"
          actionLabel="Yes, Clear Board"
          onCancel={() => setIsClearConfirmOpen(false)}
          onAction={handleClearBoardConfirmed}
        >
          This will remove all stories from board slots and reset tile display options to defaults.
        </AlertDialog>
      )}

      {isPublishConfirmOpen && (
        <AlertDialog
          title="Publish this board live?"
          titleId="publish-board-heading"
          actionLabel="Yes, Publish Board"
          onCancel={() => setIsPublishConfirmOpen(false)}
          onAction={handlePublishBoardConfirmed}
        >
          This will replace the currently published board and make these stories live for all viewers.
        </AlertDialog>
      )}

      <ToastContainer toasts={toasts} onCloseToast={handleCloseToast} />
    </div>
  );
};

export default StoryDragDrop;
