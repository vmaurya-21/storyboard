/**
 * A single story as the UI consumes it.
 *
 * This is the view-model shape, not the raw SharePoint list item: ids are
 * prefixed strings and dates are pre-formatted for display. See
 * `formatStories` in `StoryDragDrop` for the mapping from `IStoryItem`.
 */
export interface IStory {
  /** Stable identifier in the form `story-{listItemId}`. */
  id: string;
  /** Headline shown on the card and in the Available Stories row. */
  title: string;
  /** Optional summary body. May be an empty string. */
  description: string;
  /** Absolute URL of the hero/thumbnail image. */
  imageUrl: string;
  /** Absolute URL the card links out to. `'#'` when unset. */
  linkToPost: string;
  /** Display-ready publish date, e.g. `Mar 4`. */
  date: string;
  /** Raw SharePoint `Created` timestamp, used for sorting. */
  created?: string;
  /** Origin of the post: `internal`, `linkedin` or `external`. */
  source?: string;
  /**
   * Byline — the display name from the story's SharePoint Created By.
   *
   * Optional on purpose: board states published before this field existed have
   * no `author` key, and slides fall back to showing the date alone rather
   * than an empty byline. Snapshotted into `BoardStateJson` alongside the rest
   * of the story, so a published board keeps its byline even if the source
   * story is later edited or deleted.
   */
  author?: string;
}

/**
 * A named bucket of stories.
 *
 * Retained for the drag-and-drop container model; the board itself addresses
 * stories by slot id rather than by box.
 */
export interface IDropBox {
  /** Droppable container id used by dnd-kit. */
  id: string;
  /** Heading rendered above the box. */
  title: string;
  /** Stories currently held in the box, in display order. */
  stories: IStory[];
}

/** Identifier of a board layout preset. */
export type LayoutType = 'reporterDaily' | 'general' | 'highlight' | 'connectHomepage';

/**
 * Size/shape family of a card slot.
 *
 * Drives the slot's dimensions and typography scale; the layout preset decides
 * which variant each slot uses.
 */
export type CardVariant =
  | 'large'
  | 'medium'
  | 'small'
  | 'tall'
  | 'textOnly'
  | 'fullHeight'
  | 'horizontal'
  | 'general';

/**
 * How a card renders its content.
 *
 * `full-image` overlays text on the image, `thumbnail-text` splits the two
 * side by side, and `text-only` drops the image entirely.
 */
export type CardLayout = 'full-image' | 'thumbnail-text' | 'text-only';

/** One slot position within a {@link LayoutConfig}. */
export interface CardSlotConfig {
  /** Slot id, e.g. `slot-1`. Doubles as the dnd-kit droppable id. */
  id: string;
  /** Size family for this position. */
  variant: CardVariant;
  /** Default content layout; editors can override it per slot. */
  cardLayout?: CardLayout;
  /** `grid-area` name this slot occupies in the preset's grid. */
  gridArea?: string;
  /** Set to `false` to suppress the image regardless of `cardLayout`. */
  showImage?: boolean;
}

/** A board layout preset: its slots and the CSS grid they sit in. */
export interface LayoutConfig {
  /** Preset identifier. */
  id: LayoutType;
  /** Display name shown in the layout picker and in toasts. */
  name: string;
  /** One-line summary of the preset shown in the picker. */
  description: string;
  /** Slot definitions in render order. */
  slots: CardSlotConfig[];
  /** CSS grid definition backing the preset. */
  gridTemplate: {
    /** Value for `grid-template-columns`. */
    columns: string;
    /** Value for `grid-template-rows`. */
    rows: string;
    /** Value for `grid-template-areas`, when the preset uses named areas. */
    areas?: string;
  };
}

/**
 * Board state: which story sits in which slot.
 *
 * Keys are slot ids; a missing or `undefined` value means the slot is empty.
 */
export type SlotStoryMap = Record<string, IStory | undefined>;

/** A board arrangement saved for publication at a future date and time. */
export interface ScheduledStoryGroup {
  /** Client-side identifier in the form `scheduled-{timestamp}`. */
  id: string;
  /** SharePoint list item id — present once the group is persisted. */
  spId?: number;
  /**
   * Local date the group is scheduled for. The time component is unused.
   *
   * This and {@link ScheduledStoryGroup.time} are the *viewer's* rendering of
   * {@link ScheduledStoryGroup.scheduledAtUtc}: an admin in another timezone
   * sees the same instant as a different day and hour.
   */
  date: Date;
  /** Scheduled time of day in 24-hour `HH:mm` form, in the viewer's zone. */
  time: string;
  /**
   * The absolute instant the group is scheduled for. Absent only on groups
   * held in memory before their first save.
   */
  scheduledAtUtc?: Date;
  /** IANA zone the group was scheduled from, e.g. `Asia/Kolkata`. */
  authorTimeZone?: string;
  /** The author zone's offset from UTC, in minutes, at that instant. */
  authorOffsetMinutes?: number;
  /** The slot-to-story arrangement captured when the group was saved. */
  slotStories: SlotStoryMap;
  /** Per-slot card layout overrides captured alongside the arrangement. */
  slotLayoutPreferences?: Record<string, CardLayout>;
  /** Layout preset the arrangement was built for. */
  layoutType: LayoutType;
  /** Display name of {@link ScheduledStoryGroup.layoutType} at save time. */
  layoutName: string;
  /** When the group was created. */
  createdAt: Date;
}
