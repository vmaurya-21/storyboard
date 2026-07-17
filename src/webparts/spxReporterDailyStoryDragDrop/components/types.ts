export interface IStory {
  id: string;
  title: string;
  description: string;
  imageUrl: string;
  linkToPost: string;
  date: string;
  created?: string;
}

export interface IDropBox {
  id: string;
  title: string;
  stories: IStory[];
}

// Layout System Types
export type LayoutType = 'reporterDaily' | 'general' | 'highlight' | 'connectHomepage';

// `variant` doubles as the slot *size*. Legacy values (large/small/textOnly/
// fullHeight/horizontal/general) are retained for back-compat; `medium` and
// `tall` were added to mirror the reference board's slot heights.
export type CardVariant =
  | 'large'
  | 'medium'
  | 'small'
  | 'tall'
  | 'textOnly'
  | 'fullHeight'
  | 'horizontal'
  | 'general';

// How a filled card renders its content, ported from the reference StorySlot.
export type CardLayout = 'full-image' | 'thumbnail-text' | 'text-only';

export interface CardSlotConfig {
  id: string;
  variant: CardVariant;
  cardLayout?: CardLayout;
  gridArea?: string;
  showImage?: boolean; // For text-only variant
}

export interface LayoutConfig {
  id: LayoutType;
  name: string;
  description: string;
  slots: CardSlotConfig[];
  gridTemplate: {
    columns: string;
    rows: string;
    areas?: string;
  };
}

export type SlotStoryMap = Record<string, IStory | undefined>;

export interface ScheduledStoryGroup {
  id: string;
  /** SharePoint list item id — present once the group is persisted. */
  spId?: number;
  date: Date;
  time: string;
  slotStories: SlotStoryMap;
  layoutType: LayoutType;
  layoutName: string;
  createdAt: Date;
}
