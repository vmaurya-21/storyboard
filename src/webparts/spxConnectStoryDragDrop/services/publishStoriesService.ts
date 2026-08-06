
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { getSp } from "./availableStoriesService";
import { SCHEDULE_LIST_NAME, getScheduledStoriesListName, setGroupStatus } from "./scheduleStoriesService";
import { IStory, LayoutType, SlotStoryMap, CardLayout } from "../components/types";
import { getLayoutConfig } from "../components/layouts/layoutConfig";

/** Default title of the SharePoint list holding published board snapshots. */
export const DEFAULT_PUBLISHED_STORIES_LIST_NAME = "PublishedStories";

/** Backward-compatible alias used by existing tests/imports. */
export const PUBLISH_LIST_NAME = DEFAULT_PUBLISHED_STORIES_LIST_NAME;

let publishedStoriesListName = DEFAULT_PUBLISHED_STORIES_LIST_NAME;

/** Override the SharePoint list title used by the publish service. */
export const setPublishedStoriesListName = (listName?: string): void => {
  const next = (listName || "").trim();
  publishedStoriesListName = next || DEFAULT_PUBLISHED_STORIES_LIST_NAME;
};

/** Current SharePoint list title used by the publish service. */
export const getPublishedStoriesListName = (): string => publishedStoriesListName;

const VALID_LAYOUT_TYPES: LayoutType[] = ["reporterDaily", "general", "highlight", "connectHomepage"];
const DEFAULT_LAYOUT_TYPE: LayoutType = "connectHomepage";

/** Value of the `IsLive` column marking the snapshot currently on air. */
export const IS_LIVE_YES = "Yes";

/** Value of the `IsLive` column marking a superseded snapshot. */
export const IS_LIVE_NO = "No";

/** Internal name of the lookup column pointing back at a Schedule Stories row. */
const SOURCE_LOOKUP_ID_FIELD = "SourceScheduleIdId";

/** Everything {@link publishBoard} needs to write a snapshot. */
export interface IPublishBoardInput {
  /** Layout preset the board was arranged in. */
  layoutType: LayoutType;
  /** Display name of the preset. Resolved from the preset when omitted. */
  layoutName?: string;
  /** The arrangement to publish. Empty slots are dropped from the snapshot. */
  slotStories: SlotStoryMap;
  /** Per-slot card layout overrides to capture alongside the arrangement. */
  slotLayoutPreferences?: Record<string, CardLayout>;
  /** SharePoint item id of the Schedule Stories row, when publishing a schedule. */
  sourceScheduleSpId?: number;
  /** App id of the schedule group, kept in the snapshot for traceability. */
  sourceScheduleGroupId?: string;
  /** Guard against accidentally publishing an empty board. Default: false (throws). */
  allowEmpty?: boolean;
}

/** A published snapshot, rehydrated from its list item. */
export interface IPublishedBoardRecord {
  /** SharePoint item id of the snapshot. */
  spId: number;
  /** When the board was published. */
  publishedAt: Date;
  /** Layout preset captured in the snapshot. */
  layoutType: LayoutType;
  /** Display name of the preset at publish time. */
  layoutName: string;
  /** The published arrangement. */
  slotStories: SlotStoryMap;
  /** Per-slot card layout overrides captured with the arrangement. */
  slotLayoutPreferences?: Record<string, CardLayout>;
  /** Whether this snapshot is the one currently on air. */
  isLive: boolean;
  /** Schedule Stories item this snapshot was published from, if any. */
  sourceScheduleSpId?: number;
}

/** Raw shape of a "Publish Stories" list item as read from SharePoint. */
interface IRawPublishItem {
  /** List item id. */
  Id: number;
  /** Optional alternate id key some clients/projectors return. */
  ID?: number;
  /** Item title, used only for readability in the list view. */
  Title?: string;
  /** Publish timestamp, in ISO form. */
  PublishedDateTime?: string;
  /** Layout preset id, validated on read. */
  LayoutType?: string;
  /** The serialized {@link IPublishBoardStateJson} payload. */
  BoardStateJson?: string;
  /** Lookup id of the Schedule Stories row this was published from. */
  SourceScheduleIdId?: number | null;
  /** Fallback lookup projection shape. */
  SourceScheduleId?: number | { Id?: number } | null;
  /** `Yes` or `No`; see {@link IS_LIVE_YES}. */
  IsLive?: string;
}

/** Serialized board snapshot stored in BoardStateJson. */
interface IPublishBoardStateJson {
  /** Layout preset the arrangement belongs to. */
  layoutType: LayoutType;
  /** Display name of the preset at publish time. */
  layoutName: string;
  /** Publish timestamp, in ISO form. */
  publishedAt: string;
  /** App-side id of the schedule group this came from, for traceability. */
  sourceScheduleGroupId?: string;
  /** Filled slots only, each with its story and optional layout override. */
  slots: Array<{ slotId: string; story: IStory; cardLayout?: CardLayout }>;
}


const getSpOrThrow = (): SPFI => {
  const sp = getSp();
  if (!sp || !sp.web) {
    throw new Error(
      "SharePoint not initialized. Call initializeSharePoint(context) before using the Publish Stories service."
    );
  }
  return sp;
};

const isValidLayoutType = (value: string | undefined): value is LayoutType =>
  !!value && VALID_LAYOUT_TYPES.indexOf(value as LayoutType) !== -1;

const countStories = (slotStories: SlotStoryMap): number =>
  Object.keys(slotStories).filter((k) => !!slotStories[k]).length;


const serializeBoardState = (input: IPublishBoardInput, publishedAt: Date): string => {
  const layoutName = input.layoutName || getLayoutConfig(input.layoutType).name;
  const slots: Array<{ slotId: string; story: IStory; cardLayout?: CardLayout }> = Object.keys(input.slotStories)
    .map((slotId) => {
      const story = input.slotStories[slotId];
      const cardLayout = input.slotLayoutPreferences ? input.slotLayoutPreferences[slotId] : undefined;
      return { slotId, story, cardLayout } as { slotId: string; story: IStory | undefined; cardLayout?: CardLayout };
    })
    .filter((e): e is { slotId: string; story: IStory; cardLayout?: CardLayout } => !!e.story);

  const payload: IPublishBoardStateJson = {
    layoutType: input.layoutType,
    layoutName,
    publishedAt: publishedAt.toISOString(),
    sourceScheduleGroupId: input.sourceScheduleGroupId,
    slots,
  };
  return JSON.stringify(payload);
};

const parseBoardState = (json: string | undefined): IPublishBoardStateJson | undefined => {
  if (!json) return undefined;
  try {
    const parsed = JSON.parse(json) as Partial<IPublishBoardStateJson>;
    if (!parsed || !Array.isArray(parsed.slots)) return undefined;
    return {
      layoutType: isValidLayoutType(parsed.layoutType) ? parsed.layoutType : DEFAULT_LAYOUT_TYPE,
      layoutName: parsed.layoutName || "",
      publishedAt: parsed.publishedAt || "",
      sourceScheduleGroupId: parsed.sourceScheduleGroupId,
      slots: parsed.slots.filter(
        (s): s is { slotId: string; story: IStory; cardLayout?: CardLayout } => !!s && !!s.slotId && !!s.story
      ),
    };
  } catch (err) {
    console.error("Publish Stories: failed to parse BoardStateJson", err);
    return undefined;
  }
};

const boardStateToSlotMap = (state: IPublishBoardStateJson): SlotStoryMap => {
  const map: SlotStoryMap = {};
  state.slots.forEach(({ slotId, story }) => {
    map[slotId] = story;
  });
  return map;
};

const mapRawToRecord = (raw: IRawPublishItem): IPublishedBoardRecord | undefined => {
  const state = parseBoardState(raw.BoardStateJson);
  if (!state) {
    console.warn(`Publish Stories: skipping item ${raw.Id} — unreadable BoardStateJson.`);
    return undefined;
  }

  const layoutType = isValidLayoutType(state.layoutType)
    ? state.layoutType
    : isValidLayoutType(raw.LayoutType)
      ? raw.LayoutType
      : DEFAULT_LAYOUT_TYPE;

  const publishedAt = state.publishedAt
    ? new Date(state.publishedAt)
    : raw.PublishedDateTime
      ? new Date(raw.PublishedDateTime)
      : new Date();

  const slotLayoutPreferences: Record<string, CardLayout> = {};
  state.slots.forEach(({ slotId, cardLayout }) => {
    if (cardLayout) {
      slotLayoutPreferences[slotId] = cardLayout;
    }
  });

  const sourceScheduleSpId = typeof raw.SourceScheduleIdId === "number"
    ? raw.SourceScheduleIdId
    : typeof raw.SourceScheduleId === "number"
      ? raw.SourceScheduleId
      : typeof raw.SourceScheduleId === "object" && typeof raw.SourceScheduleId?.Id === "number"
        ? raw.SourceScheduleId.Id
        : undefined;

  const spId = typeof raw.Id === "number" ? raw.Id : raw.ID;
  if (typeof spId !== "number") {
    return undefined;
  }

  return {
    spId,
    publishedAt,
    layoutType,
    layoutName: state.layoutName || getLayoutConfig(layoutType).name,
    slotStories: boardStateToSlotMap(state),
    slotLayoutPreferences,
    isLive: raw.IsLive === IS_LIVE_YES,
    sourceScheduleSpId,
  };
};

const SELECT_FIELDS: string[] = [
  "Id",
  "Title",
  "PublishedDateTime",
  "LayoutType",
  "BoardStateJson",
  "SourceScheduleIdId",
  "IsLive",
];


/**
 * Flip every currently-live record to IsLive=No. Normally there is at most one,
 * but we sweep all matches to self-heal any accidental duplicates.
 */
const supersedeLive = async (sp: SPFI): Promise<void> => {
  const live: Array<{ Id: number }> = await sp.web.lists
    .getByTitle(getPublishedStoriesListName())
    .items.select("Id")
    .filter(`IsLive eq '${IS_LIVE_YES}'`)
    .top(50)();

  for (const r of live) {
    await sp.web.lists
      .getByTitle(getPublishedStoriesListName())
      .items.getById(r.Id)
      .update({ IsLive: IS_LIVE_NO });
  }
};

/**
 * Keep the "Publish Stories" list at (at most) two rows: the one that is
 * currently live and the one published just before it — the board "Reset Board"
 * restores to. Everything older is deleted.
 *
 * Best-effort: a failure here must never fail an otherwise successful publish.
 */
const prunePublishHistory = async (sp: SPFI): Promise<void> => {
  const items: Array<{ Id: number }> = await sp.web.lists
    .getByTitle(getPublishedStoriesListName())
    .items.select("Id")
    .orderBy("PublishedDateTime", false)
    .top(500)();

  if (items.length <= 2) return;

  for (const item of items.slice(2)) {
    await sp.web.lists.getByTitle(getPublishedStoriesListName()).items.getById(item.Id).delete();
  }
};


/**
 * Publish the current board. Supersedes the previous live record, writes a new
 * immutable snapshot as the live one, and (when published from a schedule) marks
 * that schedule row Published. Returns the new live record.
 *
 * Throws on an empty board unless `allowEmpty: true` — publishing nothing would
 * silently blank the live board; use `unpublishBoard` for that intent instead.
 */
export const publishBoard = async (
  input: IPublishBoardInput
): Promise<IPublishedBoardRecord> => {
  const sp = getSpOrThrow();

  const storyCount = countStories(input.slotStories);
  if (storyCount === 0 && !input.allowEmpty) {
    throw new Error(
      "Cannot publish an empty board. Add at least one story, or call unpublishBoard to take the current board offline."
    );
  }

  try {
    await supersedeLive(sp);

    const publishedAt = new Date();
    const layoutName = input.layoutName || getLayoutConfig(input.layoutType).name;
    const label = `${publishedAt.toISOString().slice(0, 16).replace("T", " ")} · ${layoutName}`;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {
      Title: label,
      PublishedDateTime: publishedAt.toISOString(),
      LayoutType: input.layoutType,
      BoardStateJson: serializeBoardState(input, publishedAt),
      IsLive: IS_LIVE_YES,
    };
    if (typeof input.sourceScheduleSpId === "number") {
      payload[SOURCE_LOOKUP_ID_FIELD] = input.sourceScheduleSpId;
    }

    const result = await sp.web.lists.getByTitle(getPublishedStoriesListName()).items.add(payload);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data: any = result.data;

    try {
      await prunePublishHistory(sp);
    } catch (pruneErr) {
      console.warn("Publish Stories: published ok, but history pruning failed", pruneErr);
    }

    if (typeof input.sourceScheduleSpId === "number") {
      try {
        await setGroupStatus(input.sourceScheduleSpId, "Published");
      } catch (statusErr) {
        console.warn(
          `Publish Stories: published ok, but failed to mark schedule ${input.sourceScheduleSpId} as Published`,
          statusErr
        );
      }
    }

    return {
      spId: data.Id,
      publishedAt,
      layoutType: input.layoutType,
      layoutName,
      slotStories: { ...input.slotStories },
      isLive: true,
      sourceScheduleSpId: input.sourceScheduleSpId,
    };
  } catch (error) {
    console.error("Publish Stories: failed to publish board", error);
    throw error;
  }
};

/** The current live board, or undefined if nothing is live. */
export const getLiveBoard = async (): Promise<IPublishedBoardRecord | undefined> => {
  const sp = getSpOrThrow();
  try {
    const items: IRawPublishItem[] = await sp.web.lists
      .getByTitle(getPublishedStoriesListName())
      .items.select(...SELECT_FIELDS)
      .filter(`IsLive eq '${IS_LIVE_YES}'`)
      .orderBy("PublishedDateTime", false)
      .top(1)();

    if (!items.length) return undefined;
    return mapRawToRecord(items[0]);
  } catch (error) {
    console.error("Publish Stories: failed to fetch live board", error);
    throw error;
  }
};

/** Full publish history, newest first. Corrupt rows are skipped, not thrown. */
export const getPublishedBoards = async (): Promise<IPublishedBoardRecord[]> => {
  const sp = getSpOrThrow();
  try {
    const items: IRawPublishItem[] = await sp.web.lists
      .getByTitle(getPublishedStoriesListName())
      .items.select(...SELECT_FIELDS)
      .orderBy("PublishedDateTime", false)
      .top(500)();

    return items
      .map(mapRawToRecord)
      .filter((r): r is IPublishedBoardRecord => r !== undefined);
  } catch (error) {
    console.error("Publish Stories: failed to fetch published boards", error);
    throw error;
  }
};

/**
 * Take the current board offline without publishing a new one. Marks all live
 * records IsLive=No. Safe when nothing is live (no-op).
 */
export const unpublishBoard = async (): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    const live: Array<{ Id: number }> = await sp.web.lists
      .getByTitle(getPublishedStoriesListName())
      .items.select("Id")
      .filter(`IsLive eq '${IS_LIVE_YES}'`)
      .top(50)();

    for (const r of live) {
      await sp.web.lists
        .getByTitle(getPublishedStoriesListName())
        .items.getById(r.Id)
        .update({ IsLive: IS_LIVE_NO });
    }
  } catch (error) {
    console.error("Publish Stories: failed to unpublish board", error);
    throw error;
  }
};

/** Delete a publish record by SharePoint item id (history cleanup). */
export const deletePublishedBoard = async (spId: number): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    await sp.web.lists.getByTitle(getPublishedStoriesListName()).items.getById(spId).delete();
  } catch (error) {
    console.error(`Publish Stories: failed to delete published board ${spId}`, error);
    throw error;
  }
};


/**
 * Ensure the "Publish Stories" list and its columns exist. Safe to call
 * repeatedly. The SourceScheduleId lookup points at the "Schedule Stories" list,
 * which therefore must exist first. Only a safety net — you created the list
 * manually.
 */
export const ensurePublishStoriesList = async (): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    const lists: Array<{ Title: string }> = await sp.web.lists();
    const listName = getPublishedStoriesListName();
    const exists = lists.some((l) => l.Title === listName);
    if (!exists) {
      await sp.web.lists.add(listName);
    }
    const list = sp.web.lists.getByTitle(listName);

    const addSafely = async (fn: () => Promise<unknown>): Promise<void> => {
      try {
        await fn();
      } catch {
      }
    };

    await addSafely(() => list.fields.addDateTime("PublishedDateTime"));
    await addSafely(() => list.fields.addChoice("LayoutType", { Choices: VALID_LAYOUT_TYPES }));
    await addSafely(() => list.fields.addMultilineText("BoardStateJson"));
    await addSafely(() => list.fields.addChoice("IsLive", { Choices: [IS_LIVE_YES, IS_LIVE_NO] }));

    await addSafely(async () => {
      const scheduleListTitle = getScheduledStoriesListName() || SCHEDULE_LIST_NAME;
      const scheduleList = await sp.web.lists.getByTitle(scheduleListTitle).select("Id")();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      await (list.fields as any).addLookup("SourceScheduleId", {
        LookupListId: (scheduleList as { Id: string }).Id,
        LookupFieldName: "Title",
      });
    });
  } catch (error) {
    console.error("Publish Stories: ensurePublishStoriesList failed", error);
    throw error;
  }
};
