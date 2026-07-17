// =============================================================
// Publish Stories service
// -------------------------------------------------------------
// Records boards that have been pushed live to the "Publish Stories"
// SharePoint list. This is an append-only history: each publish writes a NEW,
// immutable snapshot and supersedes the previous live one. The point of an
// immutable snapshot is that later edits/deletes in "Available Stories" never
// rewrite what actually went out.
//
// List columns this service targets:
//   Title              Single line   - human label (auto-generated)
//   PublishedDateTime  Date & Time   - when it went live (sort key)
//   LayoutType         Choice        - reporterDaily | general | highlight
//   BoardStateJson     Multi-line    - immutable board snapshot (source of truth)
//   SourceScheduleId   Lookup        - -> Schedule Stories item (if published from a schedule)
//   IsLive             Choice        - "Yes" | "No" (exactly one "Yes" at a time)
//
// Notes:
//  * A Lookup is written through its id field ("<InternalName>Id"), i.e.
//    "SourceScheduleIdId", set to the Schedule Stories list-item id.
//  * IsLive is the sole live indicator: publishing flips the previous live row to
//    IsLive=No and adds the new one as IsLive=Yes. When published from a schedule
//    it also flips that schedule row to Status=Published (on the Schedule list).
// =============================================================

import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { getSp } from "./sharePointService";
import { SCHEDULE_LIST_NAME, setGroupStatus } from "./scheduleStoriesService";
import { IStory, LayoutType, SlotStoryMap } from "../components/types";
import { getLayoutConfig } from "../components/layouts/layoutConfig";

// ---- Constants ----------------------------------------------
export const PUBLISH_LIST_NAME = "Publish Stories";

const VALID_LAYOUT_TYPES: LayoutType[] = ["reporterDaily", "general", "highlight"];
const DEFAULT_LAYOUT_TYPE: LayoutType = "reporterDaily";

// IsLive is a Choice column, so its values are strings. Keep them centralized.
export const IS_LIVE_YES = "Yes";
export const IS_LIVE_NO = "No";

// The lookup's id field: internal name of the column + "Id".
const SOURCE_LOOKUP_ID_FIELD = "SourceScheduleIdId";

// ---- Domain types -------------------------------------------

export interface IPublishBoardInput {
  layoutType: LayoutType;
  layoutName?: string;
  slotStories: SlotStoryMap;
  /** SharePoint item id of the Schedule Stories row, when publishing a schedule. */
  sourceScheduleSpId?: number;
  /** App id of the schedule group, kept in the snapshot for traceability. */
  sourceScheduleGroupId?: string;
  /** Guard against accidentally publishing an empty board. Default: false (throws). */
  allowEmpty?: boolean;
}

export interface IPublishedBoardRecord {
  spId: number;
  publishedAt: Date;
  layoutType: LayoutType;
  layoutName: string;
  slotStories: SlotStoryMap;
  isLive: boolean;
  sourceScheduleSpId?: number;
}

/** Raw shape of a "Publish Stories" list item as read from SharePoint. */
interface IRawPublishItem {
  Id: number;
  Title?: string;
  PublishedDateTime?: string;
  LayoutType?: string;
  BoardStateJson?: string;
  SourceScheduleIdId?: number | null;
  IsLive?: string;
}

/** Serialized board snapshot stored in BoardStateJson. */
interface IPublishBoardStateJson {
  layoutType: LayoutType;
  layoutName: string;
  publishedAt: string; // ISO — source of truth for the timestamp
  sourceScheduleGroupId?: string;
  slots: Array<{ slotId: string; story: IStory }>;
}

// ---- Guards / helpers ---------------------------------------

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

// ---- (De)serialization --------------------------------------

const serializeBoardState = (input: IPublishBoardInput, publishedAt: Date): string => {
  const layoutName = input.layoutName || getLayoutConfig(input.layoutType).name;
  const slots = Object.keys(input.slotStories)
    .map((slotId) => ({ slotId, story: input.slotStories[slotId] }))
    .filter((e): e is { slotId: string; story: IStory } => !!e.story);

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
        (s): s is { slotId: string; story: IStory } => !!s && !!s.slotId && !!s.story
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

  // Prefer the JSON timestamp; fall back to the column.
  const publishedAt = state.publishedAt
    ? new Date(state.publishedAt)
    : raw.PublishedDateTime
      ? new Date(raw.PublishedDateTime)
      : new Date();

  return {
    spId: raw.Id,
    publishedAt,
    layoutType,
    layoutName: state.layoutName || getLayoutConfig(layoutType).name,
    slotStories: boardStateToSlotMap(state),
    isLive: raw.IsLive === IS_LIVE_YES,
    sourceScheduleSpId:
      typeof raw.SourceScheduleIdId === "number" ? raw.SourceScheduleIdId : undefined,
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

// ---- Internal: supersede the current live record(s) ---------

/**
 * Flip every currently-live record to IsLive=No. Normally there is at most one,
 * but we sweep all matches to self-heal any accidental duplicates.
 */
const supersedeLive = async (sp: SPFI): Promise<void> => {
  const live: Array<{ Id: number }> = await sp.web.lists
    .getByTitle(PUBLISH_LIST_NAME)
    .items.select("Id")
    .filter(`IsLive eq '${IS_LIVE_YES}'`)
    .top(50)();

  for (const r of live) {
    await sp.web.lists
      .getByTitle(PUBLISH_LIST_NAME)
      .items.getById(r.Id)
      .update({ IsLive: IS_LIVE_NO });
  }
};

// ---- Public API ---------------------------------------------

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
    // 1) Supersede whatever is currently live so only one row stays live.
    await supersedeLive(sp);

    // 2) Write the new live snapshot.
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

    const result = await sp.web.lists.getByTitle(PUBLISH_LIST_NAME).items.add(payload);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data: any = result.data;

    // 3) Best-effort: mark the source schedule as Published. Never let this
    //    failure roll back a successful publish.
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
      .getByTitle(PUBLISH_LIST_NAME)
      .items.select(...SELECT_FIELDS)
      .filter(`IsLive eq '${IS_LIVE_YES}'`)
      .orderBy("PublishedDateTime", false) // newest first
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
      .getByTitle(PUBLISH_LIST_NAME)
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
      .getByTitle(PUBLISH_LIST_NAME)
      .items.select("Id")
      .filter(`IsLive eq '${IS_LIVE_YES}'`)
      .top(50)();

    for (const r of live) {
      await sp.web.lists
        .getByTitle(PUBLISH_LIST_NAME)
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
    await sp.web.lists.getByTitle(PUBLISH_LIST_NAME).items.getById(spId).delete();
  } catch (error) {
    console.error(`Publish Stories: failed to delete published board ${spId}`, error);
    throw error;
  }
};

// ---- Optional provisioning (self-heal missing list/columns) --

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
    const exists = lists.some((l) => l.Title === PUBLISH_LIST_NAME);
    if (!exists) {
      await sp.web.lists.add(PUBLISH_LIST_NAME);
    }
    const list = sp.web.lists.getByTitle(PUBLISH_LIST_NAME);

    const addSafely = async (fn: () => Promise<unknown>): Promise<void> => {
      try {
        await fn();
      } catch {
        /* column already exists — ignore */
      }
    };

    await addSafely(() => list.fields.addDateTime("PublishedDateTime"));
    await addSafely(() => list.fields.addChoice("LayoutType", { Choices: VALID_LAYOUT_TYPES }));
    await addSafely(() => list.fields.addMultilineText("BoardStateJson"));
    await addSafely(() => list.fields.addChoice("IsLive", { Choices: [IS_LIVE_YES, IS_LIVE_NO] }));

    // Lookup to Schedule Stories (needs that list's id).
    await addSafely(async () => {
      const scheduleList = await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).select("Id")();
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
