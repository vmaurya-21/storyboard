
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { getSp } from "./availableStoriesService";
import { IStory, LayoutType, SlotStoryMap, CardLayout } from "../components/types";
import { getLayoutConfig } from "../components/layouts/layoutConfig";

/** Title of the SharePoint list holding scheduled groups. */
export const SCHEDULE_LIST_NAME = "Schedule Stories";

/** Most scheduled groups that may exist at once. */
export const MAX_SCHEDULED_GROUPS = 10;

const VALID_LAYOUT_TYPES: LayoutType[] = ["reporterDaily", "general", "highlight", "connectHomepage"];
const DEFAULT_LAYOUT_TYPE: LayoutType = "connectHomepage";

/**
 * Lifecycle state of a scheduled group.
 *
 * A group starts `Scheduled`, becomes `Published` once its board goes live,
 * and is `Cancelled` or `Expired` when it will not publish.
 */
export type ScheduleStatus = "Scheduled" | "Published" | "Cancelled" | "Expired";
const VALID_STATUSES: ScheduleStatus[] = ["Scheduled", "Published", "Cancelled", "Expired"];
const DEFAULT_STATUS: ScheduleStatus = "Scheduled";


/**
 * A persisted scheduled group. Structurally a superset of the app's
 * `ScheduledStoryGroup` (same id/date/time/slotStories/layoutType/layoutName/
 * createdAt), plus the SharePoint item id and lifecycle status. This means a
 * record returned here can be used directly wherever `ScheduledStoryGroup` is
 * expected.
 */
export interface IScheduledGroupRecord {
  /** App id, stored in the GroupId column (e.g. "scheduled-1720000000000"). */
  id: string;
  /** SharePoint list item Id — needed for update/delete. */
  spId: number;
  /**
   * Scheduled calendar day, as local midnight. Never convert it to UTC — it
   * denotes a civil date, so shifting it by an offset changes which day it is.
   */
  date: Date;
  /** Scheduled time as "HH:mm". */
  time: string;
  /** slotId -> story snapshot. */
  slotStories: SlotStoryMap;
  /** Per-slot card layout overrides captured with the arrangement. */
  slotLayoutPreferences?: Record<string, CardLayout>;
  /** Layout preset the arrangement was built for. */
  layoutType: LayoutType;
  /** Display name of the preset at save time. */
  layoutName: string;
  /** Lifecycle state of the group. */
  status: ScheduleStatus;
  /** When the list item was created. */
  createdAt: Date;
}

/** Input for creating a new scheduled group (id supplied by the app). */
export interface INewScheduledGroup {
  /** App id, stored in the GroupId column. */
  id: string;
  /** Scheduled calendar day. */
  date: Date;
  /** Scheduled time as "HH:mm". */
  time: string;
  /** The arrangement to schedule. */
  slotStories: SlotStoryMap;
  /** Per-slot card layout overrides to capture alongside the arrangement. */
  slotLayoutPreferences?: Record<string, CardLayout>;
  /** Layout preset the arrangement was built for. */
  layoutType: LayoutType;
  /** Display name of the preset. Resolved from the preset when omitted. */
  layoutName?: string;
  /** Initial lifecycle state. Defaults to `Scheduled`. */
  status?: ScheduleStatus;
}

/** Raw shape of a "Schedule Stories" list item as read from SharePoint. */
interface IRawScheduleItem {
  /** List item id. */
  Id: number;
  /** Item title, used only for readability in the list view. */
  Title?: string;
  /** App-side group id. */
  GroupId?: string;
  /** Scheduled day, in ISO form. */
  ScheduleDate?: string;
  /** Scheduled day and time combined, in ISO form. */
  ScheduledDateTime?: string;
  /** Layout preset id, validated on read. */
  LayoutType?: string;
  /** The serialized {@link IBoardStateJson} payload. */
  BoardStateJson?: string;
  /** Lifecycle state, validated on read. */
  Status?: string;
  /** Creation timestamp. */
  Created?: string;
}

/** Serialized board snapshot stored in BoardStateJson. */
interface IBoardStateJson {
  /** Layout preset the arrangement belongs to. */
  layoutType: LayoutType;
  /** Display name of the preset at save time. */
  layoutName: string;
  /** Scheduled day, in ISO form. */
  scheduleDate: string;
  /** Scheduled time as "HH:mm". */
  time: string;
  /** Filled slots only, each with its story and optional layout override. */
  slots: Array<{ slotId: string; story: IStory; cardLayout?: CardLayout }>;
}


const getSpOrThrow = (): SPFI => {
  const sp = getSp();
  if (!sp || !sp.web) {
    throw new Error(
      "SharePoint not initialized. Call initializeSharePoint(context) before using the Schedule Stories service."
    );
  }
  return sp;
};

const isValidLayoutType = (value: string | undefined): value is LayoutType =>
  !!value && VALID_LAYOUT_TYPES.indexOf(value as LayoutType) !== -1;

const coerceStatus = (value: string | undefined): ScheduleStatus =>
  value && VALID_STATUSES.indexOf(value as ScheduleStatus) !== -1
    ? (value as ScheduleStatus)
    : DEFAULT_STATUS;

/** "HH:mm" or a safe default. */
const normalizeTime = (time: string | undefined): string =>
  time && /^\d{2}:\d{2}$/.test(time) ? time : "09:00";

/**
 * Hour of day the schedule's calendar-day columns are anchored at.
 *
 * A scheduled day is a civil date, not an instant, but SharePoint stores it in
 * a DateTime column and normalises that to UTC. Anchoring at midday means the
 * stored instant can absorb up to twelve hours of skew in either direction
 * without crossing a day boundary, so the calendar day survives the round trip
 * from any timezone. Anchoring at midnight — the obvious choice — is exactly
 * the one that slips a day for every negative UTC offset.
 */
const DAY_ANCHOR_HOUR = "12:00:00";

/**
 * Canonical "YYYY-MM-DD" for a Date.
 *
 * Reads **local** parts, because a scheduled day is a civil date and every
 * consumer displays it with local getters (`toLocaleDateString`, `getDate`).
 * Reading UTC parts here would report a different day than the UI shows for
 * any viewer west of Greenwich.
 *
 * @param d - Date whose calendar day is wanted.
 * @returns The calendar day as `YYYY-MM-DD`.
 */
export const toDateKey = (d: Date): string => {
  const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  const y = d.getFullYear();
  const m = pad2(d.getMonth() + 1);
  const day = pad2(d.getDate());
  return `${y}-${m}-${day}`;
};

/**
 * Reconstructs the app's Date from a "YYYY-MM-DD" key.
 *
 * Built from local parts rather than parsed, since `new Date("YYYY-MM-DD")`
 * is defined to parse as UTC midnight — which is the drift this module exists
 * to avoid. The result is local midnight, so `toDateKey` returns the same key
 * it was given, for every viewer.
 *
 * @param key - Calendar day as `YYYY-MM-DD`.
 * @returns Local midnight on that day.
 */
export const fromDateKey = (key: string): Date => {
  const [y, m, d] = key.split("-").map((part) => parseInt(part, 10));
  return new Date(y, m - 1, d);
};


/** Build the BoardStateJson payload from an app group. */
export const serializeBoardState = (group: INewScheduledGroup): string => {
  const layoutName = group.layoutName || getLayoutConfig(group.layoutType).name;
  const slots: Array<{ slotId: string; story: IStory; cardLayout?: CardLayout }> = Object.keys(group.slotStories)
    .map((slotId) => {
      const story = group.slotStories[slotId];
      const cardLayout = group.slotLayoutPreferences ? group.slotLayoutPreferences[slotId] : undefined;
      return { slotId, story, cardLayout } as { slotId: string; story: IStory | undefined; cardLayout?: CardLayout };
    })
    .filter((entry): entry is { slotId: string; story: IStory; cardLayout?: CardLayout } => !!entry.story);

  const payload: IBoardStateJson = {
    layoutType: group.layoutType,
    layoutName,
    scheduleDate: toDateKey(group.date),
    time: normalizeTime(group.time),
    slots,
  };
  return JSON.stringify(payload);
};

/** Parse BoardStateJson back into a slotId -> story map. Returns null if unusable. */
const parseBoardState = (json: string | undefined): IBoardStateJson | undefined => {
  if (!json) return undefined;
  try {
    const parsed = JSON.parse(json) as Partial<IBoardStateJson>;
    if (!parsed || !Array.isArray(parsed.slots)) return undefined;
    return {
      layoutType: isValidLayoutType(parsed.layoutType) ? parsed.layoutType : DEFAULT_LAYOUT_TYPE,
      layoutName: parsed.layoutName || "",
      scheduleDate: parsed.scheduleDate || "",
      time: normalizeTime(parsed.time),
      slots: parsed.slots.filter(
        (s): s is { slotId: string; story: IStory; cardLayout?: CardLayout } => !!s && !!s.slotId && !!s.story
      ),
    };
  } catch (err) {
    console.error("Schedule Stories: failed to parse BoardStateJson", err);
    return undefined;
  }
};

const boardStateToSlotMap = (state: IBoardStateJson): SlotStoryMap => {
  const map: SlotStoryMap = {};
  state.slots.forEach(({ slotId, story }) => {
    map[slotId] = story;
  });
  return map;
};

/**
 * Map a raw SharePoint row to a domain record. Returns null for rows that can't
 * be reconstructed (missing/corrupt board state) so one bad row never breaks the
 * whole load.
 */
const mapRawToRecord = (raw: IRawScheduleItem): IScheduledGroupRecord | undefined => {
  const state = parseBoardState(raw.BoardStateJson);
  if (!state) {
    console.warn(`Schedule Stories: skipping item ${raw.Id} — unreadable BoardStateJson.`);
    return undefined;
  }

  // BoardStateJson carries the calendar day as a plain string, so it is the
  // only source that cannot have drifted. The SharePoint columns are instants
  // and are consulted only when it is absent — rows written before the field
  // existed, or edited outside the app. Those are read with local parts to
  // match how the day was anchored on write; see DAY_ANCHOR_HOUR.
  const dateKey =
    state.scheduleDate ||
    (raw.ScheduleDate ? toDateKey(new Date(raw.ScheduleDate)) : "") ||
    (raw.ScheduledDateTime ? toDateKey(new Date(raw.ScheduledDateTime)) : "");

  if (!dateKey) {
    console.warn(`Schedule Stories: skipping item ${raw.Id} — no resolvable date.`);
    return undefined;
  }

  const layoutType = isValidLayoutType(state.layoutType)
    ? state.layoutType
    : isValidLayoutType(raw.LayoutType)
      ? raw.LayoutType
      : DEFAULT_LAYOUT_TYPE;

  const slotLayoutPreferences: Record<string, CardLayout> = {};
  state.slots.forEach(({ slotId, cardLayout }) => {
    if (cardLayout) {
      slotLayoutPreferences[slotId] = cardLayout;
    }
  });

  return {
    id: raw.GroupId || `sp-${raw.Id}`,
    spId: raw.Id,
    date: fromDateKey(dateKey),
    time: normalizeTime(state.time),
    slotStories: boardStateToSlotMap(state),
    slotLayoutPreferences,
    layoutType,
    layoutName: state.layoutName || getLayoutConfig(layoutType).name,
    status: coerceStatus(raw.Status),
    createdAt: raw.Created ? new Date(raw.Created) : new Date(),
  };
};


// eslint-disable-next-line @typescript-eslint/no-explicit-any
const buildItemPayload = (group: INewScheduledGroup): any => {
  const dateKey = toDateKey(group.date);
  const time = normalizeTime(group.time);
  const layoutName = group.layoutName || getLayoutConfig(group.layoutType).name;

  return {
    Title: `${dateKey} · ${layoutName}`,
    GroupId: group.id,
    // Day-only column, anchored at midday so timezone normalisation cannot
    // push it onto the neighbouring day.
    ScheduleDate: `${dateKey}T${DAY_ANCHOR_HOUR}`,
    // Genuine civil date-time — carries the scheduled time and backs the
    // soonest-first ordering, so it keeps the real hour.
    ScheduledDateTime: `${dateKey}T${time}:00`,
    LayoutType: group.layoutType,
    BoardStateJson: serializeBoardState(group),
    Status: group.status || DEFAULT_STATUS,
  };
};


/**
 * Fetch all scheduled groups, soonest-scheduled first. Rows with corrupt board
 * state are skipped (logged) rather than throwing, so a single bad item can't
 * blank the whole schedule panel.
 */
export const getScheduledGroups = async (): Promise<IScheduledGroupRecord[]> => {
  const sp = getSpOrThrow();
  try {
    const items: IRawScheduleItem[] = await sp.web.lists
      .getByTitle(SCHEDULE_LIST_NAME)
      .items.select(
        "Id",
        "Title",
        "GroupId",
        "ScheduleDate",
        "ScheduledDateTime",
        "LayoutType",
        "BoardStateJson",
        "Status",
        "Created"
      )
      .orderBy("ScheduledDateTime", true)
      .top(200)();

    return items
      .map(mapRawToRecord)
      .filter((r): r is IScheduledGroupRecord => r !== undefined);
  } catch (error) {
    console.error("Schedule Stories: failed to fetch groups", error);
    throw error;
  }
};

/**
 * Create a new scheduled group. Returns the persisted record (with spId).
 */
export const createScheduledGroup = async (
  group: INewScheduledGroup
): Promise<IScheduledGroupRecord> => {
  const sp = getSpOrThrow();
  try {
    const payload = buildItemPayload(group);
    const result = await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).items.add(payload);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data: any = result.data;

    return {
      id: group.id,
      spId: data.Id,
      date: fromDateKey(toDateKey(group.date)),
      time: normalizeTime(group.time),
      slotStories: { ...group.slotStories },
      layoutType: group.layoutType,
      layoutName: group.layoutName || getLayoutConfig(group.layoutType).name,
      status: group.status || DEFAULT_STATUS,
      createdAt: data.Created ? new Date(data.Created) : new Date(),
    };
  } catch (error) {
    console.error("Schedule Stories: failed to create group", error);
    throw error;
  }
};

/**
 * Update the board and/or date/time of an existing group (used by reschedule).
 * Only the provided fields are written.
 */
export const updateScheduledGroup = async (
  spId: number,
  changes: Partial<Pick<INewScheduledGroup, "date" | "time" | "slotStories" | "layoutType" | "layoutName" | "status" | "slotLayoutPreferences">>
): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {};

    const hasDate = changes.date instanceof Date;
    const hasTime = typeof changes.time === "string";
    const hasBoard = !!changes.slotStories;
    const hasLayout = !!changes.layoutType;
    const hasPrefs = !!changes.slotLayoutPreferences;

    if (hasDate || hasTime) {
      const existing = await sp.web.lists
         .getByTitle(SCHEDULE_LIST_NAME)
         .items.getById(spId)
         .select("ScheduleDate", "ScheduledDateTime", "BoardStateJson")();

      const prevState = parseBoardState(existing.BoardStateJson);
      const dateKey = hasDate
        ? toDateKey(changes.date as Date)
        : prevState?.scheduleDate || toDateKey(new Date(existing.ScheduledDateTime));
      const time = hasTime ? normalizeTime(changes.time) : normalizeTime(prevState?.time);

      payload.ScheduleDate = `${dateKey}T${DAY_ANCHOR_HOUR}`;
      payload.ScheduledDateTime = `${dateKey}T${time}:00`;
    }

    if (hasBoard || hasDate || hasTime || hasLayout || hasPrefs) {
      const existing = await sp.web.lists
        .getByTitle(SCHEDULE_LIST_NAME)
        .items.getById(spId)
        .select("GroupId", "BoardStateJson")();
      const prevState = parseBoardState(existing.BoardStateJson);

      const layoutType = hasLayout
        ? (changes.layoutType as LayoutType)
        : prevState?.layoutType || DEFAULT_LAYOUT_TYPE;

      const rebuilt: INewScheduledGroup = {
        id: existing.GroupId || `sp-${spId}`,
        date: hasDate
          ? (changes.date as Date)
          : fromDateKey(prevState?.scheduleDate || toDateKey(new Date())),
        time: hasTime ? normalizeTime(changes.time) : normalizeTime(prevState?.time),
        slotStories: hasBoard
          ? (changes.slotStories as SlotStoryMap)
          : prevState
            ? boardStateToSlotMap(prevState)
            : {},
        slotLayoutPreferences: hasPrefs
          ? changes.slotLayoutPreferences
          : prevState
            ? (() => {
                const prefs: Record<string, CardLayout> = {};
                prevState.slots.forEach(({ slotId, cardLayout }) => {
                  if (cardLayout) prefs[slotId] = cardLayout;
                });
                return prefs;
              })()
            : undefined,
        layoutType,
        layoutName: changes.layoutName || getLayoutConfig(layoutType).name,
      };

      payload.BoardStateJson = serializeBoardState(rebuilt);
      payload.LayoutType = layoutType;
      payload.Title = `${toDateKey(rebuilt.date)} · ${rebuilt.layoutName || getLayoutConfig(layoutType).name}`;
    }

    if (changes.status) {
      payload.Status = changes.status;
    }

    if (Object.keys(payload).length === 0) return;

    await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).items.getById(spId).update(payload);
  } catch (error) {
    console.error(`Schedule Stories: failed to update group ${spId}`, error);
    throw error;
  }
};

/** Convenience wrapper for the reschedule flow. */
export const rescheduleGroup = async (
  spId: number,
  date: Date,
  time: string,
  slotStories?: SlotStoryMap,
  slotLayoutPreferences?: Record<string, CardLayout>
): Promise<void> => {
  await updateScheduledGroup(spId, { date, time, slotStories, slotLayoutPreferences });
};

/** Set a group's lifecycle status (e.g. mark Published/Cancelled/Expired). */
export const setGroupStatus = async (spId: number, status: ScheduleStatus): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).items.getById(spId).update({ Status: status });
  } catch (error) {
    console.error(`Schedule Stories: failed to set status for ${spId}`, error);
    throw error;
  }
};

/** Delete by SharePoint item id. */
export const deleteScheduledGroup = async (spId: number): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).items.getById(spId).delete();
  } catch (error) {
    console.error(`Schedule Stories: failed to delete group ${spId}`, error);
    throw error;
  }
};

/**
 * Delete by the app's GroupId (when the caller only has the app id, not the
 * SharePoint item id). No-op-safe: if nothing matches, it resolves quietly.
 */
export const deleteScheduledGroupByGroupId = async (groupId: string): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    const matches: Array<{ Id: number }> = await sp.web.lists
      .getByTitle(SCHEDULE_LIST_NAME)
      .items.select("Id")
      .filter(`GroupId eq '${groupId.replace(/'/g, "''")}'`)
      .top(10)();

    for (const m of matches) {
      await sp.web.lists.getByTitle(SCHEDULE_LIST_NAME).items.getById(m.Id).delete();
    }
  } catch (error) {
    console.error(`Schedule Stories: failed to delete group by GroupId ${groupId}`, error);
    throw error;
  }
};


/**
 * Is a group already scheduled for the given date? Pure helper over records you
 * already fetched — mirrors the app's date-only "one group per date" rule.
 * Pass `excludeId` to ignore a group being rescheduled onto its own date.
 */
export const isDateScheduled = (
  records: IScheduledGroupRecord[],
  date: Date,
  excludeId?: string
): boolean => {
  const key = toDateKey(date);
  return records.some((r) => r.id !== excludeId && toDateKey(r.date) === key);
};

/** Has the group cap been reached? */
export const isAtGroupLimit = (records: IScheduledGroupRecord[]): boolean =>
  records.length >= MAX_SCHEDULED_GROUPS;


/**
 * Ensure the "Schedule Stories" list and its columns exist. Safe to call
 * repeatedly — each field add is guarded, so an already-present column is
 * ignored. You created the list manually, so this is only a safety net for
 * fresh environments.
 */
export const ensureScheduleStoriesList = async (): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    const lists: Array<{ Title: string }> = await sp.web.lists();
    const exists = lists.some((l) => l.Title === SCHEDULE_LIST_NAME);
    if (!exists) {
      await sp.web.lists.add(SCHEDULE_LIST_NAME);
    }
    const list = sp.web.lists.getByTitle(SCHEDULE_LIST_NAME);

    const addSafely = async (fn: () => Promise<unknown>): Promise<void> => {
      try {
        await fn();
      } catch {
      }
    };

    await addSafely(() => list.fields.addText("GroupId"));
    await addSafely(() => list.fields.addDateTime("ScheduleDate"));
    await addSafely(() => list.fields.addDateTime("ScheduledDateTime"));
    await addSafely(() => list.fields.addChoice("LayoutType", { Choices: VALID_LAYOUT_TYPES }));
    await addSafely(() => list.fields.addMultilineText("BoardStateJson"));
    await addSafely(() => list.fields.addChoice("Status", { Choices: VALID_STATUSES }));
  } catch (error) {
    console.error("Schedule Stories: ensureScheduleStoriesList failed", error);
    throw error;
  }
};
