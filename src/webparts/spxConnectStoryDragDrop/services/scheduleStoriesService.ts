
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { getSp } from "./availableStoriesService";
import { IStory, LayoutType, SlotStoryMap, CardLayout } from "../components/types";
import { getLayoutConfig } from "../components/layouts/layoutConfig";

/** Default title of the SharePoint list holding scheduled groups. */
export const DEFAULT_SCHEDULED_STORIES_LIST_NAME = "ScheduledStories";

/** Backward-compatible alias used by existing tests/imports. */
export const SCHEDULE_LIST_NAME = DEFAULT_SCHEDULED_STORIES_LIST_NAME;

let scheduledStoriesListName = DEFAULT_SCHEDULED_STORIES_LIST_NAME;

/** Override the SharePoint list title used by the schedule service. */
export const setScheduledStoriesListName = (listName?: string): void => {
  const next = (listName || "").trim();
  scheduledStoriesListName = next || DEFAULT_SCHEDULED_STORIES_LIST_NAME;
};

/** Current SharePoint list title used by the schedule service. */
export const getScheduledStoriesListName = (): string => scheduledStoriesListName;

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
  /**
   * The absolute instant this group is scheduled for. `date` and `time` are
   * this same moment expressed in the *viewer's* zone; this is what to compare
   * and what a publishing job should fire on.
   */
  scheduledAtUtc: Date;
  /** IANA zone the group was scheduled from, when it was recorded. */
  authorTimeZone?: string;
  /** The author zone's offset from UTC, in minutes, at the scheduled instant. */
  authorOffsetMinutes?: number;
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
  /** Optional alternate id key some clients/projectors return. */
  ID?: number;
  /** Item title, used only for readability in the list view. */
  Title?: string;
  /** App-side group id. */
  GroupId?: string;
  /** Scheduled day, in ISO form. */
  ScheduleDate?: string;
  /** Scheduled day and time combined, in ISO form. Zone-less; legacy. */
  ScheduledDateTime?: string;
  /** The absolute scheduled instant, in ISO form ending `Z`. Authoritative. */
  ScheduledDateTimeUtc?: string;
  /** IANA zone the group was scheduled from. */
  ScheduledTimeZone?: string;
  /** The author zone's offset from UTC, in minutes, at the scheduled instant. */
  ScheduledOffsetMinutes?: number | null;
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
  /** Scheduled day, in the author's civil terms. */
  scheduleDate: string;
  /** Scheduled time as "HH:mm", in the author's civil terms. */
  time: string;
  /** The absolute scheduled instant, ISO ending `Z`. Absent on legacy rows. */
  scheduledAtUtc?: string;
  /** IANA zone the group was scheduled from. */
  timeZone?: string;
  /** The author zone's offset from UTC, in minutes, at that instant. */
  offsetMinutes?: number;
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
const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/**
 * The timezone-aware form a scheduled moment is persisted in.
 *
 * The instant is what actually fires and what sorts; the zone and offset record
 * *whose* wall clock produced it, so a viewer elsewhere can be shown both their
 * own time and the author's.
 */
export interface IScheduledMoment {
  /** The absolute instant, as an ISO string ending in `Z`. */
  utcIso: string;
  /** IANA zone the group was scheduled from, e.g. `Asia/Kolkata`. */
  timeZone: string;
  /** That zone's offset from UTC at that instant, in minutes. East is positive. */
  offsetMinutes: number;
}

/** The runtime's IANA zone, or "" where `Intl` cannot report one. */
const resolveTimeZone = (): string => {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "";
  } catch {
    return "";
  }
};

/**
 * Converts an author's civil day and time into a storable instant.
 *
 * The `Date` is built from **local** parts, so it carries the author's wall
 * clock; `toISOString` then resolves it to the one instant every timezone
 * agrees on.
 *
 * @param dateKey - Civil day as `YYYY-MM-DD`.
 * @param time - Civil time as `HH:mm`.
 */
export const toScheduledMoment = (dateKey: string, time: string): IScheduledMoment => {
  const [y, m, d] = dateKey.split("-").map((part) => parseInt(part, 10));
  const [hr, min] = normalizeTime(time).split(":").map((part) => parseInt(part, 10));
  const local = new Date(y, m - 1, d, hr, min, 0, 0);

  return {
    utcIso: local.toISOString(),
    timeZone: resolveTimeZone(),
    offsetMinutes: -local.getTimezoneOffset(),
  };
};

/**
 * Reads an instant back as civil parts in *the viewer's* zone.
 *
 * A `Date` is already an absolute instant, so the local getters are themselves
 * the conversion — no offset arithmetic, and correct for any zone.
 *
 * @param instant - The stored moment.
 */
export const toLocalParts = (instant: Date): { dateKey: string; time: string } => ({
  dateKey: toDateKey(instant),
  time: `${pad2(instant.getHours())}:${pad2(instant.getMinutes())}`,
});

export const toDateKey = (d: Date): string => {
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

  const dateKey = toDateKey(group.date);
  const time = normalizeTime(group.time);
  const moment = toScheduledMoment(dateKey, time);

  const payload: IBoardStateJson = {
    layoutType: group.layoutType,
    layoutName,
    scheduleDate: dateKey,
    time,
    // Mirrored here as well as in its own column: this JSON is the one field
    // that cannot be reshaped by SharePoint's own timezone handling.
    scheduledAtUtc: moment.utcIso,
    timeZone: moment.timeZone,
    offsetMinutes: moment.offsetMinutes,
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
      scheduledAtUtc: parsed.scheduledAtUtc,
      timeZone: parsed.timeZone,
      offsetMinutes: typeof parsed.offsetMinutes === "number" ? parsed.offsetMinutes : undefined,
      slots: parsed.slots.filter(
        (s): s is { slotId: string; story: IStory; cardLayout?: CardLayout } => !!s && !!s.slotId && !!s.story
      ),
    };
  } catch (err) {
    console.error("Schedule Stories: failed to parse BoardStateJson", err);
    return undefined;
  }
};

/** A usable Date from an ISO string, or undefined for missing/unparseable input. */
const parseInstant = (iso: string | undefined): Date | undefined => {
  if (!iso) return undefined;
  const parsed = new Date(iso);
  return isNaN(parsed.getTime()) ? undefined : parsed;
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
  const spId = typeof raw.Id === "number" ? raw.Id : raw.ID;
  if (typeof spId !== "number") {
    return undefined;
  }

  const state = parseBoardState(raw.BoardStateJson);
  if (!state) {
    console.warn(`Schedule Stories: skipping item ${spId} — unreadable BoardStateJson.`);
    return undefined;
  }

  // Resolution order for the moment, most trustworthy first:
  //  1. the UTC column — an unambiguous instant;
  //  2. the same value mirrored in BoardStateJson, for rows edited outside the
  //     app in a way that disturbed the column;
  //  3. legacy rows with no instant at all, whose civil day and time are read
  //     as the *viewer's* local wall clock, preserving pre-migration behaviour.
  const legacyDateKey =
    state.scheduleDate ||
    (raw.ScheduleDate ? toDateKey(new Date(raw.ScheduleDate)) : "") ||
    (raw.ScheduledDateTime ? toDateKey(new Date(raw.ScheduledDateTime)) : "");

  const storedInstant = parseInstant(raw.ScheduledDateTimeUtc) || parseInstant(state.scheduledAtUtc);
  const instant =
    storedInstant ||
    (legacyDateKey ? new Date(toScheduledMoment(legacyDateKey, state.time).utcIso) : undefined);

  if (!instant) {
    console.warn(`Schedule Stories: skipping item ${spId} — no resolvable date.`);
    return undefined;
  }

  // The conversion the whole change exists for: an instant in, the viewer's
  // own civil day and time out.
  const local = toLocalParts(instant);

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
    id: raw.GroupId || `sp-${spId}`,
    spId,
    date: fromDateKey(local.dateKey),
    time: local.time,
    slotStories: boardStateToSlotMap(state),
    slotLayoutPreferences,
    layoutType,
    layoutName: state.layoutName || getLayoutConfig(layoutType).name,
    status: coerceStatus(raw.Status),
    createdAt: raw.Created ? new Date(raw.Created) : new Date(),
    scheduledAtUtc: instant,
    authorTimeZone: raw.ScheduledTimeZone || state.timeZone || undefined,
    authorOffsetMinutes:
      typeof raw.ScheduledOffsetMinutes === "number" ? raw.ScheduledOffsetMinutes : state.offsetMinutes,
  };
};


// eslint-disable-next-line @typescript-eslint/no-explicit-any
const buildItemPayload = (group: INewScheduledGroup): any => {
  const dateKey = toDateKey(group.date);
  const time = normalizeTime(group.time);
  const moment = toScheduledMoment(dateKey, time);
  const layoutName = group.layoutName || getLayoutConfig(group.layoutType).name;

  return {
    Title: `${dateKey} · ${layoutName}`,
    GroupId: group.id,
    // Day-only column, anchored at midday so timezone normalisation cannot
    // push it onto the neighbouring day.
    ScheduleDate: `${dateKey}T${DAY_ANCHOR_HOUR}`,
    // Civil date-time, kept zone-less for anything still reading it.
    ScheduledDateTime: `${dateKey}T${time}:00`,
    // The authoritative instant, plus whose wall clock produced it. This is
    // what sorts, what compares, and what a publishing job should fire on.
    ScheduledDateTimeUtc: moment.utcIso,
    ScheduledTimeZone: moment.timeZone,
    ScheduledOffsetMinutes: moment.offsetMinutes,
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
      .getByTitle(getScheduledStoriesListName())
      .items.select(
        "Id",
        "Title",
        "GroupId",
        "ScheduleDate",
        "ScheduledDateTime",
        "ScheduledDateTimeUtc",
        "ScheduledTimeZone",
        "ScheduledOffsetMinutes",
        "LayoutType",
        "BoardStateJson",
        "Status",
        "Created"
      )
      // Ordered on the instant, so the sequence is the same in every timezone.
      .orderBy("ScheduledDateTimeUtc", true)
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
    const result = await sp.web.lists.getByTitle(getScheduledStoriesListName()).items.add(payload);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data: any = result.data;
    const createdSpId = typeof data.Id === "number" ? data.Id : data.ID;
    if (typeof createdSpId !== "number") {
      throw new Error("Schedule Stories: create did not return a SharePoint item id.");
    }

    // Same moment the payload was built from, so a freshly created group and
    // one read back on the next load carry identical fields.
    const moment = toScheduledMoment(toDateKey(group.date), normalizeTime(group.time));

    return {
      id: group.id,
      spId: createdSpId,
      date: fromDateKey(toDateKey(group.date)),
      time: normalizeTime(group.time),
      scheduledAtUtc: new Date(moment.utcIso),
      authorTimeZone: moment.timeZone || undefined,
      authorOffsetMinutes: moment.offsetMinutes,
      slotStories: { ...group.slotStories },
      // Always an object, as {@link mapRawToRecord} returns on reload. Omitting
      // it left the freshly created group looking like it had no card layouts,
      // and saving an edit in that state wrote the emptiness back to the list.
      slotLayoutPreferences: { ...(group.slotLayoutPreferences || {}) },
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
         .getByTitle(getScheduledStoriesListName())
         .items.getById(spId)
         .select("ScheduleDate", "ScheduledDateTime", "BoardStateJson")();

      const prevState = parseBoardState(existing.BoardStateJson);
      const dateKey = hasDate
        ? toDateKey(changes.date as Date)
        : prevState?.scheduleDate || toDateKey(new Date(existing.ScheduledDateTime));
      const time = hasTime ? normalizeTime(changes.time) : normalizeTime(prevState?.time);

      const moment = toScheduledMoment(dateKey, time);
      payload.ScheduleDate = `${dateKey}T${DAY_ANCHOR_HOUR}`;
      payload.ScheduledDateTime = `${dateKey}T${time}:00`;
      payload.ScheduledDateTimeUtc = moment.utcIso;
      payload.ScheduledTimeZone = moment.timeZone;
      payload.ScheduledOffsetMinutes = moment.offsetMinutes;
    }

    if (hasBoard || hasDate || hasTime || hasLayout || hasPrefs) {
      const existing = await sp.web.lists
        .getByTitle(getScheduledStoriesListName())
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

    await sp.web.lists.getByTitle(getScheduledStoriesListName()).items.getById(spId).update(payload);
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
    await sp.web.lists.getByTitle(getScheduledStoriesListName()).items.getById(spId).update({ Status: status });
  } catch (error) {
    console.error(`Schedule Stories: failed to set status for ${spId}`, error);
    throw error;
  }
};

/** Delete by SharePoint item id. */
export const deleteScheduledGroup = async (spId: number): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    await sp.web.lists.getByTitle(getScheduledStoriesListName()).items.getById(spId).delete();
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
      .getByTitle(getScheduledStoriesListName())
      .items.select("Id")
      .filter(`GroupId eq '${groupId.replace(/'/g, "''")}'`)
      .top(10)();

    for (const m of matches) {
      await sp.web.lists.getByTitle(getScheduledStoriesListName()).items.getById(m.Id).delete();
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
    const listName = getScheduledStoriesListName();
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

    await addSafely(() => list.fields.addText("GroupId"));
    await addSafely(() => list.fields.addDateTime("ScheduleDate"));
    await addSafely(() => list.fields.addDateTime("ScheduledDateTime"));
    await addSafely(() => list.fields.addDateTime("ScheduledDateTimeUtc"));
    await addSafely(() => list.fields.addText("ScheduledTimeZone"));
    await addSafely(() => list.fields.addNumber("ScheduledOffsetMinutes"));
    await addSafely(() => list.fields.addChoice("LayoutType", { Choices: VALID_LAYOUT_TYPES }));
    await addSafely(() => list.fields.addMultilineText("BoardStateJson"));
    await addSafely(() => list.fields.addChoice("Status", { Choices: VALID_STATUSES }));
  } catch (error) {
    console.error("Schedule Stories: ensureScheduleStoriesList failed", error);
    throw error;
  }
};
