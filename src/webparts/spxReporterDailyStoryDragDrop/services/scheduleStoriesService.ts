// =============================================================
// Schedule Stories service
// -------------------------------------------------------------
// Persists the "scheduled board groups" (currently React-only state in
// StoryDragDrop.tsx) to the "Schedule Stories" SharePoint list.
//
// List columns this service targets:
//   Title              Single line   - human label (auto-generated)
//   GroupId            Single line   - the app id ("scheduled-<ts>"), our key
//   ScheduleDate       Date & Time   - date bucket (00:00), for filter/uniqueness
//   ScheduledDateTime  Date & Time   - full instant, for sorting / "is it due"
//   LayoutType         Choice        - reporterDaily | general | highlight
//   BoardStateJson     Multi-line    - the exact board snapshot (source of truth)
//   Status             Choice        - Scheduled | Published | Cancelled | Expired
//

import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { getSp } from "./sharePointService";
import { IStory, LayoutType, SlotStoryMap } from "../components/types";
import { getLayoutConfig } from "../components/layouts/layoutConfig";

// ---- Constants ----------------------------------------------
export const SCHEDULE_LIST_NAME = "Schedule Stories";

// Reference cap: at most 10 scheduled groups.
export const MAX_SCHEDULED_GROUPS = 10;

const VALID_LAYOUT_TYPES: LayoutType[] = ["reporterDaily", "general", "highlight"];
const DEFAULT_LAYOUT_TYPE: LayoutType = "reporterDaily";

export type ScheduleStatus = "Scheduled" | "Published" | "Cancelled" | "Expired";
const VALID_STATUSES: ScheduleStatus[] = ["Scheduled", "Published", "Cancelled", "Expired"];
const DEFAULT_STATUS: ScheduleStatus = "Scheduled";

// ---- Domain types -------------------------------------------

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
  /** Scheduled calendar day (UTC-midnight, matching the app's Date origin). */
  date: Date;
  /** Scheduled time as "HH:mm". */
  time: string;
  /** slotId -> story snapshot. */
  slotStories: SlotStoryMap;
  layoutType: LayoutType;
  layoutName: string;
  status: ScheduleStatus;
  createdAt: Date;
}

/** Input for creating a new scheduled group (id supplied by the app). */
export interface INewScheduledGroup {
  id: string;
  date: Date;
  time: string;
  slotStories: SlotStoryMap;
  layoutType: LayoutType;
  layoutName?: string;
  status?: ScheduleStatus;
}

/** Raw shape of a "Schedule Stories" list item as read from SharePoint. */
interface IRawScheduleItem {
  Id: number;
  Title?: string;
  GroupId?: string;
  ScheduleDate?: string;
  ScheduledDateTime?: string;
  LayoutType?: string;
  BoardStateJson?: string;
  Status?: string;
  Created?: string;
}

/** Serialized board snapshot stored in BoardStateJson. */
interface IBoardStateJson {
  layoutType: LayoutType;
  layoutName: string;
  scheduleDate: string; // "YYYY-MM-DD" — canonical, timezone-proof
  time: string; // "HH:mm"
  slots: Array<{ slotId: string; story: IStory }>;
}

// ---- Small guards / helpers ---------------------------------

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
 * Canonical "YYYY-MM-DD" for a Date. Uses UTC parts because the app builds its
 * schedule date via `new Date("YYYY-MM-DD")` (parsed as UTC midnight); reading
 * it back with UTC parts recovers the intended calendar day with no drift.
 */
export const toDateKey = (d: Date): string => {
  const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  const y = d.getUTCFullYear();
  const m = pad2(d.getUTCMonth() + 1);
  const day = pad2(d.getUTCDate());
  return `${y}-${m}-${day}`;
};

/** Reconstruct the app's Date (UTC midnight) from a "YYYY-MM-DD" key. */
const dateKeyToDate = (key: string): Date => new Date(`${key}T00:00:00Z`);

// ---- (De)serialization --------------------------------------

/** Build the BoardStateJson payload from an app group. */
export const serializeBoardState = (group: INewScheduledGroup): string => {
  const layoutName = group.layoutName || getLayoutConfig(group.layoutType).name;
  const slots: Array<{ slotId: string; story: IStory }> = Object.keys(group.slotStories)
    .map((slotId) => ({ slotId, story: group.slotStories[slotId] }))
    // Drop empty slots — only persist real placements.
    .filter((entry): entry is { slotId: string; story: IStory } => !!entry.story);

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
        (s): s is { slotId: string; story: IStory } => !!s && !!s.slotId && !!s.story
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

  // Prefer the JSON (timezone-proof); fall back to the columns for old rows.
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

  return {
    id: raw.GroupId || `sp-${raw.Id}`, // fall back to a stable id if GroupId was cleared
    spId: raw.Id,
    date: dateKeyToDate(dateKey),
    time: normalizeTime(state.time),
    slotStories: boardStateToSlotMap(state),
    layoutType,
    layoutName: state.layoutName || getLayoutConfig(layoutType).name,
    status: coerceStatus(raw.Status),
    createdAt: raw.Created ? new Date(raw.Created) : new Date(),
  };
};

// ---- Write payload builder ----------------------------------

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const buildItemPayload = (group: INewScheduledGroup): any => {
  const dateKey = toDateKey(group.date);
  const time = normalizeTime(group.time);
  const layoutName = group.layoutName || getLayoutConfig(group.layoutType).name;

  return {
    Title: `${dateKey} · ${layoutName}`,
    GroupId: group.id,
    // Date bucket at local midnight (for filter / one-per-date uniqueness).
    ScheduleDate: `${dateKey}T00:00:00`,
    // Full instant in local time (for ordering and "is it due" checks).
    ScheduledDateTime: `${dateKey}T${time}:00`,
    LayoutType: group.layoutType,
    BoardStateJson: serializeBoardState(group),
    Status: group.status || DEFAULT_STATUS,
  };
};

// ---- Public API ---------------------------------------------

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
      date: dateKeyToDate(toDateKey(group.date)),
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
  changes: Partial<Pick<INewScheduledGroup, "date" | "time" | "slotStories" | "layoutType" | "layoutName" | "status">>
): Promise<void> => {
  const sp = getSpOrThrow();
  try {
    // Re-read is avoided: reschedule always passes date+time+board together in
    // this app, but we guard so a partial update stays consistent.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {};

    const hasDate = changes.date instanceof Date;
    const hasTime = typeof changes.time === "string";
    const hasBoard = !!changes.slotStories;
    const hasLayout = !!changes.layoutType;

    if (hasDate || hasTime) {
      // Need a full date+time to rewrite the datetime columns coherently.
      const existing = await sp.web.lists
        .getByTitle(SCHEDULE_LIST_NAME)
        .items.getById(spId)
        .select("ScheduleDate", "ScheduledDateTime", "BoardStateJson")();

      const prevState = parseBoardState(existing.BoardStateJson);
      const dateKey = hasDate
        ? toDateKey(changes.date as Date)
        : prevState?.scheduleDate || toDateKey(new Date(existing.ScheduledDateTime));
      const time = hasTime ? normalizeTime(changes.time) : normalizeTime(prevState?.time);

      payload.ScheduleDate = `${dateKey}T00:00:00`;
      payload.ScheduledDateTime = `${dateKey}T${time}:00`;
    }

    if (hasBoard || hasDate || hasTime || hasLayout) {
      // Rebuild the JSON snapshot so it stays the source of truth. Read current
      // state to fill any fields the caller didn't pass.
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
          : dateKeyToDate(prevState?.scheduleDate || toDateKey(new Date())),
        time: hasTime ? normalizeTime(changes.time) : normalizeTime(prevState?.time),
        slotStories: hasBoard
          ? (changes.slotStories as SlotStoryMap)
          : prevState
            ? boardStateToSlotMap(prevState)
            : {},
        layoutType,
        layoutName: changes.layoutName || getLayoutConfig(layoutType).name,
      };

      payload.BoardStateJson = serializeBoardState(rebuilt);
      payload.LayoutType = layoutType;
      // Keep the human-readable Title in sync with the (possibly changed) date
      // and layout — same format buildItemPayload uses on create.
      payload.Title = `${toDateKey(rebuilt.date)} · ${rebuilt.layoutName || getLayoutConfig(layoutType).name}`;
    }

    if (changes.status) {
      payload.Status = changes.status;
    }

    if (Object.keys(payload).length === 0) return; // nothing to do

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
  slotStories?: SlotStoryMap
): Promise<void> => {
  await updateScheduledGroup(spId, { date, time, slotStories });
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

// ---- Client-side rules (mirror the app's constraints) -------

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

// ---- Optional provisioning (self-heal missing list/columns) --

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
        /* column already exists — ignore */
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
