import {
  MAX_SCHEDULED_GROUPS,
  createScheduledGroup,
  deleteScheduledGroup,
  deleteScheduledGroupByGroupId,
  ensureScheduleStoriesList,
  fromDateKey,
  getScheduledGroups,
  isAtGroupLimit,
  isDateScheduled,
  rescheduleGroup,
  setGroupStatus,
  serializeBoardState,
  toDateKey,
  toScheduledMoment,
  updateScheduledGroup,
} from '../scheduleStoriesService';
import * as availableStoriesService from '../availableStoriesService';
import { mockStories } from '../../components/__tests__/mockData';

jest.mock('../availableStoriesService', () => ({
  getSp: jest.fn(),
}));

describe('scheduleStoriesService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('round-trips date keys in local time', () => {
    const original = new Date(2026, 6, 30, 18, 22, 0, 0);
    const key = toDateKey(original);
    const rebuilt = fromDateKey(key);

    expect(key).toBe('2026-07-30');
    expect(toDateKey(rebuilt)).toBe('2026-07-30');
    expect(rebuilt.getHours()).toBe(0);
  });

  it('serializes only filled slots and normalizes time', () => {
    const json = serializeBoardState({
      id: 'scheduled-1',
      date: new Date(2026, 6, 30),
      time: 'bad',
      slotStories: {
        'slot-1': mockStories[0],
        'slot-2': undefined,
      },
      layoutType: 'connectHomepage',
    });

    const parsed = JSON.parse(json);
    expect(parsed.layoutType).toBe('connectHomepage');
    expect(parsed.time).toBe('09:00');
    expect(parsed.slots).toHaveLength(1);
    expect(parsed.slots[0].slotId).toBe('slot-1');
  });

  it('checks date conflicts and group limits', () => {
    const records = [
      {
        id: 'scheduled-1',
        spId: 11,
        date: new Date(2026, 6, 30),
      },
      {
        id: 'scheduled-2',
        spId: 12,
        date: new Date(2026, 6, 31),
      },
    ] as any;

    expect(isDateScheduled(records, new Date(2026, 6, 30))).toBe(true);
    expect(isDateScheduled(records, new Date(2026, 6, 30), 'scheduled-1')).toBe(false);

    const atLimit = Array.from({ length: MAX_SCHEDULED_GROUPS }, (_, i) => ({ id: `${i}` }));
    expect(isAtGroupLimit(atLimit as any)).toBe(true);
    expect(isAtGroupLimit(records as any)).toBe(false);
  });

  it('maps scheduled groups from SharePoint rows and skips corrupt rows', async () => {
    const itemsCall = jest.fn().mockResolvedValue([
      {
        Id: 10,
        GroupId: 'scheduled-10',
        LayoutType: 'connectHomepage',
        BoardStateJson: JSON.stringify({
          layoutType: 'connectHomepage',
          layoutName: 'Connect Homepage',
          scheduleDate: '2026-07-30',
          time: '14:30',
          slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        }),
        Status: 'Scheduled',
        Created: '2026-07-01T10:00:00.000Z',
      },
      {
        Id: 11,
        GroupId: 'scheduled-11',
        BoardStateJson: '{not-json}',
      },
    ]);

    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      orderBy: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(itemsCall),
    };

    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;

    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    const rows = await getScheduledGroups();
    expect(rows).toHaveLength(1);
    expect(rows[0].spId).toBe(10);
    expect(rows[0].time).toBe('14:30');
    expect(rows[0].slotStories['slot-1']?.title).toBe('Test Story 1');
  });

  describe('timezone handling', () => {
    /** Wires `getSp` to a rows-returning mock and returns the mapped groups. */
    const readRows = async (rows: unknown[]) => {
      const itemsApi = {
        select: jest.fn().mockReturnThis(),
        orderBy: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(rows)),
      };
      (availableStoriesService.getSp as jest.Mock).mockReturnValue({
        web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: itemsApi }) } },
      } as any);

      return getScheduledGroups();
    };

    const boardState = (extra: Record<string, unknown> = {}) =>
      JSON.stringify({
        layoutType: 'connectHomepage',
        layoutName: 'Connect Homepage',
        scheduleDate: '2026-07-30',
        time: '14:30',
        slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        ...extra,
      });

    it('turns an author wall clock into an instant, with zone and offset', () => {
      const moment = toScheduledMoment('2026-07-30', '14:30');
      const local = new Date(2026, 6, 30, 14, 30, 0, 0);

      // The instant is whatever 14:30 on that day means where this runs.
      expect(moment.utcIso).toBe(local.toISOString());
      expect(moment.utcIso.endsWith('Z')).toBe(true);
      expect(moment.offsetMinutes).toBe(-local.getTimezoneOffset());
      expect(typeof moment.timeZone).toBe('string');
    });

    it('writes the instant, zone and offset when a group is created', async () => {
      const add = jest.fn().mockResolvedValue({ data: { Id: 42, Created: '2026-07-01T10:00:00.000Z' } });
      (availableStoriesService.getSp as jest.Mock).mockReturnValue({
        web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: { add } }) } },
      } as any);

      const saved = await createScheduledGroup({
        id: 'scheduled-1',
        date: new Date(2026, 6, 30),
        time: '14:30',
        slotStories: { 'slot-1': mockStories[0] },
        layoutType: 'connectHomepage',
      });

      const payload = add.mock.calls[0][0];
      const expected = toScheduledMoment('2026-07-30', '14:30');
      expect(payload.ScheduledDateTimeUtc).toBe(expected.utcIso);
      expect(payload.ScheduledTimeZone).toBe(expected.timeZone);
      expect(payload.ScheduledOffsetMinutes).toBe(expected.offsetMinutes);
      // Mirrored into the JSON as well, so a disturbed column can be recovered.
      expect(JSON.parse(payload.BoardStateJson).scheduledAtUtc).toBe(expected.utcIso);
      expect(saved.scheduledAtUtc.toISOString()).toBe(expected.utcIso);
    });

    it('reads the stored instant back as the viewer local day and time', async () => {
      // 03:30Z — a deliberately different wall clock from the JSON's civil
      // values, so the column has to be what wins.
      const instant = '2026-07-30T03:30:00Z';
      const rows = await readRows([
        {
          Id: 10,
          GroupId: 'scheduled-10',
          ScheduledDateTimeUtc: instant,
          ScheduledTimeZone: 'Asia/Kolkata',
          ScheduledOffsetMinutes: 330,
          BoardStateJson: boardState(),
          Status: 'Scheduled',
        },
      ]);

      const local = new Date(instant);
      expect(rows).toHaveLength(1);
      expect(rows[0].scheduledAtUtc.toISOString()).toBe(local.toISOString());
      expect(rows[0].authorTimeZone).toBe('Asia/Kolkata');
      expect(rows[0].authorOffsetMinutes).toBe(330);
      // Whatever zone this runs in, the civil parts describe that same instant.
      expect(rows[0].date.getDate()).toBe(local.getDate());
      expect(rows[0].time).toBe(
        `${`0${local.getHours()}`.slice(-2)}:${`0${local.getMinutes()}`.slice(-2)}`
      );
    });

    it('falls back to the JSON mirror when the column is missing', async () => {
      const instant = '2026-07-30T03:30:00Z';
      const rows = await readRows([
        {
          Id: 11,
          GroupId: 'scheduled-11',
          BoardStateJson: boardState({ scheduledAtUtc: instant, timeZone: 'Europe/London', offsetMinutes: 60 }),
        },
      ]);

      expect(rows[0].scheduledAtUtc.toISOString()).toBe(new Date(instant).toISOString());
      expect(rows[0].authorTimeZone).toBe('Europe/London');
    });

    it('reads legacy rows with no instant as local civil time, unchanged', async () => {
      const rows = await readRows([
        {
          Id: 12,
          GroupId: 'scheduled-12',
          BoardStateJson: boardState(),
        },
      ]);

      // Pre-migration behaviour: the stored wall clock is shown as-is.
      expect(rows[0].time).toBe('14:30');
      expect(toDateKey(rows[0].date)).toBe('2026-07-30');
      expect(rows[0].scheduledAtUtc.toISOString()).toBe(new Date(2026, 6, 30, 14, 30).toISOString());
    });
  });

  it('creates a scheduled group and returns persisted shape', async () => {
    const add = jest.fn().mockResolvedValue({
      data: {
        Id: 99,
        Created: '2026-07-20T08:00:00.000Z',
      },
    });

    const itemsApi = { add };
    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    const result = await createScheduledGroup({
      id: 'scheduled-99',
      date: new Date(2026, 6, 30),
      time: '16:45',
      slotStories: { 'slot-1': mockStories[0] },
      slotLayoutPreferences: { 'slot-1': 'thumbnail-text' },
      layoutType: 'connectHomepage',
    });

    expect(add).toHaveBeenCalled();
    expect(result.id).toBe('scheduled-99');
    expect(result.spId).toBe(99);
    expect(result.time).toBe('16:45');
    expect(result.layoutType).toBe('connectHomepage');
    // The caller holds this record until the next load, and editing the group
    // in the meantime writes its preferences back — so they must survive here.
    expect(result.slotLayoutPreferences).toEqual({ 'slot-1': 'thumbnail-text' });
  });

  it('returns an empty preference map when a group is created without one', async () => {
    const add = jest.fn().mockResolvedValue({ data: { Id: 100 } });
    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: { add } }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    const result = await createScheduledGroup({
      id: 'scheduled-100',
      date: new Date(2026, 6, 31),
      time: '09:00',
      slotStories: { 'slot-1': mockStories[0] },
      layoutType: 'connectHomepage',
    });

    expect(result.slotLayoutPreferences).toEqual({});
  });

  it('deletes scheduled groups by GroupId (including apostrophes)', async () => {
    const deleteA = jest.fn().mockResolvedValue(undefined);
    const deleteB = jest.fn().mockResolvedValue(undefined);

    const itemsCall = jest.fn().mockResolvedValue([{ Id: 5 }, { Id: 7 }]);
    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(itemsCall),
      getById: jest.fn((id: number) => ({
        delete: id === 5 ? deleteA : deleteB,
      })),
    };

    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    await deleteScheduledGroupByGroupId("group-'x'");

    expect(itemsApi.filter).toHaveBeenCalledWith("GroupId eq 'group-''x'''"
    );
    expect(deleteA).toHaveBeenCalled();
    expect(deleteB).toHaveBeenCalled();
  });

  it('updates an existing scheduled group with date, time, layout, and board changes', async () => {
    const update = jest.fn().mockResolvedValue(undefined);
    const selectInvoker = jest
      .fn()
      .mockResolvedValueOnce({
        ScheduleDate: '2026-07-30T12:00:00',
        ScheduledDateTime: '2026-07-30T09:00:00',
        BoardStateJson: JSON.stringify({
          layoutType: 'connectHomepage',
          layoutName: 'Connect Homepage',
          scheduleDate: '2026-07-30',
          time: '09:00',
          slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        }),
      })
      .mockResolvedValueOnce({
        GroupId: 'scheduled-22',
        BoardStateJson: JSON.stringify({
          layoutType: 'connectHomepage',
          layoutName: 'Connect Homepage',
          scheduleDate: '2026-07-30',
          time: '09:00',
          slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        }),
      });

    const itemsApi = {
      getById: jest.fn().mockReturnValue({
        select: jest.fn(() => selectInvoker),
        update,
      }),
    };

    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    await updateScheduledGroup(22, {
      date: new Date(2026, 7, 1),
      time: '15:30',
      slotStories: { 'slot-2': mockStories[1] },
      layoutType: 'general',
      layoutName: 'General',
      status: 'Published',
    });

    expect(update).toHaveBeenCalled();
    const payload = update.mock.calls[0][0];
    expect(payload.ScheduleDate).toContain('2026-08-01');
    expect(payload.ScheduledDateTime).toContain('2026-08-01T15:30:00');
    expect(payload.LayoutType).toBe('general');
    expect(payload.Title).toContain('2026-08-01');
    expect(payload.Status).toBe('Published');
    expect(payload.BoardStateJson).toContain('slot-2');
  });

  it('reschedules through updateScheduledGroup wrapper', async () => {
    const update = jest.fn().mockResolvedValue(undefined);
    const selectInvoker = jest
      .fn()
      .mockResolvedValueOnce({
        ScheduleDate: '2026-07-30T12:00:00',
        ScheduledDateTime: '2026-07-30T09:00:00',
        BoardStateJson: JSON.stringify({
          layoutType: 'connectHomepage',
          layoutName: 'Connect Homepage',
          scheduleDate: '2026-07-30',
          time: '09:00',
          slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        }),
      })
      .mockResolvedValueOnce({
        GroupId: 'scheduled-30',
        BoardStateJson: JSON.stringify({
          layoutType: 'connectHomepage',
          layoutName: 'Connect Homepage',
          scheduleDate: '2026-07-30',
          time: '09:00',
          slots: [{ slotId: 'slot-1', story: mockStories[0] }],
        }),
      });

    const itemsApi = {
      getById: jest.fn().mockReturnValue({
        select: jest.fn(() => selectInvoker),
        update,
      }),
    };

    (availableStoriesService.getSp as jest.Mock).mockReturnValue({
      web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: itemsApi }) } },
    });

    await rescheduleGroup(30, new Date(2026, 7, 2), '11:15');
    expect(update).toHaveBeenCalled();
  });

  it('sets status and deletes by SharePoint id', async () => {
    const update = jest.fn().mockResolvedValue(undefined);
    const del = jest.fn().mockResolvedValue(undefined);
    const itemsApi = {
      getById: jest.fn((id: number) => ({
        update: id === 77 ? update : jest.fn(),
        delete: id === 88 ? del : jest.fn(),
      })),
    };

    (availableStoriesService.getSp as jest.Mock).mockReturnValue({
      web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: itemsApi }) } },
    });

    await setGroupStatus(77, 'Cancelled');
    await deleteScheduledGroup(88);

    expect(update).toHaveBeenCalledWith({ Status: 'Cancelled' });
    expect(del).toHaveBeenCalled();
  });

  it('ensures schedule list and fields exist', async () => {
    const addList = jest.fn().mockResolvedValue(undefined);
    const fields = {
      addText: jest.fn().mockResolvedValue(undefined),
      addDateTime: jest.fn().mockResolvedValue(undefined),
      addChoice: jest.fn().mockResolvedValue(undefined),
      addMultilineText: jest.fn().mockResolvedValue(undefined),
    };

    const listsFn: any = jest.fn().mockResolvedValue([]);
    listsFn.add = addList;
    listsFn.getByTitle = jest.fn().mockReturnValue({ fields });

    (availableStoriesService.getSp as jest.Mock).mockReturnValue({
      web: { lists: listsFn },
    });

    await ensureScheduleStoriesList();

    expect(addList).toHaveBeenCalled();
    expect(fields.addText).toHaveBeenCalledWith('GroupId');
    expect(fields.addDateTime).toHaveBeenCalledWith('ScheduleDate');
    expect(fields.addDateTime).toHaveBeenCalledWith('ScheduledDateTime');
    expect(fields.addMultilineText).toHaveBeenCalledWith('BoardStateJson');
  });
});
