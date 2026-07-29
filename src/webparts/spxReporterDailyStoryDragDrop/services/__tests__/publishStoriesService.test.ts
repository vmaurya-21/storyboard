import {
  deletePublishedBoard,
  ensurePublishStoriesList,
  getLiveBoard,
  getPublishedBoards,
  IS_LIVE_YES,
  publishBoard,
  unpublishBoard,
} from '../publishStoriesService';
import * as availableStoriesService from '../availableStoriesService';
import * as scheduleStoriesService from '../scheduleStoriesService';
import { mockStories } from '../../components/__tests__/mockData';

jest.mock('../availableStoriesService', () => ({
  getSp: jest.fn(),
}));

jest.mock('../scheduleStoriesService', () => ({
  ...jest.requireActual('../scheduleStoriesService'),
  setGroupStatus: jest.fn(),
  SCHEDULE_LIST_NAME: 'Schedule Stories',
}));

describe('publishStoriesService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('rejects publishing an empty board unless allowEmpty is true', async () => {
    (availableStoriesService.getSp as jest.Mock).mockReturnValue({ web: {} });

    await expect(
      publishBoard({
        layoutType: 'connectHomepage',
        slotStories: {},
      })
    ).rejects.toThrow('Cannot publish an empty board');
  });

  it('publishes board, supersedes prior live records, and marks source schedule', async () => {
    const updateLiveA = jest.fn().mockResolvedValue(undefined);
    const updateLiveB = jest.fn().mockResolvedValue(undefined);
    const add = jest.fn().mockResolvedValue({ data: { Id: 45 } });
    const deleteOld = jest.fn().mockResolvedValue(undefined);

    const queuedTopResults = [
      [{ Id: 1 }, { Id: 2 }],
      [{ Id: 45 }, { Id: 44 }, { Id: 43 }],
    ];

    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      orderBy: jest.fn().mockReturnThis(),
      top: jest.fn(() => jest.fn().mockResolvedValue(queuedTopResults.shift() || [])),
      add,
      getById: jest.fn((id: number) => ({
        update: id === 1 ? updateLiveA : id === 2 ? updateLiveB : jest.fn(),
        delete: id === 43 ? deleteOld : jest.fn().mockResolvedValue(undefined),
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

    const result = await publishBoard({
      layoutType: 'connectHomepage',
      slotStories: { 'slot-1': mockStories[0] },
      sourceScheduleSpId: 77,
    });

    expect(add).toHaveBeenCalled();
    expect(updateLiveA).toHaveBeenCalledWith({ IsLive: 'No' });
    expect(updateLiveB).toHaveBeenCalledWith({ IsLive: 'No' });
    expect(deleteOld).toHaveBeenCalled();
    expect(scheduleStoriesService.setGroupStatus).toHaveBeenCalledWith(77, 'Published');
    expect(result.spId).toBe(45);
    expect(result.isLive).toBe(true);
  });

  it('returns undefined when there is no live board', async () => {
    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      orderBy: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue([])),
    };

    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    const live = await getLiveBoard();
    expect(live).toBeUndefined();
  });

  it('maps the live board record when one exists', async () => {
    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      orderBy: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(
        jest.fn().mockResolvedValue([
          {
            Id: 200,
            LayoutType: 'connectHomepage',
            PublishedDateTime: '2026-07-30T09:30:00.000Z',
            BoardStateJson: JSON.stringify({
              layoutType: 'connectHomepage',
              layoutName: 'Connect Homepage',
              publishedAt: '2026-07-30T09:30:00.000Z',
              slots: [{ slotId: 'slot-1', story: mockStories[0] }],
            }),
            IsLive: IS_LIVE_YES,
          },
        ])
      ),
    };

    (availableStoriesService.getSp as jest.Mock).mockReturnValue({
      web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: itemsApi }) } },
    });

    const live = await getLiveBoard();
    expect(live?.spId).toBe(200);
    expect(live?.isLive).toBe(true);
    expect(live?.slotStories['slot-1']?.title).toBe('Test Story 1');
  });

  it('maps valid publish history and skips invalid rows', async () => {
    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      orderBy: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(
        jest.fn().mockResolvedValue([
          {
            Id: 100,
            PublishedDateTime: '2026-07-30T09:30:00.000Z',
            LayoutType: 'connectHomepage',
            BoardStateJson: JSON.stringify({
              layoutType: 'connectHomepage',
              layoutName: 'Connect Homepage',
              publishedAt: '2026-07-30T09:30:00.000Z',
              slots: [{ slotId: 'slot-1', story: mockStories[0] }],
            }),
            IsLive: IS_LIVE_YES,
          },
          {
            Id: 101,
            BoardStateJson: '{bad-json}',
          },
        ])
      ),
    };

    const sp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({ items: itemsApi }),
        },
      },
    } as any;
    (availableStoriesService.getSp as jest.Mock).mockReturnValue(sp);

    const boards = await getPublishedBoards();
    expect(boards).toHaveLength(1);
    expect(boards[0].spId).toBe(100);
    expect(boards[0].slotStories['slot-1']?.title).toBe('Test Story 1');
  });

  it('unpublishes and deletes publish rows', async () => {
    const updateLive = jest.fn().mockResolvedValue(undefined);
    const deleteRow = jest.fn().mockResolvedValue(undefined);

    const itemsApi = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue([{ Id: 5 }])),
      getById: jest.fn((id: number) => ({
        update: id === 5 ? updateLive : jest.fn(),
        delete: id === 9 ? deleteRow : jest.fn().mockResolvedValue(undefined),
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

    await unpublishBoard();
    await deletePublishedBoard(9);

    expect(updateLive).toHaveBeenCalledWith({ IsLive: 'No' });
    expect(deleteRow).toHaveBeenCalled();
  });

  it('ensures publish list and required fields exist', async () => {
    const addList = jest.fn().mockResolvedValue(undefined);
    const fields = {
      addDateTime: jest.fn().mockResolvedValue(undefined),
      addChoice: jest.fn().mockResolvedValue(undefined),
      addMultilineText: jest.fn().mockResolvedValue(undefined),
      addLookup: jest.fn().mockResolvedValue(undefined),
    };

    const listsFn: any = jest.fn().mockResolvedValue([]);
    listsFn.add = addList;
    listsFn.getByTitle = jest.fn((title: string) => {
      if (title === 'Schedule Stories') {
        return {
          select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue({ Id: 'sched-list-id' })),
        };
      }
      return { fields };
    });

    (availableStoriesService.getSp as jest.Mock).mockReturnValue({
      web: { lists: listsFn },
    });

    await ensurePublishStoriesList();

    expect(addList).toHaveBeenCalled();
    expect(fields.addDateTime).toHaveBeenCalledWith('PublishedDateTime');
    expect(fields.addMultilineText).toHaveBeenCalledWith('BoardStateJson');
    expect(fields.addLookup).toHaveBeenCalled();
  });
});
