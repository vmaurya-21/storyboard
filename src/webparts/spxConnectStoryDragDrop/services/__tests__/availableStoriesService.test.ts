import {
  initializeSharePoint,
  getSp,
  deriveSourceFromUrl,
  createStory,
  getStories,
  updateStory,
  deleteStory,
  ensureListExists,
  IStoryItem
} from '../availableStoriesService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';


const LIST_NAME = 'AvailableStories';

describe('availableStoriesService', () => {
  let mockContext: IWebPartContext;
  let mockSpfi: jest.Mock;
  let mockGetByTitle: jest.Mock;
  let mockItems: jest.Mock;
  let mockAdd: jest.Mock;
  let mockGetById: jest.Mock;
  let mockUpdate: jest.Mock;
  let mockDelete: jest.Mock;
  let mockSelect: jest.Mock;
  let mockExpand: jest.Mock;
  // Terminal call of the `.select(...).expand(...)()` chain — the expand is
  // what projects the Created By byline.
  let mockSelectInvoke: jest.Mock;
  let mockLists: jest.Mock;
  let mockListsAdd: jest.Mock;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let mockFields: any;
  let mockAddText: jest.Mock;
  let mockAddDateTime: jest.Mock;

  beforeEach(() => {
    jest.clearAllMocks();

    mockContext = {
      pageContext: {
        web: {
          absoluteUrl: 'https://contoso.sharepoint.com',
          title: 'Test Site',
          serverRelativeUrl: '/sites/test'
        },
        user: {
          displayName: 'Test User',
          email: 'test@contoso.com',
          loginName: 'i:0#.f|membership|test@contoso.com'
        }
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;

    mockAdd = jest.fn();
    mockUpdate = jest.fn();
    mockDelete = jest.fn();
    mockGetById = jest.fn();
    mockItems = jest.fn();
    mockSelectInvoke = jest.fn();
    mockExpand = jest.fn().mockReturnValue(mockSelectInvoke);
    mockSelect = jest.fn().mockReturnValue({ expand: mockExpand });
    mockAddText = jest.fn();
    mockAddDateTime = jest.fn();
    mockFields = {
      addText: mockAddText,
      addDateTime: mockAddDateTime
    };
    mockListsAdd = jest.fn();
    mockLists = jest.fn();
    mockGetByTitle = jest.fn();

    mockGetById.mockReturnValue({
      update: mockUpdate,
      delete: mockDelete
    });

    mockGetByTitle.mockReturnValue({
      items: Object.assign(mockItems, {
        add: mockAdd,
        getById: mockGetById,
        select: mockSelect
      }),
      fields: mockFields
    });

    mockSpfi = spfi as jest.Mock;
    mockSpfi.mockReturnValue({
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      using: jest.fn((_config: any) => ({
        web: {
          lists: Object.assign(mockLists, {
            getByTitle: mockGetByTitle,
            add: mockListsAdd
          })
        }
      }))
    });
  });

  describe('initializeSharePoint', () => {
    it('should initialize SharePoint with context', () => {
      const result = initializeSharePoint(mockContext);

      expect(result).toBeDefined();
      expect(mockSpfi).toHaveBeenCalled();
    });

    it('should return SPFI instance with web property', () => {
      const sp = initializeSharePoint(mockContext);

      expect(sp).toHaveProperty('web');
      expect(sp.web).toHaveProperty('lists');
    });

    it('should reinitialize on subsequent calls', () => {
      const sp1 = initializeSharePoint(mockContext);
      const sp2 = initializeSharePoint(mockContext);

      expect(sp1).toBeDefined();
      expect(sp2).toBeDefined();
      expect(mockSpfi).toHaveBeenCalledTimes(2);
    });
  });

  describe('getSp', () => {
    it('should return SPFI instance after initialization', () => {
      initializeSharePoint(mockContext);
      const sp = getSp();

      expect(sp).toBeDefined();
    });
  });

  describe('deriveSourceFromUrl', () => {
    it('classifies SharePoint domains as internal and whitecase.com as external', () => {
      expect(deriveSourceFromUrl('https://whitecasempsaemea.sharepoint.com/sites/comms')).toBe('internal');
      expect(deriveSourceFromUrl('https://contoso.sharepoint.com/sites/comms')).toBe('external');
      expect(deriveSourceFromUrl('https://whitecase.com/news')).toBe('external');
      expect(deriveSourceFromUrl('https://sub.whitecase.com/news')).toBe('external');
    });

    it('classifies LinkedIn domains as linkedin', () => {
      expect(deriveSourceFromUrl('https://www.linkedin.com/posts/abc')).toBe('linkedin');
    });
  });

  describe('createStory', () => {
    const mockStory: IStoryItem = {
      title: 'Test Story',
      description: 'Test Description',
      imageUrl: 'https://example.com/image.jpg',
      linkToPost: 'https://example.com/post'
    };

    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should create a story successfully (with URL fields)', async () => {
      const mockAddResult = { data: { ID: 1 }, item: {} };
      mockAdd.mockResolvedValue(mockAddResult);

      const result = await createStory(mockStory);

      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockAdd).toHaveBeenCalledWith({
        Title: mockStory.title,
        Description: mockStory.description,
        Source: 'external',
        ImageUrl: { Url: mockStory.imageUrl },
        LinkToPost: { Url: mockStory.linkToPost }
      });
      expect(result).toEqual(mockAddResult);
    });

    it('should create story without optional fields', async () => {
      const storyWithoutOptionals: IStoryItem = {
        title: 'Test Story',
        imageUrl: 'https://example.com/image.jpg'
      };

      mockAdd.mockResolvedValue({ data: { ID: 1 }, item: {} });

      await createStory(storyWithoutOptionals);

      expect(mockAdd).toHaveBeenCalledWith({
        Title: storyWithoutOptionals.title,
        Description: '',
        Source: 'internal',
        ImageUrl: { Url: storyWithoutOptionals.imageUrl }
      });
    });

    it('should handle API errors', async () => {
      mockAdd.mockRejectedValue(new Error('API Error'));

      await expect(createStory(mockStory)).rejects.toThrow('API Error');
    });
  });

  describe('getStories', () => {
    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should fetch stories successfully with URL objects', async () => {
      const mockItemsData = [
        {
          ID: 1,
          Title: 'Story 1',
          description: 'Description 1',
          imageUrl: { Url: 'https://example.com/image1.jpg' },
          linkToPost: { Url: 'https://example.com/post1' },
          Author: { Title: 'Ada Lovelace' }
        },
        {
          ID: 2,
          Title: 'Story 2',
          description: 'Description 2',
          imageUrl: { Url: 'https://example.com/image2.jpg' },
          linkToPost: { Url: 'https://example.com/post2' }
        }
      ];

      mockSelectInvoke.mockResolvedValue(mockItemsData);

      const stories = await getStories();

      expect(stories).toHaveLength(2);
      expect(stories[0]).toEqual({
        id: 1,
        title: 'Story 1',
        description: 'Description 1',
        imageUrl: 'https://example.com/image1.jpg',
        linkToPost: 'https://example.com/post1',
        date: '',
        created: '',
        source: 'external',
        author: 'Ada Lovelace'
      });
      expect(stories[1].imageUrl).toBe('https://example.com/image2.jpg');
      // Created By is expanded, but an item can still come back without it.
      expect(stories[1].author).toBe('');
    });

    it('requests the Created By byline via select and expand', async () => {
      mockSelectInvoke.mockResolvedValue([]);

      await getStories();

      expect(mockSelect).toHaveBeenCalledWith(
        expect.anything(), expect.anything(), expect.anything(), expect.anything(),
        expect.anything(), expect.anything(), expect.anything(), expect.anything(),
        'Author/Title'
      );
      expect(mockExpand).toHaveBeenCalledWith('Author');
    });

    it('should fetch stories successfully with URL strings', async () => {
      mockSelectInvoke.mockResolvedValue([
        {
          ID: 1,
          Title: 'Story 1',
          description: 'Description 1',
          imageUrl: 'https://example.com/image1.jpg',
          linkToPost: 'https://example.com/post1'
        }
      ]);

      const stories = await getStories();

      expect(stories).toHaveLength(1);
      expect(stories[0].imageUrl).toBe('https://example.com/image1.jpg');
      expect(stories[0].linkToPost).toBe('https://example.com/post1');
    });

    it('should handle stories with missing optional fields', async () => {
      mockSelectInvoke.mockResolvedValue([{ ID: 1, Title: 'Story 1' }]);

      const stories = await getStories();

      expect(stories).toHaveLength(1);
      expect(stories[0]).toEqual({
        id: 1,
        title: 'Story 1',
        description: '',
        imageUrl: '',
        linkToPost: '',
        date: '',
        created: '',
        source: 'internal',
        author: ''
      });
    });

    it('should map Modified -> date and Created -> created', async () => {
      mockSelectInvoke.mockResolvedValue([
        {
          ID: 7,
          Title: 'Dated Story',
          imageUrl: 'https://example.com/x.jpg',
          Modified: '2026-01-02T03:04:05Z',
          Created: '2025-12-31T00:00:00Z'
        }
      ]);

      const stories = await getStories();

      expect(stories[0].date).toBe('2026-01-02T03:04:05Z');
      expect(stories[0].created).toBe('2025-12-31T00:00:00Z');
    });

    it('should handle API errors', async () => {
      mockSelectInvoke.mockRejectedValue(new Error('API Error'));

      await expect(getStories()).rejects.toThrow('API Error');
    });
  });

  describe('updateStory', () => {
    const mockStory: IStoryItem = {
      title: 'Updated Story',
      description: 'Updated Description',
      imageUrl: 'https://example.com/updated-image.jpg',
      linkToPost: 'https://example.com/updated-post'
    };

    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should update a story successfully', async () => {
      mockUpdate.mockResolvedValue(undefined);

      await updateStory(1, mockStory);

      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockGetById).toHaveBeenCalledWith(1);
      expect(mockUpdate).toHaveBeenCalledWith({
        Title: mockStory.title,
        Description: mockStory.description,
        Source: 'external',
        ImageUrl: { Url: mockStory.imageUrl },
        LinkToPost: { Url: mockStory.linkToPost }
      });
    });

    it('should update story without optional fields', async () => {
      const storyWithoutOptionals: IStoryItem = {
        title: 'Updated Story',
        imageUrl: 'https://example.com/image.jpg'
      };

      mockUpdate.mockResolvedValue(undefined);

      await updateStory(1, storyWithoutOptionals);

      expect(mockUpdate).toHaveBeenCalledWith({
        Title: storyWithoutOptionals.title,
        Description: '',
        Source: 'internal',
        ImageUrl: { Url: storyWithoutOptionals.imageUrl }
      });
    });

    it('should handle API errors', async () => {
      mockUpdate.mockRejectedValue(new Error('Update failed'));

      await expect(updateStory(1, mockStory)).rejects.toThrow('Update failed');
    });
  });

  describe('deleteStory (hard delete)', () => {
    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should permanently delete the story from the list', async () => {
      mockDelete.mockResolvedValue(undefined);

      await deleteStory(1);

      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockGetById).toHaveBeenCalledWith(1);
      expect(mockDelete).toHaveBeenCalled();
      expect(mockUpdate).not.toHaveBeenCalled();
    });

    it('should handle API errors', async () => {
      mockDelete.mockRejectedValue(new Error('Delete failed'));

      await expect(deleteStory(1)).rejects.toThrow('Delete failed');
    });
  });

  describe('ensureListExists', () => {
    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should return early if list already exists', async () => {
      mockLists.mockResolvedValue([{ Title: LIST_NAME }, { Title: 'Other List' }]);

      await ensureListExists();

      expect(mockListsAdd).not.toHaveBeenCalled();
    });

    it('should create list if it does not exist', async () => {
      mockLists.mockResolvedValue([{ Title: 'Other List' }]);
      mockListsAdd.mockResolvedValue(undefined);
      mockAddText.mockResolvedValue(undefined);
      mockAddDateTime.mockResolvedValue(undefined);

      await ensureListExists();

      expect(mockListsAdd).toHaveBeenCalledWith(LIST_NAME);
      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockAddText).toHaveBeenCalledWith('Description');
      expect(mockAddText).toHaveBeenCalledWith('ImageUrl');
      expect(mockAddText).toHaveBeenCalledWith('LinkToPost');
      expect(mockAddDateTime).toHaveBeenCalledWith('StoryDate');
    });

    it('should handle errors gracefully (never throws)', async () => {
      mockLists.mockRejectedValue(new Error('List creation failed'));

      await expect(ensureListExists()).resolves.toBeUndefined();
    });
  });

  describe('when SharePoint is not initialized', () => {
    it('getStories rejects with a clear error', async () => {
      let fresh: typeof import('../availableStoriesService') | undefined;
      jest.isolateModules(() => {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        fresh = require('../availableStoriesService');
      });

      await expect(fresh!.getStories()).rejects.toThrow('SharePoint not initialized');
    });
  });
});
