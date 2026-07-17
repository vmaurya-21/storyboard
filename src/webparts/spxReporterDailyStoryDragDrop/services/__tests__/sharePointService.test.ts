import {
  initializeSharePoint,
  getSp,
  createStory,
  getStories,
  updateStory,
  deleteStory,
  ensureListExists,
  IStoryItem
} from '../sharePointService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';

// The @pnp/sp mock is already defined in src/__mocks__/@pnp/sp.ts
// We just need to get the mock functions

// Current list the service targets.
const LIST_NAME = 'Available Stories';

describe('sharePointService', () => {
  let mockContext: IWebPartContext;
  let mockSpfi: jest.Mock;
  let mockGetByTitle: jest.Mock;
  let mockItems: jest.Mock;
  let mockAdd: jest.Mock;
  let mockGetById: jest.Mock;
  let mockUpdate: jest.Mock;
  let mockDelete: jest.Mock;
  let mockSelect: jest.Mock;
  let mockSelectInvoke: jest.Mock; // `.select(...)` returns this; calling it resolves the rows
  let mockLists: jest.Mock;
  let mockListsAdd: jest.Mock;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let mockFields: any; // object so we can attach the field-add helpers
  let mockAddText: jest.Mock;
  let mockAddDateTime: jest.Mock;

  beforeEach(() => {
    // Reset all mocks before each test
    jest.clearAllMocks();

    // Create mock context
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

    // Setup mock functions
    mockAdd = jest.fn();
    mockUpdate = jest.fn();
    mockDelete = jest.fn();
    mockGetById = jest.fn();
    mockItems = jest.fn();
    mockSelectInvoke = jest.fn();
    mockSelect = jest.fn().mockReturnValue(mockSelectInvoke);
    mockAddText = jest.fn();
    mockAddDateTime = jest.fn();
    mockFields = {
      addText: mockAddText,
      addDateTime: mockAddDateTime
    };
    mockListsAdd = jest.fn();
    mockLists = jest.fn();
    mockGetByTitle = jest.fn();

    // Setup the mock chain
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

    // Setup spfi mock
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

    it('should create a story successfully (active, with URL fields)', async () => {
      const mockAddResult = { data: { ID: 1 }, item: {} };
      mockAdd.mockResolvedValue(mockAddResult);

      const result = await createStory(mockStory);

      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockAdd).toHaveBeenCalledWith({
        Title: mockStory.title,
        description: mockStory.description,
        IsActive: 'Yes',
        imageUrl: { Url: mockStory.imageUrl },
        linkToPost: { Url: mockStory.linkToPost }
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
        description: '',
        IsActive: 'Yes',
        imageUrl: { Url: storyWithoutOptionals.imageUrl }
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
          linkToPost: { Url: 'https://example.com/post1' }
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
        created: ''
      });
      expect(stories[1].imageUrl).toBe('https://example.com/image2.jpg');
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
        created: ''
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

    it('should exclude inactive (IsActive = "No") stories', async () => {
      mockSelectInvoke.mockResolvedValue([
        { ID: 1, Title: 'Active (no flag)' },
        { ID: 2, Title: 'Inactive', IsActive: 'No' },
        { ID: 3, Title: 'Active', IsActive: 'Yes' }
      ]);

      const stories = await getStories();

      expect(stories.map(s => s.id)).toEqual([1, 3]);
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
        description: mockStory.description,
        imageUrl: { Url: mockStory.imageUrl },
        linkToPost: { Url: mockStory.linkToPost }
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
        description: '',
        imageUrl: { Url: storyWithoutOptionals.imageUrl }
      });
    });

    it('should handle API errors', async () => {
      mockUpdate.mockRejectedValue(new Error('Update failed'));

      await expect(updateStory(1, mockStory)).rejects.toThrow('Update failed');
    });
  });

  describe('deleteStory (soft delete)', () => {
    beforeEach(() => {
      initializeSharePoint(mockContext);
    });

    it('should deactivate the story (IsActive = "No") instead of hard-deleting', async () => {
      mockUpdate.mockResolvedValue(undefined);

      await deleteStory(1);

      expect(mockGetByTitle).toHaveBeenCalledWith(LIST_NAME);
      expect(mockGetById).toHaveBeenCalledWith(1);
      expect(mockUpdate).toHaveBeenCalledWith({ IsActive: 'No' });
      expect(mockDelete).not.toHaveBeenCalled();
    });

    it('should handle API errors', async () => {
      mockUpdate.mockRejectedValue(new Error('Delete failed'));

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
      // Load a fresh, un-initialized copy of the module.
      let fresh: typeof import('../sharePointService') | undefined;
      jest.isolateModules(() => {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        fresh = require('../sharePointService');
      });

      await expect(fresh!.getStories()).rejects.toThrow('SharePoint not initialized');
    });
  });
});
