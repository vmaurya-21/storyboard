import { IStory } from '../types';

export const mockStories: IStory[] = [
  {
    id: 'story-1',
    title: 'Test Story 1',
    description: 'This is a test story description',
    imageUrl: 'https://via.placeholder.com/300x200',
    linkToPost: 'https://example.com/story1',
    date: '2024-01-15'
  },
  {
    id: 'story-2',
    title: 'Test Story 2',
    description: 'Another test story',
    imageUrl: 'https://via.placeholder.com/300x200',
    linkToPost: 'https://example.com/story2',
    date: '2024-01-16'
  },
  {
    id: 'story-3',
    title: 'Test Story 3',
    description: 'Third test story',
    imageUrl: 'https://via.placeholder.com/300x200',
    linkToPost: 'https://example.com/story3',
    date: '2024-01-17'
  }
];

export const mockContext = {
  pageContext: {
    web: {
      absoluteUrl: 'https://contoso.sharepoint.com',
      title: 'Test Site'
    },
    user: {
      displayName: 'Test User',
      email: 'test@contoso.com'
    }
  },
  sdks: {
    microsoftTeams: undefined
  }
};
