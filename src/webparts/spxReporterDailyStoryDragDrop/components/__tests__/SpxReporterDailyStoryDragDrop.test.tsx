import * as React from 'react';
import { render, screen } from '@testing-library/react';
import SpxReporterDailyStoryDragDrop from '../SpxReporterDailyStoryDragDrop';
import { mockContext } from './mockData';
import { AppMode } from '../../SpxReporterDailyStoryDragDropWebPart';

// Mock the StoryDragDrop component since it's complex
jest.mock('../StoryDragDrop', () => {
  return function MockStoryDragDrop({ context }: any) {
    return <div data-testid="story-drag-drop">StoryDragDrop Component - Context: {context ? 'provided' : 'missing'}</div>;
  };
});

describe('SpxReporterDailyStoryDragDrop Component', () => {
  const defaultProps = {
    description: 'Test Description',
    isDarkTheme: false,
    context: mockContext as any,
    appMode: AppMode.Spfx,
    hasTeamsContext: false,
    userDisplayName: 'Test User'
  };

  it('should render without crashing', () => {
    render(<SpxReporterDailyStoryDragDrop {...defaultProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should pass context to StoryDragDrop', () => {
    render(<SpxReporterDailyStoryDragDrop {...defaultProps} />);
    expect(screen.getByText(/Context: provided/)).toBeInTheDocument();
  });

  it('should render when context is missing', () => {
    const propsWithoutContext = { ...defaultProps, context: undefined } as any;
    render(<SpxReporterDailyStoryDragDrop {...propsWithoutContext} />);
    expect(screen.getByText(/Context: missing/)).toBeInTheDocument();
  });

  it('should render with dark theme', () => {
    const darkThemeProps = { ...defaultProps, isDarkTheme: true };
    render(<SpxReporterDailyStoryDragDrop {...darkThemeProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with Teams context', () => {
    const teamsProps = { ...defaultProps, hasTeamsContext: true };
    render(<SpxReporterDailyStoryDragDrop {...teamsProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with different user display name', () => {
    const userProps = { ...defaultProps, userDisplayName: 'John Doe' };
    render(<SpxReporterDailyStoryDragDrop {...userProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with custom description', () => {
    const descProps = { ...defaultProps, description: 'Custom Description' };
    render(<SpxReporterDailyStoryDragDrop {...descProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });
});
