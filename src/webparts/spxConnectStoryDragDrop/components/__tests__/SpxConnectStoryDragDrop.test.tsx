import * as React from 'react';
import { render, screen } from '@testing-library/react';
import SpxConnectStoryDragDrop from '../SpxConnectStoryDragDrop';
import { mockContext } from './mockData';
import { AppMode } from '../../SpxConnectStoryDragDropWebPart';

jest.mock('../StoryDragDrop', () => {
  return function MockStoryDragDrop({ context }: any) {
    return <div data-testid="story-drag-drop">StoryDragDrop Component - Context: {context ? 'provided' : 'missing'}</div>;
  };
});

describe('SpxConnectStoryDragDrop Component', () => {
  const defaultProps = {
    description: 'Test Description',
    isDarkTheme: false,
    context: mockContext as any,
    appMode: AppMode.Spfx,
    hasTeamsContext: false,
    userDisplayName: 'Test User'
  };

  it('should render without crashing', () => {
    render(<SpxConnectStoryDragDrop {...defaultProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should pass context to StoryDragDrop', () => {
    render(<SpxConnectStoryDragDrop {...defaultProps} />);
    expect(screen.getByText(/Context: provided/)).toBeInTheDocument();
  });

  it('should render when context is missing', () => {
    const propsWithoutContext = { ...defaultProps, context: undefined } as any;
    render(<SpxConnectStoryDragDrop {...propsWithoutContext} />);
    expect(screen.getByText(/Context: missing/)).toBeInTheDocument();
  });

  it('should render with dark theme', () => {
    const darkThemeProps = { ...defaultProps, isDarkTheme: true };
    render(<SpxConnectStoryDragDrop {...darkThemeProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with Teams context', () => {
    const teamsProps = { ...defaultProps, hasTeamsContext: true };
    render(<SpxConnectStoryDragDrop {...teamsProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with different user display name', () => {
    const userProps = { ...defaultProps, userDisplayName: 'John Doe' };
    render(<SpxConnectStoryDragDrop {...userProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });

  it('should render with custom description', () => {
    const descProps = { ...defaultProps, description: 'Custom Description' };
    render(<SpxConnectStoryDragDrop {...descProps} />);
    expect(screen.getByTestId('story-drag-drop')).toBeInTheDocument();
  });
});
