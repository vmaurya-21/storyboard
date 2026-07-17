import * as React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import StoryCard from '../../cards/StoryCard';
import { IStory } from '../../types';

describe('StoryCard Component', () => {
  const mockStory: IStory = {
    id: 'story-1',
    title: 'Test Story',
    description: 'Test Description',
    imageUrl: 'https://via.placeholder.com/300x200',
    linkToPost: 'https://example.com',
    date: '2024-01-15'
  };

  it('should render story with title and date', () => {
    render(<StoryCard story={mockStory} variant="large" />);
    
    expect(screen.getByText('Test Story')).toBeInTheDocument();
    expect(screen.getByText('2024-01-15')).toBeInTheDocument();
  });

  it('should render image when showImage is true', () => {
    render(<StoryCard story={mockStory} variant="large" showImage={true} />);
    
    const image = screen.getByAltText('Test Story');
    expect(image).toBeInTheDocument();
    expect(image).toHaveAttribute('src', mockStory.imageUrl);
  });

  it('should not render image when showImage is false', () => {
    render(<StoryCard story={mockStory} variant="textOnly" showImage={false} />);
    
    expect(screen.queryByAltText('Test Story')).not.toBeInTheDocument();
  });

  it('should render remove button when onRemove is provided', () => {
    const onRemove = jest.fn();
    render(<StoryCard story={mockStory} variant="large" onRemove={onRemove} />);
    
    const removeBtn = screen.getByLabelText('Remove story');
    expect(removeBtn).toBeInTheDocument();
  });

  it('should call onRemove when remove button is clicked', () => {
    const onRemove = jest.fn();
    render(<StoryCard story={mockStory} variant="large" onRemove={onRemove} />);
    
    const removeBtn = screen.getByLabelText('Remove story');
    fireEvent.click(removeBtn);
    
    expect(onRemove).toHaveBeenCalledTimes(1);
  });

  it('should not render remove button when onRemove is not provided', () => {
    render(<StoryCard story={mockStory} variant="large" />);
    
    expect(screen.queryByLabelText('Remove story')).not.toBeInTheDocument();
  });

  it('should apply dragging class when isDragging is true', () => {
    const { container } = render(
      <StoryCard story={mockStory} variant="large" isDragging={true} />
    );
    
    const card = container.firstChild as HTMLElement;
    expect(card.className).toContain('dragging');
  });

  it('should apply correct variant class', () => {
    const { container } = render(
      <StoryCard story={mockStory} variant="small" />
    );
    
    const card = container.firstChild as HTMLElement;
    expect(card.className).toContain('small');
  });
});
