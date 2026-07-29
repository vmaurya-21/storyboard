import * as React from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import DeleteGroupModal from '../DeleteGroupModal';
import { ScheduledStoryGroup } from '../types';

describe('DeleteGroupModal', () => {
  const group: ScheduledStoryGroup = {
    id: 'scheduled-1',
    date: new Date(2026, 6, 30),
    time: '14:30',
    slotStories: {
      'slot-1': {
        id: 'story-1',
        title: 'Story 1',
        description: 'Desc',
        imageUrl: 'https://example.com/image.jpg',
        linkToPost: 'https://example.com/post',
        date: '2026-07-30',
      },
    },
    layoutType: 'connectHomepage',
    layoutName: 'Connect Homepage',
    createdAt: new Date(2026, 6, 1),
  };

  it('renders confirmation text', () => {
    render(<DeleteGroupModal group={group} onClose={jest.fn()} onConfirm={jest.fn()} />);

    expect(screen.getByText('Delete Scheduled Group?')).toBeInTheDocument();
    expect(screen.getByText(/2:30 PM/i)).toBeInTheDocument();
  });

  it('calls onConfirm when delete button is clicked', () => {
    const onConfirm = jest.fn();
    render(<DeleteGroupModal group={group} onClose={jest.fn()} onConfirm={onConfirm} />);

    fireEvent.click(screen.getByText('Delete'));
    expect(onConfirm).toHaveBeenCalledTimes(1);
  });

  it('calls onClose from cancel button and close icon', () => {
    const onClose = jest.fn();
    render(<DeleteGroupModal group={group} onClose={onClose} onConfirm={jest.fn()} />);

    fireEvent.click(screen.getByText('Cancel'));
    fireEvent.click(screen.getByTitle('Close'));

    expect(onClose).toHaveBeenCalledTimes(2);
  });

  it('closes when clicking the overlay', () => {
    const onClose = jest.fn();
    const { container } = render(<DeleteGroupModal group={group} onClose={onClose} onConfirm={jest.fn()} />);

    fireEvent.click(container.firstChild as HTMLElement);
    expect(onClose).toHaveBeenCalledTimes(1);
  });
});
