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

    expect(
      screen.getByText('Are you sure you want to delete this scheduled group?')
    ).toBeInTheDocument();
    expect(screen.getByText('Thursday, Jul 30')).toBeInTheDocument();
  });

  it('counts only the filled slots when naming the story total', () => {
    render(
      <DeleteGroupModal
        group={{ ...group, slotStories: { ...group.slotStories, 'slot-2': undefined } }}
        onClose={jest.fn()}
        onConfirm={jest.fn()}
      />
    );

    expect(screen.getByText(/return all 1 story to the available list/i)).toBeInTheDocument();
  });

  it('calls onConfirm when delete button is clicked', () => {
    const onConfirm = jest.fn();
    render(<DeleteGroupModal group={group} onClose={jest.fn()} onConfirm={onConfirm} />);

    fireEvent.click(screen.getByText('Yes, Delete Group'));
    expect(onConfirm).toHaveBeenCalledTimes(1);
  });

  it('calls onClose from cancel button and close icon', () => {
    const onClose = jest.fn();
    render(<DeleteGroupModal group={group} onClose={onClose} onConfirm={jest.fn()} />);

    fireEvent.click(screen.getByText('Cancel'));
    fireEvent.click(screen.getByTitle('Close'));

    expect(onClose).toHaveBeenCalledTimes(2);
  });

  // Radix `AlertDialog` ignores outside clicks by design: a destructive
  // confirmation should not be dismissible by a stray click on the backdrop.
  it('does not close when clicking the overlay', () => {
    const onClose = jest.fn();
    const { container } = render(<DeleteGroupModal group={group} onClose={onClose} onConfirm={jest.fn()} />);

    fireEvent.click(container.firstChild as HTMLElement);
    expect(onClose).not.toHaveBeenCalled();
  });

  it('closes on Escape', () => {
    const onClose = jest.fn();
    render(<DeleteGroupModal group={group} onClose={onClose} onConfirm={jest.fn()} />);

    fireEvent.keyDown(document, { key: 'Escape' });
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it('names itself for assistive tech', () => {
    render(<DeleteGroupModal group={group} onClose={jest.fn()} onConfirm={jest.fn()} />);

    expect(
      screen.getByRole('alertdialog', {
        name: 'Are you sure you want to delete this scheduled group?',
      })
    ).toBeInTheDocument();
  });
});
