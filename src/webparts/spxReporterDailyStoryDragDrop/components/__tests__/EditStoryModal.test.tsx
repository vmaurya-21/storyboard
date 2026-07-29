import * as React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import EditStoryModal from '../EditStoryModal';
import { IStory } from '../types';

globalThis.confirm = jest.fn() as any;

describe('EditStoryModal Component', () => {
  const mockStory: IStory = {
    id: 'story-1',
    title: 'Original Title',
    description: 'Original Description',
    imageUrl: 'https://example.com/original.jpg',
    linkToPost: 'https://example.com/original',
    date: '2024-01-15'
  };

  const mockOnClose = jest.fn();
  const mockOnUpdate = jest.fn();
  const mockOnDelete = jest.fn();

  beforeEach(() => {
    jest.clearAllMocks();
    (globalThis.confirm as jest.Mock).mockReturnValue(false);
  });

  it('should render modal with title', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    expect(screen.getByText('Edit Story')).toBeInTheDocument();
  });

  it('should pre-fill form with story data', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    expect(screen.getByDisplayValue('Original Title')).toBeInTheDocument();
    expect(screen.getByDisplayValue('Original Description')).toBeInTheDocument();
    expect(screen.getByDisplayValue('https://example.com/original.jpg')).toBeInTheDocument();
    expect(screen.getByDisplayValue('https://example.com/original')).toBeInTheDocument();
  });

  it('should show validation error when title is empty', async () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const titleInput = screen.getByLabelText(/title/i);
    await userEvent.clear(titleInput);
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Title is required')).toBeInTheDocument();
    });
    
    expect(mockOnUpdate).not.toHaveBeenCalled();
  });

  it('should show validation error for invalid image URL', async () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const imageInput = screen.getByLabelText(/image url/i);
    await userEvent.clear(imageInput);
    await userEvent.type(imageInput, 'not-a-valid-url');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Please enter a valid URL')).toBeInTheDocument();
    });
  });

  it('should show validation error for invalid link URL', async () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const linkInput = screen.getByLabelText(/link to post/i);
    await userEvent.clear(linkInput);
    await userEvent.type(linkInput, 'invalid-url');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Please enter a valid URL')).toBeInTheDocument();
    });
  });

  it('should call onUpdate with modified data when valid', async () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const titleInput = screen.getByLabelText(/title/i);
    await userEvent.clear(titleInput);
    await userEvent.type(titleInput, 'Updated Title');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(mockOnUpdate).toHaveBeenCalledWith(
        expect.objectContaining({
          id: 'story-1',
          title: 'Updated Title'
        })
      );
    });
  });

  it('should call onDelete when delete is confirmed in the dialog', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );

    fireEvent.click(screen.getByText('Delete Story'));
    fireEvent.click(screen.getByText('Delete'));

    expect(mockOnDelete).toHaveBeenCalledWith('story-1');
    expect(mockOnClose).toHaveBeenCalled();
  });

  it('should not delete when the confirmation is cancelled', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );

    fireEvent.click(screen.getByText('Delete Story'));
    const cancelButtons = screen.getAllByText('Cancel');
    fireEvent.click(cancelButtons[cancelButtons.length - 1]);

    expect(mockOnDelete).not.toHaveBeenCalled();
    expect(mockOnClose).not.toHaveBeenCalled();
  });

  it('should call onClose when cancel button is clicked', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const cancelButton = screen.getByText('Cancel');
    fireEvent.click(cancelButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should call onClose when close button (×) is clicked', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const closeButton = screen.getByTitle('Close');
    fireEvent.click(closeButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should pass a blank description when cleared', async () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );

    const descInput = screen.getByLabelText(/description/i);
    await userEvent.clear(descInput);

    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);

    await waitFor(() => {
      expect(mockOnUpdate).toHaveBeenCalledWith(
        expect.objectContaining({
          description: '',
          imageUrl: 'https://example.com/original.jpg',
          linkToPost: 'https://example.com/original'
        })
      );
    });
  });

  it('should call onClose when clicking overlay', () => {
    const { container } = render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const overlay = container.firstChild as HTMLElement;
    fireEvent.click(overlay);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should not close when clicking modal content', () => {
    render(
      <EditStoryModal
        story={mockStory}
        onClose={mockOnClose}
        onUpdate={mockOnUpdate}
        onDelete={mockOnDelete}
      />
    );
    
    const modalContent = screen.getByText('Edit Story').parentElement?.parentElement;
    if (modalContent) {
      fireEvent.click(modalContent);
    }
    
    expect(mockOnClose).not.toHaveBeenCalled();
  });
});
