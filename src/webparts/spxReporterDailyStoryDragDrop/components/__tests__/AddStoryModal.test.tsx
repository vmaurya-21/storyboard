import * as React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import AddStoryModal from '../AddStoryModal';

describe('AddStoryModal Component', () => {
  const mockOnClose = jest.fn();
  const mockOnAdd = jest.fn();

  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('should render modal with title', () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    expect(screen.getByText('Add New Story')).toBeInTheDocument();
  });

  it('should render all form fields', () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    expect(screen.getByLabelText(/title/i)).toBeInTheDocument();
    expect(screen.getByLabelText(/description/i)).toBeInTheDocument();
    expect(screen.getByLabelText(/image url/i)).toBeInTheDocument();
    expect(screen.getByLabelText(/link to post/i)).toBeInTheDocument();
  });

  it('should show validation error when title is empty', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Title is required')).toBeInTheDocument();
    });
    
    expect(mockOnAdd).not.toHaveBeenCalled();
  });

  it('should show validation error when image URL is empty', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    await userEvent.type(titleInput, 'Test Story');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Image URL is required')).toBeInTheDocument();
    });
  });

  it('should show validation error for invalid image URL', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const imageInput = screen.getByLabelText(/image url/i);
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(imageInput, 'not-a-valid-url');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Please enter a valid URL')).toBeInTheDocument();
    });
  });

  it('should show validation error when link to post is empty', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const imageInput = screen.getByLabelText(/image url/i);
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Link to Post is required')).toBeInTheDocument();
    });
  });

  it('should call onAdd with form data when valid', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const descInput = screen.getByLabelText(/description/i);
    const imageInput = screen.getByLabelText(/image url/i);
    const linkInput = screen.getByLabelText(/link to post/i);
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(descInput, 'Test Description');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'https://example.com/post');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(mockOnAdd).toHaveBeenCalledWith({
        title: 'Test Story',
        description: 'Test Description',
        imageUrl: 'https://example.com/image.jpg',
        linkToPost: 'https://example.com/post'
      });
    });
  });

  it('should use default description when empty', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const imageInput = screen.getByLabelText(/image url/i);
    const linkInput = screen.getByLabelText(/link to post/i);
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'https://example.com/post');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(mockOnAdd).toHaveBeenCalledWith(
        expect.objectContaining({
          description: '(optional)'
        })
      );
    });
  });

  it('should call onClose when cancel button is clicked', () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const cancelButton = screen.getByText('Cancel');
    fireEvent.click(cancelButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should call onClose when close button (×) is clicked', () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const closeButton = screen.getByTitle('Close');
    fireEvent.click(closeButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should call onClose when clicking overlay', () => {
    const { container } = render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const overlay = container.firstChild as HTMLElement;
    fireEvent.click(overlay);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should not close when clicking modal content', () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const modalContent = screen.getByText('Add New Story').parentElement?.parentElement;
    if (modalContent) {
      fireEvent.click(modalContent);
    }
    
    expect(mockOnClose).not.toHaveBeenCalled();
  });

  it('should show validation error for invalid link to post URL', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const imageInput = screen.getByLabelText(/image url/i);
    const linkInput = screen.getByLabelText(/link to post/i);
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'not-a-valid-url');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(screen.getByText('Please enter a valid URL')).toBeInTheDocument();
    });
    
    expect(mockOnAdd).not.toHaveBeenCalled();
  });

  it('should reset form state when cancel button is clicked', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const descInput = screen.getByLabelText(/description/i);
    const imageInput = screen.getByLabelText(/image url/i);
    const linkInput = screen.getByLabelText(/link to post/i);
    
    // Fill in the form
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(descInput, 'Test Description');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'https://example.com/post');
    
    // Click cancel
    const cancelButton = screen.getByText('Cancel');
    fireEvent.click(cancelButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should reset form state when close button (×) is clicked', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i);
    const descInput = screen.getByLabelText(/description/i);
    const imageInput = screen.getByLabelText(/image url/i);
    const linkInput = screen.getByLabelText(/link to post/i);
    
    // Fill in the form
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(descInput, 'Test Description');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'https://example.com/post');
    
    // Click close button
    const closeButton = screen.getByTitle('Close');
    fireEvent.click(closeButton);
    
    expect(mockOnClose).toHaveBeenCalledTimes(1);
  });

  it('should clear all fields after successful submission', async () => {
    render(<AddStoryModal onClose={mockOnClose} onAdd={mockOnAdd} />);
    
    const titleInput = screen.getByLabelText(/title/i) as HTMLInputElement;
    const descInput = screen.getByLabelText(/description/i) as HTMLTextAreaElement;
    const imageInput = screen.getByLabelText(/image url/i) as HTMLInputElement;
    const linkInput = screen.getByLabelText(/link to post/i) as HTMLInputElement;
    
    await userEvent.type(titleInput, 'Test Story');
    await userEvent.type(descInput, 'Test Description');
    await userEvent.type(imageInput, 'https://example.com/image.jpg');
    await userEvent.type(linkInput, 'https://example.com/post');
    
    const saveButton = screen.getByText('Save');
    fireEvent.click(saveButton);
    
    await waitFor(() => {
      expect(mockOnAdd).toHaveBeenCalledWith({
        title: 'Test Story',
        description: 'Test Description',
        imageUrl: 'https://example.com/image.jpg',
        linkToPost: 'https://example.com/post'
      });
      expect(mockOnClose).toHaveBeenCalledTimes(1);
    });
  });
});
