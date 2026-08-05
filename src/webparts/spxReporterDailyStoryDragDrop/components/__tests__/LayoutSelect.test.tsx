import * as React from 'react';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import LayoutSelect from '../LayoutSelect';

describe('LayoutSelect', () => {
  const options = [
    { value: 'connectHomepage', label: 'Connect Homepage' },
    { value: 'general', label: 'General' },
    { value: 'reporterDaily', label: 'Storyboard' },
  ];

  it('renders selected option label', () => {
    render(
      <LayoutSelect
        value="general"
        options={options}
        onChange={jest.fn()}
      />
    );

    expect(screen.getByText('General')).toBeInTheDocument();
  });

  it('opens and chooses an option', () => {
    const onChange = jest.fn();
    render(
      <LayoutSelect
        value="connectHomepage"
        options={options}
        onChange={onChange}
      />
    );

    fireEvent.click(screen.getByRole('button', { name: /Select layout/i }));
    fireEvent.click(screen.getByRole('option', { name: 'Storyboard' }));

    expect(onChange).toHaveBeenCalledWith('reporterDaily');
  });

  it('closes on escape', () => {
    render(
      <LayoutSelect
        value="connectHomepage"
        options={options}
        onChange={jest.fn()}
      />
    );

    fireEvent.click(screen.getByRole('button', { name: /Select layout/i }));
    expect(screen.getByRole('listbox')).toBeInTheDocument();

    fireEvent.keyDown(document, { key: 'Escape' });
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('closes on outside click', () => {
    render(
      <LayoutSelect
        value="connectHomepage"
        options={options}
        onChange={jest.fn()}
      />
    );

    fireEvent.click(screen.getByRole('button', { name: /Select layout/i }));
    expect(screen.getByRole('listbox')).toBeInTheDocument();

    fireEvent.mouseDown(document.body);
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('supports keyboard navigation in listbox', async () => {
    render(
      <LayoutSelect
        value="connectHomepage"
        options={options}
        onChange={jest.fn()}
      />
    );

    fireEvent.click(screen.getByRole('button', { name: /Select layout/i }));
    const listbox = screen.getByRole('listbox');

    await waitFor(() => {
      expect(screen.getByRole('option', { name: 'Connect Homepage' })).toHaveFocus();
    });

    fireEvent.keyDown(listbox, { key: 'ArrowDown' });
    expect(screen.getByRole('option', { name: 'General' })).toHaveFocus();

    fireEvent.keyDown(listbox, { key: 'End' });
    expect(screen.getByRole('option', { name: 'Storyboard' })).toHaveFocus();

    fireEvent.keyDown(listbox, { key: 'Home' });
    expect(screen.getByRole('option', { name: 'Connect Homepage' })).toHaveFocus();
  });

  it('does not open when disabled', () => {
    render(
      <LayoutSelect
        value="connectHomepage"
        options={options}
        onChange={jest.fn()}
        disabled
      />
    );

    fireEvent.click(screen.getByRole('button', { name: /Select layout/i }));
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });
});
