import * as React from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import DateTimePicker, { formatTime12 } from '../DateTimePicker';

describe('DateTimePicker', () => {
  it('formats 24-hour time into 12-hour labels', () => {
    expect(formatTime12('00:00')).toBe('12:00 AM');
    expect(formatTime12('13:05')).toBe('01:05 PM');
    expect(formatTime12('bad')).toBe('12:00 AM');
  });

  it('renders selected date preview and triggers callbacks', () => {
    const onSelectDate = jest.fn();
    const onTimeChange = jest.fn();
    const onConfirm = jest.fn();

    render(
      <DateTimePicker
        selectedDate="2026-07-30"
        time="09:15"
        onSelectDate={onSelectDate}
        onTimeChange={onTimeChange}
        onConfirm={onConfirm}
        confirmText="Schedule"
      />
    );

    expect(screen.getByText(/at 09:15 AM/i)).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText('Time'), { target: { value: '10:45' } });
    expect(onTimeChange).toHaveBeenCalledWith('10:45');

    fireEvent.click(screen.getByText('Schedule'));
    expect(onConfirm).toHaveBeenCalled();

    fireEvent.click(screen.getByRole('button', { name: '30' }));
    expect(onSelectDate).toHaveBeenCalledWith('2026-07-30');
  });

  it('navigates months and supports disabled days', () => {
    const onSelectDate = jest.fn();

    render(
      <DateTimePicker
        selectedDate=""
        time="09:00"
        onSelectDate={onSelectDate}
        onTimeChange={jest.fn()}
        onConfirm={jest.fn()}
        isDateDisabled={(date) => date.getDate() === 1}
      />
    );

    fireEvent.click(screen.getByLabelText('Next month'));
    fireEvent.click(screen.getByLabelText('Previous month'));

    const day1 = screen.getByRole('button', { name: '1' });
    expect(day1).toBeDisabled();
  });

  it('rolls across year boundaries when navigating months', () => {
    const { unmount } = render(
      <DateTimePicker
        selectedDate="2026-01-15"
        time="09:00"
        onSelectDate={jest.fn()}
        onTimeChange={jest.fn()}
        onConfirm={jest.fn()}
      />
    );

    fireEvent.click(screen.getByLabelText('Previous month'));
    expect(screen.getByText('December 2025')).toBeInTheDocument();

    unmount();

    render(
      <DateTimePicker
        selectedDate="2026-12-15"
        time="09:00"
        onSelectDate={jest.fn()}
        onTimeChange={jest.fn()}
        onConfirm={jest.fn()}
      />
    );

    fireEvent.click(screen.getByLabelText('Next month'));
    expect(screen.getByText('January 2027')).toBeInTheDocument();
  });
});
