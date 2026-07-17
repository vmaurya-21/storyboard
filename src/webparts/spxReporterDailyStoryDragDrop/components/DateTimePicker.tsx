import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './DateTimePicker.module.scss';

// Ported from the reference ui/datetime-picker.tsx: a 256px panel with a
// month calendar (prev/next nav) and a footer holding "Time:", a time input
// and a cyan confirm button, plus a preview line once a date is picked.

export interface IDateTimePickerProps {
  selectedDate: string; // 'YYYY-MM-DD' or '' when nothing selected
  time: string; // 'HH:mm'
  onSelectDate: (date: string) => void;
  onTimeChange: (time: string) => void;
  onConfirm: () => void;
  confirmText?: string;
  isDateDisabled?: (date: Date) => boolean;
}

const WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

const toDateStr = (year: number, month: number, day: number): string =>
  `${year}-${pad2(month + 1)}-${pad2(day)}`;

// '09:30' -> '09:30 AM'
export const formatTime12 = (time: string): string => {
  const parts = time.split(':');
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  const suffix = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return `${pad2(h12)}:${pad2(m)} ${suffix}`;
};

const DateTimePicker: React.FC<IDateTimePickerProps> = ({
  selectedDate,
  time,
  onSelectDate,
  onTimeChange,
  onConfirm,
  confirmText = 'Confirm',
  isDateDisabled
}) => {
  const initial = selectedDate ? new Date(`${selectedDate}T00:00:00`) : new Date();
  const [viewYear, setViewYear] = useState(initial.getFullYear());
  const [viewMonth, setViewMonth] = useState(initial.getMonth());

  const goPrevMonth = (): void => {
    if (viewMonth === 0) {
      setViewMonth(11);
      setViewYear(y => y - 1);
    } else {
      setViewMonth(m => m - 1);
    }
  };

  const goNextMonth = (): void => {
    if (viewMonth === 11) {
      setViewMonth(0);
      setViewYear(y => y + 1);
    } else {
      setViewMonth(m => m + 1);
    }
  };

  const monthLabel = new Date(viewYear, viewMonth, 1).toLocaleDateString('en-US', {
    month: 'long',
    year: 'numeric'
  });

  const now = new Date();
  const todayStr = toDateStr(now.getFullYear(), now.getMonth(), now.getDate());
  const firstWeekday = new Date(viewYear, viewMonth, 1).getDay();
  const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();

  const cells: React.ReactNode[] = [];
  for (let i = 0; i < firstWeekday; i++) {
    cells.push(<div key={`empty-${i}`} className={styles.emptyDay} />);
  }
  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = toDateStr(viewYear, viewMonth, day);
    const dateObj = new Date(viewYear, viewMonth, day);
    const disabled = isDateDisabled ? isDateDisabled(dateObj) : false;
    const isSelected = dateStr === selectedDate;
    const isToday = dateStr === todayStr;
    cells.push(
      <button
        key={dateStr}
        type="button"
        className={`${styles.day} ${isSelected ? styles.daySelected : ''} ${isToday && !isSelected ? styles.dayToday : ''}`}
        onClick={() => onSelectDate(dateStr)}
        disabled={disabled}
        aria-pressed={isSelected}
      >
        {day}
      </button>
    );
  }

  const preview = selectedDate
    ? `${new Date(`${selectedDate}T00:00:00`).toLocaleDateString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric'
      })} at ${formatTime12(time)}`
    : '';

  return (
    <div className={styles.picker}>
      <div className={styles.calendarArea}>
        <div className={styles.monthHeader}>
          <button
            type="button"
            className={styles.navBtn}
            onClick={goPrevMonth}
            aria-label="Previous month"
          >
            <Icon iconName="ChevronLeft" />
          </button>
          <span className={styles.monthLabel}>{monthLabel}</span>
          <button
            type="button"
            className={styles.navBtn}
            onClick={goNextMonth}
            aria-label="Next month"
          >
            <Icon iconName="ChevronRight" />
          </button>
        </div>
        <div className={styles.weekdays}>
          {WEEKDAYS.map(w => (
            <div key={w}>{w}</div>
          ))}
        </div>
        <div className={styles.days}>{cells}</div>
      </div>

      <div className={styles.footer}>
        <div className={styles.timeRow}>
          <label className={styles.timeLabel}>Time:</label>
          <input
            type="time"
            className={styles.timeInput}
            value={time}
            onChange={(e) => onTimeChange(e.target.value)}
            aria-label="Time"
          />
          <button
            type="button"
            className={`${styles.confirmBtn} ${!selectedDate ? styles.confirmDim : ''}`}
            onClick={onConfirm}
          >
            {confirmText}
          </button>
        </div>
        {selectedDate && <div className={styles.preview}>{preview}</div>}
      </div>
    </div>
  );
};

export default DateTimePicker;
