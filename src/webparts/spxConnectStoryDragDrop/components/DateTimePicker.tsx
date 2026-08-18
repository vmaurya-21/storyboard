import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './DateTimePicker.module.scss';


/** Props for {@link DateTimePicker}. */
export interface IDateTimePickerProps {
  /** Selected day as `YYYY-MM-DD`, or `''` when nothing is chosen yet. */
  selectedDate: string;
  /** Selected time of day in 24-hour `HH:mm` form. */
  time: string;
  /** Called with a `YYYY-MM-DD` string when a day is picked. */
  onSelectDate: (date: string) => void;
  /** Called with an `HH:mm` string when the time changes. */
  onTimeChange: (time: string) => void;
  /** Called when the confirm button is pressed. */
  onConfirm: () => void;
  /** Label for the confirm button. Defaults to `Confirm`. */
  confirmText?: string;
  /** Return `true` to render a day as unselectable, e.g. an already-booked date. */
  isDateDisabled?: (date: Date) => boolean;
}

/** Column headings for the calendar grid, starting on Sunday. */
const WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];

/**
 * Left-pads a number to two digits.
 *
 * @param n - Value to pad; expected to be non-negative.
 * @returns The value as a two-character string, e.g. `7` becomes `07`.
 */
const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/**
 * Builds a `YYYY-MM-DD` string from calendar parts.
 *
 * Assembled from local parts rather than `toISOString()` so the date never
 * shifts a day under a timezone conversion.
 *
 * @param year - Four-digit year.
 * @param month - Zero-based month, as returned by `Date.getMonth()`.
 * @param day - Day of the month, starting at 1.
 * @returns The formatted date string.
 */
const toDateStr = (year: number, month: number, day: number): string =>
  `${year}-${pad2(month + 1)}-${pad2(day)}`;

/**
 * Converts a 24-hour `HH:mm` string to a padded 12-hour label.
 *
 * Unparseable parts fall back to zero, so a malformed value renders as
 * `12:00 AM` rather than `NaN`.
 *
 * @param time - Time of day as `HH:mm`.
 * @returns The 12-hour label, e.g. `09:00` becomes `09:00 AM` and `13:05`
 * becomes `01:05 PM`.
 */
export const formatTime12 = (time: string): string => {
  const parts = time.split(':');
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  const suffix = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return `${pad2(h12)}:${pad2(m)} ${suffix}`;
};

/**
 * Month calendar plus a time field, used to schedule a board.
 *
 * The selected date and time are controlled by the parent; the component owns
 * only which month is on screen, seeded from `selectedDate` (or today when
 * nothing is selected).
 */
const DateTimePicker: React.FC<IDateTimePickerProps> = ({
  selectedDate,
  time,
  onSelectDate,
  onTimeChange,
  onConfirm,
  confirmText = 'Confirm',
  isDateDisabled
}) => {
  // The omitted `Z` is deliberate: a date-*time* string with no offset parses
  // as local, which is what a calendar day needs. Adding it — or passing the
  // bare `YYYY-MM-DD`, which parses as UTC — would land on the previous day
  // for any negative UTC offset.
  const initial = selectedDate ? new Date(`${selectedDate}T00:00:00`) : new Date();
  const [viewYear, setViewYear] = useState(initial.getFullYear());
  const [viewMonth, setViewMonth] = useState(initial.getMonth());

  /** Moves the calendar back one month, rolling over to December of the prior year. */
  const goPrevMonth = (): void => {
    if (viewMonth === 0) {
      setViewMonth(11);
      setViewYear(y => y - 1);
    } else {
      setViewMonth(m => m - 1);
    }
  };

  /** Moves the calendar forward one month, rolling over to January of the next year. */
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
          {/* Really disabled, not just dimmed: the caller's own "no date"
              guard stays as a backstop, but keyboard and screen-reader users
              need the state to be real. */}
          <button
            type="button"
            className={`${styles.confirmBtn} ${!selectedDate ? styles.confirmDim : ''}`}
            onClick={onConfirm}
            disabled={!selectedDate}
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
