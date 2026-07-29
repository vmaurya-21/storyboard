import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './LayoutSelect.module.scss';

/** One entry in the {@link LayoutSelect} dropdown. */
export interface ILayoutOption {
  /** Value reported to `onChange`; in practice a `LayoutType`. */
  value: string;
  /** Text shown for the option. */
  label: string;
}

/** Props for {@link LayoutSelect}. */
interface ILayoutSelectProps {
  /** Currently selected option value. */
  value: string;
  /** Options to offer, in display order. */
  options: ILayoutOption[];
  /** Called with the new value when an option is chosen. */
  onChange: (value: string) => void;
  /** Extra class for the root element. */
  className?: string;
  /** Whether the control is disabled. */
  disabled?: boolean;
  /** Accessible name for the trigger button. */
  ariaLabel?: string;
}

/**
 * Accessible dropdown for choosing the board layout.
 *
 * Hand-rolled rather than a native `<select>` so the popover can be styled to
 * match the reference. It implements the listbox keyboard contract (arrows,
 * Home/End, Enter/Space, Escape), moves focus to the active option while
 * open, restores focus to the trigger on close, and closes on an outside
 * pointer.
 */
const LayoutSelect: React.FC<ILayoutSelectProps> = ({
  value,
  options,
  onChange,
  className,
  disabled,
  ariaLabel
}) => {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const itemRefs = useRef<Array<HTMLButtonElement | null>>([]);

  const selectedIndex = options.findIndex(o => o.value === value);
  const selected = selectedIndex >= 0 ? options[selectedIndex] : undefined;

  useEffect(() => {
    if (!open) return undefined;
    const onOutside = (e: Event): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        setOpen(false);
        if (triggerRef.current) triggerRef.current.focus();
      }
    };
    document.addEventListener('mousedown', onOutside);
    document.addEventListener('touchstart', onOutside);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onOutside);
      document.removeEventListener('touchstart', onOutside);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  useEffect(() => {
    if (!open) return undefined;
    const idx = selectedIndex >= 0 ? selectedIndex : 0;
    const t = window.setTimeout(() => {
      const el = itemRefs.current[idx];
      if (el) el.focus();
    }, 0);
    return () => window.clearTimeout(t);
  }, [open, selectedIndex]);

  /**
   * Focuses the option at a given index, if it is mounted.
   *
   * @param idx - Zero-based index into `options`.
   */
  const focusItem = (idx: number): void => {
    const el = itemRefs.current[idx];
    if (el) el.focus();
  };

  /**
   * Handles roving-focus keys inside the open list.
   *
   * Arrow keys wrap around the ends; Home and End jump to the first and last
   * option. Each handled key has its default suppressed so the page does not
   * scroll behind the popover.
   *
   * @param e - The keyboard event from the list container.
   */
  const onListKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    const count = options.length;
    if (count === 0) return;
    const current = itemRefs.current.indexOf(
      document.activeElement as HTMLButtonElement
    );
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusItem(current < count - 1 ? current + 1 : 0);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusItem(current > 0 ? current - 1 : count - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusItem(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusItem(count - 1);
    }
  };

  /**
   * Commits a selection: reports it, closes the popover and returns focus to
   * the trigger so keyboard users are not stranded.
   *
   * @param v - Value of the chosen option.
   */
  const choose = (v: string): void => {
    onChange(v);
    setOpen(false);
    if (triggerRef.current) triggerRef.current.focus();
  };

  return (
    <div className={`${styles.select} ${className || ''}`} ref={rootRef}>
      <button
        type="button"
        ref={triggerRef}
        className={styles.trigger}
        onClick={() => !disabled && setOpen(o => !o)}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={ariaLabel || 'Select layout'}
        disabled={disabled}
      >
        <span className={styles.triggerLabel}>{selected ? selected.label : 'Select…'}</span>
        <Icon iconName="ChevronDown" className={styles.chevron} />
      </button>

      {open && !disabled && (
        <div className={styles.popover} role="listbox" onKeyDown={onListKeyDown}>
          {options.map((opt, i) => {
            const isSelected = opt.value === value;
            return (
              <button
                key={opt.value}
                type="button"
                role="option"
                aria-selected={isSelected}
                ref={el => { itemRefs.current[i] = el; }}
                className={`${styles.item} ${isSelected ? styles.itemSelected : ''}`}
                onClick={() => choose(opt.value)}
              >
                <span className={styles.itemLabel}>{opt.label}</span>
                <span className={styles.check}>
                  {isSelected && <Icon iconName="CheckMark" />}
                </span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default LayoutSelect;
