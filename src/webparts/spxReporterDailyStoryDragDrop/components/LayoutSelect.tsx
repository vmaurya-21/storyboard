import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './LayoutSelect.module.scss';

export interface ILayoutOption {
  value: string;
  label: string;
}

interface ILayoutSelectProps {
  value: string;
  options: ILayoutOption[];
  onChange: (value: string) => void;
  className?: string;
  disabled?: boolean;
  ariaLabel?: string;
}

// A lightweight shadcn-style <Select> replacement: styled trigger with a
// chevron plus a popover list whose selected item shows a check — matching
// the reference "View" dropdown (which a native <select> can't fully style).
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

  // Close on outside pointer (mouse + touch) and on Escape (returning focus).
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

  // Move focus into the list (selected item) when it opens.
  useEffect(() => {
    if (!open) return undefined;
    const idx = selectedIndex >= 0 ? selectedIndex : 0;
    const t = window.setTimeout(() => {
      const el = itemRefs.current[idx];
      if (el) el.focus();
    }, 0);
    return () => window.clearTimeout(t);
  }, [open, selectedIndex]);

  const focusItem = (idx: number): void => {
    const el = itemRefs.current[idx];
    if (el) el.focus();
  };

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
