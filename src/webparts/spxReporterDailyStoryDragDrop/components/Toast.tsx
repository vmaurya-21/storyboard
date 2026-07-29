import * as React from 'react';
import { useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Toast.module.scss';

/** A queued notification. */
export interface IToast {
  /** Unique id; also the React key and the handle passed to `onClose`. */
  id: string;
  /** Severity, which selects the icon and colour treatment. */
  type: 'success' | 'error' | 'info';
  /** Bold headline, e.g. `Storyboard Published!`. */
  title: string;
  /** Optional supporting line beneath the title. */
  description?: string;
  /** Auto-dismiss delay in ms. Defaults to 3000. */
  duration?: number;
}

/** Props for {@link ToastItem}. */
interface IToastItemProps {
  /** The notification to render. */
  toast: IToast;
  /** Called with the toast id when it is dismissed. */
  onClose: (id: string) => void;
}

/**
 * A single notification that dismisses itself after {@link IToast.duration}.
 *
 * The timer is re-established whenever the toast's id, duration or handler
 * changes, and is cleared on unmount so a dismissed toast cannot fire.
 */
const ToastItem: React.FC<IToastItemProps> = ({ toast, onClose }) => {
  useEffect(() => {
    const timer = setTimeout(() => {
      onClose(toast.id);
    }, toast.duration ?? 3000);
    return () => clearTimeout(timer);
  }, [toast.id, toast.duration, onClose]);

  /**
   * Maps the toast severity to its Fluent icon name.
   *
   * @returns The `iconName` for the leading icon.
   */
  const getIconName = (): string => {
    switch (toast.type) {
      case 'success': return 'Completed';
      case 'error': return 'ErrorBadge';
      case 'info': return 'Info';
      default: return 'Info';
    }
  };

  return (
    <div className={`${styles.toast} ${styles[toast.type]}`} role="alert">
      <div className={styles.toastIcon}>
        <Icon iconName={getIconName()} />
      </div>
      <div className={styles.toastContent}>
        <div className={styles.toastTitle}>{toast.title}</div>
        {toast.description && <div className={styles.toastDescription}>{toast.description}</div>}
      </div>
      <button className={styles.toastCloseBtn} onClick={() => onClose(toast.id)} aria-label="Close">
        <Icon iconName="Cancel" />
      </button>
    </div>
  );
};

/** Props for {@link ToastContainer}. */
interface IToastContainerProps {
  /** Active notifications, oldest first. */
  toasts: IToast[];
  /** Called with a toast id when that toast should be removed from the queue. */
  onCloseToast: (id: string) => void;
}

/**
 * Fixed-position stack of active notifications.
 *
 * Owns no state: the parent holds the queue and removes entries in response
 * to `onCloseToast`.
 */
export const ToastContainer: React.FC<IToastContainerProps> = ({ toasts, onCloseToast }) => {
  return (
    <div className={styles.toastContainer}>
      {toasts.map(toast => (
        <ToastItem key={toast.id} toast={toast} onClose={onCloseToast} />
      ))}
    </div>
  );
};
