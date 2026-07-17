import * as React from 'react';
import { useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Toast.module.scss';

export interface IToast {
  id: string;
  type: 'success' | 'error' | 'info';
  title: string;
  description?: string;
}

interface IToastItemProps {
  toast: IToast;
  onClose: (id: string) => void;
}

const ToastItem: React.FC<IToastItemProps> = ({ toast, onClose }) => {
  useEffect(() => {
    const timer = setTimeout(() => {
      onClose(toast.id);
    }, 4000); // auto-dismiss after 4 seconds
    return () => clearTimeout(timer);
  }, [toast.id, onClose]);

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

interface IToastContainerProps {
  toasts: IToast[];
  onCloseToast: (id: string) => void;
}

export const ToastContainer: React.FC<IToastContainerProps> = ({ toasts, onCloseToast }) => {
  return (
    <div className={styles.toastContainer}>
      {toasts.map(toast => (
        <ToastItem key={toast.id} toast={toast} onClose={onCloseToast} />
      ))}
    </div>
  );
};
