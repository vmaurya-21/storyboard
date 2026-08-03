import * as React from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { useDialogBehavior } from '../useDialogBehavior';

const TestDialog: React.FC<{ onClose: () => void; enabled?: boolean }> = ({ onClose, enabled }) => {
  const ref = useDialogBehavior(onClose, { enabled });
  return (
    <div ref={ref} role="dialog" aria-modal="true">
      <button type="button">First</button>
      <button type="button">Last</button>
    </div>
  );
};

describe('useDialogBehavior', () => {
  it('focuses first item, traps tab, handles escape, and restores focus on unmount', () => {
    const onClose = jest.fn();

    const outside = document.createElement('button');
    outside.textContent = 'outside';
    document.body.appendChild(outside);
    outside.focus();

    const { unmount } = render(<TestDialog onClose={onClose} />);

    const first = screen.getByText('First');
    const last = screen.getByText('Last');

    expect(first).toHaveFocus();

    fireEvent.keyDown(document, { key: 'Tab', shiftKey: true });
    expect(last).toHaveFocus();

    fireEvent.keyDown(document, { key: 'Tab' });
    expect(first).toHaveFocus();

    fireEvent.keyDown(document, { key: 'Escape' });
    expect(onClose).toHaveBeenCalled();

    unmount();
    expect(outside).toHaveFocus();

    outside.remove();
  });

  it('ignores the keyboard while disabled', () => {
    const onClose = jest.fn();
    render(<TestDialog onClose={onClose} enabled={false} />);

    fireEvent.keyDown(document, { key: 'Escape' });

    expect(onClose).not.toHaveBeenCalled();
  });

  // Refcounted, so a dialog that opens a nested confirmation does not unlock
  // the page behind it when only the confirmation closes.
  it('locks page scroll until the last dialog unmounts', () => {
    const outer = render(<TestDialog onClose={jest.fn()} />);
    expect(document.body.style.overflow).toBe('hidden');

    const inner = render(<TestDialog onClose={jest.fn()} />);
    expect(document.body.style.overflow).toBe('hidden');

    inner.unmount();
    expect(document.body.style.overflow).toBe('hidden');

    outer.unmount();
    expect(document.body.style.overflow).toBe('');
  });
});
