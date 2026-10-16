import { useEffect } from 'react';

/**
 * Desktop Keyboard Shortcuts Hook & Definitions
 * Milestone M4 (Requirement R4: Offline & Desktop Reliability)
 */

export interface ShortcutHandlers {
  onSave?: () => void;
  onPrint?: () => void;
  onNewDraft?: () => void;
  onEscape?: () => void;
}

export interface ShortcutDefinition {
  id: string;
  key: string;
  ctrlOrCmd: boolean;
  action: 'save' | 'print' | 'newDraft' | 'escape';
  description: string;
}

export const SHORTCUT_DEFINITIONS: Record<string, ShortcutDefinition> = {
  SAVE: {
    id: 'SAVE',
    key: 's',
    ctrlOrCmd: true,
    action: 'save',
    description: 'ذخیره پیش‌فاکتور (در صفحه ویرایش) یا ذخیره PDF (در صفحه پیش‌نمایش)',
  },
  PRINT: {
    id: 'PRINT',
    key: 'p',
    ctrlOrCmd: true,
    action: 'print',
    description: 'مشاهده پیش‌نمایش چاپ (در صفحه ویرایش) یا چاپ / خروجی PDF (در پیش‌نمایش)',
  },
  NEW_DRAFT: {
    id: 'NEW_DRAFT',
    key: 'n',
    ctrlOrCmd: true,
    action: 'newDraft',
    description: 'ایجاد پیش‌فاکتور جدید و بازنشانی فرم به حالت پیش‌نویس خالی',
  },
  ESCAPE: {
    id: 'ESCAPE',
    key: 'Escape',
    ctrlOrCmd: false,
    action: 'escape',
    description: 'بستن پنجره‌های باز، مودال‌ها یا بازگشت از پیش‌نمایش چاپ به ویرایش',
  },
};

export type ShortcutAction = 'save' | 'print' | 'newDraft' | 'escape' | null;

/**
 * Pure function to match a keyboard event against registered shortcut definitions.
 */
export function matchKeyboardShortcut(event: {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
}): ShortcutAction {
  const isCtrlOrCmd = Boolean(event.ctrlKey || event.metaKey);
  const keyLower = event.key ? event.key.toLowerCase() : '';

  if (isCtrlOrCmd && keyLower === 's') {
    return 'save';
  }
  if (isCtrlOrCmd && keyLower === 'p') {
    return 'print';
  }
  if (isCtrlOrCmd && keyLower === 'n') {
    return 'newDraft';
  }
  if (event.key === 'Escape') {
    return 'escape';
  }
  return null;
}

/**
 * Handles a keyboard event, executes the matching handler, and prevents default behavior.
 */
export function handleKeyboardEvent(
  event: {
    key: string;
    ctrlKey?: boolean;
    metaKey?: boolean;
    preventDefault?: () => void;
  },
  handlers: ShortcutHandlers
): boolean {
  const action = matchKeyboardShortcut(event);
  if (!action) return false;

  switch (action) {
    case 'save':
      if (handlers.onSave) {
        event.preventDefault?.();
        handlers.onSave();
        return true;
      }
      break;
    case 'print':
      if (handlers.onPrint) {
        event.preventDefault?.();
        handlers.onPrint();
        return true;
      }
      break;
    case 'newDraft':
      if (handlers.onNewDraft) {
        event.preventDefault?.();
        handlers.onNewDraft();
        return true;
      }
      break;
    case 'escape':
      if (handlers.onEscape) {
        event.preventDefault?.();
        handlers.onEscape();
        return true;
      }
      break;
  }
  return false;
}

/**
 * React hook that binds a global window keydown listener for desktop shortcuts.
 */
export function useKeyboardShortcuts(handlers: ShortcutHandlers, enabled: boolean = true): void {
  useEffect(() => {
    if (!enabled) return;

    const onKeyDown = (e: KeyboardEvent) => {
      handleKeyboardEvent(e, handlers);
    };

    window.addEventListener('keydown', onKeyDown);
    return () => {
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [handlers, enabled]);
}
