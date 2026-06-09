import type { OverlayNotificationAction, OverlayNotificationVariant } from '../types';
import type { RendererContext } from './context';
import type { OverlayNotificationEntry } from './overlay-notifications.js';
import { syncOverlayMouseIgnoreState } from './overlay-mouse-ignore.js';

export const DEFAULT_OVERLAY_NOTIFICATION_HISTORY_MAX = 200;

const OVERLAY_NOTIFICATION_HISTORY_VARIANT_CLASSES = [
  'info',
  'progress',
  'success',
  'warning',
  'error',
] as const;

export type OverlayNotificationHistoryEntry = {
  id: string;
  title: string;
  body?: string;
  image?: string;
  variant: OverlayNotificationVariant;
  actions?: OverlayNotificationAction[];
  createdAt: number;
  updatedAt: number;
};

export type OverlayNotificationHistoryStoreOptions = {
  max?: number;
  now?: () => number;
};

function normalizeVariant(
  variant: OverlayNotificationVariant | undefined,
): OverlayNotificationVariant {
  return variant ?? 'info';
}

/**
 * Session-scoped log of every overlay notification that was shown. Entries are keyed by historyId
 * when provided, otherwise by live notification id. Reusing a key updates the record in place;
 * distinct history keys preserve separate visible events. Ordering is by first-seen so the panel can
 * render newest-first.
 */
export function createOverlayNotificationHistoryStore(
  options: OverlayNotificationHistoryStoreOptions = {},
) {
  const max = Math.max(1, options.max ?? DEFAULT_OVERLAY_NOTIFICATION_HISTORY_MAX);
  const now = options.now ?? (() => Date.now());
  const entries = new Map<string, OverlayNotificationHistoryEntry>();

  function record(entry: OverlayNotificationEntry): OverlayNotificationHistoryEntry {
    const timestamp = now();
    const historyId = entry.historyId?.trim() || entry.id;
    const existing = entries.get(historyId);
    const next: OverlayNotificationHistoryEntry = {
      id: historyId,
      title: entry.title,
      body: entry.body,
      image: entry.image,
      variant: normalizeVariant(entry.variant),
      actions: entry.actions?.map((action) => ({ ...action })),
      createdAt: existing?.createdAt ?? timestamp,
      updatedAt: timestamp,
    };
    // Setting an existing key keeps its original insertion slot, so an in-place update (same id,
    // new body) refreshes content without jumping the entry to the top of the panel.
    entries.set(historyId, next);
    while (entries.size > max) {
      const oldest = entries.keys().next().value;
      if (oldest === undefined) break;
      entries.delete(oldest);
    }
    return next;
  }

  function remove(id: string): void {
    entries.delete(id);
  }

  function clear(): void {
    entries.clear();
  }

  function list(): OverlayNotificationHistoryEntry[] {
    // Newest first.
    return [...entries.values()].reverse();
  }

  function size(): number {
    return entries.size;
  }

  return { record, remove, clear, list, size };
}

export type OverlayNotificationHistorySide = 'left' | 'right';

/**
 * The history panel slides in from the same edge the notifications use: left when notifications are
 * top-left, right otherwise (including center). We read the live position class off the notification
 * stack so the panel always tracks the configured/last-used position.
 */
export function resolveHistorySideFromStack(stack: Element): OverlayNotificationHistorySide {
  return stack.classList.contains('position-top-left') ? 'left' : 'right';
}

export function createOverlayNotificationHistoryPanel(
  ctx: RendererContext,
  options: { onChanged?: () => void } = {},
) {
  const store = createOverlayNotificationHistoryStore();
  const panel = ctx.dom.overlayNotificationHistory;
  const list = panel.querySelector<HTMLUListElement>('.notification-history-list');
  const empty = panel.querySelector<HTMLElement>('.notification-history-empty');
  const clearButton = panel.querySelector<HTMLButtonElement>('.notification-history-clear');
  const closeButton = panel.querySelector<HTMLButtonElement>('.notification-history-close');
  let open = false;

  function setInteractive(value: boolean): void {
    ctx.state.isOverNotificationHistory = value;
    syncOverlayMouseIgnoreState(ctx);
  }

  function applySide(): void {
    const side = resolveHistorySideFromStack(ctx.dom.overlayNotificationStack);
    panel.classList.toggle('side-left', side === 'left');
    panel.classList.toggle('side-right', side === 'right');
  }

  function formatTime(timestamp: number): string {
    try {
      return new Date(timestamp).toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      });
    } catch {
      return '';
    }
  }

  function buildItem(entry: OverlayNotificationHistoryEntry): HTMLLIElement {
    const item = document.createElement('li');
    item.className = 'notification-history-item';
    for (const variant of OVERLAY_NOTIFICATION_HISTORY_VARIANT_CLASSES) {
      item.classList.toggle(variant, variant === entry.variant);
    }
    item.dataset.notificationId = entry.id;

    const trimmedImage = entry.image?.trim();
    const leading = trimmedImage ? document.createElement('img') : document.createElement('span');
    leading.className = trimmedImage ? 'notification-history-thumb' : 'notification-history-icon';
    leading.setAttribute('aria-hidden', 'true');
    if (trimmedImage) {
      const image = leading as HTMLImageElement;
      image.src = trimmedImage;
      image.alt = '';
      image.decoding = 'async';
    }

    const content = document.createElement('div');
    content.className = 'notification-history-content';

    const title = document.createElement('div');
    title.className = 'notification-history-item-title';
    title.textContent = entry.title;
    content.append(title);

    if (entry.body && entry.body.trim().length > 0) {
      const body = document.createElement('div');
      body.className = 'notification-history-item-body';
      body.textContent = entry.body;
      content.append(body);
    }

    const time = document.createElement('time');
    time.className = 'notification-history-time';
    time.dateTime = new Date(entry.createdAt).toISOString();
    time.textContent = formatTime(entry.createdAt);
    content.append(time);

    if (entry.actions && entry.actions.length > 0) {
      const actions = document.createElement('div');
      actions.className = 'notification-history-actions';
      for (const action of entry.actions) {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'notification-history-action';
        button.textContent = action.label;
        button.addEventListener('click', () => {
          window.electronAPI.sendOverlayNotificationAction?.(entry.id, action.id, {
            noteId: action.noteId,
          });
        });
        actions.append(button);
      }
      content.append(actions);
    }

    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'notification-history-remove';
    remove.setAttribute('aria-label', 'Remove from history');
    remove.textContent = '×';
    remove.addEventListener('click', () => {
      store.remove(entry.id);
      render();
    });

    item.append(leading, content, remove);
    return item;
  }

  function render(): void {
    if (!list || !empty) return;
    const entries = store.list();
    list.replaceChildren(...entries.map(buildItem));
    empty.classList.toggle('hidden', entries.length > 0);
    if (clearButton) clearButton.disabled = entries.length === 0;
    options.onChanged?.();
  }

  function setOpen(next: boolean): void {
    if (open === next) return;
    open = next;
    ctx.state.notificationHistoryOpen = next;
    if (next) {
      applySide();
      render();
    }
    panel.classList.toggle('open', next);
    panel.setAttribute('aria-hidden', next ? 'false' : 'true');
    setInteractive(next);
    options.onChanged?.();
  }

  clearButton?.addEventListener('click', () => {
    store.clear();
    render();
  });
  closeButton?.addEventListener('click', () => setOpen(false));
  panel.addEventListener('mouseenter', () => {
    if (open) setInteractive(true);
  });
  panel.addEventListener('mouseleave', () => setInteractive(false));
  applySide();

  function record(entry: OverlayNotificationEntry): void {
    store.record(entry);
    if (open) render();
  }

  function toggle(): void {
    setOpen(!open);
  }

  return {
    record,
    toggle,
    open: () => setOpen(true),
    close: () => setOpen(false),
    isOpen: () => open,
    syncSide: applySide,
  };
}
