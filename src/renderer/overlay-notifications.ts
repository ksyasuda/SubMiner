import type {
  OverlayNotificationDismissPayload,
  OverlayNotificationEventPayload,
  OverlayNotificationPayload,
  OverlayNotificationPosition,
  OverlayNotificationVariant,
} from '../types';
import type { RendererContext } from './context';
import { syncOverlayMouseIgnoreState } from './overlay-mouse-ignore.js';

export const DEFAULT_OVERLAY_NOTIFICATION_TIMEOUT_MS = 3000;
export const DEFAULT_OVERLAY_NOTIFICATION_MAX_VISIBLE = 3;
export const DEFAULT_OVERLAY_NOTIFICATION_POSITION: OverlayNotificationPosition = 'top-right';
const OVERLAY_NOTIFICATION_POSITION_CLASSES = [
  'position-top-left',
  'position-top',
  'position-top-right',
] as const;

export type OverlayNotificationEntry = Required<
  Pick<OverlayNotificationPayload, 'id' | 'title' | 'persistent'>
> &
  Omit<OverlayNotificationPayload, 'id' | 'title' | 'persistent'> & {
    createdAt: number;
  };

export type OverlayNotificationStoreOptions = {
  maxVisible?: number;
  now?: () => number;
};

export type OverlayNotificationController = {
  show: (payload: OverlayNotificationPayload) => string;
  remove: (id: string) => void;
};

export function createOverlayNotificationStore(options: OverlayNotificationStoreOptions = {}) {
  const maxVisible = Math.max(1, options.maxVisible ?? DEFAULT_OVERLAY_NOTIFICATION_MAX_VISIBLE);
  const now = options.now ?? (() => Date.now());
  const entries: OverlayNotificationEntry[] = [];
  let nextId = 0;

  function visible(): OverlayNotificationEntry[] {
    const pinned = entries.filter((entry) => entry.persistent);
    const transientSlots = Math.max(0, maxVisible - pinned.length);
    const transient =
      transientSlots === 0
        ? []
        : entries.filter((entry) => !entry.persistent).slice(-transientSlots);
    return [...pinned, ...transient];
  }

  function pruneHiddenTransient(): void {
    const visibleIds = new Set(visible().map((entry) => entry.id));
    for (let index = entries.length - 1; index >= 0; index -= 1) {
      const entry = entries[index];
      if (!entry) continue;
      if (!entry.persistent && !visibleIds.has(entry.id)) {
        entries.splice(index, 1);
      }
    }
  }

  function upsert(payload: OverlayNotificationPayload): OverlayNotificationEntry {
    const id = payload.id ?? `overlay-notification-${nextId++}`;
    const existingIndex = entries.findIndex((entry) => entry.id === id);
    if (existingIndex >= 0) {
      entries.splice(existingIndex, 1);
    }
    const entry: OverlayNotificationEntry = {
      ...payload,
      id,
      title: payload.title,
      persistent: Boolean(payload.persistent),
      createdAt: now(),
    };
    entries.push(entry);
    pruneHiddenTransient();
    return entry;
  }

  function remove(id: string): void {
    const index = entries.findIndex((entry) => entry.id === id);
    if (index >= 0) {
      entries.splice(index, 1);
    }
  }

  return {
    upsert,
    remove,
    visible,
  };
}

export function overlayNotificationPositionClass(
  position: OverlayNotificationPosition | undefined,
): string {
  return `position-${position ?? DEFAULT_OVERLAY_NOTIFICATION_POSITION}`;
}

function isOverlayNotificationDismissPayload(
  payload: OverlayNotificationEventPayload,
): payload is OverlayNotificationDismissPayload {
  return 'dismiss' in payload && payload.dismiss === true;
}

export function handleOverlayNotificationEvent(
  controller: OverlayNotificationController,
  payload: OverlayNotificationEventPayload,
): string | null {
  if (isOverlayNotificationDismissPayload(payload)) {
    controller.remove(payload.id);
    return null;
  }
  return controller.show(payload);
}

function normalizeVariant(
  variant: OverlayNotificationVariant | undefined,
): OverlayNotificationVariant {
  return variant ?? 'info';
}

function normalizeImageSource(image: string | undefined): string | null {
  if (!image) return null;
  const trimmed = image.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function setInteractiveState(ctx: RendererContext, value: boolean): void {
  ctx.state.isOverOverlayNotification = value;
  syncOverlayMouseIgnoreState(ctx);
}

export function createOverlayNotificationRenderer(
  ctx: RendererContext,
  options: { onChanged?: () => void } = {},
) {
  const store = createOverlayNotificationStore();
  const timers = new Map<string, number>();
  let position: OverlayNotificationPosition = DEFAULT_OVERLAY_NOTIFICATION_POSITION;

  function clearTimer(id: string): void {
    const timer = timers.get(id);
    if (timer !== undefined) {
      window.clearTimeout(timer);
      timers.delete(id);
    }
  }

  function remove(id: string): void {
    clearTimer(id);
    store.remove(id);
    render();
  }

  function render(): void {
    const visible = store.visible();
    ctx.dom.overlayNotificationStack.replaceChildren();
    ctx.dom.overlayNotificationStack.classList.toggle('hidden', visible.length === 0);
    ctx.dom.overlayNotificationStack.classList.remove(...OVERLAY_NOTIFICATION_POSITION_CLASSES);
    ctx.dom.overlayNotificationStack.classList.add(overlayNotificationPositionClass(position));

    for (const entry of visible) {
      const imageSource = normalizeImageSource(entry.image);
      const card = document.createElement('section');
      card.className = `overlay-notification-card ${normalizeVariant(entry.variant)}${
        imageSource ? ' has-image' : ''
      }`;
      card.dataset.notificationId = entry.id;
      card.setAttribute('role', 'status');

      const leading = imageSource ? document.createElement('img') : document.createElement('span');
      leading.className = imageSource ? 'overlay-notification-image' : 'overlay-notification-icon';
      leading.setAttribute('aria-hidden', 'true');
      if (imageSource) {
        const image = leading as HTMLImageElement;
        image.src = imageSource;
        image.alt = '';
        image.decoding = 'async';
      }

      const content = document.createElement('div');
      content.className = 'overlay-notification-content';

      const title = document.createElement('div');
      title.className = 'overlay-notification-title';
      title.textContent = entry.title;
      content.append(title);

      if (entry.body && entry.body.trim().length > 0) {
        const body = document.createElement('div');
        body.className = 'overlay-notification-body';
        body.textContent = entry.body;
        content.append(body);
      }

      if (entry.actions && entry.actions.length > 0) {
        const actions = document.createElement('div');
        actions.className = 'overlay-notification-actions';
        for (const action of entry.actions) {
          const button = document.createElement('button');
          button.type = 'button';
          button.className = 'overlay-notification-action';
          button.textContent = action.label;
          button.addEventListener('click', () => {
            window.electronAPI.sendOverlayNotificationAction?.(entry.id, action.id);
            remove(entry.id);
          });
          actions.append(button);
        }
        content.append(actions);
      }

      const closeButton = document.createElement('button');
      closeButton.type = 'button';
      closeButton.className = 'overlay-notification-close';
      closeButton.setAttribute('aria-label', 'Dismiss notification');
      closeButton.textContent = '×';
      closeButton.addEventListener('click', () => remove(entry.id));

      card.append(leading, content, closeButton);
      ctx.dom.overlayNotificationStack.append(card);
    }

    if (visible.length === 0) {
      setInteractiveState(ctx, false);
    }
    options.onChanged?.();
  }

  ctx.dom.overlayNotificationStack.addEventListener('mouseenter', () => {
    setInteractiveState(ctx, true);
  });
  ctx.dom.overlayNotificationStack.addEventListener('mouseleave', () => {
    setInteractiveState(ctx, false);
  });

  function show(payload: OverlayNotificationPayload): string {
    const entry = store.upsert(payload);
    position = entry.position ?? DEFAULT_OVERLAY_NOTIFICATION_POSITION;
    clearTimer(entry.id);
    if (!entry.persistent) {
      const timeoutMs = Math.max(0, entry.timeoutMs ?? DEFAULT_OVERLAY_NOTIFICATION_TIMEOUT_MS);
      timers.set(
        entry.id,
        window.setTimeout(() => remove(entry.id), timeoutMs),
      );
    }
    render();
    return entry.id;
  }

  return {
    show,
    remove,
  };
}
