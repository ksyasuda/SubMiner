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
const OVERLAY_NOTIFICATION_VARIANT_CLASSES = [
  'info',
  'progress',
  'success',
  'warning',
  'error',
] as const;
// Matches the `.leaving` animation duration in style.css; the fallback timer guards
// against `animationend` never firing (e.g. element detached or reduced-motion).
const OVERLAY_NOTIFICATION_EXIT_FALLBACK_MS = 260;

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

function hasElementClass(element: Element | undefined, className: string): boolean {
  if (!element) return false;
  const legacyClassName = (element as { className?: unknown }).className;
  return (
    element.classList.contains(className) ||
    (typeof legacyClassName === 'string' && legacyClassName.split(/\s+/).includes(className))
  );
}

function isNotificationCardIcon(element: Element | undefined): boolean {
  return hasElementClass(element, 'overlay-notification-icon');
}

function isNotificationCardContent(element: Element | undefined): element is HTMLElement {
  return hasElementClass(element, 'overlay-notification-content');
}

function isNotificationCardCloseButton(element: Element | undefined): boolean {
  return hasElementClass(element, 'overlay-notification-close');
}

function hasExplicitNotificationActions(entry: OverlayNotificationEntry): boolean {
  return (entry.actions?.length ?? 0) > 0;
}

export function createOverlayNotificationRenderer(
  ctx: RendererContext,
  options: { onChanged?: () => void; onShow?: (entry: OverlayNotificationEntry) => void } = {},
) {
  const store = createOverlayNotificationStore();
  const timers = new Map<string, number>();
  // Live card elements keyed by notification id so re-renders reuse them: the enter
  // animation only plays for freshly created cards instead of replaying on every render.
  const cards = new Map<string, HTMLElement>();
  const leaving = new Set<string>();
  let position: OverlayNotificationPosition = DEFAULT_OVERLAY_NOTIFICATION_POSITION;

  function clearTimer(id: string): void {
    const timer = timers.get(id);
    if (timer !== undefined) {
      window.clearTimeout(timer);
      timers.delete(id);
    }
  }

  function commitExit(id: string, card: HTMLElement): void {
    if (!leaving.has(id)) return;
    leaving.delete(id);
    cards.delete(id);
    card.remove();
    if (cards.size === 0) {
      ctx.dom.overlayNotificationStack.classList.add('hidden');
      setInteractiveState(ctx, false);
    }
    options.onChanged?.();
  }

  function beginExit(id: string, card: HTMLElement): void {
    if (leaving.has(id)) return;
    leaving.add(id);
    card.classList.remove('entering');
    card.classList.add('leaving');
    const finalize = () => {
      window.clearTimeout(fallback);
      commitExit(id, card);
    };
    const fallback = window.setTimeout(finalize, OVERLAY_NOTIFICATION_EXIT_FALLBACK_MS);
    card.addEventListener(
      'animationend',
      (event) => {
        if ((event as AnimationEvent).animationName?.startsWith('overlay-notification-leave')) {
          finalize();
        }
      },
      { once: true },
    );
  }

  function markEnterComplete(card: HTMLElement): void {
    card.classList.remove('entering');
  }

  function watchEnterAnimation(card: HTMLElement): void {
    if (typeof window === 'undefined') {
      return;
    }
    const fallback = window.setTimeout(() => markEnterComplete(card), 320);
    card.addEventListener(
      'animationend',
      (event) => {
        if ((event as AnimationEvent).animationName?.startsWith('overlay-notification-enter')) {
          window.clearTimeout(fallback);
          markEnterComplete(card);
        }
      },
      { once: true },
    );
  }

  function appendCardIfNeeded(card: HTMLElement): void {
    if (Array.prototype.includes.call(ctx.dom.overlayNotificationStack.children, card)) {
      return;
    }
    ctx.dom.overlayNotificationStack.append(card);
  }

  function bindInteractiveControlHover(element: HTMLElement): void {
    element.addEventListener('mouseenter', () => setInteractiveState(ctx, true));
    element.addEventListener('mouseleave', () => setInteractiveState(ctx, false));
  }

  function remove(id: string): void {
    clearTimer(id);
    store.remove(id);
    const card = cards.get(id);
    if (card) {
      beginExit(id, card);
    } else {
      render();
    }
  }

  function populateContent(content: HTMLElement, entry: OverlayNotificationEntry): void {
    content.className = 'overlay-notification-content';

    const title = document.createElement('div');
    title.className = 'overlay-notification-title';
    title.textContent = entry.title;
    const children: HTMLElement[] = [title];

    if (entry.body && entry.body.trim().length > 0) {
      const body = document.createElement('div');
      body.className = 'overlay-notification-body';
      body.textContent = entry.body;
      children.push(body);
    }

    if (entry.actions && entry.actions.length > 0) {
      const actions = document.createElement('div');
      actions.className = 'overlay-notification-actions';
      for (const action of entry.actions) {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'overlay-notification-action';
        button.textContent = action.label;
        bindInteractiveControlHover(button);
        button.addEventListener('click', () => {
          window.electronAPI.sendOverlayNotificationAction?.(entry.id, action.id, {
            noteId: action.noteId,
          });
          remove(entry.id);
        });
        actions.append(button);
      }
      children.push(actions);
    }
    content.replaceChildren(...children);
  }

  function createContent(entry: OverlayNotificationEntry): HTMLElement {
    const content = document.createElement('div');
    populateContent(content, entry);
    return content;
  }

  function populateCard(card: HTMLElement, entry: OverlayNotificationEntry): void {
    const imageSource = normalizeImageSource(entry.image);
    card.classList.add('overlay-notification-card');
    for (const variant of OVERLAY_NOTIFICATION_VARIANT_CLASSES) {
      card.classList.toggle(variant, variant === normalizeVariant(entry.variant));
    }
    card.classList.toggle('has-image', Boolean(imageSource));
    card.dataset.notificationId = entry.id;
    card.setAttribute('role', 'status');

    const leadingNode = card.children[0];
    const contentNode = card.children[1];
    const closeNode = card.children[2];
    if (
      leadingNode &&
      contentNode &&
      closeNode &&
      !imageSource &&
      !entry.actions?.length &&
      isNotificationCardIcon(leadingNode) &&
      isNotificationCardContent(contentNode) &&
      isNotificationCardCloseButton(closeNode)
    ) {
      populateContent(contentNode, entry);
      return;
    }

    const leadingEl = imageSource ? document.createElement('img') : document.createElement('span');
    leadingEl.className = imageSource ? 'overlay-notification-image' : 'overlay-notification-icon';
    leadingEl.setAttribute('aria-hidden', 'true');
    if (imageSource) {
      const image = leadingEl as HTMLImageElement;
      image.src = imageSource;
      image.alt = '';
      image.decoding = 'async';
    }

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'overlay-notification-close';
    closeButton.setAttribute('aria-label', 'Dismiss notification');
    closeButton.textContent = '×';
    if (hasExplicitNotificationActions(entry)) {
      bindInteractiveControlHover(closeButton);
    }
    closeButton.addEventListener('click', () => remove(entry.id));

    card.replaceChildren(leadingEl, createContent(entry), closeButton);
  }

  function render(): void {
    const visible = store.visible();
    const visibleIds = new Set(visible.map((entry) => entry.id));
    const hasInteractiveCard = visible.some(hasExplicitNotificationActions);
    ctx.dom.overlayNotificationStack.classList.toggle(
      'hidden',
      visible.length === 0 && leaving.size === 0,
    );
    ctx.dom.overlayNotificationStack.classList.remove(...OVERLAY_NOTIFICATION_POSITION_CLASSES);
    ctx.dom.overlayNotificationStack.classList.add(overlayNotificationPositionClass(position));

    // Cards that vanished from the store without an explicit remove() (e.g. pruned when
    // over the visible cap) still need to animate out.
    for (const [id, card] of cards) {
      if (!visibleIds.has(id)) {
        beginExit(id, card);
      }
    }

    for (const entry of visible) {
      let card = cards.get(entry.id);
      if (card && leaving.has(entry.id)) {
        // The card was animating out but has been re-shown: cancel the exit.
        leaving.delete(entry.id);
        card.classList.remove('leaving');
      }
      if (!card) {
        card = document.createElement('section');
        card.classList.add('entering');
        watchEnterAnimation(card);
        cards.set(entry.id, card);
      }
      populateCard(card, entry);
      appendCardIfNeeded(card);
    }

    if (visible.length === 0 && leaving.size === 0) {
      setInteractiveState(ctx, false);
    } else if (!hasInteractiveCard && ctx.state.isOverOverlayNotification) {
      setInteractiveState(ctx, false);
    }
    options.onChanged?.();
  }

  ctx.dom.overlayNotificationStack.addEventListener('mouseenter', () => {
    if (store.visible().some(hasExplicitNotificationActions)) {
      setInteractiveState(ctx, true);
    }
  });
  ctx.dom.overlayNotificationStack.addEventListener('mouseleave', () => {
    setInteractiveState(ctx, false);
  });

  function show(payload: OverlayNotificationPayload): string {
    const entry = store.upsert(payload);
    position = entry.position ?? DEFAULT_OVERLAY_NOTIFICATION_POSITION;
    options.onShow?.(entry);
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
