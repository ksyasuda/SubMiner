import type { OverlayNotificationEventPayload } from '../../types/notification';

export interface OverlayNotificationDeliveryDeps {
  hasReadyOverlayWindow: () => boolean;
  send: (payload: OverlayNotificationEventPayload) => void;
  maxQueuedEvents?: number;
  flushRetryDelayMs?: number;
  terminalUpdateDelayMs?: number;
  scheduleFlushRetry?: (callback: () => void, delayMs: number) => unknown;
  clearFlushRetry?: (handle: unknown) => void;
}

function getPayloadId(payload: OverlayNotificationEventPayload): string | null {
  return typeof payload.id === 'string' && payload.id.trim().length > 0 ? payload.id : null;
}

function getPayloadHistoryId(payload: OverlayNotificationEventPayload): string | null {
  if ('dismiss' in payload) {
    return null;
  }
  return typeof payload.historyId === 'string' && payload.historyId.trim().length > 0
    ? payload.historyId
    : null;
}

function isDismissPayload(
  payload: OverlayNotificationEventPayload,
): payload is Extract<OverlayNotificationEventPayload, { dismiss: true }> {
  return 'dismiss' in payload && payload.dismiss === true;
}

export function createOverlayNotificationDelivery(deps: OverlayNotificationDeliveryDeps): {
  send: (payload: OverlayNotificationEventPayload) => void;
  flush: () => void;
  getQueuedCount: () => number;
} {
  const maxQueuedEvents = Math.max(1, deps.maxQueuedEvents ?? 32);
  const flushRetryDelayMs = Math.max(1, deps.flushRetryDelayMs ?? 50);
  const terminalUpdateDelayMs = Math.max(1, deps.terminalUpdateDelayMs ?? 750);
  const queuedEvents: OverlayNotificationEventPayload[] = [];
  let flushRetryHandle: unknown = null;

  const removeQueuedPayloadsById = (id: string): void => {
    const nextEvents = queuedEvents.filter((queued) => getPayloadId(queued) !== id);
    queuedEvents.splice(0, queuedEvents.length, ...nextEvents);
  };

  const clearFlushRetry = (): void => {
    if (flushRetryHandle === null) {
      return;
    }
    deps.clearFlushRetry?.(flushRetryHandle);
    flushRetryHandle = null;
  };

  const scheduleFlushRetry = (delayMs = flushRetryDelayMs): void => {
    if (!deps.scheduleFlushRetry || flushRetryHandle !== null || queuedEvents.length === 0) {
      return;
    }
    flushRetryHandle = deps.scheduleFlushRetry(() => {
      flushRetryHandle = null;
      flush();
    }, delayMs);
  };

  const queuePayload = (payload: OverlayNotificationEventPayload): void => {
    const id = getPayloadId(payload);
    if (isDismissPayload(payload)) {
      if (id) {
        removeQueuedPayloadsById(id);
      }
      return;
    }

    if (id) {
      const payloadPersistent = payload.persistent === true;
      const payloadHistoryId = getPayloadHistoryId(payload);
      const existingIndex = queuedEvents.findIndex(
        (queued) =>
          getPayloadId(queued) === id &&
          !isDismissPayload(queued) &&
          getPayloadHistoryId(queued) === payloadHistoryId &&
          (queued.persistent === true) === payloadPersistent,
      );
      if (existingIndex >= 0) {
        queuedEvents[existingIndex] = payload;
        return;
      }
    }

    queuedEvents.push(payload);
    while (queuedEvents.length > maxQueuedEvents) {
      queuedEvents.shift();
    }
  };

  const flush = (): void => {
    if (!deps.hasReadyOverlayWindow()) {
      scheduleFlushRetry();
      return;
    }
    clearFlushRetry();
    const readyEvents = queuedEvents.splice(0, queuedEvents.length);
    const sentPersistentIds = new Set<string>();
    const deferredTerminalEvents: OverlayNotificationEventPayload[] = [];
    for (const payload of readyEvents) {
      const id = getPayloadId(payload);
      if (
        id &&
        !isDismissPayload(payload) &&
        payload.persistent !== true &&
        sentPersistentIds.has(id)
      ) {
        deferredTerminalEvents.push(payload);
        continue;
      }
      deps.send(payload);
      if (id && !isDismissPayload(payload) && payload.persistent === true) {
        sentPersistentIds.add(id);
      }
    }
    if (deferredTerminalEvents.length > 0) {
      if (!deps.scheduleFlushRetry) {
        for (const payload of deferredTerminalEvents) {
          deps.send(payload);
        }
        return;
      }
      queuedEvents.unshift(...deferredTerminalEvents);
      scheduleFlushRetry(terminalUpdateDelayMs);
    }
  };

  const send = (payload: OverlayNotificationEventPayload): void => {
    if (isDismissPayload(payload)) {
      const id = getPayloadId(payload);
      if (id) {
        removeQueuedPayloadsById(id);
      }
      if (deps.hasReadyOverlayWindow()) {
        deps.send(payload);
      }
      return;
    }

    if (!deps.hasReadyOverlayWindow()) {
      queuePayload(payload);
      return;
    }
    deps.send(payload);
  };

  return {
    send,
    flush,
    getQueuedCount: () => queuedEvents.length,
  };
}
