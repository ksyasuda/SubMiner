// Buffers main-process IPC events that arrive before the renderer subscribes,
// so early events (modal opens, the first subtitle) are not dropped while the
// renderer is still booting. `subscribe` is `ipcRenderer.on` in the preload.

type Subscribe = (channel: string, handler: (payload: unknown) => void) => void;
type EmptyListener = () => void;
type PayloadedListener<T> = (payload: T) => void;

export function createIpcListenerFactories(subscribe: Subscribe) {
  // Every event fired before the first listener registers is replayed to it.
  function createQueuedIpcListener(channel: string): (listener: EmptyListener) => void {
    let count = 0;
    const listeners: EmptyListener[] = [];

    subscribe(channel, () => {
      if (listeners.length === 0) {
        count += 1;
        return;
      }
      for (const listener of listeners) listener();
    });

    return (listener) => {
      listeners.push(listener);
      while (count > 0) {
        count -= 1;
        listener();
      }
    };
  }

  // Every payload fired before the first listener registers is replayed in order.
  function createQueuedIpcListenerWithPayload<T>(
    channel: string,
    normalize: (payload: unknown) => T,
  ): (listener: PayloadedListener<T>) => void {
    const pending: T[] = [];
    const listeners: PayloadedListener<T>[] = [];

    subscribe(channel, (payloadArg) => {
      const payload = normalize(payloadArg);
      if (listeners.length === 0) {
        pending.push(payload);
        return;
      }
      for (const listener of listeners) listener(payload);
    });

    return (listener) => {
      listeners.push(listener);
      while (pending.length > 0) listener(pending.shift() as T);
    };
  }

  // Only the newest payload fired before the first listener registers is replayed.
  function createLatestValueIpcListenerWithPayload<T>(
    channel: string,
    normalize: (payload: unknown) => T,
  ): (listener: PayloadedListener<T>) => void {
    let pending: T | undefined;
    const listeners: PayloadedListener<T>[] = [];

    subscribe(channel, (payloadArg) => {
      const payload = normalize(payloadArg);
      if (listeners.length === 0) {
        pending = payload;
        return;
      }
      for (const listener of listeners) listener(payload);
    });

    return (listener) => {
      listeners.push(listener);
      if (pending !== undefined) {
        const payload = pending;
        pending = undefined;
        listener(payload);
      }
    };
  }

  return {
    createQueuedIpcListener,
    createQueuedIpcListenerWithPayload,
    createLatestValueIpcListenerWithPayload,
  };
}
