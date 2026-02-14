export function getMpvReconnectDelay(
  attempt: number,
  hasConnectedOnce: boolean,
): number {
  if (hasConnectedOnce) {
    if (attempt < 2) {
      return 1000;
    }
    if (attempt < 4) {
      return 2000;
    }
    if (attempt < 7) {
      return 5000;
    }
    return 10000;
  }

  if (attempt < 2) {
    return 200;
  }
  if (attempt < 4) {
    return 500;
  }
  if (attempt < 6) {
    return 1000;
  }
  return 2000;
}

export interface MpvReconnectSchedulerDeps {
  attempt: number;
  hasConnectedOnce: boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  onReconnectAttempt: (attempt: number, delay: number) => void;
  connect: () => void;
}

export function scheduleMpvReconnect(
  deps: MpvReconnectSchedulerDeps,
): number {
  const reconnectTimer = deps.getReconnectTimer();
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
  }
  const delay = getMpvReconnectDelay(deps.attempt, deps.hasConnectedOnce);
  deps.setReconnectTimer(
    setTimeout(() => {
      deps.onReconnectAttempt(deps.attempt + 1, delay);
      deps.connect();
    }, delay),
  );
  return deps.attempt + 1;
}
