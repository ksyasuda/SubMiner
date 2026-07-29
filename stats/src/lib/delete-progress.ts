/**
 * Tiny external store tracking in-flight delete requests.
 *
 * Every delete goes through the API client, which registers here, so a single
 * globally mounted indicator can report progress no matter which tab, detail
 * view or overlay window started the work. Keeping the state outside React
 * means the indicator survives the deleting component unmounting mid-request
 * (navigating back after deleting a library entry, for example).
 */

export interface DeleteProgressSnapshot {
  /** Number of delete requests currently in flight. */
  count: number;
  /** Label for the oldest in-flight request, or null when idle. */
  label: string | null;
}

const IDLE_SNAPSHOT: DeleteProgressSnapshot = { count: 0, label: null };

const activeTasks = new Map<number, string>();
const listeners = new Set<() => void>();
let nextTaskId = 1;
let snapshot: DeleteProgressSnapshot = IDLE_SNAPSHOT;

function publish(): void {
  if (activeTasks.size === 0) {
    snapshot = IDLE_SNAPSHOT;
  } else {
    const [oldestLabel] = activeTasks.values();
    snapshot = { count: activeTasks.size, label: oldestLabel ?? null };
  }
  for (const listener of listeners) listener();
}

export function subscribeDeleteProgress(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

export function getDeleteProgressSnapshot(): DeleteProgressSnapshot {
  return snapshot;
}

/** Register a delete as started. Returns the function that ends it. */
export function beginDeleteTask(label: string): () => void {
  const taskId = nextTaskId++;
  activeTasks.set(taskId, label);
  publish();
  let ended = false;
  return () => {
    if (ended) return;
    ended = true;
    activeTasks.delete(taskId);
    publish();
  };
}

/** Run a delete request with the global progress indicator active. */
export async function trackDelete<T>(label: string, run: () => Promise<T>): Promise<T> {
  const end = beginDeleteTask(label);
  try {
    return await run();
  } finally {
    end();
  }
}

/** Test helper: drop every in-flight task so cases don't leak into each other. */
export function resetDeleteProgress(): void {
  activeTasks.clear();
  publish();
}
