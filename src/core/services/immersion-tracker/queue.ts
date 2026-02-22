import type { QueuedWrite } from './types';

export function enqueueWrite(
  queue: QueuedWrite[],
  write: QueuedWrite,
  queueCap: number,
): {
  dropped: number;
  queueLength: number;
} {
  let dropped = 0;
  if (queue.length >= queueCap) {
    const overflow = queue.length - queueCap + 1;
    queue.splice(0, overflow);
    dropped = overflow;
  }
  queue.push(write);
  return { dropped, queueLength: queue.length };
}
