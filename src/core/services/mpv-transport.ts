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
