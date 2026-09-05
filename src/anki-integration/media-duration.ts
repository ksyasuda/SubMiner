export function clampMediaEndTime(
  startTime: number,
  endTime: number,
  maxMediaDuration: number,
): number {
  return maxMediaDuration > 0 && endTime - startTime > maxMediaDuration
    ? startTime + maxMediaDuration
    : endTime;
}
