/** Zero or a negative cap leaves the requested end time unchanged. */
export function clampMediaEndTime(
  startTime: number,
  endTime: number,
  maxMediaDuration: number,
): number {
  return maxMediaDuration > 0 && endTime - startTime > maxMediaDuration
    ? startTime + maxMediaDuration
    : endTime;
}
