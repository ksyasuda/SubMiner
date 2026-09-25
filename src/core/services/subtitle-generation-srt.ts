export function formatTimestamp(milliseconds: number): string {
  const rounded = Math.max(0, Math.round(milliseconds));
  const hours = Math.floor(rounded / 3600000);
  const minutes = Math.floor((rounded % 3600000) / 60000);
  const seconds = Math.floor((rounded % 60000) / 1000);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')},${String(rounded % 1000).padStart(3, '0')}`;
}
