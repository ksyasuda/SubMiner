export const VIDEO_EXTENSIONS = new Set([
  'mkv',
  'mp4',
  'avi',
  'webm',
  'mov',
  'flv',
  'wmv',
  'm4v',
  'ts',
  'm2ts',
]);

export function hasVideoExtension(value: string): boolean {
  const normalized = value.trim().toLowerCase().replace(/^\./, '');
  return normalized.length > 0 && VIDEO_EXTENSIONS.has(normalized);
}
