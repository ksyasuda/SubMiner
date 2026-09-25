export interface SubtitleSelectionState {
  mediaPath: string;
  tracks: { id: number; label: string }[];
  primary: number | null;
  secondary: number | null;
}

export type SubtitleSelectionRequest = Pick<
  SubtitleSelectionState,
  'mediaPath' | 'primary' | 'secondary'
>;

export function parseSubtitleSelectionRequest(value: unknown): SubtitleSelectionRequest {
  if (
    typeof value !== 'object' ||
    value === null ||
    !('mediaPath' in value) ||
    typeof value.mediaPath !== 'string' ||
    !value.mediaPath ||
    !('primary' in value) ||
    !isTrackSelection(value.primary) ||
    !('secondary' in value) ||
    !isTrackSelection(value.secondary)
  )
    throw new Error('Invalid subtitle selection.');
  if (value.primary !== null && value.primary === value.secondary)
    throw new Error('Choose different primary and secondary tracks.');
  return { mediaPath: value.mediaPath, primary: value.primary, secondary: value.secondary };
}

function isTrackSelection(value: unknown): value is number | null {
  return value === null || (typeof value === 'number' && Number.isSafeInteger(value) && value > 0);
}
