export type MpvTrackProperty = {
  type?: string;
  id?: number;
  selected?: boolean;
  "ff-index"?: number;
};

export function resolveCurrentAudioStreamIndex(
  tracks: Array<MpvTrackProperty> | null | undefined,
  currentAudioTrackId: number | null,
): number | null {
  if (!Array.isArray(tracks)) {
    return null;
  }

  const audioTracks = tracks.filter((track) => track.type === "audio");
  const activeTrack =
    audioTracks.find((track) => track.id === currentAudioTrackId) ||
    audioTracks.find((track) => track.selected === true);

  const ffIndex = activeTrack?.["ff-index"];
  return typeof ffIndex === "number" && Number.isInteger(ffIndex) && ffIndex >= 0
    ? ffIndex
    : null;
}
