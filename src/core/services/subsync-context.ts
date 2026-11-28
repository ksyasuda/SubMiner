import * as path from 'path';
import { fileURLToPath } from 'url';
import { SubsyncManualRunRequest } from '../../types';
import { getTrackById, MpvTrack, SubsyncContext } from '../../subsync/utils';
import { isRemoteMediaPath } from '../../jimaku/utils';
import { type MpvClientLike, type SubtitleSlot } from './subsync-shared';

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number') {
    return Number.isInteger(value) ? value : null;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed.length) return null;
    const parsed = Number(trimmed);
    return Number.isInteger(parsed) && String(parsed) === trimmed ? parsed : null;
  }
  return null;
}

export function parseOptionalNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

export function normalizeTrackIds(tracks: unknown[]): MpvTrack[] {
  return tracks.map((track) => {
    if (!track || typeof track !== 'object') return track as MpvTrack;
    const typed = track as MpvTrack & { id?: unknown };
    const parsedId = parseTrackId(typed.id);
    if (parsedId === null) {
      const { id: _ignored, ...rest } = typed;
      return rest as MpvTrack;
    }
    return { ...typed, id: parsedId };
  });
}

function getSourceTrackIdentity(track: MpvTrack): string {
  if (
    track.external &&
    typeof track['external-filename'] === 'string' &&
    track['external-filename'].length > 0
  ) {
    return `external:${normalizeSubtitlePathForCompare(track['external-filename'])}`;
  }
  if (typeof track.id === 'number') {
    return `id:${track.id}`;
  }
  if (typeof track.title === 'string' && track.title.length > 0) {
    return `title:${track.title.toLowerCase()}`;
  }
  return 'unknown';
}

function isPinned(track: MpvTrack, pinnedIds: Set<number>): boolean {
  return typeof track.id === 'number' && pinnedIds.has(track.id);
}

// Pinned tracks (the active primary/secondary) always survive, even when two of
// them point at the same file; only unpinned duplicates are collapsed.
function dedupeSubtitleTracks(tracks: MpvTrack[], pinnedIds: Set<number>): MpvTrack[] {
  const pinnedIdentities = new Set(
    tracks.filter((track) => isPinned(track, pinnedIds)).map(getSourceTrackIdentity),
  );
  const winners = new Map<string, MpvTrack>();
  for (const track of tracks) {
    if (isPinned(track, pinnedIds)) continue;
    const identity = getSourceTrackIdentity(track);
    if (pinnedIdentities.has(identity)) continue;
    const existing = winners.get(identity);
    if (!existing || (track.selected && !existing.selected)) {
      winners.set(identity, track);
    }
  }
  const kept = new Set(winners.values());
  return tracks.filter((track) => isPinned(track, pinnedIds) || kept.has(track));
}

export function resolveLocalSubtitlePath(
  rawPath: string,
  workingDirectory?: string | null,
): string | null {
  if (!rawPath) return null;
  let local = rawPath;
  if (/^file:\/\//i.test(local)) {
    try {
      local = fileURLToPath(local);
    } catch {
      return null;
    }
  } else if (/^[a-z][a-z\d+.-]*:\/\//i.test(local)) {
    return null;
  }
  if (!path.isAbsolute(local) && workingDirectory) {
    local = path.resolve(workingDirectory, local);
  }
  return local;
}

export function normalizeMediaVideoPath(rawPath: string, workingDirectory?: string | null): string {
  return resolveLocalSubtitlePath(rawPath, workingDirectory) ?? rawPath;
}

export async function gatherSubsyncContext(client: MpvClientLike): Promise<SubsyncContext> {
  const [videoPathRaw, sidRaw, secondarySidRaw, trackListRaw, workingDirRaw, durationRaw] =
    await Promise.all([
      client.requestProperty('path'),
      client.requestProperty('sid'),
      client.requestProperty('secondary-sid'),
      client.requestProperty('track-list'),
      client.requestProperty('working-directory').catch(() => null),
      client.requestProperty('duration').catch(() => null),
    ]);

  const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
  if (!videoPath) {
    throw new Error('No video is currently loaded');
  }

  const workingDirectory =
    typeof workingDirRaw === 'string' && workingDirRaw.trim().length > 0 ? workingDirRaw : null;
  const resolvedVideoPath = normalizeMediaVideoPath(videoPath, workingDirectory);

  const tracks = Array.isArray(trackListRaw) ? normalizeTrackIds(trackListRaw as MpvTrack[]) : [];
  const subtitleTracks = tracks.filter((track) => track.type === 'sub');
  const sid = parseTrackId(sidRaw);
  const secondarySid = parseTrackId(secondarySidRaw);

  const primaryTrack = subtitleTracks.find((track) => track.id === sid);
  if (!primaryTrack) {
    throw new Error('No active subtitle track found');
  }

  const secondaryTrack = subtitleTracks.find((track) => track.id === secondarySid) ?? null;
  const usableTracks = subtitleTracks.filter((track) => {
    if (typeof track.id !== 'number') return false;
    if (!track.external) return true;
    const filename = track['external-filename'];
    return typeof filename === 'string' && filename.length > 0;
  });

  // Resolve the active audio stream's ff-index within the video file. An external audio track
  // (--audio-file, auto-loaded .mka dub) has an ff-index inside its own file, so mapping it as
  // `0:N` on the video would pick an unrelated stream; use the video's default audio instead.
  const selectedAudio = tracks.find((track) => track.type === 'audio' && track.selected === true);
  let audioStreamIndex: number | null = null;
  if (!selectedAudio?.external) {
    audioStreamIndex = client.currentAudioStreamIndex;
    if (audioStreamIndex === null && typeof selectedAudio?.['ff-index'] === 'number') {
      audioStreamIndex = selectedAudio['ff-index'];
    }
  }

  return {
    videoPath: resolvedVideoPath,
    workingDirectory,
    durationSec: parseOptionalNumber(durationRaw),
    primaryTrack,
    secondaryTrack,
    subtitleTracks: dedupeSubtitleTracks(
      usableTracks,
      new Set([sid, secondarySid].filter((id): id is number => typeof id === 'number')),
    ),
    audioStreamIndex,
  };
}

export async function isPlaybackSessionUnchanged(
  client: MpvClientLike,
  context: SubsyncContext,
  targetTrack: MpvTrack,
  slot: SubtitleSlot,
  signal?: AbortSignal,
): Promise<boolean> {
  const unchanged = await isPlaybackSessionUnchangedNow(client, context, targetTrack, slot);
  // Checked after the awaits: a run cancelled meanwhile must not commit anything.
  return unchanged && !signal?.aborted;
}

async function isPlaybackSessionUnchangedNow(
  client: MpvClientLike,
  context: SubsyncContext,
  targetTrack: MpvTrack,
  slot: SubtitleSlot,
): Promise<boolean> {
  if (!client.connected) return false;
  try {
    const slotProperty = slot === 'secondary' ? 'secondary-sid' : 'sid';
    const initialSlotTrackId =
      slot === 'secondary'
        ? (context.secondaryTrack?.id ?? null)
        : (context.primaryTrack.id ?? null);
    // Always compare the target's source too: mpv reuses the highest freed id, so a track
    // removed and replaced mid-sync can keep the same id (and the same slot id).
    const needsTrackIdentityCheck = typeof targetTrack.id === 'number';
    const [videoPathRaw, workingDirRaw, slotTrackRaw, trackListRaw] = await Promise.all([
      client.requestProperty('path'),
      client.requestProperty('working-directory').catch(() => null),
      client.requestProperty(slotProperty),
      needsTrackIdentityCheck
        ? client.requestProperty('track-list').catch(() => null)
        : Promise.resolve(null),
    ]);
    const currentRawPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
    if (!currentRawPath) return false;
    const currentWorkingDir =
      typeof workingDirRaw === 'string' && workingDirRaw.trim().length > 0 ? workingDirRaw : null;
    const currentVideoPath = normalizeMediaVideoPath(currentRawPath, currentWorkingDir);
    if (
      normalizeSubtitlePathForCompare(currentVideoPath) !==
      normalizeSubtitlePathForCompare(context.videoPath)
    ) {
      return false;
    }
    const currentSlotId = parseTrackId(slotTrackRaw);
    if (currentSlotId !== initialSlotTrackId) {
      return false;
    }
    if (needsTrackIdentityCheck) {
      const tracks = Array.isArray(trackListRaw)
        ? normalizeTrackIds(trackListRaw as MpvTrack[])
        : [];
      const currentTarget = tracks.find(
        (track) => track.type === 'sub' && track.id === targetTrack.id,
      );
      if (
        !currentTarget ||
        getSourceTrackIdentity(currentTarget) !== getSourceTrackIdentity(targetTrack)
      ) {
        return false;
      }
    }
    return true;
  } catch {
    return false;
  }
}

// Windows local paths are case-insensitive and mpv may echo them with different
// separators; on POSIX, backslashes are valid filename characters and must stay distinct.
export function normalizeSubtitlePathForCompare(value: string): string {
  if (process.platform === 'win32' && !isRemoteMediaPath(value)) {
    return value.replace(/\\/g, '/').toLowerCase();
  }
  return value;
}

export function resolveTargetTrack(
  request: SubsyncManualRunRequest,
  context: SubsyncContext,
): MpvTrack | null {
  if (request.targetTrackId === undefined || request.targetTrackId === null) {
    return context.primaryTrack;
  }
  return getTrackById(context.subtitleTracks, request.targetTrackId);
}

// Retiming the secondary track must not steal the primary slot: the synced file
// goes back where the out-of-sync one was.
export function resolveTargetSlot(targetTrack: MpvTrack, context: SubsyncContext): SubtitleSlot {
  if (typeof targetTrack.id !== 'number') return 'primary';
  if (targetTrack.id === context.primaryTrack.id) return 'primary';
  if (context.secondaryTrack && targetTrack.id === context.secondaryTrack.id) return 'secondary';
  return 'primary';
}
