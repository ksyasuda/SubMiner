import * as path from 'path';
import { SubsyncManualPayload, SubsyncManualRunRequest, SubsyncResult } from '../../types';
import {
  CommandResult,
  fileExists,
  formatTrackLabel,
  getTrackById,
  MpvTrack,
  resolveLocalMediaPath,
  runCommand,
  summarizeCommandFailure,
  SubsyncContext,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import {
  resolveExecutable,
  SubsyncExecutable,
  SUBSYNC_EXECUTABLE_NAMES,
} from '../../subsync/executables';
import {
  cleanupTemporaryFile,
  extractSubtitleTrackToFile,
  FileExtractionResult,
} from './subsync-extract';
import { convertSubtitleForAlass } from './subsync-alass-input';
import { resolveMpvHttpHeaders, ResolvedMpvHttpHeaders } from './mpv-http-headers';
import { isRemoteMediaPath } from '../../jimaku/utils';
import { createLogger } from '../../logger';

const logger = createLogger('subsync');

type SubtitleSlot = 'primary' | 'secondary';

const SYNCED_TRACK_LOOKUP_ATTEMPTS = 5;
const SYNCED_TRACK_LOOKUP_RETRY_MS = 100;

interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

interface SubsyncCoreDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedConfig: () => SubsyncResolvedConfig;
}

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

function normalizeTrackIds(tracks: unknown[]): MpvTrack[] {
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
    return `external:${track['external-filename'].toLowerCase()}`;
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

export interface TriggerSubsyncFromConfigDeps extends SubsyncCoreDeps {
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  runWithSubsyncSpinner: <T>(task: () => Promise<T>) => Promise<T>;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

function getMpvClientForSubsync(deps: SubsyncCoreDeps): MpvClientLike {
  const client = deps.getMpvClient();
  if (!client || !client.connected) {
    throw new Error('MPV not connected');
  }
  return client;
}

async function gatherSubsyncContext(client: MpvClientLike): Promise<SubsyncContext> {
  const [videoPathRaw, sidRaw, secondarySidRaw, trackListRaw] = await Promise.all([
    client.requestProperty('path'),
    client.requestProperty('sid'),
    client.requestProperty('secondary-sid'),
    client.requestProperty('track-list'),
  ]);

  // A dropped file is reported as a `file://` URL; alass, ffsubsync and ffmpeg
  // all need the plain path, and the stream checks must not read it as remote.
  const videoPath = typeof videoPathRaw === 'string' ? resolveLocalMediaPath(videoPathRaw) : '';
  if (!videoPath) {
    throw new Error('No video is currently loaded');
  }

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

  return {
    videoPath,
    primaryTrack,
    secondaryTrack,
    subtitleTracks: dedupeSubtitleTracks(
      usableTracks,
      new Set([sid, secondarySid].filter((id): id is number => typeof id === 'number')),
    ),
    audioStreamIndex: client.currentAudioStreamIndex,
  };
}

/**
 * Resolve a sync tool, discovering it on PATH (plus the usual install prefixes)
 * when the config leaves it unset.
 */
function resolveSubsyncExecutable(configuredPath: string, name: SubsyncExecutable): string {
  const resolved = resolveExecutable(configuredPath, SUBSYNC_EXECUTABLE_NAMES[name]);
  if (resolved) return resolved;

  throw new Error(
    configuredPath.trim()
      ? `Configured ${name} executable not found: ${configuredPath}`
      : `Could not find ${name}. Install it or set subsync.${name}_path in your config.`,
  );
}

function buildRetimedPath(subPath: string, replace: boolean): string {
  if (replace) return subPath;
  const parsed = path.parse(subPath);
  return path.join(parsed.dir, `${parsed.name}_retimed${parsed.ext || '.srt'}`);
}

async function runAlassSync(
  alassPath: string,
  referenceFile: string,
  inputSubtitlePath: string,
  outputPath: string,
): Promise<CommandResult> {
  return runCommand(alassPath, [referenceFile, inputSubtitlePath, outputPath]);
}

async function runFfsubsyncSync(
  ffsubsyncPath: string,
  videoPath: string,
  inputSubtitlePath: string,
  outputPath: string,
  audioStreamIndex: number | null,
): Promise<CommandResult> {
  const args = [videoPath, '-i', inputSubtitlePath, '-o', outputPath];
  if (audioStreamIndex !== null) {
    args.push('--reference-stream', `0:${audioStreamIndex}`);
  }
  return runCommand(ffsubsyncPath, args);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// mpv may echo the path back with different separators, and Windows paths are
// case-insensitive, so compare normalized forms instead of raw strings.
function normalizeSubtitlePathForCompare(value: string): string {
  const normalized = value.replace(/\\/g, '/');
  return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
}

async function findAddedSubtitleTrackId(
  client: MpvClientLike,
  pathToLoad: string,
): Promise<number | null> {
  const wanted = normalizeSubtitlePathForCompare(pathToLoad);
  // sub-add is queued, so the track may not appear in the first track-list reply.
  for (let attempt = 0; attempt < SYNCED_TRACK_LOOKUP_ATTEMPTS; attempt += 1) {
    let tracks: MpvTrack[] = [];
    try {
      const trackListRaw = await client.requestProperty('track-list');
      tracks = Array.isArray(trackListRaw) ? normalizeTrackIds(trackListRaw as MpvTrack[]) : [];
    } catch {
      return null;
    }
    // Re-adding a file mpv already knows appends a duplicate entry; the newest
    // one holds the retimed content, so prefer the last match.
    const matches = tracks.filter((track) => {
      if (track.type !== 'sub') return false;
      const filename = track['external-filename'];
      return typeof filename === 'string' && normalizeSubtitlePathForCompare(filename) === wanted;
    });
    const added = matches[matches.length - 1];
    if (added && typeof added.id === 'number') {
      return added.id;
    }
    if (attempt < SYNCED_TRACK_LOOKUP_ATTEMPTS - 1) {
      await delay(SYNCED_TRACK_LOOKUP_RETRY_MS);
    }
  }
  return null;
}

async function loadSyncedSubtitle(
  client: MpvClientLike,
  pathToLoad: string,
  slot: SubtitleSlot,
): Promise<void> {
  if (!client.connected) {
    throw new Error('MPV disconnected while loading subtitle');
  }

  if (slot === 'secondary') {
    // Keep the primary track untouched: load without selecting, then point
    // secondary-sid at the freshly added track.
    client.send({ command: ['sub-add', pathToLoad, 'auto'] });
    client.send({ command: ['set_property', 'secondary-sub-delay', 0] });
    const addedTrackId = await findAddedSubtitleTrackId(client, pathToLoad);
    if (addedTrackId === null) {
      throw new Error('Synchronized subtitle did not appear in the mpv track list');
    }
    client.send({ command: ['set_property', 'secondary-sid', addedTrackId] });
    return;
  }

  client.send({ command: ['sub-add', pathToLoad] });
  client.send({ command: ['set_property', 'sub-delay', 0] });
}

async function subsyncToReference(
  engine: 'alass' | 'ffsubsync',
  referenceFilePath: string,
  targetTrack: MpvTrack,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  slot: SubtitleSlot,
  httpHeaders: ResolvedMpvHttpHeaders | null,
): Promise<SubsyncResult> {
  const targetExtraction = await extractSubtitleTrackToFile({
    resolveFfmpegPath: () => resolveSubsyncExecutable(resolved.ffmpegPath, 'ffmpeg'),
    videoPath: context.videoPath,
    track: targetTrack,
    httpHeaders,
  });

  // alass reads neither the target nor the reference as WebVTT, so both are
  // rewritten as SRT first; ffsubsync parses VTT itself. A video reference is
  // left alone: alass reads its audio, and there are no cues to convert.
  const referenceIsVideo = referenceFilePath === context.videoPath;
  let convertedTarget: FileExtractionResult | null = null;
  let convertedReference: FileExtractionResult | null = null;
  if (engine === 'alass') {
    try {
      convertedTarget = convertSubtitleForAlass(targetExtraction.path);
      convertedReference = referenceIsVideo ? null : convertSubtitleForAlass(referenceFilePath);
    } catch (error) {
      if (convertedTarget) cleanupTemporaryFile(convertedTarget);
      cleanupTemporaryFile(targetExtraction);
      const message = `alass synchronization failed: ${(error as Error).message}`;
      logger.error(message);
      return { ok: false, message };
    }
  }

  const target = convertedTarget ?? targetExtraction;
  const referencePath = convertedReference?.path ?? referenceFilePath;
  const replaceTarget = resolved.replace !== false && !target.temporary;
  const outputPath = buildRetimedPath(target.path, replaceTarget);
  logger.info(
    `Running ${engine}: target=${target.path} reference=${referencePath} output=${outputPath}`,
  );

  try {
    let result: CommandResult;
    if (engine === 'alass') {
      const alassPath = resolveSubsyncExecutable(resolved.alassPath, 'alass');
      result = await runAlassSync(alassPath, referencePath, target.path, outputPath);
    } else {
      const ffsubsyncPath = resolveSubsyncExecutable(resolved.ffsubsyncPath, 'ffsubsync');
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        target.path,
        outputPath,
        context.audioStreamIndex,
      );
    }

    if (!result.ok || !fileExists(outputPath)) {
      const details = summarizeCommandFailure(engine, result);
      logger.error(`${engine} synchronization failed: ${details}`);
      return {
        ok: false,
        message: `${engine} synchronization failed: ${details}`,
      };
    }

    await loadSyncedSubtitle(client, outputPath, slot);
    logger.info(`${engine} synchronization succeeded; loaded ${outputPath}`);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    if (convertedReference) cleanupTemporaryFile(convertedReference);
    if (convertedTarget) cleanupTemporaryFile(convertedTarget, outputPath);
    cleanupTemporaryFile(targetExtraction, outputPath);
  }
}

function validateFfsubsyncReference(videoPath: string): void {
  if (isRemoteMediaPath(videoPath)) {
    throw new Error(
      'FFsubsync cannot reliably sync stream URLs because it needs direct reference media access. Use Alass with a secondary subtitle source or play a local file.',
    );
  }
}

function resolveTargetTrack(
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
function resolveTargetSlot(targetTrack: MpvTrack, context: SubsyncContext): SubtitleSlot {
  if (typeof targetTrack.id !== 'number') return 'primary';
  if (targetTrack.id === context.primaryTrack.id) return 'primary';
  if (context.secondaryTrack && targetTrack.id === context.secondaryTrack.id) return 'secondary';
  return 'primary';
}

export async function runSubsyncManual(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  // Only streams need mpv's request context, and asking costs four round-trips.
  const httpHeaders = isRemoteMediaPath(context.videoPath)
    ? await resolveMpvHttpHeaders(client)
    : null;

  logger.info(
    `Manual subsync requested: engine=${request.engine} referenceMode=${
      request.referenceMode ?? 'track'
    } referenceTrackId=${request.referenceTrackId ?? 'none'} targetTrackId=${
      request.targetTrackId ?? 'none'
    } remote=${httpHeaders !== null}`,
  );

  const targetTrack = resolveTargetTrack(request, context);
  if (!targetTrack) {
    return { ok: false, message: 'Select the out-of-sync subtitle track to retime' };
  }
  const targetSlot = resolveTargetSlot(targetTrack, context);

  if (request.engine === 'ffsubsync') {
    try {
      validateFfsubsyncReference(context.videoPath);
    } catch (error) {
      const message = `ffsubsync synchronization failed: ${(error as Error).message}`;
      logger.error(message);
      return { ok: false, message };
    }
    return subsyncToReference(
      'ffsubsync',
      context.videoPath,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
      httpHeaders,
    );
  }

  if (request.referenceMode === 'video') {
    if (isRemoteMediaPath(context.videoPath)) {
      return {
        ok: false,
        message:
          'alass cannot use a stream URL as reference. Pick a reference subtitle track instead.',
      };
    }
    return subsyncToReference(
      'alass',
      context.videoPath,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
      httpHeaders,
    );
  }

  const referenceTrack = getTrackById(context.subtitleTracks, request.referenceTrackId ?? null);
  if (!referenceTrack) {
    logger.warn('Manual subsync rejected: no alass reference track selected');
    return { ok: false, message: 'Select a reference subtitle track for alass' };
  }
  if (referenceTrack.id === targetTrack.id) {
    return { ok: false, message: 'Reference and out-of-sync subtitles must be different tracks' };
  }

  let referenceExtraction: FileExtractionResult | null = null;
  try {
    referenceExtraction = await extractSubtitleTrackToFile({
      resolveFfmpegPath: () => resolveSubsyncExecutable(resolved.ffmpegPath, 'ffmpeg'),
      videoPath: context.videoPath,
      track: referenceTrack,
      httpHeaders,
    });
    return await subsyncToReference(
      'alass',
      referenceExtraction.path,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
      httpHeaders,
    );
  } finally {
    if (referenceExtraction) {
      cleanupTemporaryFile(referenceExtraction);
    }
  }
}

export async function openSubsyncManualPicker(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const subtitleTracks = context.subtitleTracks
    .filter((track) => typeof track.id === 'number')
    .map((track) => ({
      id: track.id as number,
      label: formatTrackLabel(track),
    }));
  const primaryTrackId =
    typeof context.primaryTrack.id === 'number' ? context.primaryTrack.id : null;
  const secondaryTrackId =
    typeof context.secondaryTrack?.id === 'number' ? context.secondaryTrack.id : null;
  const payload: SubsyncManualPayload = {
    subtitleTracks,
    // The secondary track can be filtered or deduped out of the emitted list,
    // so only default to it when the picker actually offers it.
    defaultReferenceTrackId:
      subtitleTracks.find((track) => track.id === secondaryTrackId)?.id ??
      subtitleTracks.find((track) => track.id !== primaryTrackId)?.id ??
      null,
    defaultTargetTrackId: primaryTrackId,
    videoReferenceAvailable: !isRemoteMediaPath(context.videoPath),
    ffsubsyncAvailable: !isRemoteMediaPath(context.videoPath),
  };
  deps.openManualPicker(payload);
}

export async function triggerSubsyncFromConfig(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  if (deps.isSubsyncInProgress()) {
    deps.showMpvOsd('Subsync already running');
    return;
  }

  try {
    await openSubsyncManualPicker(deps);
    deps.showMpvOsd('Subsync: choose engine and subtitles');
  } catch (error) {
    deps.showMpvOsd(`Subsync failed: ${(error as Error).message}`);
  } finally {
    deps.setSubsyncInProgress(false);
  }
}
