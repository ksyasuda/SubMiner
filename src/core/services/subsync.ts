import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { SubsyncManualPayload, SubsyncManualRunRequest, SubsyncResult } from '../../types';
import {
  CommandResult,
  codecToExtension,
  fileExists,
  formatTrackLabel,
  getTrackById,
  hasPathSeparators,
  MpvTrack,
  runCommand,
  SubsyncContext,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import { isRemoteMediaPath } from '../../jimaku/utils';

interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

function summarizeCommandFailure(command: string, result: CommandResult): string {
  const parts = [
    `code=${result.code ?? 'n/a'}`,
    result.stderr ? `stderr: ${result.stderr}` : '',
    result.stdout ? `stdout: ${result.stdout}` : '',
    result.error ? `error: ${result.error}` : '',
  ]
    .map((value) => value.trim())
    .filter(Boolean);

  if (parts.length === 0) return `command failed (${command})`;
  return `command failed (${command}) ${parts.join(' | ')}`;
}

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

function dedupeSourceTracks(tracks: MpvTrack[]): MpvTrack[] {
  const deduped = new Map<string, MpvTrack>();
  for (const track of tracks) {
    const identity = getSourceTrackIdentity(track);
    const existing = deduped.get(identity);
    if (!existing || (track.selected && !existing.selected)) {
      deduped.set(identity, track);
    }
  }
  return [...deduped.values()];
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

  const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
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
  const sourceTracks = subtitleTracks
    .filter((track) => track.id !== sid)
    .filter((track) => {
      if (!track.external) return true;
      const filename = track['external-filename'];
      return typeof filename === 'string' && filename.length > 0;
    });
  const uniqueSourceTracks = dedupeSourceTracks(sourceTracks);

  return {
    videoPath,
    primaryTrack,
    secondaryTrack,
    sourceTracks: uniqueSourceTracks,
    audioStreamIndex: client.currentAudioStreamIndex,
  };
}

function ensureExecutablePath(pathOrName: string, name: string): string {
  if (!pathOrName) {
    throw new Error(`Missing ${name} path in config`);
  }

  if (hasPathSeparators(pathOrName) && !fileExists(pathOrName)) {
    throw new Error(`Configured ${name} executable not found: ${pathOrName}`);
  }
  return pathOrName;
}

async function extractSubtitleTrackToFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvTrack,
): Promise<FileExtractionResult> {
  if (track.external) {
    const externalPath = track['external-filename'];
    if (typeof externalPath !== 'string' || externalPath.length === 0) {
      throw new Error('External subtitle track has no file path');
    }
    if (!fileExists(externalPath)) {
      throw new Error(`Subtitle file not found: ${externalPath}`);
    }
    return { path: externalPath, temporary: false };
  }

  const ffIndex = track['ff-index'];
  const extension = codecToExtension(track.codec);
  if (typeof ffIndex !== 'number' || !Number.isInteger(ffIndex) || ffIndex < 0) {
    throw new Error('Internal subtitle track has no valid ff-index');
  }
  if (!extension) {
    throw new Error(`Unsupported subtitle codec: ${track.codec ?? 'unknown'}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  const extraction = await runCommand(ffmpegPath, [
    '-hide_banner',
    '-nostdin',
    '-y',
    '-loglevel',
    'error',
    '-an',
    '-vn',
    '-i',
    videoPath,
    '-map',
    `0:${ffIndex}`,
    '-f',
    extension,
    outputPath,
  ]);

  if (!extraction.ok || !fileExists(outputPath)) {
    throw new Error(
      `Failed to extract internal subtitle track with ffmpeg: ${summarizeCommandFailure(
        'ffmpeg',
        extraction,
      )}`,
    );
  }

  return { path: outputPath, temporary: true };
}

function cleanupTemporaryFile(extraction: FileExtractionResult): void {
  if (!extraction.temporary) return;
  try {
    if (fileExists(extraction.path)) {
      fs.unlinkSync(extraction.path);
    }
  } catch {}
  try {
    const dir = path.dirname(extraction.path);
    if (fs.existsSync(dir)) {
      fs.rmdirSync(dir);
    }
  } catch {}
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

function loadSyncedSubtitle(client: MpvClientLike, pathToLoad: string): void {
  if (!client.connected) {
    throw new Error('MPV disconnected while loading subtitle');
  }
  client.send({ command: ['sub_add', pathToLoad] });
  client.send({ command: ['set_property', 'sub-delay', 0] });
}

async function subsyncToReference(
  engine: 'alass' | 'ffsubsync',
  referenceFilePath: string,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
): Promise<SubsyncResult> {
  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, 'ffmpeg');
  const primaryExtraction = await extractSubtitleTrackToFile(
    ffmpegPath,
    context.videoPath,
    context.primaryTrack,
  );
  const replacePrimary = resolved.replace !== false && !primaryExtraction.temporary;
  const outputPath = buildRetimedPath(primaryExtraction.path, replacePrimary);

  try {
    let result: CommandResult;
    if (engine === 'alass') {
      const alassPath = ensureExecutablePath(resolved.alassPath, 'alass');
      result = await runAlassSync(alassPath, referenceFilePath, primaryExtraction.path, outputPath);
    } else {
      const ffsubsyncPath = ensureExecutablePath(resolved.ffsubsyncPath, 'ffsubsync');
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        primaryExtraction.path,
        outputPath,
        context.audioStreamIndex,
      );
    }

    if (!result.ok || !fileExists(outputPath)) {
      const details = summarizeCommandFailure(engine, result);
      return {
        ok: false,
        message: `${engine} synchronization failed: ${details}`,
      };
    }

    loadSyncedSubtitle(client, outputPath);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    cleanupTemporaryFile(primaryExtraction);
  }
}

function validateFfsubsyncReference(videoPath: string): void {
  if (isRemoteMediaPath(videoPath)) {
    throw new Error(
      'FFsubsync cannot reliably sync stream URLs because it needs direct reference media access. Use Alass with a secondary subtitle source or play a local file.',
    );
  }
}

export async function runSubsyncManual(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();

  if (request.engine === 'ffsubsync') {
    try {
      validateFfsubsyncReference(context.videoPath);
    } catch (error) {
      return {
        ok: false,
        message: `ffsubsync synchronization failed: ${(error as Error).message}`,
      };
    }
    return subsyncToReference('ffsubsync', context.videoPath, context, resolved, client);
  }

  const sourceTrack = getTrackById(context.sourceTracks, request.sourceTrackId ?? null);
  if (!sourceTrack) {
    return { ok: false, message: 'Select a subtitle source track for alass' };
  }

  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, 'ffmpeg');
  let sourceExtraction: FileExtractionResult | null = null;
  try {
    sourceExtraction = await extractSubtitleTrackToFile(ffmpegPath, context.videoPath, sourceTrack);
    return await subsyncToReference('alass', sourceExtraction.path, context, resolved, client);
  } finally {
    if (sourceExtraction) {
      cleanupTemporaryFile(sourceExtraction);
    }
  }
}

export async function openSubsyncManualPicker(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const payload: SubsyncManualPayload = {
    ffsubsyncAvailable: !isRemoteMediaPath(context.videoPath),
    sourceTracks: context.sourceTracks
      .filter((track) => typeof track.id === 'number')
      .map((track) => ({
        id: track.id as number,
        label: formatTrackLabel(track),
      })),
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
    deps.showMpvOsd('Subsync: choose engine and source');
  } catch (error) {
    deps.showMpvOsd(`Subsync failed: ${(error as Error).message}`);
  } finally {
    deps.setSubsyncInProgress(false);
  }
}
