import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { fileURLToPath } from 'url';
import {
  SubsyncEngine,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncReferenceCandidate,
  SubsyncResult,
} from '../../types';
import {
  buildJellyfinEmbeddedSubtitleUrl,
  CommandResult,
  codecToExtension,
  extensionToFfmpegMuxer,
  fileExists,
  findSubsyncExecutable,
  formatTrackLabel,
  getTrackById,
  hasPathSeparators,
  MpvTrack,
  runCommand,
  SubsyncContext,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import {
  alignSubtitleFileToReferenceSpans,
  evaluateReferenceAlignmentToAudio,
  evaluateReferenceAlignmentToSampledAudio,
  type AlignmentResult,
  type SampledAudioWindow,
} from '../../subsync/alass-align';
import {
  extractAudioSpeechSpans,
  extractSampledAudioWindows,
  extractSinglePassRipMedia,
  getCachedAudioSpeechSpans,
  type AudioSpeechResult,
  type EmbeddedTrackSpec,
} from '../../subsync/audio-speech';
import {
  buildTrackReferenceLabel,
  discoverSidecarReferenceFiles,
  formatReferenceCandidateDetail,
  isEnglishReferenceMeta,
  isFullDialogueHintMeta,
  rankReferenceCandidates,
  type EvaluatedSubtitleReference,
} from '../../subsync/reference-discovery';
import {
  applyTimingToSubtitleFile,
  decodeSubtitleBytesWithStatus,
  extractSpansFromParsedSubtitle,
  formatSpansAsSrt,
  isSignsOrSongsTrackMeta,
  parseSubtitleFileContent,
  type ParsedSubtitleFile,
  type SubtitleTimeSpan,
} from '../../subsync/subtitle-lines';
import { isRemoteMediaPath } from '../../jimaku/utils';

interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

type SubtitleSlot = 'primary' | 'secondary';

const SYNCED_TRACK_LOOKUP_ATTEMPTS = 5;
const SYNCED_TRACK_LOOKUP_RETRY_MS = 100;
const ALREADY_IN_SYNC_MS = 80;
const MIN_REFERENCE_DIALOGUE_LINES = 15;
const AUTO_SUBTITLE_CONFIDENCE_THRESHOLD = 0.65;
const AUTO_REMOTE_CONFIDENCE_THRESHOLD = 0.7;
const AUTO_AUDIO_CONFIDENCE_THRESHOLD = 0.65;
const AUDIO_SPEECH_LEAD_SEC = 0.1;
const JELLYFIN_SUBTITLE_FETCH_TIMEOUT_MS = 15_000;

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

export interface SubsyncVadPaths {
  vadExecutablePath: string | null;
  vadModelPath: string | null;
}

interface SubsyncCoreDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedConfig: () => SubsyncResolvedConfig;
  getVadPaths?: () => SubsyncVadPaths;
  canPickReferenceFile?: () => boolean;
  fetchImpl?: typeof fetch;
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

function parseOptionalNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    if (Number.isFinite(parsed)) return parsed;
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

function resolveLocalSubtitlePath(
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

function normalizeMediaVideoPath(rawPath: string, workingDirectory?: string | null): string {
  return resolveLocalSubtitlePath(rawPath, workingDirectory) ?? rawPath;
}

async function gatherSubsyncContext(client: MpvClientLike): Promise<SubsyncContext> {
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

  // Resolve active audio stream ff-index if not already cached on mpv client.
  let audioStreamIndex = client.currentAudioStreamIndex;
  if (audioStreamIndex === null) {
    const selectedAudio = tracks.find(
      (track) =>
        track.type === 'audio' && track.selected === true && typeof track['ff-index'] === 'number',
    );
    if (selectedAudio && typeof selectedAudio['ff-index'] === 'number') {
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

function ensureExecutablePath(pathOrName: string, name: string): string {
  if (!pathOrName) {
    throw new Error(`Missing ${name} path in config`);
  }

  if (hasPathSeparators(pathOrName) && !fileExists(pathOrName)) {
    throw new Error(`Configured ${name} executable not found: ${pathOrName}`);
  }
  return pathOrName;
}

function isExecutableAvailable(configuredPath: string, names: readonly string[]): boolean {
  if (!configuredPath) return findSubsyncExecutable(undefined, names) !== null;
  if (hasPathSeparators(configuredPath)) return fileExists(configuredPath);
  return findSubsyncExecutable(configuredPath, names) !== null;
}

async function tryFetchJellyfinSubtitleBytes(
  videoPath: string,
  ffIndex: number,
  extension: string,
  fetchImpl: typeof fetch = globalThis.fetch,
): Promise<Uint8Array | null> {
  const subtitleUrl = buildJellyfinEmbeddedSubtitleUrl(videoPath, ffIndex, extension);
  if (!subtitleUrl || typeof fetchImpl !== 'function') return null;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), JELLYFIN_SUBTITLE_FETCH_TIMEOUT_MS);
  try {
    const response = await fetchImpl(subtitleUrl, { signal: controller.signal });
    if (!response.ok) return null;
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength === 0) return null;
    return new Uint8Array(buffer);
  } catch {
    return null;
  } finally {
    clearTimeout(timeout);
  }
}

async function extractSubtitleTrackToFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvTrack,
  workingDirectory?: string | null,
  fetchImpl?: typeof fetch,
): Promise<FileExtractionResult> {
  if (track.external) {
    const externalPath = track['external-filename'];
    if (typeof externalPath !== 'string' || externalPath.length === 0) {
      throw new Error('External subtitle track has no file path');
    }
    const localPath = resolveLocalSubtitlePath(externalPath, workingDirectory) ?? externalPath;
    if (!fileExists(localPath)) {
      throw new Error(`Subtitle file not found: ${localPath}`);
    }
    return { path: localPath, temporary: false };
  }

  const ffIndex = track['ff-index'];
  const extension = codecToExtension(track.codec);
  if (typeof ffIndex !== 'number' || !Number.isInteger(ffIndex) || ffIndex < 0) {
    throw new Error('Internal subtitle track has no valid ff-index');
  }
  if (!extension) {
    throw new Error(`Unsupported subtitle codec: ${track.codec ?? 'unknown'}`);
  }

  if (isRemoteMediaPath(videoPath)) {
    const fetchedBytes = await tryFetchJellyfinSubtitleBytes(
      videoPath,
      ffIndex,
      extension,
      fetchImpl,
    );
    if (fetchedBytes) {
      const decoded = decodeSubtitleBytesWithStatus(fetchedBytes);
      const parsed = parseSubtitleFileContent(decoded.text, `track_${ffIndex}.${extension}`);
      if (parsed.events.length > 0) {
        const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
        const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
        try {
          fs.writeFileSync(outputPath, fetchedBytes);
          return { path: outputPath, temporary: true };
        } catch (error) {
          fs.rmSync(tempDir, { recursive: true, force: true });
          throw error;
        }
      }
    }
  }

  const resolvedFfmpegPath = ensureExecutablePath(ffmpegPath, 'ffmpeg');
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  try {
    const extraction = await runCommand(resolvedFfmpegPath, [
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
      extensionToFfmpegMuxer(extension),
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
  } catch (error) {
    fs.rmSync(tempDir, { recursive: true, force: true });
    throw error;
  }
}

function cleanupTemporaryFile(extraction: FileExtractionResult | null): void {
  if (!extraction || !extraction.temporary) return;
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

const LOSSY_REWRITE_ERROR_MESSAGE =
  'Cannot safely rewrite subtitle file because its text encoding could not be decoded unambiguously and losslessly. Convert the file to UTF-8 and retry.';

function readAndParseSubtitleFile(filePath: string): ParsedSubtitleFile {
  const bytes = fs.readFileSync(filePath);
  const decoded = decodeSubtitleBytesWithStatus(bytes);
  return {
    ...parseSubtitleFileContent(decoded.text, filePath),
    lossless: decoded.lossless,
  };
}

function writeRetimedContent(outputPath: string, content: string): void {
  const dir = path.dirname(outputPath);
  const tmpPath = path.join(
    dir,
    `.subminer-subsync-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`,
  );
  try {
    fs.writeFileSync(tmpPath, content, 'utf8');
    fs.renameSync(tmpPath, outputPath);
  } catch (error) {
    try {
      if (fs.existsSync(tmpPath)) {
        fs.unlinkSync(tmpPath);
      }
    } catch {
      // Best-effort cleanup of temp write artifact.
    }
    throw error;
  }
}

async function isPlaybackSessionUnchanged(
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
    const needsTrackIdentityCheck =
      typeof targetTrack.id === 'number' && targetTrack.id !== initialSlotTrackId;
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

// Windows local paths are case-insensitive and mpv may echo them with different
// separators; on POSIX, backslashes are valid filename characters and must stay distinct.
function normalizeSubtitlePathForCompare(value: string): string {
  if (process.platform === 'win32' && !isRemoteMediaPath(value)) {
    return value.replace(/\\/g, '/').toLowerCase();
  }
  return value;
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

function formatOffsetSummary(offsetMs: number): string {
  const sec = offsetMs / 1000;
  const formatted = Math.abs(sec) >= 10 ? sec.toFixed(1) : sec.toFixed(2);
  return offsetMs >= 0 ? `+${formatted}s` : `${formatted}s`;
}

function formatAlignmentSummary(sourceLabel: string, alignment: AlignmentResult): string {
  const parts = [`Synced via ${sourceLabel}`, formatOffsetSummary(alignment.offsetMs)];
  if (Math.abs(alignment.ratio - 1.0) > 1e-4) {
    parts.push(`speed ${alignment.ratio.toFixed(4)}x`);
  }
  if (alignment.segments > 1) {
    parts.push(`${alignment.segments} segments`);
  }
  const pct = Math.round(alignment.confidence * 100);
  if (pct > 0) {
    parts.push(`${pct}%`);
  }
  return parts.join(' · ');
}

function isAlreadyInSync(alignment: AlignmentResult): boolean {
  return (
    alignment.confident &&
    !alignment.cut &&
    alignment.segments === 1 &&
    Math.abs(alignment.ratio - 1.0) < 1e-4 &&
    Math.abs(alignment.offsetMs) <= ALREADY_IN_SYNC_MS
  );
}

async function resolveAudioSpeechForContext(
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
): Promise<AudioSpeechResult | null> {
  if (isRemoteMediaPath(context.videoPath) || !fileExists(context.videoPath)) {
    return null;
  }
  if (!isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg'])) {
    return null;
  }
  const vadPaths = deps.getVadPaths?.();
  try {
    const speech = await extractAudioSpeechSpans({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
      vadModelPath: vadPaths?.vadModelPath ?? null,
    });
    return speech.spans.length > 0 ? speech : null;
  } catch {
    return null;
  }
}

interface CandidateWithSpans {
  evaluated: EvaluatedSubtitleReference;
  dialogueSpans: SubtitleTimeSpan[];
  allSpans: SubtitleTimeSpan[];
}

function collectExtractableEmbeddedRefTracks(
  context: SubsyncContext,
  targetTrack: MpvTrack,
): EmbeddedTrackSpec[] {
  const specs: EmbeddedTrackSpec[] = [];
  for (const track of context.subtitleTracks) {
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    if (track.external) continue;
    // Skip embedded tracks already tagged as signs/songs or forced so we don't waste demux I/O.
    if (isSignsOrSongsTrackMeta(track)) continue;
    const ffIndex = track['ff-index'];
    const extension = codecToExtension(track.codec);
    if (typeof ffIndex === 'number' && Number.isInteger(ffIndex) && ffIndex >= 0 && extension) {
      specs.push({ ffIndex, extension });
    }
  }
  return specs;
}

async function discoverAndEvaluateSubtitleCandidates(
  context: SubsyncContext,
  targetTrack: MpvTrack,
  resolved: SubsyncResolvedConfig,
  audioSpeech: AudioSpeechResult | null,
  options: {
    sampledAudioWindows?: SampledAudioWindow[] | null;
    preExtractedEmbedded?: ReadonlyMap<number, ParsedSubtitleFile> | null;
    skipSignsEmbeddedExtraction?: boolean;
    skipEmbeddedExtraction?: boolean;
    fetchImpl?: typeof fetch;
  } = {},
): Promise<CandidateWithSpans[]> {
  const results: CandidateWithSpans[] = [];
  const loadedPaths: string[] = [];
  const ffmpegAvailable = isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  const sampledWindows = options.sampledAudioWindows ?? null;
  const preExtracted = options.preExtractedEmbedded ?? null;
  const skipSignsEmbedded = options.skipSignsEmbeddedExtraction !== false;
  const skipEmbeddedExtraction = options.skipEmbeddedExtraction === true;

  for (const track of context.subtitleTracks) {
    if (track.external && track['external-filename']) {
      const resolvedPath = resolveLocalSubtitlePath(
        track['external-filename'],
        context.workingDirectory,
      );
      if (resolvedPath) loadedPaths.push(resolvedPath);
    }
  }

  const evaluateParsedCandidate = (
    meta: Omit<
      EvaluatedSubtitleReference,
      'dialogueCount' | 'signCount' | 'audioStatus' | 'audioOffsetMs' | 'audioConfidence'
    >,
    parsed: ParsedSubtitleFile | null,
  ): CandidateWithSpans => {
    const stripped = parsed
      ? extractSpansFromParsedSubtitle(parsed, { stripSigns: true })
      : { spans: [], dialogueCount: 0, signCount: 0 };
    const all = parsed
      ? extractSpansFromParsedSubtitle(parsed, { stripSigns: false })
      : { spans: [] };

    const effectiveSignsOnly =
      meta.isSignsOnly ||
      (parsed !== null && stripped.dialogueCount < MIN_REFERENCE_DIALOGUE_LINES);

    let audioStatus: EvaluatedSubtitleReference['audioStatus'] = 'unverified';
    let audioOffsetMs: number | null = null;
    let audioConfidence = 0;

    if (!effectiveSignsOnly && stripped.spans.length >= MIN_REFERENCE_DIALOGUE_LINES) {
      if (audioSpeech && audioSpeech.spans.length > 0) {
        const audioEval = evaluateReferenceAlignmentToAudio(stripped.spans, audioSpeech.spans, {
          lengthSec: context.durationSec ?? audioSpeech.durationSec,
          audioCalibration: audioSpeech.calibration,
          audioLeadSec: AUDIO_SPEECH_LEAD_SEC,
        });
        audioStatus = audioEval.status;
        audioOffsetMs = audioEval.offsetMs;
        audioConfidence = audioEval.confidence;
      } else if (sampledWindows && sampledWindows.length >= 2) {
        const sampleEval = evaluateReferenceAlignmentToSampledAudio(
          stripped.spans,
          sampledWindows,
          { audioLeadSec: AUDIO_SPEECH_LEAD_SEC },
        );
        audioStatus = sampleEval.status;
        audioOffsetMs = sampleEval.offsetMs;
        audioConfidence = sampleEval.confidence;
      }
    }

    return {
      evaluated: {
        ...meta,
        isSignsOnly: effectiveSignsOnly,
        dialogueCount: stripped.dialogueCount,
        signCount: stripped.signCount,
        audioStatus,
        audioOffsetMs,
        audioConfidence,
      },
      dialogueSpans: stripped.spans,
      allSpans: all.spans,
    };
  };

  // 1. Evaluate all mpv subtitle tracks other than target.
  for (const track of context.subtitleTracks) {
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    const isSecondary = context.secondaryTrack !== null && track.id === context.secondaryTrack.id;
    const isEmbedded = !track.external;
    const isSignsTrack = isSignsOrSongsTrackMeta(track);
    const externalFilename = track['external-filename'];
    const baseMeta = {
      key: `track:${track.id}`,
      kind: 'track' as const,
      trackId: track.id,
      filePath: track.external
        ? (resolveLocalSubtitlePath(externalFilename ?? '', context.workingDirectory) ?? null)
        : null,
      label: buildTrackReferenceLabel(track, isSecondary),
      isSecondary,
      isEmbedded,
      isEnglish: isEnglishReferenceMeta({
        lang: track.lang,
        title: track.title,
        filename: externalFilename,
      }),
      isFullDialogueHint: isFullDialogueHintMeta({
        title: track.title,
        filename: externalFilename,
      }),
      isSignsOnly: isSignsTrack,
    };

    if (isEmbedded && skipSignsEmbedded && isSignsTrack) {
      results.push(evaluateParsedCandidate(baseMeta, null));
      continue;
    }

    if (isEmbedded && preExtracted && typeof track['ff-index'] === 'number') {
      const preParsed = preExtracted.get(track['ff-index']);
      if (preParsed) {
        results.push(evaluateParsedCandidate(baseMeta, preParsed));
        continue;
      }
    }

    if (isEmbedded && skipEmbeddedExtraction) {
      results.push(evaluateParsedCandidate(baseMeta, null));
      continue;
    }

    let extraction: FileExtractionResult | null = null;
    try {
      if (isEmbedded && (!ffmpegAvailable || !codecToExtension(track.codec))) {
        results.push(evaluateParsedCandidate(baseMeta, null));
        continue;
      }
      extraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        track,
        context.workingDirectory,
        options.fetchImpl,
      );
      const parsed = readAndParseSubtitleFile(extraction.path);
      results.push(evaluateParsedCandidate(baseMeta, parsed));
    } catch {
      results.push(evaluateParsedCandidate(baseMeta, null));
    } finally {
      cleanupTemporaryFile(extraction);
    }
  }

  // 2. Discover unloaded sidecar subtitle files beside the local video.
  if (!isRemoteMediaPath(context.videoPath) && fileExists(context.videoPath)) {
    const sidecars = discoverSidecarReferenceFiles(context.videoPath, loadedPaths);
    for (const sidecar of sidecars) {
      const baseMeta = {
        key: `file:${sidecar.path}`,
        kind: 'sidecar' as const,
        trackId: null,
        filePath: sidecar.path,
        label: `Sidecar - ${sidecar.filename}`,
        isSecondary: false,
        isEmbedded: false,
        isEnglish: isEnglishReferenceMeta({
          lang: sidecar.languageHint,
          filename: sidecar.filename,
        }),
        isFullDialogueHint: isFullDialogueHintMeta({ filename: sidecar.filename }),
        isSignsOnly: sidecar.isSignsOnly,
      };
      try {
        const parsed = readAndParseSubtitleFile(sidecar.path);
        results.push(evaluateParsedCandidate(baseMeta, parsed));
      } catch {
        // Ignore unreadable sidecars.
      }
    }
  }

  const rankedEvaluated = rankReferenceCandidates(results.map((r) => r.evaluated));
  const byKey = new Map(results.map((r) => [r.evaluated.key, r]));
  return rankedEvaluated.map((ev) => byKey.get(ev.key)!);
}

function buildManualPayload(
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
  options: {
    evaluatedCandidates?: CandidateWithSpans[];
    audioSpeechAvailable?: boolean;
    guidanceMessage?: string | null;
    defaultEngine?: SubsyncEngine;
    defaultReferenceKey?: string | null;
  } = {},
): SubsyncManualPayload {
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

  const isLocalMedia = !isRemoteMediaPath(context.videoPath);
  const ffmpegAvailable = isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  const alassAvailable = isExecutableAvailable(resolved.alassPath, ['alass', 'alass-cli']);
  const ffsubsyncAvailable =
    isLocalMedia && isExecutableAvailable(resolved.ffsubsyncPath, ['ffsubsync']);
  const audioReferenceAvailable = options.audioSpeechAvailable ?? (isLocalMedia && ffmpegAvailable);

  const referenceCandidates: SubsyncReferenceCandidate[] = [];
  if (options.evaluatedCandidates && options.evaluatedCandidates.length > 0) {
    for (const [idx, item] of options.evaluatedCandidates.entries()) {
      const ev = item.evaluated;
      referenceCandidates.push({
        key: ev.key,
        kind: ev.kind === 'sidecar' ? 'file' : 'track',
        label: ev.label,
        detail: formatReferenceCandidateDetail(ev) || undefined,
        trackId: ev.trackId,
        filePath: ev.filePath,
        audioStatus: ev.audioStatus,
        audioOffsetMs: ev.audioOffsetMs,
        isSignsOnly: ev.isSignsOnly,
        recommended: idx === 0 && ev.audioStatus === 'aligned' && !ev.isSignsOnly,
      });
    }
  } else {
    for (const track of context.subtitleTracks) {
      if (typeof track.id !== 'number' || track.id === primaryTrackId) continue;
      const isSecondary = secondaryTrackId !== null && track.id === secondaryTrackId;
      referenceCandidates.push({
        key: `track:${track.id}`,
        kind: 'track',
        label: buildTrackReferenceLabel(track, isSecondary),
        trackId: track.id,
        filePath: null,
        isSignsOnly: isSignsOrSongsTrackMeta(track),
      });
    }
  }

  if (isLocalMedia) {
    const audioRecommended =
      !referenceCandidates.some((c) => c.recommended) && audioReferenceAvailable;
    referenceCandidates.push({
      key: 'video',
      kind: 'video',
      label: 'Video audio (built-in speech detection)',
      detail: audioReferenceAvailable ? 'no reference subtitle needed' : 'requires ffmpeg',
      recommended: audioRecommended,
    });
  }

  const defaultReferenceTrackId =
    subtitleTracks.find((track) => track.id === secondaryTrackId)?.id ??
    subtitleTracks.find((track) => track.id !== primaryTrackId)?.id ??
    null;

  const recommendedKey =
    options.defaultReferenceKey ??
    referenceCandidates.find((c) => c.recommended)?.key ??
    (defaultReferenceTrackId !== null ? `track:${defaultReferenceTrackId}` : null) ??
    referenceCandidates[0]?.key ??
    null;

  return {
    subtitleTracks,
    defaultReferenceTrackId,
    defaultTargetTrackId: primaryTrackId,
    videoReferenceAvailable: isLocalMedia,
    ffsubsyncAvailable,
    mediaPath: context.videoPath,
    guidanceMessage: options.guidanceMessage ?? null,
    referenceCandidates,
    defaultReferenceKey: recommendedKey,
    defaultEngine: options.defaultEngine ?? 'builtin',
    builtinAvailable: true,
    alassAvailable,
    audioReferenceAvailable,
    canPickReferenceFile: deps.canPickReferenceFile?.() ?? false,
  };
}

async function runBuiltinSubsync(
  request: SubsyncManualRunRequest,
  targetTrack: MpvTrack,
  targetSlot: SubtitleSlot,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const stripSigns = request.stripSigns !== false;
  const allowSplits = request.allowSplits !== false;

  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    deps.fetchImpl,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);

  let refExtraction: FileExtractionResult | null = null;
  let keepTempTarget = false;
  try {
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    if (parsedTarget.events.length === 0) {
      return { ok: false, message: 'Target subtitle file has no readable dialogue cues.' };
    }
    if (parsedTarget.lossless === false) {
      return { ok: false, message: LOSSY_REWRITE_ERROR_MESSAGE };
    }

    let referenceSpans: SubtitleTimeSpan[] = [];
    let referenceLabel = 'reference';
    let calibration: 'clean' | 'noisy' = 'clean';
    let leadSec = 0;
    let lengthSec = context.durationSec ?? 0;

    if (request.referenceMode === 'video') {
      if (isRemoteMediaPath(context.videoPath)) {
        return {
          ok: false,
          message:
            'Built-in audio sync cannot read remote stream URLs directly. Pick a reference subtitle track or browse a local subtitle file.',
        };
      }
      const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, 'ffmpeg');
      const vadPaths = deps.getVadPaths?.();
      const audioSpeech = await extractAudioSpeechSpans({
        ffmpegPath,
        videoPath: context.videoPath,
        audioStreamIndex: context.audioStreamIndex,
        vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
        vadModelPath: vadPaths?.vadModelPath ?? null,
      });
      if (audioSpeech.spans.length === 0) {
        return {
          ok: false,
          message: 'No spoken dialogue was detected in the video audio.',
        };
      }
      referenceSpans = audioSpeech.spans;
      referenceLabel = audioSpeech.source === 'silero-vad' ? 'audio (Silero VAD)' : 'video audio';
      calibration = audioSpeech.calibration;
      leadSec = AUDIO_SPEECH_LEAD_SEC;
      lengthSec = Math.max(lengthSec, audioSpeech.durationSec);
    } else if (request.referenceMode === 'file') {
      const refPath = request.referenceFilePath?.trim() ?? '';
      if (!refPath || !fileExists(refPath)) {
        return { ok: false, message: 'Select an existing reference subtitle file.' };
      }
      const parsedRef = readAndParseSubtitleFile(refPath);
      referenceSpans = extractSpansFromParsedSubtitle(parsedRef, { stripSigns }).spans;
      referenceLabel = path.basename(refPath);
    } else {
      const referenceTrack = getTrackById(context.subtitleTracks, request.referenceTrackId ?? null);
      if (!referenceTrack) {
        return { ok: false, message: 'Select a reference subtitle track.' };
      }
      if (referenceTrack.id === targetTrack.id) {
        return {
          ok: false,
          message: 'Reference and out-of-sync subtitles must be different tracks',
        };
      }
      refExtraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        referenceTrack,
        context.workingDirectory,
        deps.fetchImpl,
      );
      const parsedRef = readAndParseSubtitleFile(refExtraction.path);
      referenceSpans = extractSpansFromParsedSubtitle(parsedRef, { stripSigns }).spans;
      referenceLabel = buildTrackReferenceLabel(
        referenceTrack,
        context.secondaryTrack?.id === referenceTrack.id,
      );
    }

    if (referenceSpans.length === 0) {
      return { ok: false, message: 'Selected reference has no usable dialogue timestamps.' };
    }

    const alignment = alignSubtitleFileToReferenceSpans(parsedTarget, referenceSpans, {
      stripSigns,
      allowSplits,
      calibration,
      leadSec,
      lengthSec,
    });

    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
      return { ok: false, message: 'Playback changed while syncing; discarded stale result.' };
    }

    writeRetimedContent(outputPath, alignment.rewrittenContent);
    if (targetExtraction.temporary) {
      keepTempTarget = true;
    }
    await loadSyncedSubtitle(client, outputPath, targetSlot);
    return {
      ok: true,
      message: formatAlignmentSummary(referenceLabel, alignment),
    };
  } finally {
    cleanupTemporaryFile(refExtraction);
    if (!keepTempTarget) {
      cleanupTemporaryFile(targetExtraction);
    }
  }
}

async function runManualOffsetSubsync(
  offsetMs: number,
  targetTrack: MpvTrack,
  targetSlot: SubtitleSlot,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  fetchImpl?: typeof fetch,
): Promise<SubsyncResult> {
  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    fetchImpl,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);

  let keepTempTarget = false;
  try {
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    if (parsedTarget.events.length === 0) {
      return { ok: false, message: 'Target subtitle file has no readable cues to shift.' };
    }
    if (parsedTarget.lossless === false) {
      return { ok: false, message: LOSSY_REWRITE_ERROR_MESSAGE };
    }
    const allIndices = parsedTarget.events.map((_, idx) => idx);
    const moves = new Array<number>(allIndices.length).fill(offsetMs / 1000);
    const rewritten = applyTimingToSubtitleFile(parsedTarget, allIndices, moves, 1.0);
    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
      return { ok: false, message: 'Playback changed while syncing; discarded stale result.' };
    }
    writeRetimedContent(outputPath, rewritten);
    if (targetExtraction.temporary) {
      keepTempTarget = true;
    }
    await loadSyncedSubtitle(client, outputPath, targetSlot);
    return {
      ok: true,
      message: `Shifted subtitle by ${formatOffsetSummary(offsetMs)}`,
    };
  } finally {
    if (!keepTempTarget) {
      cleanupTemporaryFile(targetExtraction);
    }
  }
}

async function subsyncToReference(
  engine: 'alass' | 'ffsubsync',
  referenceFilePath: string,
  targetTrack: MpvTrack,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  slot: SubtitleSlot,
  fetchImpl?: typeof fetch,
): Promise<SubsyncResult> {
  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    fetchImpl,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);
  const parsedOutput = path.parse(outputPath);
  const stagedOutputPath = path.join(
    parsedOutput.dir,
    `.subminer-subsync-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}${parsedOutput.ext || '.srt'}`,
  );

  try {
    let result: CommandResult;
    if (engine === 'alass') {
      const alassPath = ensureExecutablePath(resolved.alassPath, 'alass');
      result = await runAlassSync(
        alassPath,
        referenceFilePath,
        targetExtraction.path,
        stagedOutputPath,
      );
    } else {
      const ffsubsyncPath = ensureExecutablePath(resolved.ffsubsyncPath, 'ffsubsync');
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        targetExtraction.path,
        stagedOutputPath,
        context.audioStreamIndex,
      );
    }

    if (!result.ok || !fileExists(stagedOutputPath)) {
      const details = summarizeCommandFailure(engine, result);
      return {
        ok: false,
        message: `${engine} synchronization failed: ${details}`,
      };
    }

    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, slot))) {
      return { ok: false, message: 'Playback changed while syncing; discarded stale result.' };
    }

    fs.renameSync(stagedOutputPath, outputPath);
    await loadSyncedSubtitle(client, outputPath, slot);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    try {
      if (fs.existsSync(stagedOutputPath)) {
        fs.unlinkSync(stagedOutputPath);
      }
    } catch {
      // Best-effort cleanup of staged external-engine output.
    }
    cleanupTemporaryFile(targetExtraction);
  }
}

function validateFfsubsyncReference(videoPath: string): void {
  if (isRemoteMediaPath(videoPath)) {
    throw new Error(
      'FFsubsync cannot reliably sync stream URLs because it needs direct reference media access. Use built-in or alass sync with a subtitle track, or play a local file.',
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

function maybeWriteCleanedReferenceSrt(
  sourcePath: string,
  stripSigns: boolean,
): FileExtractionResult | null {
  if (!stripSigns) return null;
  try {
    const parsed = readAndParseSubtitleFile(sourcePath);
    const extracted = extractSpansFromParsedSubtitle(parsed, { stripSigns: true });
    if (extracted.signCount === 0 || extracted.dialogueCount < MIN_REFERENCE_DIALOGUE_LINES) {
      return null;
    }
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-cleanref-'));
    const cleanPath = path.join(tempDir, 'clean_ref.srt');
    fs.writeFileSync(cleanPath, formatSpansAsSrt(extracted.spans), 'utf8');
    return { path: cleanPath, temporary: true };
  } catch {
    return null;
  }
}

export async function runSubsyncManual(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();

  if (request.mediaPath !== undefined && request.mediaPath !== null) {
    const expectedVideoPath =
      request.mediaPath.trim().length > 0
        ? normalizeMediaVideoPath(request.mediaPath, context.workingDirectory)
        : '';
    if (
      !expectedVideoPath ||
      normalizeSubtitlePathForCompare(expectedVideoPath) !==
        normalizeSubtitlePathForCompare(context.videoPath)
    ) {
      return {
        ok: false,
        message:
          'Active video changed since the subsync picker opened. Reopen subsync for the current video.',
      };
    }
  }

  const targetTrack = resolveTargetTrack(request, context);
  if (!targetTrack) {
    return { ok: false, message: 'Select the out-of-sync subtitle track to retime' };
  }
  const targetSlot = resolveTargetSlot(targetTrack, context);

  if (request.engine === 'offset') {
    const offsetMs = request.offsetMs ?? 0;
    if (!Number.isFinite(offsetMs) || offsetMs === 0) {
      return { ok: false, message: 'Enter a non-zero offset in milliseconds' };
    }
    return runManualOffsetSubsync(
      offsetMs,
      targetTrack,
      targetSlot,
      context,
      resolved,
      client,
      deps.fetchImpl,
    );
  }

  if (request.engine === 'builtin') {
    return runBuiltinSubsync(request, targetTrack, targetSlot, context, resolved, client, deps);
  }

  if (request.engine === 'ffsubsync') {
    try {
      validateFfsubsyncReference(context.videoPath);
    } catch (error) {
      return {
        ok: false,
        message: `ffsubsync synchronization failed: ${(error as Error).message}`,
      };
    }
    return subsyncToReference(
      'ffsubsync',
      context.videoPath,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
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
    );
  }

  if (request.referenceMode === 'file') {
    const refPath = request.referenceFilePath?.trim() ?? '';
    if (!refPath || !fileExists(refPath)) {
      return { ok: false, message: 'Select an existing reference subtitle file for alass' };
    }
    const cleaned = maybeWriteCleanedReferenceSrt(refPath, request.stripSigns !== false);
    try {
      return await subsyncToReference(
        'alass',
        cleaned ? cleaned.path : refPath,
        targetTrack,
        context,
        resolved,
        client,
        targetSlot,
        deps.fetchImpl,
      );
    } finally {
      cleanupTemporaryFile(cleaned);
    }
  }

  const referenceTrack = getTrackById(context.subtitleTracks, request.referenceTrackId ?? null);
  if (!referenceTrack) {
    return { ok: false, message: 'Select a reference subtitle track for alass' };
  }
  if (referenceTrack.id === targetTrack.id) {
    return { ok: false, message: 'Reference and out-of-sync subtitles must be different tracks' };
  }

  let referenceExtraction: FileExtractionResult | null = null;
  let cleanedRef: FileExtractionResult | null = null;
  try {
    referenceExtraction = await extractSubtitleTrackToFile(
      resolved.ffmpegPath,
      context.videoPath,
      referenceTrack,
      context.workingDirectory,
      deps.fetchImpl,
    );
    cleanedRef = maybeWriteCleanedReferenceSrt(
      referenceExtraction.path,
      request.stripSigns === true,
    );
    return await subsyncToReference(
      'alass',
      cleanedRef ? cleanedRef.path : referenceExtraction.path,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
      deps.fetchImpl,
    );
  } finally {
    cleanupTemporaryFile(cleanedRef);
    cleanupTemporaryFile(referenceExtraction);
  }
}

async function prepareRipCandidatesAndAudio(
  context: SubsyncContext,
  targetTrack: MpvTrack,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
  options: { allowSampledAudioFastPath: boolean },
): Promise<{
  audioSpeech: AudioSpeechResult | null;
  sampledAudioWindows: SampledAudioWindow[] | null;
  preExtractedEmbedded: Map<number, ParsedSubtitleFile> | null;
}> {
  const isLocal = !isRemoteMediaPath(context.videoPath) && fileExists(context.videoPath);
  const ffmpegAvailable = isLocal && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  if (!ffmpegAvailable) {
    return { audioSpeech: null, sampledAudioWindows: null, preExtractedEmbedded: null };
  }

  const embeddedSpecs = collectExtractableEmbeddedRefTracks(context, targetTrack);
  const vadPaths = deps.getVadPaths?.();
  const cachedFull = getCachedAudioSpeechSpans(
    context.videoPath,
    context.audioStreamIndex,
    vadPaths,
  );

  // Fast path: when a non-signs external or sidecar reference already exists and
  // full audio is not cached yet, sample 4 seeked 60s windows (~0.3s local / ~3s NAS)
  // before ever streaming the full container over disk/NFS for embedded tracks or full VAD.
  const loadedExternalPaths: string[] = [];
  for (const track of context.subtitleTracks) {
    if (track.external && track['external-filename']) {
      const resolvedPath = resolveLocalSubtitlePath(
        track['external-filename'],
        context.workingDirectory,
      );
      if (resolvedPath) loadedExternalPaths.push(resolvedPath);
    }
  }
  const hasExternalOrSidecarHint =
    context.subtitleTracks.some(
      (t) =>
        typeof t.id === 'number' &&
        t.id !== targetTrack.id &&
        t.external &&
        !isSignsOrSongsTrackMeta(t),
    ) ||
    discoverSidecarReferenceFiles(context.videoPath, loadedExternalPaths).some(
      (s) => !s.isSignsOnly,
    );

  if (
    options.allowSampledAudioFastPath &&
    !cachedFull &&
    hasExternalOrSidecarHint &&
    context.durationSec !== null &&
    context.durationSec !== undefined &&
    context.durationSec >= 300
  ) {
    const sampled = await extractSampledAudioWindows({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      durationSec: context.durationSec,
    });
    if (sampled && sampled.length >= 2) {
      return { audioSpeech: null, sampledAudioWindows: sampled, preExtractedEmbedded: null };
    }
  }

  // Single-pass demux when 1+ non-signs embedded tracks exist: reads the MKV ONCE
  // for all embedded subtitle tracks + full audio simultaneously.
  if (embeddedSpecs.length > 0) {
    const rip = await extractSinglePassRipMedia({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
      vadModelPath: vadPaths?.vadModelPath ?? null,
      embeddedTracks: embeddedSpecs,
      includeAudio: true,
    });
    if (rip) {
      return {
        audioSpeech: rip.audioSpeech,
        sampledAudioWindows: null,
        preExtractedEmbedded: rip.embeddedSubtitles,
      };
    }
    // Single-pass failed (e.g. a broken track in container): fall back cleanly.
    const audioSpeech = await resolveAudioSpeechForContext(context, resolved, deps);
    return { audioSpeech, sampledAudioWindows: null, preExtractedEmbedded: null };
  }

  if (cachedFull) {
    return { audioSpeech: cachedFull, sampledAudioWindows: null, preExtractedEmbedded: null };
  }

  const audioSpeech = await resolveAudioSpeechForContext(context, resolved, deps);
  return { audioSpeech, sampledAudioWindows: null, preExtractedEmbedded: null };
}

function collectRankedRemoteEmbeddedTracks(
  context: SubsyncContext,
  targetTrack: MpvTrack,
): Array<{ track: MpvTrack; ffIndex: number; extension: string; key: string; label: string }> {
  const metadataEvaluated: Array<{
    evaluated: EvaluatedSubtitleReference;
    track: MpvTrack;
    ffIndex: number;
    extension: string;
  }> = [];

  for (const track of context.subtitleTracks) {
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    if (track.external || isSignsOrSongsTrackMeta(track)) continue;
    const ffIndex = track['ff-index'];
    const extension = codecToExtension(track.codec);
    if (typeof ffIndex !== 'number' || !Number.isInteger(ffIndex) || ffIndex < 0 || !extension) {
      continue;
    }
    const isSecondary = context.secondaryTrack !== null && track.id === context.secondaryTrack.id;
    metadataEvaluated.push({
      track,
      ffIndex,
      extension,
      evaluated: {
        key: `track:${track.id}`,
        kind: 'track',
        trackId: track.id,
        filePath: null,
        label: buildTrackReferenceLabel(track, isSecondary),
        isSecondary,
        isEmbedded: true,
        isEnglish: isEnglishReferenceMeta({ lang: track.lang, title: track.title }),
        isFullDialogueHint: isFullDialogueHintMeta({ title: track.title }),
        isSignsOnly: false,
        dialogueCount: MIN_REFERENCE_DIALOGUE_LINES + 5,
        signCount: 0,
        audioStatus: 'unverified',
        audioOffsetMs: null,
        audioConfidence: 0,
      },
    });
  }

  const rankedKeys = rankReferenceCandidates(metadataEvaluated.map((item) => item.evaluated));
  const byKey = new Map(metadataEvaluated.map((item) => [item.evaluated.key, item]));
  return rankedKeys.map((ev) => {
    const item = byKey.get(ev.key)!;
    return {
      track: item.track,
      ffIndex: item.ffIndex,
      extension: item.extension,
      key: ev.key,
      label: ev.label,
    };
  });
}

export async function openSubsyncManualPicker(
  deps: TriggerSubsyncFromConfigDeps,
  options: {
    guidanceMessage?: string | null;
    skipAudioAnalysis?: boolean;
  } = {},
): Promise<boolean> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  const isRemote = isRemoteMediaPath(context.videoPath);

  let evaluatedCandidates: CandidateWithSpans[] | undefined;
  let audioSpeechAvailable: boolean | undefined;

  if (!options.skipAudioAnalysis) {
    const prepared = await prepareRipCandidatesAndAudio(
      context,
      context.primaryTrack,
      resolved,
      deps,
      {
        allowSampledAudioFastPath: true,
      },
    );
    audioSpeechAvailable =
      prepared.audioSpeech !== null ||
      prepared.sampledAudioWindows !== null ||
      (!isRemote && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']));
    evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
      context,
      context.primaryTrack,
      resolved,
      prepared.audioSpeech,
      {
        sampledAudioWindows: prepared.sampledAudioWindows,
        preExtractedEmbedded: prepared.preExtractedEmbedded,
        // Never stream a remote video over HTTP or full-read a local MKV during sampled fast-path just to populate the picker.
        skipEmbeddedExtraction:
          isRemote ||
          prepared.sampledAudioWindows !== null ||
          prepared.preExtractedEmbedded !== null,
        fetchImpl: deps.fetchImpl,
      },
    );
    if (!(await isPlaybackSessionUnchanged(client, context, context.primaryTrack, 'primary'))) {
      return false;
    }
  }

  const payload = buildManualPayload(context, resolved, deps, {
    evaluatedCandidates,
    audioSpeechAvailable,
    guidanceMessage: options.guidanceMessage ?? null,
  });
  deps.openManualPicker(payload);
  return true;
}

async function runAutoSubsync(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  const isRemote = isRemoteMediaPath(context.videoPath);

  const targetTrack = context.primaryTrack;
  const targetSlot = resolveTargetSlot(targetTrack, context);

  let targetExtraction: FileExtractionResult | null = null;
  let keepTempTarget = false;
  try {
    try {
      targetExtraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        targetTrack,
        context.workingDirectory,
        deps.fetchImpl,
      );
    } catch (extractError) {
      if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
        return;
      }
      const payload = buildManualPayload(context, resolved, deps, {
        guidanceMessage: `Automatic sync could not read the active subtitle track (${(extractError as Error).message}). Choose an engine and subtitle tracks below.`,
      });
      deps.openManualPicker(payload);
      deps.showMpvOsd('Subsync: choose engine and subtitles');
      return;
    }
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    const targetInfo = extractSpansFromParsedSubtitle(parsedTarget, { stripSigns: true });

    if (parsedTarget.events.length === 0 || targetInfo.spans.length === 0) {
      if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
        return;
      }
      const payload = buildManualPayload(context, resolved, deps, {
        guidanceMessage:
          'Could not read dialogue cues from the active subtitle track. Choose an engine or reference below.',
      });
      deps.openManualPicker(payload);
      deps.showMpvOsd('Subsync: choose engine and subtitles');
      return;
    }

    const prepared = await prepareRipCandidatesAndAudio(context, targetTrack, resolved, deps, {
      allowSampledAudioFastPath: true,
    });
    let audioSpeech = prepared.audioSpeech;
    let evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
      context,
      targetTrack,
      resolved,
      audioSpeech,
      {
        sampledAudioWindows: prepared.sampledAudioWindows,
        preExtractedEmbedded: prepared.preExtractedEmbedded,
        // Defer embedded extraction on remote streams (try external tracks first) and on
        // local sampled-audio fast path (try external/sidecar tracks before full MKV demux).
        skipEmbeddedExtraction:
          isRemote ||
          prepared.sampledAudioWindows !== null ||
          prepared.preExtractedEmbedded !== null,
        fetchImpl: deps.fetchImpl,
      },
    );

    const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
    const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);
    let lengthSec = context.durationSec ?? audioSpeech?.durationSec ?? 0;

    const applyAlignment = async (
      alignment: AlignmentResult,
      sourceLabel: string,
    ): Promise<void> => {
      const alreadyInSync = isAlreadyInSync(alignment);
      const delayProperty = targetSlot === 'secondary' ? 'secondary-sub-delay' : 'sub-delay';
      const activeDelay = alreadyInSync
        ? parseOptionalNumber(await client.requestProperty(delayProperty).catch(() => null))
        : null;
      if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
        return;
      }
      if (alreadyInSync) {
        if (activeDelay !== null && Math.abs(activeDelay) > 0.005) {
          client.send({ command: ['set_property', delayProperty, 0] });
          deps.showMpvOsd(`Reset subtitle delay to 0s (already in sync with ${sourceLabel})`);
          return;
        }
        deps.showMpvOsd(`Subtitles are already in sync (${sourceLabel})`);
        return;
      }
      if (parsedTarget.lossless === false) {
        deps.showMpvOsd(`Subsync failed: ${LOSSY_REWRITE_ERROR_MESSAGE}`);
        return;
      }
      writeRetimedContent(outputPath, alignment.rewrittenContent);
      if (targetExtraction!.temporary) {
        keepTempTarget = true;
      }
      await loadSyncedSubtitle(client, outputPath, targetSlot);
      deps.showMpvOsd(formatAlignmentSummary(sourceLabel, alignment));
    };

    // Fast path: if sampled audio already verified a clean aligned subtitle reference
    // and target alignment succeeds with high confidence, finish immediately without
    // reading the full video audio or embedded container tracks!
    if (!audioSpeech && prepared.sampledAudioWindows) {
      const fastAligned = evaluatedCandidates.filter(
        (c) =>
          c.evaluated.audioStatus === 'aligned' &&
          !c.evaluated.isSignsOnly &&
          c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES,
      );
      for (const candidate of fastAligned) {
        const alignment = alignSubtitleFileToReferenceSpans(parsedTarget, candidate.dialogueSpans, {
          stripSigns: true,
          allowSplits: true,
          calibration: 'clean',
          lengthSec,
          confidenceThreshold: AUTO_SUBTITLE_CONFIDENCE_THRESHOLD,
        });
        if (alignment.confident) {
          await applyAlignment(alignment, candidate.evaluated.label);
          return;
        }
      }

      // Sampled verification did not yield a confident aligned subtitle sync (e.g. all
      // external references were misaligned or sampled windows were inconclusive). Escalate
      // to a single full-episode pass for embedded tracks + complete audio verification.
      const escalated = await prepareRipCandidatesAndAudio(context, targetTrack, resolved, deps, {
        allowSampledAudioFastPath: false,
      });
      audioSpeech = escalated.audioSpeech;
      lengthSec = context.durationSec ?? audioSpeech?.durationSec ?? lengthSec;
      evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
        context,
        targetTrack,
        resolved,
        audioSpeech,
        {
          // Preserve sampled-audio verdicts (including misaligned references) whenever full audio is unavailable.
          sampledAudioWindows: audioSpeech ? null : prepared.sampledAudioWindows,
          preExtractedEmbedded: escalated.preExtractedEmbedded,
          fetchImpl: deps.fetchImpl,
        },
      );
    }

    let misalignedCandidates = evaluatedCandidates.filter(
      (c) => c.evaluated.audioStatus === 'misaligned' && !c.evaluated.isSignsOnly,
    );
    let alignedCandidates = evaluatedCandidates.filter(
      (c) =>
        c.evaluated.audioStatus === 'aligned' &&
        !c.evaluated.isSignsOnly &&
        c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES,
    );
    let unverifiedDialogueCandidates = evaluatedCandidates.filter(
      (c) =>
        c.evaluated.audioStatus === 'unverified' &&
        !c.evaluated.isSignsOnly &&
        c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES,
    );

    let bestAttempt: {
      alignment: AlignmentResult;
      sourceLabel: string;
      referenceKey: string;
    } | null = null;

    // 1. Best path: an audio-verified aligned dialogue subtitle reference (loaded or unloaded!).
    for (const candidate of alignedCandidates) {
      const alignment = alignSubtitleFileToReferenceSpans(parsedTarget, candidate.dialogueSpans, {
        stripSigns: true,
        allowSplits: true,
        calibration: 'clean',
        lengthSec,
        confidenceThreshold: AUTO_SUBTITLE_CONFIDENCE_THRESHOLD,
      });
      if (alignment.confident) {
        await applyAlignment(alignment, candidate.evaluated.label);
        return;
      }
      if (!bestAttempt || alignment.confidence > bestAttempt.alignment.confidence) {
        bestAttempt = {
          alignment,
          sourceLabel: candidate.evaluated.label,
          referenceKey: candidate.evaluated.key,
        };
      }
    }

    // 2. Fallback when no subtitle track is audio-aligned (none exist, only signs/songs, or all
    //    reference subtitles are shifted off the audio): sync directly to video audio speech!
    if (audioSpeech && audioSpeech.spans.length >= MIN_REFERENCE_DIALOGUE_LINES) {
      const audioAlignment = alignSubtitleFileToReferenceSpans(parsedTarget, audioSpeech.spans, {
        stripSigns: true,
        allowSplits: true,
        calibration: audioSpeech.calibration,
        leadSec: AUDIO_SPEECH_LEAD_SEC,
        lengthSec,
        confidenceThreshold: AUTO_AUDIO_CONFIDENCE_THRESHOLD,
      });
      const audioLabel = audioSpeech.source === 'silero-vad' ? 'audio (Silero VAD)' : 'video audio';
      if (audioAlignment.confident) {
        await applyAlignment(audioAlignment, audioLabel);
        return;
      }
      if (!bestAttempt || audioAlignment.confidence > bestAttempt.alignment.confidence) {
        bestAttempt = {
          alignment: audioAlignment,
          sourceLabel: audioLabel,
          referenceKey: 'video',
        };
      }
    }

    // 3. When audio verification wasn't available (e.g. remote stream or no ffmpeg):
    //    try already-loaded external dialogue tracks first with a stricter confidence bar.
    if (!audioSpeech) {
      for (const candidate of unverifiedDialogueCandidates) {
        const alignment = alignSubtitleFileToReferenceSpans(parsedTarget, candidate.dialogueSpans, {
          stripSigns: true,
          allowSplits: true,
          calibration: 'clean',
          lengthSec,
          confidenceThreshold: AUTO_REMOTE_CONFIDENCE_THRESHOLD,
        });
        if (alignment.confident) {
          await applyAlignment(alignment, candidate.evaluated.label);
          return;
        }
        if (!bestAttempt || alignment.confidence > bestAttempt.alignment.confidence) {
          bestAttempt = {
            alignment,
            sourceLabel: candidate.evaluated.label,
            referenceKey: candidate.evaluated.key,
          };
        }
      }

      // 3b. On remote streams where no external reference succeeded, fetch embedded tracks
      //     via Jellyfin's subtitle endpoint first (KB instead of GB), or fall back to a
      //     single ffmpeg pass across embedded candidates instead of N sequential downloads.
      if (isRemote) {
        const rankedEmbedded = collectRankedRemoteEmbeddedTracks(context, targetTrack);
        if (rankedEmbedded.length > 0) {
          const remoteExtracted = new Map<number, ParsedSubtitleFile>();
          const triedRemoteEmbedded = new Set<number>();

          const tryAlignRemoteEmbeddedTrack = async (
            item: (typeof rankedEmbedded)[number],
            parsed: ParsedSubtitleFile,
          ): Promise<boolean> => {
            remoteExtracted.set(item.ffIndex, parsed);
            triedRemoteEmbedded.add(item.ffIndex);

            const extractedSpans = extractSpansFromParsedSubtitle(parsed, { stripSigns: true });
            if (extractedSpans.dialogueCount < MIN_REFERENCE_DIALOGUE_LINES) return false;
            const alignment = alignSubtitleFileToReferenceSpans(
              parsedTarget,
              extractedSpans.spans,
              {
                stripSigns: true,
                allowSplits: true,
                calibration: 'clean',
                lengthSec,
                confidenceThreshold: AUTO_REMOTE_CONFIDENCE_THRESHOLD,
              },
            );
            if (alignment.confident) {
              await applyAlignment(alignment, item.label);
              return true;
            }
            if (!bestAttempt || alignment.confidence > bestAttempt.alignment.confidence) {
              bestAttempt = {
                alignment,
                sourceLabel: item.label,
                referenceKey: item.key,
              };
            }
            return false;
          };

          for (const item of rankedEmbedded) {
            if (
              !buildJellyfinEmbeddedSubtitleUrl(context.videoPath, item.ffIndex, item.extension)
            ) {
              break;
            }
            const bytes = await tryFetchJellyfinSubtitleBytes(
              context.videoPath,
              item.ffIndex,
              item.extension,
              deps.fetchImpl,
            );
            if (!bytes) continue;
            const decoded = decodeSubtitleBytesWithStatus(bytes);
            const parsed = parseSubtitleFileContent(
              decoded.text,
              `track_${item.ffIndex}.${item.extension}`,
            );
            if (parsed.events.length === 0) continue;
            if (await tryAlignRemoteEmbeddedTrack(item, parsed)) {
              return;
            }
          }

          const remainingItems = rankedEmbedded.filter(
            (item) => !remoteExtracted.has(item.ffIndex),
          );
          const remainingSpecs = remainingItems.map((item) => ({
            ffIndex: item.ffIndex,
            extension: item.extension,
          }));

          if (remainingSpecs.length > 0 && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg'])) {
            const singlePass = await extractSinglePassRipMedia({
              ffmpegPath: resolved.ffmpegPath,
              videoPath: context.videoPath,
              audioStreamIndex: context.audioStreamIndex,
              embeddedTracks: remainingSpecs,
              includeAudio: false,
            });
            if (singlePass) {
              for (const item of remainingItems) {
                const parsed = singlePass.embeddedSubtitles.get(item.ffIndex);
                if (!parsed) continue;
                if (await tryAlignRemoteEmbeddedTrack(item, parsed)) {
                  return;
                }
              }
            }

            // Retry any ranked track omitted by the combined pass (whether the pass returned null
            // or a partial map) individually in rank order, stopping on the first confident match.
            const missingItems = remainingItems.filter(
              (item) => !remoteExtracted.has(item.ffIndex),
            );
            for (const fallbackItem of missingItems) {
              let extraction: FileExtractionResult | null = null;
              let parsed: ParsedSubtitleFile | null = null;
              try {
                extraction = await extractSubtitleTrackToFile(
                  resolved.ffmpegPath,
                  context.videoPath,
                  fallbackItem.track,
                  context.workingDirectory,
                  deps.fetchImpl,
                );
                parsed = readAndParseSubtitleFile(extraction.path);
              } catch {
                // Continue to lower-ranked sibling tracks if extracting or parsing this track fails.
                continue;
              } finally {
                cleanupTemporaryFile(extraction);
              }
              if (parsed.events.length === 0) continue;
              if (await tryAlignRemoteEmbeddedTrack(fallbackItem, parsed)) {
                return;
              }
            }
          }

          if (remoteExtracted.size > 0) {
            evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
              context,
              targetTrack,
              resolved,
              null,
              {
                preExtractedEmbedded: remoteExtracted,
                skipEmbeddedExtraction: true,
                fetchImpl: deps.fetchImpl,
              },
            );
            misalignedCandidates = evaluatedCandidates.filter(
              (c) => c.evaluated.audioStatus === 'misaligned' && !c.evaluated.isSignsOnly,
            );
            alignedCandidates = evaluatedCandidates.filter(
              (c) =>
                c.evaluated.audioStatus === 'aligned' &&
                !c.evaluated.isSignsOnly &&
                c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES,
            );
            unverifiedDialogueCandidates = evaluatedCandidates.filter(
              (c) =>
                c.evaluated.audioStatus === 'unverified' &&
                !c.evaluated.isSignsOnly &&
                c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES,
            );

            for (const candidate of unverifiedDialogueCandidates) {
              if (!candidate.evaluated.isEmbedded) continue;
              const matchedItem = rankedEmbedded.find(
                (item) => item.key === candidate.evaluated.key,
              );
              if (matchedItem && triedRemoteEmbedded.has(matchedItem.ffIndex)) continue;
              const alignment = alignSubtitleFileToReferenceSpans(
                parsedTarget,
                candidate.dialogueSpans,
                {
                  stripSigns: true,
                  allowSplits: true,
                  calibration: 'clean',
                  lengthSec,
                  confidenceThreshold: AUTO_REMOTE_CONFIDENCE_THRESHOLD,
                },
              );
              if (alignment.confident) {
                await applyAlignment(alignment, candidate.evaluated.label);
                return;
              }
              if (!bestAttempt || alignment.confidence > bestAttempt.alignment.confidence) {
                bestAttempt = {
                  alignment,
                  sourceLabel: candidate.evaluated.label,
                  referenceKey: candidate.evaluated.key,
                };
              }
            }
          }
        }
      }
    }

    // 4. Automatic sync could not finish with high confidence -> open guided modal with explanation!
    const reasons: string[] = [];
    if (misalignedCandidates.length > 0 && alignedCandidates.length === 0) {
      const sample = misalignedCandidates[0]!.evaluated;
      const shiftStr =
        sample.audioOffsetMs !== null ? formatOffsetSummary(sample.audioOffsetMs) : '';
      reasons.push(
        `Reference subtitle "${sample.label}" appears misaligned with the audio${shiftStr ? ` (${shiftStr})` : ''}.`,
      );
    } else if (
      evaluatedCandidates.length > 0 &&
      alignedCandidates.length === 0 &&
      unverifiedDialogueCandidates.length === 0
    ) {
      reasons.push(
        'Available subtitle tracks look like signs/songs or have too few dialogue lines.',
      );
    } else if (evaluatedCandidates.length === 0 && !audioSpeech) {
      reasons.push('No reference subtitle track or local audio source was available.');
    }

    if (bestAttempt) {
      const pct = Math.round(bestAttempt.alignment.confidence * 100);
      reasons.push(
        `Best automatic match (${bestAttempt.sourceLabel}) had low confidence (${pct}%).`,
      );
    }
    reasons.push('Pick a reference, browse a subtitle file, or adjust sync options below.');

    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
      return;
    }

    const payload = buildManualPayload(context, resolved, deps, {
      evaluatedCandidates,
      audioSpeechAvailable: audioSpeech !== null,
      guidanceMessage: reasons.join(' '),
      defaultEngine: 'builtin',
      defaultReferenceKey: bestAttempt?.referenceKey ?? null,
    });
    deps.openManualPicker(payload);
    deps.showMpvOsd('Subsync needs guidance: opened options');
  } finally {
    if (!keepTempTarget) {
      cleanupTemporaryFile(targetExtraction);
    }
  }
}

export async function triggerSubsyncFromConfig(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  if (deps.isSubsyncInProgress()) {
    deps.showMpvOsd('Subsync already running');
    return;
  }

  deps.setSubsyncInProgress(true);
  try {
    const resolved = deps.getResolvedConfig();
    if (resolved.mode === 'manual') {
      const opened = await openSubsyncManualPicker(deps, { skipAudioAnalysis: true });
      if (opened) {
        deps.showMpvOsd('Subsync: choose engine and subtitles');
      }
      return;
    }

    await deps.runWithSubsyncSpinner(() => runAutoSubsync(deps));
  } catch (error) {
    deps.showMpvOsd(`Subsync failed: ${(error as Error).message}`);
  } finally {
    deps.setSubsyncInProgress(false);
  }
}
