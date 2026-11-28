import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  buildJellyfinEmbeddedSubtitleUrl,
  CommandResult,
  canFfmpegExtractSubtitleExtension,
  codecToExtension,
  extensionToFfmpegMuxer,
  fileExists,
  MpvTrack,
  runCommand,
} from '../../subsync/utils';
import {
  decodeSubtitleBytesWithStatus,
  extractSpansFromParsedSubtitle,
  formatSpansAsSrt,
  parseSubtitleFileContent,
  type ParsedSubtitleFile,
} from '../../subsync/subtitle-lines';
import { isRemoteMediaPath } from '../../jimaku/utils';
import {
  normalizeSubtitlePathForCompare,
  normalizeTrackIds,
  resolveLocalSubtitlePath,
} from './subsync-context';
import {
  type FileExtractionResult,
  JELLYFIN_SUBTITLE_FETCH_TIMEOUT_MS,
  MIN_REFERENCE_DIALOGUE_LINES,
  type MpvClientLike,
  SYNCED_TRACK_LOOKUP_ATTEMPTS,
  SYNCED_TRACK_LOOKUP_RETRY_MS,
  type SubsyncIo,
  type SubtitleSlot,
  delay,
  ensureExecutablePath,
  summarizeCommandFailure,
} from './subsync-shared';

export async function tryFetchJellyfinSubtitleBytes(
  videoPath: string,
  ffIndex: number,
  extension: string,
  io: SubsyncIo = {},
): Promise<Uint8Array | null> {
  const subtitleUrl = buildJellyfinEmbeddedSubtitleUrl(videoPath, ffIndex, extension);
  const fetchImpl = io.fetchImpl ?? globalThis.fetch;
  if (!subtitleUrl || typeof fetchImpl !== 'function') return null;

  const timeoutSignal = AbortSignal.timeout(JELLYFIN_SUBTITLE_FETCH_TIMEOUT_MS);
  const signal = io.signal ? AbortSignal.any([timeoutSignal, io.signal]) : timeoutSignal;
  try {
    const response = await fetchImpl(subtitleUrl, { signal });
    if (!response.ok) return null;
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength === 0) return null;
    return new Uint8Array(buffer);
  } catch {
    return null;
  }
}

export async function extractSubtitleTrackToFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvTrack,
  workingDirectory?: string | null,
  io: SubsyncIo = {},
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
    const fetchedBytes = await tryFetchJellyfinSubtitleBytes(videoPath, ffIndex, extension, io);
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

  if (!canFfmpegExtractSubtitleExtension(extension)) {
    throw new Error(
      `${extension.toUpperCase()} subtitle tracks cannot be extracted locally (ffmpeg has no decoder for them)`,
    );
  }
  const resolvedFfmpegPath = ensureExecutablePath(ffmpegPath, 'ffmpeg');
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  try {
    const extraction = await runCommand(
      resolvedFfmpegPath,
      [
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
      ],
      undefined,
      io.signal,
    );

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

const SUBSYNC_TEMP_DIR_PREFIX = 'subminer-subsync-';
const STALE_SUBSYNC_TEMP_DIR_MS = 24 * 60 * 60 * 1000;

/**
 * Temp dirs still holding a retimed subtitle that mpv loaded (embedded or remote target).
 * Pruned by `pruneRetainedSubsyncTempDirs` once mpv no longer lists the file, and on exit.
 */
const retainedSubsyncTempDirs = new Set<string>();
let exitCleanupRegistered = false;
let staleSweepDone = false;

function removeDirQuietly(dir: string): void {
  try {
    fs.rmSync(dir, { recursive: true, force: true });
  } catch {
    // Best-effort temp cleanup.
  }
}

function retainSubsyncTempDir(dir: string): void {
  retainedSubsyncTempDirs.add(dir);
  if (exitCleanupRegistered) return;
  exitCleanupRegistered = true;
  process.once('exit', () => {
    for (const retained of retainedSubsyncTempDirs) removeDirQuietly(retained);
  });
}

/**
 * Remove a temporary extraction and its mkdtemp dir. A dir that still holds the retimed
 * subtitle mpv loaded stays on disk (mpv and the subtitle sidebar read it) and is retained.
 */
export function cleanupTemporaryFile(extraction: FileExtractionResult | null): void {
  if (!extraction || !extraction.temporary) return;
  removeIfExists(extraction.path);
  const dir = path.dirname(extraction.path);
  try {
    fs.rmdirSync(dir);
  } catch {
    if (fs.existsSync(dir)) retainSubsyncTempDir(dir);
  }
}

/**
 * Drop retained temp dirs whose files mpv no longer lists (`referencedPaths` are the current
 * external subtitle filenames), and once per process sweep subsync temp dirs a crashed
 * session left behind more than a day ago.
 */
export function pruneRetainedSubsyncTempDirs(
  referencedPaths: readonly string[],
  options: { now?: number; tmpRoot?: string } = {},
): void {
  const isReferenced = (dir: string): boolean => {
    const prefix = `${dir}${path.sep}`;
    return referencedPaths.some((filePath) => filePath.startsWith(prefix));
  };
  for (const dir of retainedSubsyncTempDirs) {
    if (isReferenced(dir)) continue;
    removeDirQuietly(dir);
    // Keep failed removals registered so the next prune or exit cleanup retries them.
    if (!fs.existsSync(dir)) retainedSubsyncTempDirs.delete(dir);
  }
  if (staleSweepDone) return;
  staleSweepDone = true;
  const tmpRoot = options.tmpRoot ?? os.tmpdir();
  const now = options.now ?? Date.now();
  let entries: string[] = [];
  try {
    entries = fs.readdirSync(tmpRoot);
  } catch {
    return;
  }
  for (const name of entries) {
    if (!name.startsWith(SUBSYNC_TEMP_DIR_PREFIX)) continue;
    const dir = path.join(tmpRoot, name);
    if (retainedSubsyncTempDirs.has(dir)) continue;
    // mpv may still play a file from an old dir (e.g. the app restarted under a running mpv):
    // keep it until mpv drops the file, then the regular prune removes it.
    if (isReferenced(dir)) {
      retainSubsyncTempDir(dir);
      continue;
    }
    try {
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > STALE_SUBSYNC_TEMP_DIR_MS) {
        removeDirQuietly(dir);
      }
    } catch {
      // Raced with another cleanup; ignore.
    }
  }
}

const MAX_REFERENCE_FILE_BYTES = 20 * 1024 * 1024;

/**
 * Validate a renderer-supplied reference subtitle path before main reads it synchronously:
 * absolute regular file under a subtitle-sized cap (rejects directories, FIFOs, and videos
 * picked by mistake through the "All files" filter). Returns the path or null.
 */
export function resolveReferenceFilePath(rawPath: string | null | undefined): string | null {
  const refPath = rawPath?.trim() ?? '';
  if (!refPath || !path.isAbsolute(refPath)) return null;
  try {
    const stat = fs.statSync(refPath);
    return stat.isFile() && stat.size <= MAX_REFERENCE_FILE_BYTES ? refPath : null;
  } catch {
    return null;
  }
}

export function buildRetimedPath(subPath: string, replace: boolean): string {
  if (replace) return subPath;
  const parsed = path.parse(subPath);
  return path.join(parsed.dir, `${parsed.name}_retimed${parsed.ext || '.srt'}`);
}

export const LOSSY_REWRITE_ERROR_MESSAGE =
  'Cannot safely rewrite subtitle file because its text encoding could not be decoded unambiguously and losslessly. Convert the file to UTF-8 and retry.';

/** Size + mtime of a subtitle file, compared before an in-place rewrite to catch edits made mid-sync. */
export function readSubtitleFileStamp(filePath: string): string | null {
  try {
    const stat = fs.statSync(filePath);
    return `${stat.size}\0${stat.mtimeMs}`;
  } catch {
    return null;
  }
}

export type ReadSubtitleFile = ParsedSubtitleFile & {
  /** `readSubtitleFileStamp` taken before reading, for `RetimedWrite.sourceStamp`. */
  stamp: string | null;
};

export function readAndParseSubtitleFile(filePath: string): ReadSubtitleFile {
  const stamp = readSubtitleFileStamp(filePath);
  const bytes = fs.readFileSync(filePath);
  const decoded = decodeSubtitleBytesWithStatus(bytes);
  return {
    ...parseSubtitleFileContent(decoded.text, filePath),
    lossless: decoded.lossless,
    stamp,
  };
}

export interface RetimedWrite {
  /** In-place `subsync.replace` rewrite of the user's subtitle (backup + source-change check). */
  replacesSource: boolean;
  /** Stamp of the source when it was read; a different stamp now means it was edited mid-sync. */
  sourceStamp?: string | null;
}

const SOURCE_CHANGED_MESSAGE =
  'The subtitle file changed on disk while syncing; discarded the result. Run subsync again.';

/** Suffix for the one-time backup of a replaced subtitle; not a subtitle extension, so mpv and sidecar discovery ignore it. */
const SUBSYNC_BACKUP_SUFFIX = '.orig';

function resolveRealPathOrSelf(filePath: string): string {
  try {
    return fs.realpathSync(filePath);
  } catch {
    return filePath;
  }
}

/**
 * Move a staged retimed file over `outputPath`. Writes through a symlink to the real file
 * and keeps its permissions. When `replacesSource` is set (in-place `subsync.replace`), the
 * first replacement keeps the untouched original as `<file>.orig`; later runs leave it alone.
 */
export function installRetimedFile(
  stagedPath: string,
  outputPath: string,
  { replacesSource, sourceStamp }: RetimedWrite,
): void {
  const targetPath = resolveRealPathOrSelf(outputPath);
  if (
    replacesSource &&
    sourceStamp !== undefined &&
    readSubtitleFileStamp(targetPath) !== sourceStamp
  ) {
    throw new Error(SOURCE_CHANGED_MESSAGE);
  }
  const existing = fs.existsSync(targetPath) ? fs.statSync(targetPath) : null;
  if (existing) {
    if (replacesSource) {
      const backupPath = `${targetPath}${SUBSYNC_BACKUP_SUFFIX}`;
      if (!fs.existsSync(backupPath)) {
        fs.copyFileSync(targetPath, backupPath, fs.constants.COPYFILE_EXCL);
      }
    }
    fs.chmodSync(stagedPath, existing.mode & 0o7777);
  }
  try {
    fs.renameSync(stagedPath, targetPath);
  } catch (error) {
    // A staged file next to a symlink can sit on another filesystem than the link target:
    // copy it next to the target first so the final swap is still an atomic rename.
    if ((error as NodeJS.ErrnoException).code !== 'EXDEV') throw error;
    const localStaged = `${targetPath}.subminer-subsync-${process.pid}-${Date.now()}.tmp`;
    try {
      fs.copyFileSync(stagedPath, localStaged);
      if (existing) fs.chmodSync(localStaged, existing.mode & 0o7777);
      fs.renameSync(localStaged, targetPath);
    } catch (copyError) {
      removeIfExists(localStaged);
      throw copyError;
    }
    fs.unlinkSync(stagedPath);
  }
}

export function removeIfExists(filePath: string): void {
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  } catch {
    // Best-effort cleanup of a staged write artifact.
  }
}

export function writeRetimedContent(
  outputPath: string,
  content: string,
  write: RetimedWrite,
): void {
  const dir = path.dirname(resolveRealPathOrSelf(outputPath));
  const tmpPath = path.join(
    dir,
    `.subminer-subsync-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`,
  );
  try {
    fs.writeFileSync(tmpPath, content, 'utf8');
    installRetimedFile(tmpPath, outputPath, write);
  } catch (error) {
    removeIfExists(tmpPath);
    throw error;
  }
}

export async function runAlassSync(
  alassPath: string,
  referenceFile: string,
  inputSubtitlePath: string,
  outputPath: string,
  signal?: AbortSignal,
): Promise<CommandResult> {
  return runCommand(alassPath, [referenceFile, inputSubtitlePath, outputPath], undefined, signal);
}

export async function runFfsubsyncSync(
  ffsubsyncPath: string,
  videoPath: string,
  inputSubtitlePath: string,
  outputPath: string,
  audioStreamIndex: number | null,
  signal?: AbortSignal,
): Promise<CommandResult> {
  const args = [videoPath, '-i', inputSubtitlePath, '-o', outputPath];
  if (audioStreamIndex !== null) {
    args.push('--reference-stream', `0:${audioStreamIndex}`);
  }
  return runCommand(ffsubsyncPath, args, undefined, signal);
}

async function readSubtitleTracks(client: MpvClientLike): Promise<MpvTrack[]> {
  const trackListRaw = await client.requestProperty('track-list');
  return Array.isArray(trackListRaw)
    ? normalizeTrackIds(trackListRaw as MpvTrack[]).filter((track) => track.type === 'sub')
    : [];
}

/**
 * Find the track `sub-add` created for `pathToLoad`. Only ids missing from `knownTrackIds`
 * (the list before the add) count, so an older track for the same file is never picked.
 */
async function findAddedSubtitleTrackId(
  client: MpvClientLike,
  pathToLoad: string,
  knownTrackIds: ReadonlySet<number>,
): Promise<number | null> {
  const wanted = normalizeSubtitlePathForCompare(pathToLoad);
  // sub-add is queued, so the track may not appear in the first track-list reply.
  for (let attempt = 0; attempt < SYNCED_TRACK_LOOKUP_ATTEMPTS; attempt += 1) {
    let tracks: MpvTrack[] = [];
    try {
      tracks = await readSubtitleTracks(client);
    } catch {
      return null;
    }
    const added = tracks.filter((track) => {
      const filename = track['external-filename'];
      return (
        typeof track.id === 'number' &&
        !knownTrackIds.has(track.id) &&
        typeof filename === 'string' &&
        normalizeSubtitlePathForCompare(filename) === wanted
      );
    });
    const newest = added[added.length - 1];
    if (newest && typeof newest.id === 'number') {
      return newest.id;
    }
    if (attempt < SYNCED_TRACK_LOOKUP_ATTEMPTS - 1) {
      await delay(SYNCED_TRACK_LOOKUP_RETRY_MS);
    }
  }
  return null;
}

/** `sub-add` args that keep the retimed track's original title/lang (used for ranking and labels). */
function buildSubAddCommand(
  pathToLoad: string,
  flag: 'select' | 'auto',
  track: MpvTrack,
): string[] {
  const title = track.title?.trim() ?? '';
  const lang = track.lang?.trim() ?? '';
  if (!title && !lang)
    return flag === 'select' ? ['sub-add', pathToLoad] : ['sub-add', pathToLoad, flag];
  return ['sub-add', pathToLoad, flag, title || path.basename(pathToLoad), lang];
}

export async function loadSyncedSubtitle(
  client: MpvClientLike,
  pathToLoad: string,
  slot: SubtitleSlot,
  sourceTrack: MpvTrack,
  signal?: AbortSignal,
): Promise<void> {
  if (!client.connected) {
    throw new Error('MPV disconnected while loading subtitle');
  }

  if (slot === 'secondary') {
    // Keep the primary track untouched: load without selecting, then point
    // secondary-sid at the freshly added track.
    // Without a baseline an older track for the same file could pass as the new one.
    const knownTrackIds = new Set(
      (await readSubtitleTracks(client)).flatMap((track) =>
        typeof track.id === 'number' ? [track.id] : [],
      ),
    );
    signal?.throwIfAborted();
    client.send({ command: buildSubAddCommand(pathToLoad, 'auto', sourceTrack) });
    const addedTrackId = await findAddedSubtitleTrackId(client, pathToLoad, knownTrackIds);
    if (addedTrackId === null) {
      throw new Error('Synchronized subtitle did not appear in the mpv track list');
    }
    signal?.throwIfAborted();
    client.send({ command: ['set_property', 'secondary-sid', addedTrackId] });
    // The retimed file carries the timing now; only drop the old delay once it is selected.
    client.send({ command: ['set_property', 'secondary-sub-delay', 0] });
    return;
  }

  client.send({ command: buildSubAddCommand(pathToLoad, 'select', sourceTrack) });
  client.send({ command: ['set_property', 'sub-delay', 0] });
}

export function validateFfsubsyncReference(videoPath: string): void {
  if (isRemoteMediaPath(videoPath)) {
    throw new Error(
      'FFsubsync cannot reliably sync stream URLs because it needs direct reference media access. Use built-in or alass sync with a subtitle track, or play a local file.',
    );
  }
}

export function maybeWriteCleanedReferenceSrt(
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
    try {
      fs.writeFileSync(cleanPath, formatSpansAsSrt(extracted.spans), 'utf8');
    } catch (error) {
      removeDirQuietly(tempDir);
      throw error;
    }
    return { path: cleanPath, temporary: true };
  } catch {
    return null;
  }
}
