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

export function cleanupTemporaryFile(extraction: FileExtractionResult | null): void {
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

export function readAndParseSubtitleFile(filePath: string): ParsedSubtitleFile {
  const bytes = fs.readFileSync(filePath);
  const decoded = decodeSubtitleBytesWithStatus(bytes);
  return {
    ...parseSubtitleFileContent(decoded.text, filePath),
    lossless: decoded.lossless,
  };
}

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
  replacesSource: boolean,
): void {
  const targetPath = resolveRealPathOrSelf(outputPath);
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
    // A staged file next to a symlink can sit on another filesystem than the link target.
    if ((error as NodeJS.ErrnoException).code !== 'EXDEV') throw error;
    fs.copyFileSync(stagedPath, targetPath);
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
  replacesSource: boolean,
): void {
  const dir = path.dirname(resolveRealPathOrSelf(outputPath));
  const tmpPath = path.join(
    dir,
    `.subminer-subsync-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`,
  );
  try {
    fs.writeFileSync(tmpPath, content, 'utf8');
    installRetimedFile(tmpPath, outputPath, replacesSource);
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
): Promise<void> {
  if (!client.connected) {
    throw new Error('MPV disconnected while loading subtitle');
  }

  if (slot === 'secondary') {
    // Keep the primary track untouched: load without selecting, then point
    // secondary-sid at the freshly added track.
    client.send({ command: buildSubAddCommand(pathToLoad, 'auto', sourceTrack) });
    client.send({ command: ['set_property', 'secondary-sub-delay', 0] });
    const addedTrackId = await findAddedSubtitleTrackId(client, pathToLoad);
    if (addedTrackId === null) {
      throw new Error('Synchronized subtitle did not appear in the mpv track list');
    }
    client.send({ command: ['set_property', 'secondary-sid', addedTrackId] });
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
    fs.writeFileSync(cleanPath, formatSpansAsSrt(extracted.spans), 'utf8');
    return { path: cleanPath, temporary: true };
  } catch {
    return null;
  }
}
