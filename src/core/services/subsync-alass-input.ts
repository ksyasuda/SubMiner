import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { FileExtractionResult } from './subsync-extract';
import { createLogger } from '../../logger';

const logger = createLogger('subsync');

/**
 * alass picks its parser purely from the file extension, and WebVTT is not one
 * of the formats it knows. A `.vtt` reference is therefore treated as a *video*
 * ("no audio stream in file"), and the same file renamed to `.srt` dies in the
 * SubRip parser on the `WEBVTT` header. Extension-backed anime streams hand out
 * VTT almost exclusively, so every alass input is converted to SRT first.
 */
const ALASS_SUBTITLE_EXTENSIONS = new Set(['srt', 'ass', 'ssa', 'sub', 'idx']);

/** How much of the file is decoded to recognise a WebVTT header. */
const SNIFF_BYTES = 1024;

const CUE_TIMING_PATTERN =
  /^\s*(?:(\d{1,3}):)?(\d{1,2}):(\d{2})[.,](\d{1,3})\s*-->\s*(?:(\d{1,3}):)?(\d{1,2}):(\d{2})[.,](\d{1,3})/;

interface RawCue {
  start: number;
  end: number;
  text: string;
}

function toSeconds(
  hours: string | undefined,
  minutes: string,
  seconds: string,
  millis: string,
): number {
  return (
    Number(hours ?? 0) * 3600 +
    Number(minutes) * 60 +
    Number(seconds) +
    Number(millis.padEnd(3, '0')) / 1000
  );
}

function formatTimestamp(totalSeconds: number): string {
  const clamped = Math.max(0, totalSeconds);
  const millis = Math.round(clamped * 1000);
  const hours = Math.floor(millis / 3_600_000);
  const minutes = Math.floor((millis % 3_600_000) / 60_000);
  const seconds = Math.floor((millis % 60_000) / 1000);
  const remainder = millis % 1000;
  const pad = (value: number, width: number): string => String(value).padStart(width, '0');
  return `${pad(hours, 2)}:${pad(minutes, 2)}:${pad(seconds, 2)},${pad(remainder, 3)}`;
}

/**
 * Read timed cues out of a WebVTT (or SRT-shaped) file.
 *
 * The payload is kept verbatim rather than sanitised: this same file is what
 * alass retimes and mpv then loads, so stripping inline tags here would show up
 * on screen. Cue identifiers, `NOTE`/`STYLE`/`REGION` blocks and the `WEBVTT`
 * header all fall out for free — only lines that follow a timing line are kept.
 */
export function parseTimedCues(content: string): RawCue[] {
  const lines = content.replace(/^﻿/, '').split(/\r?\n/);
  const cues: RawCue[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const timing = CUE_TIMING_PATTERN.exec(lines[index]!);
    if (!timing) continue;

    const start = toSeconds(timing[1], timing[2]!, timing[3]!, timing[4]!);
    const end = toSeconds(timing[5], timing[6]!, timing[7]!, timing[8]!);

    const textLines: string[] = [];
    index += 1;
    while (index < lines.length && lines[index]!.trim().length > 0) {
      textLines.push(lines[index]!);
      index += 1;
    }

    const text = textLines.join('\n').trim();
    if (text.length > 0) {
      cues.push({ start, end, text });
    }
  }

  return cues;
}

export function formatCuesAsSrt(cues: RawCue[]): string {
  return cues
    .map(
      (cue, index) =>
        `${index + 1}\n${formatTimestamp(cue.start)} --> ${formatTimestamp(cue.end)}\n${cue.text}\n`,
    )
    .join('\n');
}

function readHead(filePath: string): string {
  const handle = fs.openSync(filePath, 'r');
  try {
    const buffer = Buffer.alloc(SNIFF_BYTES);
    const bytesRead = fs.readSync(handle, buffer, 0, SNIFF_BYTES, 0);
    return buffer.subarray(0, bytesRead).toString('utf8');
  } finally {
    fs.closeSync(handle);
  }
}

function isWebVttContent(head: string): boolean {
  return /^WEBVTT(\s|$)/.test(head.replace(/^﻿/, '').trimStart());
}

/**
 * Decide whether alass can read the file as-is.
 *
 * Content wins over the extension in one direction only: a file alass would
 * accept by extension still needs converting when it turns out to hold VTT,
 * while an unrecognised extension is always converted. Anything else is passed
 * through untouched, which also keeps its original charset intact — alass
 * detects encodings that a UTF-8 round trip here would mangle.
 */
export function needsAlassConversion(filePath: string, head: string): boolean {
  const extension = path.extname(filePath).slice(1).toLowerCase();
  if (!ALASS_SUBTITLE_EXTENSIONS.has(extension)) return true;
  return isWebVttContent(head);
}

/**
 * Rewrite a subtitle file as SRT for alass, or return null when alass can
 * already read it. The result is a temporary file in its own directory, so
 * `cleanupTemporaryFile` can remove it (and keep the retimed output beside it).
 */
export function convertSubtitleForAlass(filePath: string): FileExtractionResult | null {
  if (!needsAlassConversion(filePath, readHead(filePath))) return null;

  const cues = parseTimedCues(fs.readFileSync(filePath, 'utf8'));
  if (cues.length === 0) {
    throw new Error(`Could not read subtitle timings for alass from ${path.basename(filePath)}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-alass-'));
  const outputPath = path.join(tempDir, `${path.parse(filePath).name}.srt`);
  try {
    fs.writeFileSync(outputPath, formatCuesAsSrt(cues), 'utf8');
  } catch (error) {
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch {}
    throw error;
  }

  logger.info(`Converted ${filePath} to SRT for alass: ${outputPath} (${cues.length} cues)`);
  return { path: outputPath, temporary: true };
}
