import { execFile as nodeExecFile } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import type { AnkiConnectConfig } from '../types/anki';

type NoteInfoLike = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

interface ResolveAnimatedImageLeadInSecondsArgs<TNoteInfo extends NoteInfoLike> {
  config: Pick<AnkiConnectConfig, 'fields' | 'media'>;
  noteInfo: TNoteInfo;
  resolveConfiguredFieldName: (
    noteInfo: TNoteInfo,
    ...preferredNames: (string | undefined)[]
  ) => string | null;
  retrieveMediaFileBase64: (filename: string) => Promise<string>;
  probeAudioDurationSeconds?: (buffer: Buffer, filename: string) => Promise<number | null>;
  logWarn?: (message: string, ...args: unknown[]) => void;
}

interface ProbeAudioDurationDeps {
  execFile?: typeof nodeExecFile;
  mkdtempSync?: typeof fs.mkdtempSync;
  writeFileSync?: typeof fs.writeFileSync;
  rmSync?: typeof fs.rmSync;
}

export function extractSoundFilenames(value: string): string[] {
  const matches = value.matchAll(/\[sound:([^\]]+)\]/gi);
  return Array.from(matches, (match) => match[1]?.trim() || '').filter((value) => value.length > 0);
}

function shouldSyncAnimatedImageToWordAudio(config: Pick<AnkiConnectConfig, 'media'>): boolean {
  return config.media?.imageType === 'avif' && config.media?.syncAnimatedImageToWordAudio !== false;
}

function resolveSentenceAudioStartOffsetSeconds(config: Pick<AnkiConnectConfig, 'media'>): number {
  const configuredPadding = config.media?.audioPadding;
  if (typeof configuredPadding === 'number' && Number.isFinite(configuredPadding)) {
    return configuredPadding;
  }
  return DEFAULT_ANKI_CONNECT_CONFIG.media.audioPadding;
}

export async function probeAudioDurationSeconds(
  buffer: Buffer,
  filename: string,
  deps: ProbeAudioDurationDeps = {},
): Promise<number | null> {
  const execFile = deps.execFile ?? nodeExecFile;
  const mkdtempSync = deps.mkdtempSync ?? fs.mkdtempSync;
  const writeFileSync = deps.writeFileSync ?? fs.writeFileSync;
  const rmSync = deps.rmSync ?? fs.rmSync;

  const tempDir = mkdtempSync(path.join(os.tmpdir(), 'subminer-audio-probe-'));
  const ext = path.extname(filename) || '.bin';
  const tempPath = path.join(tempDir, `probe${ext}`);
  writeFileSync(tempPath, buffer);

  return new Promise((resolve) => {
    execFile(
      'ffprobe',
      [
        '-v',
        'error',
        '-show_entries',
        'format=duration',
        '-of',
        'default=noprint_wrappers=1:nokey=1',
        tempPath,
      ],
      (error, stdout) => {
        try {
          if (error) {
            resolve(null);
            return;
          }

          const durationSeconds = Number.parseFloat((stdout || '').trim());
          resolve(Number.isFinite(durationSeconds) && durationSeconds > 0 ? durationSeconds : null);
        } finally {
          rmSync(tempDir, { recursive: true, force: true });
        }
      },
    );
  });
}

export async function resolveAnimatedImageLeadInSeconds<TNoteInfo extends NoteInfoLike>({
  config,
  noteInfo,
  resolveConfiguredFieldName,
  retrieveMediaFileBase64,
  probeAudioDurationSeconds: probeDuration = probeAudioDurationSeconds,
  logWarn,
}: ResolveAnimatedImageLeadInSecondsArgs<TNoteInfo>): Promise<number> {
  if (!shouldSyncAnimatedImageToWordAudio(config)) {
    return 0;
  }

  const wordAudioFieldName = resolveConfiguredFieldName(
    noteInfo,
    config.fields?.audio,
    DEFAULT_ANKI_CONNECT_CONFIG.fields.audio,
  );
  if (!wordAudioFieldName) {
    return 0;
  }

  const wordAudioValue = noteInfo.fields[wordAudioFieldName]?.value || '';
  const filenames = extractSoundFilenames(wordAudioValue);
  if (filenames.length === 0) {
    return 0;
  }

  let totalLeadInSeconds = 0;
  for (const filename of filenames) {
    const encoded = await retrieveMediaFileBase64(filename);
    if (!encoded) {
      logWarn?.('Animated image sync skipped: failed to retrieve word audio', filename);
      return 0;
    }

    const durationSeconds = await probeDuration(Buffer.from(encoded, 'base64'), filename);
    if (!(typeof durationSeconds === 'number' && Number.isFinite(durationSeconds))) {
      logWarn?.('Animated image sync skipped: failed to probe word audio duration', filename);
      return 0;
    }

    totalLeadInSeconds += durationSeconds;
  }

  return totalLeadInSeconds + resolveSentenceAudioStartOffsetSeconds(config);
}
