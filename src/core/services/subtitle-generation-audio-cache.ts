import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rename,
  rm,
  stat,
  utimes,
  writeFile,
} from 'node:fs/promises';
import path from 'node:path';
import type { ResolvedMpvHttpHeaders } from './mpv-http-headers';

const MAX_BYTES = 512 * 1024 * 1024;
const ENTRY_NAME = /^[a-f0-9]{64}$/;

export function subtitleAudioCacheKey(
  mediaPath: string,
  index: number | undefined,
  headers: ResolvedMpvHttpHeaders,
): string {
  return createHash('sha256')
    .update(
      JSON.stringify([
        1,
        mediaPath,
        index ?? null,
        headers.userAgent,
        Object.entries(headers.headers).sort(([a], [b]) => a.localeCompare(b)),
      ]),
    )
    .digest('hex');
}

async function digest(file: string): Promise<string> {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}

/** Copy into the job's directory so eviction by another job cannot remove active audio. */
export async function readSubtitleAudioCache(input: {
  directory: string;
  key: string;
  destination: string;
}): Promise<{ offset: number } | null> {
  const entry = path.join(input.directory, 'audio', input.key);
  try {
    const metadata: unknown = JSON.parse(await readFile(path.join(entry, 'metadata.json'), 'utf8'));
    if (
      typeof metadata !== 'object' ||
      metadata === null ||
      !('offset' in metadata) ||
      typeof metadata.offset !== 'number' ||
      !Number.isFinite(metadata.offset) ||
      !('size' in metadata) ||
      typeof metadata.size !== 'number' ||
      metadata.size <= 44 ||
      metadata.size > MAX_BYTES ||
      !('sha256' in metadata) ||
      typeof metadata.sha256 !== 'string'
    )
      return null;
    const source = path.join(entry, 'audio.wav');
    if ((await stat(source)).size !== metadata.size) return null;
    await copyFile(source, input.destination);
    if (
      (await stat(input.destination)).size !== metadata.size ||
      (await digest(input.destination)) !== metadata.sha256
    ) {
      await rm(input.destination, { force: true });
      return null;
    }
    await utimes(entry, new Date(), new Date());
    return { offset: metadata.offset };
  } catch {
    await rm(input.destination, { force: true }).catch(() => undefined);
    return null;
  }
}

async function prune(directory: string): Promise<void> {
  const entries = [];
  for (const name of await readdir(directory)) {
    if (!ENTRY_NAME.test(name)) continue;
    const entry = path.join(directory, name);
    try {
      entries.push({
        path: entry,
        modified: (await stat(entry)).mtimeMs,
        size: (await stat(path.join(entry, 'audio.wav'))).size,
      });
    } catch {
      /* A concurrent eviction may already have removed it. */
    }
  }
  let retained = 0;
  for (const entry of entries.sort((a, b) => b.modified - a.modified)) {
    if (retained + entry.size > MAX_BYTES) {
      await rm(entry.path, { recursive: true, force: true });
    } else retained += entry.size;
  }
}

/** Publish only complete extracts. Cache failures must never discard generated subtitles. */
export async function writeSubtitleAudioCache(input: {
  directory: string;
  key: string;
  wavPath: string;
  offset: number;
}): Promise<void> {
  const size = (await stat(input.wavPath)).size;
  if (size <= 44 || size > MAX_BYTES) return;
  const directory = path.join(input.directory, 'audio');
  await mkdir(directory, { recursive: true });
  const staged = await mkdtemp(path.join(directory, '.pending-'));
  try {
    await copyFile(input.wavPath, path.join(staged, 'audio.wav'));
    await writeFile(
      path.join(staged, 'metadata.json'),
      JSON.stringify({ size, offset: input.offset, sha256: await digest(input.wavPath) }),
    );
    const destination = path.join(directory, input.key);
    // Replace a stale or corrupt entry only after its replacement is complete.
    await rm(destination, { recursive: true, force: true });
    await rename(staged, destination);
    await prune(directory);
  } finally {
    await rm(staged, { recursive: true, force: true });
  }
}
