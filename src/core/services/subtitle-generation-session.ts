import { mkdir, mkdtemp, readdir, rm } from 'node:fs/promises';
import path from 'node:path';

/** Owns downloaded audio and job files, never saved subtitles or speech models. */
export function createSubtitleGenerationSession(cacheDirectory: string) {
  const root = path.join(cacheDirectory, 'sessions');
  let directory: Promise<string> | undefined;
  let disposed = false;
  let disposal: Promise<void> | undefined;
  async function prepare(): Promise<string> {
    await mkdir(root, { recursive: true });
    // Remove audio retained by the previous persistent-cache implementation.
    await rm(path.join(cacheDirectory, 'audio'), { recursive: true, force: true });
    for (const entry of await readdir(root, { withFileTypes: true })) {
      const match = /^(\d+)-[a-zA-Z0-9]+$/.exec(entry.name);
      if (!entry.isDirectory() || !match) continue;
      try {
        process.kill(Number(match[1]), 0);
      } catch (error) {
        if (error instanceof Error && 'code' in error && error.code === 'ESRCH') {
          await rm(path.join(root, entry.name), { recursive: true, force: true });
        }
      }
    }
    return mkdtemp(path.join(root, `${process.pid}-`));
  }
  return {
    directory(): Promise<string> {
      if (disposed) return Promise.reject(new Error('Subtitle generation session has closed.'));
      return (directory ??= prepare());
    },
    dispose(): Promise<void> {
      disposed = true;
      return (disposal ??= (async () => {
        if (directory) await rm(await directory, { recursive: true, force: true });
      })());
    },
  };
}
