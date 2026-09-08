import { constants } from 'node:fs';
import { copyFile, link } from 'node:fs/promises';
import { homedir } from 'node:os';
import path from 'node:path';

export function expandSubtitleGenerationPath(value: string): string {
  if (value === '~') return homedir();
  if (value.startsWith('~/') || value.startsWith('~\\'))
    return path.join(homedir(), value.slice(2));
  return value;
}

// Prefer atomic publication. Filesystems without hard links still get exclusive creation.
export async function publishSubtitleGenerationFile(
  source: string,
  destination: string,
): Promise<void> {
  try {
    await link(source, destination);
  } catch (error) {
    if (
      !(error instanceof Error) ||
      !('code' in error) ||
      (error.code !== 'ENOTSUP' &&
        error.code !== 'EOPNOTSUPP' &&
        error.code !== 'EPERM' &&
        error.code !== 'EXDEV')
    )
      throw error;
    await copyFile(source, destination, constants.COPYFILE_EXCL);
  }
}
