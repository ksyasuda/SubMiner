import * as fs from 'fs';
import * as path from 'path';

/**
 * A GUI launch inherits a minimal PATH that omits the package-manager prefixes
 * users actually install these tools under, so PATH alone cannot find them.
 * Probing the usual prefixes is what makes "leave the config empty and we will
 * discover it" true rather than aspirational.
 */
const FALLBACK_BIN_DIRS = [
  '/opt/homebrew/bin',
  '/usr/local/bin',
  '/opt/local/bin',
  '/usr/bin',
  '/bin',
];

/** Names each tool ships under, in the order they should be preferred. */
export const SUBSYNC_EXECUTABLE_NAMES = {
  alass: ['alass', 'alass-cli'],
  ffsubsync: ['ffsubsync'],
  ffmpeg: ['ffmpeg'],
} as const;

export type SubsyncExecutable = keyof typeof SUBSYNC_EXECUTABLE_NAMES;

function unique(values: string[]): string[] {
  return values.filter((value, index) => value.length > 0 && values.indexOf(value) === index);
}

/**
 * A same-named non-executable file earlier on PATH would otherwise shadow the
 * real binary and surface as a spawn EACCES rather than "keep looking". Windows
 * has no execute bit, so the regular-file check is all it can offer there.
 */
export function isExecutableFile(filePath: string): boolean {
  try {
    if (!fs.statSync(filePath).isFile()) return false;
    if (process.platform === 'win32') return true;
    fs.accessSync(filePath, fs.constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

function searchDirectories(): string[] {
  const entries = (process.env.PATH ?? '')
    .split(path.delimiter)
    .map((entry) => entry.trim())
    .filter(Boolean);
  return unique([...entries, ...FALLBACK_BIN_DIRS]);
}

function executableNames(name: string): string[] {
  if (process.platform !== 'win32') return [name];
  const extensions = (process.env.PATHEXT ?? '.EXE;.CMD;.BAT')
    .split(';')
    .map((entry) => entry.trim())
    .filter(Boolean);
  if (path.extname(name)) return [name];
  return [name, ...extensions.map((extension) => `${name}${extension}`)];
}

/**
 * Resolve the first of `names` that exists, returning '' when none do.
 *
 * A name carrying a directory component is taken literally — the caller spelled
 * out a path, so falling back to a same-named binary elsewhere would silently
 * run something they did not point at.
 */
export function findExecutable(names: readonly string[]): string {
  for (const name of names) {
    if (path.dirname(name) !== '.') {
      return isExecutableFile(name) ? name : '';
    }
  }

  for (const dir of searchDirectories()) {
    for (const name of names) {
      for (const executableName of executableNames(name)) {
        const candidate = path.join(dir, executableName);
        if (isExecutableFile(candidate)) return candidate;
      }
    }
  }

  return '';
}

/** Resolve a configured tool path, discovering it when the config is empty. */
export function resolveExecutable(
  configuredPath: string | null | undefined,
  names: readonly string[],
): string {
  const trimmed = configuredPath?.trim() ?? '';
  if (trimmed) return findExecutable([trimmed]);
  return findExecutable(names);
}
