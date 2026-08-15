export type DropFileLike = { path?: string } | { name: string };

export interface DropDataTransferLike {
  files?: ArrayLike<DropFileLike>;
  getData?: (format: string) => string;
}

const VIDEO_EXTENSIONS = new Set([
  '.3gp',
  '.avi',
  '.flv',
  '.m2ts',
  '.m4v',
  '.mkv',
  '.mov',
  '.mp4',
  '.mpeg',
  '.mpg',
  '.mts',
  '.ts',
  '.webm',
  '.wmv',
]);

const SUBTITLE_EXTENSIONS = new Set(['.ass', '.srt', '.ssa', '.sub', '.vtt']);

function getPathExtension(pathValue: string): string {
  const normalized = pathValue.split(/[?#]/, 1)[0] ?? '';
  const dot = normalized.lastIndexOf('.');
  return dot >= 0 ? normalized.slice(dot).toLowerCase() : '';
}

function isSupportedVideoPath(pathValue: string): boolean {
  return VIDEO_EXTENSIONS.has(getPathExtension(pathValue));
}

function isSupportedSubtitlePath(pathValue: string): boolean {
  return SUBTITLE_EXTENSIONS.has(getPathExtension(pathValue));
}

function parseUriList(data: string, isSupportedPath: (pathValue: string) => boolean): string[] {
  if (!data.trim()) return [];
  const out: string[] = [];

  for (const line of data.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    if (!trimmed.toLowerCase().startsWith('file://')) continue;

    try {
      const parsed = new URL(trimmed);
      let filePath = decodeURIComponent(parsed.pathname);
      if (/^\/[A-Za-z]:\//.test(filePath)) {
        filePath = filePath.slice(1);
      }
      if (filePath && isSupportedPath(filePath)) {
        out.push(filePath);
      }
    } catch {
      continue;
    }
  }

  return out;
}

export function parseClipboardVideoPath(text: string): string | null {
  const trimmed = text.trim();
  if (!trimmed) return null;

  const unquoted =
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
      ? trimmed.slice(1, -1).trim()
      : trimmed;
  if (!unquoted) return null;

  if (unquoted.toLowerCase().startsWith('file://')) {
    try {
      const parsed = new URL(unquoted);
      let filePath = decodeURIComponent(parsed.pathname);
      if (/^\/[A-Za-z]:\//.test(filePath)) {
        filePath = filePath.slice(1);
      }
      return filePath && isSupportedVideoPath(filePath) ? filePath : null;
    } catch {
      return null;
    }
  }

  return isSupportedVideoPath(unquoted) ? unquoted : null;
}

export function collectDroppedVideoPaths(
  dataTransfer: DropDataTransferLike | null | undefined,
  resolvedFilePaths: ArrayLike<string> = [],
): string[] {
  return collectDroppedPaths(dataTransfer, resolvedFilePaths, isSupportedVideoPath);
}

export function collectDroppedSubtitlePaths(
  dataTransfer: DropDataTransferLike | null | undefined,
  resolvedFilePaths: ArrayLike<string> = [],
): string[] {
  return collectDroppedPaths(dataTransfer, resolvedFilePaths, isSupportedSubtitlePath);
}

function collectDroppedPaths(
  dataTransfer: DropDataTransferLike | null | undefined,
  resolvedFilePaths: ArrayLike<string>,
  isSupportedPath: (pathValue: string) => boolean,
): string[] {
  if (!dataTransfer) return [];

  const out: string[] = [];
  const seen = new Set<string>();

  const addPath = (candidate: string | null | undefined): void => {
    if (!candidate) return;
    const trimmed = candidate.trim();
    if (!trimmed || !isSupportedPath(trimmed) || seen.has(trimmed)) return;
    seen.add(trimmed);
    out.push(trimmed);
  };

  if (dataTransfer.files) {
    for (let i = 0; i < dataTransfer.files.length; i += 1) {
      const file = dataTransfer.files[i] as { path?: string } | undefined;
      addPath(file?.path);
      addPath(resolvedFilePaths[i]);
    }
  }

  if (typeof dataTransfer.getData === 'function') {
    for (const pathValue of parseUriList(dataTransfer.getData('text/uri-list'), isSupportedPath)) {
      addPath(pathValue);
    }
  }

  return out;
}

export function buildMpvLoadfileCommands(
  paths: string[],
  append: boolean,
): Array<(string | number)[]> {
  if (append) {
    return paths.map((pathValue) => ['loadfile', pathValue, 'append']);
  }
  return paths.map((pathValue, index) => [
    'loadfile',
    pathValue,
    index === 0 ? 'replace' : 'append',
  ]);
}

export function buildMpvSubtitleAddCommands(paths: string[]): Array<(string | number)[]> {
  return paths.map((pathValue, index) =>
    index === 0 ? ['sub-add', pathValue, 'select'] : ['sub-add', pathValue],
  );
}
