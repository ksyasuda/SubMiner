import * as path from 'node:path';

const DEFAULT_JIMAKU_LANGUAGE_SUFFIX = 'ja';
const DEFAULT_SUBTITLE_EXTENSION = '.srt';

function stripFileExtension(name: string): string {
  const ext = path.extname(name);
  return ext ? name.slice(0, -ext.length) : name;
}

function sanitizeFilenameSegment(value: string, fallback: string): string {
  const sanitized = value
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return sanitized || fallback;
}

function resolveMediaFilename(mediaPath: string): string {
  if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(mediaPath)) {
    return path.basename(path.resolve(mediaPath));
  }

  try {
    const parsedUrl = new URL(mediaPath);
    const decodedPath = decodeURIComponent(parsedUrl.pathname);
    const fromPath = path.basename(decodedPath);
    if (fromPath) {
      return fromPath;
    }
    return parsedUrl.hostname.replace(/^www\./, '') || 'subtitle';
  } catch {
    return path.basename(mediaPath);
  }
}

export function buildJimakuSubtitleFilenameFromMediaPath(
  mediaPath: string,
  downloadedSubtitleName: string,
  languageSuffix = DEFAULT_JIMAKU_LANGUAGE_SUFFIX,
): string {
  const mediaFilename = resolveMediaFilename(mediaPath);
  const mediaBasename = sanitizeFilenameSegment(stripFileExtension(mediaFilename), 'subtitle');
  const subtitleName = path.basename(downloadedSubtitleName);
  const subtitleExt = path.extname(subtitleName) || DEFAULT_SUBTITLE_EXTENSION;
  const normalizedLanguageSuffix = sanitizeFilenameSegment(languageSuffix, 'ja').replace(
    /\s+/g,
    '-',
  );
  return `${mediaBasename}.${normalizedLanguageSuffix}${subtitleExt}`;
}
