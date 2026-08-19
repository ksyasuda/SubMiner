import * as fs from 'fs';
import * as path from 'path';
import type { CharacterNameImage } from '../../types';
import { readCachedSnapshots } from './cache';
import type {
  CharacterDictionaryGlossaryEntry,
  CharacterDictionarySnapshot,
  CharacterDictionarySnapshotImage,
  CharacterDictionaryTermEntry,
} from './types';

const CHARACTER_IMAGE_PATH_PATTERN = /^img\/m\d+-c\d+\.[a-z0-9]+$/i;

type StructuredContentNode = {
  tag?: unknown;
  path?: unknown;
  alt?: unknown;
  title?: unknown;
  content?: unknown;
};

function normalizeLookupTerm(term: string): string {
  return term.trim();
}

function normalizeLookupMediaId(mediaId: unknown): number | null {
  if (typeof mediaId !== 'number' || !Number.isFinite(mediaId)) {
    return null;
  }
  const normalized = Math.floor(mediaId);
  return normalized > 0 ? normalized : null;
}

function getSnapshotsDir(outputDir: string): string {
  return path.join(outputDir, 'snapshots');
}

function getImageMimeType(imagePath: string, dataBase64: string): string {
  const signature = Buffer.from(dataBase64.slice(0, 64), 'base64');
  if (
    signature.length >= 8 &&
    signature[0] === 0x89 &&
    signature[1] === 0x50 &&
    signature[2] === 0x4e &&
    signature[3] === 0x47
  ) {
    return 'image/png';
  }
  if (
    signature.length >= 12 &&
    signature.subarray(0, 4).toString('ascii') === 'RIFF' &&
    signature.subarray(8, 12).toString('ascii') === 'WEBP'
  ) {
    return 'image/webp';
  }
  if (
    signature.length >= 6 &&
    (signature.subarray(0, 6).toString('ascii') === 'GIF89a' ||
      signature.subarray(0, 6).toString('ascii') === 'GIF87a')
  ) {
    return 'image/gif';
  }
  if (signature.length >= 3 && signature[0] === 0xff && signature[1] === 0xd8) {
    return 'image/jpeg';
  }
  if (
    signature.length >= 12 &&
    signature.subarray(4, 8).toString('ascii') === 'ftyp' &&
    signature.subarray(8, 12).toString('ascii') === 'avif'
  ) {
    return 'image/avif';
  }

  const ext = path.extname(imagePath).toLowerCase();
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.png') return 'image/png';
  if (ext === '.webp') return 'image/webp';
  if (ext === '.gif') return 'image/gif';
  if (ext === '.avif') return 'image/avif';
  return 'image/jpeg';
}

function buildImageByPath(
  images: ReadonlyArray<CharacterDictionarySnapshotImage>,
): Map<string, CharacterDictionarySnapshotImage> {
  const imageByPath = new Map<string, CharacterDictionarySnapshotImage>();
  for (const image of images) {
    if (image.path && image.dataBase64) {
      imageByPath.set(image.path, image);
    }
  }
  return imageByPath;
}

function findCharacterImageNode(value: unknown): StructuredContentNode | null {
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findCharacterImageNode(item);
      if (found) return found;
    }
    return null;
  }

  if (!value || typeof value !== 'object') {
    return null;
  }

  const node = value as StructuredContentNode;
  if (
    node.tag === 'img' &&
    typeof node.path === 'string' &&
    CHARACTER_IMAGE_PATH_PATTERN.test(node.path)
  ) {
    return node;
  }

  return findCharacterImageNode(node.content);
}

function findCharacterImageNodeInGlossary(
  glossary: ReadonlyArray<CharacterDictionaryGlossaryEntry>,
): StructuredContentNode | null {
  for (const entry of glossary) {
    const found = findCharacterImageNode(entry);
    if (found) return found;
  }
  return null;
}

function createCharacterNameImage(
  entry: CharacterDictionaryTermEntry,
  imageByPath: ReadonlyMap<string, CharacterDictionarySnapshotImage>,
): CharacterNameImage | null {
  const term = normalizeLookupTerm(entry[0]);
  if (!term) {
    return null;
  }

  const imageNode = findCharacterImageNodeInGlossary(entry[5]);
  const imagePath = typeof imageNode?.path === 'string' ? imageNode.path : '';
  const image = imageByPath.get(imagePath);
  if (!image) {
    return null;
  }

  const rawAlt =
    typeof imageNode?.alt === 'string'
      ? imageNode.alt
      : typeof imageNode?.title === 'string'
        ? imageNode.title
        : term;
  const alt = rawAlt.trim() || term;
  return {
    src: `data:${getImageMimeType(image.path, image.dataBase64)};base64,${image.dataBase64}`,
    alt,
  };
}

function appendSnapshotImages(
  index: Map<string, CharacterNameImage>,
  snapshot: CharacterDictionarySnapshot,
): void {
  const imageByPath = buildImageByPath(snapshot.images);
  for (const entry of snapshot.termEntries) {
    const term = normalizeLookupTerm(entry[0]);
    if (!term || index.has(term)) {
      continue;
    }
    const image = createCharacterNameImage(entry, imageByPath);
    if (image) {
      index.set(term, image);
    }
  }
}

export function snapshotHasCharacterNameImages(snapshot: CharacterDictionarySnapshot): boolean {
  const imageByPath = buildImageByPath(snapshot.images);
  return snapshot.termEntries.some(
    (entry) => createCharacterNameImage(entry, imageByPath) !== null,
  );
}

function getSnapshotDirectorySignature(outputDir: string): string {
  let entries: fs.Dirent[] = [];
  try {
    entries = fs.readdirSync(getSnapshotsDir(outputDir), { withFileTypes: true });
  } catch {
    return '';
  }

  const parts: string[] = [];
  for (const entry of entries) {
    if (!entry.isFile() || !/^anilist-\d+\.json$/.test(entry.name)) {
      continue;
    }
    const snapshotPath = path.join(getSnapshotsDir(outputDir), entry.name);
    try {
      const stat = fs.statSync(snapshotPath);
      parts.push(`${entry.name}:${stat.mtimeMs}:${stat.size}`);
    } catch {
      // Ignore files that disappear during refresh; next lookup will rebuild.
    }
  }
  return parts.sort().join('|');
}

export async function buildCharacterNameImageIndexFromSnapshots(
  outputDir: string,
): Promise<Map<string, CharacterNameImage>> {
  const index = new Map<string, CharacterNameImage>();
  for (const snapshot of await readCachedSnapshots(outputDir)) {
    appendSnapshotImages(index, snapshot);
  }
  return index;
}

export function createCharacterDictionaryImageLookup(deps: {
  userDataPath?: string;
  outputDir?: string;
  getCurrentMediaId?: () => number | null | undefined;
  onIndexReady?: () => void;
  onIndexReadyError?: (error: unknown) => void;
}): {
  get: (term: string, mediaId?: number | null) => CharacterNameImage | null;
  invalidate: () => void;
} {
  const outputDir =
    deps.outputDir ??
    (deps.userDataPath ? path.join(deps.userDataPath, 'character-dictionaries') : '');
  let signature: string | null = null;
  let index = new Map<string, CharacterNameImage>();
  let indexByMediaId = new Map<number, Map<string, CharacterNameImage>>();
  let refreshInFlight = false;
  let indexReadyDeliveryPending = false;

  function deliverIndexReadyIfPending(): void {
    if (!indexReadyDeliveryPending || !deps.onIndexReady) {
      return;
    }
    indexReadyDeliveryPending = false;
    try {
      deps.onIndexReady();
    } catch (error) {
      indexReadyDeliveryPending = true;
      try {
        deps.onIndexReadyError?.(error);
      } catch {
        // Error reporting must not reject the detached index refresh task.
      }
    }
  }

  // Rebuilding means re-reading every cached snapshot (potentially GBs of JSON), which used to run
  // synchronously inside a lookup and froze the whole app right after a snapshot changed. Lookups
  // now serve the previous index while a single background rebuild catches up; the swap is atomic
  // and the signature only advances once the rebuild it belongs to has landed.
  function refreshIfNeeded(): void {
    if (!outputDir) {
      index = new Map<string, CharacterNameImage>();
      indexByMediaId = new Map<number, Map<string, CharacterNameImage>>();
      signature = '';
      return;
    }
    deliverIndexReadyIfPending();
    const nextSignature = getSnapshotDirectorySignature(outputDir);
    if (nextSignature === signature || refreshInFlight) {
      return;
    }
    refreshInFlight = true;
    void (async () => {
      try {
        const snapshots = await readCachedSnapshots(outputDir);
        const nextIndex = new Map<string, CharacterNameImage>();
        const nextIndexByMediaId = new Map<number, Map<string, CharacterNameImage>>();
        for (const snapshot of snapshots) {
          appendSnapshotImages(nextIndex, snapshot);
          const mediaIndex = new Map<string, CharacterNameImage>();
          appendSnapshotImages(mediaIndex, snapshot);
          if (mediaIndex.size > 0) {
            nextIndexByMediaId.set(snapshot.mediaId, mediaIndex);
          }
        }
        index = nextIndex;
        indexByMediaId = nextIndexByMediaId;
        signature = nextSignature;
        indexReadyDeliveryPending = deps.onIndexReady !== undefined;
        deliverIndexReadyIfPending();
      } finally {
        refreshInFlight = false;
      }
    })();
  }

  return {
    get(term: string, mediaId?: number | null): CharacterNameImage | null {
      const normalizedTerm = normalizeLookupTerm(term);
      if (!normalizedTerm) {
        return null;
      }
      refreshIfNeeded();
      const scopedMediaId = normalizeLookupMediaId(mediaId ?? deps.getCurrentMediaId?.() ?? null);
      if (scopedMediaId !== null) {
        return indexByMediaId.get(scopedMediaId)?.get(normalizedTerm) ?? null;
      }
      return index.get(normalizedTerm) ?? null;
    },
    invalidate(): void {
      signature = null;
    },
  };
}
