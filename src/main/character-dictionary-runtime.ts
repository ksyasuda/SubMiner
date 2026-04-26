import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { hasVideoExtension } from '../shared/video-extensions';
import {
  applyCollapsibleOpenStatesToTermEntries,
  buildDictionaryTitle,
  buildDictionaryZip,
  buildSnapshotFromCharacters,
  buildSnapshotImagePath,
  buildVaImagePath,
} from './character-dictionary-runtime/build';
import {
  buildMergedRevision,
  getMergedZipPath,
  getSnapshotPath,
  normalizeMergedMediaIds,
  readSnapshot,
  writeSnapshot,
} from './character-dictionary-runtime/cache';
import {
  ANILIST_REQUEST_DELAY_MS,
  CHARACTER_DICTIONARY_MERGED_TITLE,
  CHARACTER_IMAGE_DOWNLOAD_DELAY_MS,
} from './character-dictionary-runtime/constants';
import {
  downloadCharacterImage,
  fetchAniListMediaCandidateById,
  fetchCharactersForMedia,
  resolveAniListMediaIdFromGuess,
  searchAniListMediaCandidates,
} from './character-dictionary-runtime/fetch';
import {
  buildCharacterDictionarySeriesKey,
  createCharacterDictionaryManualSelectionStore,
} from './character-dictionary-runtime/manual-selection';
import type {
  AniListMediaCandidate,
  CharacterDictionaryBuildResult,
  CharacterDictionaryGenerateOptions,
  CharacterDictionaryManualSelectionResult,
  CharacterDictionaryManualSelectionSnapshot,
  CharacterDictionaryRuntimeDeps,
  CharacterDictionarySnapshotImage,
  CharacterDictionarySnapshotProgress,
  CharacterDictionarySnapshotProgressCallbacks,
  CharacterDictionarySnapshotResult,
  MergedCharacterDictionaryBuildResult,
  ResolvedAniListMedia,
} from './character-dictionary-runtime/types';

export type {
  CharacterDictionaryBuildResult,
  CharacterDictionaryGenerateOptions,
  CharacterDictionaryRuntimeDeps,
  CharacterDictionarySnapshot,
  CharacterDictionarySnapshotProgress,
  CharacterDictionarySnapshotProgressCallbacks,
  CharacterDictionarySnapshotResult,
  MergedCharacterDictionaryBuildResult,
} from './character-dictionary-runtime/types';

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function expandUserPath(input: string): string {
  if (input.startsWith('~')) {
    return path.join(os.homedir(), input.slice(1));
  }
  return input;
}

function isVideoFile(filePath: string): boolean {
  return hasVideoExtension(path.extname(filePath));
}

function findFirstVideoFileInDirectory(directoryPath: string): string | null {
  const queue: string[] = [directoryPath];
  while (queue.length > 0) {
    const current = queue.shift()!;
    let entries: fs.Dirent[] = [];
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      continue;
    }
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const entry of entries) {
      const fullPath = path.join(current, entry.name);
      if (entry.isFile() && isVideoFile(fullPath)) {
        return fullPath;
      }
      if (entry.isDirectory() && !entry.name.startsWith('.')) {
        queue.push(fullPath);
      }
    }
  }
  return null;
}

function resolveDictionaryGuessInputs(targetPath: string): {
  mediaPath: string;
  mediaTitle: string | null;
} {
  const trimmed = targetPath.trim();
  if (!trimmed) {
    throw new Error('Dictionary target path is empty.');
  }
  const resolvedPath = path.resolve(expandUserPath(trimmed));
  let stats: fs.Stats;
  try {
    stats = fs.statSync(resolvedPath);
  } catch {
    throw new Error(`Dictionary target path not found: ${targetPath}`);
  }

  if (stats.isFile()) {
    return {
      mediaPath: resolvedPath,
      mediaTitle: path.basename(resolvedPath),
    };
  }

  if (stats.isDirectory()) {
    const firstVideo = findFirstVideoFileInDirectory(resolvedPath);
    if (firstVideo) {
      return {
        mediaPath: firstVideo,
        mediaTitle: path.basename(firstVideo),
      };
    }
    return {
      mediaPath: resolvedPath,
      mediaTitle: path.basename(resolvedPath),
    };
  }

  throw new Error(`Dictionary target must be a file or directory path: ${targetPath}`);
}

export function createCharacterDictionaryRuntimeService(deps: CharacterDictionaryRuntimeDeps): {
  getOrCreateCurrentSnapshot: (
    targetPath?: string,
    progress?: CharacterDictionarySnapshotProgressCallbacks,
  ) => Promise<CharacterDictionarySnapshotResult>;
  buildMergedDictionary: (mediaIds: number[]) => Promise<MergedCharacterDictionaryBuildResult>;
  getManualSelectionSnapshot: (
    targetPath?: string,
  ) => Promise<CharacterDictionaryManualSelectionSnapshot>;
  setManualSelection: (request: {
    targetPath?: string;
    mediaId: number;
  }) => Promise<CharacterDictionaryManualSelectionResult>;
  generateForCurrentMedia: (
    targetPath?: string,
    options?: CharacterDictionaryGenerateOptions,
  ) => Promise<CharacterDictionaryBuildResult>;
} {
  const outputDir = path.join(deps.userDataPath, 'character-dictionaries');
  const sleepMs = deps.sleep ?? sleep;
  const getCollapsibleSectionOpenState = deps.getCollapsibleSectionOpenState ?? (() => false);
  const manualSelectionStore = createCharacterDictionaryManualSelectionStore({
    userDataPath: deps.userDataPath,
  });

  const createAniListRequestSlot = (): (() => Promise<void>) => {
    let hasAniListRequest = false;
    return async (): Promise<void> => {
      if (!hasAniListRequest) {
        hasAniListRequest = true;
        return;
      }
      await sleepMs(ANILIST_REQUEST_DELAY_MS);
    };
  };

  const resolveGuessInput = (
    targetPath?: string,
  ): { mediaPath: string | null; mediaTitle: string | null } => {
    const dictionaryTarget = targetPath?.trim() || '';
    return dictionaryTarget.length > 0
      ? resolveDictionaryGuessInputs(dictionaryTarget)
      : {
          mediaPath: deps.getCurrentMediaPath(),
          mediaTitle: deps.getCurrentMediaTitle(),
        };
  };

  const guessCurrentMedia = async (targetPath?: string) => {
    const guessInput = resolveGuessInput(targetPath);
    const mediaPathForGuess = deps.resolveMediaPathForJimaku(guessInput.mediaPath);
    const guessed = await deps.guessAnilistMediaInfo(mediaPathForGuess, guessInput.mediaTitle);
    if (!guessed || !guessed.title.trim()) {
      throw new Error('Unable to resolve current anime from media path/title.');
    }
    return {
      guessed,
      seriesKey: buildCharacterDictionarySeriesKey({
        mediaPath: mediaPathForGuess,
        mediaTitle: guessInput.mediaTitle,
        guess: guessed,
      }),
    };
  };

  const resolveCurrentMedia = async (
    targetPath?: string,
    beforeRequest?: () => Promise<void>,
  ): Promise<ResolvedAniListMedia> => {
    deps.logInfo?.('[dictionary] resolving current anime for character dictionary generation');
    const { guessed, seriesKey } = await guessCurrentMedia(targetPath);
    deps.logInfo?.(
      `[dictionary] current anime guess: ${guessed.title.trim()}${
        typeof guessed.episode === 'number' && guessed.episode > 0
          ? ` (episode ${guessed.episode})`
          : ''
      }`,
    );
    const override = await manualSelectionStore.getOverride(seriesKey);
    if (override) {
      deps.logInfo?.(
        `[dictionary] manual AniList override: ${override.mediaTitle} -> AniList ${override.mediaId}`,
      );
      return {
        id: override.mediaId,
        title: override.mediaTitle,
        staleMediaIds: override.staleMediaIds,
      };
    }
    const resolved = await resolveAniListMediaIdFromGuess(guessed, beforeRequest);
    deps.logInfo?.(`[dictionary] AniList match: ${resolved.title} -> AniList ${resolved.id}`);
    return resolved;
  };

  const getOrCreateSnapshot = async (
    mediaId: number,
    mediaTitleHint?: string,
    beforeRequest?: () => Promise<void>,
    progress?: CharacterDictionarySnapshotProgressCallbacks,
  ): Promise<CharacterDictionarySnapshotResult> => {
    const snapshotPath = getSnapshotPath(outputDir, mediaId);
    const cachedSnapshot = readSnapshot(snapshotPath);
    if (cachedSnapshot) {
      deps.logInfo?.(`[dictionary] snapshot hit for AniList ${mediaId}`);
      return {
        mediaId: cachedSnapshot.mediaId,
        mediaTitle: cachedSnapshot.mediaTitle,
        entryCount: cachedSnapshot.entryCount,
        fromCache: true,
        updatedAt: cachedSnapshot.updatedAt,
      };
    }

    progress?.onGenerating?.({
      mediaId,
      mediaTitle: mediaTitleHint || `AniList ${mediaId}`,
    });
    deps.logInfo?.(`[dictionary] snapshot miss for AniList ${mediaId}, fetching characters`);

    const { mediaTitle: fetchedMediaTitle, characters } = await fetchCharactersForMedia(
      mediaId,
      beforeRequest,
      (page) => {
        deps.logInfo?.(
          `[dictionary] downloaded AniList character page ${page} for AniList ${mediaId}`,
        );
      },
    );
    if (characters.length === 0) {
      throw new Error(`No characters returned for AniList media ${mediaId}.`);
    }

    const imagesByCharacterId = new Map<number, CharacterDictionarySnapshotImage>();
    const imagesByVaId = new Map<number, CharacterDictionarySnapshotImage>();
    const allImageUrls: Array<{ id: number; url: string; kind: 'character' | 'va' }> = [];
    const seenVaIds = new Set<number>();
    for (const character of characters) {
      if (character.imageUrl) {
        allImageUrls.push({ id: character.id, url: character.imageUrl, kind: 'character' });
      }
      for (const va of character.voiceActors) {
        if (va.imageUrl && !seenVaIds.has(va.id)) {
          seenVaIds.add(va.id);
          allImageUrls.push({ id: va.id, url: va.imageUrl, kind: 'va' });
        }
      }
    }
    if (allImageUrls.length > 0) {
      deps.logInfo?.(
        `[dictionary] downloading ${allImageUrls.length} images for AniList ${mediaId}`,
      );
    }
    let hasAttemptedImageDownload = false;
    for (const entry of allImageUrls) {
      if (hasAttemptedImageDownload) {
        await sleepMs(CHARACTER_IMAGE_DOWNLOAD_DELAY_MS);
      }
      hasAttemptedImageDownload = true;
      const image = await downloadCharacterImage(entry.url, entry.id);
      if (!image) continue;
      if (entry.kind === 'character') {
        imagesByCharacterId.set(entry.id, {
          path: buildSnapshotImagePath(mediaId, entry.id, image.ext),
          dataBase64: image.bytes.toString('base64'),
        });
      } else {
        imagesByVaId.set(entry.id, {
          path: buildVaImagePath(mediaId, entry.id, image.ext),
          dataBase64: image.bytes.toString('base64'),
        });
      }
    }

    const snapshot = buildSnapshotFromCharacters(
      mediaId,
      fetchedMediaTitle || mediaTitleHint || `AniList ${mediaId}`,
      characters,
      imagesByCharacterId,
      imagesByVaId,
      deps.now(),
      getCollapsibleSectionOpenState,
    );
    writeSnapshot(snapshotPath, snapshot);
    deps.logInfo?.(
      `[dictionary] stored snapshot for AniList ${mediaId}: ${snapshot.entryCount} terms`,
    );

    return {
      mediaId: snapshot.mediaId,
      mediaTitle: snapshot.mediaTitle,
      entryCount: snapshot.entryCount,
      fromCache: false,
      updatedAt: snapshot.updatedAt,
    };
  };

  return {
    getOrCreateCurrentSnapshot: async (
      targetPath?: string,
      progress?: CharacterDictionarySnapshotProgressCallbacks,
    ) => {
      const waitForAniListRequestSlot = createAniListRequestSlot();
      const resolvedMedia = await resolveCurrentMedia(targetPath, waitForAniListRequestSlot);
      progress?.onChecking?.({
        mediaId: resolvedMedia.id,
        mediaTitle: resolvedMedia.title,
      });
      const snapshot = await getOrCreateSnapshot(
        resolvedMedia.id,
        resolvedMedia.title,
        waitForAniListRequestSlot,
        progress,
      );
      return {
        ...snapshot,
        staleMediaIds: resolvedMedia.staleMediaIds,
      };
    },
    buildMergedDictionary: async (mediaIds: number[]) => {
      const normalizedMediaIds = normalizeMergedMediaIds(mediaIds);
      const snapshotResults = await Promise.all(
        normalizedMediaIds.map((mediaId) => getOrCreateSnapshot(mediaId)),
      );
      const snapshots = snapshotResults.map(({ mediaId }) => {
        const snapshot = readSnapshot(getSnapshotPath(outputDir, mediaId));
        if (!snapshot) {
          throw new Error(`Missing character dictionary snapshot for AniList ${mediaId}.`);
        }
        return snapshot;
      });
      const revision = buildMergedRevision(normalizedMediaIds, snapshots);
      const description =
        snapshots.length === 1
          ? `Character names from ${snapshots[0]!.mediaTitle}`
          : `Character names from ${snapshots.length} recent anime`;
      const { zipPath, entryCount } = buildDictionaryZip(
        getMergedZipPath(outputDir),
        CHARACTER_DICTIONARY_MERGED_TITLE,
        description,
        revision,
        applyCollapsibleOpenStatesToTermEntries(
          snapshots.flatMap((snapshot) => snapshot.termEntries),
          getCollapsibleSectionOpenState,
        ),
        snapshots.flatMap((snapshot) => snapshot.images),
      );
      deps.logInfo?.(
        `[dictionary] rebuilt merged dictionary: ${normalizedMediaIds.join(', ') || '<empty>'} -> ${zipPath}`,
      );
      return {
        zipPath,
        revision,
        dictionaryTitle: CHARACTER_DICTIONARY_MERGED_TITLE,
        entryCount,
      };
    },
    getManualSelectionSnapshot: async (targetPath?: string) => {
      const waitForAniListRequestSlot = createAniListRequestSlot();
      const { guessed, seriesKey } = await guessCurrentMedia(targetPath);
      const [candidates, override] = await Promise.all([
        searchAniListMediaCandidates(guessed.title, waitForAniListRequestSlot),
        manualSelectionStore.getOverride(seriesKey),
      ]);
      const current = await resolveAniListMediaIdFromGuess(guessed, waitForAniListRequestSlot)
        .then(
          (entry): AniListMediaCandidate => ({
            id: entry.id,
            title: entry.title,
            episodes: candidates.find((candidate) => candidate.id === entry.id)?.episodes ?? null,
          }),
        )
        .catch(() => null);
      return {
        seriesKey,
        guessTitle: guessed.title,
        current,
        override: override
          ? { id: override.mediaId, title: override.mediaTitle, episodes: null }
          : null,
        candidates,
      };
    },
    setManualSelection: async ({ targetPath, mediaId }) => {
      const waitForAniListRequestSlot = createAniListRequestSlot();
      const { guessed, seriesKey } = await guessCurrentMedia(targetPath);
      const [selected, current] = await Promise.all([
        fetchAniListMediaCandidateById(mediaId, waitForAniListRequestSlot),
        resolveAniListMediaIdFromGuess(guessed, waitForAniListRequestSlot).catch(() => null),
      ]);
      const staleMediaIds = current && current.id !== selected.id ? [current.id] : [];
      await manualSelectionStore.setOverride({
        seriesKey,
        mediaId: selected.id,
        mediaTitle: selected.title,
        staleMediaIds,
      });
      return {
        ok: true,
        seriesKey,
        selected,
        staleMediaIds,
      };
    },
    generateForCurrentMedia: async (
      targetPath?: string,
      _options?: CharacterDictionaryGenerateOptions,
    ) => {
      const waitForAniListRequestSlot = createAniListRequestSlot();
      const resolvedMedia = await resolveCurrentMedia(targetPath, waitForAniListRequestSlot);
      const snapshot = await getOrCreateSnapshot(
        resolvedMedia.id,
        resolvedMedia.title,
        waitForAniListRequestSlot,
      );
      const storedSnapshot = readSnapshot(getSnapshotPath(outputDir, resolvedMedia.id));
      if (!storedSnapshot) {
        throw new Error(`Snapshot missing after generation for AniList ${resolvedMedia.id}.`);
      }
      const revision = String(storedSnapshot.updatedAt);
      const dictionaryTitle = buildDictionaryTitle(resolvedMedia.id);
      const description = `Character names from ${storedSnapshot.mediaTitle} [AniList media ID ${resolvedMedia.id}]`;
      const zipPath = path.join(outputDir, `anilist-${resolvedMedia.id}.zip`);
      deps.logInfo?.(`[dictionary] building ZIP for AniList ${resolvedMedia.id}`);
      buildDictionaryZip(
        zipPath,
        dictionaryTitle,
        description,
        revision,
        applyCollapsibleOpenStatesToTermEntries(
          storedSnapshot.termEntries,
          getCollapsibleSectionOpenState,
        ),
        storedSnapshot.images,
      );
      deps.logInfo?.(
        `[dictionary] generated AniList ${resolvedMedia.id}: ${storedSnapshot.entryCount} terms -> ${zipPath}`,
      );
      return {
        zipPath,
        fromCache: snapshot.fromCache,
        mediaId: resolvedMedia.id,
        mediaTitle: storedSnapshot.mediaTitle,
        entryCount: storedSnapshot.entryCount,
        dictionaryTitle,
        revision,
      };
    },
  };
}
