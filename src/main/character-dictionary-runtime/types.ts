import type { AnilistMediaGuess } from '../../core/services/anilist/anilist-updater';
import type { AnilistCharacterDictionaryCollapsibleSectionKey } from '../../types';

export type CharacterDictionaryRole = 'main' | 'primary' | 'side' | 'appears';

export type CharacterDictionaryGlossaryEntry = string | Record<string, unknown>;

export type CharacterDictionaryTermEntry = [
  string,
  string,
  string,
  string,
  number,
  CharacterDictionaryGlossaryEntry[],
  number,
  string,
];

export type CharacterDictionarySnapshotImage = {
  path: string;
  dataBase64: string;
};

export type CharacterBirthday = [number, number];

export type JapaneseNameParts = {
  hasSpace: boolean;
  original: string;
  combined: string;
  family: string | null;
  given: string | null;
};

export type ResolvedNameSplit = {
  family: string;
  given: string;
};

export type ResolvedNameSplits = ReadonlyMap<string, ResolvedNameSplit>;

export type NameSplitToken = {
  word: string;
  pos1?: string;
  pos2?: string;
  pos3?: string;
  pos4?: string;
  katakanaReading?: string;
};

export type NameSplitTokenizer = (text: string) => Promise<NameSplitToken[] | null>;

export type NameSplitSource = 'mecab' | 'heuristic';

export type NameReadings = {
  hasSpace: boolean;
  original: string;
  full: string;
  family: string;
  given: string;
};

export type CharacterDictionarySnapshot = {
  formatVersion: number;
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  updatedAt: number;
  nameSplitSource?: NameSplitSource;
  termEntries: CharacterDictionaryTermEntry[];
  images: CharacterDictionarySnapshotImage[];
};

export type VoiceActorRecord = {
  id: number;
  fullName: string;
  nativeName: string;
  imageUrl: string | null;
};

export type CharacterRecord = {
  id: number;
  role: CharacterDictionaryRole;
  firstNameHint: string;
  fullName: string;
  lastNameHint: string;
  nativeName: string;
  alternativeNames: string[];
  bloodType: string;
  birthday: CharacterBirthday | null;
  description: string;
  imageUrl: string | null;
  age: string;
  sex: string;
  voiceActors: VoiceActorRecord[];
};

export type CharacterDictionaryBuildResult = {
  zipPath: string;
  fromCache: boolean;
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  dictionaryTitle?: string;
  revision?: string;
};

export type CharacterDictionaryGenerateOptions = {
  refreshTtlMs?: number;
};

export type CharacterDictionarySnapshotResult = {
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  fromCache: boolean;
  updatedAt: number;
  staleMediaIds?: number[];
};

export type CharacterDictionarySnapshotProgress = {
  mediaId: number;
  mediaTitle: string;
};

export type CharacterDictionarySnapshotProgressCallbacks = {
  onChecking?: (progress: CharacterDictionarySnapshotProgress) => void;
  onGenerating?: (progress: CharacterDictionarySnapshotProgress) => void;
};

export type MergedCharacterDictionaryBuildResult = {
  zipPath: string;
  revision: string;
  dictionaryTitle: string;
  entryCount: number;
};

export type AniListMediaCandidate = {
  id: number;
  title: string;
  episodes: number | null;
};

export type CharacterDictionaryManualSelectionSnapshot = {
  seriesKey: string;
  guessTitle: string | null;
  current: AniListMediaCandidate | null;
  override: AniListMediaCandidate | null;
  candidates: AniListMediaCandidate[];
};

export type CharacterDictionaryManualSelectionResult = {
  ok: boolean;
  seriesKey: string;
  selected: AniListMediaCandidate;
  staleMediaIds: number[];
};

export interface CharacterDictionaryRuntimeDeps {
  userDataPath: string;
  getCurrentMediaPath: () => string | null;
  getCurrentVideoPath?: () => string | null | undefined;
  getCurrentMediaTitle: () => string | null;
  resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
  guessAnilistMediaInfo: (
    mediaPath: string | null,
    mediaTitle: string | null,
  ) => Promise<AnilistMediaGuess | null>;
  now: () => number;
  sleep?: (ms: number) => Promise<void>;
  logInfo?: (message: string) => void;
  logWarn?: (message: string) => void;
  getNameMatchImagesEnabled?: () => boolean;
  getCollapsibleSectionOpenState?: (
    section: AnilistCharacterDictionaryCollapsibleSectionKey,
  ) => boolean;
  tokenizeJapaneseName?: NameSplitTokenizer;
  getJapaneseNameTokenizerAvailable?: () => boolean;
}

export type ResolvedAniListMedia = {
  id: number;
  title: string;
  staleMediaIds?: number[];
  /** False when a season >= 2 was requested but only the season 1 entry could be found. */
  seasonResolved?: boolean;
  requestedSeason?: number | null;
};
