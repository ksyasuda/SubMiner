import type { AnilistCharacterDictionaryCollapsibleSectionKey } from '../../types';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import { createDefinitionGlossary } from './glossary';
import { generateNameReadings, splitJapaneseName } from './name-reading';
import {
  buildNameTerms,
  buildReadingForTerm,
  buildTermEntry,
  buildVisibleNameTerms,
} from './term-building';
import type {
  CharacterDictionaryGlossaryEntry,
  CharacterDictionarySnapshot,
  CharacterDictionarySnapshotImage,
  CharacterDictionaryTermEntry,
  CharacterRecord,
} from './types';

export function buildSnapshotImagePath(mediaId: number, charId: number, ext: string): string {
  return `img/m${mediaId}-c${charId}.${ext}`;
}

export function buildVaImagePath(mediaId: number, vaId: number, ext: string): string {
  return `img/m${mediaId}-va${vaId}.${ext}`;
}

export function buildSnapshotFromCharacters(
  mediaId: number,
  mediaTitle: string,
  characters: CharacterRecord[],
  imagesByCharacterId: Map<number, CharacterDictionarySnapshotImage>,
  imagesByVaId: Map<number, CharacterDictionarySnapshotImage>,
  updatedAt: number,
  getCollapsibleSectionOpenState: (
    section: AnilistCharacterDictionaryCollapsibleSectionKey,
  ) => boolean,
): CharacterDictionarySnapshot {
  const termEntries: CharacterDictionaryTermEntry[] = [];

  for (const character of characters) {
    const seenTerms = new Set<string>();
    const imagePath = imagesByCharacterId.get(character.id)?.path ?? null;
    const vaImagePaths = new Map<number, string>();
    for (const va of character.voiceActors) {
      const vaImg = imagesByVaId.get(va.id);
      if (vaImg) vaImagePaths.set(va.id, vaImg.path);
    }
    const candidateTerms = buildNameTerms(character);
    const glossary = createDefinitionGlossary(
      character,
      mediaId,
      mediaTitle,
      imagePath,
      vaImagePaths,
      buildVisibleNameTerms(candidateTerms),
      getCollapsibleSectionOpenState,
    );
    const nameParts = splitJapaneseName(
      character.nativeName,
      character.firstNameHint,
      character.lastNameHint,
    );
    const readings = generateNameReadings(
      character.nativeName,
      character.fullName,
      character.firstNameHint,
      character.lastNameHint,
    );
    for (const term of candidateTerms) {
      if (seenTerms.has(term)) continue;
      seenTerms.add(term);
      const reading = buildReadingForTerm(term, character, readings, nameParts);
      termEntries.push(buildTermEntry(term, reading, character.role, glossary));
    }
  }

  if (termEntries.length === 0) {
    throw new Error('No dictionary entries generated from AniList character data.');
  }

  return {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId,
    mediaTitle,
    entryCount: termEntries.length,
    updatedAt,
    termEntries,
    images: [...imagesByCharacterId.values(), ...imagesByVaId.values()],
  };
}

function getCollapsibleSectionKeyFromTitle(
  title: string,
): AnilistCharacterDictionaryCollapsibleSectionKey | null {
  if (title === 'Description') return 'description';
  if (title === 'Character Information') return 'characterInformation';
  if (title === 'Voiced by') return 'voicedBy';
  return null;
}

function applyCollapsibleOpenStatesToStructuredValue(
  value: unknown,
  getCollapsibleSectionOpenState: (
    section: AnilistCharacterDictionaryCollapsibleSectionKey,
  ) => boolean,
): unknown {
  if (Array.isArray(value)) {
    return value.map((item) =>
      applyCollapsibleOpenStatesToStructuredValue(item, getCollapsibleSectionOpenState),
    );
  }
  if (!value || typeof value !== 'object') {
    return value;
  }

  const record = value as Record<string, unknown>;
  const next: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(record)) {
    next[key] = applyCollapsibleOpenStatesToStructuredValue(child, getCollapsibleSectionOpenState);
  }

  if (record.tag === 'details') {
    const content = Array.isArray(record.content) ? record.content : [];
    const summary = content[0];
    if (summary && typeof summary === 'object' && !Array.isArray(summary)) {
      const summaryContent = (summary as Record<string, unknown>).content;
      if (typeof summaryContent === 'string') {
        const section = getCollapsibleSectionKeyFromTitle(summaryContent);
        if (section) {
          next.open = getCollapsibleSectionOpenState(section);
        }
      }
    }
  }

  return next;
}

export function applyCollapsibleOpenStatesToTermEntries(
  termEntries: CharacterDictionaryTermEntry[],
  getCollapsibleSectionOpenState: (
    section: AnilistCharacterDictionaryCollapsibleSectionKey,
  ) => boolean,
): CharacterDictionaryTermEntry[] {
  return termEntries.map((entry) => {
    const glossary = entry[5].map((item) =>
      applyCollapsibleOpenStatesToStructuredValue(item, getCollapsibleSectionOpenState),
    ) as CharacterDictionaryGlossaryEntry[];
    return [...entry.slice(0, 5), glossary, ...entry.slice(6)] as CharacterDictionaryTermEntry;
  });
}
