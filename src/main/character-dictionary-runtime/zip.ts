import * as path from 'path';
import { readStoredZipFirstFile, writeStoredZipAsync } from '../../shared/stored-zip';
import { ensureDir } from './fs-utils';
import type { CharacterDictionarySnapshotImage, CharacterDictionaryTermEntry } from './types';

export function buildDictionaryTitle(mediaId: number): string {
  return `SubMiner Character Dictionary (AniList ${mediaId})`;
}

function createIndex(
  dictionaryTitle: string,
  description: string,
  revision: string,
): Record<string, unknown> {
  return {
    title: dictionaryTitle,
    revision,
    format: 3,
    author: 'SubMiner',
    description,
  };
}

function createTagBank(): Array<[string, string, number, string, number]> {
  return [
    ['name', 'partOfSpeech', 0, 'Character name', 0],
    ['main', 'name', 0, 'Protagonist', 0],
    ['primary', 'name', 0, 'Main character', 0],
    ['side', 'name', 0, 'Side character', 0],
    ['appears', 'name', 0, 'Minor appearance', 0],
  ];
}

/**
 * Revision recorded inside a built dictionary ZIP, or null when the archive is missing, truncated,
 * or not one of ours. `index.json` is always the first entry written by {@link buildDictionaryZip}.
 */
export function readDictionaryZipRevision(zipPath: string): string | null {
  const firstFile = readStoredZipFirstFile(zipPath);
  if (!firstFile || firstFile.name !== 'index.json') {
    return null;
  }
  try {
    const index = JSON.parse(firstFile.data.toString('utf8')) as { revision?: unknown };
    return typeof index.revision === 'string' && index.revision.length > 0 ? index.revision : null;
  } catch {
    return null;
  }
}

export async function buildDictionaryZip(
  outputPath: string,
  dictionaryTitle: string,
  description: string,
  revision: string,
  termEntries: CharacterDictionaryTermEntry[],
  images: CharacterDictionarySnapshotImage[],
): Promise<{ zipPath: string; entryCount: number }> {
  ensureDir(path.dirname(outputPath));

  function* zipFiles(): Iterable<{ name: string; data: Buffer }> {
    yield {
      name: 'index.json',
      data: Buffer.from(
        JSON.stringify(createIndex(dictionaryTitle, description, revision), null, 2),
        'utf8',
      ),
    };
    yield {
      name: 'tag_bank_1.json',
      data: Buffer.from(JSON.stringify(createTagBank()), 'utf8'),
    };

    for (const image of images) {
      yield {
        name: image.path,
        data: Buffer.from(image.dataBase64, 'base64'),
      };
    }

    // Each bank is stringified in one shot, so the bank size sets the longest single block in the
    // build. 10k entries measured ~38MB and ~135ms per bank on a real merged dictionary; 2k keeps
    // every bank under the archive writer's yield budget at ~27ms. Yomitan reads any number of
    // term_bank_N.json files, so this only changes how the terms are split across them.
    const entriesPerBank = 2_000;
    for (let i = 0; i < termEntries.length; i += entriesPerBank) {
      yield {
        name: `term_bank_${Math.floor(i / entriesPerBank) + 1}.json`,
        data: Buffer.from(JSON.stringify(termEntries.slice(i, i + entriesPerBank)), 'utf8'),
      };
    }
  }

  await writeStoredZipAsync(outputPath, zipFiles());
  return { zipPath: outputPath, entryCount: termEntries.length };
}
