import * as path from 'path';
import { writeStoredZip } from '../../shared/stored-zip';
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

export function buildDictionaryZip(
  outputPath: string,
  dictionaryTitle: string,
  description: string,
  revision: string,
  termEntries: CharacterDictionaryTermEntry[],
  images: CharacterDictionarySnapshotImage[],
): { zipPath: string; entryCount: number } {
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

    const entriesPerBank = 10_000;
    for (let i = 0; i < termEntries.length; i += entriesPerBank) {
      yield {
        name: `term_bank_${Math.floor(i / entriesPerBank) + 1}.json`,
        data: Buffer.from(JSON.stringify(termEntries.slice(i, i + entriesPerBank)), 'utf8'),
      };
    }
  }

  writeStoredZip(outputPath, zipFiles());
  return { zipPath: outputPath, entryCount: termEntries.length };
}
