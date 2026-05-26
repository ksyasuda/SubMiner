import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { getSnapshotPath, writeSnapshot } from './cache';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import {
  buildCharacterNameImageIndexFromSnapshots,
  createCharacterDictionaryImageLookup,
} from './image-lookup';
import type { CharacterDictionarySnapshot } from './types';

const PNG_1X1_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-image-lookup-'));
}

test('buildCharacterNameImageIndexFromSnapshots maps name terms to character portrait data URLs', () => {
  const outputDir = makeTempDir();
  const snapshot: CharacterDictionarySnapshot = {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 130298,
    mediaTitle: 'The Eminence in Shadow',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [
      [
        'アレクシア',
        'あれくしあ',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: {
              tag: 'div',
              content: [
                { tag: 'div', content: 'アレクシア・ミドガル' },
                {
                  tag: 'div',
                  content: {
                    tag: 'img',
                    path: 'img/m130298-c123.png',
                    alt: 'アレクシア・ミドガル',
                  },
                },
                {
                  tag: 'details',
                  content: [
                    { tag: 'summary', content: 'Voiced by' },
                    {
                      tag: 'div',
                      content: {
                        tag: 'img',
                        path: 'img/m130298-va456.png',
                        alt: 'VA',
                      },
                    },
                  ],
                },
              ],
            },
          },
        ],
        0,
        '',
      ],
    ],
    images: [
      { path: 'img/m130298-c123.png', dataBase64: 'AAAA' },
      { path: 'img/m130298-va456.png', dataBase64: 'BBBB' },
    ],
  };
  writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const index = buildCharacterNameImageIndexFromSnapshots(outputDir);

  assert.deepEqual(index.get('アレクシア'), {
    src: 'data:image/png;base64,AAAA',
    alt: 'アレクシア・ミドガル',
  });
});

test('buildCharacterNameImageIndexFromSnapshots sniffs image MIME from bytes before path extension', () => {
  const outputDir = makeTempDir();
  const snapshot: CharacterDictionarySnapshot = {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 130298,
    mediaTitle: 'The Eminence in Shadow',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [
      [
        'アレクシア',
        'あれくしあ',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: {
              tag: 'img',
              path: 'img/m130298-c123.jpg',
              alt: 'アレクシア・ミドガル',
            },
          },
        ],
        0,
        '',
      ],
    ],
    images: [{ path: 'img/m130298-c123.jpg', dataBase64: PNG_1X1_BASE64 }],
  };
  writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const index = buildCharacterNameImageIndexFromSnapshots(outputDir);

  assert.equal(index.get('アレクシア')?.src, `data:image/png;base64,${PNG_1X1_BASE64}`);
});

test('createCharacterDictionaryImageLookup can scope duplicate names to the current media', () => {
  const outputDir = makeTempDir();
  const towerSnapshot: CharacterDictionarySnapshot = {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 115230,
    mediaTitle: 'Tower of God',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [
      [
        'カズ',
        'かず',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: { tag: 'img', path: 'img/m115230-c1.png', alt: 'Kaz' },
          },
        ],
        0,
        '',
      ],
    ],
    images: [{ path: 'img/m115230-c1.png', dataBase64: 'TOWER' }],
  };
  const konosubaSnapshot: CharacterDictionarySnapshot = {
    ...towerSnapshot,
    mediaId: 21202,
    mediaTitle: 'KonoSuba',
    termEntries: [
      [
        'カズ',
        'かず',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: { tag: 'img', path: 'img/m21202-c2.png', alt: 'Kazuma' },
          },
        ],
        0,
        '',
      ],
    ],
    images: [{ path: 'img/m21202-c2.png', dataBase64: 'KONOSUBA' }],
  };
  writeSnapshot(getSnapshotPath(outputDir, towerSnapshot.mediaId), towerSnapshot);
  writeSnapshot(getSnapshotPath(outputDir, konosubaSnapshot.mediaId), konosubaSnapshot);

  const lookup = createCharacterDictionaryImageLookup({ outputDir });

  assert.equal(lookup.get('カズ', 21202)?.alt, 'Kazuma');
});

test('createCharacterDictionaryImageLookup does not fall back globally on scoped miss', () => {
  const outputDir = makeTempDir();
  const snapshot: CharacterDictionarySnapshot = {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 115230,
    mediaTitle: 'Tower of God',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [
      [
        'カズ',
        'かず',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: { tag: 'img', path: 'img/m115230-c1.png', alt: 'Kaz' },
          },
        ],
        0,
        '',
      ],
    ],
    images: [{ path: 'img/m115230-c1.png', dataBase64: 'TOWER' }],
  };
  writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const lookup = createCharacterDictionaryImageLookup({ outputDir });

  assert.equal(lookup.get('カズ', 21202), null);
  assert.equal(lookup.get('カズ')?.alt, 'Kaz');
});
