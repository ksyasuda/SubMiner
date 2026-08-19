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

// Lookup indexes rebuild in the background while gets serve stale data, so tests poll until the
// refresh they triggered has landed.
async function waitForRefresh<T>(probe: () => T | null | undefined): Promise<T> {
  const deadline = Date.now() + 5000;
  for (;;) {
    const value = probe();
    if (value !== null && value !== undefined) {
      return value;
    }
    if (Date.now() > deadline) {
      throw new Error('timed out waiting for background snapshot refresh');
    }
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

const PNG_1X1_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-image-lookup-'));
}

test('buildCharacterNameImageIndexFromSnapshots maps name terms to character portrait data URLs', async () => {
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
  await writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const index = await buildCharacterNameImageIndexFromSnapshots(outputDir);

  assert.deepEqual(index.get('アレクシア'), {
    src: 'data:image/png;base64,AAAA',
    alt: 'アレクシア・ミドガル',
  });
});

test('buildCharacterNameImageIndexFromSnapshots sniffs image MIME from bytes before path extension', async () => {
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
  await writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const index = await buildCharacterNameImageIndexFromSnapshots(outputDir);

  assert.equal(index.get('アレクシア')?.src, `data:image/png;base64,${PNG_1X1_BASE64}`);
});

test('createCharacterDictionaryImageLookup can scope duplicate names to the current media', async () => {
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
  await writeSnapshot(getSnapshotPath(outputDir, towerSnapshot.mediaId), towerSnapshot);
  await writeSnapshot(getSnapshotPath(outputDir, konosubaSnapshot.mediaId), konosubaSnapshot);

  const lookup = createCharacterDictionaryImageLookup({ outputDir });

  const scoped = await waitForRefresh(() => lookup.get('カズ', 21202));
  assert.equal(scoped.alt, 'Kazuma');
});

test('createCharacterDictionaryImageLookup reports and retries a failed index-ready callback', async () => {
  const outputDir = makeTempDir();
  const snapshot: CharacterDictionarySnapshot = {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 21858,
    mediaTitle: 'Little Witch Academia',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [
      [
        'ダイアナ',
        'だいあな',
        'name primary',
        '',
        75,
        [
          {
            type: 'structured-content',
            content: {
              tag: 'img',
              path: 'img/m21858-c81709.png',
              alt: 'ダイアナ・キャベンディッシュ',
            },
          },
        ],
        0,
        '',
      ],
    ],
    images: [{ path: 'img/m21858-c81709.png', dataBase64: PNG_1X1_BASE64 }],
  };
  await writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);
  const callbackError = new Error('annotation refresh failed');
  let readyCount = 0;
  const reportedErrors: unknown[] = [];
  const lookup = createCharacterDictionaryImageLookup({
    outputDir,
    onIndexReady: () => {
      readyCount += 1;
      if (readyCount === 1) {
        throw callbackError;
      }
    },
    onIndexReadyError: (error) => reportedErrors.push(error),
  });

  assert.equal(lookup.get('ダイアナ', snapshot.mediaId), null);
  await waitForRefresh(() => lookup.get('ダイアナ', snapshot.mediaId));

  assert.equal(readyCount, 2);
  assert.deepEqual(reportedErrors, [callbackError]);
  lookup.get('ダイアナ', snapshot.mediaId);
  assert.equal(readyCount, 2);
});

test('createCharacterDictionaryImageLookup does not fall back globally on scoped miss', async () => {
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
  await writeSnapshot(getSnapshotPath(outputDir, snapshot.mediaId), snapshot);

  const lookup = createCharacterDictionaryImageLookup({ outputDir });

  const unscoped = await waitForRefresh(() => lookup.get('カズ'));
  assert.equal(unscoped.alt, 'Kaz');
  assert.equal(lookup.get('カズ', 21202), null);
});
