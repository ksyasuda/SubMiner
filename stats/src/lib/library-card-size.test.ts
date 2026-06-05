import assert from 'node:assert/strict';
import test from 'node:test';
import {
  DEFAULT_LIBRARY_CARD_SIZE,
  LIBRARY_CARD_SIZE_STORAGE_KEY,
  getLibraryCardSizeStorage,
  readLibraryCardSizePreference,
  writeLibraryCardSizePreference,
} from './library-card-size';

function createStorage(initial: Record<string, string | null> = {}): Storage {
  const values = new Map(
    Object.entries(initial).filter((entry): entry is [string, string] => {
      return entry[1] !== null;
    }),
  );

  return {
    get length() {
      return values.size;
    },
    clear() {
      values.clear();
    },
    getItem(key: string) {
      return values.get(key) ?? null;
    },
    key(index: number) {
      return Array.from(values.keys())[index] ?? null;
    },
    removeItem(key: string) {
      values.delete(key);
    },
    setItem(key: string, value: string) {
      values.set(key, value);
    },
  };
}

test('readLibraryCardSizePreference returns saved valid sizes', () => {
  const storage = createStorage({ [LIBRARY_CARD_SIZE_STORAGE_KEY]: 'lg' });

  assert.equal(readLibraryCardSizePreference(storage), 'lg');
});

test('readLibraryCardSizePreference falls back for missing or invalid saved sizes', () => {
  assert.equal(readLibraryCardSizePreference(createStorage()), DEFAULT_LIBRARY_CARD_SIZE);
  assert.equal(
    readLibraryCardSizePreference(createStorage({ [LIBRARY_CARD_SIZE_STORAGE_KEY]: 'xl' })),
    DEFAULT_LIBRARY_CARD_SIZE,
  );
});

test('library card size preference helpers ignore storage failures', () => {
  const storage = {
    getItem() {
      throw new Error('blocked');
    },
    setItem() {
      throw new Error('blocked');
    },
  } as unknown as Storage;

  assert.equal(readLibraryCardSizePreference(storage), DEFAULT_LIBRARY_CARD_SIZE);
  assert.doesNotThrow(() => writeLibraryCardSizePreference(storage, 'sm'));
});

test('getLibraryCardSizeStorage returns null when localStorage access is blocked', () => {
  const source = {
    get localStorage(): Storage {
      throw new Error('blocked');
    },
  };

  assert.equal(getLibraryCardSizeStorage(source), null);
});
