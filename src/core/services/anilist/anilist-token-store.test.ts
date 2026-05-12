import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { createAnilistTokenStore, type SafeStorageLike } from './anilist-token-store';

function createTempTokenFile(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anilist-token-'));
  return path.join(dir, 'token.json');
}

function createLogger() {
  return {
    info: (_message: string) => {},
    warn: (_message: string) => {},
    error: (_message: string) => {},
  };
}

function createStorage(encryptionAvailable: boolean): SafeStorageLike {
  return {
    isEncryptionAvailable: () => encryptionAvailable,
    encryptString: (value: string) => Buffer.from(`enc:${value}`, 'utf-8'),
    decryptString: (value: Buffer) => {
      const raw = value.toString('utf-8');
      return raw.startsWith('enc:') ? raw.slice(4) : raw;
    },
  };
}

function createPassthroughStorage(): SafeStorageLike {
  return {
    isEncryptionAvailable: () => true,
    encryptString: (value: string) => Buffer.from(value, 'utf-8'),
    decryptString: (value: Buffer) => value.toString('utf-8'),
  };
}

function createTransientUnavailableStorage(): SafeStorageLike & {
  setAvailable: (next: boolean) => void;
} {
  let available = false;
  return {
    isEncryptionAvailable: () => available,
    encryptString: (value: string) => Buffer.from(`enc:${value}`, 'utf-8'),
    decryptString: (value: Buffer) => {
      const raw = value.toString('utf-8');
      return raw.startsWith('enc:') ? raw.slice(4) : raw;
    },
    getSelectedStorageBackend: () => (available ? 'gnome_libsecret' : 'unknown'),
    setAvailable(next: boolean) {
      available = next;
    },
  } as SafeStorageLike & { setAvailable: (next: boolean) => void };
}

test('anilist token store saves and loads encrypted token', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(true));
  assert.equal(store.saveToken('  demo-token  '), true);

  const payload = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as {
    encryptedToken?: string;
    plaintextToken?: string;
  };
  assert.equal(typeof payload.encryptedToken, 'string');
  assert.equal(payload.plaintextToken, undefined);
  assert.equal(store.loadToken(), 'demo-token');
});

test('anilist token store refuses to persist token when encryption unavailable', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(false));
  assert.equal(store.saveToken('plain-token'), false);

  assert.equal(fs.existsSync(filePath), false);
  assert.equal(store.loadToken(), null);
});

test('anilist token store retries safeStorage after transient encryption unavailability', () => {
  const filePath = createTempTokenFile();
  fs.writeFileSync(
    filePath,
    JSON.stringify({
      encryptedToken: Buffer.from('enc:stored-token', 'utf-8').toString('base64'),
      updatedAt: Date.now(),
    }),
    'utf-8',
  );
  const storage = createTransientUnavailableStorage();
  const store = createAnilistTokenStore(filePath, createLogger(), storage);

  assert.equal(store.loadToken(), null);
  storage.setAvailable(true);

  assert.equal(store.loadToken(), 'stored-token');
  assert.equal(store.saveToken('new-token'), true);
  assert.equal(store.loadToken(), 'new-token');
});

test('anilist token store migrates legacy plaintext to encrypted', () => {
  const filePath = createTempTokenFile();
  fs.writeFileSync(
    filePath,
    JSON.stringify({ plaintextToken: 'legacy-token', updatedAt: Date.now() }),
    'utf-8',
  );

  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(true));
  assert.equal(store.loadToken(), 'legacy-token');

  const payload = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as {
    encryptedToken?: string;
    plaintextToken?: string;
  };
  assert.equal(typeof payload.encryptedToken, 'string');
  assert.equal(payload.plaintextToken, undefined);
});

test('anilist token store refuses passthrough safeStorage implementation', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createPassthroughStorage());
  assert.equal(store.saveToken('demo-token'), false);
  assert.equal(store.loadToken(), null);
});

test('anilist token store clears persisted token file', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(true));
  store.saveToken('to-clear');
  assert.equal(fs.existsSync(filePath), true);
  store.clearToken();
  assert.equal(fs.existsSync(filePath), false);
});
