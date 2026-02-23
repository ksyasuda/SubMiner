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

test('anilist token store saves and loads encrypted token', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(true));
  store.saveToken('  demo-token  ');

  const payload = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as {
    encryptedToken?: string;
    plaintextToken?: string;
  };
  assert.equal(typeof payload.encryptedToken, 'string');
  assert.equal(payload.plaintextToken, undefined);
  assert.equal(store.loadToken(), 'demo-token');
});

test('anilist token store falls back to plaintext when encryption unavailable', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(false));
  store.saveToken('plain-token');

  const payload = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as {
    plaintextToken?: string;
  };
  assert.equal(payload.plaintextToken, 'plain-token');
  assert.equal(store.loadToken(), 'plain-token');
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

test('anilist token store clears persisted token file', () => {
  const filePath = createTempTokenFile();
  const store = createAnilistTokenStore(filePath, createLogger(), createStorage(true));
  store.saveToken('to-clear');
  assert.equal(fs.existsSync(filePath), true);
  store.clearToken();
  assert.equal(fs.existsSync(filePath), false);
});
