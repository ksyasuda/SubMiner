import test from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { safeStorage } from "electron";

import { createAnilistTokenStore } from "./anilist-token-store";

function createTempTokenFile(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "subminer-anilist-token-"));
  return path.join(dir, "token.json");
}

function createLogger() {
  return {
    info: (_message: string) => {},
    warn: (_message: string) => {},
    error: (_message: string) => {},
  };
}

type SafeStorageLike = {
  isEncryptionAvailable: () => boolean;
  encryptString: (value: string) => Buffer;
  decryptString: (value: Buffer) => string;
};

const safeStorageApi = safeStorage as unknown as Partial<SafeStorageLike>;
const hasSafeStorage =
  typeof safeStorageApi?.isEncryptionAvailable === "function" &&
  typeof safeStorageApi?.encryptString === "function" &&
  typeof safeStorageApi?.decryptString === "function";

const originalSafeStorage: SafeStorageLike | null = hasSafeStorage
  ? {
      isEncryptionAvailable:
        safeStorageApi.isEncryptionAvailable as () => boolean,
      encryptString: safeStorageApi.encryptString as (value: string) => Buffer,
      decryptString: safeStorageApi.decryptString as (value: Buffer) => string,
    }
  : null;

function mockSafeStorage(encryptionAvailable: boolean): void {
  if (!hasSafeStorage) return;
  (
    safeStorage as unknown as {
      isEncryptionAvailable: typeof safeStorage.isEncryptionAvailable;
      encryptString: typeof safeStorage.encryptString;
      decryptString: typeof safeStorage.decryptString;
    }
  ).isEncryptionAvailable = () => encryptionAvailable;
  (
    safeStorage as unknown as {
      encryptString: typeof safeStorage.encryptString;
      decryptString: typeof safeStorage.decryptString;
    }
  ).encryptString = (value: string) => Buffer.from(`enc:${value}`, "utf-8");
  (
    safeStorage as unknown as {
      decryptString: typeof safeStorage.decryptString;
    }
  ).decryptString = (value: Buffer) => {
    const raw = value.toString("utf-8");
    return raw.startsWith("enc:") ? raw.slice(4) : raw;
  };
}

function restoreSafeStorage(): void {
  if (!hasSafeStorage || !originalSafeStorage) return;
  (
    safeStorage as unknown as {
      isEncryptionAvailable: typeof safeStorage.isEncryptionAvailable;
      encryptString: typeof safeStorage.encryptString;
      decryptString: typeof safeStorage.decryptString;
    }
  ).isEncryptionAvailable = originalSafeStorage.isEncryptionAvailable;
  (
    safeStorage as unknown as {
      encryptString: typeof safeStorage.encryptString;
      decryptString: typeof safeStorage.decryptString;
    }
  ).encryptString = originalSafeStorage.encryptString;
  (
    safeStorage as unknown as {
      decryptString: typeof safeStorage.decryptString;
    }
  ).decryptString = originalSafeStorage.decryptString;
}

test(
  "anilist token store saves and loads encrypted token",
  { skip: !hasSafeStorage },
  () => {
    mockSafeStorage(true);
    try {
      const filePath = createTempTokenFile();
      const store = createAnilistTokenStore(filePath, createLogger());
      store.saveToken("  demo-token  ");

      const payload = JSON.parse(fs.readFileSync(filePath, "utf-8")) as {
        encryptedToken?: string;
        plaintextToken?: string;
      };
      assert.equal(typeof payload.encryptedToken, "string");
      assert.equal(payload.plaintextToken, undefined);
      assert.equal(store.loadToken(), "demo-token");
    } finally {
      restoreSafeStorage();
    }
  },
);

test(
  "anilist token store falls back to plaintext when encryption unavailable",
  { skip: !hasSafeStorage },
  () => {
    mockSafeStorage(false);
    try {
      const filePath = createTempTokenFile();
      const store = createAnilistTokenStore(filePath, createLogger());
      store.saveToken("plain-token");

      const payload = JSON.parse(fs.readFileSync(filePath, "utf-8")) as {
        plaintextToken?: string;
      };
      assert.equal(payload.plaintextToken, "plain-token");
      assert.equal(store.loadToken(), "plain-token");
    } finally {
      restoreSafeStorage();
    }
  },
);

test(
  "anilist token store migrates legacy plaintext to encrypted",
  { skip: !hasSafeStorage },
  () => {
    const filePath = createTempTokenFile();
    fs.writeFileSync(
      filePath,
      JSON.stringify({ plaintextToken: "legacy-token", updatedAt: Date.now() }),
      "utf-8",
    );

    mockSafeStorage(true);
    try {
      const store = createAnilistTokenStore(filePath, createLogger());
      assert.equal(store.loadToken(), "legacy-token");

      const payload = JSON.parse(fs.readFileSync(filePath, "utf-8")) as {
        encryptedToken?: string;
        plaintextToken?: string;
      };
      assert.equal(typeof payload.encryptedToken, "string");
      assert.equal(payload.plaintextToken, undefined);
    } finally {
      restoreSafeStorage();
    }
  },
);

test(
  "anilist token store clears persisted token file",
  { skip: !hasSafeStorage },
  () => {
    mockSafeStorage(true);
    try {
      const filePath = createTempTokenFile();
      const store = createAnilistTokenStore(filePath, createLogger());
      store.saveToken("to-clear");
      assert.equal(fs.existsSync(filePath), true);
      store.clearToken();
      assert.equal(fs.existsSync(filePath), false);
    } finally {
      restoreSafeStorage();
    }
  },
);
