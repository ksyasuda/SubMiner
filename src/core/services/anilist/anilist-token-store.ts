import * as fs from 'fs';
import * as electron from 'electron';
import { ensureDirForFile } from '../../../shared/fs-utils';

interface PersistedTokenPayload {
  encryptedToken?: string;
  plaintextToken?: string;
  updatedAt?: number;
}

export interface AnilistTokenStore {
  loadToken: () => string | null;
  saveToken: (token: string) => boolean;
  clearToken: () => void;
}

export interface SafeStorageLike {
  isEncryptionAvailable: () => boolean;
  encryptString: (value: string) => Buffer;
  decryptString: (value: Buffer) => string;
  getSelectedStorageBackend?: () => string;
}

function writePayload(filePath: string, payload: PersistedTokenPayload): void {
  ensureDirForFile(filePath);
  fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
}

export function createAnilistTokenStore(
  filePath: string,
  logger: {
    info: (message: string) => void;
    warn: (message: string, details?: unknown) => void;
    error: (message: string, details?: unknown) => void;
    warnUser?: (message: string) => void;
  },
  storage: SafeStorageLike = electron.safeStorage,
): AnilistTokenStore {
  let safeStorageUsable: boolean | null = null;

  const getSelectedBackend = (): string => {
    if (typeof storage.getSelectedStorageBackend !== 'function') {
      return 'unsupported';
    }
    try {
      return storage.getSelectedStorageBackend();
    } catch {
      return 'error';
    }
  };

  const getSafeStorageDebugContext = (): string =>
    JSON.stringify({
      platform: process.platform,
      dbusSession: process.env.DBUS_SESSION_BUS_ADDRESS,
      xdgRuntimeDir: process.env.XDG_RUNTIME_DIR,
      display: process.env.DISPLAY,
      waylandDisplay: process.env.WAYLAND_DISPLAY,
      hasDefaultApp: Boolean(process.defaultApp),
      selectedSafeStorageBackend: getSelectedBackend(),
    });

  const isSafeStorageUsable = (): boolean => {
    if (safeStorageUsable != null) return safeStorageUsable;

    try {
      if (!storage.isEncryptionAvailable()) {
        notifyUser(
          `AniList token encryption unavailable: safeStorage.isEncryptionAvailable() is false. ` +
            `Context: ${getSafeStorageDebugContext()}`,
        );
        return false;
      }
      const probe = storage.encryptString('__subminer_anilist_probe__');
      if (probe.equals(Buffer.from('__subminer_anilist_probe__'))) {
        notifyUser(
          'AniList token encryption probe failed: safeStorage.encryptString() returned plaintext bytes.',
        );
        return false;
      }
      const roundTrip = storage.decryptString(probe);
      if (roundTrip !== '__subminer_anilist_probe__') {
        notifyUser(
          'AniList token encryption probe failed: encrypt/decrypt round trip returned unexpected content.',
        );
        return false;
      }
      safeStorageUsable = true;
      return true;
    } catch (error) {
      logger.error('AniList token encryption probe failed.', error);
      notifyUser(
        `AniList token encryption unavailable: safeStorage probe threw an error. ` +
          `Context: ${getSafeStorageDebugContext()}`,
      );
      return false;
    }
  };

  const notifyUser = (message: string): void => {
    logger.warn(message);
    logger.warnUser?.(message);
  };

  return {
    loadToken(): string | null {
      if (!fs.existsSync(filePath)) {
        return null;
      }
      try {
        const raw = fs.readFileSync(filePath, 'utf-8');
        const parsed = JSON.parse(raw) as PersistedTokenPayload;
        if (typeof parsed.encryptedToken === 'string' && parsed.encryptedToken.length > 0) {
          const encrypted = Buffer.from(parsed.encryptedToken, 'base64');
          if (!isSafeStorageUsable()) {
            return null;
          }
          const decrypted = storage.decryptString(encrypted).trim();
          if (decrypted.length === 0) {
            return null;
          }
          return decrypted;
        }
        if (typeof parsed.plaintextToken === 'string' && parsed.plaintextToken.trim().length > 0) {
          if (storage.isEncryptionAvailable()) {
            if (!isSafeStorageUsable()) {
              return null;
            }
            const plaintext = parsed.plaintextToken.trim();
            notifyUser(
              'AniList token plaintext fallback payload found. Migrating to encrypted storage.',
            );
            this.saveToken(plaintext);
            return plaintext;
          }
          notifyUser(
            'AniList token plaintext was found but ignored because safe storage is unavailable.',
          );
          this.clearToken();
          return null;
        }
      } catch (error) {
        logger.error('Failed to read AniList token store.', error);
      }
      return null;
    },

    saveToken(token: string): boolean {
      const trimmed = token.trim();
      if (trimmed.length === 0) {
        this.clearToken();
        return true;
      }
      try {
        if (!isSafeStorageUsable()) {
          notifyUser(
            'AniList token encryption is unavailable; refusing to store access token. Re-login required after restart.',
          );
          return false;
        }
        const encrypted = storage.encryptString(trimmed);
        writePayload(filePath, {
          encryptedToken: encrypted.toString('base64'),
          updatedAt: Date.now(),
        });
        return true;
      } catch (error) {
        logger.error('Failed to persist AniList token.', error);
        return false;
      }
    },

    clearToken(): void {
      if (!fs.existsSync(filePath)) return;
      try {
        fs.unlinkSync(filePath);
        logger.info('Cleared stored AniList token.');
      } catch (error) {
        logger.error('Failed to clear stored AniList token.', error);
      }
    },
  };
}
