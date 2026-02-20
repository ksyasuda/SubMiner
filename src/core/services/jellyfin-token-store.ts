import * as fs from 'fs';
import * as path from 'path';
import { safeStorage } from 'electron';

interface PersistedTokenPayload {
  encryptedToken?: string;
  plaintextToken?: string;
  updatedAt?: number;
}

export interface JellyfinTokenStore {
  loadToken: () => string | null;
  saveToken: (token: string) => void;
  clearToken: () => void;
}

function ensureDirectory(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function writePayload(filePath: string, payload: PersistedTokenPayload): void {
  ensureDirectory(filePath);
  fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
}

export function createJellyfinTokenStore(
  filePath: string,
  logger: {
    info: (message: string) => void;
    warn: (message: string, details?: unknown) => void;
    error: (message: string, details?: unknown) => void;
  },
): JellyfinTokenStore {
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
          if (!safeStorage.isEncryptionAvailable()) {
            logger.warn('Jellyfin token encryption is not available on this system.');
            return null;
          }
          const decrypted = safeStorage.decryptString(encrypted).trim();
          return decrypted.length > 0 ? decrypted : null;
        }
        if (typeof parsed.plaintextToken === 'string' && parsed.plaintextToken.trim().length > 0) {
          const plaintext = parsed.plaintextToken.trim();
          this.saveToken(plaintext);
          return plaintext;
        }
      } catch (error) {
        logger.error('Failed to read Jellyfin token store.', error);
      }
      return null;
    },

    saveToken(token: string): void {
      const trimmed = token.trim();
      if (trimmed.length === 0) {
        this.clearToken();
        return;
      }
      try {
        if (!safeStorage.isEncryptionAvailable()) {
          logger.warn('Jellyfin token encryption unavailable; storing token in plaintext fallback.');
          writePayload(filePath, {
            plaintextToken: trimmed,
            updatedAt: Date.now(),
          });
          return;
        }
        const encrypted = safeStorage.encryptString(trimmed);
        writePayload(filePath, {
          encryptedToken: encrypted.toString('base64'),
          updatedAt: Date.now(),
        });
      } catch (error) {
        logger.error('Failed to persist Jellyfin token.', error);
      }
    },

    clearToken(): void {
      if (!fs.existsSync(filePath)) return;
      try {
        fs.unlinkSync(filePath);
        logger.info('Cleared stored Jellyfin token.');
      } catch (error) {
        logger.error('Failed to clear stored Jellyfin token.', error);
      }
    },
  };
}
