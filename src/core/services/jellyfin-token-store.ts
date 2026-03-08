import * as fs from 'fs';
import * as path from 'path';
import electron from 'electron';

const { safeStorage } = electron;

interface PersistedSessionPayload {
  encryptedSession?: string;
  plaintextSession?: {
    accessToken?: string;
    userId?: string;
  };
  // Legacy payload fields (token only).
  encryptedToken?: string;
  plaintextToken?: string;
  updatedAt?: number;
}

export interface JellyfinStoredSession {
  accessToken: string;
  userId: string;
}

export interface JellyfinTokenStore {
  loadSession: () => JellyfinStoredSession | null;
  saveSession: (session: JellyfinStoredSession) => void;
  clearSession: () => void;
}

function ensureDirectory(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function writePayload(filePath: string, payload: PersistedSessionPayload): void {
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
    loadSession(): JellyfinStoredSession | null {
      if (!fs.existsSync(filePath)) {
        return null;
      }
      try {
        const raw = fs.readFileSync(filePath, 'utf-8');
        const parsed = JSON.parse(raw) as PersistedSessionPayload;

        if (typeof parsed.encryptedSession === 'string' && parsed.encryptedSession.length > 0) {
          const encrypted = Buffer.from(parsed.encryptedSession, 'base64');
          if (!safeStorage.isEncryptionAvailable()) {
            logger.warn('Jellyfin session encryption is not available on this system.');
            return null;
          }
          const decrypted = safeStorage.decryptString(encrypted).trim();
          const session = JSON.parse(decrypted) as Partial<JellyfinStoredSession>;
          const accessToken =
            typeof session.accessToken === 'string' ? session.accessToken.trim() : '';
          const userId = typeof session.userId === 'string' ? session.userId.trim() : '';
          if (!accessToken || !userId) return null;
          return { accessToken, userId };
        }

        if (parsed.plaintextSession && typeof parsed.plaintextSession === 'object') {
          const accessToken =
            typeof parsed.plaintextSession.accessToken === 'string'
              ? parsed.plaintextSession.accessToken.trim()
              : '';
          const userId =
            typeof parsed.plaintextSession.userId === 'string'
              ? parsed.plaintextSession.userId.trim()
              : '';
          if (accessToken && userId) {
            const session = { accessToken, userId };
            this.saveSession(session);
            return session;
          }
        }

        if (
          (typeof parsed.encryptedToken === 'string' && parsed.encryptedToken.length > 0) ||
          (typeof parsed.plaintextToken === 'string' && parsed.plaintextToken.trim().length > 0)
        ) {
          logger.warn(
            'Ignoring legacy Jellyfin token-only store payload because userId is missing.',
          );
        }
      } catch (error) {
        logger.error('Failed to read Jellyfin session store.', error);
      }
      return null;
    },

    saveSession(session: JellyfinStoredSession): void {
      const accessToken = session.accessToken.trim();
      const userId = session.userId.trim();
      if (!accessToken || !userId) {
        this.clearSession();
        return;
      }
      try {
        if (!safeStorage.isEncryptionAvailable()) {
          logger.warn(
            'Jellyfin session encryption unavailable; storing session in plaintext fallback.',
          );
          writePayload(filePath, {
            plaintextSession: {
              accessToken,
              userId,
            },
            updatedAt: Date.now(),
          });
          return;
        }
        const encrypted = safeStorage.encryptString(JSON.stringify({ accessToken, userId }));
        writePayload(filePath, {
          encryptedSession: encrypted.toString('base64'),
          updatedAt: Date.now(),
        });
      } catch (error) {
        logger.error('Failed to persist Jellyfin session.', error);
      }
    },

    clearSession(): void {
      if (!fs.existsSync(filePath)) return;
      try {
        fs.unlinkSync(filePath);
        logger.info('Cleared stored Jellyfin session.');
      } catch (error) {
        logger.error('Failed to clear stored Jellyfin session.', error);
      }
    },
  };
}
