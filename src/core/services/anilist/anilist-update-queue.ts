import * as fs from 'fs';
import * as path from 'path';

const INITIAL_BACKOFF_MS = 30_000;
const MAX_BACKOFF_MS = 6 * 60 * 60 * 1000;
const MAX_ATTEMPTS = 8;
const MAX_ITEMS = 500;

export interface AnilistQueuedUpdate {
  key: string;
  title: string;
  episode: number;
  createdAt: number;
  attemptCount: number;
  nextAttemptAt: number;
  lastError: string | null;
}

interface AnilistRetryQueuePayload {
  pending?: AnilistQueuedUpdate[];
  deadLetter?: AnilistQueuedUpdate[];
}

export interface AnilistRetryQueueSnapshot {
  pending: number;
  ready: number;
  deadLetter: number;
}

export interface AnilistUpdateQueue {
  enqueue: (key: string, title: string, episode: number) => void;
  nextReady: (nowMs?: number) => AnilistQueuedUpdate | null;
  markSuccess: (key: string) => void;
  markFailure: (key: string, reason: string, nowMs?: number) => void;
  getSnapshot: (nowMs?: number) => AnilistRetryQueueSnapshot;
}

function ensureDir(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function clampBackoffMs(attemptCount: number): number {
  const computed = INITIAL_BACKOFF_MS * Math.pow(2, Math.max(0, attemptCount - 1));
  return Math.min(MAX_BACKOFF_MS, computed);
}

export function createAnilistUpdateQueue(
  filePath: string,
  logger: {
    info: (message: string) => void;
    warn: (message: string, details?: unknown) => void;
    error: (message: string, details?: unknown) => void;
  },
): AnilistUpdateQueue {
  let pending: AnilistQueuedUpdate[] = [];
  let deadLetter: AnilistQueuedUpdate[] = [];

  const persist = () => {
    try {
      ensureDir(filePath);
      const payload: AnilistRetryQueuePayload = { pending, deadLetter };
      fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
    } catch (error) {
      logger.error('Failed to persist AniList retry queue.', error);
    }
  };

  const load = () => {
    if (!fs.existsSync(filePath)) {
      return;
    }
    try {
      const raw = fs.readFileSync(filePath, 'utf-8');
      const parsed = JSON.parse(raw) as AnilistRetryQueuePayload;
      const parsedPending = Array.isArray(parsed.pending) ? parsed.pending : [];
      const parsedDeadLetter = Array.isArray(parsed.deadLetter) ? parsed.deadLetter : [];
      pending = parsedPending
        .filter(
          (item): item is AnilistQueuedUpdate =>
            item &&
            typeof item.key === 'string' &&
            typeof item.title === 'string' &&
            typeof item.episode === 'number' &&
            item.episode > 0 &&
            typeof item.createdAt === 'number' &&
            typeof item.attemptCount === 'number' &&
            typeof item.nextAttemptAt === 'number' &&
            (typeof item.lastError === 'string' || item.lastError === null),
        )
        .slice(0, MAX_ITEMS);
      deadLetter = parsedDeadLetter
        .filter(
          (item): item is AnilistQueuedUpdate =>
            item &&
            typeof item.key === 'string' &&
            typeof item.title === 'string' &&
            typeof item.episode === 'number' &&
            item.episode > 0 &&
            typeof item.createdAt === 'number' &&
            typeof item.attemptCount === 'number' &&
            typeof item.nextAttemptAt === 'number' &&
            (typeof item.lastError === 'string' || item.lastError === null),
        )
        .slice(0, MAX_ITEMS);
    } catch (error) {
      logger.error('Failed to load AniList retry queue.', error);
    }
  };

  load();

  return {
    enqueue(key: string, title: string, episode: number): void {
      const existing = pending.find((item) => item.key === key);
      if (existing) {
        return;
      }
      if (pending.length >= MAX_ITEMS) {
        pending.shift();
      }
      pending.push({
        key,
        title,
        episode,
        createdAt: Date.now(),
        attemptCount: 0,
        nextAttemptAt: Date.now(),
        lastError: null,
      });
      persist();
      logger.info(`Queued AniList retry for "${title}" episode ${episode}.`);
    },

    nextReady(nowMs: number = Date.now()): AnilistQueuedUpdate | null {
      const ready = pending.find((item) => item.nextAttemptAt <= nowMs);
      return ready ?? null;
    },

    markSuccess(key: string): void {
      const before = pending.length;
      pending = pending.filter((item) => item.key !== key);
      if (pending.length !== before) {
        persist();
      }
    },

    markFailure(key: string, reason: string, nowMs: number = Date.now()): void {
      const item = pending.find((candidate) => candidate.key === key);
      if (!item) {
        return;
      }
      item.attemptCount += 1;
      item.lastError = reason;
      if (item.attemptCount >= MAX_ATTEMPTS) {
        pending = pending.filter((candidate) => candidate.key !== key);
        if (deadLetter.length >= MAX_ITEMS) {
          deadLetter.shift();
        }
        deadLetter.push({
          ...item,
          nextAttemptAt: nowMs,
        });
        logger.warn('AniList retry moved to dead-letter queue.', {
          key,
          reason,
          attempts: item.attemptCount,
        });
        persist();
        return;
      }
      item.nextAttemptAt = nowMs + clampBackoffMs(item.attemptCount);
      persist();
      logger.warn('AniList retry scheduled with backoff.', {
        key,
        attemptCount: item.attemptCount,
        nextAttemptAt: item.nextAttemptAt,
        reason,
      });
    },

    getSnapshot(nowMs: number = Date.now()): AnilistRetryQueueSnapshot {
      const ready = pending.filter((item) => item.nextAttemptAt <= nowMs).length;
      return {
        pending: pending.length,
        ready,
        deadLetter: deadLetter.length,
      };
    },
  };
}
