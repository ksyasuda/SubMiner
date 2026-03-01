import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import { SecondarySubMode, SubtitlePosition } from '../../types';
import { createLogger } from '../../logger';

const logger = createLogger('main:subtitle-position');

export interface CycleSecondarySubModeDeps {
  getSecondarySubMode: () => SecondarySubMode;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  getLastSecondarySubToggleAtMs: () => number;
  setLastSecondarySubToggleAtMs: (timestampMs: number) => void;
  broadcastSecondarySubMode: (mode: SecondarySubMode) => void;
  showMpvOsd: (text: string) => void;
  now?: () => number;
}

const SECONDARY_SUB_CYCLE: SecondarySubMode[] = ['hidden', 'visible', 'hover'];
const SECONDARY_SUB_TOGGLE_DEBOUNCE_MS = 120;

export function cycleSecondarySubMode(deps: CycleSecondarySubModeDeps): void {
  const now = deps.now ? deps.now() : Date.now();
  if (now - deps.getLastSecondarySubToggleAtMs() < SECONDARY_SUB_TOGGLE_DEBOUNCE_MS) {
    return;
  }
  deps.setLastSecondarySubToggleAtMs(now);

  const currentMode = deps.getSecondarySubMode();
  const currentIndex = SECONDARY_SUB_CYCLE.indexOf(currentMode);
  const safeIndex = currentIndex >= 0 ? currentIndex : 0;
  const nextMode = SECONDARY_SUB_CYCLE[(safeIndex + 1) % SECONDARY_SUB_CYCLE.length]!;
  deps.setSecondarySubMode(nextMode);
  deps.broadcastSecondarySubMode(nextMode);
  deps.showMpvOsd(`Secondary subtitle: ${nextMode}`);
}

function getSubtitlePositionFilePath(mediaPath: string, subtitlePositionsDir: string): string {
  const key = normalizeMediaPathForSubtitlePosition(mediaPath);
  const hash = crypto.createHash('sha256').update(key).digest('hex');
  return path.join(subtitlePositionsDir, `${hash}.json`);
}

function normalizeMediaPathForSubtitlePosition(mediaPath: string): string {
  const trimmed = mediaPath.trim();
  if (!trimmed) return trimmed;

  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(trimmed) || /^ytsearch:/.test(trimmed)) {
    return trimmed;
  }

  const resolved = path.resolve(trimmed);
  let normalized = resolved;
  try {
    if (fs.existsSync(resolved)) {
      normalized = fs.realpathSync(resolved);
    }
  } catch {
    normalized = resolved;
  }

  if (process.platform === 'win32') {
    normalized = normalized.toLowerCase();
  }

  return normalized;
}

function persistSubtitlePosition(
  position: SubtitlePosition,
  currentMediaPath: string | null,
  subtitlePositionsDir: string,
): void {
  if (!currentMediaPath) return;
  if (!fs.existsSync(subtitlePositionsDir)) {
    fs.mkdirSync(subtitlePositionsDir, { recursive: true });
  }
  const positionPath = getSubtitlePositionFilePath(currentMediaPath, subtitlePositionsDir);
  fs.writeFileSync(positionPath, JSON.stringify(position, null, 2));
}

export function loadSubtitlePosition(
  options: {
    currentMediaPath: string | null;
    fallbackPosition: SubtitlePosition;
  } & { subtitlePositionsDir: string },
): SubtitlePosition | null {
  if (!options.currentMediaPath) {
    return options.fallbackPosition;
  }

  try {
    const positionPath = getSubtitlePositionFilePath(
      options.currentMediaPath,
      options.subtitlePositionsDir,
    );
    if (!fs.existsSync(positionPath)) {
      return options.fallbackPosition;
    }

    const data = fs.readFileSync(positionPath, 'utf-8');
    const parsed = JSON.parse(data) as Partial<SubtitlePosition>;
    if (parsed && typeof parsed.yPercent === 'number' && Number.isFinite(parsed.yPercent)) {
      return { yPercent: parsed.yPercent };
    }
    return options.fallbackPosition;
  } catch (err) {
    logger.error('Failed to load subtitle position:', (err as Error).message);
    return options.fallbackPosition;
  }
}

export function saveSubtitlePosition(options: {
  position: SubtitlePosition;
  currentMediaPath: string | null;
  subtitlePositionsDir: string;
  onQueuePending: (position: SubtitlePosition) => void;
  onPersisted: () => void;
}): void {
  if (!options.currentMediaPath) {
    options.onQueuePending(options.position);
    logger.warn('Queued subtitle position save - no media path yet');
    return;
  }

  try {
    persistSubtitlePosition(
      options.position,
      options.currentMediaPath,
      options.subtitlePositionsDir,
    );
    options.onPersisted();
  } catch (err) {
    logger.error('Failed to save subtitle position:', (err as Error).message);
  }
}

export function updateCurrentMediaPath(options: {
  mediaPath: unknown;
  currentMediaPath: string | null;
  pendingSubtitlePosition: SubtitlePosition | null;
  subtitlePositionsDir: string;
  loadSubtitlePosition: () => SubtitlePosition | null;
  setCurrentMediaPath: (mediaPath: string | null) => void;
  clearPendingSubtitlePosition: () => void;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
  broadcastSubtitlePosition: (position: SubtitlePosition | null) => void;
}): void {
  const nextPath =
    typeof options.mediaPath === 'string' && options.mediaPath.trim().length > 0
      ? options.mediaPath
      : null;
  if (nextPath === options.currentMediaPath) return;
  options.setCurrentMediaPath(nextPath);

  if (nextPath && options.pendingSubtitlePosition) {
    try {
      persistSubtitlePosition(
        options.pendingSubtitlePosition,
        nextPath,
        options.subtitlePositionsDir,
      );
      options.setSubtitlePosition(options.pendingSubtitlePosition);
      options.clearPendingSubtitlePosition();
    } catch (err) {
      logger.error('Failed to persist queued subtitle position:', (err as Error).message);
    }
  }

  const position = options.loadSubtitlePosition();
  options.broadcastSubtitlePosition(position);
}
